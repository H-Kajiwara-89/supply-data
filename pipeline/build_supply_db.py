#!/usr/bin/env python3
"""厚生労働省「医療用医薬品供給状況」Excel からアプリ配信用の supply.sqlite を生成する。

標準ライブラリのみで動作する（Python 3.9+）。

出力（--out ディレクトリ）:
  supply-YYYYMMDD.sqlite.deflate  raw DEFLATE 圧縮した SQLite（iOS: NSData.decompressed(using: .zlib)）
  manifest.json                   スキーマ版・データ時点・ファイル名・sha256・サイズ
  supply.sqlite                   非圧縮版（アプリ同梱用・確認用）

スキーマは docs/xcode-prompts.md「プロンプト0」と同一（schema_version = 1）。
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import io
import json
import re
import sqlite3
import sys
import unicodedata
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
import zlib
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

SCHEMA_VERSION = 1
MHLW_PAGE_URL = (
    "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
)
MHLW_BASE = "https://www.mhlw.go.jp"
USER_AGENT = "SupplyChecker-DataPipeline/1.0"
JST = dt.timezone(dt.timedelta(hours=9))
MIN_ROWS = 10_000

# Excel 列（A〜U）と、ヘッダに含まれているべき文字列。列構成の変更を検出するために使う。
EXPECTED_HEADERS = {
    "A": "薬剤区分",
    "C": "成分名",
    "D": "規格単位",
    "E": "YJコード",
    "F": "品名",
    "G": "製造販売業者名",
    "H": "製品区分",
    "L": "出荷対応",
    "M": "⑫の情報を更新した日",
    "N": "理由",
    "O": "解除見込み",
    "P": "在庫消尽時期",
    "Q": "出荷量",
    "R": "改善（増加）見込み時期",
    "U": "今回掲載時の更新有無",
}

STATUS_PREFIX = {"①": 1, "②": 2, "③": 3, "④": 4, "⑤": 5}
# 出荷量は "Aプラス" を "A" より先に判定する
VOLUME_PREFIX = [("Aプラス", 1), ("A", 2), ("B", 3), ("C", 4), ("D", 5)]

SCHEMA_SQL = """
CREATE TABLE meta(key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE products(
  yj TEXT PRIMARY KEY,
  group9 TEXT NOT NULL,
  group7 TEXT NOT NULL,
  name TEXT NOT NULL,
  ingredient TEXT NOT NULL,
  standard TEXT NOT NULL,
  maker TEXT NOT NULL,
  category TEXT NOT NULL,
  route TEXT NOT NULL,
  status INTEGER NOT NULL,
  status_updated TEXT,
  reason TEXT,
  outlook TEXT,
  outlook_detail TEXT,
  volume INTEGER,
  volume_outlook TEXT,
  is_new INTEGER NOT NULL DEFAULT 0,
  name_norm TEXT NOT NULL,
  ingredient_norm TEXT NOT NULL
);
CREATE INDEX idx_group9 ON products(group9);
CREATE INDEX idx_group7 ON products(group7);
CREATE TABLE code_map(code TEXT PRIMARY KEY, code_type INTEGER NOT NULL, yj TEXT NOT NULL);
"""


class PipelineError(Exception):
    """配信を中止すべき異常。"""


# ---------------------------------------------------------------- 正規化

def normalize(text: Optional[str]) -> str:
    """アプリ側 TextNormalizer と同一ルール。

    1. NFKC  2. ひらがな→カタカナ  3. ASCII 小文字化  4. 空白と「【般】」を除去
    """
    if not text:
        return ""
    s = unicodedata.normalize("NFKC", text)
    s = "".join(chr(ord(c) + 0x60) if "ぁ" <= c <= "ゖ" else c for c in s)
    s = "".join(c.lower() if c.isascii() else c for c in s)
    s = s.replace("【般】", "")
    return re.sub(r"\s+", "", s)


# ---------------------------------------------------------------- 取得

def http_get(url: str, timeout: int = 60) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as res:
        return res.read()


def find_latest_xlsx_url(page_html: str) -> str:
    links = re.findall(r'href="([^"]*iyakuhinkyoukyu[^"]*\.xlsx)"', page_html)
    if not links:
        links = re.findall(r'href="([^"]+\.xlsx)"', page_html)
    if not links:
        raise PipelineError("厚労省ページに .xlsx へのリンクが見つかりません")
    href = links[0]
    return href if href.startswith("http") else MHLW_BASE + href


def data_date_from_filename(url: str) -> Optional[dt.date]:
    m = re.search(r"/(\d{6})iyakuhinkyoukyu\.xlsx$", url)
    if not m:
        return None
    yy, mm, dd = m.group(1)[:2], m.group(1)[2:4], m.group(1)[4:]
    try:
        return dt.date(2000 + int(yy), int(mm), int(dd))
    except ValueError:
        return None


# ---------------------------------------------------------------- xlsx 読み込み（標準ライブラリのみ）

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _col_letters(ref: str) -> str:
    return "".join(ch for ch in ref if ch.isalpha())


def read_xlsx_rows(data: bytes) -> List[Dict[str, str]]:
    """最初のシートを {列記号: 文字列} の行リストとして返す。"""
    z = zipfile.ZipFile(io.BytesIO(data))
    shared: List[str] = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(_NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(_NS + "t")))
    sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    rows: List[Dict[str, str]] = []
    for r in sheet.iter(_NS + "row"):
        row: Dict[str, str] = {}
        for c in r.findall(_NS + "c"):
            t = c.get("t")
            if t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter(_NS + "t"))
            else:
                v = c.find(_NS + "v")
                if v is None or v.text is None:
                    continue
                val = shared[int(v.text)] if t == "s" else v.text
            row[_col_letters(c.get("r", ""))] = val
        rows.append(row)
    return rows


# ---------------------------------------------------------------- 値の変換

def clean(v: Optional[str]) -> Optional[str]:
    if v is None:
        return None
    v = v.strip()
    return v or None


def strip_code_prefix(v: Optional[str]) -> Optional[str]:
    """「2：後発品」「３．製造トラブル」「ア. あり」のような番号付きラベルから番号を外す。"""
    v = clean(v)
    if v is None:
        return None
    return re.sub(r"^[0-9０-９A-Za-zア-ン]+\s*[：:．.]\s*", "", v) or v


def is_dash(v: Optional[str]) -> bool:
    """「７．－」「エ. －」のような「該当なし」表記か。"""
    v = clean(v)
    return v is None or v.endswith("－") or v.endswith("-")


def excel_date(v: Optional[str]) -> Optional[str]:
    """Excel シリアル値または日付文字列を YYYY-MM-DD にする。変換できない文字列はそのまま返す。"""
    v = clean(v)
    if v is None:
        return None
    if re.fullmatch(r"\d+(\.0+)?", v):
        n = int(float(v))
        if 20000 < n < 80000:
            return (dt.date(1899, 12, 30) + dt.timedelta(days=n)).isoformat()
    return v


def label_or_date(v: Optional[str]) -> Optional[str]:
    """日付シリアルなら日付に、「e．３か月超」のような番号付きラベルなら番号を外す。"""
    d = excel_date(v)
    if d is None or re.fullmatch(r"\d{4}-\d{2}-\d{2}", d):
        return d
    return strip_code_prefix(d)


def parse_status(v: Optional[str]) -> int:
    v = clean(v) or ""
    for prefix, code in STATUS_PREFIX.items():
        if v.startswith(prefix):
            return code
    raise PipelineError(f"想定外の出荷対応状況: {v!r}")


def parse_volume(v: Optional[str]) -> Optional[int]:
    v = clean(v)
    if v is None:
        return None
    for prefix, code in VOLUME_PREFIX:
        if v.startswith(prefix):
            return code
    raise PipelineError(f"想定外の出荷量状況: {v!r}")


# ---------------------------------------------------------------- 供給状況 Excel → products

def find_header_row(rows: List[Dict[str, str]]) -> int:
    for i, row in enumerate(rows[:10]):
        if "YJコード" in (row.get("E") or ""):
            return i
    raise PipelineError("ヘッダ行（⑤YJコード）が見つかりません。列構成が変わった可能性があります")


def validate_headers(header: Dict[str, str]) -> None:
    problems = [
        f"{col}列: 期待「{want}」 実際「{(header.get(col) or '').replace(chr(10), ' ')}」"
        for col, want in EXPECTED_HEADERS.items()
        if want not in (header.get(col) or "")
    ]
    if problems:
        raise PipelineError("Excel の列構成が想定と異なります:\n  " + "\n  ".join(problems))


def build_products(rows: List[Dict[str, str]]) -> List[Tuple]:
    h = find_header_row(rows)
    validate_headers(rows[h])
    products: List[Tuple] = []
    seen = set()
    for i, r in enumerate(rows[h + 1:], start=h + 2):
        yj = clean(r.get("E"))
        if not yj:
            continue
        if yj in seen:
            raise PipelineError(f"YJコードが重複しています: {yj}（{i}行目）")
        seen.add(yj)
        name = clean(r.get("F")) or ""
        ingredient = clean(r.get("C")) or ""
        try:
            status = parse_status(r.get("L"))
            volume = parse_volume(r.get("Q"))
        except PipelineError as e:
            raise PipelineError(f"{i}行目: {e}") from None
        products.append((
            yj,
            yj[:9],
            yj[:7],
            name,
            ingredient,
            clean(r.get("D")) or "",
            clean(r.get("G")) or "",
            strip_code_prefix(r.get("H")) or "",
            strip_code_prefix(r.get("A")) or "",
            status,
            excel_date(r.get("M")),
            None if is_dash(r.get("N")) else strip_code_prefix(r.get("N")),
            None if is_dash(r.get("O")) else strip_code_prefix(r.get("O")),
            excel_date(r.get("P")),
            volume,
            label_or_date(r.get("R")),
            1 if (clean(r.get("U")) or "").lower() == "new" else 0,
            normalize(name),
            normalize(ingredient),
        ))
    if len(products) < MIN_ROWS:
        raise PipelineError(f"品目数が少なすぎます: {len(products)}（下限 {MIN_ROWS}）")
    return products


# ---------------------------------------------------------------- 支払基金 医薬品マスター → code_map

YAKKA_CODE_RE = re.compile(r"^\d{4}[0-9A-Z]\d{2}[A-Z][0-9A-Z]\d{3}$")
RECE_CODE_RE = re.compile(r"^6\d{8}$")


def read_ssk_master(data: bytes) -> List[List[str]]:
    """医薬品マスター（zip または csv、Shift_JIS）を行リストにする。"""
    if data[:2] == b"PK":
        z = zipfile.ZipFile(io.BytesIO(data))
        names = [n for n in z.namelist() if n.lower().endswith((".csv", ".txt"))]
        if not names:
            raise PipelineError("医薬品マスター zip に CSV がありません")
        data = z.read(names[0])
    text = data.decode("cp932", errors="replace")
    return list(csv.reader(io.StringIO(text)))


def build_code_map(master_rows: Iterable[List[str]], known_yj: set) -> List[Tuple[str, int, str]]:
    """レセプト電算コード(9桁) → YJ の対応を作る。

    マスターの列位置に依存しないよう、各行から「6で始まる9桁」と「薬価基準コード形式の12桁」を探す。
    薬価基準コードが供給状況 Excel の YJ に存在するものだけを採用する（統一名収載品などは対象外）。
    """
    out: Dict[str, str] = {}
    rows_with_codes = 0
    for row in master_rows:
        fields = [f.strip() for f in row]
        rece = next((f for f in fields[:5] if RECE_CODE_RE.match(f)), None)
        yakka = next((f for f in fields if YAKKA_CODE_RE.match(f)), None)
        if not rece or not yakka:
            continue
        rows_with_codes += 1
        if yakka in known_yj:
            out[rece] = yakka
    if rows_with_codes == 0:
        raise PipelineError("医薬品マスターからコードを読み取れませんでした（形式変更の可能性）")
    return [(code, 2, yj) for code, yj in sorted(out.items())]


def fetch_previous_db(manifest_url: str, dest: Path) -> Path:
    """配信中の manifest.json から前回の DB を取得して展開する。"""
    manifest = json.loads(http_get(manifest_url))
    packed_url = urllib.parse.urljoin(manifest_url, manifest["url"])
    packed = http_get(packed_url)
    if hashlib.sha256(packed).hexdigest() != manifest["sha256"]:
        raise PipelineError("前回DBの sha256 が manifest と一致しません")
    dest.write_bytes(zlib.decompress(packed, wbits=-15))
    return dest


def code_map_from_previous_db(path: Path, known_yj: set) -> List[Tuple[str, int, str]]:
    con = sqlite3.connect(str(path))
    try:
        rows = con.execute("SELECT code, code_type, yj FROM code_map").fetchall()
    finally:
        con.close()
    return [r for r in rows if r[2] in known_yj]


# ---------------------------------------------------------------- 出力

def write_sqlite(path: Path, products: List[Tuple], code_map: List[Tuple], meta: Dict[str, str]) -> None:
    if path.exists():
        path.unlink()
    con = sqlite3.connect(str(path))
    try:
        con.executescript(SCHEMA_SQL)
        con.executemany(f"INSERT INTO products VALUES ({','.join('?' * 19)})", products)
        con.executemany("INSERT INTO code_map VALUES (?,?,?)", code_map)
        con.executemany("INSERT INTO meta VALUES (?,?)", sorted(meta.items()))
        con.commit()
        con.execute("VACUUM")
    finally:
        con.close()


def deflate(data: bytes) -> bytes:
    """raw DEFLATE（RFC 1951、ヘッダなし）。iOS の Compression .zlib と互換。"""
    c = zlib.compressobj(level=9, wbits=-15)
    return c.compress(data) + c.flush()


def verify_db(path: Path) -> Dict[str, int]:
    con = sqlite3.connect(str(path))
    try:
        stats = {
            "products": con.execute("SELECT COUNT(*) FROM products").fetchone()[0],
            "code_map": con.execute("SELECT COUNT(*) FROM code_map").fetchone()[0],
        }
        for s, n in con.execute("SELECT status, COUNT(*) FROM products GROUP BY status"):
            stats[f"status_{s}"] = n
        if con.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise PipelineError("SQLite integrity_check に失敗しました")
    finally:
        con.close()
    return stats


def run(args: argparse.Namespace) -> Dict:
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    # 1. 供給状況 Excel
    if args.xlsx:
        xlsx_bytes = Path(args.xlsx).read_bytes()
        source_url = args.source_url or f"file:{Path(args.xlsx).name}"
    else:
        source_url = find_latest_xlsx_url(http_get(MHLW_PAGE_URL).decode("utf-8", errors="replace"))
        xlsx_bytes = http_get(source_url)
    data_date = (
        dt.date.fromisoformat(args.data_date) if args.data_date
        else data_date_from_filename(source_url) or dt.datetime.now(JST).date()
    )
    products = build_products(read_xlsx_rows(xlsx_bytes))
    known_yj = {p[0] for p in products}

    # 2. code_map（支払基金マスター → 前回DB → なし の順で試す）
    code_map: List[Tuple] = []
    code_map_source = "none"
    warnings: List[str] = []
    master_src = args.ssk_master
    if master_src:
        try:
            raw = http_get(master_src) if master_src.startswith("http") else Path(master_src).read_bytes()
            code_map = build_code_map(read_ssk_master(raw), known_yj)
            code_map_source = "ssk_master"
        except Exception as e:  # 取得失敗でも配信は止めない
            warnings.append(f"医薬品マスターを使えませんでした: {e}")
    previous_db = Path(args.previous_db) if args.previous_db else None
    if not code_map and not (previous_db and previous_db.exists()) and args.previous_manifest_url:
        try:
            previous_db = fetch_previous_db(args.previous_manifest_url, out / "previous.sqlite")
        except Exception as e:  # 初回配信時などは存在しない
            warnings.append(f"前回DBを取得できませんでした: {e}")
    if not code_map and previous_db and previous_db.exists():
        code_map = code_map_from_previous_db(previous_db, known_yj)
        code_map_source = "previous_db"
    if not code_map:
        warnings.append("code_map が空です（レセ電算コードは名称照合にフォールバックします）")

    # 3. 書き出し
    generated_at = dt.datetime.now(JST).replace(microsecond=0).isoformat()
    meta = {
        "schema_version": str(SCHEMA_VERSION),
        "data_date": data_date.isoformat(),
        "generated_at": generated_at,
        "source_url": source_url,
        "code_map_source": code_map_source,
    }
    db_path = out / "supply.sqlite"
    write_sqlite(db_path, products, code_map, meta)
    stats = verify_db(db_path)

    raw_db = db_path.read_bytes()
    packed = deflate(raw_db)
    if zlib.decompress(packed, wbits=-15) != raw_db:
        raise PipelineError("圧縮結果の検証に失敗しました")
    packed_name = f"supply-{data_date.strftime('%Y%m%d')}.sqlite.deflate"
    for old in out.glob("supply-*.sqlite.deflate"):
        if old.name != packed_name:
            old.unlink()
    (out / packed_name).write_bytes(packed)

    (out / "previous.sqlite").unlink(missing_ok=True)  # 配信物に含めない

    manifest = {
        "schema_version": SCHEMA_VERSION,
        "data_date": data_date.isoformat(),
        "generated_at": generated_at,
        "url": packed_name,  # manifest.json からの相対パス
        "sha256": hashlib.sha256(packed).hexdigest(),
        "size": len(packed),
        "uncompressed_size": len(raw_db),
        "stats": stats,
    }
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {"manifest": manifest, "warnings": warnings}


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--out", default="dist", help="出力ディレクトリ")
    p.add_argument("--xlsx", help="ローカルの供給状況 Excel（省略時は厚労省から取得）")
    p.add_argument("--source-url", help="--xlsx 使用時に meta に記録する元 URL")
    p.add_argument("--data-date", help="データ時点 YYYY-MM-DD（省略時はファイル名から推定）")
    p.add_argument("--ssk-master", help="支払基金 医薬品マスター（zip/csv の URL またはパス）")
    p.add_argument("--previous-db", help="前回配信の supply.sqlite（code_map の引き継ぎ用）")
    p.add_argument("--previous-manifest-url", help="配信中の manifest.json の URL（前回DBを自動取得）")
    args = p.parse_args(argv)
    try:
        result = run(args)
    except PipelineError as e:
        print(f"::error::{e}", file=sys.stderr)
        return 1
    for w in result["warnings"]:
        print(f"::warning::{w}", file=sys.stderr)
    print(json.dumps(result["manifest"], ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
