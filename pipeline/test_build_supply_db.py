"""build_supply_db のテスト。実行: python3 -m unittest -v"""
import argparse
import hashlib
import io
import json
import sqlite3
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path
from unittest import mock
from xml.sax.saxutils import escape

import build_supply_db as b

HEADER = {
    "A": "①薬剤区分", "B": "②薬効分類", "C": "③成分名", "D": "④規格単位\n※全角", "E": "⑤YJコード",
    "F": "⑥品名", "G": "⑦製造販売業者名", "H": "⑧製品区分", "I": "⑨基礎的\n医薬品",
    "J": "⑩（重要）供給確保医薬品", "K": "⑪薬価収載年月日", "L": "⑫製造販売業者の\n「出荷対応」の状況",
    "M": "⑬当該品目の⑫の情報を更新した日（…）", "N": "⑭限定出荷/供給停止の理由",
    "O": "⑮限定出荷の解除見込み／\n供給停止の解消見込み",
    "P": "⑯限定出荷の解除見込み／\n供給停止の解消見込み／\n販売中止品の在庫消尽時期",
    "Q": "⑰製造販売業者の\n「出荷量」の現在の状況", "R": "⑱製造販売業者の「出荷量」の改善（増加）見込み時期",
    "S": "⑲改善見込み量", "T": "⑳当該品目の⑫以外の情報を更新した日", "U": "今回掲載時の更新有無（更新有りの場合、Newと表示）",
}


def row(yj, name, status="①通常出荷", **kw):
    r = {
        "A": "1：内用薬", "C": "アムロジピンベシル酸塩", "D": "２．５ｍｇ１錠", "E": yj, "F": name,
        "G": "テスト製薬", "H": "2：後発品", "L": status, "N": "７．－", "O": "エ. －",
        "Q": "A．出荷量通常", "U": "",
    }
    r.update(kw)
    return r


def make_xlsx(rows):
    """inlineStr だけで構成した最小の xlsx を作る。"""
    body = []
    for i, r in enumerate(rows, start=1):
        cells = "".join(
            f'<c r="{col}{i}" t="inlineStr"><is><t>{escape(v)}</t></is></c>'
            if not v.isdigit() else f'<c r="{col}{i}"><v>{v}</v></c>'
            for col, v in r.items() if v != ""
        )
        body.append(f'<row r="{i}">{cells}</row>')
    sheet = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        + "".join(body) + "</sheetData></worksheet>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return buf.getvalue()


SAMPLE = [
    {},
    HEADER,
    row("2171022F1029", "ノルバスク錠２．５ｍｇ", H="1：先発品"),
    row("2171022F1037", "アムロジピン錠２．５ｍｇ「Ａ」", "③限定出荷（他社品の影響）",
        M="46260", N="１．需要増", O="ウ. 未定", Q="B．出荷量減少", R="e．３か月超", U="New"),
    row("2171022F1045", "アムロジピン錠２．５ｍｇ「Ｂ」", "⑤供給停止",
        N="３．製造トラブル（製造委託を含む）", O="ア. あり", P="2027年3月", Q="C．出荷停止"),
]


class NormalizeTests(unittest.TestCase):
    def test_rules(self):
        self.assertEqual(b.normalize("【般】ニフェジピンカプセル１０ｍｇ"), "ニフェジピンカプセル10mg")
        self.assertEqual(b.normalize("のるばすく"), "ノルバスク")
        self.assertEqual(b.normalize("ｄ－クロルフェニラミン　マレイン酸塩"), "d-クロルフェニラミンマレイン酸塩")
        self.assertEqual(b.normalize(None), "")


class ValueTests(unittest.TestCase):
    def test_status(self):
        self.assertEqual(b.parse_status("①通常出荷"), 1)
        self.assertEqual(b.parse_status("④限定出荷（その他）"), 4)
        self.assertEqual(b.parse_status("⑤供給停止"), 5)
        with self.assertRaises(b.PipelineError):
            b.parse_status("出荷調整")

    def test_volume(self):
        self.assertEqual(b.parse_volume("Aプラス．出荷量増加"), 1)
        self.assertEqual(b.parse_volume("A．出荷量通常"), 2)
        self.assertEqual(b.parse_volume("D．薬価削除予定"), 5)
        self.assertIsNone(b.parse_volume(None))
        with self.assertRaises(b.PipelineError):
            b.parse_volume("X．不明")

    def test_labels_and_dates(self):
        self.assertEqual(b.strip_code_prefix("2：後発品"), "後発品")
        self.assertEqual(b.strip_code_prefix("３．製造トラブル（製造委託を含む）"), "製造トラブル（製造委託を含む）")
        self.assertEqual(b.strip_code_prefix("ア. あり"), "あり")
        self.assertTrue(b.is_dash("７．－"))
        self.assertTrue(b.is_dash("エ. －"))
        self.assertFalse(b.is_dash("ウ. 未定"))
        self.assertEqual(b.excel_date("46260"), "2026-08-26")
        self.assertEqual(b.excel_date("在庫消尽済"), "在庫消尽済")
        self.assertEqual(b.label_or_date("e．３か月超"), "３か月超")
        self.assertEqual(b.label_or_date("46260"), "2026-08-26")

    def test_data_date_from_filename(self):
        url = "https://www.mhlw.go.jp/content/10800000/260929iyakuhinkyoukyu.xlsx"
        self.assertEqual(str(b.data_date_from_filename(url)), "2026-09-29")
        self.assertIsNone(b.data_date_from_filename("https://example.com/x.xlsx"))

    def test_find_xlsx_link(self):
        html = '<a href="/content/10800000/260930iyakuhinkyoukyu.xlsx">供給状況</a>'
        self.assertEqual(b.find_latest_xlsx_url(html),
                         "https://www.mhlw.go.jp/content/10800000/260930iyakuhinkyoukyu.xlsx")
        with self.assertRaises(b.PipelineError):
            b.find_latest_xlsx_url("<html></html>")


@mock.patch.object(b, "MIN_ROWS", 1)
class ProductTests(unittest.TestCase):
    def test_build_products(self):
        products = b.build_products(b.read_xlsx_rows(make_xlsx(SAMPLE)))
        self.assertEqual(len(products), 3)
        by_yj = {p[0]: p for p in products}
        p = by_yj["2171022F1037"]
        self.assertEqual(p[1:3], ("2171022F1", "2171022"))
        self.assertEqual(p[7:10], ("後発品", "内用薬", 3))
        self.assertEqual(p[10:17], ("2026-08-26", "需要増", "未定", None, 3, "３か月超", 1))
        self.assertEqual(p[17], "アムロジピン錠2.5mg「a」")
        stop = by_yj["2171022F1045"]
        self.assertEqual((stop[9], stop[11], stop[12], stop[13], stop[14]),
                         (5, "製造トラブル（製造委託を含む）", "あり", "2027年3月", 4))
        normal = by_yj["2171022F1029"]
        self.assertEqual((normal[11], normal[12]), (None, None))

    def test_header_change_is_rejected(self):
        broken = dict(HEADER, E="⑤薬価基準コード")
        with self.assertRaises(b.PipelineError):
            b.build_products(b.read_xlsx_rows(make_xlsx([{}, broken] + SAMPLE[2:])))
        moved = dict(HEADER, L="⑫出荷量")
        with self.assertRaisesRegex(b.PipelineError, "L列"):
            b.build_products(b.read_xlsx_rows(make_xlsx([{}, moved] + SAMPLE[2:])))

    def test_duplicate_yj_is_rejected(self):
        with self.assertRaisesRegex(b.PipelineError, "重複"):
            b.build_products(b.read_xlsx_rows(make_xlsx(SAMPLE + [SAMPLE[2]])))

    def test_unknown_status_is_rejected(self):
        bad = row("2171022F1053", "X", "⑥新区分")
        with self.assertRaisesRegex(b.PipelineError, "出荷対応"):
            b.build_products(b.read_xlsx_rows(make_xlsx(SAMPLE + [bad])))

    def test_too_few_rows(self):
        with mock.patch.object(b, "MIN_ROWS", 100):
            with self.assertRaisesRegex(b.PipelineError, "少なすぎ"):
                b.build_products(b.read_xlsx_rows(make_xlsx(SAMPLE)))


class CodeMapTests(unittest.TestCase):
    def test_build_code_map(self):
        csv_text = (
            "0,Y,612170709,ﾉﾙﾊﾞｽｸ,x,,,,2171022F1029,\n"
            "0,Y,620000001,不明品,x,,,,9999999F9999,\n"   # 供給状況に無い → 除外
            "0,Y,620000002,コードなし,x,,,,,\n"
        ).encode("cp932")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            z.writestr("y_ALL20260317.csv", csv_text)
        rows = b.read_ssk_master(buf.getvalue())
        self.assertEqual(b.build_code_map(rows, {"2171022F1029"}), [("612170709", 2, "2171022F1029")])

    def test_unreadable_master(self):
        with self.assertRaises(b.PipelineError):
            b.build_code_map([["a", "b"]], set())


@mock.patch.object(b, "MIN_ROWS", 1)
class EndToEndTests(unittest.TestCase):
    def run_pipeline(self, tmp, **kw):
        xlsx = Path(tmp) / "in.xlsx"
        xlsx.write_bytes(make_xlsx(SAMPLE))
        args = argparse.Namespace(
            out=str(Path(tmp) / "dist"), xlsx=str(xlsx), source_url=None, data_date="2026-09-29",
            ssk_master=None, previous_db=None, previous_manifest_url=None,
        )
        for k, v in kw.items():
            setattr(args, k, v)
        return b.run(args), Path(args.out)

    def test_outputs(self):
        with tempfile.TemporaryDirectory() as tmp:
            result, out = self.run_pipeline(tmp)
            m = json.loads((out / "manifest.json").read_text())
            self.assertEqual(m["schema_version"], 1)
            self.assertEqual(m["data_date"], "2026-09-29")
            self.assertEqual(m["url"], "supply-20260929.sqlite.deflate")
            packed = (out / m["url"]).read_bytes()
            self.assertEqual(hashlib.sha256(packed).hexdigest(), m["sha256"])
            self.assertEqual(len(packed), m["size"])
            db = zlib.decompress(packed, wbits=-15)
            self.assertEqual(db, (out / "supply.sqlite").read_bytes())
            con = sqlite3.connect(str(out / "supply.sqlite"))
            meta = dict(con.execute("SELECT key, value FROM meta"))
            self.assertEqual(meta["schema_version"], "1")
            self.assertEqual(con.execute("SELECT COUNT(*) FROM products WHERE group9='2171022F1'").fetchone()[0], 3)
            self.assertIn("code_map が空", " ".join(result["warnings"]))
            self.assertFalse((out / "previous.sqlite").exists())

    def test_code_map_is_carried_over_from_previous_db(self):
        with tempfile.TemporaryDirectory() as tmp:
            prev = Path(tmp) / "prev.sqlite"
            con = sqlite3.connect(str(prev))
            con.execute("CREATE TABLE code_map(code TEXT PRIMARY KEY, code_type INTEGER NOT NULL, yj TEXT NOT NULL)")
            con.executemany("INSERT INTO code_map VALUES (?,?,?)",
                            [("612170709", 2, "2171022F1029"), ("699999999", 2, "0000000X0000")])
            con.commit()
            con.close()
            _, out = self.run_pipeline(tmp, ssk_master=str(Path(tmp) / "missing.zip"), previous_db=str(prev))
            con = sqlite3.connect(str(out / "supply.sqlite"))
            self.assertEqual(con.execute("SELECT code, yj FROM code_map").fetchall(), [("612170709", "2171022F1029")])
            self.assertEqual(con.execute("SELECT value FROM meta WHERE key='code_map_source'").fetchone()[0],
                             "previous_db")


if __name__ == "__main__":
    unittest.main()
