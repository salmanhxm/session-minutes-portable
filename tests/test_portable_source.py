from __future__ import annotations

import importlib.util
import re
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace


ROOT = Path(__file__).resolve().parents[1]
ENGINE_PATH = ROOT / "src" / "session_minutes.py"


def load_engine():
    spec = importlib.util.spec_from_file_location("portable_session_minutes", ENGINE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load portable engine")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class PortableSourceTests(unittest.TestCase):
    def test_source_contains_no_embedded_case_number_pilot_list(self):
        source = ENGINE_PATH.read_text(encoding="utf-8")
        self.assertNotIn("PILOT_PATHS", source)
        self.assertIsNone(re.search(r"4703\d{4}", source))

    def test_dynamic_pilot_selection_is_bounded_and_diverse(self):
        engine = load_engine()
        records = []
        for index in range(9):
            records.append(
                {
                    "status": "ready",
                    "source_path": f"8/1/example-{index}.docx",
                    "session_action": "insert_before_basis",
                    "current_opening_embedded": False,
                    "existing_notice_count": 1,
                    "table_count_before": 1,
                    "date_pickers_removed": 0,
                    "comment_markers_before": {},
                    "closing_action": "keep_current",
                }
            )
        records[1]["current_opening_embedded"] = True
        records[1]["session_action"] = "keep_current"
        records[2]["existing_notice_count"] = 2
        records[3]["table_count_before"] = 2
        records[4]["date_pickers_removed"] = 1
        records[4]["comment_markers_before"] = {"start": 1}
        records[5]["closing_action"] = "insert_current"

        selected = engine.select_pilot(records)

        self.assertEqual(len(selected), engine.PILOT_COUNT)
        self.assertEqual(len({record["source_path"] for record in selected}), len(selected))
        for expected in records[1:6]:
            self.assertIn(expected, selected)

    def test_direct_final_flag_is_explicit_and_available(self):
        engine = load_engine()
        args = engine.build_parser().parse_args(
            ["apply", "--allow-direct-apply"]
        )
        self.assertTrue(args.allow_direct_apply)
        self.assertEqual(args.mode, "apply")

    def test_closing_keep_properties_preserve_semantics_and_schema_order(self):
        engine = load_engine()
        paragraph = engine.etree.fromstring(
            f'''<w:p xmlns:w="{engine.W_NS}"><w:pPr><w:bidi/><w:jc w:val="right"/><w:rPr/></w:pPr><w:r><w:t>صدر هذا القرار</w:t></w:r></w:p>'''
        )
        before = engine.closing_semantic_xml_hash(paragraph)

        engine.keep_closing_with_following_table(paragraph)

        properties = paragraph.find("w:pPr", engine.NS)
        self.assertIsNotNone(properties)
        names = [engine.etree.QName(child).localname for child in properties]
        self.assertEqual(names[:2], ["keepNext", "keepLines"])
        self.assertLess(names.index("keepLines"), names.index("bidi"))
        self.assertEqual(before, engine.closing_semantic_xml_hash(paragraph))

    def test_signature_table_preserves_layout_and_allocates_safe_drawing_ids(self):
        engine = load_engine()
        root = engine.etree.fromstring(
            f'''<w:document xmlns:w="{engine.W_NS}" xmlns:wp="{engine.WP_NS}">
            <w:body>
              <w:p><w:r><w:drawing><wp:inline><wp:docPr id="31" name="existing"/></wp:inline></w:drawing></w:r></w:p>
              <w:tbl>
                <w:tblPr><w:tblpPr/><w:tblOverlap w:val="never"/></w:tblPr>
                <w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr><w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:drawing><wp:anchor><wp:docPr id="31" name="signature"/></wp:anchor></w:drawing></w:r><w:r><w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr><w:t>A deliberately long signature name for fitting</w:t></w:r></w:p></w:tc></w:tr>
              </w:tbl>
            </w:body></w:document>'''
        )
        table = root.find(".//w:tbl", engine.NS)
        before_layout = engine.signature_table_layout_hash(table)

        engine.normalize_signature_table_flow(
            table, root, reserved_drawing_ids={50, 51}
        )

        self.assertEqual(before_layout, engine.signature_table_layout_hash(table))
        self.assertIsNotNone(table.find("w:tblPr/w:tblpPr", engine.NS))
        self.assertIsNotNone(table.find("w:tblPr/w:tblOverlap", engine.NS))
        row_properties = table.find("w:tr/w:trPr", engine.NS)
        names = [engine.etree.QName(child).localname for child in row_properties]
        self.assertNotIn("cantSplit", names)
        self.assertIsNone(table.find(".//w:tcFitText", engine.NS))
        fitted_run = table.find(".//w:r[w:t]/w:rPr", engine.NS)
        self.assertEqual(fitted_run.find("w:sz", engine.NS).get(engine.W_VAL), "21")
        self.assertEqual(fitted_run.find("w:szCs", engine.NS).get(engine.W_VAL), "21")
        self.assertEqual(engine.duplicate_drawing_ids(root), [])
        self.assertGreater(
            min(engine.drawing_ids(table)), 51
        )

    def test_signature_table_uses_fixed_template_grid_and_inline_centering(self):
        engine = load_engine()
        table = engine.etree.fromstring(
            f'''<w:tbl xmlns:w="{engine.W_NS}">
              <w:tblPr><w:tblpPr w:horzAnchor="margin" w:tblpX="300"/><w:tblW w:w="5230" w:type="pct"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2600"/></w:tblGrid>
              <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="2400" w:type="pct"/></w:tcPr><w:p/></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="2600" w:type="pct"/></w:tcPr><w:p/></w:tc>
              </w:tr>
            </w:tbl>'''
        )
        expected_hash = engine.adapted_signature_table_visual_hash(table)

        engine.stabilize_signature_table_geometry(table)

        preferred = table.find("w:tblPr/w:tblW", engine.NS)
        floating = table.find("w:tblPr/w:tblpPr", engine.NS)
        justification = table.find("w:tblPr/w:jc", engine.NS)
        widths = table.xpath("./w:tr/w:tc/w:tcPr/w:tcW", namespaces=engine.NS)
        self.assertEqual(preferred.get(engine.qn(engine.W_NS, "type")), "dxa")
        self.assertEqual(preferred.get(engine.qn(engine.W_NS, "w")), "5000")
        self.assertIsNone(floating)
        self.assertEqual(justification.get(engine.W_VAL), "center")
        self.assertEqual(
            [item.get(engine.qn(engine.W_NS, "w")) for item in widths],
            ["2400", "2600"],
        )
        self.assertEqual(engine.signature_table_visual_hash(table), expected_hash)

    def test_signature_table_imports_missing_style_and_conflict_safe_base(self):
        engine = load_engine()
        template_styles_root = engine.etree.fromstring(
            f'''<w:styles xmlns:w="{engine.W_NS}">
              <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:pPr><w:spacing w:after="0"/></w:pPr></w:style>
              <w:style w:type="table" w:styleId="TableGridLight"><w:name w:val="Light Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders/></w:tblPr></w:style>
            </w:styles>'''
        )
        closing_styles = engine.collect_style_closure(
            template_styles_root, ["TableGridLight"], context="test template"
        )
        target_styles = engine.etree.fromstring(
            f'''<w:styles xmlns:w="{engine.W_NS}">
              <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Target Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr></w:style>
            </w:styles>'''
        )
        table = engine.etree.fromstring(
            f'''<w:tbl xmlns:w="{engine.W_NS}"><w:tblPr><w:tblStyle w:val="TableGridLight"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'''
        )
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "source.docx"
            with engine.ZipFile(source, "w") as package:
                package.writestr(engine.STYLES_XML, engine.xml_bytes(target_styles))
            replacements = {}
            allowed_changes = set()
            engine.import_signature_styles(
                source,
                table,
                SimpleNamespace(closing_styles=closing_styles),
                replacements,
                allowed_changes,
            )

        self.assertIn(engine.STYLES_XML, replacements)
        self.assertIn(engine.STYLES_XML, allowed_changes)
        imported_root = engine.parse_xml(replacements[engine.STYLES_XML])
        self.assertEqual(
            engine.signature_style_hashes(
                imported_root, table, context="test output"
            ),
            engine.Counter(
                engine.style_definition_visual_hash(style)
                for style in closing_styles.values()
            ),
        )

    def test_modified_template_dropdown_items_are_discovered_and_preserved(self):
        engine = load_engine()
        paragraph = engine.etree.fromstring(
            f'''<w:p xmlns:w="{engine.W_NS}"><w:sdt><w:sdtPr>
            <w:dropDownList>
              <w:listItem w:displayText="الخيار الأول" w:value="A"/>
              <w:listItem w:displayText="خيار مضاف من المستخدم" w:value="CUSTOM"/>
            </w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>الخيار الأول</w:t></w:r></w:sdtContent></w:sdt></w:p>'''
        )

        expected = (
            "dropDownList",
            (("الخيار الأول", "A"), ("خيار مضاف من المستخدم", "CUSTOM")),
        )
        self.assertEqual(
            engine.control_definition_signature(paragraph.find("w:sdt", engine.NS)),
            expected,
        )
        clone = engine.copy.deepcopy(paragraph)
        self.assertEqual(
            engine.paragraph_control_signature(clone),
            (expected,),
        )

    def test_session_discovery_accepts_variable_nested_control_counts(self):
        engine = load_engine()
        root = engine.etree.fromstring(
            f'''<w:document xmlns:w="{engine.W_NS}"><w:body><w:tbl><w:tr><w:tc>
            <w:p><w:r><w:t>{engine.SESSION_PREFIX}، في تمام الساعة 3:00 مساءً.</w:t></w:r>
              <w:sdt><w:sdtPr><w:comboBox/></w:sdtPr><w:sdtContent>
                <w:sdt><w:sdtPr><w:dropDownList/></w:sdtPr><w:sdtContent><w:r><w:t>خيار</w:t></w:r></w:sdtContent></w:sdt>
              </w:sdtContent></w:sdt>
            </w:p></w:tc></w:tr></w:tbl></w:body></w:document>'''
        )
        paragraph = engine.body_paragraphs(root, nonempty=True)[0]

        self.assertTrue(engine.is_template_session_paragraph(paragraph))
        self.assertEqual(
            engine.control_counter(paragraph),
            engine.Counter({"comboBox": 1, "dropDownList": 1}),
        )
        self.assertEqual(engine.paragraph_number(root, paragraph), 1)
        self.assertEqual(
            engine.paragraph_location(root, paragraph),
            "body/table[1]/row[1]/cell[1]/paragraph[1]",
        )

    def test_control_count_alone_never_classifies_a_session(self):
        engine = load_engine()
        controls = "".join(
            f'''<w:sdt><w:sdtPr><w:comboBox/></w:sdtPr><w:sdtContent><w:r><w:t>{index}</w:t></w:r></w:sdtContent></w:sdt>'''
            for index in range(7)
        )
        paragraph = engine.etree.fromstring(
            f'''<w:p xmlns:w="{engine.W_NS}"><w:r><w:t>فقرة إدارية غير مخصصة للجلسة</w:t></w:r>{controls}</w:p>'''
        )

        self.assertEqual(engine.control_counter(paragraph)["comboBox"], 7)
        self.assertFalse(engine.is_template_session_paragraph(paragraph))

    def test_template_error_names_file_paragraph_text_and_actual_counts(self):
        engine = load_engine()
        root = engine.etree.fromstring(
            f'''<w:document xmlns:w="{engine.W_NS}"><w:body><w:p>
            <w:r><w:t>{engine.SESSION_PREFIX} دون وقت صالح</w:t></w:r>
            <w:sdt><w:sdtPr><w:comboBox/></w:sdtPr><w:sdtContent><w:r><w:t>خيار</w:t></w:r></w:sdtContent></w:sdt>
            </w:p></w:body></w:document>'''
        )
        paragraph = engine.body_paragraphs(root, nonempty=True)[0]

        error = engine.template_validation_error(
            Path("regional-template.docx"),
            root,
            paragraph,
            "missing time",
            expected="one valid time; control counts may vary",
        )
        message = str(error)

        self.assertIn("File: regional-template.docx", message)
        self.assertIn("Paragraph: 1", message)
        self.assertIn("Text:", message)
        self.assertIn("ComboBox found: 1", message)
        self.assertIn("DropDownList found: 0", message)
        self.assertIn("Issue: missing time", message)

    def test_existing_session_remains_current_after_control_schema_edit(self):
        engine = load_engine()
        paragraph = engine.etree.fromstring(
            f'''<w:p xmlns:w="{engine.W_NS}"><w:r><w:t>{engine.SESSION_PREFIX}، في تمام الساعة 3:00 مساءً.</w:t></w:r>
            <w:sdt><w:sdtPr><w:comboBox/></w:sdtPr><w:sdtContent><w:r><w:t>قديم</w:t></w:r></w:sdtContent></w:sdt></w:p>'''
        )
        opening = SimpleNamespace(
            normalized_text=engine.normalize_text(engine.paragraph_text(paragraph)),
            control_signature=(
                ("comboBox", (("قديم", "OLD"),)),
                ("comboBox", (("مضاف حديثًا", "NEW"),)),
            ),
        )

        text_count, current_count, locations = engine.functional_opening_matches(
            [paragraph], opening
        )

        self.assertEqual(text_count, 1)
        self.assertEqual(current_count, 1)
        self.assertEqual(locations, [1])


if __name__ == "__main__":
    unittest.main(verbosity=2)
