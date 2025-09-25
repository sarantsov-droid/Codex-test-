 (cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF' 
diff --git a//dev/null b/clean_data.py
index 0000000000000000000000000000000000000000..4c5ee39b3d7dfde12d475798c973b622d2870efa 100644
--- a//dev/null
+++ b/clean_data.py
@@ -0,0 +1,302 @@
+"""Transform Data.csv into a long-format Excel file with cleaned values."""
+from __future__ import annotations
+
+import csv
+import zipfile
+from datetime import datetime, date
+from pathlib import Path
+from typing import Dict, Iterable, List, Sequence, Tuple
+from xml.sax.saxutils import escape
+
+DATA_FILE = Path(__file__).resolve().parent / "Data.csv"
+OUTPUT_FILE = Path(__file__).resolve().parent / "akcii_clean.xlsx"
+
+DATE_LABEL = "Дата акції від дати завантаження до дати закінчення, план"
+FIELD_MAP = {
+    "Мережа": "Мережа",
+    "Продукція": "Продукт",
+    "Еталонний місяць": "Еталонний місяць",
+    "Акційний місяць": "Акційний місяць",
+    "Продажі в еталоний період, грн": "Продажі грн",
+    "Продажі в еталоний період, шт": "Продажі шт",
+    "Маржа акційний період, %": "Маржа %",
+    "Відсоток акційної знижки, %": "Знижка %",
+}
+FINAL_COLUMNS = [
+    "Менеджер",
+    "Дата початку акції",
+    "Дата закінчення акції",
+    "Мережа",
+    "Продукт",
+    "Еталонний місяць",
+    "Акційний місяць",
+    "Продажі грн",
+    "Продажі шт",
+    "Маржа %",
+    "Знижка %",
+]
+TEXT_COLUMNS = {"Менеджер", "Мережа", "Продукт", "Еталонний місяць", "Акційний місяць"}
+NUMERIC_COLUMNS = {"Продажі грн", "Продажі шт"}
+PERCENT_COLUMNS = {"Маржа %", "Знижка %"}
+
+CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
+<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
+  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
+  <Default Extension="xml" ContentType="application/xml"/>
+  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
+  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
+  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
+</Types>
+"""
+
+RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
+<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
+  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
+</Relationships>
+"""
+
+WORKBOOK_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
+<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
+          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
+  <sheets>
+    <sheet name="Акції" sheetId="1" r:id="rId1"/>
+  </sheets>
+</workbook>
+"""
+
+WORKBOOK_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
+<Relationships xmlns="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
+  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
+  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
+</Relationships>
+"""
+
+STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
+<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
+  <fonts count="1">
+    <font>
+      <sz val="11"/>
+      <color theme="1"/>
+      <name val="Calibri"/>
+      <family val="2"/>
+      <scheme val="minor"/>
+    </font>
+  </fonts>
+  <fills count="2">
+    <fill><patternFill patternType="none"/></fill>
+    <fill><patternFill patternType="gray125"/></fill>
+  </fills>
+  <borders count="1">
+    <border>
+      <left/><right/><top/><bottom/><diagonal/>
+    </border>
+  </borders>
+  <cellStyleXfs count="1">
+    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
+  </cellStyleXfs>
+  <cellXfs count="1">
+    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
+  </cellXfs>
+  <cellStyles count="1">
+    <cellStyle name="Normal" xfId="0" builtinId="0"/>
+  </cellStyles>
+</styleSheet>
+"""
+
+
+def read_raw_data(path: Path) -> Tuple[List[str], List[List[str]]]:
+    """Read the semicolon-separated CSV file."""
+    with path.open("r", newline="", encoding="utf-8-sig") as handle:
+        reader = csv.reader(handle, delimiter=";")
+        header = next(reader)
+        rows = [row for row in reader]
+    return header, rows
+
+
+def value_at(row: Sequence[str], index: int) -> str:
+    """Return the value at index or empty string if missing."""
+    if index < len(row):
+        return row[index]
+    return ""
+
+
+def normalise_whitespace(value: str) -> str:
+    """Normalise whitespace by removing non-breaking spaces and trimming."""
+    if value is None:
+        return ""
+    text = value.replace("\u00a0", " ").replace("\u202f", " ") if isinstance(value, str) else str(value)
+    return text.strip()
+
+
+def parse_single_date(value: str) -> date | None:
+    """Parse a date string in dd.mm.yy or dd.mm.yyyy format."""
+    text = normalise_whitespace(value)
+    if not text:
+        return None
+    for fmt in ("%d.%m.%Y", "%d.%m.%y"):
+        try:
+            parsed = datetime.strptime(text, fmt)
+            return parsed.date()
+        except ValueError:
+            continue
+    return None
+
+
+def format_date(value: date | None) -> str | None:
+    """Return ISO formatted date string if value is not None."""
+    if value is None:
+        return None
+    return value.isoformat()
+
+
+def parse_date_range(value: str) -> Tuple[str | None, str | None]:
+    """Split the date range string into separate ISO formatted dates."""
+    text = normalise_whitespace(value)
+    if not text:
+        return None, None
+    if "-" not in text:
+        single = parse_single_date(text)
+        return format_date(single), format_date(single)
+    start_part, end_part = text.split("-", 1)
+    start_part = normalise_whitespace(start_part)
+    end_part = normalise_whitespace(end_part)
+    if start_part.count(".") == 1 and end_part.count(".") >= 2:
+        year_suffix = end_part.split(".")[-1]
+        start_part = f"{start_part}.{year_suffix}"
+    start_date = parse_single_date(start_part)
+    end_date = parse_single_date(end_part)
+    return format_date(start_date), format_date(end_date)
+
+
+def clean_number(value: str) -> int | float | None:
+    """Convert a textual number into int or float."""
+    text = normalise_whitespace(value)
+    if not text:
+        return None
+    text = text.replace(" ", "").replace(",", ".")
+    if not text or text == "-":
+        return None
+    try:
+        number = float(text)
+    except ValueError:
+        return None
+    if number.is_integer():
+        return int(number)
+    return number
+
+
+def clean_percentage(value: str) -> float | None:
+    """Convert a percentage string into a float without the percent sign."""
+    text = normalise_whitespace(value).replace("%", "")
+    cleaned = clean_number(text)
+    if isinstance(cleaned, int):
+        return float(cleaned)
+    return cleaned
+
+
+def build_records(header: Sequence[str], rows: Iterable[Sequence[str]]) -> List[Dict[str, object]]:
+    """Transform the wide table into a list of cleaned records."""
+    records: List[Dict[str, object]] = []
+    for col_idx in range(1, len(header)):
+        manager_name = normalise_whitespace(header[col_idx])
+        if not manager_name:
+            manager_name = f"Менеджер {col_idx}"
+        record: Dict[str, object] = {key: None for key in FINAL_COLUMNS}
+        record["Менеджер"] = manager_name
+        for row in rows:
+            if not row:
+                continue
+            label = normalise_whitespace(row[0])
+            if not label:
+                continue
+            raw_value = value_at(row, col_idx)
+            if label == DATE_LABEL:
+                start, end = parse_date_range(raw_value)
+                record["Дата початку акції"] = start
+                record["Дата закінчення акції"] = end
+            elif label in FIELD_MAP:
+                record[FIELD_MAP[label]] = normalise_whitespace(raw_value)
+        # Final clean-up for text fields
+        for field in TEXT_COLUMNS:
+            value = record.get(field)
+            record[field] = normalise_whitespace(value) if value else None
+        for field in NUMERIC_COLUMNS:
+            record[field] = clean_number(record.get(field, "") or "")
+        for field in PERCENT_COLUMNS:
+            record[field] = clean_percentage(record.get(field, "") or "")
+        records.append(record)
+    return records
+
+
+def column_letter(index: int) -> str:
+    """Convert a 1-based column index into an Excel column letter."""
+    result = ""
+    current = index
+    while current > 0:
+        current, remainder = divmod(current - 1, 26)
+        result = chr(65 + remainder) + result
+    return result
+
+
+def build_sheet_xml(headers: Sequence[str], records: Sequence[Dict[str, object]]) -> str:
+    """Construct the XML for the worksheet with inline strings."""
+    lines = [
+        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
+        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
+        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
+        "  <sheetData>",
+    ]
+    # Header row
+    lines.append("    <row r=\"1\">")
+    for col_idx, header in enumerate(headers, start=1):
+        cell_ref = f"{column_letter(col_idx)}1"
+        lines.append(
+            f"      <c r=\"{cell_ref}\" t=\"inlineStr\"><is><t>{escape(header)}</t></is></c>"
+        )
+    lines.append("    </row>")
+    # Data rows
+    for row_idx, record in enumerate(records, start=2):
+        lines.append(f"    <row r=\"{row_idx}\">")
+        for col_idx, key in enumerate(headers, start=1):
+            value = record.get(key)
+            if value in (None, ""):
+                continue
+            cell_ref = f"{column_letter(col_idx)}{row_idx}"
+            if isinstance(value, (int, float)):
+                lines.append(f"      <c r=\"{cell_ref}\"><v>{value}</v></c>")
+            else:
+                lines.append(
+                    "      <c r=\"{ref}\" t=\"inlineStr\"><is><t>{text}</t></is></c>".format(
+                        ref=cell_ref, text=escape(str(value))
+                    )
+                )
+        lines.append("    </row>")
+    lines.append("  </sheetData>")
+    lines.append("</worksheet>")
+    return "\n".join(lines)
+
+
+def write_xlsx(path: Path, headers: Sequence[str], records: Sequence[Dict[str, object]]) -> None:
+    """Write the cleaned data to an XLSX file."""
+    sheet_xml = build_sheet_xml(headers, records)
+    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
+        archive.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
+        archive.writestr("_rels/.rels", RELS_XML)
+        archive.writestr("xl/workbook.xml", WORKBOOK_XML)
+        archive.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
+        archive.writestr("xl/styles.xml", STYLES_XML)
+        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
+
+
+def main() -> None:
+    header, rows = read_raw_data(DATA_FILE)
+    records = build_records(header, rows)
+    # Remove records without product information
+    filtered_records = [record for record in records if record.get("Продукт")]
+    OUTPUT_FILE.unlink(missing_ok=True)
+    write_xlsx(OUTPUT_FILE, FINAL_COLUMNS, filtered_records)
+    print(f"Cleaned data saved to {OUTPUT_FILE.name}")
+
+
+if __name__ == "__main__":
+    main()
 
EOF
)
