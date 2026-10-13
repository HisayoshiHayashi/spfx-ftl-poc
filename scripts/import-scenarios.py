"""Import the scenario workbook and ID-named images without modifying the originals."""

import argparse
import collections
import json
import pathlib
import shutil
import xml.etree.ElementTree as ET
import zipfile

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def read_sheets(workbook):
    with zipfile.ZipFile(workbook) as archive:
        strings = [
            "".join(node.itertext())
            for node in ET.fromstring(archive.read("xl/sharedStrings.xml")).findall("m:si", NS)
        ]
        relationships = {
            item.attrib["Id"]: item.attrib["Target"]
            for item in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        }
        sheets = {}
        for sheet in ET.fromstring(archive.read("xl/workbook.xml")).findall("m:sheets/m:sheet", NS):
            target = relationships[sheet.attrib[REL_NS + "id"]]
            path = target.lstrip("/") if target.startswith("/") else "xl/" + target
            rows = []
            for row in ET.fromstring(archive.read(path)).findall("m:sheetData/m:row", NS):
                values = {}
                for cell in row:
                    value = cell.find("m:v", NS)
                    text = value.text if value is not None else ""
                    if cell.get("t") == "s":
                        text = strings[int(text)]
                    elif cell.get("t") == "inlineStr":
                        text = "".join(cell.find("m:is", NS).itertext())
                    values[cell.attrib["r"].rstrip("0123456789")] = text
                rows.append(values)
            header = rows[0]
            sheets[sheet.attrib["name"]] = [
                {header[column]: value for column, value in row.items() if column in header}
                for row in rows[1:] if row.get("A")
            ]
        return sheets


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=pathlib.Path)
    parser.add_argument("images", type=pathlib.Path)
    args = parser.parse_args()
    sheets = read_sheets(args.workbook)
    scenarios = []
    for row in sheets["未来シナリオ"]:
        scenario_id = row["ID"]
        if not scenario_id.startswith("MS") or not scenario_id[2:].isdigit():
            raise ValueError(f"Invalid scenario ID: {scenario_id}")
        scenarios.append({
            "id": scenario_id,
            "title": row["タイトル"],
            "summary": row["概要"],
            "year": int(float(row["年度"])),
            "categories": [row[key] for key in ["カテゴリ1", "カテゴリ2", "カテゴリ3"] if row.get(key)],
            "before": [row.get(f"変化のポイント前-{i}", "") for i in range(1, 4)],
            "after": [row.get(f"変化のポイント後-{i}", "") for i in range(1, 4)],
            "backgroundTitle": row.get("背景タイトル", ""),
            "background": row.get("背景文章", ""),
            "detailTitle": row.get("詳細タイトル", ""),
            "detail": row.get("詳細文章", ""),
            "sourceClip": row.get("元兆しクリップ", ""),
            "socialData": row.get("参照社会データ", ""),
        })
    ids = {item["id"] for item in scenarios}
    if len(ids) != len(scenarios):
        raise ValueError("Duplicate scenario IDs")
    consumers = []
    for row in sheets["生活者シナリオ"]:
        if row["未来シナリオID"] not in ids:
            raise ValueError(f"Unknown parent ID: {row['未来シナリオID']}")
        consumers.append({
            "id": row["ID"],
            "scenarioId": row["未来シナリオID"],
            "persona": int(float(row["生活者"])),
            "title": row["タイトル"],
            "story": row["生活者シナリオ"],
            "year": int(float(row["年度"])),
        })
    if len({item["id"] for item in consumers}) != len(consumers):
        raise ValueError("Duplicate consumer IDs")
    for item in scenarios:
        image = args.images / f"{item['id']}.jpg"
        if not image.is_file():
            raise FileNotFoundError(image)
    root = pathlib.Path(__file__).resolve().parents[1] / "src/webparts/ftlApp"
    data = root / "data"
    assets = root / "assets/scenarios"
    data.mkdir(parents=True, exist_ok=True)
    assets.mkdir(parents=True, exist_ok=True)
    for item in scenarios:
        shutil.copyfile(args.images / f"{item['id']}.jpg", assets / f"{item['id']}.jpg")
    for filename, records in [("scenarios.json", scenarios), ("consumers.json", consumers)]:
        (data / filename).write_text(json.dumps(records, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    imports = [f"import image{i} from '../assets/scenarios/{item['id']}.jpg';" for i, item in enumerate(scenarios)]
    entries = [f"  {item['id']}: image{i}," for i, item in enumerate(scenarios)]
    (data / "images.ts").write_text(
        "\n".join(imports) + "\n\nexport const scenarioImages: Record<string, string> = {\n"
        + "\n".join(entries) + "\n};\n", encoding="utf-8"
    )
    print(f"Imported {len(scenarios)} scenarios, {len(consumers)} consumer stories and {len(scenarios)} images.")
    print("Years:", dict(collections.Counter(item["year"] for item in scenarios)))


if __name__ == "__main__":
    main()
