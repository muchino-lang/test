#!/usr/bin/env python3
"""template.xlsx の書式パーツを src/index.src.html に埋め込んで index.html を生成する。

テンプレート（書式・フォント・罫線・列幅）を変えたいときは template.xlsx を
Excel で編集し直して、このスクリプトを実行するだけでツールに反映される。

    python3 tools/quote-xlsx/build.py
"""
import json
import re
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
TEMPLATE = HERE / "template.xlsx"
SRC = HERE / "src" / "index.src.html"
OUT = HERE / "index.html"

# シートは JS 側で組み立てるので、共有文字列と計算チェーンは同梱しない
PARTS = {
    "contentTypes": "[Content_Types].xml",
    "rels": "_rels/.rels",
    "app": "docProps/app.xml",
    "core": "docProps/core.xml",
    "workbook": "xl/workbook.xml",
    "workbookRels": "xl/_rels/workbook.xml.rels",
    "styles": "xl/styles.xml",
    "theme": "xl/theme/theme1.xml",
}


def main() -> None:
    with zipfile.ZipFile(TEMPLATE) as z:
        parts = {key: z.read(name).decode("utf-8") for key, name in PARTS.items()}

    # sharedStrings / calcChain への参照を外す
    parts["contentTypes"] = re.sub(
        r'<Override PartName="/xl/(sharedStrings|calcChain)\.xml"[^>]*/>', "", parts["contentTypes"]
    )
    parts["workbookRels"] = re.sub(
        r'<Relationship [^>]*Target="(sharedStrings|calcChain)\.xml"[^>]*/>', "", parts["workbookRels"]
    )

    # 開いた時点で数式を再計算させる
    parts["workbook"] = parts["workbook"].replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ')
    # 作成環境のローカルパスは持ち回らない
    parts["workbook"] = re.sub(
        r"<mc:AlternateContent .*?</mc:AlternateContent>", "", parts["workbook"], flags=re.S
    )

    for key, xml in parts.items():
        assert xml.lstrip().startswith("<?xml"), key

    src = SRC.read_text(encoding="utf-8")
    marker = "/*__TEMPLATE_PARTS__*/{}"
    assert marker in src, "src/index.src.html に埋め込みマーカーがありません"
    OUT.write_text(src.replace(marker, json.dumps(parts, ensure_ascii=False)), encoding="utf-8")
    print(f"wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
