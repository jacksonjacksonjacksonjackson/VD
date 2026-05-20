"""
scripts/update_template_font.py — One-time utility: set Avenir LT Std Book as the
theme body font in assets/template_default.pptx.

Run once from the project root:
    python3 scripts/update_template_font.py

After running, all python-pptx chart text that inherits from the theme body font
(axis labels, legends, data labels) will use Avenir LT Std Book automatically
without per-element code overrides.
"""

import sys
import os
import zipfile
import shutil
import tempfile
from lxml import etree

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(PROJECT_ROOT, "assets", "template_default.pptx")
TARGET_FONT = "Calibri"

NSMAP = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def _patch_font_scheme(tree: etree._Element, font_name: str) -> bool:
    """Patch <a:minorFont> (body font) in the theme XML tree. Returns True if changed."""
    changed = False
    for fs in tree.findall(".//a:fontScheme", NSMAP):
        minor = fs.find("a:minorFont", NSMAP)
        if minor is None:
            continue
        for tag in ("a:latin", "a:ea", "a:cs"):
            el = minor.find(tag, NSMAP)
            if el is None:
                el = etree.SubElement(
                    minor,
                    etree.QName(NSMAP["a"], tag.split(":")[1]),
                )
            old = el.get("typeface", "")
            if old != font_name:
                el.set("typeface", font_name)
                print(f"  {tag}: '{old}' → '{font_name}'")
                changed = True
            else:
                print(f"  {tag}: already '{font_name}' (no change)")
    return changed


def main() -> None:
    if not os.path.exists(TEMPLATE_PATH):
        print(f"ERROR: Template not found at {TEMPLATE_PATH}")
        sys.exit(1)

    print(f"Opening: {TEMPLATE_PATH}")

    # .pptx files are ZIP archives — edit the theme XML directly in-place
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        changed = False
        with zipfile.ZipFile(TEMPLATE_PATH, "r") as zin, \
             zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():
                data = zin.read(item.filename)

                # Theme files live at ppt/theme/theme*.xml
                if item.filename.startswith("ppt/theme/") and item.filename.endswith(".xml"):
                    print(f"Patching {item.filename} ...")
                    tree = etree.fromstring(data)
                    if _patch_font_scheme(tree, TARGET_FONT):
                        data = etree.tostring(tree, xml_declaration=True,
                                              encoding="UTF-8", standalone=True)
                        changed = True
                    else:
                        print(f"  (no changes needed)")

                zout.writestr(item, data)

        if changed:
            shutil.move(tmp_path, TEMPLATE_PATH)
            print(f"\nSaved: {TEMPLATE_PATH}")
            print("Done. Charts inheriting theme body font will now use Avenir LT Std Book.")
        else:
            os.unlink(tmp_path)
            print("\nNo changes needed — template already uses the target font.")

    except Exception:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise


if __name__ == "__main__":
    main()
