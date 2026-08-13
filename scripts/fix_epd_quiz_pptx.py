#!/usr/bin/env python3
"""Fix quiz PPTX: answer delay (+10s / robust timing) + followed-hyperlink red cells.

Preserves user layout/text/shapes; only patches animation timing, theme colors,
and board-cell hyperlink wiring.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "Квиз_1С-ЭПД_Доки_Логистика ИИ ПОПРАВИТЬ.pptx"
OUT = ROOT / "Квиз_1С-ЭПД_Доки_Логистика ИИ ПОПРАВИТЬ.pptx"
OUT_COPY = ROOT / "квиз 1с-эпд" / "Квиз_1С-ЭПД_Доки_Логистика_100к1.pptx"

# Answer currently appears ~10s early vs the 30s timer → wait 40s
ANSWER_DELAY_MS = 40000

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
A = NS["a"]
P = NS["p"]
R = NS["r"]


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def timing_xml(spid: str, delay_ms: int) -> etree._Element:
    """PowerPoint-compatible Appear after delay (with bldLst hide-until-animate)."""
    xml = f"""
    <p:timing xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="3" fill="hold" nodeType="clickPar">
                        <p:stCondLst>
                          <p:cond delay="indefinite"/>
                          <p:cond evt="onBegin" delay="0">
                            <p:tn val="2"/>
                          </p:cond>
                        </p:stCondLst>
                        <p:childTnLst>
                          <p:par>
                            <p:cTn id="4" fill="hold" nodeType="afterGroup">
                              <p:stCondLst>
                                <p:cond delay="0"/>
                              </p:stCondLst>
                              <p:childTnLst>
                                <p:par>
                                  <p:cTn id="5" presetID="1" presetClass="entr" presetSubtype="0"
                                         fill="hold" grpId="0" nodeType="afterEffect">
                                    <p:stCondLst>
                                      <p:cond delay="{delay_ms}"/>
                                    </p:stCondLst>
                                    <p:childTnLst>
                                      <p:set>
                                        <p:cBhvr>
                                          <p:cTn id="6" dur="1" fill="hold">
                                            <p:stCondLst>
                                              <p:cond delay="0"/>
                                            </p:stCondLst>
                                          </p:cTn>
                                          <p:tgtEl>
                                            <p:spTgt spid="{spid}"/>
                                          </p:tgtEl>
                                          <p:attrNameLst>
                                            <p:attrName>style.visibility</p:attrName>
                                          </p:attrNameLst>
                                        </p:cBhvr>
                                        <p:to>
                                          <p:strVal val="visible"/>
                                        </p:to>
                                      </p:set>
                                    </p:childTnLst>
                                  </p:cTn>
                                </p:par>
                              </p:childTnLst>
                            </p:cTn>
                          </p:par>
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst>
                  <p:cond evt="onPrev" delay="0">
                    <p:tgtEl><p:sldTgt/></p:tgtEl>
                  </p:cond>
                </p:prevCondLst>
                <p:nextCondLst>
                  <p:cond evt="onNext" delay="0">
                    <p:tgtEl><p:sldTgt/></p:tgtEl>
                  </p:cond>
                </p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
      <p:bldLst>
        <p:bldP spid="{spid}" grpId="0" animBg="1"/>
      </p:bldLst>
    </p:timing>
    """
    return etree.fromstring(xml)


def find_answer_spid(slide_root: etree._Element) -> str | None:
    for sp in slide_root.xpath(".//p:sp", namespaces=NS):
        texts = "".join(t.text or "" for t in sp.xpath(".//a:t", namespaces=NS))
        if texts.startswith("Ответ") and "появится" not in texts:
            c_nv = sp.find("./p:nvSpPr/p:cNvPr", NS)
            if c_nv is not None:
                return c_nv.get("id")
    # fallback: existing animation target
    tgts = slide_root.xpath(".//p:spTgt", namespaces=NS)
    return tgts[0].get("spid") if tgts else None


def fix_question_slide(slide_xml: bytes) -> bytes:
    root = etree.fromstring(slide_xml)
    spid = find_answer_spid(root)
    if not spid:
        return slide_xml

    for old in root.findall(qn("p", "timing")):
        root.remove(old)

    timing = timing_xml(spid, ANSWER_DELAY_MS)
    # Insert before extLst if present, else append
    ext = root.find(qn("p", "extLst"))
    if ext is not None:
        ext.addprevious(timing)
    else:
        root.append(timing)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def patch_theme_followed_hyperlink(theme_xml: bytes) -> bytes:
    """Unused links = white score text; opened (followed) = red."""
    root = etree.fromstring(theme_xml)
    changed = False
    for tag, color in (("hlink", "FFFFFF"), ("folHlink", "C0392B")):
        el = root.find(f".//{{{A}}}{tag}")
        if el is None:
            continue
        for child in list(el):
            el.remove(child)
        srgb = etree.SubElement(el, qn("a", "srgbClr"))
        srgb.set("val", color)
        changed = True
    if not changed:
        return theme_xml
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _set_run_hyperlink_color(r_pr: etree._Element, r_id: str, action: str | None) -> None:
    # Remove old solidFill / hlinkClick
    for child in list(r_pr):
        if child.tag in (qn("a", "solidFill"), qn("a", "hlinkClick")):
            r_pr.remove(child)
    # No underline — only color should change when the link is followed
    r_pr.set("u", "none")
    solid = etree.SubElement(r_pr, qn("a", "solidFill"))
    scheme = etree.SubElement(solid, qn("a", "schemeClr"))
    scheme.set("val", "hlink")
    hl = etree.SubElement(r_pr, qn("a", "hlinkClick"))
    hl.set(qn("r", "id"), r_id)
    if action:
        hl.set("action", action)


def fix_board_followed_hyperlinks(slide_xml: bytes) -> bytes:
    """Put hyperlinks on score text runs with scheme hlink color (followed → red)."""
    root = etree.fromstring(slide_xml)
    score_values = {"10", "30", "50", "70", "100"}

    for sp in root.xpath(".//p:sp", namespaces=NS):
        c_nv = sp.find("./p:nvSpPr/p:cNvPr", NS)
        if c_nv is None:
            continue
        shape_hl = c_nv.find(qn("a", "hlinkClick"))
        if shape_hl is None:
            continue

        texts = [t.text or "" for t in sp.xpath(".//a:t", namespaces=NS)]
        joined = "".join(texts).strip()
        if joined not in score_values:
            continue

        r_id = shape_hl.get(qn("r", "id"))
        action = shape_hl.get("action")
        if not r_id:
            continue

        # Ensure every text run in the score cell uses hyperlink scheme color
        for r in sp.xpath(".//a:r", namespaces=NS):
            r_pr = r.find(qn("a", "rPr"))
            if r_pr is None:
                r_pr = etree.Element(qn("a", "rPr"))
                r.insert(0, r_pr)
            _set_run_hyperlink_color(r_pr, r_id, action)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def process(src: Path, dst: Path) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        buf, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            name = info.filename

            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                # question slides are slide3.xml .. slide27.xml
                base = Path(name).stem  # slideN
                try:
                    num = int(base.replace("slide", ""))
                except ValueError:
                    num = -1
                if num == 2:
                    data = fix_board_followed_hyperlinks(data)
                elif 3 <= num <= 27:
                    data = fix_question_slide(data)
            elif name == "ppt/theme/theme1.xml":
                data = patch_theme_followed_hyperlink(data)

            zout.writestr(info, data)

    dst.write_bytes(buf.getvalue())
    print(f"Saved: {dst} ({dst.stat().st_size} bytes)")
    print(f"Answer delay: {ANSWER_DELAY_MS} ms (was 30000; +10s to match timer)")
    print("Board cells: score text uses theme hyperlink color; opened → red folHlink")


def main() -> None:
    if not SRC.exists():
        raise SystemExit(f"Not found: {SRC}")
    process(SRC, OUT)
    OUT_COPY.parent.mkdir(parents=True, exist_ok=True)
    OUT_COPY.write_bytes(OUT.read_bytes())
    print(f"Copy:  {OUT_COPY}")


if __name__ == "__main__":
    main()
