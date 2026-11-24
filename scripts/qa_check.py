# -*- coding: utf-8 -*-
"""Automated QA checklist over all 6 deliverable decks."""
import zipfile, re, os, glob
from lxml import etree
from pptx import Presentation

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
OUT = "/home/user/pr123/deliverables"

DECK_FILES = [
    ("Alatau",   "Alatau_SmartCity_RU.pptx", 17),
    ("Master",   "SOKOL_Master_RU.pptx", 97),
    ("Investor", "SOKOL_Investor_RU.pptx", 51),
    ("Hub100",   "SOKOL_Hub100_RU.pptx", 53),
    ("Park",     "SOKOL_Park_RU.pptx", 53),
    ("Animated", "SOKOL_Animated_RU.pptx", 77),
]

def slide_texts(path):
    """yield (idx, xml_str, visible_text, notes_text)"""
    prs = Presentation(path)
    for i, s in enumerate(prs.slides, start=1):
        xml = s.part.blob.decode("utf-8", "ignore")
        vis = " ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
        notes = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
        yield i, xml, vis, notes

def all_text(path):
    with zipfile.ZipFile(path) as z:
        chunks = []
        for n in z.namelist():
            if n.startswith("ppt/slides/slide") and n.endswith(".xml"):
                chunks.append(z.read(n).decode("utf-8", "ignore"))
        return "\n".join(chunks)

def all_notes(path):
    with zipfile.ZipFile(path) as z:
        chunks = []
        for n in z.namelist():
            if n.startswith("ppt/notesSlides/notesSlide") and n.endswith(".xml"):
                chunks.append(z.read(n).decode("utf-8", "ignore"))
        return "\n".join(chunks)

PASS = "\033[32mPASS\033[0m"; FAIL = "\033[31mFAIL\033[0m"
report = []

def check(name, ok, detail=""):
    report.append((name, ok, detail := name if False else detail))
    print(f"  [{PASS if ok else FAIL}] {name}" + (f" — {detail}" if detail else ""))

grand_ok = True
for key, fn, exp_n in DECK_FILES:
    path = os.path.join(OUT, fn)
    print(f"\n===== {key}: {fn} =====")
    if not os.path.exists(path):
        check(f"{key}: file exists", False, "missing"); grand_ok = False; continue
    prs = Presentation(path)
    # 1. 16:9
    ok = abs(prs.slide_width / prs.slide_height - 16/9) < 0.001
    check(f"{key}: 16:9 ({prs.slide_width}x{prs.slide_height})", ok)
    # 2. slide count
    ok = len(prs.slides._sldIdLst) == exp_n
    check(f"{key}: slide count == {exp_n}", ok, f"got {len(prs.slides._sldIdLst)}")
    # 3. explicit bg on every slide
    no_bg = []
    fonts_bad = set()
    no_notes = []
    no_source_note = []
    with zipfile.ZipFile(path) as z:
        slide_names = sorted([n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)],
                             key=lambda s: int(re.search(r"\d+", s.split("slide")[-1]).group()))
        for n in slide_names:
            xml = z.read(n).decode("utf-8", "ignore")
            if "<p:bg>" not in xml:
                no_bg.append(n)
            for m in re.finditer(r'typeface="([^"+][^"]*)"', xml):
                if m.group(1) != "Segoe UI":
                    fonts_found = m.group(1)
                    fonts_bad.add(fonts_found)
    prs2 = Presentation(path)
    for i, s in enumerate(prs2.slides, start=1):
        if not (s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip()):
            no_notes.append(i)
        else:
            nt = s.notes_slide.notes_text_frame.text
            if "ИСТОЧНИК:" not in nt or "СТАТУС:" not in nt:
                no_source_note.append(i)
    check(f"{key}: explicit <p:bg> on all slides", not no_bg, f"missing in {len(no_bg)}")
    check(f"{key}: all fonts Segoe UI", not fonts_bad, f"bad: {sorted(fonts_bad)[:5]}")
    check(f"{key}: notes on 100% slides", not no_notes, f"missing: {no_notes[:8]}")
    check(f"{key}: source+status in every note", not no_source_note, f"bad: {no_source_note[:8]}")

    txt = all_text(path)
    nts = all_notes(path)
    combined = txt + "\n" + nts

    # 4. Legal greps
    bad1 = re.findall(r"входит в СЭЗ", combined)
    bad2 = re.findall(r"входит в её границы", combined)
    check(f"{key}: zero unqualified «входит в СЭЗ» / «входит в её границы»", not bad1 and not bad2,
          f"hits: {len(bad1)}+{len(bad2)}")
    bad3 = re.findall(r"по данным прокуратуры", combined, re.I)
    check(f"{key}: zero «по данным прокуратуры»", not bad3, f"hits: {len(bad3)}")
    bad4 = re.findall(r"прокуратур", combined, re.I)
    check(f"{key}: zero «прокуратур*» anywhere", not bad4, f"hits: {len(bad4)}")

    # 5. Xiaomi/MG/VOYAH only with disclaimer
    brand_hits = len(re.findall(r"Xiaomi", combined)) + len(re.findall(r"VOYAH", combined))
    disc_hits = len(re.findall(r"не подтверждено", combined)) + len(re.findall(r"не\s+подтверждено", combined))
    if brand_hits:
        check(f"{key}: brand mentions ({brand_hits}) covered by «не подтверждено» ({disc_hits})", disc_hits >= 1)
    else:
        check(f"{key}: no unconfirmed brands", True)

    # 6. Attributions
    if key in ("Master", "Animated"):
        ok = "Tengrinews" in combined and "24KZ" in combined
        check(f"{key}: AutoFlight attribution preserved", ok)
    if key in ("Master", "Hub100"):
        ok = "DGT" in combined and "duisport" in combined
        check(f"{key}: Duisburg DGT/duisport attribution", ok)
    ok = "OpenStreetMap" in combined
    check(f"{key}: OpenStreetMap attribution", ok)

    # 7. TBD placeholders & caveats
    ok = "TBD" in combined
    check(f"{key}: TBD placeholders present", ok)
    ok = "предварительн" in combined.lower()
    check(f"{key}: «предварительно» caveats present", ok)
    ok = "требует согласования" in combined.lower() or "согласовать с КТЖ" in combined
    check(f"{key}: «требует согласования» caveats present", ok)

    # 8. Roadmap dates
    if key in ("Master", "Investor", "Hub100", "Park", "Animated"):
        ok = "2027–2030" in combined and "2030–2035" in combined and "после 2035" in combined
        check(f"{key}: roadmap dates 2027–2030 / 2030–2035 / после 2035", ok)
    if key == "Alatau":
        ok = "2027—2030" in combined or "2027–2030" in combined
        check(f"{key}: Alatau roadmap dates present", ok)
        ok = "286-VIII" in combined
        check(f"{key}: Constitutional Law № 286-VIII", ok)
        ok = "одобрения ERG" in combined or "одобрение ERG" in combined
        check(f"{key}: ERG approval caveat", ok)
        ok = "Alatau City Authority" in combined
        check(f"{key}: Alatau City Authority single window", ok)

    # 9. EN duplicates (title/offer/contacts)
    ok = ("INVEST" in combined or "INDUSTRIAL PARK" in combined)
    check(f"{key}: EN key-slide lines present", ok)
    if key in ("Master", "Investor", "Hub100", "Park", "Animated"):
        ok = "KZ:" in combined or "Инвесторлар" in combined
        check(f"{key}: KZ line in contacts", ok)

    # 10. Registry exists & row count
    reg = os.path.join(OUT, "registry", f"REESTR_{'SOKOL_' if key!='Alatau' else ''}{key}.md")
    if key == "Alatau":
        reg = os.path.join(OUT, "registry", "REESTR_Alatau.md")
    elif key == "Master":
        reg = os.path.join(OUT, "registry", "REESTR_SOKOL_Master.md")
    ok = os.path.exists(reg)
    check(f"{key}: registry exists ({os.path.basename(reg)})", ok)
    if ok:
        with open(reg) as f:
            rows = [l for l in f if l.startswith("| ") and "----" not in l][1:]
        check(f"{key}: registry rows == {exp_n}", len(rows) == exp_n, f"got {len(rows)}")

print("\n" + "="*60)
fails = [r for r in report if not r[1]]
print(f"TOTAL: {len(report)} checks, {len(fails)} failed")
for name, ok, detail in fails:
    print(f"  FAILED: {name} — {detail}")
