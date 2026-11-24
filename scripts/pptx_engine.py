import os, re, io, json, copy, hashlib, zipfile
from lxml import etree
from pptx import Presentation

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

LIGHT_TEXT_HEX = {"FFFFFF", "F8F4ED", "F8F5ED", "F7F3EB", "F5F2EB", "F7F4EB", "E8B66C", "EBC48A", "EDC186", "E9C28E", "F4BD70", "E7C087", "EDC58A"}
DARK_TEXT_HEX  = {"0B202B", "112C38", "0E2841", "000000", "172D48", "142C35", "112933"}

def make_bg_xml(hex_color):
    return etree.fromstring(
        f'<p:bg xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">'
        f'<p:bgPr><a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill><a:effectLst/></p:bgPr>'
        f'</p:bg>'
    )

def detect_slide_bg_color(slide_root, default_dark="0B202B", default_light="F7F5F0"):
    csld = slide_root.find("p:cSld", NS)
    if csld is None:
        return None
    if csld.find("p:bg", NS) is not None:
        return None  # already has explicit bg
    sptree = csld.find("p:spTree", NS)
    if sptree is None:
        return None
    # Check if there is already a full-bleed background shape/pic/video covering >= 12000000 x 6800000
    for el in sptree:
        tag = el.tag.split("}")[-1]
        if tag in ("sp", "pic", "grpSp"):
            xfrm = el.find(".//a:xfrm", NS)
            if xfrm is not None:
                off = xfrm.find("a:off", NS)
                ext = xfrm.find("a:ext", NS)
                if off is not None and ext is not None:
                    x, y = int(off.get("x", "0")), int(off.get("y", "0"))
                    cx, cy = int(ext.get("cx", "0")), int(ext.get("cy", "0"))
                    if x <= 50000 and y <= 50000 and cx >= 12000000 and cy >= 6800000:
                        # check if it has noFill (transparent overlay) or actual fill/blip
                        if tag == "sp" and el.find("p:spPr/a:noFill", NS) is not None:
                            continue
                        # full-bleed opaque visual: inject default_dark behind it (invisible, but explicit)
                        return default_dark
    # Otherwise inspect top-level text colors (not inside dark cards, check header/title shapes near top y < 1600000)
    top_colors = []
    all_colors = []
    for sp in sptree.findall("p:sp", NS):
        xfrm = sp.find("p:spPr/a:xfrm", NS)
        y = int(xfrm.find("a:off", NS).get("y", "9999999")) if (xfrm is not None and xfrm.find("a:off", NS) is not None) else 9999999
        has_solid_box = sp.find("p:spPr/a:solidFill", NS) is not None
        for clr in sp.findall(".//p:txBody//a:rPr//a:srgbClr", NS):
            v = clr.get("val", "").upper()
            all_colors.append(v)
            if y < 1600000 and not has_solid_box:
                top_colors.append(v)
    probe = top_colors if top_colors else all_colors[:6]
    if any(c in DARK_TEXT_HEX for c in probe):
        return default_light
    if any(c in LIGHT_TEXT_HEX for c in probe):
        return default_dark
    return default_dark

def set_explicit_bg(slide_root, hex_color):
    csld = slide_root.find("p:cSld", NS)
    old_bg = csld.find("p:bg", NS)
    if old_bg is not None:
        csld.remove(old_bg)
    csld.insert(0, make_bg_xml(hex_color))

def normalize_fonts(slide_root, font_family="Segoe UI"):
    for tag in ("latin", "ea", "cs"):
        for el in slide_root.findall(f".//a:{tag}", NS):
            tf = el.get("typeface", "")
            if tf and not tf.startswith("+"):
                el.set("typeface", font_family)

def make_notes_xml(notes_text):
    # Clean control chars except newline and tab
    clean_lines = [line for line in notes_text.replace("\r\n", "\n").split("\n")]
    paras = []
    for line in clean_lines:
        esc = (line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        paras.append(
            f'<a:p><a:r><a:rPr lang="ru-RU" sz="1100" dirty="0">'
            f'<a:latin typeface="Segoe UI"/></a:rPr><a:t>{esc}</a:t></a:r></a:p>'
        )
    p_xml = "".join(paras) if paras else '<a:p><a:endParaRPr lang="ru-RU"/></a:p>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
        f'<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{p_xml}</p:txBody></p:sp>'
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
    ).encode("utf-8")

def make_textbox_sp(sp_id, name, x, y, cx, cy, paragraphs, fill_hex=None, border_hex=None, border_w=12700, radius=False, anchor="t", pad=(91440, 55000, 91440, 55000)):
    """
    paragraphs: list of dicts or tuples:
      {"text": str, "sz": int(pt*100), "bold": bool, "color": hex, "align": "l"|"ctr"|"r", "spcBef": int, "spcAft": int}
    """
    geom = 'roundRect' if radius else 'rect'
    av = '<a:gd name="adj" fmla="val 6000"/>' if radius else ''
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill>' if fill_hex else '<a:noFill/>'
    ln_xml = f'<a:ln w="{border_w}"><a:solidFill><a:srgbClr val="{border_hex}"/></a:solidFill></a:ln>' if border_hex else '<a:ln><a:noFill/></a:ln>'
    lIns, tIns, rIns, bIns = pad
    p_list = []
    for p in paragraphs:
        txt = (p.get("text", "")
               .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        sz = p.get("sz", 1400)
        b = ' b="1"' if p.get("bold", False) else ''
        col = p.get("color", "F8F4ED")
        algn = p.get("align", "l")
        sb = p.get("spcBef", 0)
        sa = p.get("spcAft", 0)
        lnspc = p.get("lnSpc", None)
        lnspc_xml = f'<a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc>' if lnspc else ''
        # support multi-line via \n -> <a:br/>
        parts = txt.split("\n")
        runs = []
        for idx_p, part in enumerate(parts):
            if idx_p > 0:
                runs.append(f'<a:br><a:rPr lang="ru-RU" sz="{sz}"{b}/></a:br>')
            runs.append(
                f'<a:r><a:rPr lang="ru-RU" sz="{sz}"{b} dirty="0">'
                f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
                f'<a:latin typeface="Segoe UI"/><a:ea typeface="Segoe UI"/><a:cs typeface="Segoe UI"/>'
                f'</a:rPr><a:t>{part}</a:t></a:r>'
            )
        p_list.append(
            f'<a:p><a:pPr algn="{algn}">{lnspc_xml}'
            f'<a:spcBef><a:spcPts val="{sb}"/></a:spcBef>'
            f'<a:spcAft><a:spcPts val="{sa}"/></a:spcAft></a:pPr>'
            f'{"".join(runs)}</a:p>'
        )
    xml_str = (
        f'<p:sp xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}">'
        f'<p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst>{av}</a:avLst></a:prstGeom>{fill_xml}{ln_xml}</p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" lIns="{lIns}" tIns="{tIns}" rIns="{rIns}" bIns="{bIns}" anchor="{anchor}"><a:noAutofit/></a:bodyPr>'
        f'<a:lstStyle/>{"".join(p_list)}</p:txBody></p:sp>'
    )
    return etree.fromstring(xml_str)

def next_sp_id(slide_root):
    ids = [int(el.get("id", "1")) for el in slide_root.findall(".//p:cNvPr", NS) if el.get("id", "").isdigit()]
    return (max(ids) + 1) if ids else 100

def replace_text_in_slide(slide_root, replacements):
    """
    replacements: list of (old_substring, new_substring) applied to individual <a:t> or paragraph text.
    First tries direct <a:t> substring replacement; if multi-run paragraph contains old_substring across runs, merges runs.
    """
    for old_s, new_s in replacements:
        # 1. check individual <a:t>
        hit = False
        for t_el in slide_root.findall(".//a:t", NS):
            if t_el.text and old_s in t_el.text:
                t_el.text = t_el.text.replace(old_s, new_s)
                hit = True
        if not hit:
            # 2. check paragraph level
            for p_el in slide_root.findall(".//a:p", NS):
                runs = p_el.findall("a:r", NS)
                if not runs:
                    continue
                full = "".join((r.find("a:t", NS).text or "") for r in runs if r.find("a:t", NS) is not None)
                if old_s in full:
                    updated = full.replace(old_s, new_s)
                    first_t = runs[0].find("a:t", NS)
                    if first_t is not None:
                        first_t.text = updated
                    for extra_r in runs[1:]:
                        p_el.remove(extra_r)

def strip_video_elements(slide_root):
    """Remove video/audio playback markup from a slide, keeping the poster image (a:blip) as a static visual.
    Returns the number of video/audio pic elements converted."""
    count = 0
    # 1) pics with videoFile/audioFile: remove media refs + media extLst + media hyperlink
    for nvPr in slide_root.findall(".//p:pic/p:nvPicPr/p:nvPr", NS):
        has_media = False
        for tag in ("videoFile", "audioFile"):
            for vf in nvPr.findall(f"a:{tag}", NS):
                nvPr.remove(vf)
                has_media = True
        for extLst in nvPr.findall("p:extLst", NS):
            if b"media" in etree.tostring(extLst):
                nvPr.remove(extLst)
                has_media = True
        if has_media:
            count += 1
    # remove ppaction://media hyperlinks on pic cNvPr
    for cNvPr in slide_root.findall(".//p:pic/p:nvPicPr/p:cNvPr", NS):
        for hl in cNvPr.findall("a:hlinkClick", NS):
            if hl.get("action", "") == "ppaction://media":
                cNvPr.remove(hl)
    # 2) timing tree: remove p:video / p:audio media nodes
    timing = slide_root.find("p:timing", NS)
    if timing is not None:
        for parent in timing.iter():
            for child in list(parent):
                local = child.tag.split("}")[-1]
                if local in ("video", "audio"):
                    parent.remove(child)
    return count

class DeckAssembler:
    """
    Assembles a valid OOXML .pptx package from a template .pptx (for theme, slideMaster, slideLayouts, notesMaster)
    and a list of slide specs (either cloned from source .pptx files or newly constructed XML slides).
    """
    def __init__(self, template_pptx_path, default_dark="0B202B", default_light="F7F5F0", layout_target="../slideLayouts/slideLayout7.xml", strip_video=False):
        self.template_pptx_path = template_pptx_path
        self.default_dark = default_dark
        self.default_light = default_light
        self.layout_target = layout_target
        self.strip_video = strip_video
        self.stripped_videos = 0   # stats counter
        self.slides = []       # list of dicts: {xml_bytes, rels_list, notes_text}
        self.media_blobs = {}  # md5 -> (partname_in_ppt_media, bytes, ext)
        self.media_counter = 1

    def _register_media(self, blob, ext):
        ext = ext.lower().lstrip(".")
        if ext == "jpg":
            ext = "jpeg"
        h = hashlib.md5(blob).hexdigest()
        if h in self.media_blobs:
            return self.media_blobs[h][0]
        fname = f"media_{self.media_counter:04d}.{ext}"
        self.media_counter += 1
        self.media_blobs[h] = (fname, blob, ext)
        return fname

    def add_cloned_slide(self, src_pptx_path, local_slide_idx, notes_text="", mutate_fn=None):
        with zipfile.ZipFile(src_pptx_path, "r") as z:
            s_num = local_slide_idx + 1
            slide_xml_bytes = z.read(f"ppt/slides/slide{s_num}.xml")
            rels_xml_bytes = z.read(f"ppt/slides/_rels/slide{s_num}.xml.rels")
            slide_root = etree.fromstring(slide_xml_bytes)
            rels_root = etree.fromstring(rels_xml_bytes)

            new_rels = []
            # Keep exact rIds for media/image/video/hyperlink, drop tags & old notesSlide, normalize slideLayout
            has_layout = False
            for rel in rels_root.findall("rel:Relationship", NS):
                rid = rel.get("Id")
                rtype = rel.get("Type")
                target = rel.get("Target")
                tmode = rel.get("TargetMode")
                short_type = rtype.split("/")[-1]
                if short_type in ("tags", "notesSlide"):
                    continue
                elif short_type == "slideLayout":
                    new_rels.append({"Id": rid, "Type": rtype, "Target": self.layout_target})
                    has_layout = True
                elif tmode == "External":
                    new_rels.append({"Id": rid, "Type": rtype, "Target": target, "TargetMode": "External"})
                elif self.strip_video and short_type in ("video", "media"):
                    # drop video / p14:media relationship entirely (poster image rel stays)
                    continue
                else:
                    # internal media or image or video
                    # resolve path relative to ppt/slides/
                    norm_path = os.path.normpath(os.path.join("ppt/slides", target)).replace("\\", "/")
                    blob = z.read(norm_path)
                    ext = os.path.splitext(norm_path)[1]
                    new_media_fname = self._register_media(blob, ext)
                    new_rels.append({"Id": rid, "Type": rtype, "Target": f"../media/{new_media_fname}"})

            if not has_layout:
                new_rels.append({
                    "Id": "rId99",
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                    "Target": self.layout_target
                })

            # Strip video playback elements -> keep poster frame (a:blip) as static visual
            if self.strip_video:
                self.stripped_videos += strip_video_elements(slide_root)

            # Fix implicit dark/light background BEFORE mutation
            bg_hex = detect_slide_bg_color(slide_root, self.default_dark, self.default_light)
            if bg_hex:
                set_explicit_bg(slide_root, bg_hex)

            normalize_fonts(slide_root, "Segoe UI")

            if mutate_fn is not None:
                mutate_fn(slide_root)

            self.slides.append({
                "root": slide_root,
                "rels": new_rels,
                "notes": notes_text,
            })

    def add_custom_slide(self, bg_hex, build_fn, notes_text=""):
        slide_xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<p:sld xmlns:a="{NS["a"]}" xmlns:r="{NS["r"]}" xmlns:p="{NS["p"]}">'
            f'<p:cSld>'
            f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg_hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
            f'<p:spTree>'
            f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            f'</p:spTree>'
            f'</p:cSld>'
            f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            f'</p:sld>'
        )
        slide_root = etree.fromstring(slide_xml.encode("utf-8"))
        build_fn(slide_root)
        new_rels = [{
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
            "Target": self.layout_target
        }]
        self.slides.append({
            "root": slide_root,
            "rels": new_rels,
            "notes": notes_text,
        })

    def save(self, out_pptx_path):
        os.makedirs(os.path.dirname(os.path.abspath(out_pptx_path)), exist_ok=True)
        with zipfile.ZipFile(self.template_pptx_path, "r") as zin, \
             zipfile.ZipFile(out_pptx_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:

            # 1. Copy base template parts (excluding slides, notesSlides, tags, media, presentation.xml, presentation.xml.rels, [Content_Types].xml)
            skip_prefixes = ("ppt/slides/", "ppt/notesSlides/", "ppt/tags/", "ppt/media/")
            skip_exact = {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml"}
            copied_parts = set()
            for item in zin.infolist():
                fn = item.filename
                if fn in skip_exact or any(fn.startswith(p) for p in skip_prefixes):
                    continue
                zout.writestr(fn, zin.read(fn))
                copied_parts.add(fn)

            # 2. Write deduplicated media
            ext_used = set()
            for _, (fname, blob, ext) in self.media_blobs.items():
                zout.writestr(f"ppt/media/{fname}", blob)
                ext_used.add(ext)

            # 3. Write slides, slide rels, notesSlides, notesSlide rels
            pres_root = etree.fromstring(zin.read("ppt/presentation.xml"))
            sldIdLst = pres_root.find("p:sldIdLst", NS)
            for child in list(sldIdLst):
                sldIdLst.remove(child)

            # Build presentation.xml.rels cleanly
            pres_rels_old = etree.fromstring(zin.read("ppt/_rels/presentation.xml.rels"))
            kept_pres_rels = []
            max_rid_num = 1
            for rel in pres_rels_old.findall("rel:Relationship", NS):
                rtype = rel.get("Type")
                if rtype.endswith("/slide"):
                    continue
                rid = rel.get("Id")
                kept_pres_rels.append((rid, rtype, rel.get("Target")))
                m = re.match(r"rId(\d+)", rid)
                if m:
                    max_rid_num = max(max_rid_num, int(m.group(1)))

            slide_pres_rels = []
            for i, sdata in enumerate(self.slides):
                s_idx = i + 1
                pres_rid = f"rId{max_rid_num + s_idx}"
                slide_pres_rels.append((pres_rid, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", f"slides/slide{s_idx}.xml"))
                sld_el = etree.SubElement(sldIdLst, f"{{{NS['p']}}}sldId")
                sld_el.set("id", str(255 + s_idx))
                sld_el.set(f"{{{NS['r']}}}id", pres_rid)

                # Add notesSlide relationship to slide rels
                used_rids = {r["Id"] for r in sdata["rels"]}
                n_rid_num = 1
                while f"rId{n_rid_num}" in used_rids:
                    n_rid_num += 1
                notes_rid = f"rId{n_rid_num}"
                s_rels = list(sdata["rels"]) + [{
                    "Id": notes_rid,
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
                    "Target": f"../notesSlides/notesSlide{s_idx}.xml"
                }]

                # Write slide XML
                zout.writestr(
                    f"ppt/slides/slide{s_idx}.xml",
                    etree.tostring(sdata["root"], xml_declaration=True, encoding="UTF-8", standalone="yes")
                )

                # Write slide rels XML
                rels_lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                              f'<Relationships xmlns="{NS["rel"]}">']
                for r in s_rels:
                    tm = f' TargetMode="{r["TargetMode"]}"' if "TargetMode" in r else ""
                    t_esc = r["Target"].replace("&", "&amp;")
                    rels_lines.append(f'<Relationship Id="{r["Id"]}" Type="{r["Type"]}" Target="{t_esc}"{tm}/>')
                rels_lines.append('</Relationships>')
                zout.writestr(f"ppt/slides/_rels/slide{s_idx}.xml.rels", "".join(rels_lines).encode("utf-8"))

                # Write notesSlide XML & rels
                zout.writestr(f"ppt/notesSlides/notesSlide{s_idx}.xml", make_notes_xml(sdata["notes"] or ""))
                notes_rels_xml = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    f'<Relationships xmlns="{NS["rel"]}">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>'
                    f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide{s_idx}.xml"/>'
                    '</Relationships>'
                ).encode("utf-8")
                zout.writestr(f"ppt/notesSlides/_rels/notesSlide{s_idx}.xml.rels", notes_rels_xml)

            # Write presentation.xml & presentation.xml.rels
            zout.writestr("ppt/presentation.xml", etree.tostring(pres_root, xml_declaration=True, encoding="UTF-8", standalone="yes"))
            p_rels_lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                            f'<Relationships xmlns="{NS["rel"]}">']
            for rid, rtype, targ in kept_pres_rels + slide_pres_rels:
                p_rels_lines.append(f'<Relationship Id="{rid}" Type="{rtype}" Target="{targ}"/>')
            p_rels_lines.append('</Relationships>')
            zout.writestr("ppt/_rels/presentation.xml.rels", "".join(p_rels_lines).encode("utf-8"))

            # Build [Content_Types].xml
            ct_old = etree.fromstring(zin.read("[Content_Types].xml"))
            ct_new = etree.Element(f"{{{NS['ct']}}}Types", nsmap={None: NS["ct"]})
            seen_defaults = set()
            default_mime_map = {
                "rels": "application/vnd.openxmlformats-package.relationships+xml",
                "xml": "application/xml",
                "jpeg": "image/jpeg",
                "jpg": "image/jpeg",
                "png": "image/png",
                "gif": "image/gif",
                "svg": "image/svg+xml",
                "mp4": "video/mp4",
                "wdp": "image/vnd.ms-photo",
            }
            for d in ct_old.findall("ct:Default", NS):
                ext = d.get("Extension")
                seen_defaults.add(ext)
                el = etree.SubElement(ct_new, f"{{{NS['ct']}}}Default")
                el.set("Extension", ext)
                el.set("ContentType", d.get("ContentType"))
            for ext in sorted(ext_used | {"jpeg", "png", "mp4"}):
                if ext not in seen_defaults:
                    el = etree.SubElement(ct_new, f"{{{NS['ct']}}}Default")
                    el.set("Extension", ext)
                    el.set("ContentType", default_mime_map.get(ext, "application/octet-stream"))
                    seen_defaults.add(ext)

            # Keep non-slide/non-notes/non-tag overrides that actually exist in zout
            for ov in ct_old.findall("ct:Override", NS):
                pn = ov.get("PartName")
                rel_pn = pn.lstrip("/")
                if rel_pn in copied_parts or rel_pn == "ppt/presentation.xml":
                    el = etree.SubElement(ct_new, f"{{{NS['ct']}}}Override")
                    el.set("PartName", pn)
                    el.set("ContentType", ov.get("ContentType"))

            for i in range(1, len(self.slides) + 1):
                el_s = etree.SubElement(ct_new, f"{{{NS['ct']}}}Override")
                el_s.set("PartName", f"/ppt/slides/slide{i}.xml")
                el_s.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
                el_n = etree.SubElement(ct_new, f"{{{NS['ct']}}}Override")
                el_n.set("PartName", f"/ppt/notesSlides/notesSlide{i}.xml")
                el_n.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")

            zout.writestr("[Content_Types].xml", etree.tostring(ct_new, xml_declaration=True, encoding="UTF-8", standalone="yes"))

        # Validate with python-pptx immediately
        prs_check = Presentation(out_pptx_path)
        assert len(prs_check.slides) == len(self.slides), f"Slide count mismatch: {len(prs_check.slides)} != {len(self.slides)}"
        return len(prs_check.slides), os.path.getsize(out_pptx_path)

print("pptx_engine.py written successfully.")
