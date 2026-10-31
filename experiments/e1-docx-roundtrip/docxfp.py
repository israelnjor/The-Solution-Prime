"""Structural fingerprint of a DOCX file, used to measure round-trip fidelity.

Reads the OOXML directly (no python-docx), so it sees exactly what a candidate
wrote. Style inheritance is resolved, so formatting that a candidate moves
between styles and direct formatting still compares as equal.
"""
import hashlib
import posixpath
import re
import unicodedata
import zipfile
from collections import Counter
import xml.etree.ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}
MARKER_RE = re.compile(r"\[[A-Z]-\d{3}\]")
# Effective character properties compared per character, in this order.
PROPS = ("b", "i", "u", "strike", "color", "sz", "font", "highlight", "vertAlign")
DEFAULT_RUN = {"b": False, "i": False, "u": None, "strike": False, "color": None,
               "sz": 10.0, "font": None, "highlight": None, "vertAlign": None}
# Field results that are computed at render time and are not document content.
COMPUTED_FIELDS = {"PAGE", "NUMPAGES", "SECTIONPAGES", "DATE", "TIME", "TOC", "REF", "PAGEREF", "SEQ"}


def q(tag):
    prefix, name = tag.split(":")
    return f"{{{NS[prefix]}}}{name}"


def wattr(el, name):
    return None if el is None else el.get(q("w:" + name))


def onoff(el):
    if el is None:
        return None
    return wattr(el, "val") not in ("0", "false", "off")


class Package:
    def __init__(self, path):
        self.zip = zipfile.ZipFile(path)
        self.names = set(self.zip.namelist())
        self._rels = {}
        root_rels = self.rels("")
        main = next((t for typ, t, ext in root_rels.values() if typ.endswith("/officeDocument")), None)
        if main is None:
            raise ValueError("no officeDocument relationship; not a WordprocessingML package")
        self.main = main

    def read(self, name):
        return self.zip.read(name) if name in self.names else None

    def xml(self, name):
        data = self.read(name)
        return None if data is None else ET.fromstring(data)

    def rels(self, part):
        """Map relationship id -> (type, resolved target, is_external) for a part."""
        if part in self._rels:
            return self._rels[part]
        base = posixpath.dirname(part)
        rels_name = posixpath.join(base, "_rels", posixpath.basename(part) + ".rels") if part else "_rels/.rels"
        out = {}
        root = self.xml(rels_name)
        if root is not None:
            for rel in root.findall(q("pr:Relationship")):
                target = rel.get("Target", "")
                external = rel.get("TargetMode") == "External"
                if not external:
                    target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
                out[rel.get("Id")] = (rel.get("Type", ""), target, external)
        self._rels[part] = out
        return out

    def related(self, part, type_suffix):
        return next((t for typ, t, ext in self.rels(part).values() if typ.endswith(type_suffix)), None)


class Styles:
    def __init__(self, pkg):
        self.styles = {}
        self.default_para = None
        self.doc_rpr = None
        self.doc_ppr = None
        self.theme_fonts = {}
        styles_part = pkg.related(pkg.main, "/styles")
        root = pkg.xml(styles_part) if styles_part else None
        if root is not None:
            defaults = root.find(q("w:docDefaults"))
            if defaults is not None:
                self.doc_rpr = defaults.find(f"{q('w:rPrDefault')}/{q('w:rPr')}")
                self.doc_ppr = defaults.find(f"{q('w:pPrDefault')}/{q('w:pPr')}")
            for st in root.findall(q("w:style")):
                sid = wattr(st, "styleId")
                name_el = st.find(q("w:name"))
                based = st.find(q("w:basedOn"))
                self.styles[sid] = {
                    "type": wattr(st, "type"),
                    "name": (wattr(name_el, "val") or sid or "").lower(),
                    "based": wattr(based, "val"),
                    "pPr": st.find(q("w:pPr")),
                    "rPr": st.find(q("w:rPr")),
                }
                if wattr(st, "type") == "paragraph" and wattr(st, "default") in ("1", "true"):
                    self.default_para = sid
        theme_part = pkg.related(pkg.main, "/theme")
        theme = pkg.xml(theme_part) if theme_part else None
        if theme is not None:
            for kind in ("major", "minor"):
                latin = theme.find(f".//{q('a:' + kind + 'Font')}/{q('a:latin')}")
                if latin is not None:
                    self.theme_fonts[kind] = latin.get("typeface")

    def chain(self, sid):
        """Styles from the root of the basedOn chain down to sid."""
        out, seen = [], set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            out.append(self.styles[sid])
            sid = self.styles[sid]["based"]
        return out[::-1]

    def name(self, sid):
        return self.styles[sid]["name"] if sid in self.styles else (sid or "").lower()

    def run_props(self, rpr):
        """Partial property dict from a w:rPr element."""
        d = {}
        if rpr is None:
            return d
        for key in ("b", "i", "strike"):
            v = onoff(rpr.find(q("w:" + key)))
            if v is not None:
                d[key] = v
        u = rpr.find(q("w:u"))
        if u is not None:
            d["u"] = None if wattr(u, "val") == "none" else (wattr(u, "val") or "single")
        color = rpr.find(q("w:color"))
        if color is not None:
            val = (wattr(color, "val") or "auto").upper()
            d["color"] = None if val in ("AUTO", "000000") else val
        sz = rpr.find(q("w:sz"))
        if sz is not None and wattr(sz, "val"):
            d["sz"] = int(wattr(sz, "val")) / 2
        fonts = rpr.find(q("w:rFonts"))
        if fonts is not None:
            font = wattr(fonts, "ascii") or wattr(fonts, "hAnsi")
            theme = wattr(fonts, "asciiTheme") or wattr(fonts, "hAnsiTheme")
            if theme and not font:
                font = self.theme_fonts.get("major" if theme.startswith("major") else "minor")
            if font:
                d["font"] = font
        hl = rpr.find(q("w:highlight"))
        if hl is not None:
            d["highlight"] = None if wattr(hl, "val") == "none" else wattr(hl, "val")
        va = rpr.find(q("w:vertAlign"))
        if va is not None:
            d["vertAlign"] = None if wattr(va, "val") == "baseline" else wattr(va, "val")
        return d

    def effective_run(self, para_chain, rstyle, rpr):
        props = dict(DEFAULT_RUN)
        props.update(self.run_props(self.doc_rpr))
        for st in para_chain:
            props.update(self.run_props(st["rPr"]))
        for st in self.chain(rstyle):
            props.update(self.run_props(st["rPr"]))
        props.update(self.run_props(rpr))
        return tuple(props[k] for k in PROPS)

    def para_prop(self, para_chain, ppr, path):
        """Last value of a pPr child element along docDefaults -> styles -> direct."""
        found = None
        for src in [self.doc_ppr] + [st["pPr"] for st in para_chain] + [ppr]:
            if src is not None:
                el = src.find(path)
                if el is not None:
                    found = el
        return found


class Numbering:
    def __init__(self, pkg):
        self.nums, self.abstract = {}, {}
        part = pkg.related(pkg.main, "/numbering")
        root = pkg.xml(part) if part else None
        if root is None:
            return
        for an in root.findall(q("w:abstractNum")):
            levels = {}
            for lvl in an.findall(q("w:lvl")):
                levels[wattr(lvl, "ilvl")] = wattr(lvl.find(q("w:numFmt")), "val") or "decimal"
            self.abstract[wattr(an, "abstractNumId")] = levels
        for num in root.findall(q("w:num")):
            self.nums[wattr(num, "numId")] = wattr(num.find(q("w:abstractNumId")), "val")

    def fmt(self, num_id, ilvl):
        return self.abstract.get(self.nums.get(num_id), {}).get(ilvl or "0", "unknown")


class ParagraphWalker:
    """Collects text, per-character effective formatting, links, images and fields of one paragraph."""

    def __init__(self, fp, part, p):
        self.fp, self.part, self.rels = fp, part, fp.pkg.rels(part)
        styles = fp.styles
        ppr = p.find(q("w:pPr"))
        sid = wattr(ppr.find(q("w:pStyle")), "val") if ppr is not None else None
        sid = sid or styles.default_para
        self.chain = styles.chain(sid)
        self.style = styles.name(sid) if sid else "normal"
        jc = styles.para_prop(self.chain, ppr, q("w:jc"))
        align = wattr(jc, "val") or "left"
        self.align = {"both": "justify", "start": "left", "end": "right", "distribute": "justify"}.get(align, align)
        num_id = styles.para_prop(self.chain, ppr, f"{q('w:numPr')}/{q('w:numId')}")
        ilvl = styles.para_prop(self.chain, ppr, f"{q('w:numPr')}/{q('w:ilvl')}")
        nid = wattr(num_id, "val")
        self.list = None if nid in (None, "0") else {"fmt": fp.numbering.fmt(nid, wattr(ilvl, "val")),
                                                     "level": int(wattr(ilvl, "val") or 0)}
        pbb = styles.para_prop(self.chain, ppr, q("w:pageBreakBefore"))
        self.page_breaks = 1 if onoff(pbb) else 0
        self.text, self.props, self.links, self.images, self.fields = [], [], [], [], []
        self.suppress = 0
        self.fld_stack = []
        for child in p:
            if child.tag != q("w:pPr"):
                self.visit(child)

    def emit(self, s, props):
        hidden = self.suppress or any(f["phase"] == "result" and f["kind"] != "HYPERLINK" for f in self.fld_stack)
        if hidden:
            return
        for ch in s:
            self.text.append(ch)
            self.props.append(props)

    def visit(self, el):
        tag = el.tag
        if tag == q("w:r"):
            self.run(el)
        elif tag == q("w:hyperlink"):
            rel = self.rels.get(el.get(q("r:id")))
            target = rel[1] if rel else ("#" + (wattr(el, "anchor") or ""))
            start = len(self.text)
            for c in el:
                self.visit(c)
            self.links.append({"target": target, "text": "".join(self.text[start:])})
        elif tag == q("w:fldSimple"):
            instr = (wattr(el, "instr") or "").strip()
            kind = instr.split()[0].upper() if instr else ""
            self.fields.append(kind)
            start = len(self.text)
            hide = kind in COMPUTED_FIELDS
            self.suppress += hide
            for c in el:
                self.visit(c)
            self.suppress -= hide
            if kind == "HYPERLINK":
                self.links.append({"target": _field_target(instr), "text": "".join(self.text[start:])})
        elif tag in (q("w:del"), q("w:moveFrom")):
            return  # tracked deletions are not visible content
        else:
            for c in el:
                self.visit(c)

    def run(self, r):
        rpr = r.find(q("w:rPr"))
        rstyle = wattr(rpr.find(q("w:rStyle")), "val") if rpr is not None else None
        props = self.fp.styles.effective_run(self.chain, rstyle, rpr)
        for c in r:
            tag = c.tag
            if tag == q("w:t"):
                self.emit(c.text or "", props)
            elif tag == q("w:tab"):
                self.emit("\t", props)
            elif tag == q("w:br"):
                kind = wattr(c, "type")
                if kind == "page":
                    self.page_breaks += 1
                elif kind != "column":
                    self.emit("\n", props)
            elif tag == q("w:cr"):
                self.emit("\n", props)
            elif tag == q("w:noBreakHyphen"):
                self.emit("‑", props)
            elif tag == q("w:sym"):
                self.emit("�", props)
            elif tag == q("w:fldChar"):
                self.fld_char(wattr(c, "fldCharType"))
            elif tag == q("w:instrText"):
                if self.fld_stack:
                    self.fld_stack[-1]["instr"] += c.text or ""
            elif tag in (q("w:drawing"), q("w:pict")) or tag.endswith("}AlternateContent"):
                self.images.extend(self.fp.find_images(self.part, c))

    def fld_char(self, kind):
        if kind == "begin":
            self.fld_stack.append({"instr": "", "phase": "instr", "kind": ""})
        elif kind == "separate" and self.fld_stack:
            f = self.fld_stack[-1]
            f["phase"] = "result"
            f["kind"] = f["instr"].strip().split()[0].upper() if f["instr"].strip() else ""
            f["start"] = len(self.text)
            self.fields.append(f["kind"])
        elif kind == "end" and self.fld_stack:
            f = self.fld_stack.pop()
            if f["phase"] == "instr":  # field without a result section
                f["kind"] = f["instr"].strip().split()[0].upper() if f["instr"].strip() else ""
                self.fields.append(f["kind"])
            elif f["kind"] == "HYPERLINK":
                self.links.append({"target": _field_target(f["instr"]), "text": "".join(self.text[f["start"]:])})


def _field_target(instr):
    m = re.search(r'HYPERLINK\s+"([^"]*)"', instr, re.I)
    return m.group(1) if m else instr


def normalize_text(s):
    return unicodedata.normalize("NFC", s)


class Fingerprint:
    def __init__(self, path):
        self.pkg = Package(path)
        self.styles = Styles(self.pkg)
        self.numbering = Numbering(self.pkg)
        self.paragraphs = []   # every paragraph in reading order, body and table cells
        self.tables = []
        self.sections = []
        doc = self.pkg.xml(self.pkg.main)
        body = doc.find(q("w:body"))
        self._container(self.pkg.main, body, "body", self.paragraphs, top=True)
        self._resolve_section_headers()

    # -- body ---------------------------------------------------------------
    def _container(self, part, el, where, sink, top=False):
        for child in el:
            if child.tag == q("w:p"):
                sink.append(self._paragraph(part, child, where))
                if top:
                    sect = child.find(f"{q('w:pPr')}/{q('w:sectPr')}")
                    if sect is not None:
                        self.sections.append(self._section(sect))
            elif child.tag == q("w:tbl"):
                self._table(part, child, where, sink)
            elif child.tag == q("w:sdt"):
                content = child.find(q("w:sdtContent"))
                if content is not None:
                    self._container(part, content, where, sink, top)
            elif child.tag == q("w:sectPr") and top:
                self.sections.append(self._section(child))

    def _paragraph(self, part, p, where):
        w = ParagraphWalker(self, part, p)
        return {"where": where, "text": normalize_text("".join(w.text)), "chars": w.props,
                "style": w.style, "align": w.align, "list": w.list, "links": w.links,
                "images": w.images, "fields": w.fields, "page_breaks": w.page_breaks}

    def _table(self, part, tbl, where, sink):
        index = len(self.tables)
        table = {"index": index, "rows": [], "grid_cols": len(tbl.findall(f"{q('w:tblGrid')}/{q('w:gridCol')}")),
                 "header_rows": 0}
        self.tables.append(table)
        for r, tr in enumerate(tbl.findall(q("w:tr"))):
            if onoff(tr.find(f"{q('w:trPr')}/{q('w:tblHeader')}")):
                table["header_rows"] += 1
            row = []
            for c, tc in enumerate(tr.findall(q("w:tc"))):
                tcpr = tc.find(q("w:tcPr"))
                span = wattr(tcpr.find(q("w:gridSpan")), "val") if tcpr is not None else None
                vmerge_el = tcpr.find(q("w:vMerge")) if tcpr is not None else None
                shd = tcpr.find(q("w:shd")) if tcpr is not None else None
                fill = (wattr(shd, "fill") or "").upper()
                cell_paras = []
                self._container(part, tc, f"table {index} row {r} cell {c}", cell_paras)
                sink.extend(cell_paras)
                row.append({
                    "text": "\n".join(p["text"] for p in cell_paras),
                    "span": int(span or 1),
                    "vmerge": None if vmerge_el is None else (wattr(vmerge_el, "val") or "continue"),
                    "fill": None if fill in ("", "AUTO", "FFFFFF") else fill,
                })
            table["rows"].append(row)

    # -- sections, headers, footers -----------------------------------------
    def _section(self, sect):
        pgsz = sect.find(q("w:pgSz"))
        mar = sect.find(q("w:pgMar"))
        w = int(wattr(pgsz, "w") or 12240)
        h = int(wattr(pgsz, "h") or 15840)

        def inches(name):
            v = wattr(mar, name)
            return round(int(v) / 1440, 2) if v not in (None, "") else None

        refs = {}
        for kind in ("header", "footer"):
            for ref in sect.findall(q(f"w:{kind}Reference")):
                rel = self.pkg.rels(self.pkg.main).get(ref.get(q("r:id")))
                if rel:
                    refs[f"{kind}/{wattr(ref, 'type') or 'default'}"] = self._header_part(rel[1])
        return {
            "width_in": round(w / 1440, 2), "height_in": round(h / 1440, 2),
            "orient": wattr(pgsz, "orient") or ("landscape" if w > h else "portrait"),
            "margins": {k: inches(k) for k in ("top", "bottom", "left", "right")},
            "type": wattr(sect.find(q("w:type")), "val") or "nextPage",
            "title_pg": bool(onoff(sect.find(q("w:titlePg")))),
            "parts": refs,
        }

    def _header_part(self, part):
        root = self.pkg.xml(part)
        paras = []
        if root is not None:
            self._container(part, root, part, paras)
        return {"text": "\n".join(p["text"] for p in paras if p["text"].strip()),
                "fields": sorted(f for p in paras for f in p["fields"]),
                "images": [(i["sha1"], i["size_in"], i["effects"]) for p in paras for i in p["images"]]}

    def _resolve_section_headers(self):
        """A section without a header/footer reference of some type inherits the previous section's."""
        inherited = {}
        for s in self.sections:
            inherited.update(s["parts"])
            s["parts"] = dict(inherited)

    # -- images --------------------------------------------------------------
    def find_images(self, part, el):
        out = []
        placement = "floating" if el.find(f".//{q('wp:anchor')}") is not None else "inline"
        extent = el.find(f".//{q('wp:extent')}")
        size = None
        if extent is not None:
            size = (round(int(extent.get("cx")) / 914400, 2), round(int(extent.get("cy")) / 914400, 2))
        rids = [b.get(q("r:embed")) or b.get(q("r:link")) for b in el.iter(q("a:blip"))]
        rids += [d.get(q("r:id")) for d in el.iter(q("v:imagedata"))]
        effects = self._picture_effects(el)
        for rid in rids:
            rel = self.pkg.rels(part).get(rid)
            data = self.pkg.read(rel[1]) if rel and not rel[2] else None
            out.append({"sha1": hashlib.sha1(data).hexdigest()[:12] if data else None,
                        "size_in": size, "placement": placement, "effects": effects})
        return out

    @staticmethod
    def _picture_effects(el):
        """Visible picture decoration: borders, effects, crops, recolouring, picture links."""
        def sig(e):
            return ET.tostring(e, encoding="unicode")
        found = []
        for sp in el.iter(q("pic:spPr")):
            for c in sp:
                name = c.tag.split("}")[1]
                if name in ("xfrm", "prstGeom", "noFill"):
                    continue
                if name == "ln" and all(x.tag == q("a:noFill") for x in c):
                    continue
                found.append(sig(c))
        for blip in el.iter(q("a:blip")):
            found += [sig(c) for c in blip if c.tag != q("a:extLst")]
        for rect in el.iter(q("a:srcRect")):
            if rect.attrib:
                found.append(sig(rect))
        for link in el.iter(q("a:hlinkClick")):
            found.append("hlinkClick")
        return hashlib.sha1("|".join(sorted(found)).encode()).hexdigest()[:8] if found else None

    # -- summaries -----------------------------------------------------------
    @property
    def markers(self):
        c = Counter()
        for p in self.paragraphs:
            c.update(MARKER_RE.findall(p["text"]))
        # Inherited headers/footers are shared objects; count each part once.
        parts = {id(part): part for s in self.sections for part in s["parts"].values()}
        for part in parts.values():
            c.update(MARKER_RE.findall(part["text"]))
        return c

    @property
    def images(self):
        return [img for p in self.paragraphs for img in p["images"]]

    @property
    def links(self):
        return [lnk for p in self.paragraphs for lnk in p["links"]]

    @property
    def page_breaks(self):
        return sum(p["page_breaks"] for p in self.paragraphs)


if __name__ == "__main__":
    import json
    import sys
    fp = Fingerprint(sys.argv[1])
    print(json.dumps({
        "paragraphs": len(fp.paragraphs), "tables": len(fp.tables), "sections": fp.sections,
        "markers": len(fp.markers), "images": fp.images, "links": fp.links, "page_breaks": fp.page_breaks,
        "styles_used": Counter(p["style"] for p in fp.paragraphs),
        "lists": Counter(f"{p['list']['fmt']}/{p['list']['level']}" for p in fp.paragraphs if p["list"]),
    }, indent=2, default=str, ensure_ascii=False))
