"""Compare two DOCX fingerprints and produce findings scored against the Experiment 1 rubric.

Severity:
  CRITICAL  content lost, duplicated or changed (violates "zero unexplained text loss")
  MAJOR     structure or layout the user will notice (tables, images, links, headings, sections)
  MINOR     formatting drift (fonts, colors, alignment, style names)
"""
import difflib
import re
from collections import Counter, defaultdict

from docxfp import PROPS

AREAS = ("Content", "Formatting", "Structure", "Pagination", "Reliability", "Transparency")
SEVERITY_ORDER = {"CRITICAL": 0, "MAJOR": 1, "MINOR": 2}


class Findings(list):
    def add(self, severity, area, message, examples=()):
        self.append({"severity": severity, "area": area, "message": message, "examples": list(examples)[:5]})


def _tokens(fp):
    text = " ".join(p["text"] for p in fp.paragraphs).replace(" ", " ")
    return re.findall(r"\S+", text)


def _context(tokens, i, j):
    return " ".join(tokens[max(0, i - 4):j + 4])


def _marker(text):
    m = re.search(r"\[[A-Z]-\d{3}\]", text)
    return m.group(0) if m else repr(text[:40])


def compare_content(a, b, f, metrics):
    ma, mb = a.markers, b.markers
    missing = sorted((ma - mb).elements())
    extra = sorted((mb - ma).elements())
    if missing:
        f.add("CRITICAL", "Content", f"{len(missing)} content markers missing", missing)
    if extra:
        f.add("CRITICAL", "Content", f"{len(extra)} content markers duplicated or added", extra)

    ta, tb = _tokens(a), _tokens(b)
    sm = difflib.SequenceMatcher(a=ta, b=tb, autojunk=False)
    removed = added = 0
    examples = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue
        removed += i2 - i1
        added += j2 - j1
        examples.append(f"{op}: «{' '.join(ta[i1:i2])[:80]}» → «{' '.join(tb[j1:j2])[:80]}» near «{_context(ta, i1, i2)[:100]}»")
    metrics["words_original"] = len(ta)
    metrics["words_preserved_pct"] = round(100 * sum(blk.size for blk in sm.get_matching_blocks()) / max(1, len(ta)), 2)
    if removed or added:
        f.add("CRITICAL", "Content", f"text changed: {removed} words removed/replaced, {added} words added", examples)

    nbsp_a = sum(p["text"].count(" ") for p in a.paragraphs)
    nbsp_b = sum(p["text"].count(" ") for p in b.paragraphs)
    if nbsp_a != nbsp_b:
        f.add("MINOR", "Formatting", f"non-breaking spaces changed ({nbsp_a} → {nbsp_b})")
    breaks_a = sum(p["text"].count("\n") for p in a.paragraphs)
    breaks_b = sum(p["text"].count("\n") for p in b.paragraphs)
    if breaks_a != breaks_b:
        f.add("MINOR", "Formatting", f"manual line breaks changed ({breaks_a} → {breaks_b})")


def paired_paragraphs(a, b):
    ka = [p["text"] for p in a.paragraphs]
    kb = [p["text"] for p in b.paragraphs]
    sm = difflib.SequenceMatcher(a=ka, b=kb, autojunk=False)
    for blk in sm.get_matching_blocks():
        for k in range(blk.size):
            yield a.paragraphs[blk.a + k], b.paragraphs[blk.b + k]


def compare_formatting(a, b, f, metrics):
    style_changes = Counter()
    heading_changes = defaultdict(list)
    align_changes = defaultdict(list)
    list_lost, list_changed = [], defaultdict(list)
    prop_changes = defaultdict(lambda: [0, []])
    chars_total = chars_same = 0

    for pa, pb in paired_paragraphs(a, b):
        if not pa["text"].strip():
            continue
        if pa["style"] != pb["style"]:
            key = f"{pa['style']} → {pb['style']}"
            if pa["style"].startswith(("heading", "title")) or pb["style"].startswith(("heading", "title")):
                heading_changes[key].append(_marker(pa["text"]))
            else:
                style_changes[key] += 1
        if pa["align"] != pb["align"]:
            align_changes[f"{pa['align']} → {pb['align']}"].append(_marker(pa["text"]))
        if pa["list"] and not pb["list"]:
            list_lost.append(_marker(pa["text"]))
        elif pa["list"] != pb["list"]:
            list_changed[f"{pa['list']} → {pb['list']}"].append(_marker(pa["text"]))
        if len(pa["chars"]) == len(pb["chars"]):
            for ca, cb in zip(pa["chars"], pb["chars"]):
                chars_total += 1
                if ca == cb:
                    chars_same += 1
                    continue
                for k, va, vb in zip(PROPS, ca, cb):
                    if va != vb:
                        entry = prop_changes[(k, va, vb)]
                        entry[0] += 1
                        if len(entry[1]) < 3 and _marker(pa["text"]) not in entry[1]:
                            entry[1].append(_marker(pa["text"]))

    metrics["formatting_identical_pct"] = round(100 * chars_same / max(1, chars_total), 2)
    for key, where in heading_changes.items():
        f.add("MAJOR", "Structure", f"heading/title style changed {key} ({len(where)} paragraphs)", where)
    for key, n in style_changes.most_common():
        f.add("MINOR", "Formatting", f"paragraph style changed {key} ({n} paragraphs)")
    for key, where in align_changes.items():
        f.add("MINOR", "Formatting", f"alignment changed {key} ({len(where)} paragraphs)", where)
    if list_lost:
        f.add("MAJOR", "Structure", f"list formatting lost on {len(list_lost)} paragraphs", list_lost)
    for key, where in list_changed.items():
        f.add("MINOR", "Formatting", f"list changed {key} ({len(where)} paragraphs)", where)
    for (k, va, vb), (n, where) in sorted(prop_changes.items(), key=lambda kv: -kv[1][0]):
        f.add("MINOR", "Formatting", f"{k}: {va!r} → {vb!r} on {n} characters", where)


def compare_structure(a, b, f):
    if len(a.tables) != len(b.tables):
        f.add("MAJOR", "Structure", f"table count changed ({len(a.tables)} → {len(b.tables)})")
    for ta, tb in zip(a.tables, b.tables):
        name = f"table {ta['index'] + 1}"
        shape_a = [len(r) for r in ta["rows"]]
        shape_b = [len(r) for r in tb["rows"]]
        if len(shape_a) != len(shape_b):
            f.add("MAJOR", "Structure", f"{name}: row count changed ({len(shape_a)} → {len(shape_b)})")
        elif shape_a != shape_b:
            f.add("MAJOR", "Structure", f"{name}: cells per row changed")
        merges_a = [(r, c, x["span"], x["vmerge"]) for r, row in enumerate(ta["rows"]) for c, x in enumerate(row)
                    if x["span"] > 1 or x["vmerge"]]
        merges_b = [(r, c, x["span"], x["vmerge"]) for r, row in enumerate(tb["rows"]) for c, x in enumerate(row)
                    if x["span"] > 1 or x["vmerge"]]
        if merges_a != merges_b:
            diff = [f"row {r} cell {c}: span={s} vMerge={v}" for r, c, s, v in sorted(set(merges_b) - set(merges_a))]
            f.add("MAJOR", "Structure", f"{name}: merged cells changed", diff)
        if ta["header_rows"] != tb["header_rows"]:
            f.add("MINOR", "Pagination", f"{name}: repeating header rows changed ({ta['header_rows']} → {tb['header_rows']})")
        fills_a = [x["fill"] for row in ta["rows"] for x in row]
        fills_b = [x["fill"] for row in tb["rows"] for x in row]
        if len(fills_a) == len(fills_b) and fills_a != fills_b:
            f.add("MINOR", "Formatting", f"{name}: cell shading changed")

    ia, ib = a.images, b.images
    if len(ia) != len(ib):
        f.add("MAJOR", "Structure", f"image count changed ({len(ia)} → {len(ib)})")
    for n, (x, y) in enumerate(zip(ia, ib), 1):
        if x["sha1"] != y["sha1"]:
            f.add("MINOR", "Structure", f"image {n}: image data re-encoded or replaced")
        if x["size_in"] and y["size_in"] and any(abs(p - q) > 0.05 for p, q in zip(x["size_in"], y["size_in"])):
            f.add("MINOR", "Structure", f"image {n}: size changed {x['size_in']} → {y['size_in']} in")
        if x.get("effects") != y.get("effects"):
            f.add("MAJOR", "Structure", f"image {n}: border, effect, crop or picture link changed")
        if x["placement"] != y["placement"]:
            f.add("MINOR", "Structure", f"image {n}: placement changed {x['placement']} → {y['placement']}")

    la = Counter((l["target"], l["text"]) for l in a.links)
    lb = Counter((l["target"], l["text"]) for l in b.links)
    lost = la - lb
    if lost:
        by_text_b = {txt: t for t, txt in lb}
        targets_b = {t for t, _ in lb}
        retargeted = [f"{t} → {by_text_b[txt]}" for t, txt in lost if txt in by_text_b]
        relabeled = [f"{t} ({txt})" for t, txt in lost if txt not in by_text_b and t in targets_b]
        missing = [f"{t} ({txt})" for t, txt in lost if txt not in by_text_b and t not in targets_b]
        if retargeted:
            f.add("MAJOR", "Structure", f"{len(retargeted)} hyperlinks now point to a different URL", retargeted)
        if missing:
            f.add("MAJOR", "Structure", f"{len(missing)} hyperlinks lost", missing)
        if relabeled:
            f.add("MINOR", "Structure", f"{len(relabeled)} hyperlinks with changed text", relabeled)

    if len(a.paragraphs) != len(b.paragraphs):
        f.add("MINOR", "Structure", f"paragraph count changed ({len(a.paragraphs)} → {len(b.paragraphs)})")


def compare_pagination(a, b, f):
    sa, sb = a.sections, b.sections
    if len(sa) != len(sb):
        f.add("MAJOR", "Pagination", f"section count changed ({len(sa)} → {len(sb)})")
    for n, (x, y) in enumerate(zip(sa, sb), 1):
        if (x["width_in"], x["height_in"], x["orient"]) != (y["width_in"], y["height_in"], y["orient"]):
            f.add("MAJOR", "Pagination", f"section {n}: page size/orientation changed "
                  f"{x['width_in']}x{x['height_in']} {x['orient']} → {y['width_in']}x{y['height_in']} {y['orient']}")
        if x["margins"] != y["margins"]:
            f.add("MAJOR", "Pagination", f"section {n}: margins changed {x['margins']} → {y['margins']}")
        if x["title_pg"] != y["title_pg"]:
            f.add("MAJOR", "Pagination", f"section {n}: different-first-page setting changed")
        for key, part in x["parts"].items():
            other = y["parts"].get(key)
            if key.endswith("/first") and not x["title_pg"]:
                continue  # inactive first-page header
            if other is None:
                f.add("MAJOR", "Pagination", f"section {n}: {key} lost", [part["text"][:80]])
                continue
            if part["text"] != other["text"]:
                f.add("MAJOR", "Pagination", f"section {n}: {key} text changed",
                      [f"«{part['text'][:60]}» → «{other['text'][:60]}»"])
            if part.get("images") != other.get("images"):
                f.add("MAJOR", "Pagination", f"section {n}: {key} pictures changed "
                      f"({len(part.get('images', []))} → {len(other.get('images', []))})")
            if part["fields"] != other["fields"]:
                f.add("MAJOR", "Pagination", f"section {n}: {key} fields changed {part['fields']} → {other['fields']}")
    if a.page_breaks != b.page_breaks:
        f.add("MAJOR", "Pagination", f"explicit page breaks changed ({a.page_breaks} → {b.page_breaks})")


def compare(a, b):
    """Return (findings, metrics) for original fingerprint a vs round-tripped fingerprint b."""
    f, metrics = Findings(), {}
    compare_content(a, b, f, metrics)
    compare_formatting(a, b, f, metrics)
    compare_structure(a, b, f)
    compare_pagination(a, b, f)
    return f, metrics


def area_status(findings):
    # A candidate that crashed or produced no readable output was not evaluated on anything else.
    if any(x["severity"] == "CRITICAL" and x["area"] == "Reliability" for x in findings):
        return {area: ("FAIL" if area == "Reliability" else "SKIP") for area in AREAS}
    status = {}
    for area in AREAS:
        sev = [x["severity"] for x in findings if x["area"] == area]
        status[area] = "FAIL" if {"CRITICAL", "MAJOR"} & set(sev) else ("WARN" if sev else "PASS")
    return status
