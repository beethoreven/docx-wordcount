# -*- coding: utf-8 -*-
"""
docx_wordcount.py - Count .docx text volume exactly the way Microsoft Word does,
without opening Word. Pure stdlib (zipfile + ElementTree), no dependencies.

Every rule below was measured against Word's own ComputeStatistics() over
controlled probe documents plus a set of real-world files.

WORDS ("字數")
  - each East-Asian wide character counts as 1 word: CJK ideographs, kana,
    Hangul, bopomofo, CJK punctuation, fullwidth forms
  - each maximal run of anything else counts as 1 word: letters, digits, ASCII
    punctuation, ellipsis, / - _ | circled digits, roman numerals, emoji
  - en dash U+2013 and em dash U+2014 SEPARATE words and are worth 0 words
    (they do still count as characters)
  - whitespace separates: SP, TAB, NBSP, U+3000; a <w:br/> separates but is
    not a character at all
  - Unicode "ambiguous width" characters (·, “ ”, ①, Ⅰ …) switch sides
    depending on the run: East-Asian in a run hinted w:hint="eastAsia",
    Latin otherwise. En/em dash stay separators either way.
  - the label of an automatically numbered list ("1、", "a)", a bullet) IS
    counted, along with the tab or space that follows it

CHARACTERS
  no-spaces : every character except whitespace
  w/ spaces : every character; paragraph marks are NOT counted
  a footnote/endnote reference mark counts as 1 character but 0 words

PARAGRAPHS
  a paragraph counts only when it holds at least one non-whitespace character

ZERO-WIDTH
  U+200B, U+200C, U+200D and U+FEFF are invisible to Word: no character, no
  word, and they do not even break a word run

SCOPE
  Word's headline number covers the main story only - body text and tables.
  Headers, footers, footnotes, endnotes, comments, textboxes and hidden text
  are excluded; Word's dialog checkbox "include textboxes, footnotes and
  endnotes" is reproduced by --include-boxes / --include-notes.
"""
import sys, os, zipfile, argparse, unicodedata
import xml.etree.ElementTree as ET

# This tool prints CJK filenames. On a machine whose console codepage is Big5
# (cp950) an unencodable character - any simplified-only filename, kana, emoji -
# aborts the whole run with UnicodeEncodeError, and whether that happens depends
# entirely on ambient PYTHONIOENCODING/PYTHONUTF8. Pin it here, do not inherit it.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"

WHITESPACE = frozenset(chr(c) for c in (0x20, 0x09, 0xA0, 0x3000, 0x0A, 0x0D, 0x0B))  # SP TAB NBSP IDEO-SP LF CR VT
REF_MARK = chr(0xE000)      # footnote/endnote mark: 1 character, 0 words
SEP_NOCHAR = chr(0xE001)    # <w:br/>: separates words, counts as nothing
# Measured: exactly these three are charged 1 character but 0 words and break a word run.
ZERO_WORD_SEPARATORS = frozenset((chr(0x2013), chr(0x2014), chr(0x2022), REF_MARK))

# Measured: these four are invisible to Word entirely - 0 characters, 0 words,
# and they do NOT break a word run (A + ZWSP + B counts as one word).
# Deliberately narrow: soft hyphen, LRM/RLM, WORD JOINER and combining marks
# are all Cf/Mn too, and Word counts every one of them as a normal character.
IGNORED_CHARS = frozenset(chr(c) for c in (0x200B, 0x200C, 0x200D, 0xFEFF))

_REF_MARK_TAGS = tuple(W + t for t in
                       ("footnoteReference", "endnoteReference", "footnoteRef", "endnoteRef"))

_EA_RANGES = (
    (0x1100, 0x11FF), (0x2E80, 0x2EFF), (0x2F00, 0x2FDF), (0x2FF0, 0x2FFF),
    (0x3001, 0x303F),                               # CJK punctuation (U+3000 is a space)
    (0x3040, 0x309F), (0x30A0, 0x30FF), (0x3100, 0x312F), (0x3130, 0x318F),
    (0x3190, 0x319F), (0x31C0, 0x31EF), (0x31F0, 0x31FF), (0x3200, 0x32FF),
    (0x3300, 0x33FF), (0x3400, 0x4DBF), (0x4E00, 0x9FFF), (0xA000, 0xA4CF),
    (0xA960, 0xA97F), (0xAC00, 0xD7FF), (0xF900, 0xFAFF), (0xFE10, 0xFE1F),
    (0xFE30, 0xFE4F), (0xFE50, 0xFE6F),
    (0xFF01, 0xFF60), (0xFF61, 0xFF9F), (0xFFE0, 0xFFE6),   # fullwidth + halfwidth forms
    (0x1F200, 0x1F2FF), (0x20000, 0x2FA1F),
)

_ea_cache = {}


def is_east_asian(ch):
    """True for characters Word always treats as one wide character = one word."""
    v = _ea_cache.get(ch)
    if v is None:
        cp = ord(ch)
        v = any(lo <= cp <= hi for lo, hi in _EA_RANGES)
        _ea_cache[ch] = v
    return v


_amb_cache = {}


def is_ambiguous_width(ch):
    """True for Unicode 'ambiguous width' characters, which Word classifies
    as East-Asian only inside a run hinted w:hint="eastAsia"."""
    v = _amb_cache.get(ch)
    if v is None:
        v = unicodedata.east_asian_width(ch) == "A"
        _amb_cache[ch] = v
    return v


def count_segments(segments):
    """Count one paragraph given as [(text, east_asian_hint), ...]."""
    words = chars = chars_ws = fareast = 0
    in_latin_run = False
    for text, ea_hint in segments:
        for ch in text:
            if ch in IGNORED_CHARS:
                continue
            if ch in WHITESPACE:
                chars_ws += 1
                in_latin_run = False
            elif ch == SEP_NOCHAR:
                in_latin_run = False
            elif ch in ZERO_WORD_SEPARATORS:
                chars += 1
                chars_ws += 1
                in_latin_run = False
            elif is_east_asian(ch) or (ea_hint and is_ambiguous_width(ch)):
                chars += 1
                chars_ws += 1
                fareast += 1
                words += 1
                in_latin_run = False
            else:
                chars += 1
                chars_ws += 1
                if not in_latin_run:
                    words += 1
                    in_latin_run = True
    return words, chars, chars_ws, fareast


def count_text(text, east_asian_hint=False):
    """Convenience wrapper for counting a bare string."""
    return count_segments([(text, east_asian_hint)])


# --------------------------------------------------------------------------
# XML traversal
# --------------------------------------------------------------------------

def _children(node):
    """Children of node, resolving mc:AlternateContent down to one branch.

    A DrawingML shape is stored twice - under <mc:Choice> and again under
    <mc:Fallback> as legacy VML - so walking both would count its text twice.
    """
    if node.tag == MC + "AlternateContent":
        branch = node.find(MC + "Choice")
        if branch is None:
            branch = node.find(MC + "Fallback")
        return list(branch) if branch is not None else []
    return list(node)


def _hint_of(rpr):
    """'eastAsia'/'default'/None from a <w:rPr>'s <w:rFonts w:hint=...>."""
    if rpr is None:
        return None
    rfonts = rpr.find(W + "rFonts")
    if rfonts is None:
        return None
    return rfonts.get(W + "hint")


def _run_is_hidden(run):
    rpr = run.find(W + "rPr")
    if rpr is None:
        return False
    vanish = rpr.find(W + "vanish")
    if vanish is None:
        return False
    return vanish.get(W + "val", "true") not in ("0", "false")


def _paragraph_segments(p, include_deleted=False):
    """Visible text of one <w:p> as [(text, east_asian_hint), ...].

    Text box content is skipped - Word keeps it in a separate story.
    """
    segments = []

    def walk(node, ea):
        for child in _children(node):
            tag = child.tag
            if tag == W + "txbxContent" or tag == W + "pPr" or tag == W + "rPr":
                continue        # pPr holds tab STOP definitions, not tab characters
            elif tag == W + "t":
                segments.append((child.text or "", ea))
            elif tag == W + "delText":
                if include_deleted:
                    segments.append((child.text or "", ea))
            elif tag == W + "instrText":
                pass                                    # field codes are not text
            elif tag == W + "tab":
                segments.append((chr(0x09), ea))
            elif tag == W + "br" or tag == W + "cr":
                segments.append((SEP_NOCHAR, ea))
            elif tag in _REF_MARK_TAGS:
                segments.append((REF_MARK, ea))
            elif tag == W + "noBreakHyphen":
                segments.append((chr(0x2011), ea))   # non-breaking hyphen
            elif tag == W + "softHyphen":
                segments.append((chr(0x00AD), ea))   # soft hyphen
            elif tag == W + "sym":
                segments.append((chr(0xE002), ea))   # a symbol glyph: 1 Latin char
            elif tag == W + "r":
                if _run_is_hidden(child):
                    continue
                h = _hint_of(child.find(W + "rPr"))
                walk(child, ea if h is None else (h == "eastAsia"))
            else:
                walk(child, ea)

    # Only a run's own w:rFonts/@w:hint decides the ambiguous-width characters;
    # measurement showed the paragraph mark and docDefaults do NOT propagate it.
    walk(p, False)
    return segments


def _iter_paragraphs(root):
    """Main-story paragraphs in document order, text box content excluded."""
    def walk(node):
        for child in _children(node):
            if child.tag == W + "txbxContent":
                continue
            if child.tag == W + "p":
                yield child
                continue                                # <w:p> nests only via a text box
            for x in walk(child):
                yield x
    body = root.find(W + "body")
    for x in walk(body if body is not None else root):
        yield x


def _iter_textbox_paragraphs(root):
    """Paragraphs inside text boxes, Choice/Fallback duplicates collapsed."""
    def walk(node, inbox):
        for child in _children(node):
            if child.tag == W + "txbxContent":
                for x in walk(child, True):
                    yield x
                continue
            if child.tag == W + "p":
                if inbox:
                    yield child
                for x in walk(child, inbox):            # a box may sit inside a paragraph
                    yield x
                continue
            for x in walk(child, inbox):
                yield x
    for x in walk(root, False):
        yield x


# --------------------------------------------------------------------------
# Automatic list numbering - Word counts the label, so we have to render it
# --------------------------------------------------------------------------

_ROMAN = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
          (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"))
_CN_DIGITS = "〇一二三四五六七八九"


def _roman(n):
    out = []
    for v, s in _ROMAN:
        while n >= v:
            out.append(s)
            n -= v
    return "".join(out)


def _alpha(n):
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(ord("A") + r) + out
    return out


def _chinese(n):
    if n <= 0 or n > 99:
        return str(n)
    if n < 10:
        return _CN_DIGITS[n]
    tens, ones = divmod(n, 10)
    return (("" if tens == 1 else _CN_DIGITS[tens]) + "十" +
            ("" if ones == 0 else _CN_DIGITS[ones]))


def _format_number(n, fmt):
    if fmt in ("none", "bullet"):
        return ""
    if fmt == "decimalZero":
        return "%02d" % n
    if fmt == "lowerLetter":
        return _alpha(n).lower()
    if fmt == "upperLetter":
        return _alpha(n)
    if fmt == "lowerRoman":
        return _roman(n).lower()
    if fmt == "upperRoman":
        return _roman(n)
    if fmt == "decimalFullWidth" or fmt == "decimalFullWidth2":
        return "".join(chr(0xFF10 + int(d)) for d in str(n))
    if fmt == "decimalEnclosedCircle" or fmt == "decimalEnclosedCircleChinese":
        return chr(0x245F + n) if 1 <= n <= 20 else str(n)
    if fmt == "decimalEnclosedParen":
        return "(%d)" % n
    if fmt.startswith("chinese") or fmt.startswith("taiwanese") or fmt.startswith("ideograph"):
        return _chinese(n)
    return str(n)


def _load_numbering(z, names):
    """numId -> {ilvl: {fmt, text, suff, start}}."""
    if "word/numbering.xml" not in names:
        return {}
    root = ET.fromstring(z.read("word/numbering.xml"))

    def read_levels(parent):
        levels = {}
        for lvl in parent.findall(W + "lvl"):
            try:
                ilvl = int(lvl.get(W + "ilvl", "0"))
            except ValueError:
                continue

            def val(tag, default):
                e = lvl.find(W + tag)
                return default if e is None else e.get(W + "val", default)

            try:
                start = int(val("start", "1"))
            except ValueError:
                start = 1
            levels[ilvl] = dict(fmt=val("numFmt", "decimal"), text=val("lvlText", ""),
                                suff=val("suff", "tab"), start=start)
        return levels

    abstracts = {an.get(W + "abstractNumId"): read_levels(an)
                 for an in root.findall(W + "abstractNum")}
    nums = {}
    for num in root.findall(W + "num"):
        aid = num.find(W + "abstractNumId")
        if aid is None:
            continue
        levels = dict(abstracts.get(aid.get(W + "val"), {}))
        for ov in num.findall(W + "lvlOverride"):
            try:
                ilvl = int(ov.get(W + "ilvl", "0"))
            except ValueError:
                continue
            replacement = read_levels(ov)
            if ilvl in replacement:
                levels[ilvl] = replacement[ilvl]
            so = ov.find(W + "startOverride")
            if so is not None and ilvl in levels:
                try:
                    levels[ilvl] = dict(levels[ilvl], start=int(so.get(W + "val", "1")))
                except ValueError:
                    pass
        nums[num.get(W + "numId")] = levels
    return nums


def _load_style_numbering(z, names):
    """styleId -> (numId, ilvl) for paragraph styles that carry numbering."""
    if "word/styles.xml" not in names:
        return {}
    root = ET.fromstring(z.read("word/styles.xml"))
    out = {}
    for st in root.findall(W + "style"):
        numpr = st.find(W + "pPr/" + W + "numPr")
        if numpr is None:
            continue
        nid = numpr.find(W + "numId")
        ilvl = numpr.find(W + "ilvl")
        if nid is not None:
            out[st.get(W + "styleId")] = (nid.get(W + "val"),
                                          int(ilvl.get(W + "val", "0")) if ilvl is not None else 0)
    return out


def _numbering_of(p, style_num):
    """(numId, ilvl) for a paragraph, from direct formatting or its style."""
    ppr = p.find(W + "pPr")
    if ppr is None:
        return None
    numpr = ppr.find(W + "numPr")
    if numpr is not None:
        nid = numpr.find(W + "numId")
        ilvl = numpr.find(W + "ilvl")
        if nid is not None:
            return (nid.get(W + "val"),
                    int(ilvl.get(W + "val", "0")) if ilvl is not None else 0)
    pstyle = ppr.find(W + "pStyle")
    if pstyle is not None:
        return style_num.get(pstyle.get(W + "val"))
    return None


class _ListCounters(object):
    """Renders the label Word would show in front of a numbered paragraph."""

    def __init__(self, numbering):
        self.numbering = numbering
        self.state = {}

    def label(self, numId, ilvl):
        levels = self.numbering.get(numId)
        if not levels or ilvl not in levels:
            return ""
        counters = self.state.setdefault(numId, {})
        counters[ilvl] = counters.get(ilvl, levels[ilvl]["start"] - 1) + 1
        for deeper in [k for k in counters if k > ilvl]:
            del counters[deeper]

        spec = levels[ilvl]
        text = spec["text"]
        for i in range(9, 0, -1):
            token = "%%%d" % i
            if token not in text:
                continue
            lv = i - 1
            n = counters.get(lv, levels.get(lv, {}).get("start", 1))
            text = text.replace(token, _format_number(n, levels.get(lv, spec)["fmt"]))
        suffix = {"tab": "\t", "space": " ", "nothing": ""}.get(spec["suff"], "\t")
        return text + suffix if text else ""


# --------------------------------------------------------------------------
# Document level
# --------------------------------------------------------------------------

def count_docx(path, include_boxes=False, include_notes=False, include_deleted=False):
    res = dict(words=0, chars=0, chars_ws=0, fareast=0, paragraphs=0)
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        style_num = _load_style_numbering(z, names)
        numbering = _load_numbering(z, names)

        def tally(part, count_paras=True, boxes_only=False, numbered=False):
            if part not in names:
                return
            root = ET.fromstring(z.read(part))
            paras = list(_iter_textbox_paragraphs(root) if boxes_only
                         else _iter_paragraphs(root))
            counters = _ListCounters(numbering)
            for p in paras:
                segments = _paragraph_segments(p, include_deleted)
                if numbered:
                    num = _numbering_of(p, style_num)
                    if num and num[0] not in (None, "0"):
                        label = counters.label(num[0], num[1])
                        if label:
                            segments.insert(0, (label, False))
                w, c, cw, fe = count_segments(segments)
                res["words"] += w
                res["chars"] += c
                res["chars_ws"] += cw
                res["fareast"] += fe
                if count_paras and c:
                    # Measured: a paragraph holding only whitespace (two U+3000,
                    # for instance) is not counted, so the test is the
                    # no-spaces character count, not the with-spaces one.
                    res["paragraphs"] += 1

        tally("word/document.xml", numbered=True)
        if include_boxes:
            tally("word/document.xml", count_paras=False, boxes_only=True)
        if include_notes:
            for part in ("word/footnotes.xml", "word/endnotes.xml"):
                tally(part, count_paras=False)
    return res


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _dw(s):
    """Display width: East-Asian wide characters take two terminal columns."""
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in s)


def main():
    ap = argparse.ArgumentParser(
        description="Word-identical .docx word count, without opening Word.")
    ap.add_argument("paths", nargs="+", help=".docx files or directories")
    ap.add_argument("-r", "--recursive", action="store_true", help="recurse into directories")
    ap.add_argument("--include-boxes", action="store_true", help="also count text boxes")
    ap.add_argument("--include-notes", action="store_true", help="also count footnotes/endnotes")
    ap.add_argument("--include-deleted", action="store_true", help="count tracked deletions")
    ap.add_argument("--tsv", action="store_true", help="tab-separated output for piping")
    args = ap.parse_args()

    files = []
    for p in args.paths:
        if os.path.isdir(p):
            if args.recursive:
                for dirpath, _, filenames in os.walk(p):
                    files += [os.path.join(dirpath, f) for f in sorted(filenames)
                              if f.lower().endswith(".docx") and not f.startswith("~$")]
            else:
                files += [os.path.join(p, f) for f in sorted(os.listdir(p))
                          if f.lower().endswith(".docx") and not f.startswith("~$")]
        else:
            files.append(p)

    # PowerShell and cmd do not expand wildcards for external programs, so a
    # literal "*.docx" arrives here as a filename. Say so plainly instead of
    # letting it surface as OSError errno 22.
    missing = [f for f in files if not os.path.isfile(f)]

    rows = []
    total = dict(words=0, chars=0, chars_ws=0, fareast=0, paragraphs=0)
    for f in files:
        if f in missing:
            hint = "找不到這個路徑"
            if any(c in f for c in "*?"):
                hint += "（PowerShell/cmd 不會展開通配字元，請直接給資料夾或用 "
                hint += "(Get-ChildItem *.docx).FullName）"
            rows.append((os.path.basename(f), None, hint))
            continue
        try:
            r = count_docx(f, args.include_boxes, args.include_notes, args.include_deleted)
        except Exception as exc:
            rows.append((os.path.basename(f), None, str(exc)))
            continue
        for k in total:
            total[k] += r[k]
        rows.append((os.path.basename(f), r, None))

    keys = ("words", "chars", "chars_ws", "fareast", "paragraphs")
    if args.tsv:
        # Full path, not the basename: a recursive run easily hits the same
        # filename in two folders, and consumers key on this column.
        print("file\t" + "\t".join(keys))
        for (name, r, err), path in zip(rows, files):
            print("%s\tERROR\t%s" % (path, err) if err
                  else path + "\t" + "\t".join(str(r[k]) for k in keys))
        return

    labels = ("字數", "字元(不計空白)", "字元(含空白)", "中文字元", "段落")
    widths = (9, 14, 12, 10, 6)
    # Numbers first, filename last: CJK filenames make column alignment
    # unreliable in any terminal that measures wide characters differently.
    head = "  ".join(" " * max(0, w - _dw(l)) + l for l, w in zip(labels, widths))
    print(head + "  檔案")
    print("-" * (_dw(head) + 30))
    for name, r, err in rows:
        if err:
            print(" " * _dw(head) + "  %s  <-- ERROR: %s" % (name, err))
            continue
        print("  ".join("%*d" % (w, r[k]) for k, w in zip(keys, widths)) + "  " + name)
    if len(rows) > 1:
        print("-" * (_dw(head) + 30))
        print("  ".join("%*d" % (w, total[k]) for k, w in zip(keys, widths)) +
              "  總計 %d 個檔案" % len(rows))


if __name__ == "__main__":
    main()
