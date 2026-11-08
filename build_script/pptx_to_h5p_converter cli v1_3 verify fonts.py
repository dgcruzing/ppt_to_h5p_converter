"""Read used PPTX fonts and compare them with fonts visible to the renderer.

No font installation, document changes, GUI automation, or network access.
"""

import ctypes
from functools import lru_cache
import io
import logging
import os
from pathlib import Path
import posixpath
import re
import struct
import subprocess
import sys
from typing import NamedTuple
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


class FontCheckError(Exception):
    """Font availability could not be established; distinct from missing fonts."""


class Font(NamedTuple):
    family: str
    bold: bool = False
    italic: bool = False

    @property
    def style(self):
        return "Bold Italic" if self.bold and self.italic else (
            "Bold" if self.bold else "Italic" if self.italic else "Regular")


def normalized(name):
    return "".join(c for c in unicodedata.normalize("NFKC", name).casefold() if c.isalnum())


def _properties(node, result):
    if node is None:
        return
    for key in ("b", "i", "lang"):
        if key in node.attrib:
            result[key] = node.get(key)
    for key in ("latin", "ea", "cs", "sym"):
        child = node.find("a:" + key, NS)
        if child is not None and child.get("typeface"):
            result[key] = child.get("typeface")


def _paragraph(node, result):
    if node is None:
        return
    _properties(node.find("a:defRPr", NS), result)
    for child in node:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag in ("buNone", "buChar", "buAutoNum"):
            result["bullet"] = tag
        elif tag == "buFont":
            result["bullet_font"] = child.get("typeface")
        elif tag == "buFontTx":
            result["bullet_font"] = None


def _list_style(node, level, result):
    if node is not None:
        _paragraph(node.find("a:defPPr", NS), result)
        _paragraph(node.find(f"a:lvl{level + 1}pPr", NS), result)


def _script(char, language):
    """Select the theme's script font only for characters of that script."""
    name = unicodedata.name(char, "")
    if name.startswith(("CJK", "IDEOGRAPHIC", "HIRAGANA", "KATAKANA")):
        return "ea", ("Jpan" if language.startswith("ja") else
                      "Hant" if language.lower().startswith(("zh-tw", "zh-hk", "zh-mo")) else "Hans")
    if name.startswith("HANGUL"):
        return "ea", "Hang"
    scripts = (("ARABIC", "Arab"), ("HEBREW", "Hebr"), ("SYRIAC", "Syrc"),
               ("THAI", "Thai"), ("DEVANAGARI", "Deva"), ("BENGALI", "Beng"),
               ("GUJARATI", "Gujr"), ("GURMUKHI", "Guru"), ("TAMIL", "Taml"),
               ("TELUGU", "Telu"), ("KANNADA", "Knda"), ("MALAYALAM", "Mlym"),
               ("SINHALA", "Sinh"), ("KHMER", "Khmr"), ("LAO", "Laoo"),
               ("MYANMAR", "Mymr"), ("ETHIOPIC", "Ethi"), ("TIBETAN", "Tibt"),
               ("MONGOLIAN", "Mong"), ("THAANA", "Thaa"))
    for prefix, script in scripts:
        if name.startswith(prefix):
            return "cs", script
    return "latin", None


class DeckFonts:
    def __init__(self, archive):
        self.archive = archive
        self.cache = {}
        self.required = set()
        self.unresolved = False

    def root(self, part):
        if part not in self.cache:
            root = ET.fromstring(self.archive.read(part))
            # Count the modern representation, never both Choice and Fallback.
            self._alternates(root)
            self.cache[part] = root
        return self.cache[part]

    def _alternates(self, node):
        for child in list(node):
            if child.tag == "{" + NS["mc"] + "}AlternateContent":
                branch = child.find("mc:Choice", NS)
                if branch is None:
                    branch = child.find("mc:Fallback", NS)
                index = list(node).index(child)
                node.remove(child)
                if branch is not None:
                    self._alternates(branch)
                    for entry in list(branch):
                        node.insert(index, entry)
                        index += 1
            else:
                self._alternates(child)

    def rels(self, part):
        directory, filename = posixpath.split(part)
        path = posixpath.join(directory, "_rels", filename + ".rels")
        if path not in self.archive.namelist():
            return []
        return [(r.get("Id"), r.get("Type", "").rsplit("/", 1)[-1],
                 r.get("Target").lstrip("/") if r.get("Target", "").startswith("/") else
                 posixpath.normpath(posixpath.join(directory, r.get("Target", ""))))
                for r in self.root(path) if r.get("TargetMode") != "External"]

    def related(self, part, kind):
        return next((target for _, typ, target in self.rels(part) if typ == kind), None) if part else None

    @staticmethod
    def placeholder(shape):
        return shape.find("p:nvSpPr/p:nvPr/p:ph", NS)

    def parent_shape(self, part, shape, by_index):
        marker = self.placeholder(shape)
        if part is None or marker is None:
            return None
        kind = marker.get("type", "obj")
        master_kind = "title" if kind == "ctrTitle" else "body" if kind == "obj" else kind
        for candidate in self.root(part).findall(".//p:sp", NS):
            other = self.placeholder(candidate)
            if other is None:
                continue
            if ((other.get("idx", "0") == marker.get("idx", "0")) if by_index else
                    (other.get("type", "obj") == master_kind)):
                return candidate
        return None

    def theme(self, part, fallback=None):
        target = self.related(part, "theme")
        result = dict(fallback or {})
        if target:
            scheme = self.root(target).find("a:themeElements/a:fontScheme", NS)
            if scheme is not None:
                for kind in ("major", "minor"):
                    node = scheme.find(f"a:{kind}Font", NS)
                    result[kind] = node
        override = self.related(part, "themeOverride")
        if override:
            scheme = self.root(override).find("a:fontScheme", NS)
            if scheme is not None:
                for kind in ("major", "minor"):
                    node = scheme.find(f"a:{kind}Font", NS)
                    if node is not None:
                        result[kind] = node
        return result

    def resolve(self, token, theme, script=None):
        if not token:
            return None
        match = re.fullmatch(r"\+(mj|mn)-(lt|ea|cs)", token)
        if not match:
            return token
        group = theme.get("major" if match[1] == "mj" else "minor")
        if group is None:
            return None
        slot = {"lt": "latin", "ea": "ea", "cs": "cs"}[match[2]]
        node = group.find("a:" + slot, NS)
        if node is not None and node.get("typeface"):
            return node.get("typeface")
        if script:
            for node in group.findall("a:font", NS):
                if node.get("script") == script:
                    return node.get("typeface")
        node = group.find("a:latin", NS)
        return node.get("typeface") if node is not None else None

    def text(self, text, attrs, theme):
        for slot, script in {_script(c, attrs.get("lang", "")) for c in text if not c.isspace()}:
            family = self.resolve(attrs.get(slot) or attrs.get("latin"), theme, script)
            if family:
                self.required.add(Font(family, attrs.get("b") in ("1", "true"),
                                       attrs.get("i") in ("1", "true")))
            else:
                self.unresolved = True

    def body(self, body, base, theme, chain=(), defaults=None):
        for paragraph in body.findall("a:p", NS):
            ppr = paragraph.find("a:pPr", NS)
            level = int(ppr.get("lvl", "0")) if ppr is not None else 0
            attrs = dict(base)
            _list_style(self.presentation.find("p:defaultTextStyle", NS), level, attrs)
            _list_style(self.master_style, level, attrs)
            attrs.update(defaults or {})
            for shape in chain:
                if shape is None:
                    continue
                ref = shape.find("p:style/a:fontRef", NS)
                if ref is not None and ref.get("idx") in ("major", "minor"):
                    prefix = "+mj" if ref.get("idx") == "major" else "+mn"
                    attrs.update(latin=prefix + "-lt", ea=prefix + "-ea", cs=prefix + "-cs")
                _list_style(shape.find("p:txBody/a:lstStyle", NS), level, attrs)
                if shape.find("p:txBody", NS) is not body:
                    for parent_para in shape.findall("p:txBody/a:p", NS):
                        parent_ppr = parent_para.find("a:pPr", NS)
                        parent_level = int(parent_ppr.get("lvl", "0")) if parent_ppr is not None else 0
                        if parent_level == level:
                            _paragraph(parent_ppr, attrs)
                            break
            _list_style(body.find("a:lstStyle", NS), level, attrs)
            _paragraph(ppr, attrs)
            has_text = False
            for run in paragraph:
                if run.tag not in ("{" + NS["a"] + "}r", "{" + NS["a"] + "}fld"):
                    continue
                text = run.findtext("a:t", "", NS)
                has_text |= bool(text.strip())
                effective = dict(attrs)
                _properties(run.find("a:rPr", NS), effective)
                # Explicit symbol fonts apply to symbol/private-use characters.
                if effective.get("sym") and any(unicodedata.category(c) == "Co" for c in text):
                    effective["latin"] = effective["sym"]
                self.text(text, effective, theme)
            for math in paragraph.findall(".//m:oMath", NS):
                if any((n.text or "").strip() for n in math.findall(".//m:t", NS)):
                    has_text = True
                    self.required.add(Font("Cambria Math"))
            if has_text and attrs.get("bullet") == "buChar":
                family = self.resolve(attrs.get("bullet_font") or attrs.get("latin"), theme)
                if family:
                    self.required.add(Font(family))

    def shapes(self, part, layout, master, theme, inherited=False):
        node = self.root(part)
        for shape in node.findall(".//p:sp", NS):
            nv = shape.find("p:nvSpPr/p:cNvPr", NS)
            if nv is not None and nv.get("hidden") in ("1", "true"):
                continue
            marker = self.placeholder(shape)
            if inherited and marker is not None:
                # Prompt text on layouts/masters is not displayed on slides.
                continue
            body = shape.find("p:txBody", NS)
            if body is None:
                continue
            parent = self.parent_shape(layout, shape, True) if layout else None
            ancestor = self.parent_shape(master, parent if parent is not None else shape, False)
            kind = marker.get("type", "obj") if marker is not None else "other"
            style = "titleStyle" if kind in ("title", "ctrTitle") else (
                "bodyStyle" if kind in ("body", "obj") else "otherStyle")
            self.master_style = self.root(master).find("p:txStyles/p:" + style, NS) if master else None
            base = {"latin": "+mn-lt", "ea": "+mn-ea", "cs": "+mn-cs", "b": "0", "i": "0"}
            self.body(body, base, theme, (ancestor, parent, shape))
        # Table formatting can supply fonts even when a cell has no run properties.
        self.master_style = None
        table_bodies = set()
        table_styles = self.related("ppt/presentation.xml", "tableStyles")
        for table in node.findall(".//a:tbl", NS):
            properties = table.find("a:tblPr", NS)
            identifier = properties.findtext("a:tableStyleId", "", NS) if properties is not None else ""
            definition = None
            if table_styles:
                definitions = self.root(table_styles)
                identifier = identifier or definitions.get("def", "")
                definition = next((s for s in definitions if s.get("styleId") == identifier), None)
            rows = table.findall("a:tr", NS)
            for row_index, row in enumerate(rows):
                cells = row.findall("a:tc", NS)
                for column_index, cell in enumerate(cells):
                    defaults = {}
                    regions = ["wholeTbl"]
                    if properties is not None:
                        enabled = lambda key: properties.get(key) in ("1", "true", "on")
                        if enabled("bandRow"):
                            regions.append("band1H" if row_index % 2 == 0 else "band2H")
                        if enabled("bandCol"):
                            regions.append("band1V" if column_index % 2 == 0 else "band2V")
                        if enabled("firstCol") and column_index == 0:
                            regions.append("firstCol")
                        if enabled("lastCol") and column_index == len(cells) - 1:
                            regions.append("lastCol")
                        if enabled("firstRow") and row_index == 0:
                            regions.append("firstRow")
                        if enabled("lastRow") and row_index == len(rows) - 1:
                            regions.append("lastRow")
                    if definition is not None:
                        for region in regions:
                            style = definition.find(f"a:{region}/a:tcTxStyle", NS)
                            if style is None:
                                continue
                            _properties(style, defaults)
                            _properties(style.find("a:font", NS), defaults)
                            for key in ("b", "i"):
                                if defaults.get(key) == "on":
                                    defaults[key] = "1"
                                elif defaults.get(key) == "off":
                                    defaults[key] = "0"
                            ref = style.find("a:fontRef", NS)
                            if ref is not None and ref.get("idx") in ("major", "minor"):
                                prefix = "+mj" if ref.get("idx") == "major" else "+mn"
                                defaults.update(latin=prefix + "-lt", ea=prefix + "-ea", cs=prefix + "-cs")
                    body = cell.find("a:txBody", NS)
                    if body is not None:
                        table_bodies.add(body)
                        self.body(body, {"latin": "+mn-lt", "ea": "+mn-ea", "cs": "+mn-cs"},
                                  theme, defaults=defaults)
        # Other DrawingML text bodies, e.g. diagram text.
        for body in node.findall(".//a:txBody", NS):
            if body not in table_bodies:
                self.body(body, {"latin": "+mn-lt", "ea": "+mn-ea", "cs": "+mn-cs"}, theme)
        for _, kind, target in self.rels(part):
            if kind in ("chart", "diagramData", "diagramDrawing"):
                graphic_theme = self.theme(target, theme)
                self.master_style = None
                graphic = self.root(target)
                defaults = {"latin": "+mn-lt", "ea": "+mn-ea", "cs": "+mn-cs"}
                # Chart txPr sets defaults for data labels/titles/axes.
                for child in graphic:
                    if child.tag.rsplit("}", 1)[-1] == "txPr":
                        _list_style(child.find("a:lstStyle", NS), 0, defaults)
                        _paragraph(child.find("a:p/a:pPr", NS), defaults)
                for para in graphic.findall(".//a:p", NS):
                    wrapper = ET.Element("body")
                    wrapper.append(para)
                    self.body(wrapper, {}, graphic_theme, defaults=defaults)
                if kind == "chart" and not graphic.findall(".//a:t", NS):
                    self.text("chart labels", defaults, graphic_theme)

    def collect(self, numbers):
        self.presentation = self.root("ppt/presentation.xml")
        targets = {key: target for key, _, target in self.rels("ppt/presentation.xml")}
        slides = self.presentation.find("p:sldIdLst", NS)
        fallback_theme = self.theme("ppt/presentation.xml")
        for number in numbers:
            part = targets[slides[number - 1].get("{" + NS["r"] + "}id")]
            layout = self.related(part, "slideLayout")
            master = self.related(layout, "slideMaster")
            theme = self.theme(part, self.theme(layout, self.theme(master, fallback_theme)))
            self.shapes(part, layout, master, theme)
            if layout:
                self.shapes(layout, None, master, theme, inherited=True)
            if master and self.root(part).get("showMasterSp", "1") not in ("0", "false"):
                self.shapes(master, None, master, theme, inherited=True)
        if self.unresolved:
            raise FontCheckError("Some text fonts could not be resolved from the PPTX theme.")
        return self.required


def _fonttools():
    try:
        from fontTools import ttLib
        return ttLib
    except ImportError as exc:
        raise FontCheckError("FontTools is required for font checks. Run "
                             "python -m pip install -r requirements.txt.") from exc


def font_faces(file, embedded=False):
    # FontTools can log irrelevant timestamp warnings while reading system-font
    # headers. Surface actionable missing-font messages instead of those notices.
    logger = logging.getLogger("fontTools.ttLib.tables._h_e_a_d")
    previous_level = logger.level
    logger.setLevel(logging.ERROR)
    try:
        return _read_font_faces(file, embedded)
    finally:
        logger.setLevel(previous_level)


def _read_font_faces(file, embedded=False):
    """Return family/style aliases from real font metadata, including collections."""
    ttLib = _fonttools()
    faces = set()
    stream = io.BytesIO(file) if isinstance(file, bytes) else open(file, "rb")
    with stream:
        magic = stream.read(4)
        stream.seek(0)
        # Uncompressed Embedded OpenType (EOT) wraps an SFNT at the end.
        if embedded and magic not in (b"ttcf", b"OTTO", b"\x00\x01\x00\x00", b"true"):
            data = stream.read()
            if len(data) < 82 or data[34:36] != b"LP":
                return faces
            size = struct.unpack_from("<I", data, 4)[0]
            flags = struct.unpack_from("<I", data, 12)[0]
            if not size or size > len(data) - 82 or flags & 4:  # MTX-compressed EOT
                return faces
            return font_faces(data[-size:], embedded=True)
        collection = ttLib.TTCollection(stream, lazy=True) if magic == b"ttcf" else None
        fonts = collection.fonts if collection is not None else [ttLib.TTFont(stream, lazy=True)]
        try:
            for font in fonts:
                if embedded and (font.get("OS/2") is None or font["OS/2"].fsType & 2):
                    continue
                if not all(tag in font for tag in ("name", "head", "cmap", "maxp")):
                    continue
                family_names = set()
                face_names = set()
                for name in font["name"].names:
                    if name.nameID in (1, 16, 21, 4, 6):
                        value = name.toUnicode()
                        (family_names if name.nameID in (1, 16, 21) else face_names).add(normalized(value))
                selection = font["OS/2"].fsSelection if "OS/2" in font else 0
                mac_style = font["head"].macStyle if "OS/2" not in font else 0
                bold = bool(selection & 32 or mac_style & 1)
                italic = bool(selection & 1 or mac_style & 2)
                styles = {(bold, italic)}
                if "fvar" in font:
                    axes = {a.axisTag: a for a in font["fvar"].axes}
                    weights = {bold}
                    italics = {italic}
                    if "wght" in axes:
                        axis = axes["wght"]
                        weights |= {b for b, w in ((False, 400), (True, 700)) if axis.minValue <= w <= axis.maxValue}
                    if "ital" in axes:
                        axis = axes["ital"]
                        italics |= {i for i in (False, True) if axis.minValue <= int(i) <= axis.maxValue}
                    styles |= {(b, i) for b in weights for i in italics}
                faces.update((name, b, i) for name in family_names | face_names for b, i in styles)
                faces.update((name, False, False) for name in face_names)
        finally:
            for font in fonts:
                font.close()
    return faces


def _mac_font_files():
    cf = ctypes.CDLL("/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation")
    ct = ctypes.CDLL("/System/Library/Frameworks/CoreText.framework/CoreText")
    ct.CTFontManagerCopyAvailableFontURLs.restype = ctypes.c_void_p
    cf.CFArrayGetCount.argtypes = [ctypes.c_void_p]
    cf.CFArrayGetCount.restype = ctypes.c_long
    cf.CFArrayGetValueAtIndex.argtypes = [ctypes.c_void_p, ctypes.c_long]
    cf.CFArrayGetValueAtIndex.restype = ctypes.c_void_p
    cf.CFURLGetFileSystemRepresentation.argtypes = [ctypes.c_void_p, ctypes.c_bool,
                                                   ctypes.c_char_p, ctypes.c_long]
    cf.CFURLGetFileSystemRepresentation.restype = ctypes.c_bool
    cf.CFRelease.argtypes = [ctypes.c_void_p]
    array = ct.CTFontManagerCopyAvailableFontURLs()
    paths = set()
    if not array:
        raise FontCheckError("macOS did not return its available fonts.")
    try:
        for index in range(cf.CFArrayGetCount(array)):
            url = cf.CFArrayGetValueAtIndex(array, index)
            buffer = ctypes.create_string_buffer(32768)
            if cf.CFURLGetFileSystemRepresentation(url, True, buffer, len(buffer)):
                paths.add(Path(os.fsdecode(buffer.value)))
    finally:
        cf.CFRelease(array)
    return paths


def _windows_font_files():
    import winreg
    system = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    user = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Microsoft/Windows/Fonts"
    paths = set()
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for location in (r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
                         r"SOFTWARE\Microsoft\Windows\CurrentVersion\Fonts"):
            try:
                with winreg.OpenKey(hive, location) as key:
                    for index in range(winreg.QueryInfoKey(key)[1]):
                        _, value, _ = winreg.EnumValue(key, index)
                        path = Path(os.path.expandvars(value))
                        paths.add(path if path.is_absolute() else
                                  (user if hive == winreg.HKEY_CURRENT_USER else system) / path)
            except FileNotFoundError:
                pass
    return paths


def _fontconfig_files():
    result = subprocess.run(["fc-list", "--format", "%{file}\\n"], capture_output=True,
                            check=True, timeout=20)
    return {Path(os.fsdecode(line)) for line in result.stdout.splitlines() if line}


def _bundled_font_files(executable):
    if executable is None:
        return set()
    path = Path(executable).resolve()
    bases = {path.parent.parent}
    if sys.platform == "darwin":
        # Homebrew's soffice symlink can point to a shell wrapper, not the app.
        if path.suffix == ".sh":
            with path.open(errors="replace") as handle:
                launcher = handle.read(4096)
            for match in re.finditer(r"['\"]([^'\"]+\.app/Contents)/MacOS/soffice['\"]", launcher):
                bases.add(Path(match[1]))
    roots = [base / suffix for base in bases for suffix in (
        "Resources/fonts", "Resources/resource/common/fonts", "share/fonts", "share/resource/common/fonts")]
    return {path for root in roots if root.is_dir() for path in root.rglob("*")
            if path.is_file() and path.suffix.lower() in (".ttf", ".ttc", ".otf", ".otc")}


@lru_cache(maxsize=4)
def available_fonts(app, executable=None):
    """Cache metadata per renderer for a single CLI run, not across invocations."""
    ttLib = _fonttools()
    try:
        paths = (_mac_font_files() if sys.platform == "darwin" else
                 _windows_font_files() if sys.platform == "win32" else _fontconfig_files())
        if app == "libreoffice":
            paths |= _bundled_font_files(executable)
        faces = set()
        for path in paths:
            try:
                faces |= font_faces(path)
            except (OSError, ValueError, KeyError, struct.error, ttLib.TTLibError, UnicodeError):
                # System collections may include bitmap/Type1 fonts unsupported by
                # the metadata reader. They cannot satisfy an SFNT font check.
                continue
        if not faces:
            raise FontCheckError("No readable fonts were found in the renderer's font environment.")
        return frozenset(faces)
    except (OSError, subprocess.SubprocessError) as exc:
        hint = " Install fontconfig (fc-list) on Linux." if sys.platform not in ("darwin", "win32") else ""
        raise FontCheckError(f"Cannot inspect installed fonts: {exc}.{hint}") from exc


def missing_fonts(source, numbers, app, executable=None):
    try:
        with zipfile.ZipFile(source) as archive:
            deck = DeckFonts(archive)
            required = deck.collect(numbers)
            if not required:
                return []
            embedded = set()
            # The renderer must be instructed to use fonts embedded in this deck.
            if deck.presentation.get("embedTrueTypeFonts", "0") in ("1", "true"):
                targets = {key: target for key, _, target in deck.rels("ppt/presentation.xml")}
                for font in deck.presentation.findall("p:embeddedFontLst/p:embeddedFont", NS):
                    for face in font:
                        target = targets.get(face.get("{" + NS["r"] + "}id"))
                        if target:
                            try:
                                embedded |= font_faces(archive.read(target), embedded=True)
                            except (OSError, ValueError, KeyError, struct.error, _fonttools().TTLibError):
                                pass  # Broken/unsupported embedded fonts are not available.
        available = available_fonts(app, str(executable) if executable is not None else None) | embedded
        unique = {(normalized(f.family), f.bold, f.italic): f for f in sorted(required, reverse=True)}
        return sorted((f for key, f in unique.items() if key not in available),
                      key=lambda f: (f.family.casefold(), f.bold, f.italic))
    except (zipfile.BadZipFile, KeyError, ET.ParseError, ValueError, IndexError) as exc:
        raise FontCheckError(f"Cannot resolve fonts in {source.name}: {exc}") from exc
