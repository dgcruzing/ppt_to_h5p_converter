"""Font preflight tests use synthetic decks and generated, valid miniature fonts."""

import contextlib
import importlib
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "build_script"))
fonts = importlib.import_module("pptx_to_h5p_converter cli v1_3 verify fonts")
import test_cli as harness

cli = harness.cli
NS = fonts.NS


def shape(text="body", kind="body", index=1, run_properties="", list_style="", paragraph="", style=""):
    marker = f'<p:ph type="{kind}" idx="{index}"/>' if kind else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="1" name="Test"/><p:nvPr>{marker}</p:nvPr></p:nvSpPr>'
            f'{style}<p:txBody><a:bodyPr/><a:lstStyle>{list_style}</a:lstStyle><a:p>'
            f'<a:pPr>{paragraph}</a:pPr><a:r><a:rPr>{run_properties}</a:rPr><a:t>{text}</a:t></a:r>'
            '</a:p></p:txBody></p:sp>')


def deck(path, content=None, layout_extra="", master_extra="", embedded=None, hidden_second=None):
    declarations = ' '.join(f'xmlns:{p}="{uri}"' for p, uri in NS.items())
    def xml(tag, inner):
        return f'<{tag} {declarations}>{inner}</{tag}>'
    def relationships(entries):
        return ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                ''.join(f'<Relationship Id="{key}" Type="{NS["r"]}/{typ}" Target="{target}"/>'
                        for key, typ, target in entries) + '</Relationships>')
    presentation_rels = [('s1', 'slide', 'slides/chosen.xml')]
    ids = '<p:sldId id="256" r:id="s1"/>'
    if hidden_second is not None:
        presentation_rels.append(('s2', 'slide', 'slides/hidden.xml'))
        ids += '<p:sldId id="257" r:id="s2"/>'
    embedded_list = ''
    if embedded is not None:
        family, data = embedded
        embedded_list = (f'<p:embeddedFontLst><p:embeddedFont><p:font typeface="{family}"/>'
                         '<p:regular r:id="f1"/></p:embeddedFont></p:embeddedFontLst>')
        presentation_rels.append(('f1', 'font', 'fonts/embedded.fntdata'))
    with zipfile.ZipFile(path, 'w') as archive:
        presentation = xml('p:presentation', '<p:sldIdLst>' + ids + '</p:sldIdLst>' + embedded_list +
                           '<p:defaultTextStyle><a:lvl1pPr><a:defRPr><a:latin typeface="+mn-lt"/>'
                           '</a:defRPr></a:lvl1pPr></p:defaultTextStyle>')
        if embedded is not None:
            presentation = presentation.replace('<p:presentation ', '<p:presentation embedTrueTypeFonts="1" ', 1)
            archive.writestr('ppt/fonts/embedded.fntdata', data)
        archive.writestr('ppt/presentation.xml', presentation)
        archive.writestr('ppt/_rels/presentation.xml.rels', relationships(presentation_rels))
        for name, slide_content in [('chosen', content if content is not None else shape()),
                                     *([('hidden', hidden_second)] if hidden_second is not None else [])]:
            slide_xml = xml('p:sld', '<p:cSld><p:spTree>' + slide_content + '</p:spTree></p:cSld>')
            if name == 'hidden':
                slide_xml = slide_xml.replace('<p:sld ', '<p:sld show="0" ', 1)
            archive.writestr(f'ppt/slides/{name}.xml', slide_xml)
            archive.writestr(f'ppt/slides/_rels/{name}.xml.rels', relationships([
                ('l1', 'slideLayout', '../slideLayouts/layout.xml')]))
        archive.writestr('ppt/slideLayouts/layout.xml', xml('p:sldLayout', '<p:cSld><p:spTree>' +
            shape('layout prompt', list_style='<a:lvl1pPr><a:defRPr><a:latin typeface="Layout Body"/>'
                  '</a:defRPr></a:lvl1pPr>') + layout_extra + '</p:spTree></p:cSld>'))
        archive.writestr('ppt/slideLayouts/_rels/layout.xml.rels', relationships([
            ('m1', 'slideMaster', '../slideMasters/master.xml')]))
        archive.writestr('ppt/slideMasters/master.xml', xml('p:sldMaster', '<p:cSld><p:spTree>' +
            shape('master prompt', kind='title', index=0) + shape('master prompt') + master_extra +
            '</p:spTree></p:cSld><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr><a:latin '
            'typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr>'
            '<a:defRPr><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:bodyStyle></p:txStyles>'))
        archive.writestr('ppt/slideMasters/_rels/master.xml.rels', relationships([
            ('t1', 'theme', '../theme/theme.xml')]))
        archive.writestr('ppt/theme/theme.xml', xml('a:theme', '<a:themeElements><a:fontScheme name="Test">'
            '<a:majorFont><a:latin typeface="Theme Heading"/></a:majorFont>'
            '<a:minorFont><a:latin typeface="Theme Body"/><a:ea typeface=""/><a:cs typeface=""/>'
            '<a:font script="Arab" typeface="Theme Arabic"/>'
            '<a:font script="Jpan" typeface="Unused Japanese"/></a:minorFont>'
            '</a:fontScheme></a:themeElements>'))


def tiny_font(family="Test Family", bold=False, italic=False):
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen
    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder(['.notdef', 'space', 'a'])
    builder.setupCharacterMap({32: 'space', 97: 'a'})
    builder.setupGlyf({name: TTGlyphPen(None).glyph() for name in ['.notdef', 'space', 'a']})
    builder.setupHorizontalMetrics({name: (500, 0) for name in ['.notdef', 'space', 'a']})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    style = fonts.Font(family, bold, italic).style
    ps = family.replace(' ', '') + '-' + style.replace(' ', '')
    builder.setupNameTable({'familyName': family, 'styleName': style,
                           'fullName': family + ' ' + style, 'psName': ps, 'uniqueFontIdentifier': ps})
    builder.setupOS2(sTypoAscender=800, sTypoDescender=-200, usWinAscent=800, usWinDescent=200,
                     usWeightClass=700 if bold else 400,
                     fsSelection=(32 if bold else 0) | (1 if italic else 0) | (64 if not bold and not italic else 0))
    builder.setupPost()
    builder.setupMaxp()
    builder.font['head'].macStyle = (1 if bold else 0) | (2 if italic else 0)
    stream = io.BytesIO()
    builder.save(stream)
    return stream.getvalue()


class FontTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        self.source = self.directory / 'slides.pptx'
        fonts.available_fonts.cache_clear()
        self.addCleanup(fonts.available_fonts.cache_clear)

    def required(self, numbers=(1,)):
        with zipfile.ZipFile(self.source) as archive:
            return fonts.DeckFonts(archive).collect(numbers)

    def test_slide_layout_master_theme_precedence_and_unused_fonts(self):
        deck(self.source, shape('Title', kind='title', index=0) +
             shape('Inherited body') + shape('Explicit', run_properties='<a:latin typeface="Run Face"/>',
                 paragraph='<a:defRPr b="1" i="1"><a:latin typeface="Paragraph Face"/></a:defRPr>'),
             layout_extra=shape('banner', kind=None, run_properties='<a:latin typeface="Layout Banner"/>'),
             master_extra=shape('copyright', kind=None, run_properties='<a:latin typeface="Master Footer"/>'))
        self.assertEqual(self.required(), {fonts.Font('Theme Heading'), fonts.Font('Layout Body'),
            fonts.Font('Run Face', True, True), fonts.Font('Layout Banner'), fonts.Font('Master Footer')})

    def test_hidden_slides_and_relationship_order(self):
        deck(self.source, hidden_second=shape('hidden', run_properties='<a:latin typeface="Hidden Font"/>'))
        self.assertEqual(cli.eligible_slides(self.source), [1])
        self.assertNotIn(fonts.Font('Hidden Font'), self.required(cli.eligible_slides(self.source)))
        self.assertIn(fonts.Font('Hidden Font'), self.required(cli.eligible_slides(self.source, True)))

    def test_only_active_bullets_math_and_used_script_fonts(self):
        content = shape('English', paragraph='<a:buFont typeface="Unused Bullet"/><a:buNone/>')
        content += shape('Bullet', paragraph='<a:buFont typeface="Wingdings"/><a:buChar char="§"/>')
        content += shape('مرحبا', kind=None)
        content += shape().replace('<a:r><a:rPr></a:rPr><a:t>body</a:t></a:r>',
            '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>')
        deck(self.source, content)
        self.assertEqual(self.required(), {fonts.Font('Layout Body'), fonts.Font('Wingdings'),
                                          fonts.Font('Theme Arabic'), fonts.Font('Cambria Math')})

    def test_fontref_tables_empty_text_and_alternate_fallback(self):
        content = shape('styled', kind=None,
                        style='<p:style><a:fontRef idx="major"/></p:style>')
        content += shape('', run_properties='<a:latin typeface="Empty Font"/>')
        content += ('<p:graphicFrame><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:rPr b="1">'
                    '<a:latin typeface="Table Font"/></a:rPr><a:t>cell</a:t></a:r></a:p>'
                    '</a:txBody></a:tc></a:tr></a:tbl></p:graphicFrame>')
        content += ('<mc:AlternateContent><mc:Choice Requires="a">' + shape('modern') +
                    '</mc:Choice><mc:Fallback>' + shape('old', run_properties='<a:latin typeface="Old Font"/>') +
                    '</mc:Fallback></mc:AlternateContent>')
        deck(self.source, content)
        self.assertEqual(self.required(), {fonts.Font('Theme Heading'), fonts.Font('Layout Body'),
                                          fonts.Font('Table Font', True)})

    def test_font_aliases_styles_collections_and_embedded_fonts(self):
        from fontTools.ttLib import TTCollection, TTFont
        regular = tiny_font()
        bold = tiny_font(bold=True)
        self.assertIn((fonts.normalized('TestFamily-Regular'), False, False), fonts.font_faces(regular))
        self.assertNotIn((fonts.normalized('Test Family'), True, False), fonts.font_faces(regular))
        collection = TTCollection()
        collection.fonts = [TTFont(io.BytesIO(regular)), TTFont(io.BytesIO(bold))]
        data = io.BytesIO()
        collection.save(data)
        collection.close()
        self.assertIn((fonts.normalized('Test Family'), True, False), fonts.font_faces(data.getvalue()))
        deck(self.source, shape('embedded', run_properties='<a:latin typeface="Test Family"/>'),
             embedded=('Test Family', regular))
        with mock.patch.object(fonts, 'available_fonts', return_value=frozenset()):
            self.assertEqual(fonts.missing_fonts(self.source, [1], 'libreoffice'), [])
        deck(self.source, shape('broken', run_properties='<a:latin typeface="Test Family"/>'),
             embedded=('Test Family', b'not a font'))
        with mock.patch.object(fonts, 'available_fonts', return_value=frozenset()):
            self.assertEqual(fonts.missing_fonts(self.source, [1], 'libreoffice'), [fonts.Font('Test Family')])

    def test_bundled_fonts_count_and_fontconfig_failure_is_actionable(self):
        executable = self.directory / 'LibreOffice/program/soffice'
        executable.parent.mkdir(parents=True)
        bundled = executable.parent.parent / 'share/fonts/truetype'
        bundled.mkdir(parents=True)
        (bundled / 'font.ttf').write_bytes(tiny_font())
        (bundled / 'unsupported.ttf').write_bytes(b'not a font')
        with mock.patch.object(fonts.sys, 'platform', 'linux'), \
                mock.patch.object(fonts, '_fontconfig_files', return_value=set()):
            self.assertIn((fonts.normalized('Test Family'), False, False),
                          fonts.available_fonts('libreoffice', str(executable)))
        fonts.available_fonts.cache_clear()
        with mock.patch.object(fonts.sys, 'platform', 'linux'), \
                mock.patch.object(fonts.subprocess, 'run', side_effect=FileNotFoundError('fc-list')):
            with self.assertRaisesRegex(fonts.FontCheckError, 'Install fontconfig'):
                fonts.available_fonts('libreoffice', str(executable))

    def test_homebrew_wrapper_uses_the_selected_apps_fonts(self):
        executable = self.directory / 'soffice.wrapper.sh'
        base = self.directory / 'Custom LibreOffice.app/Contents'
        bundled = base / 'Resources/fonts/truetype'
        bundled.mkdir(parents=True)
        (bundled / 'font.ttf').write_bytes(tiny_font())
        executable.write_text(f"#!/bin/sh\n'{base}/MacOS/soffice' \"$@\"\n")
        with mock.patch.object(fonts.sys, 'platform', 'darwin'):
            self.assertEqual(fonts._bundled_font_files(executable), {bundled / 'font.ttf'})

    def test_chart_rich_text_and_labels(self):
        deck(self.source)
        # Link the chart through layout relationships, retaining its master.
        rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="chart" Type="{NS["r"]}/chart" Target="../charts/chart.xml"/>'
                f'<Relationship Id="m1" Type="{NS["r"]}/slideMaster" '
                'Target="../slideMasters/master.xml"/></Relationships>')
        with zipfile.ZipFile(self.source) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        entries['ppt/slideLayouts/_rels/layout.xml.rels'] = rels.encode()
        entries['ppt/charts/chart.xml'] = (f'<chart xmlns:a="{NS["a"]}"><rich><a:p><a:r>'
            '<a:rPr><a:latin typeface="Chart Font"/></a:rPr><a:t>Chart</a:t></a:r></a:p></rich></chart>').encode()
        with zipfile.ZipFile(self.source, 'w') as archive:
            for name, data in entries.items():
                archive.writestr(name, data)
        self.assertIn(fonts.Font('Chart Font'), self.required())

    def test_table_style_inheritance_and_cell_override(self):
        def cell(text, properties=''):
            return ('<a:tc><a:txBody><a:p><a:r><a:rPr>' + properties +
                    '</a:rPr><a:t>' + text + '</a:t></a:r></a:p></a:txBody></a:tc>')
        table = ('<p:graphicFrame><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>custom</a:tableStyleId>'
                 '</a:tblPr><a:tr>' + cell('Header') + '</a:tr><a:tr>' + cell('Body') +
                 cell('Override', '<a:latin typeface="Explicit Cell"/>') + '</a:tr></a:tbl></p:graphicFrame>')
        deck(self.source, table)
        with zipfile.ZipFile(self.source) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        rels = entries['ppt/_rels/presentation.xml.rels'].decode().replace('</Relationships>',
            f'<Relationship Id="tables" Type="{NS["r"]}/tableStyles" Target="tableStyles.xml"/></Relationships>')
        entries['ppt/_rels/presentation.xml.rels'] = rels.encode()
        entries['ppt/tableStyles.xml'] = (f'<a:tblStyleLst xmlns:a="{NS["a"]}" def="custom">'
            '<a:tblStyle styleId="custom"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="major"/>'
            '</a:tcTxStyle></a:wholeTbl><a:firstRow><a:tcTxStyle b="on"><a:font>'
            '<a:latin typeface="Table Header"/></a:font></a:tcTxStyle></a:firstRow>'
            '</a:tblStyle></a:tblStyleLst>').encode()
        with zipfile.ZipFile(self.source, 'w') as archive:
            for name, data in entries.items():
                archive.writestr(name, data)
        self.assertEqual(self.required(), {fonts.Font('Table Header', True),
                                          fonts.Font('Theme Heading'), fonts.Font('Explicit Cell')})

    def test_empty_deck_needs_no_inventory_and_missing_names_are_deduplicated(self):
        deck(self.source, '')
        with mock.patch.object(fonts, 'available_fonts', side_effect=AssertionError('no text')):
            self.assertEqual(fonts.missing_fonts(self.source, [1], 'libreoffice'), [])
        deck(self.source, shape('A', run_properties='<a:latin typeface="Missing Font"/>') +
             shape('B', run_properties='<a:latin typeface="missing font"/>'))
        with mock.patch.object(fonts, 'available_fonts', return_value=frozenset()):
            self.assertEqual(fonts.missing_fonts(self.source, [1], 'libreoffice'), [fonts.Font('Missing Font')])

    def test_override_silences_and_skips_check_without_hiding_export_errors(self):
        args = cli.create_parser().parse_args([str(harness.SOURCE), '--ignore-missing-fonts'])
        destination = self.directory / 'result.h5p'
        with mock.patch.object(cli, 'missing_fonts', side_effect=AssertionError('must skip')) as check, \
                mock.patch.object(cli, 'export_pdf', side_effect=cli.ConversionError('export failed')), \
                contextlib.redirect_stderr(io.StringIO()) as stderr:
            with self.assertRaisesRegex(cli.ConversionError, 'export failed'):
                cli.convert(harness.SOURCE, destination, 'libreoffice', None, args)
            check.assert_not_called()
        self.assertEqual(stderr.getvalue(), '')
        self.assertFalse(destination.exists())

    def test_block_before_rendering_creating_output_or_replacing_existing(self):
        args = cli.create_parser().parse_args([str(harness.SOURCE), '--force'])
        destination = self.directory / 'output/result.h5p'
        for existing in (False, True):
            if existing:
                destination.parent.mkdir()
                destination.write_bytes(b'original')
            with mock.patch.object(cli, 'missing_fonts', return_value=[fonts.Font('Missing Font', True)]), \
                    mock.patch.object(cli, 'export_pdf') as export:
                with self.assertRaisesRegex(cli.ConversionError, 'Missing Font — Bold') as error:
                    cli.convert(harness.SOURCE, destination, 'libreoffice', None, args)
                self.assertIn('Install these fonts', str(error.exception))
                self.assertIn('--ignore-missing-fonts', str(error.exception))
                export.assert_not_called()
            self.assertEqual(destination.exists(), existing)
            if existing:
                self.assertEqual(destination.read_bytes(), b'original')
            else:
                self.assertFalse(destination.parent.exists())

    def test_folder_skips_missing_continues_and_override_succeeds(self):
        sources = self.directory / 'sources'
        sources.mkdir()
        for name in ('a', 'b'):
            shutil.copyfile(harness.SOURCE, sources / (name + '.pptx'))
        output = self.directory / 'courses'
        def check(source, *_):
            return [fonts.Font('Missing Font')] if source.stem == 'a' else []
        for override in (False, True):
            with mock.patch.object(cli, 'missing_fonts', side_effect=check) as preflight, \
                    mock.patch.object(cli, 'export_pdf') as export, \
                    mock.patch.object(cli, 'render_pdf', return_value=[(1, harness.REFERENCE_IMAGE)]), \
                    mock.patch.object(cli, 'require_module'), mock.patch.object(cli, 'find_libreoffice'), \
                    contextlib.redirect_stderr(io.StringIO()) as stderr, contextlib.redirect_stdout(io.StringIO()):
                code = cli.main([str(sources), '-o', str(output), '--force', '--app', 'libreoffice'] +
                                (['--ignore-missing-fonts'] if override else []))
                self.assertEqual(code, 0 if override else 1)
                self.assertEqual(export.call_count, 2 if override else 1)
                self.assertEqual(preflight.call_count, 0 if override else 2)
                self.assertEqual('Missing fonts' in stderr.getvalue(), not override)
            self.assertTrue((output / 'b.h5p').exists())
            self.assertEqual((output / 'a.h5p').exists(), override)

    def test_unverifiable_inventory_blocks_with_a_distinct_message(self):
        args = cli.create_parser().parse_args([str(harness.SOURCE)])
        with mock.patch.object(cli, 'missing_fonts', side_effect=fonts.FontCheckError('no inventory')):
            with self.assertRaisesRegex(cli.ConversionError, 'Font check failed: no inventory'):
                cli.convert(harness.SOURCE, self.directory / 'result.h5p', 'libreoffice', None, args)


@unittest.skipUnless(os.environ.get('H5P_RUN_INTEGRATION') == '1', 'Set H5P_RUN_INTEGRATION=1')
class FontIntegrationTests(unittest.TestCase):
    def test_actual_cli_blocks_missing_font_and_converts_with_override_or_available_fonts(self):
        with tempfile.TemporaryDirectory(prefix='h5p-font-integration-') as temporary:
            directory = Path(temporary)
            source = directory / 'missing fonts.pptx'
            destination = directory / 'result.h5p'
            with zipfile.ZipFile(harness.SOURCE) as original, zipfile.ZipFile(source, 'w') as output:
                for item in original.infolist():
                    data = original.read(item.filename)
                    if item.filename.endswith('.xml'):
                        data = data.replace(b'typeface="Open Sans"', b'typeface="H5P Deliberately Missing Font"')
                    output.writestr(item, data)
            command = [sys.executable, str(harness.SCRIPT), str(source), '-o', str(destination),
                       '--app', 'libreoffice', '--resolution', 'low', '--timeout', '60']
            blocked = subprocess.run(command, capture_output=True, timeout=90)
            self.assertEqual(blocked.returncode, 1, blocked.stderr.decode())
            self.assertIn(b'Missing fonts:', blocked.stderr)
            self.assertIn(b'H5P Deliberately Missing Font', blocked.stderr)
            self.assertIn(b'Install these fonts', blocked.stderr)
            self.assertFalse(destination.exists())
            allowed = subprocess.run(command + ['--ignore-missing-fonts'], capture_output=True, timeout=90)
            self.assertEqual(allowed.returncode, 0, allowed.stderr.decode())
            self.assertNotIn(b'Missing fonts:', allowed.stderr)
            self.assertNotIn(b'H5P Deliberately Missing Font', allowed.stderr)
            with zipfile.ZipFile(destination) as archive:
                self.assertIsNone(archive.testzip())
                slides = json.loads(archive.read('content/content.json'))['presentation']['slides']
                self.assertEqual(len(slides), 7)
            # LibreOffice bundles Liberation Sans, so exercise default preflight
            # success without relying on the user's separately installed fonts.
            with zipfile.ZipFile(harness.SOURCE) as original, zipfile.ZipFile(source, 'w') as output:
                for item in original.infolist():
                    data = original.read(item.filename)
                    if item.filename.endswith('.xml'):
                        data = data.replace(b'typeface="Open Sans"', b'typeface="Liberation Sans"')
                    output.writestr(item, data)
            checked = subprocess.run(command + ['--force'], capture_output=True, timeout=90)
            self.assertEqual(checked.returncode, 0, checked.stderr.decode())
            self.assertNotIn(b'Missing fonts:', checked.stderr)
            with zipfile.ZipFile(destination) as archive:
                self.assertIsNone(archive.testzip())
                self.assertEqual(len(json.loads(archive.read('content/content.json'))['presentation']['slides']), 7)


if __name__ == '__main__':
    unittest.main()
