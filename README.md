# PowerPoint to H5P Converter

Convert PowerPoint `.pptx` presentations into image-based H5P Course
Presentations. Choose the **Windows GUI** for file selection with buttons or the
**cross-platform CLI** for terminal use and folder conversion.

- [Version differences](#version-differences)
- [Windows GUI: v1.1 and v1.2](#windows-gui-v11-and-v12)
- [Command line: v1.3](#command-line-v13)
- [H5P output and compatibility](#h5p-output-and-compatibility)
- [Development and verification](#development-and-verification)

## Version differences

| Behavior | v1.1 GUI | v1.2 GUI | v1.3 CLI |
| --- | --- | --- | --- |
| Interface | Tkinter file-selection window | Tkinter window with resolution buttons | Terminal arguments |
| Operating system | Windows | Windows | Windows, Mac, Linux |
| Rendering app | Microsoft PowerPoint | Microsoft PowerPoint | PowerPoint by default on Windows; LibreOffice elsewhere; selectable with `--app` |
| Rendering sequence | PowerPoint → PNG → H5P | PowerPoint → PNG → H5P | PowerPoint → PNG → H5P, or LibreOffice → temporary PDF → PNG → H5P |
| Image resolution | Fixed 3840×2160 | High 1920×1080 (default) or low 1280×720 | High 1920×1080 (default) or low 1280×720 |
| Aspect ratio | Requests the fixed output dimensions directly | Requests the selected output dimensions directly | Preserves proportions, centering on a white canvas |
| Input | One presentation | One presentation | One presentation or a folder, without recursion |
| Hidden slides | Included | Included | Skipped by default; `--include-hidden` includes them |
| Existing output | No explicit overwrite guard in conversion code | No explicit overwrite guard in conversion code | Protected unless `--force` is supplied |
| Export timeout | None | None | 30 seconds per presentation; adjustable with `--timeout` |
| Diagnostics | Dialogs and a log file | Dialogs and a log file | stderr, exit codes, optional tracebacks with `--verbose` |

The three versions are separate scripts. v1.2 adds resolution selection to the
original GUI; v1.3 adds a command-line interface, cross-platform app selection,
folder processing, hidden-slide controls, and safer output handling. The GUI
scripts have not been changed by the CLI work. v1.3 also explicitly declares the
`H5P.Image 1.1` dependency used by each slide; the GUI archives omit that declaration.

Exact source filenames and Windows executable build instructions are listed in
[build_script/README.md](build_script/README.md).

## Windows GUI: v1.1 and v1.2

### Requirements and setup

The GUI scripts require Windows, installed and activated desktop Microsoft
PowerPoint, Python with Tkinter, Pillow, and pywin32. The source setup below uses
Python 3.11 or newer. Tkinter is supplied with the standard Windows Python
installer; it is not installed through pip.

From the repository root in PowerShell:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m tkinter
```

The final command opens a small Tkinter verification window; close it before
starting the converter. The shared requirements install dependencies for both
the GUI and CLI, including pywin32 only on Windows. For a GUI-only environment,
Pillow and pywin32 are the only pip dependencies.

### Launch and convert

Start v1.2 for resolution selection:

```powershell
.\.venv\Scripts\python.exe "build_script/pptx_to_h5p_converter buttons v1_2.py"
```

Or start v1.1 for fixed 3840×2160 exports:

```powershell
.\.venv\Scripts\python.exe "build_script/ppt_to_h5p_converter v1.1.py"
```

1. Browse to the input `.pptx` file.
2. Choose the destination `.h5p` file.
3. In v1.2, select **High** or **Low** resolution. v1.1 has no resolution control.
4. Click **Convert**, then wait for the success or error dialog.
5. Upload the result to an H5P destination with the required libraries installed.

If you already have a Windows GUI `.exe` from a project release, double-click it
to open the same interface. A bundled executable does not require a separate
Python installation, but still requires Microsoft PowerPoint. No `.exe` is
included in this checkout; see the [build instructions](build_script/README.md#building-a-windows-gui-executable).

### GUI troubleshooting

- Run PowerPoint once to finish activation and initial setup.
- Check that both the source and destination are accessible and the output
  directory is writable.
- Both GUI versions write `ppt_to_h5p_converter.log` in the current working
  directory. When using the commands above, that is the repository root; it is
  not necessarily the directory containing an executable.
- The GUI versions call `PowerPoint.Application.Quit()` after conversion and can
  share an existing PowerPoint instance. Save and close other presentations
  before using them.
- v1.1 exports at 4K, which can produce larger files. Use v1.2's low-resolution
  setting when smaller images are sufficient.
- GUI conversion runs synchronously without a timeout or cancellation control;
  the window can appear unresponsive while PowerPoint is exporting.

## Command line: v1.3

### Requirements and setup

Use Python 3.11 or newer. Windows defaults to installed desktop Microsoft
PowerPoint. Mac and Linux use [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/).
Windows users can also choose LibreOffice with `--app libreoffice`.

On Mac, Homebrew users can install LibreOffice with
`brew install --cask libreoffice`. Then, from the repository root:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

On Windows, use the PowerShell environment setup from the GUI section. Replace
`python3` in the examples below with `.\.venv\Scripts\python.exe`.
Quote the script path because its name contains spaces.

### Launch and convert

```bash
# Convert the bundled example into the converted directory
python3 "build_script/pptx_to_h5p_converter cli v1_3.py" "tests/fixtures/introduction_to_h5p/source.pptx" -o "converted/introduction_to_h5p.h5p"

# Select a destination and smaller images
python3 "build_script/pptx_to_h5p_converter cli v1_3.py" "slides.pptx" -o "output/course.h5p" --resolution low

# All PPTX files immediately inside a folder
python3 "build_script/pptx_to_h5p_converter cli v1_3.py" "presentations" -o "converted"

# Choose LibreOffice, include hidden slides, and allow 90 seconds for export
python3 "build_script/pptx_to_h5p_converter cli v1_3.py" "slides.pptx" --app libreoffice --include-hidden --timeout 90

# Override LibreOffice discovery only if needed
python3 "build_script/pptx_to_h5p_converter cli v1_3.py" "slides.pptx" --app libreoffice --libreoffice-path "/Applications/LibreOffice.app/Contents/MacOS/soffice"

# Continue with substitute fonts and suppress the missing-font warning
python3 "build_script/pptx_to_h5p_converter cli v1_3.py" "slides.pptx" --ignore-missing-fonts

# List every option
python3 "build_script/pptx_to_h5p_converter cli v1_3.py" --help
```

| Option | Behavior |
| --- | --- |
| `input` | A `.pptx` file or folder. Folder processing is not recursive; uppercase extensions are supported and `~$` lock files are ignored. |
| `-o`, `--output` | H5P filename for single input; directory for folder input. Parent directories are created. Defaults to beside each source. |
| `--app auto/powerpoint/libreoffice` | Default `auto`: PowerPoint on Windows; LibreOffice elsewhere. Explicit PowerPoint selection requires Windows. |
| `--libreoffice-path PATH` | Override the LibreOffice executable location. Only valid when LibreOffice is selected. |
| `--include-hidden` | Include hidden slides. By default they are skipped; all-hidden decks fail without creating output. |
| `--ignore-missing-fonts` | Skip font verification and its warnings, allowing conversion with substitute fonts. Other conversion errors remain visible. |
| `--timeout SECONDS` | Positive number, default **30** per presentation's app export. PDF rasterization and H5P packaging are outside this timeout. |
| `--resolution high/low` | Default `high`: 1920×1080; `low`: 1280×720. Slides preserve proportions and are centered on a white canvas. |
| `--force` | Replace existing output files after successful conversion. Without this flag, existing outputs are reported as failures. |
| `--verbose` | Print diagnostics and tracebacks to stderr. |

Folder processing uses filename order, continues after individual failures, and
prints a final success/failure count. Conflicting output filenames are rejected
before processing. Successful output paths go to stdout; status and errors go to
stderr. Exit codes: `0` all succeeded, `1` conversion failure or existing output,
`2` invalid arguments/input or no matching files, `130` interrupted.

### Missing-font check

Before rendering, the CLI resolves the fonts and bold/italic styles used by the
included slides, including inherited layout/master/theme styles, static template
text, active bullets, equations, table text, and rich chart/diagram text. It does
not count notes, excluded hidden slides, empty text runs, unused theme script
fonts, or placeholder prompt text.

If a required font or style is missing, conversion stops before creating output
and prints a message such as:

```text
Missing fonts:
  - Tw Cen MT — Regular
  - Calibri — Bold

Install these fonts on this computer before converting. Missing fonts may change text size, spacing, and line breaks.
To continue with substitute fonts, add --ignore-missing-fonts.
```

Install the listed fonts for the user running the converter, then start the CLI
again. Office fonts stored privately inside PowerPoint may need installing for
other applications to see them. The converter does not install or download fonts.
Use `--ignore-missing-fonts` to skip this check silently when substitution is
acceptable. This override also allows conversion when the font inventory cannot
be inspected; it does not suppress rendering, timeout, or output errors.

Folder conversion skips presentations with missing fonts, processes the remaining
files, and exits with code `1` if any fail. With the override, it attempts every
presentation. Existing output is protected even when `--force` is supplied until
font verification and conversion succeed.

The check reads real font metadata using FontTools. On Mac it uses CoreText's
available font files; on Windows it reads the installed-font registry; Linux
requires `fc-list` from fontconfig. LibreOffice's own bundled fonts and supported,
readable embedded fonts are included. A replacement font such as Carlito does
not count as the original Calibri. Corrupt, restricted, obfuscated, or unsupported
embedded fonts are not assumed usable. Font inventory is cached within the
converter process, so rerun the CLI after installing fonts.

This is a font availability check, not a guarantee of PowerPoint-equivalent
rendering. It does not check individual glyph coverage, text baked into images,
or every font chosen implicitly by complex chart/SmartArt rendering. Installed
fonts do not fix separate equation-import or layout issues in LibreOffice.

### CLI troubleshooting

- If LibreOffice is missing, install it or use `--libreoffice-path` pointing to
  `soffice` (Mac/Linux) or `soffice.com` / `soffice.exe` (Windows). Discovery checks
  `PATH` and standard installation locations.
- Windows PowerPoint requires activation and initial setup. Alternatively select
  `--app libreoffice`. The CLI does not silently switch apps after a failure.
- For a font-check failure on Linux, ensure fontconfig (`fc-list`) is installed.
  For other font errors, install the listed faces/styles or explicitly opt into
  substitutions with `--ignore-missing-fonts`.
- If export times out, try `--timeout 90`. Converter-owned processes and temporary
  files are cleaned up. Existing PowerPoint sessions are not terminated; if
  automation stalled in a reused session, an opened copy may need closing manually.
- On Mac, an `Abort trap: 6` / exit `134` popup can occur when a restricted app
  sandbox blocks LibreOffice's macOS startup. Try the command from Terminal
  outside that sandbox. The integration tests succeeded outside the sandbox.
- For dependency errors, activate the environment and rerun
  `python -m pip install -r requirements.txt`. For other errors, retry with
  `--verbose` and include the output in an issue report.

## H5P output and compatibility

All versions produce one PNG image per included slide inside an H5P Course
Presentation. Text and objects are not editable; animations, video, and speaker
notes are not transferred. LibreOffice may render fonts, diagrams, and layouts
differently from PowerPoint. Install the source fonts and inspect the result
before publishing.

The CLI excludes review comments from the slide images, including comment icons
in the LibreOffice PDF rendering route. Comments in the source PPTX are preserved.

The destination needs Course Presentation **1.25**, Image **1.1**, FontAwesome
**4.5**, H5P.FontIcons **1.0**, H5P.JoubelUI **1.3**, and H5P.Transition **1.0**,
including their dependencies. Archives contain content and metadata rather than
bundled runtime libraries. For missing-library import errors, ask the destination
administrator to install compatible libraries. See [H5P dependency validation](https://h5p.org/creating-your-own-h5p-plugin).
An archive structure check does not establish compatibility with every H5P host.

## Development and verification

See [build_script/README.md](build_script/README.md) for source entry points and
Windows GUI executable packaging. Use the root requirements file for the shared
development environment.

```bash
python3 -m unittest discover -s tests -v
H5P_RUN_INTEGRATION=1 python3 -m unittest discover -s tests -v
```

The integration tests require LibreOffice and the bundled
[Introduction to H5P fixture](tests/fixtures/introduction_to_h5p/README.md).
They cover missing-font rejection, silent override, and conversion with available
fonts as well as the existing packaging and folder-conversion checks.
Fixtures group the source deck and reference exports by presentation. Derived
decks, conversion outputs, and temporary PDFs are created in temporary directories.
PowerPoint reference exports are for visual comparison; different apps need not
produce identical pixels.

The CLI was tested on Mac with Python 3.14 and LibreOffice. Windows PowerPoint
has mocked tests; a real Windows end-to-end run remains required before claiming
runtime verification. The original GUI scripts were inspected but were not run
on Windows during this work.

## License and attribution

MIT. See [LICENSE](LICENSE) and [ATTRIBUTION.md](ATTRIBUTION.md).
