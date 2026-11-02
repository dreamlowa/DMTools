# Third-Party Licenses

This project uses third-party open-source libraries. The project license does not relicense those libraries; each dependency remains under its own license.

## Runtime Dependencies

| Dependency | Version | License | Notes |
| --- | --- | --- | --- |
| PyMuPDF / MuPDF | `1.26.7` | GNU AGPL v3.0 or commercial Artifex license | Used for opening and rendering PDF pages. Because the released EXE includes PyMuPDF/MuPDF, this project is distributed under GNU AGPL v3.0. |
| Pillow | `11.2.1` | MIT-CMU License | Used for image handling before barcode decoding. |
| zxing-cpp | `3.0.0` | Apache License 2.0 | Used for DataMatrix barcode decoding. |

## Build Dependencies

| Dependency | Version | License | Notes |
| --- | --- | --- | --- |
| PyInstaller | `6.21.0` | GPL v2.0 or later with bootloader exception | Used only to build the standalone Windows EXE. |

## References

- PyMuPDF / MuPDF license: https://mupdf.readthedocs.io/en/latest/license.html
- Artifex commercial licensing: https://artifex.com/licensing
- Pillow license: https://pillow.readthedocs.io/en/stable/about.html#license
- zxing-cpp package metadata: https://pypi.org/project/zxing-cpp/
- PyInstaller license: https://pyinstaller.org/en/stable/license.html
