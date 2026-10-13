# WPS Office on Linux

Load this for installing, launching or debugging WPS on Linux. WPS was not installed on the box;
everything here is from the sources named.

## Install
- Download the `.deb` or `.rpm` from the official WPS Linux page and install the local file:
  `sudo apt install ./wps-office_VERSION_amd64.deb` (the `./` matters; plain
  `apt install wps-office` fails unless a configured repository has it) (wps.com install guides).
- Arch: `wps-office` from the AUR, fonts from `ttf-wps-fonts` (ArchWiki).
- Newest Linux build listed on wps.com/whatsnew/linux when read on 2026-10-13: 12.1.2.29182
  (international accounts and cloud, AI features, 20 interface languages). The Chinese site
  `linux.wps.cn` has its own builds. Ask before installing.

## Layout and launchers
| Item | Value | Source |
| :--- | :--- | :--- |
| Program files | `/opt/kingsoft/wps-office/office6` | nixpkgs package, Gentoo ebuild |
| Writer | `wps` | ArchWiki |
| Spreadsheets | `et` | ArchWiki |
| Presentation | `wpp` | ArchWiki |
| PDF | `wpspdf` | ArchWiki |
| MIME types | `application/wps-office.wps`, `.wpt`, `.et`, `.ett`, `.dps`, `.dpt` | ArchWiki |

Check what this build really has (names can vary by package build, wps.com):
```bash
command -v wps et wpp wpspdf
dpkg -L wps-office | grep /usr/bin/      # Debian/Ubuntu
```
The launchers start the GUI; they are not batch converters. Do not start WPS on a received file
to "look at it": that renders the file and can prompt for macros.

## Known problems
| Symptom | Cause | Fix | Source |
| :--- | :--- | :--- | :--- |
| PDF export or opening PDFs does nothing | `wpspdf` needs `libtiff.so.5` | Install libtiff5 (AUR `libtiff5`) | ArchWiki |
| Boxes or wrong glyphs | Missing fonts | Install the fonts the document uses (Arch: `ttf-wps-fonts`) | ArchWiki |
| `command not found: et` | Build without that launcher, or half-installed package | `dpkg -l wps-office`, reinstall the exact `.deb` | wps.com |
