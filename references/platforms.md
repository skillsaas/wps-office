# Platforms and editions: who can run what

Load this when the question is "can my WPS run this macro?", or the user is on a phone, a Mac,
or the international (wps.com) build. Nothing here was tested on a device; each row names its
source and how strong it is.

Knowledge freshness: every row was read on 2026-10-13 from the pages listed under Sources.
Editions and macro support change with WPS releases and differ between wps.cn and wps.com, so
recheck a row before telling a user something they will pay for or depend on.

Evidence labels used below: **official** (wps.cn or wps.com article or staff answer),
**community** (forum user), **unverified** (no source found). Nothing in this file is
**tested here**: WPS was not installed.

## Chinese (wps.cn) vs international (wps.com) builds
| Topic | Chinese build (wps.cn, 金山办公) | International build (wps.com) | Evidence |
| :--- | :--- | :--- | :--- |
| Account | WPS account (Kingsoft), phone or WeChat sign-in common | WPS account at wps.com; Linux 12.1.2.29182 added international sign-in | official (wps.com Linux update log); Chinese sign-in methods unverified |
| Cloud | KDocs (金山文档, kdocs.cn) and WPS cloud | WPS Cloud; KDocs availability unverified | official (wps.com update log); rest unverified |
| AirScript | On KDocs documents; script-token API at airsheet.wps.cn / kdocs.cn | Not found in wps.com docs | official (airsheet.wps.cn); international unverified |
| VBA | Paid editions plus the VBA component (table below) | Unverified | official (wps.cn) / unverified |
| JS macros (JSA) | Yes (macro editor) | Unverified in docs; same client family | official (open.wps.cn) / unverified |
| Fonts | Chinese fonts expected by Chinese documents | Same documents need the same fonts; on Linux install them yourself (Arch `ttf-wps-fonts`) | community (ArchWiki) |
| Linux launchers | `wps`, `et`, `wpp`, `wpspdf` (linux.wps.cn builds) | Same names in the wps.com `.deb`/`.rpm` | community (ArchWiki); not checked per build |
| Downloads | wps.cn, linux.wps.cn | wps.com | official |

## VBA by edition (WPS for Windows, Chinese builds)
| Edition | VBA | Source |
| :--- | :--- | :--- |
| Personal / free (个人版) | No VBA environment; the Macro button is grey | wps.cn learning room answers d/263839, d/263826, d/302526 (staff-style answers in the user Q&A area, WPS 2019, Windows 10) |
| Professional Enhanced (专业版增强版), WPS+ (WPS+云办公) | Yes | wps.cn learning room d/300529, d/263839 |
| Business Standard / Advanced (商业标准版 / 商业高级版) | Yes, after downloading and installing the VBA component (`vba_for_wps_2052.exe`, VBA 7.1) and restarting WPS | wps.cn learning article id 12181 (official article) |
| International build (wps.com) | Not verified here | |

So "WPS ran my macro" depends on the edition. A received file with VBA is still a risk on every
edition: the user may have, or later install, the VBA component.

## JS macros (JSA) by platform
| Platform | JSA | Source and strength |
| :--- | :--- | :--- |
| Windows | Yes, in every edition (the WPS macro editor) | open.wps.cn JS API; aliyun article (default since WPS 2021) |
| macOS, Linux | Yes | forum.wps.cn topic 93172 (forum, 2023): JSA macro files run on Windows, Mac and Linux |
| Android | No; described as planned | forum.wps.cn topic 8474 (forum, 2023) |
| iOS | No macros | bbs.wps.cn topic 48981 (official reply: iOS does not support macros) |

## VBA on other platforms
| Platform | VBA | Source and strength |
| :--- | :--- | :--- |
| macOS | Reported absent | bbs.wps.cn topic 61880 (community answer) |
| Linux | Not documented in any source found | |
| Android, iOS | No | bbs.wps.cn topics 48981 (official, iOS) and 61880 (community, Android) |

## What this means for an agent
- On a phone, macros in a received file do not run in WPS, but the file can still carry links,
  remote templates and embedded objects: run `inspect_wps.py` the same way.
- On macOS and Windows, `convert_wps.py` refuses (no bubblewrap). Use a Linux VM or container:
  WSL2 on Windows, UTM or a Docker Linux VM on macOS, then install `libreoffice-core` and
  `bubblewrap` inside it. Docker needs `--security-opt seccomp=unconfined` or a user-namespace
  enabled runtime for bwrap to start; if bwrap fails there, a disposable VM is the fallback.
- WPS mobile apps cannot be scripted by this skill.

## Sources
- https://www.wps.cn/learning/room/d/263839
- https://www.wps.cn/learning/room/d/263826
- https://www.wps.cn/learning/room/d/302526
- https://www.wps.cn/learning/room/d/300529
- https://www.wps.cn/learning/article/detail_simple/id/12181
- https://forum.wps.cn/topic/93172 and https://forum.wps.cn/topic/8474
- https://bbs.wps.cn/topic/48981 and https://bbs.wps.cn/topic/61880
