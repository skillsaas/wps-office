# Automating WPS: JS macros (JSA), add-ins (wpsjs) and AirScript

Load this when the user wants code that runs in WPS. The agent writes and explains the code;
the user runs it in their own WPS. Nothing here was run against WPS (not installed on the box).

## Which surface
| Need | Surface | Where it runs | Docs |
| :--- | :--- | :--- | :--- |
| Automate one document or workbook, share as one file | JS macro (JSA) in the WPS macro editor, saved in `.xlsm`/`.docm`/`.pptm` | Desktop WPS (Windows, macOS, Linux) | open.wps.cn JS API (Application, Document, CreateObject) |
| A ribbon button or task pane for many users | WPS add-in (加载项) built with the `wpsjs` CLI | Desktop WPS | open.wps.cn add-in guide; npm `wpsjs` |
| Cloud document in KDocs (金山文档) | AirScript, optionally triggered over HTTP with a script token | Kingsoft cloud | airsheet.wps.cn AirScript docs |
| Existing VBA | VBA in WPS | Depends on edition; not verified | |

## JS macros (JSA)
- Runtime: V8, most ES6 syntax, JavaScript built-ins only; no browser objects such as
  `window`, `document` or `fetch` (open.wps.cn "JavaScript runtime").
- The object model mirrors Office's: `Application`, `ActiveDocument`, `Documents`, `Workbooks`,
  `Range`, and so on. In the macro editor the common members are global, so
  `ActiveDocument.PrintOut()` works without `Application.` (open.wps.cn Application object).
- `ActiveDocument` throws when no document is open: check `Documents.Count` first.
- Cross-app: `CreateObject("kwps.application")` Writer, `"ket.application"` Spreadsheets,
  `"kwpp.application"` Presentation (case-insensitive; open.wps.cn CreateObject).
- Optional arguments: pass an array with holes and spread it, for example
  `let a = []; a[8] = true; ThisWorkbook.SaveAs(...a)` sets only the 9th argument (aliyun article).
- Save the macro in a macro-enabled format; plain `.docx`/`.xlsx`/`.pptx` cannot hold it
  (wps.cn help). In a workbook the code lands in `xl/JDEData.bin`, as plain text (bbs.wps.cn 36292).
- External programs can call a JSA macro with `Application.Run`, text arguments only
  (bbs.wps.cn 36290).

Example from the open.wps.cn Document page (documented, not run here):
```javascript
function test() {
    Documents.Item("Document1.docx").Activate()
    ActiveDocument.PageSetup.Orientation = wdOrientLandscape
    ActiveDocument.PrintOut()
}
```
Right pattern: read with the object model, change the minimum, `SaveAs` to a new name. Wrong
pattern: a macro that `Save()`s over the user's file, or prints, sends or deletes without asking.

## WPS add-ins (wpsjs)
Documented flow (open.wps.cn "generate the first WPS add-in"):
```bash
npm install -g wpsjs          # npm view wpsjs: 2.2.3, "用于开发wps加载项的工具包" (checked 2026-10-13)
wpsjs create HelloWps         # pick Writer, Spreadsheets or Presentation; plain, Vue or React template
cd HelloWps && wpsjs debug    # starts WPS, loads the add-in, hot reload
wpsjs publish                 # builds wps-addon-build and publish.html; users install from that page
```
`jsplugins.xml` is only for repackaged (OEM) deployments pointed to by `oem.ini`
`JSPluginsServer` (open.wps.cn add-in guide). Installing an add-in for the user is a system
change: ask first.

## AirScript on KDocs (cloud)
- Scripts live in a cloud document. A script token (APIToken) created in the script editor lets
  an external system run a script over HTTPS (airsheet.wps.cn).
- Sync run: `POST https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task`,
  headers `Content-Type: application/json` and `AirScript-Token: <token>`, body
  `{"Context":{"argv":{}}}`; arguments arrive in the script as `Context.argv`. Optional
  `Context.sheet_name`, `Context.range`. Response has `data.result`, `data.logs`, `status`, `error`.
- Async run: `POST .../task` returns `task_id`; poll `GET /api/v3/script/task?task_id=<url-encoded id>`.
- A token runs code that can change the document: treat it like a password. Never ask the user
  to paste it into chat; have them set an environment variable themselves, and ask before every run.
