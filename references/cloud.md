# WPS cloud documents and KDocs

Load this when the file lives in WPS cloud (international) or KDocs (金山文档, kdocs.cn).

- WPS for Linux 12.1.2.29182 added international account sign-in, cloud document sync and
  real-time co-editing (wps.com update log). Desktop WPS on all platforms can sync to the cloud.
- KDocs is Kingsoft's online editor; AirScript runs scripts inside its documents
  (`references/automation.md`).
- WPS WebOffice is Kingsoft's embeddable online editor for third-party systems; its supported
  formats list includes `.et`, `.ett`, `.dps`, `.dpt` and the Microsoft formats
  (solution.wps.cn).

Agent rules for cloud files:
1. Ask which account and service. Never sign in for the user, never handle passwords, cookies or
   tokens in chat.
2. Prefer a downloaded copy: the user exports the file, then the local scripts inspect it.
3. Sharing links, permission changes and uploads are visible to others: show exactly what will
   change and wait for a yes.
4. Do not upload a received file to the cloud to "convert" or "preview" it without consent.
5. An online editor renders the file on Kingsoft's servers and may run AirScript; it is not a
   sandbox for the user's privacy.
