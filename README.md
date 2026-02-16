# OnlyOffice Document Server WebForms Demo

ASP.NET Web Forms sample that integrates OnlyOffice Document Server to upload a document, edit it in the browser, and fetch the saved version through the server. It targets .NET Framework 4.7.2 and stores uploaded files under `App_Data/uploads`.

## Overview
- Upload a document, generate a stable OnlyOffice `document.key`, and load the editor with JWT-protected config.
- Receive OnlyOffice callbacks, validate the JWT, and cache the download URL in memory.
- Let users trigger a download of the saved version; the server proxies the file from Document Server.
- Minimal UI built with Web Forms, Bootstrap, and the OnlyOffice JS API.

## Architecture
- Server: [Default.aspx.cs](Default.aspx.cs) hosts all request handling (upload, OnlyOffice config, callback processing, proxy download, save polling) plus a lightweight in-memory callback store.
- Client: [Default.aspx](Default.aspx) renders the toolbar, editor surface, and JavaScript that instantiates `DocsAPI.DocEditor`, wires events, and drives the save/download flow.
- JWT helpers: [App_Code/OnlyOfficeJwt.cs](App_Code/OnlyOfficeJwt.cs) and the nested `OnlyOfficeJwtHelper` in the default page provide HS256 signing/validation for OnlyOffice payloads.
- Routing/bundles: [Global.asax.cs](Global.asax.cs), [App_Start/RouteConfig.cs](App_Start/RouteConfig.cs), and [App_Start/BundleConfig.cs](App_Start/BundleConfig.cs) enable Friendly URLs and script bundles (including a jQuery mapping).
- Storage: Uploaded files are saved to `App_Data/uploads` with a GUID-based name; callback download URLs are kept in-memory only.

## Document Flow
1) Upload: User selects a file and clicks **Subir y abrir**; the server saves it under a GUID and computes a deterministic `document.key`.
2) Editor config: The server builds a JSON config (document URL, callback URL, user info, language, JWT token) and the page loads `DocsAPI.DocEditor` using that config.
3) Callback: When the user saves in OnlyOffice (status 2/6/7), Document Server posts back to `Default.aspx?onlyoffice=callback&fileId=...`; the app validates the JWT and caches the `downloadUrl` by `document.key`.
4) Download: The **Guardar y descargar** flow calls `Default.aspx?onlyoffice=proxydownload&key=...` (or `fileId` fallback), which proxies the saved file from Document Server to the browser. A polling endpoint (`onlyoffice=savestatus`) supports the legacy flow if the editor does not return a URL directly.

## Configuration
- Document Server script: Update the `<script src="http://192.168.10.34:8085/web-apps/apps/api/documents/api.js"></script>` in [Default.aspx](Default.aspx) to point to your OnlyOffice Document Server.
- Public base URL: Set `OnlyOfficeSettings.PublicBaseUrlOverride` in [Default.aspx.cs](Default.aspx.cs) so Document Server can reach this app (include scheme and port, no trailing slash). If left blank, the app derives it from the incoming request URL.
- JWT secret: Set `OnlyOfficeSettings.JwtSecret` in [Default.aspx.cs](Default.aspx.cs) to the same secret configured on Document Server.
- Optional logging: Flip `EnableOnlyOfficeCallbackLogging` in [Default.aspx.cs](Default.aspx.cs) to `true` to persist callback payloads under `App_Data/onlyoffice-callback.log` (disabled by default).

## Running Locally
1) Prerequisites: Visual Studio 2022 (or 2019) with .NET Framework 4.7.2, OnlyOffice Document Server reachable from your machine.
2) Restore packages: Open the solution and let NuGet restore from `packages.config` (packages are checked in under `packages/`).
3) Configure: Adjust the Document Server script URL, `PublicBaseUrlOverride`, and `JwtSecret` as described above.
4) Run: Start the Web Forms app (IIS Express is fine). Upload a `.docx` or compatible file, edit it in the embedded OnlyOffice editor, then download the saved version.

## Endpoints (all on Default.aspx)
- `?onlyoffice=download&fileId={id}`: Serve the originally uploaded file.
- `?onlyoffice=callback&fileId={id}`: OnlyOffice callback endpoint; validates JWT, records `downloadUrl` for statuses 2/6/7, returns `{ "error":0 }`.
- `?onlyoffice=proxydownload&key={docKey}` (or `fileId` fallback): Proxy the saved file from OnlyOffice using the cached callback URL.
- `?onlyoffice=savestatus&key={docKey}` (or `fileId` fallback): Poll whether a saved URL has been received.

## Notes and Limitations
- Callback data is in-memory only; any app recycle clears the cache and pending downloads.
- Secrets and URLs are currently hard-coded for convenience—move them to config before production use.
- The sample uses HTTP in the OnlyOffice script tag; prefer HTTPS in real deployments.