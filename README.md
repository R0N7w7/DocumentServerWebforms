# OnlyOffice Document Server WebForms Demo

ASP.NET Web Forms sample that integrates OnlyOffice Document Server to upload a document, edit it in the browser, and download it directly from the editor client. It targets .NET Framework 4.7.2 and stores uploaded files under `App_Data/uploads`.

## Overview
- Upload a document, generate a stable OnlyOffice `document.key`, and load the editor with JWT-protected config.
- Download the edited document directly from the client via `editor.downloadAs()`.
- Minimal UI built with Web Forms, Bootstrap, and the OnlyOffice JS API.

## Architecture
- Server: [Default.aspx.cs](Default.aspx.cs) handles file upload, serves files to Document Server, generates the JWT-signed editor config, and responds to OnlyOffice callbacks with `{"error":0}`.
- Client: [Default.aspx](Default.aspx) renders the toolbar, editor surface, and JavaScript that instantiates `DocsAPI.DocEditor`, wires events, and triggers `downloadAs()` for the save/download flow.
- JWT helpers: [App_Code/OnlyOfficeJwt.cs](App_Code/OnlyOfficeJwt.cs) and the nested `OnlyOfficeJwtHelper` in the default page provide HS256 signing for OnlyOffice payloads.
- Routing/bundles: [Global.asax.cs](Global.asax.cs), [App_Start/RouteConfig.cs](App_Start/RouteConfig.cs), and [App_Start/BundleConfig.cs](App_Start/BundleConfig.cs) enable Friendly URLs and script bundles.
- Storage: Uploaded files are saved to `App_Data/uploads` with a GUID-based name.

## Document Flow
1) Upload: User selects a file and clicks **Subir y abrir**; the server saves it under a GUID and computes a deterministic `document.key`.
2) Editor config: The server builds a JSON config (document URL, callback URL, user info, language, JWT token) and the page loads `DocsAPI.DocEditor` using that config.
3) Download: The **Guardar y descargar** button calls `editor.downloadAs()` on the client, which triggers the `onDownloadAs` event with a direct download URL from Document Server.

## Configuration
- Document Server script: Update the `<script src="...">` in [Default.aspx](Default.aspx) to point to your OnlyOffice Document Server.
- Public base URL: Set `OnlyOfficeSettings.PublicBaseUrlOverride` in [Default.aspx.cs](Default.aspx.cs) so Document Server can reach this app (include scheme and port, no trailing slash). If left blank, the app derives it from the incoming request URL.
- JWT secret: Set `OnlyOfficeSettings.JwtSecret` in [Default.aspx.cs](Default.aspx.cs) to the same secret configured on Document Server.

## Running Locally
1) Prerequisites: Visual Studio 2022 (or 2019) with .NET Framework 4.7.2, OnlyOffice Document Server reachable from your machine.
2) Restore packages: Open the solution and let NuGet restore from `packages.config` (packages are checked in under `packages/`).
3) Configure: Adjust the Document Server script URL, `PublicBaseUrlOverride`, and `JwtSecret` as described above.
4) Run: Start the Web Forms app (IIS Express is fine). Upload a `.docx` or compatible file, edit it in the embedded OnlyOffice editor, then download the saved version.

## Endpoints (all on Default.aspx)
- `?onlyoffice=download&fileId={id}`: Serve the originally uploaded file (used by Document Server to load the document).
- `?onlyoffice=callback&fileId={id}`: OnlyOffice callback endpoint; returns `{ "error":0 }` (required by Document Server).

## Notes and Limitations
- Secrets and URLs are currently hard-coded for convenience — move them to config before production use.
- The sample uses HTTPS in the OnlyOffice script tag; adjust the URL for your environment.