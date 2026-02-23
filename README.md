# OnlyOffice Editor (WebForms) — comportamiento real del control

Este proyecto expone un `UserControl` reutilizable para incrustar **OnlyOffice Document Server** en ASP.NET WebForms (.NET Framework 4.7.2).

La integración actual funciona así:

1. El servidor guarda/expone el archivo fuente en `App_Data/uploads`.
2. El control genera `document.url`, `editorConfig.callbackUrl` y `token` JWT.
3. En `PreRender`, el control registra scripts e inicializa `DocsAPI.DocEditor`.
4. Para obtener el editado, el cliente llama `downloadAs()` y recibe una URL temporal (`onDownloadAs`).
5. Si se requiere postback WebForms, esa descarga se convierte a base64 y se guarda en un `HiddenField`.

> Importante: el endpoint `callback` **no persiste** automáticamente la versión editada; solo responde `{"error":0}`.

---

## Estructura relevante

```
Controls/
  OnlyOfficeEditor.ascx
  OnlyOfficeEditor.ascx.cs

Scripts/
  OnlyOfficeEditor.js

Handlers/
  OnlyOfficeHandler.ashx
  OnlyOfficeHandler.ashx.cs

App_Code/
  OnlyOfficeJwt.cs
```

---

## Flujo del control

### 1) Cargar documento

Puedes cargar desde bytes, archivo físico o `fileId` existente:

```csharp
docEditor.SetDocumentFromBytes(fileBytes, "contrato.docx");
// o
docEditor.SetDocumentFromFile(@"C:\docs\contrato.docx");
// o
docEditor.SetDocumentFromUpload(fileId, "contrato.docx");
```

`SetDocumentFromBytes` guarda el archivo en `App_Data/uploads/<fileId>.<ext>` y configura:

- `DocumentUrl = .../Handlers/OnlyOfficeHandler.ashx?action=download&fileId=...`
- `CallbackUrl = .../Handlers/OnlyOfficeHandler.ashx?action=callback&fileId=...`
- `DocumentKey` derivada de `fileId`

### 2) Render e inicialización

En `Page_PreRender`:

- Se construye `ConfigJson` firmado con JWT (`OnlyOfficeJwt.Create(...)`).
- Se registra `OnlyOfficeApiUrl` (script `api.js` de Document Server).
- Se registra `~/Scripts/OnlyOfficeEditor.js`.
- Se registra `window.__onlyOfficeProxyUrl` apuntando a `action=proxy`.
- Se ejecuta `OnlyOfficeEditorModule.init(containerId, config)`.

### 3) Obtener el documento editado

El módulo JS no lee el archivo desde callback. El flujo real es:

- `getEditedDocumentUrl(containerId)` llama `editor.downloadAs()`.
- OnlyOffice dispara `onDownloadAs` con la URL temporal del archivo editado.
- `getEditedDocumentBlob(containerId)` descarga ese URL usando `action=proxy` (server-to-server) para evitar CORS del browser.

### 4) Captura para postback WebForms

Si defines `CaptureTriggerId`, el control intercepta el click del botón indicado:

1. Ejecuta `OnlyOfficeEditorModule.captureToHiddenField(...)`.
2. Convierte el blob a base64.
3. Lo guarda en `hfEditedDocumentBase64`.
4. Dispara `__doPostBack` al control origen.

En el evento del botón del servidor puedes hacer:

```csharp
var bytes = docEditor.GetEditedDocumentBytes();
if (bytes != null)
{
    // guardar o descargar
    docEditor.ClearEditedDocument();
}
```

---

## Propiedades del control

### Configurables

| Propiedad | Default | Descripción |
|---|---|---|
| `Mode` | `"edit"` | `"edit"` o `"view"`. |
| `Lang` | `"es"` | Idioma del editor. |
| `EditorHeight` | `"520px"` | Altura mínima del contenedor. |
| `OnlyOfficeApiUrl` | URL local hardcodeada | Script `api.js` de Document Server. |
| `JwtSecret` | `"secreto_personalizado"` | Secreto compartido para JWT HS256. |
| `PublicBaseUrl` | URL local hardcodeada | Base pública por la que Document Server alcanza esta app. |
| `UserId` | `"1"` | Usuario de sesión en editor. |
| `UserDisplayName` | `"Usuario"` | Nombre mostrado en editor. |
| `CaptureTriggerId` | `null` | ID (o IDs separados por coma) de botones que capturan el documento antes del postback. |

### Solo lectura / estado

| Propiedad / método | Descripción |
|---|---|
| `EditorContainerId` | ID HTML único de la instancia de editor. |
| `ConfigJson` | Config firmado con JWT, generado en `PreRender`. |
| `HasDocument` | Indica si hay `DocumentUrl`, `DocumentName` y `DocumentKey`. |
| `HiddenFieldClientId` | ClientID del hidden donde se guarda base64. |
| `HasEditedDocument` | Indica si el hidden ya contiene contenido. |
| `GetEditedDocumentBytes()` | Devuelve bytes capturados en postback. |
| `ClearEditedDocument()` | Limpia el hidden field. |

---

## API JavaScript disponible (`Scripts/OnlyOfficeEditor.js`)

Todas reciben `containerId` como primer parámetro.

- `init(containerId, config, options?)`
- `getEditedDocumentUrl(containerId) : Promise<string>`
- `getEditedDocumentBlob(containerId) : Promise<Blob>`
- `downloadDocument(containerId, fileName?) : Promise<string>`
- `captureToHiddenField(containerId, hiddenFieldId, options?) : Promise<string>`
- `getEditor(containerId)`
- `destroy(containerId)`
- `setBusy(containerId, isBusy)`

`captureToHiddenField` es la base del flujo WebForms de “capturar y luego procesar en OnClick”.

---

## Endpoints del handler (`OnlyOfficeHandler.ashx`)

| Endpoint | Método | Qué hace actualmente |
|---|---|---|
| `?action=download&fileId=...` | GET | Sirve `App_Data/uploads/<fileId>.*`. |
| `?action=callback&fileId=...` | POST | Confirma recepción (`{"error":0}`), sin guardar versión editada. |
| `?action=proxy&url=...` | GET | Descarga server-to-server y reenvía al cliente para evitar CORS. |

---

## Ejemplo real del proyecto (`Default.aspx`)

```aspx
<oo:Editor ID="docEditor" runat="server" CaptureTriggerId="btnDescargar" />
```

- `btnUpload_Click`: llama `SetDocumentFromBytes(...)`.
- `btnDescargar_Click`: lee `GetEditedDocumentBytes()` y responde descarga HTTP.

Este patrón ya está implementado en la página principal del repositorio.

---

## Configuración mínima para que funcione

1. `OnlyOfficeApiUrl` debe apuntar al `api.js` real de tu Document Server.
2. `PublicBaseUrl` debe ser alcanzable desde Document Server.
3. `JwtSecret` debe coincidir en app y Document Server.
4. Asegura permisos de escritura en `App_Data/uploads`.

---

## Limitaciones actuales

- No hay limpieza automática de `App_Data/uploads`.
- `callback` no descarga/guarda la versión final (si lo necesitas, debes implementarlo ahí).
- El control inyecta defaults de personalización UI en JS (`compactToolbar`, `hideRightMenu`, etc.).
- Los defaults de URL/secreto están hardcodeados para entorno local; mover a `Web.config` antes de producción.
