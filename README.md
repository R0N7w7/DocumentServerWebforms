# OnlyOffice Editor — Control Reutilizable para ASP.NET WebForms

Control modular (`UserControl`) que encapsula la integración con **OnlyOffice Document Server**.  
Permite incrustar un editor de documentos en cualquier página `.aspx`, pasarle un documento de múltiples formas, y obtener el documento editado desde JavaScript para descargarlo, enviarlo o guardarlo en base de datos.

Targets **.NET Framework 4.7.2**. Los archivos temporales se almacenan en `App_Data/uploads`.

---

## Estructura del proyecto

```
Controls/
  OnlyOfficeEditor.ascx            ← UserControl (markup)
  OnlyOfficeEditor.ascx.cs         ← Code-behind con propiedades y métodos
  OnlyOfficeEditor.ascx.designer.cs

Handlers/
  OnlyOfficeHandler.ashx            ← HTTP Handler para servir archivos y callbacks
  OnlyOfficeHandler.ashx.cs

Scripts/
  OnlyOfficeEditor.js               ← Módulo JS con API pública del cliente

App_Code/
  OnlyOfficeJwt.cs                  ← Generación y validación de tokens JWT (HS256)

Content/
  WebEditor.css                     ← Estilos del editor y componentes visuales
```

---

## Inicio rápido

### 1. Registrar el control en la página

```aspx
<%@ Register Src="~/Controls/OnlyOfficeEditor.ascx" TagPrefix="oo" TagName="Editor" %>
```

### 2. Colocar el control en el markup

```aspx
<oo:Editor ID="docEditor" runat="server" />
```

### 3. Pasar un documento desde el code-behind

```csharp
// En el evento de un botón, Page_Load, o donde se necesite:
docEditor.SetDocumentFromBytes(fileBytes, "contrato.docx");
```

Eso es todo. El control se encarga de:
- Almacenar el archivo temporalmente en `App_Data/uploads/`.
- Generar las URLs de descarga y callback para Document Server.
- Firmar la configuración con JWT.
- Renderizar el editor con el documento cargado.

---

## Formas de cargar un documento (code-behind)

### Opción A — Desde bytes en memoria

Ideal para archivos que vienen de un `FileUpload`, de base de datos o de una API externa.

```csharp
byte[] fileBytes = fuFile.FileBytes; // desde un FileUpload
docEditor.SetDocumentFromBytes(fileBytes, "reporte.docx");
```

### Opción B — Desde una ruta de archivo en el servidor

```csharp
docEditor.SetDocumentFromFile(@"C:\documentos\reporte.docx");

// Opcionalmente se puede dar un nombre para mostrar diferente:
docEditor.SetDocumentFromFile(@"C:\documentos\reporte.docx", "Mi Reporte Final.docx");
```

### Opción C — Desde un archivo que ya está en `App_Data/uploads/`

Si ya se tiene el `fileId` de un archivo subido previamente:

```csharp
docEditor.SetDocumentFromUpload("a1b2c3d4e5f6...", "reporte.docx");
```

### Opción D — Configuración manual de URLs

Para escenarios donde el documento se sirve desde otra fuente (CDN, otro servidor, etc.):

```csharp
docEditor.DocumentUrl  = "https://mi-servidor.com/api/documentos/123/download";
docEditor.DocumentName = "contrato.docx";
docEditor.DocumentKey  = "clave-unica-por-version";
docEditor.CallbackUrl  = "https://mi-servidor.com/api/documentos/123/callback";
```

> **Nota:** `DocumentKey` debe ser única por cada versión del documento. Si el contenido cambia, la clave debe cambiar también para que Document Server no use una versión cacheada.

---

## Propiedades configurables

Se pueden establecer en el markup o en el code-behind antes del render.

| Propiedad | Tipo | Default | Descripción |
|---|---|---|---|
| `Mode` | `string` | `"edit"` | Modo del editor: `"edit"` o `"view"`. |
| `Lang` | `string` | `"es"` | Idioma de la interfaz del editor. |
| `EditorHeight` | `string` | `"520px"` | Altura mínima CSS del contenedor. |
| `OnlyOfficeApiUrl` | `string` | *(URL del Document Server)* | URL completa al `api.js` de OnlyOffice. |
| `JwtSecret` | `string` | `"secreto_personalizado"` | Secreto JWT compartido con Document Server. |
| `PublicBaseUrl` | `string` | `"http://192.168.10.34:2355"` | URL base pública de la app (como la ve Document Server). |
| `UserId` | `string` | `"1"` | ID del usuario para la sesión del editor. |
| `UserDisplayName` | `string` | `"Usuario"` | Nombre del usuario en el editor. |

### Ejemplo con propiedades en markup

```aspx
<oo:Editor ID="docEditor" runat="server"
    Mode="edit"
    Lang="es"
    EditorHeight="700px"
    UserId="42"
    UserDisplayName="Ana García" />
```

### Ejemplo en code-behind

```csharp
protected void Page_Load(object sender, EventArgs e)
{
    docEditor.Mode = "view";  // Solo lectura
    docEditor.Lang = "en";
    docEditor.EditorHeight = "800px";
}
```

---

## Propiedades de solo lectura

| Propiedad | Tipo | Descripción |
|---|---|---|
| `EditorContainerId` | `string` | ID del DOM del contenedor del editor (único por instancia). Se usa en las llamadas a la API JS. |
| `HasDocument` | `bool` | `true` si el control tiene un documento válido configurado. |
| `ConfigJson` | `string` | JSON de configuración firmado con JWT (se genera automáticamente en `PreRender`). |

---

## API JavaScript del cliente

Todas las funciones del módulo `OnlyOfficeEditorModule` reciben el `containerId` como primer parámetro. En el markup se obtiene así:

```javascript
var containerId = '<%= docEditor.EditorContainerId %>';
```

### `getEditedDocumentUrl(containerId)` → `Promise<string>`

Solicita a Document Server que genere la URL de descarga del documento editado.

```javascript
OnlyOfficeEditorModule.getEditedDocumentUrl(containerId)
    .then(function(url) {
        console.log('URL del documento editado:', url);
        // Enviar la URL al servidor, guardar en BD, etc.
    })
    .catch(function(err) {
        console.error('Error:', err);
    });
```

### `downloadDocument(containerId, fileName?)` → `Promise<string>`

Descarga el documento editado directamente en el navegador.

```javascript
OnlyOfficeEditorModule.downloadDocument(containerId)
    .then(function(url) {
        console.log('Descarga iniciada desde:', url);
    });
```

### `getEditedDocumentBlob(containerId)` → `Promise<Blob>`

Obtiene el documento como un `Blob` de JavaScript. Útil para subirlo a otro servidor o procesarlo en el cliente.

> **Requisito:** Document Server debe tener CORS habilitado para el origen de esta aplicación.

```javascript
OnlyOfficeEditorModule.getEditedDocumentBlob(containerId)
    .then(function(blob) {
        // Ejemplo: subir el blob a un API
        var fd = new FormData();
        fd.append('archivo', blob, 'documento.docx');
        return fetch('/api/documentos/guardar', {
            method: 'POST',
            body: fd
        });
    })
    .then(function(response) {
        console.log('Guardado exitosamente');
    });
```

### `getEditor(containerId)` → `DocsAPI.DocEditor | null`

Retorna la instancia nativa de `DocsAPI.DocEditor` para acceso directo al API de OnlyOffice.

```javascript
var editor = OnlyOfficeEditorModule.getEditor(containerId);
if (editor) {
    editor.downloadAs();
}
```

### `destroy(containerId)`

Destruye la instancia del editor y libera recursos del DOM.

```javascript
OnlyOfficeEditorModule.destroy(containerId);
```

### `setBusy(containerId, isBusy)`

Muestra u oculta el overlay de "Procesando…" sobre el editor.

```javascript
OnlyOfficeEditorModule.setBusy(containerId, true);  // Mostrar
OnlyOfficeEditorModule.setBusy(containerId, false); // Ocultar
```

### `init(containerId, config, options?)` → `DocsAPI.DocEditor | null`

Inicializa manualmente el editor (normalmente el control lo invoca automáticamente). Acepta callbacks opcionales:

```javascript
OnlyOfficeEditorModule.init('mi-editor', configJson, {
    onReady: function() { console.log('Editor listo'); },
    onDocumentReady: function() { console.log('Documento cargado'); },
    onError: function(err) { console.error('Error:', err); },
    applyTheme: true  // inyecta tema visual al iframe (default: true)
});
```

---

## Escenarios de uso completos

### Subir, editar y descargar

```aspx
<%@ Register Src="~/Controls/OnlyOfficeEditor.ascx" TagPrefix="oo" TagName="Editor" %>

<asp:FileUpload ID="fuFile" runat="server" />
<asp:Button ID="btnUpload" runat="server" Text="Abrir" OnClick="btnUpload_Click" />

<oo:Editor ID="docEditor" runat="server" />

<button type="button" onclick="descargar();">Descargar editado</button>

<script>
    function descargar() {
        var id = '<%= docEditor.EditorContainerId %>';
        OnlyOfficeEditorModule.downloadDocument(id);
    }
</script>
```

```csharp
protected void btnUpload_Click(object sender, EventArgs e)
{
    if (fuFile.HasFile)
        docEditor.SetDocumentFromBytes(fuFile.FileBytes, fuFile.FileName);
}
```

### Cargar desde base de datos y guardar de vuelta

```csharp
protected void Page_Load(object sender, EventArgs e)
{
    if (!IsPostBack)
    {
        byte[] contenido = MiRepositorio.ObtenerDocumento(idDocumento);
        docEditor.SetDocumentFromBytes(contenido, "expediente.docx");
    }
}
```

```javascript
function guardarEnBD() {
    var id = '<%= docEditor.EditorContainerId %>';
    OnlyOfficeEditorModule.getEditedDocumentBlob(id)
        .then(function(blob) {
            var fd = new FormData();
            fd.append('archivo', blob, 'expediente.docx');
            fd.append('idDocumento', '<%= idDocumento %>');
            return fetch('/api/documentos/actualizar', {
                method: 'POST',
                body: fd
            });
        })
        .then(function() { alert('Guardado en base de datos'); });
}
```

### Solo lectura (vista previa)

```aspx
<oo:Editor ID="vistaPrevia" runat="server" Mode="view" EditorHeight="400px" />
```

```csharp
vistaPrevia.SetDocumentFromFile(@"C:\plantillas\formato.docx");
```

### Múltiples editores en la misma página

```aspx
<oo:Editor ID="editorOriginal" runat="server" Mode="view" />
<oo:Editor ID="editorCopia" runat="server" Mode="edit" />
```

```csharp
editorOriginal.SetDocumentFromBytes(bytesOriginal, "original.docx");
editorCopia.SetDocumentFromBytes(bytesCopia, "copia.docx");
```

Cada instancia genera un `EditorContainerId` único, por lo que no hay conflictos.

---

## Handler HTTP (`OnlyOfficeHandler.ashx`)

El handler desacopla la lógica de servir archivos y recibir callbacks de cualquier página específica.

| Endpoint | Método | Descripción |
|---|---|---|
| `?action=download&fileId=xxx` | GET | Sirve el archivo de `App_Data/uploads/` para Document Server. |
| `?action=callback&fileId=xxx` | POST | Recibe notificaciones de guardado de Document Server. Responde `{"error":0}`. |

El control configura estas URLs automáticamente al usar `SetDocumentFromBytes`, `SetDocumentFromFile` o `SetDocumentFromUpload`.

---

## Configuración del entorno

| Qué configurar | Dónde | Descripción |
|---|---|---|
| URL del Document Server | Propiedad `OnlyOfficeApiUrl` del control | URL completa al `api.js` del Document Server. |
| URL base pública | Propiedad `PublicBaseUrl` del control | URL por la cual Document Server puede alcanzar esta app (esquema + host + puerto, sin slash final). |
| Secreto JWT | Propiedad `JwtSecret` del control | Debe coincidir con el configurado en Document Server. |

## Ejecución local

1. **Prerrequisitos:** Visual Studio 2022 (o 2019) con .NET Framework 4.7.2 y un OnlyOffice Document Server accesible.
2. **Restaurar paquetes:** Abrir la solución y dejar que NuGet restaure desde `packages.config`.
3. **Configurar:** Ajustar `OnlyOfficeApiUrl`, `PublicBaseUrl` y `JwtSecret` según el entorno.
4. **Ejecutar:** Iniciar con IIS Express. Subir un `.docx` o archivo compatible, editarlo y descargarlo.

---

## Notas importantes

- **ViewState:** Las propiedades del documento (`DocumentUrl`, `DocumentName`, `DocumentKey`, `CallbackUrl`) se persisten en ViewState, por lo que el editor sobrevive postbacks sin reconfiguración.
- **JWT:** La configuración se firma automáticamente con el secreto definido en `JwtSecret`. Asegúrate de que coincida con el configurado en Document Server.
- **Archivos temporales:** Los documentos se almacenan en `App_Data/uploads/`. No se eliminan automáticamente; implementa una tarea de limpieza según tu necesidad.
- **CORS para Blob:** Si usas `getEditedDocumentBlob()`, Document Server debe permitir peticiones cross-origin desde el dominio de tu aplicación.
- **Tipos de documento:** El control detecta automáticamente el tipo (`word`, `cell`, `slide`) según la extensión del archivo (`.docx`, `.xlsx`, `.pptx`, etc.).
- **Secretos hardcodeados:** Los valores por defecto de URLs y secretos son para desarrollo local. Muévelos a `Web.config` o `appSettings` antes de producción.