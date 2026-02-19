using System;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;

namespace WebEditor.Controls
{
    /// <summary>
    /// Control reutilizable que renderiza un editor OnlyOffice Document Server.
    /// 
    /// ── Uso en .aspx ──────────────────────────────────────────────────
    ///   &lt;%@ Register Src="~/Controls/OnlyOfficeEditor.ascx" TagPrefix="oo" TagName="Editor" %&gt;
    ///   &lt;oo:Editor ID="docEditor" runat="server" Mode="edit" EditorHeight="600px" /&gt;
    /// 
    /// ── Uso en code-behind ────────────────────────────────────────────
    ///   // Opción A: pasar bytes directamente (el control guarda el archivo y genera URLs)
    ///   docEditor.SetDocumentFromBytes(fileBytes, "reporte.docx");
    /// 
    ///   // Opción B: pasar ruta de archivo en el servidor
    ///   docEditor.SetDocumentFromFile(@"C:\docs\reporte.docx");
    /// 
    ///   // Opción C: configurar URLs manualmente
    ///   docEditor.DocumentUrl  = "http://…/download?file=1";
    ///   docEditor.DocumentName = "reporte.docx";
    ///   docEditor.DocumentKey  = "clave-unica";
    ///   docEditor.CallbackUrl  = "http://…/callback";
    /// 
    /// ── API JavaScript del lado del cliente ────────────────────────────
    ///   // Obtener URL del documento editado (Promise)
    ///   OnlyOfficeEditorModule.getEditedDocumentUrl('containerId')
    ///       .then(function(url) { /* url del archivo convertido */ });
    /// 
    ///   // Descargar directamente en el navegador
    ///   OnlyOfficeEditorModule.downloadDocument('containerId');
    /// 
    ///   // Obtener como Blob (requiere CORS habilitado en Document Server)
    ///   OnlyOfficeEditorModule.getEditedDocumentBlob('containerId')
    ///       .then(function(blob) { /* subir blob a servidor, etc. */ });
    /// 
    ///   // Obtener instancia nativa DocsAPI.DocEditor
    ///   var editor = OnlyOfficeEditorModule.getEditor('containerId');
    /// </summary>
    public partial class OnlyOfficeEditor : UserControl
    {
        // ═══════════════════════════════════════════════════════════════
        //  Propiedades del documento (persistidas en ViewState)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>URL del documento accesible por Document Server.</summary>
        public string DocumentUrl
        {
            get => (string)ViewState["DocumentUrl"];
            set => ViewState["DocumentUrl"] = value;
        }

        /// <summary>Nombre original del documento (ej. "reporte.docx").</summary>
        public string DocumentName
        {
            get => (string)ViewState["DocumentName"];
            set => ViewState["DocumentName"] = value;
        }

        /// <summary>Clave única para el caché de Document Server. Cambiarla fuerza recarga.</summary>
        public string DocumentKey
        {
            get => (string)ViewState["DocumentKey"];
            set => ViewState["DocumentKey"] = value;
        }

        /// <summary>URL donde Document Server envía notificaciones de guardado.</summary>
        public string CallbackUrl
        {
            get => (string)ViewState["CallbackUrl"];
            set => ViewState["CallbackUrl"] = value;
        }

        // ═══════════════════════════════════════════════════════════════
        //  Propiedades de configuración (se definen en el markup o en Page_Load)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>"edit" o "view". Por defecto: "edit".</summary>
        public string Mode { get; set; } = "edit";

        /// <summary>Código de idioma (ej. "es", "en"). Por defecto: "es".</summary>
        public string Lang { get; set; } = "es";

        /// <summary>Altura mínima CSS del contenedor del editor. Por defecto: "520px".</summary>
        public string EditorHeight { get; set; } = "520px";

        /// <summary>URL completa al JS del API de OnlyOffice Document Server.</summary>
        public string OnlyOfficeApiUrl { get; set; } = "https://192.168.10.34:4443/web-apps/apps/api/documents/api.js";

        /// <summary>Secreto JWT compartido con Document Server.</summary>
        public string JwtSecret { get; set; } = "secreto_personalizado";

        /// <summary>URL base pública de esta aplicación (como la ve Document Server).</summary>
        public string PublicBaseUrl { get; set; } = "http://192.168.10.34:2355";

        /// <summary>ID del usuario para la sesión del editor.</summary>
        public string UserId { get; set; } = "1";

        /// <summary>Nombre del usuario para la sesión del editor.</summary>
        public string UserDisplayName { get; set; } = "Usuario";

        // ═══════════════════════════════════════════════════════════════
        //  Propiedades computadas (solo lectura)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>ID del DOM del contenedor del editor (único por instancia).</summary>
        public string EditorContainerId => ClientID + "_editor";

        /// <summary>JSON de configuración de OnlyOffice (se renderiza en la página).</summary>
        public string ConfigJson { get; private set; } = "null";

        /// <summary>Indica si el control tiene un documento válido para mostrar.</summary>
        public bool HasDocument =>
            !string.IsNullOrWhiteSpace(DocumentUrl)
            && !string.IsNullOrWhiteSpace(DocumentName)
            && !string.IsNullOrWhiteSpace(DocumentKey);

        // ═══════════════════════════════════════════════════════════════
        //  Métodos de conveniencia para cargar documentos
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Carga un documento desde bytes en memoria.
        /// Guarda el archivo en App_Data/uploads y configura todas las URLs automáticamente.
        /// </summary>
        /// <param name="data">Contenido binario del archivo.</param>
        /// <param name="fileName">Nombre original del archivo (ej. "reporte.docx").</param>
        public void SetDocumentFromBytes(byte[] data, string fileName)
        {
            if (data == null || data.Length == 0 || string.IsNullOrWhiteSpace(fileName))
                return;

            var fileId = Guid.NewGuid().ToString("N");
            var ext = Path.GetExtension(fileName);
            var storedName = fileId + ext;
            var uploadsDir = HttpContext.Current.Server.MapPath("~/App_Data/uploads");
            Directory.CreateDirectory(uploadsDir);
            File.WriteAllBytes(Path.Combine(uploadsDir, storedName), data);

            DocumentName = Path.GetFileName(fileName);
            DocumentKey = GenerateDocumentKey(fileId);
            DocumentUrl = BuildAbsoluteUrl(
                "~/Handlers/OnlyOfficeHandler.ashx?action=download&fileId=" + HttpUtility.UrlEncode(fileId));
            CallbackUrl = BuildAbsoluteUrl(
                "~/Handlers/OnlyOfficeHandler.ashx?action=callback&fileId=" + HttpUtility.UrlEncode(fileId));
        }

        /// <summary>
        /// Carga un documento desde una ruta física en el servidor.
        /// </summary>
        /// <param name="serverFilePath">Ruta absoluta al archivo en el servidor.</param>
        /// <param name="displayName">Nombre a mostrar (opcional; si es null usa el nombre del archivo).</param>
        public void SetDocumentFromFile(string serverFilePath, string displayName = null)
        {
            if (!File.Exists(serverFilePath)) return;
            SetDocumentFromBytes(
                File.ReadAllBytes(serverFilePath),
                displayName ?? Path.GetFileName(serverFilePath));
        }

        /// <summary>
        /// Configura el editor para un archivo que ya existe en App_Data/uploads
        /// (por ejemplo, uno subido previamente con su fileId conocido).
        /// </summary>
        /// <param name="fileId">Identificador del archivo almacenado (sin extensión).</param>
        /// <param name="originalName">Nombre original del archivo.</param>
        public void SetDocumentFromUpload(string fileId, string originalName)
        {
            if (string.IsNullOrWhiteSpace(fileId)) return;

            DocumentName = originalName ?? fileId;
            DocumentKey = GenerateDocumentKey(fileId);
            DocumentUrl = BuildAbsoluteUrl(
                "~/Handlers/OnlyOfficeHandler.ashx?action=download&fileId=" + HttpUtility.UrlEncode(fileId));
            CallbackUrl = BuildAbsoluteUrl(
                "~/Handlers/OnlyOfficeHandler.ashx?action=callback&fileId=" + HttpUtility.UrlEncode(fileId));
        }

        // ═══════════════════════════════════════════════════════════════
        //  Ciclo de vida
        // ═══════════════════════════════════════════════════════════════

        protected void Page_PreRender(object sender, EventArgs e)
        {
            ConfigJson = HasDocument ? BuildConfigJson() : "null";
        }

        // ═══════════════════════════════════════════════════════════════
        //  Helpers privados
        // ═══════════════════════════════════════════════════════════════

        private string BuildConfigJson()
        {
            var ext = Path.GetExtension(DocumentName);
            var fileType = string.IsNullOrWhiteSpace(ext) ? "" : ext.TrimStart('.');

            var config = new
            {
                document = new
                {
                    fileType,
                    key = DocumentKey,
                    title = DocumentName,
                    url = DocumentUrl
                },
                documentType = ResolveDocumentType(fileType),
                editorConfig = new
                {
                    callbackUrl = CallbackUrl ?? "",
                    mode = Mode ?? "edit",
                    lang = Lang ?? "es",
                    user = new { id = UserId ?? "1", name = UserDisplayName ?? "Usuario" }
                }
            };

            var serializer = new JavaScriptSerializer();
            var json = serializer.Serialize(config);

            // Firmar con JWT
            var token = OnlyOfficeJwt.Create(json, JwtSecret);

            return "{\"token\":" + serializer.Serialize(token)
                + ",\"document\":" + serializer.Serialize(config.document)
                + ",\"documentType\":" + serializer.Serialize(config.documentType)
                + ",\"editorConfig\":" + serializer.Serialize(config.editorConfig)
                + "}";
        }

        private string BuildAbsoluteUrl(string virtualPath)
        {
            var resolved = ResolveUrl(virtualPath);
            if (!string.IsNullOrWhiteSpace(PublicBaseUrl))
            {
                var baseUri = new Uri(PublicBaseUrl.TrimEnd('/') + "/");
                var rel = resolved.StartsWith("~") ? resolved.Substring(1) : resolved;
                return new Uri(baseUri, rel.TrimStart('/')).ToString();
            }
            return new Uri(Page.Request.Url, resolved).ToString();
        }

        private static string GenerateDocumentKey(string fileId)
        {
            if (string.IsNullOrWhiteSpace(fileId))
                return Guid.NewGuid().ToString("N");
            var c = fileId.Replace("-", "");
            if (c.Length < 24) c = c.PadRight(24, '0');
            return c.Substring(0, 20) + "_" + c.Substring(c.Length - 4);
        }

        private static string ResolveDocumentType(string fileType)
        {
            switch ((fileType ?? "").ToLowerInvariant())
            {
                case "xls": case "xlsx": case "ods": case "csv": return "cell";
                case "ppt": case "pptx": case "odp": return "slide";
                default: return "word";
            }
        }
    }
}
