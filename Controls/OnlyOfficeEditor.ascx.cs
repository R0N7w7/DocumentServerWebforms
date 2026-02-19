using System;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Web.UI.WebControls;

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
        //  Propiedades para enlazar botones externos (sin JS en la página)
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// IDs de los controles (botón, LinkButton, etc.) que al hacer clic capturarán
        /// el documento editado, lo almacenarán en un HiddenField interno, y
        /// dispararán un postback. Tras el postback se pueden leer los bytes
        /// con <see cref="GetEditedDocumentBytes()"/>.
        /// Acepta un solo ID o varios separados por coma: "btn1,btn2,btn3".
        /// El control conecta automáticamente el JavaScript necesario.
        /// </summary>
        public string CaptureTriggerId { get; set; }

        // ═══════════════════════════════════════════════════════════════
        //  Evento de documento capturado
        // ═══════════════════════════════════════════════════════════════



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

        /// <summary>
        /// ID del HiddenField que almacena el documento editado en base64.
        /// Se usa desde JavaScript para inyectar el contenido antes del postback.
        /// </summary>
        public string HiddenFieldClientId => hfEditedDocumentBase64.ClientID;

        // ═══════════════════════════════════════════════════════════════
        //  Acceso al documento editado desde code-behind
        // ═══════════════════════════════════════════════════════════════

        /// <summary>
        /// Indica si el HiddenField contiene un documento editado (tras un postback
        /// disparado por <c>captureToHiddenField</c> en JavaScript).
        /// </summary>
        public bool HasEditedDocument =>
            !string.IsNullOrWhiteSpace(hfEditedDocumentBase64.Value);

        /// <summary>
        /// Obtiene los bytes del documento editado que fue capturado en el HiddenField
        /// por la función JS <c>OnlyOfficeEditorModule.captureToHiddenField()</c>.
        /// Retorna <c>null</c> si no hay documento capturado.
        /// </summary>
        public byte[] GetEditedDocumentBytes()
        {
            var b64 = hfEditedDocumentBase64.Value;
            if (string.IsNullOrWhiteSpace(b64)) return null;
            try { return Convert.FromBase64String(b64); }
            catch { return null; }
        }

        /// <summary>
        /// Limpia el contenido del HiddenField (para liberar memoria del ViewState
        /// después de haber consumido los bytes).
        /// </summary>
        public void ClearEditedDocument()
        {
            hfEditedDocumentBase64.Value = string.Empty;
        }

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

        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void Page_PreRender(object sender, EventArgs e)
        {
            ConfigJson = HasDocument ? BuildConfigJson() : "null";
            RegisterTriggerScripts();
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

        /// <summary>
        /// Registra automáticamente los scripts de arranque del editor y
        /// los handlers de clic para DownloadTriggerId / CaptureTriggerId.
        /// </summary>
        private void RegisterTriggerScripts()
        {
            if (!HasDocument) return;

            var cs = Page.ClientScript;
            var type = GetType();
            var uid = ClientID;

            // 1─ Registrar los <script src> del API y del módulo (una sola vez por página)
            var apiKey = "oo_api_script";
            if (!cs.IsClientScriptIncludeRegistered(type, apiKey))
                cs.RegisterClientScriptInclude(type, apiKey, OnlyOfficeApiUrl);

            var moduleKey = "oo_module_script";
            if (!cs.IsClientScriptIncludeRegistered(type, moduleKey))
                cs.RegisterClientScriptInclude(type, moduleKey, ResolveUrl("~/Scripts/OnlyOfficeEditor.js"));

            // 2─ Registrar la URL del proxy (una sola vez por página)
            var proxyKey = "oo_proxy_url";
            if (!cs.IsStartupScriptRegistered(type, proxyKey))
            {
                var proxyUrl = ResolveUrl("~/Handlers/OnlyOfficeHandler.ashx?action=proxy&url=");
                var proxyScript = string.Format(
                    "window.__onlyOfficeProxyUrl='{0}';", proxyUrl);
                cs.RegisterStartupScript(type, proxyKey, proxyScript, true);
            }

            // 3─ Script de inicialización del editor (por instancia)
            var initKey = "oo_init_" + uid;
            if (!cs.IsStartupScriptRegistered(type, initKey))
            {
                var initScript = string.Format(
                    @"(function(){{ var cfg={0}; if(cfg) OnlyOfficeEditorModule.init('{1}',cfg); }})();",
                    ConfigJson, EditorContainerId);
                cs.RegisterStartupScript(type, initKey, initScript, true);
            }

            // 4─ Conectar botones de captura (guarda base64 → postback → OnClick)
            if (!string.IsNullOrWhiteSpace(CaptureTriggerId))
            {
                var ids = CaptureTriggerId.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var rawId in ids)
                {
                    var id = rawId.Trim();
                    if (string.IsNullOrEmpty(id)) continue;

                    var capBtn = FindControlRecursive(Page, id);
                    if (capBtn == null) continue;

                    var capJs = string.Format(
                        "if(typeof OnlyOfficeEditorModule!=='undefined'){{OnlyOfficeEditorModule.captureToHiddenField('{0}','{1}',{{autoPostBack:true,postBackTarget:'{2}'}}).catch(function(e){{console.error(e)}});}};return false;",
                        EditorContainerId,
                        hfEditedDocumentBase64.ClientID,
                        capBtn.UniqueID);

                    if (capBtn is IAttributeAccessor acc)
                        acc.SetAttribute("onclick", capJs);
                    else if (capBtn is WebControl wc)
                        wc.Attributes["onclick"] = capJs;
                }
            }
        }

        /// <summary>Busca un control por ID recursión arriba y abajo en el árbol.</summary>
        private static Control FindControlRecursive(Control root, string id)
        {
            if (root == null || string.IsNullOrWhiteSpace(id)) return null;
            if (root.ID == id) return root;
            foreach (Control c in root.Controls)
            {
                var found = FindControlRecursive(c, id);
                if (found != null) return found;
            }
            return null;
        }
    }

}
