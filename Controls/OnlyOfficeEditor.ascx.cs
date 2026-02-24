using System;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebEditor.Controls
{
    /// Control reutilizable que renderiza un editor OnlyOffice Document Server.
    public partial class OnlyOfficeEditor : UserControl
    {
        // ═══════════════════════════════════════════════════════════════
        //  Propiedades del documento (persistidas en ViewState)
        // ═══════════════════════════════════════════════════════════════

        /// URL del documento accesible por Document Server.
        public string DocumentUrl
        {
            get => (string)ViewState["DocumentUrl"];
            set => ViewState["DocumentUrl"] = value;
        }

        /// Nombre original del documento (ej. "reporte.docx").
        public string DocumentName
        {
            get => (string)ViewState["DocumentName"];
            set => ViewState["DocumentName"] = value;
        }

        /// Clave única para el caché de Document Server. Cambiarla fuerza recarga.
        public string DocumentKey
        {
            get => (string)ViewState["DocumentKey"];
            set => ViewState["DocumentKey"] = value;
        }

        /// URL donde Document Server envía notificaciones de guardado.
        public string CallbackUrl
        {
            get => (string)ViewState["CallbackUrl"];
            set => ViewState["CallbackUrl"] = value;
        }

        // ═══════════════════════════════
        //  Configuración de conexión con Document Server
        // ═══════════════════════════════

        /// URL completa al JS del API de OnlyOffice Document Server.
        public string OnlyOfficeApiUrl { get; set; } = "https://192.168.10.14:4443/web-apps/apps/api/documents/api.js";

        /// Secreto JWT compartido con Document Server.
        public string JwtSecret { get; set; } = "secreto_personalizado";

        /// URL base pública de esta aplicación (para que document server devuelva el documento).
        public string PublicBaseUrl { get; set; } = "http://192.168.10.34:2355";

        // ═══════════════════════════════
        //  Propiedades de configuración
        // ═══════════════════════════════

        /// "edit" o "view". Por defecto: "edit".
        public string Mode { get; set; } = "edit";

        /// Código de idioma (ej. "es", "en"). Por defecto: "es".
        public string Lang { get; set; } = "es";

        /// Altura mínima CSS del contenedor del editor. Por defecto: "520px".
        public string EditorHeight { get; set; } = "520px";


        /// ID del usuario para la sesión del editor.
        public string UserId { get; set; } = "1";

        /// Nombre del usuario para la sesión del editor.
        public string UserDisplayName { get; set; } = "Usuario";

        // ═══════════════════════════════════════════════════════════════
        //  Propiedades para enlazar botones externos (sin JS en la página)
        // ═══════════════════════════════════════════════════════════════

        /// IDs de los controles (botón, LinkButton, etc.) que al hacer clic capturarán
        /// el documento editado, lo almacenarán en un HiddenField interno, y
        /// dispararán un postback. Tras el postback se pueden leer los bytes
        /// con <see cref="GetEditedDocumentBytes()"/>.
        /// Acepta un solo ID o varios separados por coma: "btn1,btn2,btn3".
        /// El control conecta automáticamente el JavaScript necesario.
        public string CaptureTriggerId { get; set; }

        // ═══════════════════════════════════════════════════════════════
        //  Propiedades computadas
        // ═══════════════════════════════════════════════════════════════

        /// ID del DOM del contenedor del editor (único por instancia).
        public string EditorContainerId => ClientID + "_editor";

        /// JSON de configuración de OnlyOffice (se renderiza en la página).
        public string ConfigJson { get; private set; } = "null";

        /// Indica si el control tiene un documento válido para mostrar.
        public bool HasDocument =>
            !string.IsNullOrWhiteSpace(DocumentUrl)
            && !string.IsNullOrWhiteSpace(DocumentName)
            && !string.IsNullOrWhiteSpace(DocumentKey);

        /// ID del HiddenField que almacena el documento editado en base64.
        /// Se usa desde JavaScript para inyectar el contenido antes del postback.
        public string HiddenFieldClientId => hfEditedDocumentBase64.ClientID;

        // ═══════════════════════════════════════════════════════════════
        //  Acceso al documento editado desde code-behind
        // ═══════════════════════════════════════════════════════════════

        /// Indica si el HiddenField contiene un documento editado (tras un postback
        /// disparado por <c>captureToHiddenField</c> en JavaScript).
        public bool HasEditedDocument =>
            !string.IsNullOrWhiteSpace(hfEditedDocumentBase64.Value);

        /// Obtiene los bytes del documento editado que fue capturado en el HiddenField
        /// por la función JS <c>OnlyOfficeEditorModule.captureToHiddenField()</c>.
        /// Retorna <c>null</c> si no hay documento capturado.
        public byte[] GetEditedDocumentBytes()
        {
            var b64 = hfEditedDocumentBase64.Value;
            if (string.IsNullOrWhiteSpace(b64)) return null;
            try { return Convert.FromBase64String(b64); }
            catch { return null; }
        }

        /// Limpia el contenido del HiddenField (para liberar memoria del ViewState
        /// después de haber consumido los bytes).
        public void ClearEditedDocument()
        {
            hfEditedDocumentBase64.Value = string.Empty;
        }

        // ═══════════════════════════════════════════════════════════════
        //  Métodos para cargar documentos
        // ═══════════════════════════════════════════════════════════════

        /// Carga un documento desde bytes en memoria.
        /// Guarda el archivo en App_Data/uploads y configura todas las URLs automáticamente.
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

        /// Carga un documento desde una ruta física en el servidor.
        /// "serverFilePath" => Ruta absoluta al archivo en el servidor.
        public void SetDocumentFromFile(string serverFilePath, string displayName = null)
        {
            if (!File.Exists(serverFilePath)) return;
            SetDocumentFromBytes(
                File.ReadAllBytes(serverFilePath),
                displayName ?? Path.GetFileName(serverFilePath));
        }

        /// Configura el editor para un archivo que ya existe en App_Data/uploads
        /// (por ejemplo, uno subido previamente con su fileId conocido).
        /// "fileId" => Identificador del archivo almacenado (sin extensión).
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

        /// Registra automáticamente los scripts de arranque del editor y
        /// los handlers de clic para DownloadTriggerId / CaptureTriggerId.
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

        /// Busca un control por ID recursión arriba y abajo en el árbol.
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
