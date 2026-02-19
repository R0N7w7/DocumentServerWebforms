using System;
using System.IO;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;

namespace WebEditor
{
    public partial class _Default : Page
    {
        private const string UploadFolderVirtual = "~/App_Data/uploads";

        // Centraliza aquí la configuración de red/OnlyOffice.
        // - PublicBaseUrl: URL pública de esta app (la que el Document Server puede alcanzar)
        // - JwtSecret: secreto compartido con Document Server
        private static class OnlyOfficeSettings
        {
            // Ejemplo: "http://192.168.10.34:2355" (sin slash final)
            public static string PublicBaseUrlOverride { get; set; } = "http://192.168.10.34:2355";

            // Debe coincidir con el valor configurado en OnlyOffice Document Server
            public static string JwtSecret { get; set; } = "secreto_personalizado";

            public static string GetPublicBaseUrl(HttpRequest request)
            {
                if (!string.IsNullOrWhiteSpace(PublicBaseUrlOverride))
                    return PublicBaseUrlOverride.TrimEnd('/');

                if (request?.Url == null)
                    return null;

                var baseUri = request.Url;
                var appPath = request.ApplicationPath ?? "/";
                var path = appPath == "/" ? "/" : appPath.TrimEnd('/') + "/";
                var builder = new UriBuilder(baseUri.Scheme, baseUri.Host, baseUri.Port, path);
                return builder.Uri.ToString().TrimEnd('/');
            }
        }

        protected string OnlyOfficeConfigJson { get; private set; }

        protected void Page_Load(object sender, EventArgs e)
        {
            // OnlyOffice envía callbacks POST al callbackUrl; respondemos siempre OK.
            if (IsCallbackRequest())
            {
                RespondJson("{\"error\":0}");
                return;
            }

            if (!IsPostBack)
            {
                OnlyOfficeConfigJson = "null";
                return;
            }

            OnlyOfficeConfigJson = BuildOnlyOfficeConfigJson();
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (!fuFile.HasFile)
            {
                litStatus.Text = string.Empty;
                return;
            }

            var uploadsPhysical = Server.MapPath(UploadFolderVirtual);
            Directory.CreateDirectory(uploadsPhysical);

            var originalName = Path.GetFileName(fuFile.FileName);
            var fileId = Guid.NewGuid().ToString("N");
            var storedName = fileId + Path.GetExtension(originalName);
            var physicalPath = Path.Combine(uploadsPhysical, storedName);
            fuFile.SaveAs(physicalPath);

            hfFileId.Value = fileId;
            hfDocKey.Value = GenerateOnlyOfficeDocumentKey(fileId);

            OnlyOfficeConfigJson = BuildOnlyOfficeConfigJson(originalName, storedName);
            litStatus.Text = string.Empty;
        }

        private static string GenerateOnlyOfficeDocumentKey(string fileId)
        {
            if (string.IsNullOrWhiteSpace(fileId))
                return Guid.NewGuid().ToString("N");

            var clean = fileId.Replace("-", "");
            if (clean.Length < 24)
                clean = clean.PadRight(24, '0');

            var shard = clean.Substring(0, 20);
            var suffix = clean.Substring(clean.Length - 4);
            return shard + "_" + suffix;
        }

        private bool IsCallbackRequest()
        {
            return string.Equals(Request.QueryString["onlyoffice"], "callback", StringComparison.OrdinalIgnoreCase);
        }

        private string BuildOnlyOfficeConfigJson(string originalName = null, string storedName = null)
        {
            if (string.IsNullOrWhiteSpace(hfFileId.Value) || string.IsNullOrWhiteSpace(hfDocKey.Value))
                return "null";

            if (string.IsNullOrWhiteSpace(storedName))
                storedName = FindStoredName(hfFileId.Value);
            if (storedName == null)
                return "null";

            if (string.IsNullOrWhiteSpace(originalName))
                originalName = storedName;

            var fileUrl = AbsoluteUrl("~/Default.aspx?onlyoffice=download&fileId=" + HttpUtility.UrlEncode(hfFileId.Value));
            var callbackUrl = AbsoluteUrl("~/Default.aspx?onlyoffice=callback&fileId=" + HttpUtility.UrlEncode(hfFileId.Value));

            var ext = Path.GetExtension(originalName);
            var fileType = string.IsNullOrWhiteSpace(ext) ? "" : ext.TrimStart('.');
            var documentType = "word";

            var configObject = new
            {
                document = new
                {
                    fileType = fileType,
                    key = hfDocKey.Value,
                    title = originalName,
                    url = fileUrl
                },
                documentType = documentType,
                editorConfig = new
                {
                    callbackUrl = callbackUrl,
                    mode = "edit",
                    lang = "es",
                    user = new { id = "1", name = "Usuario" }
                }
            };

            var serializer = new JavaScriptSerializer();
            var configJson = serializer.Serialize(configObject);
            var token = OnlyOfficeJwtHelper.Create(configJson, OnlyOfficeSettings.JwtSecret);

            var finalJson = "{\"token\":" + serializer.Serialize(token) + ",\"document\":" + serializer.Serialize(configObject.document) + ",\"documentType\":" + serializer.Serialize(configObject.documentType) + ",\"editorConfig\":" + serializer.Serialize(configObject.editorConfig) + "}";
            return finalJson;
        }

        protected override void Render(HtmlTextWriter writer)
        {
            // Sirve el archivo original para que Document Server lo cargue en el editor.
            if (string.Equals(Request.QueryString["onlyoffice"], "download", StringComparison.OrdinalIgnoreCase))
            {
                var fileId = Request.QueryString["fileId"];
                var storedName = string.IsNullOrWhiteSpace(fileId) ? null : FindStoredName(fileId);
                if (storedName == null)
                {
                    Response.StatusCode = 404;
                    Response.End();
                    return;
                }
                var physicalPath = GetPhysicalPath(storedName);
                TransmitFile(physicalPath, storedName);
                return;
            }

            base.Render(writer);
        }

        private static class OnlyOfficeJwtHelper
        {
            public static string Create(string jsonPayload, string secret)
            {
                if (string.IsNullOrWhiteSpace(jsonPayload)) return null;
                if (secret == null) secret = string.Empty;

                var headerJson = "{\"alg\":\"HS256\",\"typ\":\"JWT\"}";
                var header = Base64UrlEncode(Encoding.UTF8.GetBytes(headerJson));
                var payload = Base64UrlEncode(Encoding.UTF8.GetBytes(jsonPayload));
                var signingInput = header + "." + payload;
                var signatureBytes = HmacSha256(Encoding.UTF8.GetBytes(secret), Encoding.UTF8.GetBytes(signingInput));
                var signature = Base64UrlEncode(signatureBytes);
                return signingInput + "." + signature;
            }

            private static string Base64UrlEncode(byte[] input)
            {
                return Convert.ToBase64String(input)
                    .TrimEnd('=')
                    .Replace('+', '-')
                    .Replace('/', '_');
            }

            private static byte[] HmacSha256(byte[] key, byte[] data)
            {
                using (var h = new System.Security.Cryptography.HMACSHA256(key))
                {
                    return h.ComputeHash(data);
                }
            }
        }

        private string FindStoredName(string fileId)
        {
            var uploadsPhysical = Server.MapPath(UploadFolderVirtual);
            if (!Directory.Exists(uploadsPhysical))
                return null;

            var matches = Directory.GetFiles(uploadsPhysical, fileId + ".*");
            if (matches.Length == 0)
                return null;

            return Path.GetFileName(matches[0]);
        }

        private string GetPhysicalPath(string storedName)
        {
            var uploadsPhysical = Server.MapPath(UploadFolderVirtual);
            return Path.Combine(uploadsPhysical, storedName);
        }

        private string AbsoluteUrl(string relative)
        {
            var url = ResolveUrl(relative);

            var cfgBase = OnlyOfficeSettings.GetPublicBaseUrl(Request);
            if (!string.IsNullOrWhiteSpace(cfgBase))
            {
                var baseUri = new Uri(cfgBase.TrimEnd('/') + "/");
                var relativeWithoutTilde = url.StartsWith("~", StringComparison.Ordinal) ? url.Substring(1) : url;
                return new Uri(baseUri, relativeWithoutTilde.TrimStart('/')).ToString();
            }

            return new Uri(Request.Url, url).ToString();
        }

        private void TransmitFile(string physicalPath, string downloadName)
        {
            if (!File.Exists(physicalPath))
            {
                Response.StatusCode = 404;
                Response.End();
                return;
            }

            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("Content-Disposition", "attachment; filename=\"" + downloadName.Replace("\"", "") + "\"");
            Response.TransmitFile(physicalPath);
            Response.End();
        }

        private void RespondJson(string json)
        {
            Response.Clear();
            Response.ContentType = "application/json";
            Response.Write(json);
            Response.End();
        }
    }
}