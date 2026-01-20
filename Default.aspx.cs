using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Configuration;

namespace WebEditor
{
    public partial class _Default : Page
    {
        private const string UploadFolderVirtual = "~/App_Data/uploads";
        private const string OnlyOfficeJwtSecret = "Y1EOwRcQmDQlgzBBTP3aQLAwWvVFlLz2";

        protected string OnlyOfficeConfigJson { get; private set; }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsCallbackRequest())
            {
                ProcessOnlyOfficeCallback();
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
                litStatus.Text = "<div class='text-danger mt-2'>Selecciona un archivo.</div>";
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
            hfDocKey.Value = ComputeKey(fileId, physicalPath);
            btnDownload.Enabled = true;

            OnlyOfficeConfigJson = BuildOnlyOfficeConfigJson(originalName, storedName);
            litStatus.Text = "<div class='text-success mt-2'>Archivo subido. Cargando editor…</div>";
        }

        protected void btnDownload_Click(object sender, EventArgs e)
        {
            var fileId = hfFileId.Value;
            if (string.IsNullOrWhiteSpace(fileId))
            {
                litStatus.Text = "<div class='text-danger mt-2'>No hay archivo para descargar.</div>";
                return;
            }

            var storedName = FindStoredName(fileId);
            if (storedName == null)
            {
                litStatus.Text = "<div class='text-danger mt-2'>Archivo no encontrado.</div>";
                return;
            }

            var physicalPath = GetPhysicalPath(storedName);
            // Entrega el archivo actual en disco (que OnlyOffice sobreescribe vía callback)
            TransmitFile(physicalPath, storedName);
        }

        private bool IsCallbackRequest()
        {
            return string.Equals(Request.QueryString["onlyoffice"], "callback", StringComparison.OrdinalIgnoreCase);
        }

        private void ProcessOnlyOfficeCallback()
        {
            // OnlyOffice server POSTs JSON here.
            string body;
            using (var reader = new StreamReader(Request.InputStream))
            {
                body = reader.ReadToEnd();
            }

            var serializer = new JavaScriptSerializer();
            dynamic payload;
            try
            {
                payload = serializer.DeserializeObject(body);
            }
            catch
            {
                RespondJson("{\"error\":1}");
                return;
            }

            var callbackToken = TryGetString(payload, "token");
            if (!string.IsNullOrWhiteSpace(callbackToken) && !ValidateJwt(callbackToken))
            {
                RespondJson("{\"error\":1}");
                return;
            }

            // Expected fields: status, url, key
            var status = TryGetInt(payload, "status");
            var downloadUrl = TryGetString(payload, "url");

            var fileId = Request.QueryString["fileId"];
            if (string.IsNullOrWhiteSpace(fileId))
            {
                RespondJson("{\"error\":1}");
                return;
            }

            // status 2 or 6 => ready to save.
            if ((status == 2 || status == 6) && !string.IsNullOrWhiteSpace(downloadUrl))
            {
                try
                {
                    var storedName = FindStoredName(fileId);
                    if (storedName == null)
                    {
                        RespondJson("{\"error\":1}");
                        return;
                    }

                    var physicalPath = GetPhysicalPath(storedName);
                    using (var wc = new System.Net.WebClient())
                    {
                        wc.DownloadFile(downloadUrl, physicalPath);
                    }
                }
                catch
                {
                    RespondJson("{\"error\":1}");
                    return;
                }
            }

            RespondJson("{\"error\":0}");
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

            var fileUrl = AbsoluteUrl("~/OnlyOffice.ashx?type=download&fileId=" + HttpUtility.UrlEncode(hfFileId.Value));
            var callbackUrl = AbsoluteUrl("~/OnlyOffice.ashx?type=callback&fileId=" + HttpUtility.UrlEncode(hfFileId.Value));

            var ext = Path.GetExtension(originalName);
            var fileType = string.IsNullOrWhiteSpace(ext) ? "" : ext.TrimStart('.');

            var configObject = new
            {
                document = new
                {
                    fileType = fileType,
                    key = hfDocKey.Value,
                    title = originalName,
                    url = fileUrl
                },
                documentType = "word",
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
            var token = CreateJwt(configJson);

            // Return final config with token field at top-level (OnlyOffice expects it there when JWT is enabled)
            var finalJson = "{\"token\":" + serializer.Serialize(token) + ",\"document\":" + serializer.Serialize(configObject.document) + ",\"documentType\":" + serializer.Serialize(configObject.documentType) + ",\"editorConfig\":" + serializer.Serialize(configObject.editorConfig) + "}";
            return finalJson;
        }

        protected override void Render(HtmlTextWriter writer)
        {
            // Support direct download request without touching UI.
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

        private static string Base64UrlEncode(byte[] input)
        {
            return Convert.ToBase64String(input)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private static byte[] HmacSha256(byte[] key, byte[] data)
        {
            using (var h = new HMACSHA256(key))
            {
                return h.ComputeHash(data);
            }
        }

        private static bool FixedTimeEquals(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length) return false;
            var diff = 0;
            for (int i = 0; i < a.Length; i++) diff |= a[i] ^ b[i];
            return diff == 0;
        }

        private string CreateJwt(string payloadJson)
        {
            var headerJson = "{\"alg\":\"HS256\",\"typ\":\"JWT\"}";
            var header = Base64UrlEncode(Encoding.UTF8.GetBytes(headerJson));
            var payload = Base64UrlEncode(Encoding.UTF8.GetBytes(payloadJson));
            var signingInput = header + "." + payload;
            var signatureBytes = HmacSha256(Encoding.UTF8.GetBytes(OnlyOfficeJwtSecret), Encoding.UTF8.GetBytes(signingInput));
            var signature = Base64UrlEncode(signatureBytes);
            return signingInput + "." + signature;
        }

        private bool ValidateJwt(string token)
        {
            var parts = token.Split('.');
            if (parts.Length != 3) return false;

            var signingInput = parts[0] + "." + parts[1];
            var expectedSig = Base64UrlEncode(HmacSha256(Encoding.UTF8.GetBytes(OnlyOfficeJwtSecret), Encoding.UTF8.GetBytes(signingInput)));
            var actualSig = parts[2];
            return FixedTimeEquals(Encoding.ASCII.GetBytes(expectedSig), Encoding.ASCII.GetBytes(actualSig));
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

        private static string GuessDocumentType(string fileType)
        {
            if (string.IsNullOrWhiteSpace(fileType)) return "text";
            fileType = fileType.ToLowerInvariant();

            switch (fileType)
            {
                case "doc":
                case "docx":
                case "odt":
                case "rtf":
                case "txt":
                case "pdf":
                    return "text";
                case "xls":
                case "xlsx":
                case "ods":
                case "csv":
                    return "spreadsheet";
                case "ppt":
                case "pptx":
                case "odp":
                    return "presentation";
                default:
                    return "text";
            }
        }

        private static string ComputeKey(string fileId, string physicalPath)
        {
            var lastWrite = File.GetLastWriteTimeUtc(physicalPath).Ticks.ToString();
            var raw = fileId + ":" + lastWrite;
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(raw));
                var sb = new StringBuilder(bytes.Length * 2);
                for (int i = 0; i < bytes.Length; i++) sb.Append(bytes[i].ToString("x2"));
                // OnlyOffice key size limitations; keep it short but stable.
                return sb.ToString(0, 20);
            }
        }

        private string AbsoluteUrl(string relative)
        {
            var url = ResolveUrl(relative);

            var cfgBase = ConfigurationManager.AppSettings["OnlyOffice.PublicBaseUrl"];
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

        private static int TryGetInt(dynamic dict, string key)
        {
            if (dict == null) return 0;
            if (!(dict is System.Collections.IDictionary d)) return 0;
            if (!d.Contains(key) || d[key] == null) return 0;
            int val;
            if (int.TryParse(d[key].ToString(), out val)) return val;
            return 0;
        }

        private static string TryGetString(dynamic dict, string key)
        {
            if (dict == null) return null;
            if (!(dict is System.Collections.IDictionary d)) return null;
            if (!d.Contains(key) || d[key] == null) return null;
            return d[key].ToString();
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