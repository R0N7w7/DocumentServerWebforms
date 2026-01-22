using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Collections.Concurrent;
using System.Net;

namespace WebEditor
{
    public partial class _Default : Page
    {
        private const string UploadFolderVirtual = "~/App_Data/uploads";

        // Centraliza aquí la configuración de red/OnlyOffice.
        // - PublicBaseUrl: URL pública de esta app (la que el Document Server puede alcanzar)
        // - DocumentServerUrl: URL del OnlyOffice Document Server (si algún día la usas desde el cliente/servidor)
        // - JwtSecret: secreto compartido con Document Server
        private static class OnlyOfficeSettings
        {
            // Ejemplo: "http://192.168.10.34:2355" (sin slash final)
            public static string PublicBaseUrlOverride { get; set; } = "http://192.168.10.34:2355";

            // Ejemplo: "http://192.168.10.50:8080" (si necesitas referenciar el DS desde tu app)
            public static string DocumentServerUrl { get; set; } = "http://192.168.10.34:8085";

            // Debe coincidir con el valor configurado en OnlyOffice Document Server
            public static string JwtSecret { get; set; } = "secreto_personalizado";

            public static string GetPublicBaseUrl(HttpRequest request)
            {
                if (!string.IsNullOrWhiteSpace(PublicBaseUrlOverride))
                    return PublicBaseUrlOverride.TrimEnd('/');

                // Fallback: usa la URL de la petición actual.
                // OJO: esto sólo funciona si Document Server puede resolver/alcanzar esa misma URL.
                if (request?.Url == null)
                    return null;

                var baseUri = request.Url;
                var appPath = request.ApplicationPath ?? "/";

                // Normaliza para que quede: scheme://host[:port]/[app]
                var path = appPath == "/" ? "/" : appPath.TrimEnd('/') + "/";
                var builder = new UriBuilder(baseUri.Scheme, baseUri.Host, baseUri.Port, path);
                return builder.Uri.ToString().TrimEnd('/');
            }
        }

        private sealed class SavedUrlInfo
        {
            public string Url { get; set; }
            public DateTime UtcSavedAt { get; set; }
        }

        private static bool TryReadOnlyOfficeKeyFromJwt(string token, out string key)
        {
            key = null;

            try
            {
                var parts = token.Split('.');
                if (parts.Length != 3) return false;

                var payloadJson = Base64UrlDecodeToString(parts[1]);
                if (string.IsNullOrWhiteSpace(payloadJson)) return false;

                var serializer = new JavaScriptSerializer();
                var root = serializer.DeserializeObject(payloadJson) as System.Collections.IDictionary;
                if (root == null) return false;

                var payload = root["payload"] as System.Collections.IDictionary;
                if (payload == null) return false;

                if (payload.Contains("key") && payload["key"] != null)
                {
                    key = payload["key"].ToString();
                }

                return !string.IsNullOrWhiteSpace(key);
            }
            catch
            {
                return false;
            }
        }

        private static readonly ConcurrentDictionary<string, SavedUrlInfo> LastSavedUrlByDocKey =
            new ConcurrentDictionary<string, SavedUrlInfo>(StringComparer.OrdinalIgnoreCase);

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
            hfDocKey.Value = GenerateOnlyOfficeDocumentKey(fileId);
            btnDownload.Enabled = true;

            OnlyOfficeConfigJson = BuildOnlyOfficeConfigJson(originalName, storedName);
            litStatus.Text = "<div class='text-success mt-2'>Archivo subido. Cargando editor…</div>";
        }

        private static string GenerateOnlyOfficeDocumentKey(string fileId)
        {
            // Mantener un formato tipo `shardkey_suffix` como el que se observa en los callbacks/logs.
            // Esto hace que `document.key` (y por ende el key devuelto por OnlyOffice) sea estable y predecible.
            // fileId ya viene como hex de 32 chars, tomamos 20 para shard y 4 para sufijo.
            if (string.IsNullOrWhiteSpace(fileId))
                return Guid.NewGuid().ToString("N");

            var clean = fileId.Replace("-", "");
            if (clean.Length < 24)
                clean = clean.PadRight(24, '0');

            var shard = clean.Substring(0, 20);
            var suffix = clean.Substring(clean.Length - 4);
            return shard + "_" + suffix;
        }

        protected void btnDownload_Click(object sender, EventArgs e)
        {
            var fileId = hfFileId.Value;
            if (string.IsNullOrWhiteSpace(fileId))
            {
                litStatus.Text = "<div class='text-danger mt-2'>No hay archivo para descargar.</div>";
                return;
            }

            var docKey = hfDocKey.Value;
            var url = "~/Default.aspx?onlyoffice=proxydownload&fileId=" + HttpUtility.UrlEncode(fileId);
            if (!string.IsNullOrWhiteSpace(docKey))
            {
                url += "&key=" + HttpUtility.UrlEncode(docKey);
            }

            Response.Redirect(url, endResponse: false);
            Context.ApplicationInstance.CompleteRequest();
        }

        private bool IsCallbackRequest()
        {
            return string.Equals(Request.QueryString["onlyoffice"], "callback", StringComparison.OrdinalIgnoreCase);
        }

        private void ProcessOnlyOfficeCallback()
        {
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
            if (string.IsNullOrWhiteSpace(callbackToken))
            {
                callbackToken = TryGetBearerToken(Request);
            }

            if (!string.IsNullOrWhiteSpace(callbackToken) && !ValidateJwt(callbackToken))
            {
                RespondJson("{\"error\":1}");
                return;
            }

            // status/url pueden venir en el body o dentro del JWT (Authorization: Bearer)
            var status = TryGetInt(payload, "status");
            var downloadUrl = TryGetString(payload, "url");
            if (string.IsNullOrWhiteSpace(downloadUrl))
            {
                var dataObj = TryGetDict(payload, "data");
                downloadUrl = TryGetString(dataObj, "url");
            }

            if (status == 0 && string.IsNullOrWhiteSpace(downloadUrl) && !string.IsNullOrWhiteSpace(callbackToken))
            {
                int jwtStatus;
                string jwtUrl;
                if (TryReadOnlyOfficeCallbackFromJwt(callbackToken, out jwtStatus, out jwtUrl))
                {
                    status = jwtStatus;
                    downloadUrl = jwtUrl;
                }
            }

            var fileId = Request.QueryString["fileId"];
            if (string.IsNullOrWhiteSpace(fileId))
            {
                RespondJson("{\"error\":1}");
                return;
            }

            // IMPORTANT: OnlyOffice identifica el documento por `key` (no por nuestro fileId). En Docker logs
            // se ve `key` como algo tipo `bfd32f7bdab188ada7ca_6047`. Ese valor viene en el callback.
            string docKey = TryGetString(payload, "key");
            if (string.IsNullOrWhiteSpace(docKey) && !string.IsNullOrWhiteSpace(callbackToken))
            {
                TryReadOnlyOfficeKeyFromJwt(callbackToken, out docKey);
            }

            TryLogOnlyOfficeCallback(body, status, downloadUrl, (docKey ?? "") + " fileId=" + fileId);

            // status 2 = MustSave, 6 = MustForceSave, 7 = Corrupted (puede venir con url en algunas configs)
            if ((status == 2 || status == 6 || status == 7) && !string.IsNullOrWhiteSpace(downloadUrl))
            {
                // Guardar por docKey si existe; fallback por fileId.
                var storeKey = !string.IsNullOrWhiteSpace(docKey) ? docKey : fileId;
                LastSavedUrlByDocKey[storeKey] = new SavedUrlInfo
                {
                    Url = downloadUrl,
                    UtcSavedAt = DateTime.UtcNow
                };
            }

            RespondJson("{\"error\":0}");
        }

        private static void TryLogOnlyOfficeCallback(string body, int status, string url, string fileId)
        {
            try
            {
                var ctx = HttpContext.Current;
                if (ctx == null) return;
                var logPath = ctx.Server.MapPath("~/App_Data/onlyoffice-callback.log");
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                File.AppendAllText(logPath,
                    DateTime.UtcNow.ToString("o") + " fileId=" + fileId + " status=" + status + " url=" + (url ?? "") + Environment.NewLine +
                    body + Environment.NewLine + Environment.NewLine);
            }
            catch
            {
                // ignore
            }
        }

        private static System.Collections.IDictionary TryGetDict(dynamic dict, string key)
        {
            if (dict == null) return null;
            if (!(dict is System.Collections.IDictionary d)) return null;
            if (!d.Contains(key) || d[key] == null) return null;
            return d[key] as System.Collections.IDictionary;
        }

        private static string TryGetBearerToken(HttpRequest request)
        {
            var auth = request?.Headers?["Authorization"];
            if (string.IsNullOrWhiteSpace(auth)) return null;
            const string prefix = "Bearer ";
            if (!auth.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return null;
            return auth.Substring(prefix.Length).Trim();
        }

        // Lee claims del JWT sin validar firma (ya se validó arriba). ONLYOFFICE mete `status`/`url` dentro de `payload`.
        private static bool TryReadOnlyOfficeCallbackFromJwt(string token, out int status, out string url)
        {
            status = 0;
            url = null;

            try
            {
                var parts = token.Split('.');
                if (parts.Length != 3) return false;

                var payloadJson = Base64UrlDecodeToString(parts[1]);
                if (string.IsNullOrWhiteSpace(payloadJson)) return false;

                var serializer = new JavaScriptSerializer();
                var root = serializer.DeserializeObject(payloadJson) as System.Collections.IDictionary;
                if (root == null) return false;

                var payload = root["payload"] as System.Collections.IDictionary;
                if (payload == null) return false;

                if (payload.Contains("status") && payload["status"] != null)
                {
                    int.TryParse(payload["status"].ToString(), out status);
                }

                if (payload.Contains("url") && payload["url"] != null)
                {
                    url = payload["url"].ToString();
                }

                return status != 0 || !string.IsNullOrWhiteSpace(url);
            }
            catch
            {
                return false;
            }
        }

        private static string Base64UrlDecodeToString(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            var s = input.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4)
            {
                case 2: s += "=="; break;
                case 3: s += "="; break;
            }

            var bytes = Convert.FromBase64String(s);
            return Encoding.UTF8.GetString(bytes);
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
            var token = CreateJwt(configJson);

            // Return final config with token field at top-level (OnlyOffice expects it there when JWT is enabled)
            var finalJson = "{\"token\":" + serializer.Serialize(token) + ",\"document\":" + serializer.Serialize(configObject.document) + ",\"documentType\":" + serializer.Serialize(configObject.documentType) + ",\"editorConfig\":" + serializer.Serialize(configObject.editorConfig) + "}";
            return finalJson;
        }

        protected override void Render(HtmlTextWriter writer)
        {
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

            if (string.Equals(Request.QueryString["onlyoffice"], "proxydownload", StringComparison.OrdinalIgnoreCase))
            {
                // La descarga debe consultar por `key` (docKey), no por `fileId`, porque el callback se indexa por key.
                var docKey = Request.QueryString["key"];
                var fileId = Request.QueryString["fileId"];
                var lookupKey = !string.IsNullOrWhiteSpace(docKey) ? docKey : fileId;

                if (string.IsNullOrWhiteSpace(lookupKey))
                {
                    Response.StatusCode = 400;
                    Response.End();
                    return;
                }

                if (!LastSavedUrlByDocKey.TryGetValue(lookupKey, out var saved) || string.IsNullOrWhiteSpace(saved?.Url))
                {
                    // Fallback: intentar encontrar una URL guardada que contenga el `fileId` (o parte) en la ruta.
                    if (!string.IsNullOrWhiteSpace(fileId))
                    {
                        foreach (var kvp in LastSavedUrlByDocKey)
                        {
                            var u = kvp.Value?.Url;
                            if (!string.IsNullOrWhiteSpace(u) && u.IndexOf(fileId, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                saved = kvp.Value;
                                break;
                            }
                        }
                    }

                    if (saved == null || string.IsNullOrWhiteSpace(saved.Url))
                    {
                    // Aún no llegó callback con status=2/6, por lo tanto no hay URL del archivo modificado.
                    Response.StatusCode = 409;
                    Response.ContentType = "application/json";
                    Response.Write("{\"error\":1,\"message\":\"No hay versión guardada todavía.\"}");
                    Response.End();
                    return;
                    }
                }

                ProxyDownloadFromOnlyOffice(saved.Url, fileId);
                return;
            }

            if (string.Equals(Request.QueryString["onlyoffice"], "savestatus", StringComparison.OrdinalIgnoreCase))
            {
                var docKey = Request.QueryString["key"];
                var fileId = Request.QueryString["fileId"];
                var lookupKey = !string.IsNullOrWhiteSpace(docKey) ? docKey : fileId;

                SavedUrlInfo saved = null;
                if (!string.IsNullOrWhiteSpace(lookupKey))
                {
                    LastSavedUrlByDocKey.TryGetValue(lookupKey, out saved);

                    if ((saved == null || string.IsNullOrWhiteSpace(saved.Url)) && !string.IsNullOrWhiteSpace(fileId))
                    {
                        foreach (var kvp in LastSavedUrlByDocKey)
                        {
                            var u = kvp.Value?.Url;
                            if (!string.IsNullOrWhiteSpace(u) && u.IndexOf(fileId, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                saved = kvp.Value;
                                break;
                            }
                        }
                    }
                }

                var has = saved != null && !string.IsNullOrWhiteSpace(saved.Url);
                var ageMs = has ? (long)Math.Max(0, (DateTime.UtcNow - saved.UtcSavedAt).TotalMilliseconds) : -1;

                Response.Clear();
                Response.StatusCode = 200;
                Response.ContentType = "application/json";
                Response.Write("{\"saved\":" + (has ? "true" : "false") + ",\"ageMs\":" + ageMs.ToString() + "}");
                Response.End();
                return;
            }

            base.Render(writer);
        }

        private void ProxyDownloadFromOnlyOffice(string url, string fileId)
        {
            // Nombre de descarga: intenta usar el original si existe.
            var storedName = FindStoredName(fileId);
            var downloadName = storedName ?? (fileId + ".docx");

            Response.Clear();
            Response.BufferOutput = false;
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("Content-Disposition", "attachment; filename=\"" + downloadName.Replace("\"", "") + "\"");

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.AllowAutoRedirect = true;
            request.ReadWriteTimeout = 30000;
            request.Timeout = 30000;

            using (var remote = (HttpWebResponse)request.GetResponse())
            {
                if (!string.IsNullOrWhiteSpace(remote.ContentType))
                {
                    Response.ContentType = remote.ContentType;
                }
                if (remote.ContentLength > 0)
                {
                    Response.AddHeader("Content-Length", remote.ContentLength.ToString());
                }

                using (var src = remote.GetResponseStream())
                {
                    if (src == null)
                    {
                        Response.StatusCode = 502;
                        Response.End();
                        return;
                    }

                    src.CopyTo(Response.OutputStream);
                }
            }

            Response.Flush();
            Response.End();
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
            var signatureBytes = HmacSha256(Encoding.UTF8.GetBytes(OnlyOfficeSettings.JwtSecret), Encoding.UTF8.GetBytes(signingInput));
            var signature = Base64UrlEncode(signatureBytes);
            return signingInput + "." + signature;
        }

        private bool ValidateJwt(string token)
        {
            var parts = token.Split('.');
            if (parts.Length != 3) return false;

            var signingInput = parts[0] + "." + parts[1];
            var expectedSig = Base64UrlEncode(HmacSha256(Encoding.UTF8.GetBytes(OnlyOfficeSettings.JwtSecret), Encoding.UTF8.GetBytes(signingInput)));
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