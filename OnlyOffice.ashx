<%@ WebHandler Language="C#" Class="WebEditor.OnlyOffice" %>

using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;

namespace WebEditor
{
    public class OnlyOffice : IHttpHandler
    {
        private const string UploadFolderVirtual = "~/App_Data/uploads";
        private const string OnlyOfficeJwtSecret = "Y1EOwRcQmDQlgzBBTP3aQLAwWvVFlLz2";

        public bool IsReusable => false;

        public void ProcessRequest(HttpContext context)
        {
            LogRequest(context);
            var type = (context.Request.QueryString["type"] ?? "").Trim();

            if (type.Equals("callback", StringComparison.OrdinalIgnoreCase))
            {
                ProcessOnlyOfficeCallback(context);
                return;
            }

            if (type.Equals("download", StringComparison.OrdinalIgnoreCase))
            {
                ProcessDownload(context);
                return;
            }

            context.Response.StatusCode = 400;
            context.Response.ContentType = "application/json";
            context.Response.Write("{\"error\":1}");
        }

        private static void LogRequest(HttpContext context)
        {
            try
            {
                var logPath = context.Server.MapPath("~/App_Data/onlyoffice.log");
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                var sb = new StringBuilder();
                sb.Append(DateTime.UtcNow.ToString("o"));
                sb.Append(" ");
                sb.Append(context.Request.HttpMethod);
                sb.Append(" ");
                sb.Append(context.Request.RawUrl);
                sb.AppendLine();

                sb.Append("  RemoteAddr: ").Append(context.Request.UserHostAddress).AppendLine();
                sb.Append("  UA: ").Append(context.Request.UserAgent).AppendLine();
                sb.Append("  ContentType: ").Append(context.Request.ContentType).AppendLine();
                sb.Append("  ContentLength: ").Append(context.Request.ContentLength).AppendLine();
                sb.Append("  Auth: ").Append(context.Request.Headers["Authorization"] ?? "").AppendLine();
                sb.Append("  Range: ").Append(context.Request.Headers["Range"] ?? "").AppendLine();

                File.AppendAllText(logPath, sb.ToString());
            }
            catch
            {
                // ignore logging errors
            }
        }

        private void ProcessDownload(HttpContext context)
        {
            var fileId = context.Request.QueryString["fileId"];
            var storedName = string.IsNullOrWhiteSpace(fileId) ? null : FindStoredName(context, fileId);
            if (storedName == null)
            {
                context.Response.StatusCode = 404;
                return;
            }

            var physicalPath = GetPhysicalPath(context, storedName);
            TransmitFile(context, physicalPath, storedName);
        }

        private void ProcessOnlyOfficeCallback(HttpContext context)
        {
            string body;
            using (var reader = new StreamReader(context.Request.InputStream))
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
                RespondJson(context, "{\"error\":1}");
                return;
            }

            var callbackToken = TryGetString(payload, "token");
            if (!string.IsNullOrWhiteSpace(callbackToken) && !ValidateJwt(callbackToken))
            {
                RespondJson(context, "{\"error\":1}");
                return;
            }

            var status = TryGetInt(payload, "status");
            var downloadUrl = TryGetString(payload, "url");

            var fileId = context.Request.QueryString["fileId"];
            if (string.IsNullOrWhiteSpace(fileId))
            {
                RespondJson(context, "{\"error\":1}");
                return;
            }

            if ((status == 2 || status == 6) && !string.IsNullOrWhiteSpace(downloadUrl))
            {
                try
                {
                    var storedName = FindStoredName(context, fileId);
                    if (storedName == null)
                    {
                        RespondJson(context, "{\"error\":1}");
                        return;
                    }

                    var physicalPath = GetPhysicalPath(context, storedName);
                    using (var wc = new System.Net.WebClient())
                    {
                        wc.DownloadFile(downloadUrl, physicalPath);
                    }
                }
                catch
                {
                    RespondJson(context, "{\"error\":1}");
                    return;
                }
            }

            RespondJson(context, "{\"error\":0}");
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

        private static void RespondJson(HttpContext context, string json)
        {
            context.Response.Clear();
            context.Response.ContentType = "application/json";
            context.Response.Write(json);
            context.Response.End();
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

        private static bool ValidateJwt(string token)
        {
            var parts = token.Split('.');
            if (parts.Length != 3) return false;

            var signingInput = parts[0] + "." + parts[1];
            var expectedSig = Base64UrlEncode(HmacSha256(Encoding.UTF8.GetBytes(OnlyOfficeJwtSecret), Encoding.UTF8.GetBytes(signingInput)));
            var actualSig = parts[2];
            return FixedTimeEquals(Encoding.ASCII.GetBytes(expectedSig), Encoding.ASCII.GetBytes(actualSig));
        }

        private static string FindStoredName(HttpContext context, string fileId)
        {
            var uploadsPhysical = context.Server.MapPath(UploadFolderVirtual);
            if (!Directory.Exists(uploadsPhysical))
                return null;

            var matches = Directory.GetFiles(uploadsPhysical, fileId + ".*");
            if (matches.Length == 0)
                return null;

            return Path.GetFileName(matches[0]);
        }

        private static string GetPhysicalPath(HttpContext context, string storedName)
        {
            var uploadsPhysical = context.Server.MapPath(UploadFolderVirtual);
            return Path.Combine(uploadsPhysical, storedName);
        }

        private static void TransmitFile(HttpContext context, string physicalPath, string downloadName)
        {
            if (!File.Exists(physicalPath))
            {
                context.Response.StatusCode = 404;
                return;
            }

            context.Response.Clear();
            context.Response.ContentType = "application/octet-stream";
            context.Response.AddHeader("Content-Disposition", "attachment; filename=\"" + downloadName.Replace("\"", "") + "\"");
            context.Response.TransmitFile(physicalPath);
        }
    }
}
