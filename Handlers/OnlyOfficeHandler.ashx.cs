using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Web;

namespace WebEditor.Handlers
{
    /// <summary>
    /// Handler HTTP genérico para servir documentos temporales al Document Server
    /// y recibir callbacks de guardado.
    ///
    /// Endpoints:
    ///   ?action=download&amp;fileId=xxx  → Sirve el archivo almacenado en App_Data/uploads.
    ///   ?action=callback&amp;fileId=xxx  → Recibe notificación POST del Document Server.
    ///   ?action=proxy&amp;url=xxx        → Descarga un archivo del Document Server (server-to-server) y lo reenvía al cliente. Evita problemas de CORS.
    /// </summary>
    public class OnlyOfficeHandler : IHttpHandler
    {
        private const string UploadFolder = "~/App_Data/uploads";

        public bool IsReusable => false;

        public void ProcessRequest(HttpContext context)
        {
            var action = context.Request.QueryString["action"];
            if (string.IsNullOrWhiteSpace(action))
            {
                context.Response.StatusCode = 400;
                context.Response.Write("Parámetro 'action' requerido.");
                return;
            }

            switch (action.ToLowerInvariant())
            {
                case "download":
                    ServeFile(context);
                    break;
                case "callback":
                    HandleCallback(context);
                    break;
                case "proxy":
                    ProxyDownload(context);
                    break;
                default:
                    context.Response.StatusCode = 400;
                    context.Response.Write("Acción no reconocida: " + action);
                    break;
            }
        }

        /// <summary>
        /// Sirve un archivo de App_Data/uploads para que Document Server lo cargue.
        /// </summary>
        private static void ServeFile(HttpContext context)
        {
            var fileId = context.Request.QueryString["fileId"];
            if (string.IsNullOrWhiteSpace(fileId))
            {
                context.Response.StatusCode = 400;
                context.Response.Write("Parámetro 'fileId' requerido.");
                return;
            }

            var uploadsDir = context.Server.MapPath(UploadFolder);
            if (!Directory.Exists(uploadsDir))
            {
                context.Response.StatusCode = 404;
                context.Response.Write("Directorio de uploads no encontrado.");
                return;
            }

            var matches = Directory.GetFiles(uploadsDir, fileId + ".*");
            if (matches.Length == 0)
            {
                context.Response.StatusCode = 404;
                context.Response.Write("Archivo no encontrado para fileId: " + fileId);
                return;
            }

            var filePath = matches[0];
            var fileName = Path.GetFileName(filePath);

            context.Response.Clear();
            context.Response.ContentType = "application/octet-stream";
            context.Response.AddHeader(
                "Content-Disposition",
                "attachment; filename=\"" + fileName.Replace("\"", "") + "\"");
            context.Response.TransmitFile(filePath);
            context.Response.End();
        }

        /// <summary>
        /// Recibe callbacks POST del Document Server (eventos de guardado).
        /// Responde siempre con {"error":0} para confirmar recepción.
        /// </summary>
        private static void HandleCallback(HttpContext context)
        {
            // Aquí se puede extender para procesar el JSON del callback:
            //   var body = new StreamReader(context.Request.InputStream).ReadToEnd();
            //   var data = new JavaScriptSerializer().Deserialize<Dictionary<string,object>>(body);
            //   int status = Convert.ToInt32(data["status"]);
            //   if (status == 2 || status == 6) { /* documento guardado → descargar de data["url"] */ }

            context.Response.Clear();
            context.Response.ContentType = "application/json";
            context.Response.Write("{\"error\":0}");
            context.Response.End();
        }

        /// <summary>
        /// Descarga un archivo del Document Server (server-to-server) y lo reenvía
        /// al cliente. Esto evita el bloqueo de CORS que ocurre cuando el browser
        /// intenta hacer fetch() directamente al Document Server.
        /// </summary>
        private static void ProxyDownload(HttpContext context)
        {
            var url = context.Request.QueryString["url"];
            if (string.IsNullOrWhiteSpace(url))
            {
                context.Response.StatusCode = 400;
                context.Response.Write("Parámetro 'url' requerido.");
                return;
            }

            try
            {
                var uri = new Uri(url);

                var request = (HttpWebRequest)WebRequest.Create(uri);
                request.Method = "GET";
                request.Timeout = 60000; // 60 segundos

                using (var response = (HttpWebResponse)request.GetResponse())
                using (var stream = response.GetResponseStream())
                {
                    context.Response.Clear();
                    context.Response.ContentType = response.ContentType
                        ?? "application/octet-stream";

                    // Pasar Content-Disposition si existe
                    var cd = response.Headers["Content-Disposition"];
                    if (!string.IsNullOrWhiteSpace(cd))
                        context.Response.AddHeader("Content-Disposition", cd);

                    if (stream != null)
                        stream.CopyTo(context.Response.OutputStream);

                    context.Response.Flush();
                    context.ApplicationInstance.CompleteRequest();
                }
            }
            catch (ThreadAbortException)
            {
                // Response.End() de otros handlers puede causar esto — ignorar
            }
            catch (WebException wex)
            {
                context.Response.StatusCode = 502;
                context.Response.ContentType = "text/plain";
                context.Response.Write("WebException: " + wex.Message);
                if (wex.Status != WebExceptionStatus.Success)
                    context.Response.Write(" | Status: " + wex.Status);
                if (wex.InnerException != null)
                    context.Response.Write(" | Inner: " + wex.InnerException.Message);
                if (wex.Response is HttpWebResponse errResp)
                    context.Response.Write(" | HTTP " + (int)errResp.StatusCode);
            }
            catch (Exception ex)
            {
                context.Response.StatusCode = 502;
                context.Response.ContentType = "text/plain";
                context.Response.Write("Error: " + ex.GetType().Name + ": " + ex.Message);
                if (ex.InnerException != null)
                    context.Response.Write(" | Inner: " + ex.InnerException.GetType().Name + ": " + ex.InnerException.Message);
            }
        }
    }
}
