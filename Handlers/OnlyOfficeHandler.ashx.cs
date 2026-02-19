using System.IO;
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
    }
}
