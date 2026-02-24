using System;
using System.IO;
using System.Web.UI;

namespace WebEditor
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (!fuFile.HasFile)
            {
                litStatus.Text = string.Empty;
                return;
            }

            docEditor.SetDocumentFromBytes(fuFile.FileBytes, fuFile.FileName); // Carga el documento en el editor
            litStatus.Text = string.Empty;
        }

        protected void btnDescargar_Click(object sender, EventArgs e)
        {
            // El evento puede acceder al documento editado solo si el ID del disparador se registró en el control OnlyOfficeEditor
            byte[] documentBytes = docEditor.GetEditedDocumentBytes(); // Obtiene el documento editado como un arreglo de bytes

            if (documentBytes == null || documentBytes.Length == 0)
            {
                litStatus.Text = "<span class='text-warning'>No hay documento editado para descargar.</span>";
                return;
            }

            docEditor.ClearEditedDocument();

            var ext = Path.GetExtension(docEditor.DocumentName ?? ".docx");
            var fileName = docEditor.DocumentName ?? ("documento" + ext);
            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + fileName);
            Response.AddHeader("Content-Length", documentBytes.Length.ToString());
            Response.BinaryWrite(documentBytes);
            Response.End();
        }
    }
}