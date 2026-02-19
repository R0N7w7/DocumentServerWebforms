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

            docEditor.SetDocumentFromBytes(fuFile.FileBytes, fuFile.FileName);
            litStatus.Text = string.Empty;
        }

        /// <summary>
        /// Se dispara cuando el usuario hace clic en "Guardar en servidor".
        /// El control ya capturó el documento; basta con llamar GetEditedDocumentBytes().
        /// </summary>
        /// <summary>
        /// Descarga el documento editado en el navegador del cliente.
        /// </summary>
        protected void btnDescargar_Click(object sender, EventArgs e)
        {
            byte[] documentBytes = docEditor.GetEditedDocumentBytes();
            if (documentBytes == null || documentBytes.Length == 0)
            {
                litStatus.Text = "<span class='text-warning'>No hay documento editado para descargar.</span>";
                return;
            }

            docEditor.ClearEditedDocument();

            var fileName = "editado_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".docx";
            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + fileName);
            Response.AddHeader("Content-Length", documentBytes.Length.ToString());
            Response.BinaryWrite(documentBytes);
            Response.End();
        }
    }
}