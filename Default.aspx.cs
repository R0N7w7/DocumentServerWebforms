using System;
using System.Web.UI;

namespace WebEditor
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            // La lógica de callbacks y descarga de archivos ahora está en
            // Handlers/OnlyOfficeHandler.ashx (desacoplada de esta página).
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (!fuFile.HasFile)
            {
                litStatus.Text = string.Empty;
                return;
            }

            // Una sola línea: el control se encarga de guardar el archivo,
            // generar URLs, firmar JWT y renderizar el editor.
            docEditor.SetDocumentFromBytes(fuFile.FileBytes, fuFile.FileName);
            litStatus.Text = string.Empty;
        }
    }
}