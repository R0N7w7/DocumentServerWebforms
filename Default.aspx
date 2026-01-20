<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebEditor._Default" ValidateRequest="false" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <style>
        iframe {
            aspect-ratio: 16 / 9;
            min-height: 100%;
        }
    </style>
    <div class="row">
        <div class="col-12">
            <h2>Editor (OnlyOffice)</h2>

            <asp:Panel runat="server" CssClass="card mb-3">
                <div class="card-body">
                    <div class="mb-3">
                        <asp:FileUpload ID="fuFile" runat="server" CssClass="form-control" />
                    </div>
                    <div class="d-flex gap-2">
                        <asp:Button ID="btnUpload" runat="server" Text="Subir y abrir" CssClass="btn btn-primary" OnClick="btnUpload_Click" />
                        <asp:Button ID="btnDownload" runat="server" Text="Descargar" CssClass="btn btn-secondary" OnClick="btnDownload_Click" Enabled="false" OnClientClick="return window.WebEditor_downloadWithSave();" />
                    </div>
                    <asp:Literal ID="litStatus" runat="server" />
                </div>
            </asp:Panel>

            <asp:HiddenField ID="hfDocKey" runat="server" />
            <asp:HiddenField ID="hfFileId" runat="server" />

            <div id="onlyoffice-editor" style="width: 100%; height: 900px; border: 1px solid #ddd;"></div>
        </div>
    </div>

    <asp:PlaceHolder runat="server">
        <script type="text/javascript" src="http://192.168.137.213:8085/web-apps/apps/api/documents/api.js"></script>
        <script type="text/javascript">
            (function () {
                var cfg = <%= OnlyOfficeConfigJson %>;
                if (!cfg) return;
                if (!cfg.document || !cfg.editorConfig) return;
                if (!cfg.document.url) return;

                if (window._docEditor && window._docEditor.destroyEditor) {
                    try { window._docEditor.destroyEditor(); } catch (e) { }
                }
                window._docEditor = new DocsAPI.DocEditor("onlyoffice-editor", cfg);
            })();

            window.WebEditor_downloadWithSave = function () {
                // Si el editor no está listo aún, deja que el postback ocurra.
                if (!window._docEditor) return true;

                // Solicita guardado explícito al Document Server y espera un momento,
                // para dar tiempo a que llegue el callback y se sobrescriba el archivo en disco.
                try {
                    if (window._docEditor.requestSave) {
                        window._docEditor.requestSave();
                        setTimeout(function () {
                            __doPostBack('<%= btnDownload.UniqueID %>', '');
                        }, 1500);
                        return false;
                    }
                } catch (e) { }

                return true;
            };
        </script>
    </asp:PlaceHolder>
</asp:Content>
