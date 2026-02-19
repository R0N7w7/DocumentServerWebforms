<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebEditor._Default" ValidateRequest="false" %>
<%@ Register Src="~/Controls/OnlyOfficeEditor.ascx" TagPrefix="oo" TagName="Editor" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <link rel="stylesheet" href="<%= ResolveUrl("~/Content/WebEditor.css") %>" />

    <div class="we">
        <section class="we-toolbar" aria-label="Acciones">
            <div class="we-toolbar__left">
                <asp:FileUpload ID="fuFile" runat="server" CssClass="form-control we-input" />
            </div>
            <div class="we-toolbar__right">
                <asp:Button ID="btnUpload" runat="server" Text="Subir y abrir" CssClass="btn we-btn we-btn-accent" OnClick="btnUpload_Click" />
            </div>
            <div class="we-toolbar__status">
                <asp:Literal ID="litStatus" runat="server" />
            </div>
        </section>

        <div class="we-layout">
            <main class="we-main" aria-label="Editor">
                <div class="we-surface">
                    <div class="we-surface__head">
                        <div class="we-surface__title">Documento</div>
                        <div class="we-surface__meta">Vista del editor OnlyOffice</div>
                    </div>
                    <%-- ═══ Editor modular reutilizable ═══ --%>
                    <oo:Editor ID="docEditor" runat="server"
                        CaptureTriggerId="btnDescargar" />
                </div>
            </main>

            <aside class="we-aside" aria-label="Lista de cambios">
                <div class="we-surface we-surface--sticky">
                    <div class="we-surface__head">
                        <div class="we-surface__title">Cambios a realizar</div>
                        <div class="we-surface__meta">Pendientes</div>
                    </div>

                    <div class="we-changes">
                        <div class="we-change">
                            <div class="we-change__top">
                                <span class="we-change__title">Actualizar el encabezado</span>
                            </div>
                            <div class="we-change__desc">Reemplazar el nombre de la empresa y fecha en la portada.</div>
                        </div>

                        <div class="we-change">
                            <div class="we-change__top">
                                <span class="we-change__title">Corregir tabla de precios</span>
                            </div>
                            <div class="we-change__desc">Ajustar totales y aplicar formato de moneda.</div>
                        </div>

                        <div class="we-change">
                            <div class="we-change__top">
                                <span class="we-change__title">Unificar tipografías</span>
                            </div>
                            <div class="we-change__desc">Usar un solo estilo de fuente en títulos y párrafos.</div>
                        </div>

                        <div class="we-change">
                            <div class="we-change__top">
                                <span class="we-change__title">Agregar nota legal</span>
                            </div>
                            <div class="we-change__desc">Incluir cláusula estándar al final del documento.</div>
                        </div>
                    </div>

                    <div class="we-divider"></div>

                    <div class="we-asideActions">
                        <asp:Button ID="btnDescargar" runat="server" Text="Guardar y descargar"
                            CssClass="btn we-btn we-btn-accent we-btn-block" OnClick="btnDescargar_Click" />
                        <div class="we-asideHint">Descarga o guarda la versión editada tras completar los cambios.</div>
                    </div>
                </div>
            </aside>
        </div>
    </div>
</asp:Content>
