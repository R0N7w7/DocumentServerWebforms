<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebEditor._Default" ValidateRequest="false" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <link rel="stylesheet" href="<%= ResolveUrl("~/Content/WebEditor.css") %>" />

    <div class="we">
        <div id="weBusyOverlay" class="we-busyOverlay" style="display:none" aria-live="polite" aria-busy="true">
            <div class="we-busyOverlay__card" role="status" aria-label="Procesando">
                <span class="we-spinner we-spinner--lg" aria-hidden="true"></span>
                <div class="we-busyOverlay__text">Procesando…</div>
            </div>
        </div>

        <section class="we-toolbar" aria-label="Acciones">
            <div class="we-toolbar__left">
                <asp:FileUpload ID="fuFile" runat="server" CssClass="form-control we-input" />
            </div>
            <div class="we-toolbar__right">
                <asp:Button ID="btnUpload" runat="server" Text="Subir y abrir" CssClass="btn we-btn we-btn-accent" OnClick="btnUpload_Click" />
                <asp:Button ID="btnDownload" runat="server" Text="Guardar y descargar" CssClass="btn we-btn we-btn-ghost" OnClick="btnDownload_Click" Enabled="false" Style="display:none" />
            </div>
            <div class="we-toolbar__status">
                <asp:Literal ID="litStatus" runat="server" />
            </div>
        </section>

        <asp:HiddenField ID="hfDocKey" runat="server" />
        <asp:HiddenField ID="hfFileId" runat="server" />
        <asp:HiddenField ID="hfDownloadUniqueId" runat="server" Value="" />

        <div class="we-layout">
            <main class="we-main" aria-label="Editor">
                <div class="we-surface">
                    <div class="we-surface__head">
                        <div class="we-surface__title">Documento</div>
                        <div class="we-surface__meta">Vista del editor OnlyOffice</div>
                    </div>
                    <div id="onlyoffice-editor" class="we-editor"></div>
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
                        <button type="button" class="btn we-btn we-btn-accent we-btn-block" onclick="return window.WebEditor_trySaveThenDownload();">Guardar y descargar</button>
                        <div class="we-asideHint">Descarga la versión guardada tras completar los cambios.</div>
                    </div>
                </div>
            </aside>
        </div>
    </div>

    <asp:PlaceHolder runat="server">
        <script type="text/javascript" src="http://192.168.10.34:8085/web-apps/apps/api/documents/api.js"></script>
        <script type="text/javascript">
            (function () {
                var cfg = <%= OnlyOfficeConfigJson %>;
                if (!cfg) return;
                if (!cfg.document || !cfg.editorConfig) return;
                if (!cfg.document.url) return;

                function setBusy(isBusy) {
                    try {
                        var el = document.getElementById('weBusyOverlay');
                        if (!el) return;
                        el.style.display = isBusy ? 'flex' : 'none';
                    } catch (e) { }
                }
                window.WebEditor_setBusy = setBusy;

                setBusy(true);

                // Prefer official customization options when available.
                cfg.editorConfig.customization = cfg.editorConfig.customization || {};
                cfg.editorConfig.customization.uiTheme = (cfg.editorConfig.customization.uiTheme || 'theme-classic-light');
                cfg.editorConfig.customization.compactToolbar = true;
                cfg.editorConfig.customization.toolbarNoTabs = true;
                cfg.editorConfig.customization.hideRightMenu = true;
                cfg.editorConfig.customization.hideRulers = true;
                cfg.editorConfig.customization.showReviewChanges = true;

                // Hook OnlyOffice events so downloadAs can return a URL immediately.
                cfg.events = cfg.events || {};
                cfg.events.onAppReady = function () { setBusy(false); };
                cfg.events.onDownloadAs = function (evt) {
                    setBusy(false);
                    try {
                        var data = evt && evt.data;
                        if (data && data.url) {
                            window.location.href = data.url;
                            return;
                        }
                    } catch (e) { }

                    // Fallback to legacy flow if the URL is not present.
                    if (typeof window.WebEditor_startLegacySaveThenDownload === 'function') {
                        window.WebEditor_startLegacySaveThenDownload();
                    }
                };

                if (window._docEditor && window._docEditor.destroyEditor) {
                    try { window._docEditor.destroyEditor(); } catch (e) { }
                }
                window._docEditor = new DocsAPI.DocEditor("onlyoffice-editor", cfg);
                setTimeout(function () { setBusy(false); }, 800);

                // Force CSS into the editor iframe (best-effort).
                // Note: will only work if the iframe is same-origin or if the browser allows access.
                window.WebEditor_applyEditorIframeTheme = function () {
                    try {
                        var host = document.getElementById('onlyoffice-editor');
                        if (!host) return false;
                        var iframe = host.querySelector('iframe');
                        if (!iframe) return false;
                        var doc = iframe.contentDocument || (iframe.contentWindow && iframe.contentWindow.document);
                        if (!doc) return false;

                        var id = 'we-iframe-theme';
                        if (doc.getElementById(id)) return true;

                        var style = doc.createElement('style');
                        style.id = id;
                        style.type = 'text/css';
                        style.appendChild(doc.createTextNode(
                            ':root{--we-accent:#7c9383;}\n'
                            + 'html,body{background:#fff !important;}\n'
                            + '.toolbar, .toolbar-box, .toolbar-group{border-color: rgba(229,231,235,.9) !important;}\n'
                            + '.btn, button{border-radius:12px !important;}\n'
                            + 'button.primary, .btn.primary, .button--primary{background:var(--we-accent) !important; border-color:var(--we-accent) !important;}\n'
                            + 'a, .link{color:var(--we-accent) !important;}\n'
                            + '*:focus{outline:none !important; box-shadow: 0 0 0 4px rgba(124,147,131,.18) !important;}\n'
                        ));
                        (doc.head || doc.documentElement).appendChild(style);
                        return true;
                    } catch (e) {
                        return false;
                    }
                };

                (function retryTheme() {
                    var attempts = 0;
                    var maxAttempts = 40; // ~10s
                    var timer = setInterval(function () {
                        attempts++;
                        if (window.WebEditor_applyEditorIframeTheme()) {
                            clearInterval(timer);
                        } else if (attempts >= maxAttempts) {
                            clearInterval(timer);
                        }
                    }, 250);
                })();
            })();

            window.WebEditor_startLegacySaveThenDownload = function () {
                var fileId = '<%= (hfFileId.Value ?? string.Empty).Replace("'", "") %>';
                var key = '<%= (hfDocKey.Value ?? string.Empty).Replace("'", "") %>';
                var downloadUniqueId = '<%= (hfDownloadUniqueId.Value ?? string.Empty).Replace("'", "") %>';

                function doDownloadPostBack() {
                    try {
                        if (typeof __doPostBack === 'function' && downloadUniqueId) {
                            __doPostBack(downloadUniqueId, '');
                            return;
                        }
                    } catch (e) { }

                    // fallback: navigate to server handler directly
                    try {
                        var url = 'Default.aspx?onlyoffice=proxydownload&fileId=' + encodeURIComponent(fileId);
                        if (key) url += '&key=' + encodeURIComponent(key);
                        window.location.href = url;
                    } catch (e) { }
                }

                try {
                    var el = document.getElementById('weBusyOverlay');
                    if (el) el.style.display = 'flex';
                } catch (e) { }

                try {
                    if (window._docEditor && window._docEditor.requestSave) {
                        window._docEditor.requestSave();
                    }
                } catch (e) { }

                var started = Date.now();
                var timeoutMs = 15000;
                var intervalMs = 500;

                function poll() {
                    var elapsed = Date.now() - started;
                    if (elapsed > timeoutMs) {
                        doDownloadPostBack();
                        return;
                    }

                    var url = 'Default?onlyoffice=savestatus&fileId=' + encodeURIComponent(fileId);
                    if (key) url += '&key=' + encodeURIComponent(key);

                    try {
                        fetch(url, { cache: 'no-store' })
                            .then(function (r) { return r.json(); })
                            .then(function (j) {
                                if (j && j.saved === true) {
                                    doDownloadPostBack();
                                    return;
                                }
                                setTimeout(poll, intervalMs);
                            })
                            .catch(function () {
                                setTimeout(poll, intervalMs);
                            });
                    } catch (e) {
                        setTimeout(poll, intervalMs);
                    }
                }

                poll();
                return false;
            };

            window.WebEditor_trySaveThenDownload = function () {
                var editor = window._docEditor;
                if (editor && typeof editor.downloadAs === 'function') {
                    try { if (typeof window.WebEditor_setBusy === 'function') window.WebEditor_setBusy(true); } catch (e) { }
                    try {
                        editor.downloadAs();
                        return false;
                    } catch (e) { }
                }

                return window.WebEditor_startLegacySaveThenDownload();
            };

        </script>
    </asp:PlaceHolder>
</asp:Content>
