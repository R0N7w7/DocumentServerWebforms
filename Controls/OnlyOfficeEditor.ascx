<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="OnlyOfficeEditor.ascx.cs" Inherits="WebEditor.Controls.OnlyOfficeEditor" %>

<div id="<%= EditorContainerId %>_wrapper" style="position:relative;">
    <div id="<%= EditorContainerId %>_busy" class="we-busyOverlay" style="display:none;" aria-live="polite" aria-busy="true">
        <div class="we-busyOverlay__card" role="status">
            <span class="we-spinner we-spinner--lg" aria-hidden="true"></span>
            <div class="we-busyOverlay__text">Procesando&#8230;</div>
        </div>
    </div>
    <div id="<%= EditorContainerId %>" class="we-editor" style="min-height:<%= EditorHeight %>;"></div>
</div>

<% if (HasDocument) { %>
<script type="text/javascript" src="<%= OnlyOfficeApiUrl %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/Scripts/OnlyOfficeEditor.js") %>"></script>
<script type="text/javascript">
    (function () {
        var cfg = <%= ConfigJson %>;
        if (cfg) {
            OnlyOfficeEditorModule.init('<%= EditorContainerId %>', cfg);
        }
    })();
</script>
<% } %>
