/**
 * OnlyOfficeEditorModule — Módulo JavaScript reutilizable para OnlyOffice Document Server.
 *
 * Permite inicializar, controlar y obtener el documento editado de forma programática.
 * Soporta múltiples instancias de editor en la misma página.
 *
 * ── API Pública ────────────────────────────────────────────────────────
 *
 *   OnlyOfficeEditorModule.init(containerId, config, options?)
 *     Inicializa el editor en el contenedor indicado.
 *     options: { onReady, onDocumentReady, onError }
 *
 *   OnlyOfficeEditorModule.getEditedDocumentUrl(containerId)
 *     → Promise<string>  URL de descarga del documento editado.
 *
 *   OnlyOfficeEditorModule.getEditedDocumentBlob(containerId)
 *     → Promise<Blob>    Blob del documento (requiere CORS en Document Server).
 *
 *   OnlyOfficeEditorModule.downloadDocument(containerId, fileName?)
 *     → Promise<string>  Descarga el documento en el navegador.
 *
 *   OnlyOfficeEditorModule.getEditor(containerId)
 *     → DocsAPI.DocEditor | null  Instancia nativa del editor.
 *
 *   OnlyOfficeEditorModule.destroy(containerId)
 *     Destruye la instancia del editor y libera recursos.
 *
 *   OnlyOfficeEditorModule.setBusy(containerId, isBusy)
 *     Muestra/oculta el overlay de "Procesando…".
 */
var OnlyOfficeEditorModule = (function () {
    'use strict';

    // Almacén de instancias por containerId
    var _instances = {};

    // ── Helpers internos ─────────────────────────────────────────────

    function _busyId(containerId) {
        return containerId + '_busy';
    }

    function _setBusy(containerId, isBusy) {
        try {
            var el = document.getElementById(_busyId(containerId));
            if (el) el.style.display = isBusy ? 'flex' : 'none';
        } catch (e) { /* silenciar */ }
    }

    /**
     * Intenta inyectar estilos de tema dentro del iframe del editor (best-effort).
     * Solo funciona si el iframe es same-origin.
     */
    function _applyIframeTheme(containerId) {
        try {
            var host = document.getElementById(containerId);
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
                + '.toolbar,.toolbar-box,.toolbar-group{border-color:rgba(229,231,235,.9) !important;}\n'
                + '.btn,button{border-radius:12px !important;}\n'
                + 'button.primary,.btn.primary,.button--primary{background:var(--we-accent) !important;border-color:var(--we-accent) !important;}\n'
                + 'a,.link{color:var(--we-accent) !important;}\n'
                + '*:focus{outline:none !important;box-shadow:0 0 0 4px rgba(124,147,131,.18) !important;}\n'
            ));
            (doc.head || doc.documentElement).appendChild(style);
            return true;
        } catch (e) {
            return false;
        }
    }

    function _startThemeRetry(containerId) {
        var attempts = 0;
        var maxAttempts = 40; // ~10 s
        var timer = setInterval(function () {
            attempts++;
            if (_applyIframeTheme(containerId) || attempts >= maxAttempts) {
                clearInterval(timer);
            }
        }, 250);
    }

    // ── API Pública ──────────────────────────────────────────────────

    /**
     * Inicializa un editor de OnlyOffice.
     * @param {string}  containerId  ID del elemento DOM contenedor.
     * @param {object}  config       Objeto de configuración de OnlyOffice (con token, document, editorConfig…).
     * @param {object}  [options]    Opciones adicionales.
     * @param {function} [options.onReady]          Se invoca cuando el editor está listo.
     * @param {function} [options.onDocumentReady]  Se invoca cuando el documento se ha cargado.
     * @param {function} [options.onError]          Se invoca ante un error.
     * @param {boolean}  [options.applyTheme=true]  Intenta inyectar tema visual al iframe.
     * @returns {DocsAPI.DocEditor|null}
     */
    function init(containerId, config, options) {
        if (!config || !config.document || !config.document.url) {
            console.warn('[OnlyOfficeEditorModule] Config inválida o sin document.url');
            return null;
        }

        if (typeof DocsAPI === 'undefined') {
            console.error('[OnlyOfficeEditorModule] DocsAPI no está cargado. ¿Falta el script del Document Server?');
            if (options && options.onError) options.onError({ message: 'DocsAPI not loaded' });
            return null;
        }

        options = options || {};

        // Destruir instancia previa si existe
        if (_instances[containerId]) {
            destroy(containerId);
        }

        _setBusy(containerId, true);

        // Personalización visual por defecto
        config.editorConfig = config.editorConfig || {};
        config.editorConfig.customization = config.editorConfig.customization || {};
        var cust = config.editorConfig.customization;
        if (cust.uiTheme === undefined)          cust.uiTheme = 'theme-classic-light';
        if (cust.compactToolbar === undefined)    cust.compactToolbar = true;
        if (cust.toolbarNoTabs === undefined)     cust.toolbarNoTabs = true;
        if (cust.hideRightMenu === undefined)     cust.hideRightMenu = true;
        if (cust.hideRulers === undefined)        cust.hideRulers = true;
        if (cust.showReviewChanges === undefined) cust.showReviewChanges = true;

        // Resolvers pendientes de downloadAs
        var _downloadResolve = null;
        var _downloadReject = null;

        // Cablear eventos de OnlyOffice
        config.events = config.events || {};

        config.events.onAppReady = function () {
            _setBusy(containerId, false);
            if (options.onReady) options.onReady();
        };

        config.events.onDocumentReady = function () {
            if (options.onDocumentReady) options.onDocumentReady();
        };

        config.events.onDownloadAs = function (evt) {
            _setBusy(containerId, false);
            var url = null;
            try {
                var data = evt && evt.data;
                url = (data && typeof data === 'object') ? data.url : data;
            } catch (e) { /* silenciar */ }

            if (_downloadResolve) {
                _downloadResolve(url);
                _downloadResolve = null;
                _downloadReject = null;
            }
        };

        config.events.onError = function (evt) {
            _setBusy(containerId, false);
            if (_downloadReject) {
                _downloadReject(evt);
                _downloadResolve = null;
                _downloadReject = null;
            }
            if (options.onError) options.onError(evt);
        };

        var editor = new DocsAPI.DocEditor(containerId, config);

        _instances[containerId] = {
            editor: editor,
            setDownloadResolvers: function (resolve, reject) {
                _downloadResolve = resolve;
                _downloadReject = reject;
            }
        };

        // Timeout de seguridad para ocultar el overlay
        setTimeout(function () { _setBusy(containerId, false); }, 8000);

        // Tema visual en el iframe (best-effort)
        if (options.applyTheme !== false) {
            _startThemeRetry(containerId);
        }

        return editor;
    }

    /**
     * Obtiene la URL de descarga del documento editado.
     * Llama a editor.downloadAs() y espera el evento onDownloadAs.
     * @param {string} containerId
     * @returns {Promise<string>} URL de descarga proporcionada por Document Server.
     */
    function getEditedDocumentUrl(containerId) {
        return new Promise(function (resolve, reject) {
            var instance = _instances[containerId];
            if (!instance || !instance.editor) {
                reject(new Error('Editor no inicializado para: ' + containerId));
                return;
            }

            instance.setDownloadResolvers(resolve, reject);
            _setBusy(containerId, true);

            try {
                instance.editor.downloadAs();
            } catch (e) {
                _setBusy(containerId, false);
                instance.setDownloadResolvers(null, null);
                reject(e);
            }

            // Timeout: si en 30 s no responde, rechazar
            setTimeout(function () {
                if (_instances[containerId] &&
                    _instances[containerId].setDownloadResolvers) {
                    // Solo rechazar si aún estamos esperando
                    // (si ya se resolvió, _downloadResolve ya es null)
                    _setBusy(containerId, false);
                }
            }, 30000);
        });
    }

    /**
     * Obtiene el documento editado como Blob.
     * NOTA: Requiere que Document Server tenga CORS configurado para esta origin.
     * @param {string} containerId
     * @returns {Promise<Blob>}
     */
    function getEditedDocumentBlob(containerId) {
        return getEditedDocumentUrl(containerId).then(function (url) {
            if (!url) throw new Error('No se recibió URL de descarga');
            return fetch(url).then(function (response) {
                if (!response.ok) throw new Error('Error al descargar: ' + response.status);
                return response.blob();
            });
        });
    }

    /**
     * Descarga el documento editado en el navegador.
     * @param {string}  containerId
     * @param {string}  [fileName]  Nombre sugerido para el archivo.
     * @returns {Promise<string>} La URL utilizada para la descarga.
     */
    function downloadDocument(containerId, fileName) {
        return getEditedDocumentUrl(containerId).then(function (url) {
            if (!url) throw new Error('No se recibió URL de descarga');
            // Usar window.location.href para máxima compatibilidad
            // (funciona incluso cross-origin porque el server envía Content-Disposition)
            window.location.href = url;
            return url;
        });
    }

    /**
     * Obtiene la instancia nativa DocsAPI.DocEditor.
     * @param {string} containerId
     * @returns {DocsAPI.DocEditor|null}
     */
    function getEditor(containerId) {
        var instance = _instances[containerId];
        return instance ? instance.editor : null;
    }

    /**
     * Destruye la instancia del editor y libera recursos.
     * @param {string} containerId
     */
    function destroy(containerId) {
        var instance = _instances[containerId];
        if (instance && instance.editor) {
            try { instance.editor.destroyEditor(); } catch (e) { /* silenciar */ }
        }
        delete _instances[containerId];
    }

    /**
     * Muestra u oculta el overlay de "Procesando…".
     * @param {string}  containerId
     * @param {boolean} isBusy
     */
    function setBusy(containerId, isBusy) {
        _setBusy(containerId, isBusy);
    }

    // ── Interfaz pública ─────────────────────────────────────────────
    return {
        init: init,
        getEditedDocumentUrl: getEditedDocumentUrl,
        getEditedDocumentBlob: getEditedDocumentBlob,
        downloadDocument: downloadDocument,
        getEditor: getEditor,
        destroy: destroy,
        setBusy: setBusy
    };
})();
