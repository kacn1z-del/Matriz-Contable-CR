// Matriz Contable CR — proceso principal de Electron (Windows)
//
// La app HTML fue escrita para correr dentro de un WebView de React Native.
// Este archivo reemplaza ese puente nativo:
//   - window.ReactNativeWebView.postMessage(...) llega aquí por el protocolo matriz://bridge
//   - "llamarNativo" hace las peticiones HTTPS desde el proceso principal (sin CORS)
//   - "guardarArchivo" / "imprimirArchivo" abren el diálogo para guardar el archivo
//   - "abrirEnlace" abre el navegador del sistema

const { app, BrowserWindow, protocol, shell, net, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { pathToFileURL } = require('url');

const SCHEME = 'matriz';
const ROOT = __dirname;

protocol.registerSchemesAsPrivileged([
  {
    scheme: SCHEME,
    privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true }
  }
]);

// Se inyecta en la página al cargar: crea el objeto que el HTML espera encontrar
const BRIDGE_SHIM = `
  (function () {
    if (window.ReactNativeWebView && window.ReactNativeWebView.postMessage) return;
    window.ReactNativeWebView = {
      postMessage: function (msg) {
        fetch('${SCHEME}://bridge/msg', {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
          body: String(msg)
        }).catch(function (e) { console.error('Puente nativo:', e); });
      }
    };
  })();
`;

function safeFilename(name) {
  const base = path.basename(String(name || 'archivo'));
  return base.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_') || 'archivo';
}

async function handleBridgeMessage(raw, webContents) {
  let data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    return;
  }

  const win = BrowserWindow.fromWebContents(webContents);

  switch (data.tipo) {
    case 'llamarNativo': {
      const { id, url, opciones = {} } = data;
      let ok = true;
      let httpOk = false;
      let texto = '';
      try {
        if (!/^https:\/\//i.test(url)) throw new Error('URL no permitida: ' + url);
        const res = await net.fetch(url, {
          method: opciones.method || 'GET',
          headers: opciones.headers || {},
          body: opciones.body
        });
        httpOk = res.ok;
        texto = await res.text();
      } catch (e) {
        ok = false;
        texto = String((e && e.message) || e);
      }
      await webContents.executeJavaScript(
        `window.recibirRespuestaNativa(${JSON.stringify(id)}, ${ok}, ${httpOk}, ${JSON.stringify(texto)})`
      );
      break;
    }

    case 'guardarArchivo':
    case 'imprimirArchivo': {
      const { filename, base64 } = data;
      const result = await dialog.showSaveDialog(win, {
        defaultPath: safeFilename(filename),
        filters: [{ name: 'Todos los archivos', extensions: ['*'] }]
      });
      if (!result.canceled && result.filePath) {
        await fs.promises.writeFile(result.filePath, Buffer.from(base64, 'base64'));
        if (data.tipo === 'imprimirArchivo') {
          // Abre el PDF con el visor del sistema para que se pueda imprimir desde ahí
          shell.openPath(result.filePath);
        }
      }
      break;
    }

    case 'abrirEnlace': {
      if (/^https?:\/\//i.test(data.url || '')) shell.openExternal(data.url);
      break;
    }

    case 'diagnosticoErrorHoja':
      console.error('Error en la hoja:', data.detalle);
      break;

    default:
      // programarRecordatorioPendiente y otros: sin efecto en escritorio
      break;
  }
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 900,
    minHeight: 600,
    title: 'Matriz Contable CR',
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: false,
      nodeIntegration: false,
      sandbox: true
    }
  });

  win.webContents.on('dom-ready', () => {
    win.webContents.executeJavaScript(BRIDGE_SHIM).catch(() => {});
  });

  win.loadURL(`${SCHEME}://app/index.html`);

  // Enlaces externos (Hacienda, proveedores, WhatsApp, etc.) abren el navegador del sistema
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:\/\//i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  win.webContents.on('will-navigate', (event, url) => {
    if (!url.startsWith(`${SCHEME}://`)) {
      event.preventDefault();
      if (/^https?:\/\//i.test(url)) shell.openExternal(url);
    }
  });
}

app.whenReady().then(() => {
  protocol.handle(SCHEME, async (request) => {
    const url = new URL(request.url);

    // Mensajes del HTML hacia el proceso principal
    if (url.host === 'bridge') {
      const body = await request.text();
      // No bloqueamos la respuesta: el resultado se devuelve con executeJavaScript
      handleBridgeMessage(body, BrowserWindow.getFocusedWindow()?.webContents || BrowserWindow.getAllWindows()[0].webContents)
        .catch((e) => console.error('Error en el puente:', e));
      return new Response('ok', { status: 200 });
    }

    // Archivos de la aplicación
    let rel = decodeURIComponent(url.pathname);
    if (rel === '/' || rel === '') rel = '/index.html';

    const filePath = path.normalize(path.join(ROOT, rel));
    if (!filePath.startsWith(ROOT)) {
      return new Response('Forbidden', { status: 403 });
    }

    return net.fetch(pathToFileURL(filePath).toString());
  });

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
