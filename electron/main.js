const path = require("path");
const { pathToFileURL } = require("url");
const { app, BrowserWindow, shell, session } = require("electron");

const indexFilePath = path.join(__dirname, "..", "public", "index.html");
const indexFileUrl = pathToFileURL(indexFilePath).href;
const allowedRemoteOrigins = new Set([
  "https://cdn.jsdelivr.net",
  "https://fonts.googleapis.com",
  "https://fonts.gstatic.com"
]);

function isAllowedInternalNavigation(url) {
  return url === indexFileUrl || url.startsWith(`${indexFileUrl}#`);
}

function isAllowedExternalUrl(rawUrl) {
  try {
    const parsed = new URL(rawUrl);
    return parsed.protocol === "https:" || parsed.protocol === "http:";
  } catch (error) {
    return false;
  }
}

function isAllowedRequestUrl(rawUrl) {
  try {
    const parsed = new URL(rawUrl);
    if (parsed.protocol === "file:" || parsed.protocol === "data:" || parsed.protocol === "blob:") {
      return true;
    }
    if (parsed.protocol === "devtools:") {
      return true;
    }
    return allowedRemoteOrigins.has(parsed.origin);
  } catch (error) {
    return false;
  }
}

function createMainWindow() {
  const iconPath = path.join(__dirname, "..", "public", "assets", "acb-white.png");
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1100,
    minHeight: 700,
    icon: iconPath,
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webSecurity: true,
      allowRunningInsecureContent: false,
      devTools: !app.isPackaged,
      navigateOnDragDrop: false,
      safeDialogs: true,
      spellcheck: false
    }
  });

  win.loadFile(indexFilePath);

  // Prevent navigation away from the packaged app entrypoint.
  win.webContents.on("will-navigate", (event, url) => {
    if (!isAllowedInternalNavigation(url)) {
      event.preventDefault();
    }
  });

  win.webContents.on("will-redirect", (event, url) => {
    if (!isAllowedInternalNavigation(url)) {
      event.preventDefault();
    }
  });

  // Keep external links in the user's default browser.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (!isAllowedExternalUrl(url)) {
      return { action: "deny" };
    }
    shell.openExternal(url).catch((err) => {
      console.error("Failed to open external URL:", err);
    });
    return { action: "deny" };
  });
}

app.whenReady().then(() => {
  if (process.platform === "win32") {
    app.setAppUserModelId("com.customs.billingportal");
  }

  // Deny all permission prompts by default (camera/mic/notifications/etc).
  session.defaultSession.setPermissionRequestHandler((_webContents, _permission, callback) => {
    callback(false);
  });
  session.defaultSession.setPermissionCheckHandler(() => false);
  session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
    callback({ cancel: !isAllowedRequestUrl(details.url) });
  });

  createMainWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
});

app.on("web-contents-created", (_event, contents) => {
  contents.on("will-attach-webview", (event) => {
    event.preventDefault();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
