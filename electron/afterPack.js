const fs = require("fs");
const path = require("path");
const rcedit = require("rcedit");

module.exports = async function afterPack(context) {
  if (context.electronPlatformName !== "win32") return;

  const projectDir = context.packager.projectDir;
  const iconPath = path.join(projectDir, "build", "icon.ico");
  const exeName = `${context.packager.appInfo.productFilename}.exe`;
  const exePath = path.join(context.appOutDir, exeName);

  if (!fs.existsSync(iconPath)) {
    console.warn(`[afterPack] Icon not found: ${iconPath}`);
    return;
  }

  if (!fs.existsSync(exePath)) {
    console.warn(`[afterPack] Executable not found: ${exePath}`);
    return;
  }

  await rcedit(exePath, { icon: iconPath });
  console.log(`[afterPack] Applied icon to ${exePath}`);
};
