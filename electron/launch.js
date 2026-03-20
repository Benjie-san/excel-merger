const path = require("path");
const { spawn } = require("child_process");

const electronBinary = require("electron");
const mainEntry = path.join(__dirname, "main.js");
const extraArgs = process.argv.slice(2);
const childEnv = { ...process.env };
delete childEnv.ELECTRON_RUN_AS_NODE;

const child = spawn(electronBinary, [mainEntry, ...extraArgs], {
  stdio: "inherit",
  env: childEnv
});

child.on("exit", (code) => process.exit(code || 0));
child.on("error", (err) => {
  console.error("Failed to launch Electron:", err);
  process.exit(1);
});
