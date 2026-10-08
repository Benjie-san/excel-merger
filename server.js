// ----------------------------------------------------------
// STATIC SERVER ONLY — NO EXCEL PROCESSING ON THE BACKEND
// ----------------------------------------------------------
// All Excel merging, cleaning, row trimming, filename labels,
// and reporting will happen CLIENT-SIDE in script.js.
// ----------------------------------------------------------

const express = require('express');
const os = require('os');
const path = require('path');

const app = express();
const port = process.env.PORT || 8080;
const host = process.env.HOST || "0.0.0.0";

app.disable("x-powered-by");
app.use((req, res, next) => {
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "DENY");
  res.setHeader("Referrer-Policy", "no-referrer");
  next();
});

// Serve everything from /public as static files
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

const server = app.listen(port, host, () => {
  const listeningPort = server.address().port;
  console.log(`Static server listening on ${host}:${listeningPort}`);

  if (host === "0.0.0.0") {
    console.log(`Local: http://localhost:${listeningPort}`);
    const lanAddresses = new Set(
      Object.values(os.networkInterfaces())
        .flat()
        .filter((address) => address?.family === "IPv4" && !address.internal)
        .map((address) => address.address)
    );
    for (const address of lanAddresses) {
      console.log(`Network: http://${address}:${listeningPort}`);
    }
  } else {
    const urlHost = host.includes(":") ? `[${host}]` : host;
    console.log(`Open: http://${urlHost}:${listeningPort}`);
  }
});
