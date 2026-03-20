// ----------------------------------------------------------
// STATIC SERVER ONLY — NO EXCEL PROCESSING ON THE BACKEND
// ----------------------------------------------------------
// All Excel merging, cleaning, row trimming, filename labels,
// and reporting will happen CLIENT-SIDE in script.js.
// ----------------------------------------------------------

const express = require('express');
const path = require('path');

const app = express();
const port = process.env.PORT || 8080;

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

app.listen(port, () =>
  console.log(`Static server running at http://localhost:${port}`)
);
