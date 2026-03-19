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

// Serve everything from /public as static files
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(port, () =>
  console.log(`Static server running at http://localhost:${port}`)
);
