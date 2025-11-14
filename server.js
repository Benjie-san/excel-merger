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

// Render your index.ejs page (optional)
app.set("views", path.join(__dirname, "views"));
app.set("view engine", "ejs");

app.get("/", (req, res) => {
  res.render("index");
});

app.get("/api", (req, res) => {
  res.json({ msg: "Hello world" });
});

app.listen(port, () =>
  console.log(`Static server running at http://localhost:${port}`)
);
