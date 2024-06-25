// index.js
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const bodyParser = require("body-parser");
const fs = require("fs");
const dataRouter = require("./routes/crud");

const path = require("path");

const excelFilePath = path.join(__dirname, "/data/data.xlsx");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(bodyParser.json());

app.use("/crud", dataRouter);

const port = 3000;

// Ensure the Excel file exists
if (!fs.existsSync(excelFilePath)) {
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet([]);
  xlsx.utils.book_append_sheet(wb, ws, "Sheet1");
  xlsx.writeFile(wb, excelFilePath);
}

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
