const express = require("express");
const router = express.Router();
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const excelFilePath = path.join(__dirname, "../data/data.xlsx");

// Ensure the Excel file exists
if (!fs.existsSync(excelFilePath)) {
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet([]);
  xlsx.utils.book_append_sheet(wb, ws, "Sheet1");
  xlsx.writeFile(wb, excelFilePath);
}

router.post("/add", (req, res) => {
  const newData = req.body;
  const wb = xlsx.readFile(excelFilePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(ws);

  data.push(newData);
  const newWs = xlsx.utils.json_to_sheet(data);
  wb.Sheets[wb.SheetNames[0]] = newWs;
  xlsx.writeFile(wb, excelFilePath);

  res.status(201).send(newData);
});

router.get("/get", (req, res) => {
  const wb = xlsx.readFile(excelFilePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(ws);

  res.status(200).json(data);
});

router.put("/update/:id", (req, res) => {
  const id = parseInt(req.params.id, 10);
  const updatedData = req.body;
  const wb = xlsx.readFile(excelFilePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  let data = xlsx.utils.sheet_to_json(ws);

  const index = data.findIndex((item) => item.id == id);
  if (index === -1) {
    return res.status(404).send({ message: "Data not found" });
  }

  data[index] = { ...data[index], ...updatedData };
  const newWs = xlsx.utils.json_to_sheet(data);
  wb.Sheets[wb.SheetNames[0]] = newWs;
  xlsx.writeFile(wb, excelFilePath);

  res.status(200).send(data[index]);
});

router.delete("/data/:id", (req, res) => {
  const id = parseInt(req.params.id, 10);
  const wb = xlsx.readFile(excelFilePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  let data = xlsx.utils.sheet_to_json(ws);

  const index = data.findIndex((item) => item.age === id);
  if (index === -1) {
    return res.status(404).send({ message: "Data not found" });
  }

  data.splice(index, 1);
  const newWs = xlsx.utils.json_to_sheet(data);
  wb.Sheets[wb.SheetNames[0]] = newWs;
  xlsx.writeFile(wb, excelFilePath);

  res.status(204).send();
});

module.exports = router;
