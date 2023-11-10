import * as XLSX from "xlsx/xlsx.mjs";
/* load 'fs' for readFile and writeFile support */
import * as fs from "fs";
XLSX.set_fs(fs);
/* load 'stream' for stream support */
import { Readable } from "stream";
XLSX.stream.set_readable(Readable);
/* load the codepage support library for extended support with older formats  */
import * as cpexcel from "xlsx/dist/cpexcel.full.mjs";
XLSX.set_cptable(cpexcel);

import { join } from "path";

const year = new Date().getFullYear();
const month = new Date().getMonth() + 1;

let folder = join("output");
// check if the folder output exists
if (!fs.existsSync(folder)) {
  fs.mkdirSync(folder);
}
// create year folder if doesn't exist
folder = join("output", `${year}`);
if (!fs.existsSync(folder)) {
  fs.mkdirSync(folder);
}
// create month folder if doesn't exist
folder = join("output", `${year}`, `${month}`);
if (!fs.existsSync(folder)) {
  fs.mkdirSync(folder);
}

const inputFileName = join("input", "default.xlsx");

// Read the workbook
const workbook = XLSX.readFile(inputFileName);

// Get the first sheet
const sheetName = workbook.SheetNames[0];
let worksheet = workbook.Sheets[sheetName];

// Convert the sheet to JSON
let data = XLSX.utils.sheet_to_json(worksheet, { raw: true, defval: "" });

for (const row of data) {
  for (const key in row) {
    if(key.startsWith("Fecha Desde")) {
      row[key] = `15/${month}/${year}`;
      continue;
    }
    if (key.startsWith("Fecha Hasta")) {
      row[key] = `15/${month}/${year}`;
      continue;
    }
    if(key.startsWith("Descripción de la Evidencia")) {
      row[key] = "La descripción es un no se que y un no se cuanto";
      continue;
    }
  }
}

// Convert the JSON back to a worksheet
worksheet = XLSX.utils.json_to_sheet(data);

// Replace the old sheet
workbook.Sheets[sheetName] = worksheet;

// Write the workbook back to the file
const outputFileName = join("output", `${year}`, `${month}`, `diego.maroto_Evidencias_${year}_${month}.xlsx`);
XLSX.writeFile(workbook, outputFileName);
