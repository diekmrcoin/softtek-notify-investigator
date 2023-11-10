import * as fs from "fs";
import { join } from "path";
import { Readable } from "stream";
import * as XLSX from "xlsx/xlsx.mjs";
import * as cpexcel from "xlsx/dist/cpexcel.full.mjs";

XLSX.set_fs(fs);
XLSX.stream.set_readable(Readable);
XLSX.set_cptable(cpexcel);

class WorkbookManager {
  constructor(year, month, day) {
    this.year = year;
    this.month = month;
    this.day = day;
    this.database = {};
    this.database[year] = {};
    this.database[year][month] = {};
    this.databaseFile = join("database");
    this.folder = join("output");
    this.inputFileName = join("input", "default.xlsx");
  }

  setup() {
    this.loadDatabase();
    // check if the folder output exists
    if (!fs.existsSync(this.folder)) {
      fs.mkdirSync(this.folder);
    }
    // create year folder if doesn't exist
    this.folder = join("output", `${this.year}`);
    if (!fs.existsSync(this.folder)) {
      fs.mkdirSync(this.folder);
    }
    // create month folder if doesn't exist
    this.folder = join("output", `${this.year}`, `${this.month}`);
    if (!fs.existsSync(this.folder)) {
      fs.mkdirSync(this.folder);
    }
  }

  loadDatabase() {
    // create database folder and file if doesn't exist
    if (!fs.existsSync(this.databaseFile)) {
      fs.mkdirSync(this.databaseFile);
    }
    this.databaseFile = join("database", "database.json");
    if (!fs.existsSync(this.databaseFile)) {
      fs.writeFileSync(
        this.databaseFile,
        JSON.stringify(this.database),
        "utf8"
      );
      console.log("Database created");
    } else {
      const data = fs.readFileSync(this.databaseFile, "utf8");
      this.database = JSON.parse(data);
      console.log("Database loaded");
    }
  }

  populate() {
    console.log(`Populating ${this.year}/${this.month}`);
    // Read the workbook
    const workbook = XLSX.readFile(this.inputFileName, { cellStyles: true });
    // Get the first sheet
    const sheetName = workbook.SheetNames[0];
    let worksheet = workbook.Sheets[sheetName];

    // Get current date
    const currentDate = new Date(this.year, this.month, this.day);
    // format DD/MM/AAAA
    const formattedDate = `${("0" + currentDate.getDate()).slice(-2)}/${(
      "0" +
      (currentDate.getMonth() + 1)
    ).slice(-2)}/${currentDate.getFullYear()}`;

    // Add data to specific cells
    worksheet["A2"] = { v: formattedDate, t: "s" };
    worksheet["B2"] = { v: formattedDate, t: "s" };
    worksheet["H2"] = { v: "dummy", t: "s" };
    worksheet["J2"] = { v: "dummy", t: "s" };

    // Write the workbook back to the file
    const outputFileName = join(
      "output",
      `${this.year}`,
      `${this.month}`,
      `diego.maroto_Evidencias_${this.year}_${this.month}.xlsx`
    );
    XLSX.writeFile(workbook, outputFileName, { bookType: "xlsx" });
  }
}

export { WorkbookManager };
