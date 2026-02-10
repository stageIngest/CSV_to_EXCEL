/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as XLSX from 'xlsx';

let fileNo = 1;
let fileInput;
var FileSelected;
var FileName;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    fileInput = document.getElementById("fileInput");
    const run = document.getElementById("run");
    const downloadEnabled = document.getElementById("download");

    fileInput.addEventListener("change", () => {
      if (fileInput.files && fileInput.files.length > 0) {
        console.log("File selezionato:", fileInput.files[0].name);
        document.getElementById("fileName").textContent =
          "Selected file: " + fileInput.files[0].name;
        if (fileInput.files.length > 1) {
          for (let no = 1; no < fileInput.files.length; no++) {
            document.getElementById("fileName").textContent += ", " + fileInput.files[no].name;
          }
        }
      }
    });
    
    run.addEventListener("click", fileImport);
  }
});


// Gestione nel caso di file multipli
async function fileImport() {
  if (!fileInput || !fileInput.files) return;

  fileNo = 1;

  for (let i = 0; i < fileInput.files.length; i++) {
    FileSelected = fileInput.files[i];
    if (FileSelected) {
      FileName = FileSelected.name;
      await ReadFile(FileSelected);
    }
  }
  
  console.log("Importazione completata!");
  alert("Importazione completata!");
}

// VERSIONE CORRETTA - Restituisce una Promise
async function ReadFile(FileSelected) {
  return new Promise((resolve, reject) => {
    const Reader = new FileReader();
    Reader.readAsArrayBuffer(FileSelected);

    Reader.onload = async () => {
      try {
        const buffer = Reader.result;
        if (!buffer) {
          resolve();
          return;
        }

        const CSVText = new TextDecoder().decode(buffer).trim();
        if (!CSVText) {
          resolve();
          return;
        }

        const rows = CSVText
          .split(/\r?\n/)
          .filter(r => r.trim() !== "");

        if (!rows.length) {
          resolve();
          return;
        }

        const CSVData = rows.map((row) => {
          const rowForSplitting = row.replace(/"(\d+),(\d+)"/g, "$1.$2"); 
          const separator = rowForSplitting.includes(";") ? ";" : ",";

          return rowForSplitting.split(separator).map((cell) => setCellAs(cell));
        });
        
        await writeExcel(CSVData);
        resolve(); 
      } catch (err) {
        console.error("Errore CSV:", err);
        reject(err);
      }
    };
    
    Reader.onerror = () => {
      console.error("Errore lettura file");
      reject(new Error("Errore lettura file"));
    };
  });
}

// Interpretazione e formattazione del contenuto della cella 
function setCellAs(cell) {
  let str = cell.trim();
  
  if (str.startsWith('"') && str.endsWith('"')) {
    str = str.slice(1, -1);
  }
  
  let num = Number(str);
  return isNaN(num) ? str : num;
}

// Formato di ogni colonna, controllando tutte le singole celle escludendo l'header
function SetNumericFormat(CSVData) {
  const colCount = CSVData[0].length;
  const numericCols = Array(colCount).fill(true);

  for (let col = 0; col < colCount; col++) {
    if (CSVData[0][col].toString().toLowerCase() == "matricola") {
      numericCols[col] = false;
    }
    else if (CSVData[0][col].toString().toLowerCase().includes('nr.')) {
      numericCols[col] = false;
    }
    
    for (let row = 1; row < CSVData.length; row++) {
      if (typeof CSVData[row][col] !== "number") {
        numericCols[col] = false;
        break;
      }
    }
  }
  return numericCols;
}

// Gestione della scrittura su excel, con controllo dell'overflow rispetto alle colonne dell'header
async function writeExcel(CSVData) {
  if (!CSVData.length) return;

  const colCount = CSVData[0].length;

  for (let i = 0; i < CSVData.length; i++) {
    if (CSVData[i].length < colCount) {
      CSVData[i].push("");
    } else if (CSVData[i].length > colCount) {
      const fixed = CSVData[i].slice(0, colCount - 1);
      const overflow = CSVData[i].slice(colCount - 1).join(",");
      fixed.push(overflow);
      CSVData[i] = fixed;
    }
  }
  await createNewExcel(CSVData);
}

// Creazione del file excel vero
// Vengono salvati i file in diversi fogli sullo stesso file excel
async function createNewExcel(CSVData) {
  await Excel.run(async (context) => {
    let worksheet;
    let CleanName;
    
    if (fileNo == 1) {
      worksheet = context.workbook.worksheets.getActiveWorksheet();
      CleanName = FileName.split(".")[0].substring(0, 31);
      worksheet.name = CleanName;
    } else {
      worksheet = context.workbook.worksheets.add();
      CleanName = FileName.split(".")[0].substring(0, 31);
      worksheet.name = CleanName;
    }

    fileNo++;

    const rowCount = CSVData.length;
    const colCount = CSVData[0].length;

    const range = worksheet.getRangeByIndexes(0, 0, rowCount, colCount);
    range.values = CSVData;

    const isNumeric = SetNumericFormat(CSVData);
    for (let column = 0; column < colCount; column++) {
      if (isNumeric[column]) {
        range.getColumn(column).numberFormat = [["#,##0.00;[Red]-#,##0.00"]];
      }
    }

    const headerRange = range.getRow(0);
    headerRange.format.font.bold = true;

    range.format.autofitColumns();
    range.format.autofitRows();

    await context.sync();
    
    const downloadEnabled = document.getElementById("download");
    if (downloadEnabled && downloadEnabled.checked) {
      await downloadWorksheetAsWorkbook(context, worksheet, CleanName);
    }
  });
}

// Nuova funzione per scaricare un worksheet come workbook separato
async function downloadWorksheetAsWorkbook(context, worksheet, sheetName) {
  try {
    const workbook = XLSX.utils.book_new();
    
    const usedRange = worksheet.getUsedRange();
    usedRange.load("values");
    await context.sync();
    
    const ws = XLSX.utils.aoa_to_sheet(usedRange.values);
    
    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
    
    const excelFileName = sheetName + ".xlsx";
    
    XLSX.writeFile(workbook, excelFileName);
    
    console.log(`File scaricato: ${excelFileName}`);

  } catch (err) {
    console.error("Errore durante il download del workbook:", err);
  }
}