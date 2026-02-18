/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import ExcelJS from 'exceljs';

let fileInput;
const workbooksToSave = []; // Array per workbook: ogni workbook ha nome, buffer di dati, NumericCol, header
let run;
let runBottom;
let save;
let saveBottom;
let csvData;
let fileName;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("fileName").textContent = "";
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    fileInput = document.getElementById("fileInput");
    run = document.getElementById("run");
    runBottom = document.getElementById("run-bottom");
    save = document.getElementById("save");
    saveBottom = document.getElementById("save-bottom");

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
    runBottom.addEventListener("click", fileImport);
    run.addEventListener("click", fileImport);


    save.addEventListener("click", saveAllWorksheets);
    saveBottom.addEventListener("click", saveAllWorksheets);
  }
});



async function fileImport() {
  if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
    alert("Seleziona almeno un file CSV");
    return;
  }
  run.disabled = true;
  runBottom.disabled = true;
  save.disabled = true;
  saveBottom.disabled = true;
  workbooksToSave.length = 0; // Reset array dei workbook 

  try {
    for (let i = 0; i < fileInput.files.length; i++) {
      const file = fileInput.files[i];
      fileName = file.name;

      await processCSVFile(file, fileName);
      await writeInExcel(csvData, fileName);
      await createExcelWorkbook(csvData, fileName);
    }

    // Abilita save solo se ci sono workbook
    if (workbooksToSave.length > 0) {
      save.disabled = false;
      saveBottom.disabled = false;
    }

    run.disabled = false;
    runBottom.disabled = false;

    document.getElementById("fileName").textContent = "";

  } catch (err) {
    console.error("Errore durante l'importazione:", err);
    run.disabled = false;
    runBottom.disabled = false;
  }
}

async function writeInExcel(csvData, fileName) {
  await Excel.run(async (context) => {
    const cleanName = fileName.split(".")[0].substring(0, 31);
    const worksheet = context.workbook.worksheets.add(cleanName);
    const rowNo = csvData.length;
    const colNo = csvData[0].length;
    const numericCols = getNumericColumns(csvData);
    const range = worksheet.getRangeByIndexes(0, 0, rowNo, colNo);
    range.values = csvData;


    for (let col = 0; col < colNo; col++) {
      if (numericCols[col]) {
        const columnRange = worksheet.getRangeByIndexes(1, col, rowNo - 1, 1);
        columnRange.numberFormat = [["#,##0.00;[Red]-#,##0.00"]];
      }
    }

    const headerRange = worksheet.getRangeByIndexes(0, 0, 1, colNo);
    headerRange.format.font.bold = true;
    range.format.autofitColumns();
    range.format.autofitRows();

    await context.sync();
  })
}

// Legge e processa un singolo CSV
async function processCSVFile(file, fileName) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);

    reader.onload = async () => {
      try {
        const buffer = reader.result;
        if (!buffer) {
          resolve();
          return;
        }
        const csvText = new TextDecoder().decode(buffer).trim();
        if (!csvText) {
          resolve();
          return;
        }
        const rows = csvText.split(/\r?\n/).filter(r => r.trim() !== ""); //divisione per righe, esclusione righe vuote
        if (!rows.length) {
          resolve();
          return;
        }
        csvData = rows.map((row) => {
          const rowForSplitting = row.replace(/"(\d+),(\d+)"/g, "$1.$2"); //formattazione campo decimale come numero.numero, semplifica lo split
          const separator = rowForSplitting.includes(";") ? ";" : ","; //split righe in celle
          return rowForSplitting.split(separator).map((cell) => processCell(cell)); //processa celle
        });
        resolve();
      } catch (err) {
        console.error("Errore CSV:", err);
        reject(err);
      }
    };

    reader.onerror = () => {
      reject(new Error("Errore lettura file"));
    };
  });
}

// processa cella CSV, se numero decimale fa parsefloat altrimenti stringa (gli interi sono stringhe)
function processCell(cell) {
  const str = cell.trim();
  const numericStr = str.replace(/,/g, '.');

  if (/\d+\.\d+/.test(numericStr)) {
    return parseFloat(numericStr);
  }
  return str;
}

// Determina quali colonne sono decimali per settarle come numeriche
function getNumericColumns(csvData) {
  const colCount = csvData[0].length;
  const numericCols = Array(colCount).fill(false);

  for (let col = 0; col < colCount; col++) {
    const header = csvData[0][col].toString().toLowerCase();

    // Escludi esplicitamente colonne specifiche (facilità e semplicità, senza scansionare tutti i dati sempre)
    if (header === "matricola" || header.includes('nr.')) {
      numericCols[col] = false;
      continue;
    }

    let hasDecimals = false;

    for (let row = 1; row < csvData.length; row++) {
      const value = csvData[row][col];

      // Se è un numero decimale (non intero) hasdecimal è true per questa colonna, altrimenti false di default
      // questo perchè nel caso in una colonna decimale ci sia solo 1 decimale anche tutti gli altri interi dovranno essere scritti coerentemente
      if (/\d+\.\d+/.test(value)) {
        hasDecimals = true;
        break; //quando trovo il primo decimale salto alla prossima colonna
      }
    }
    numericCols[col] = hasDecimals;
  }
  return numericCols;
}

// Crea workbook Excel con ExcelJS (da installare npm install exceljs)
async function createExcelWorkbook(csvData, fileName) {
  const workbook = new ExcelJS.Workbook();

  const sheetName = fileName.split(".")[0].substring(0, 31); //numero massimo caratteri accettati
  const worksheet = workbook.addWorksheet(sheetName);

  const colCount = csvData[0].length;

  // Gestisci overflow colonne, le colonne in più senza header sono accorpate nell'ultima con valori separati da ,
  for (let i = 0; i < csvData.length; i++) {
    if (csvData[i].length < colCount) {
      csvData[i].push("");
    } else if (csvData[i].length > colCount) {
      const fixed = csvData[i].slice(0, colCount - 1);
      const overflow = csvData[i].slice(colCount - 1).join(",");
      fixed.push(overflow);
      csvData[i] = fixed;
    }
  }

  const numericCols = getNumericColumns(csvData);

  //per ogni colonna gestiamo formati, header e scrittura dei numeri, registriamo questi dati nell'array columns
  const columns = [];
  for (let col = 0; col < colCount; col++) {
    const columnDef = {
      header: csvData[0][col], //set header
      key: `${col}`, //numero colonna
      width: 15 //larghezza fissa di default
    };

    if (numericCols[col]) {
      columnDef.style = {
        numFmt: '#,##0.00;[Red]-#,##0.00'
      };
    }

    columns.push(columnDef);
  }

  worksheet.columns = columns;

  // Aggiungi dati (salta header perché già definito in columns)
  for (let row = 1; row < csvData.length; row++) {
    const rowData = {};
    for (let col = 0; col < colCount; col++) {
      rowData[`${col}`] = csvData[row][col]; //per ogni riga scansioniamo valori con dati format dall'array colonna
    }
    worksheet.addRow(rowData); //aggiunta formattata nel worksheet provvisorio
  }

  // Formattazione header
  const headerRow = worksheet.getRow(1); //excel js parte da 1, Excel parte da 0
  headerRow.font = { bold: true };

  // Applica formato numerico esplicitamente a ogni cella
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) { // Skip header (riga 1)
      row.eachCell((cell, colNumber) => {
        if (numericCols[colNumber - 1] && typeof cell.value === 'number') {
          cell.numFmt = '#,##0.00;[Red]-#,##0.00';
        }
      });
    }
  });

  // Genera buffer Excel
  const buffer = await workbook.xlsx.writeBuffer();

  const excelFileName = fileName.slice(".")[0] + ".xlsx";

  // Salva con metadata delle colonne numeriche
  workbooksToSave.push({
    name: excelFileName,
    buffer: buffer,
    numericCols: numericCols,
    headers: csvData[0]
  });

  console.log(`Workbook creato: ${excelFileName}`);
}

// Aggiorna i buffer con i dati modificati dall'utente in Excel
async function updateWorkbooksFromExcel() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    // Per ogni workbook in workbooksToSave
    for (let i = 0; i < workbooksToSave.length; i++) {
      const workbookData = workbooksToSave[i];
      const sheetName = workbookData.name.replace('.xlsx', ''); //per ricerca

      // Trova il worksheet corrispondente
      const sheet = sheets.items.find(s => s.name === sheetName.substring(0,31));
      if (!sheet) {
        console.warn(`Worksheet ${sheetName} non trovato, salto aggiornamento`);
        continue;
      }

      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();

      const updatedCsvData = usedRange.values;

      // Rigenera il buffer Excel con i dati aggiornati
      const newBuffer = await regenerateExcelBuffer(
        updatedCsvData,
        workbookData.numericCols,
        sheetName
      );

      // Aggiorna il buffer in workbooksToSave
      workbooksToSave[i].buffer = newBuffer;

      console.log(`Buffer aggiornato per: ${sheetName}`);
    }
  });
}

// Rigenera solo il buffer Excel (minore peso nell'elaborazione)
async function regenerateExcelBuffer(csvData, numericCols, sheetName) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName);
  const colCount = csvData[0].length;

  const columns = [];
  for (let col = 0; col < colCount; col++) {
    const columnDef = {
      header: csvData[0][col],
      key: `${col}`,
      width: 15
    };

    if (numericCols[col]) {
      columnDef.style = {
        numFmt: '#,##0.00;[Red]-#,##0.00'
      };
    }
    columns.push(columnDef);
  }

  worksheet.columns = columns;

  for (let row = 1; row < csvData.length; row++) {
    const rowData = {};
    for (let col = 0; col < colCount; col++) {
      rowData[`${col}`] = csvData[row][col];
    }
    worksheet.addRow(rowData);
  }

  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true };

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.eachCell((cell, colNumber) => {
        if (numericCols[colNumber - 1] && typeof cell.value === 'number') {
          cell.numFmt = '#,##0.00;[Red]-#,##0.00';
        }
      });
    }
  });

  return await workbook.xlsx.writeBuffer();
}

// Salva tutti i workbook nella cartella selezionata, NELLA STESSA, SE NE SELEZIONA SOLO 1
async function saveAllWorksheets() {
  try {
    //aggiorna i buffer con i dati modificati dall'utente (sempre per sicurezza)
    await updateWorkbooksFromExcel();

    const dir = await window.showDirectoryPicker();

    for (const workbookData of workbooksToSave) {
      const { name, buffer } = workbookData;

      const fileHandle = await dir.getFileHandle(name, { create: true });
      const writable = await fileHandle.createWritable(); //stream per scrittura file

      await writable.write({
        type: 'write',
        data: buffer
      });

      await writable.close();
      document.getElementById("fileName").textContent += `${name} SALVATO CON SUCCESSO,   `;
    }
    await Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  for (const workbookData of workbooksToSave) { //elimino solo i foglio creati perchè se li elimino tutti excel da errore 
    const sheetName = workbookData.name.replace(".xlsx", "");
    const sheet = sheets.items.find(s => s.name === sheetName);

    if (sheet) {
      sheet.delete();
    }
  }

  await context.sync();
});


  } catch (err) {
    if (err.name === 'AbortError') {
      alert("Salvataggio annullato");
    } else {
      throw err;
    }
  }
}