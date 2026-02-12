/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import ExcelJS from 'exceljs';

let fileInput;
const workbooksToSave = []; // Array per workbook: ogni workbook ha nome, buffer di dati, NumericCol, header

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("fileName").textContent = "";
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    fileInput = document.getElementById("fileInput");
    const run = document.getElementById("run");

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

async function fileImport() {
  if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
    alert("Seleziona almeno un file CSV");
    return;
  }

  workbooksToSave.length = 0; // Reset array dei workbook 
  
  try {
    for (let i = 0; i < fileInput.files.length; i++) {
      const file = fileInput.files[i];
      const fileName = file.name;
      
      await processCSVFile(file, fileName);
    }
    
    if (workbooksToSave.length > 0) {
      document.getElementById("fileName").textContent = "";
      await saveAllWorkbooks();
    }    
  } catch (err) {
    console.error("Errore durante l'importazione:", err);
  }
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

        const csvData = rows.map((row) => {
          const rowForSplitting = row.replace(/"(\d+),(\d+)"/g, "$1.$2"); //formattazione campo decimale come numero.numero, semplifica lo split
          const separator = rowForSplitting.includes(";") ? ";" : ","; //split righe in celle
          return rowForSplitting.split(separator).map((cell) => processCell(cell)); //processa celle
        });
        
        await createExcelWorkbook(csvData, fileName); //i dati formattati e il nome sono usati per la creazione del nuovo workbook da salvare
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

  if (numericStr.includes('.') && /\d+\.\d+/.test(numericStr)) {
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
      if (typeof value === "number" && !Number.isInteger(value)) {
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
  
  const excelFileName = sheetName + ".xlsx";
  
  // Salva con metadata delle colonne numeriche
  workbooksToSave.push({
    name: excelFileName,
    buffer: buffer,
    numericCols: numericCols,
    headers: csvData[0]
  });
  
  console.log(`Workbook creato: ${excelFileName}`);
}

// Salva tutti i workbook nella cartella selezionata, NELLA STESSA, SE NE SELEZIONA SOLO 1
async function saveAllWorkbooks() {
  try {
    const dir = await window.showDirectoryPicker();
    
    for (const workbookData of workbooksToSave) {
      const { name, buffer, numericCols, headers } = workbookData;
      
      const fileHandle = await dir.getFileHandle(name, { create: true }); //per creare nuovo file
      const writable = await fileHandle.createWritable();
      
      await writable.write({
        type: 'write',
        data: buffer
      });
      
      await writable.close();
      console.log(`Salvato: ${name}`);
      document.getElementById("fileName").textContent += `${name} SALVATO CON SUCCESSO,   `;

    }    
  } catch (err) {
    if (err.name === 'AbortError') {
      console.log("Salvataggio annullato dall'utente");
      alert("Salvataggio annullato");
    } else {
      throw err;
    }
  }
}

