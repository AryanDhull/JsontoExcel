const fs = require('fs');
const XLSX = require('xlsx');

/**
 * Recursively processes JSON data and creates sheets accordingly.
 * @param {object} jsonData - The JSON data to process.
 * @param {object} sheets - The object to store processed sheets.
 * @param {string} prefix - The prefix for sheet names.
 */
function processJSON(jsonData, sheets, prefix = '') {
  if (!jsonData || typeof jsonData !== 'object') {
    return;
  }

  const currentSheet = {};

  Object.keys(jsonData).forEach((key) => {
    const sheetData = jsonData[key];
    const sheetName = prefix + key;

    if (Array.isArray(sheetData) && sheetData.length > 0 && typeof sheetData[0] === 'object') {
      // Treat arrays of objects as separate sheets
      sheets[sheetName] = [];
      sheetData.forEach((item, index) => {
        const newSheetName = `${sheetName}_${index + 1}`;
        processJSON(item, sheets, newSheetName + '_');
        // Add processed data to the current sheet
        sheets[sheetName].push(...sheets[newSheetName]);
        delete sheets[newSheetName];
      });
    } else if (typeof sheetData === 'object') {
      // Treat nested objects as separate sheets
      processJSON(sheetData, sheets, sheetName + '_');
    } else {
      // Keep other variables in the same sheet
      currentSheet[key] = sheetData;
    }
  });

  if (Object.keys(currentSheet).length > 0) {
    // No hyperlink, just display the sheet name
    sheets[prefix.slice(0, -1)] = [
      {
        'Sheet Data': `SHEET::${prefix.slice(0, -1)}`,
        ...currentSheet,
      },
    ];
  }
}

/**
 * Writes processed sheets to an Excel file.
 * @param {object} sheets - The object containing processed sheets.
 * @param {string} outputFileName - The name of the output Excel file.
 */
function writeToExcel(sheets, outputFileName) {
  const wb = XLSX.utils.book_new();

  Object.keys(sheets).forEach((sheetName) => {
    const ws = XLSX.utils.json_to_sheet(sheets[sheetName]);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  XLSX.writeFile(wb, outputFileName);
}

/**
 * Reads JSON from a file, processes it, and writes to an Excel file.
 * @param {string} filePath - The path to the input JSON file.
 * @param {string} outputFileName - The name of the output Excel file.
 */
function readAndProcessJSON(filePath, outputFileName) {
  try {
    const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf8'));

    if (!jsonData) {
      throw new Error('JSON data is null or undefined.');
    }

    const sheets = {};
    processJSON(jsonData, sheets);
    writeToExcel(sheets, outputFileName);
    console.log(`Data successfully written to ${outputFileName}`);
  } catch (error) {
    console.error('Error:', error.message);
  }
}

// Replace 'input.json' and 'output_file.xlsx' with your file names
readAndProcessJSON('input.json', 'output_file.xlsx');
