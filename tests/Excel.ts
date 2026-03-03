const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

function readExcel(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    let data = {};

    workbook.SheetNames.forEach(SheetName => {
      let sheet = workbook.Sheets[SheetName];
      data[SheetName] = xlsx.utils.sheet_to_json(sheet);
    });

    return data;
  } catch (error) {
    console.error('Error reading Excel file:', error);
    return {};
  }
}

function writeExcel(filePath, exceljson) {
  const newExcel = xlsx.utils.book_new();
  Object.keys(exceljson).forEach(sheetName => {
        const worksheet = xlsx.utils.json_to_sheet(exceljson[sheetName]);
        xlsx.utils.book_append_sheet(newExcel, worksheet, sheetName);
  });

  xlsx.writeFile(newExcel, filePath);
}

module.exports = {readExcel, writeExcel};