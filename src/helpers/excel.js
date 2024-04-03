import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';
import fs from 'fs';
import path from 'path';

import { getPreviousMonthAndYear } from './dates.js';

export async function exportToExcel(data, subscriptionName) {
  const workbook = new ExcelJS.Workbook();
  const excelFile = findExcelFilePath();

  if (excelFile) {
    await workbook.xlsx.readFile(excelFile);
  }
  //Cleaning Subscription Name and Creating Worksheet
  const cleanedSubscriptionName = subscriptionName.replace(/[\\*?:/[\]]/g, '');
  const worksheet = workbook.addWorksheet(cleanedSubscriptionName);
  console.log(`Exporting data for ${subscriptionName}...`);

  // Adding Title and Headings
  worksheet.addRow([]);
  const { month, year } = getPreviousMonthAndYear();
  const title = `Informe de costos: ${subscriptionName} ${month} ${year}`;
  addTitleToWorksheet(worksheet, title);
  worksheet.addRow([]);
  addHeadersToWorksheet(worksheet);

  //Data Processing and Subtotal Calculation
  const totalRowData = data.pop(); // Extract the last row data for 'TOTAL'
  data.sort((a, b) => a.Etiqueta.localeCompare(b.Etiqueta)); // Sort the data by 'Etiqueta'
  addDataAndCalculateSubtotals(worksheet, data);

  //Formatting and Addition of the 'TOTAL' Row
  const totalRow = worksheet.addRow([`TOTAL SUSCRIPCIÓN: ${subscriptionName}`, '', '', '', '', totalRowData.Costo]);
  formatTotalRow(worksheet, totalRow);

  //Adjusting Column Widths and Saving the File
  adjustColumnWidth(worksheet);
  const fileName = excelFile || `Costos-por-recursos-Suscripción--periodo-${month}-${year}.xlsx`;
  await workbook.xlsx.writeFile(fileName);
  console.log(`Data exported ${subscriptionName} to ${fileName}`);
}

function addTitleToWorksheet(worksheet, title) {
  // Add the title row and merge cells from A2 to F2
  const titleRow = worksheet.addRow([title]);
  worksheet.mergeCells('A2:F2');

  // Style for the title row
  titleRow.font = { size: 24, bold: true, color: { argb: 'FFFFFFFF' } };
  titleRow.alignment = { vertical: 'middle', horizontal: 'center' };

  // Apply background color to each cell in the title row
  titleRow.eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF16365C' },
    };
  });
}

function addHeadersToWorksheet(worksheet) {
  const headerStyle = {
    font: { bold: true, color: { argb: 'FFFFFFFF' } },
    alignment: { horizontal: 'center' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF366090' } },
    border: { bottom: { style: 'medium', color: { argb: '000000' } } },
  };

  // Column titles
  const headers = ['Suscripción', 'Grupo de recursos', 'Etiqueta(s)', 'Tipo de recurso', 'Nombre de recurso', 'Costo'];
  const headerRow = worksheet.addRow(headers);

  // Apply the style to each header cell
  headerRow.eachCell(cell => {
    cell.font = headerStyle.font;
    cell.alignment = headerStyle.alignment;
    cell.fill = headerStyle.fill;
    cell.border = headerStyle.border;
  });
}

function addDataAndCalculateSubtotals(worksheet, data) {
  // Define border style for data cells
  const dataBorderStyle = {
    border: {
      top: { style: 'thin', color: { argb: '000000' } },
      left: { style: 'thin', color: { argb: '000000' } },
      bottom: { style: 'thin', color: { argb: '000000' } },
      right: { style: 'thin', color: { argb: '000000' } },
    },
  };

  let currentTag = null;
  let subtotal = 0;

  // Iterate over each data item
  data.forEach(item => {
    // Check for a change in tag and add a subtotal row if needed
    if (currentTag && item.Etiqueta !== currentTag) {
      addSubtotalRow(worksheet, currentTag, subtotal, dataBorderStyle);
      subtotal = 0; // Reset subtotal for the new tag
    }

    // Update the current tag and accumulate cost for the subtotal
    currentTag = item.Etiqueta;
    subtotal += item.Costo;

    // Add the data row and apply styles
    const row = worksheet.addRow([
      item.Suscripción,
      item['Grupo de recursos'],
      item.Etiqueta,
      item['Tipo de recurso'],
      item.Recurso,
      item.Costo,
    ]);

    // Apply border style and currency format to each cell
    row.eachCell((cell, colNumber) => {
      cell.border = dataBorderStyle.border;
      if (colNumber === 6) {
        cell.numFmt = '"$"#,##0.00'; // Currency format for cost column
      }
    });
  });

  // Add a final subtotal row for the last tag
  if (currentTag) {
    addSubtotalRow(worksheet, currentTag, subtotal, dataBorderStyle);
  }
}

function addSubtotalRow(worksheet, tag, subtotal, borderStyle) {
  // Insert a row for the subtotal
  const subtotalRow = worksheet.addRow([`Subtotal: ${tag}`, '', '', '', '', subtotal]);

  // Combine cells 'A' to 'E' for subtotal
  const startCol = 1; // Column A
  const endCol = 5; // Column E
  const rowNumber = subtotalRow.number;
  worksheet.mergeCells(rowNumber, startCol, rowNumber, endCol);

  subtotalRow.eachCell((cell, colNumber) => {
    // Applies the background color and bold to all cells in the subtotal row
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFC5D9F1' },
    };
    cell.font = { bold: true };

    // Applies the border and currency formatting to the cost column
    cell.border = borderStyle.border;
    if (colNumber === 6) {
      cell.numFmt = '"$"#,##0.00';
    }

    // Center alignment for merged cells
    if (colNumber >= 1 && colNumber <= 5) {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  });
}

function formatTotalRow(worksheet, totalRow) {
  // Combine cells 'A' to 'E' for the 'TOTAL' row
  const totalRowNumber = totalRow.number;
  worksheet.mergeCells(totalRowNumber, 1, totalRowNumber, 5);

  // Apply styles to the 'TOTAL' row
  totalRow.eachCell((cell, colNumber) => {
    // Set fill, font, and number format for all cells in the row
    if (colNumber >= 1 && colNumber <= 6) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF366092' } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      if (colNumber === 6) {
        cell.numFmt = '"$"#,##0.00'; // Currency format for the cost column
      }
    }

    // Set the top border style for all cells and align the merged cells
    cell.border = { top: { style: 'medium', color: { argb: '000000' } }, bottom: {}, right: {} };
    if (colNumber >= 1 && colNumber <= 5) {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  });
}

function adjustColumnWidth(worksheet) {
  // Adjust column widths based on content
  worksheet.columns.forEach(column => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, cell => {
      // Consider only cells that are not part of a subtotal
      if (!cell.isMerged && cell.value) {
        let cellLength = cell.value.toString().length;
        if (cellLength > maxLength) {
          maxLength = cellLength;
        }
      }
    });
    // Set the column width based on the longest value, with a small extra margin
    column.width = maxLength + 1;
  });
}

export async function updateExcelWithChangesIfExists(deletedElements, newElements) {
  const workbook = new ExcelJS.Workbook();
  const excelFilePath = findExcelFilePath();

  try {
    await workbook.xlsx.readFile(excelFilePath);

    const addSheetWithData = (workbook, sheetName, data) => {
      const worksheet = workbook.addWorksheet(sheetName);
      worksheet.columns = [
        { header: 'Suscripción', key: 'Suscripción', width: 30 },
        { header: 'Grupo de recursos', key: 'Grupo de recursos', width: 30 },
        { header: 'Etiqueta(s)', key: 'Etiqueta', width: 30 },
        { header: 'Tipo de recurso', key: 'Tipo de recurso', width: 20 },
        { header: 'Nombre de recurso', key: 'Recurso', width: 30 },
        { header: 'Costo', key: 'Costo', width: 15, style: { numFmt: '"$"#,##0.00;[Red]"-$"#,##0.00' } },
      ];
      worksheet.getRow(1).eachCell(cell => {
        cell.alignment = { horizontal: 'center' };
        cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF366090' } };
        cell.border = {
          bottom: { style: 'medium', color: { argb: 'FF000000' } },
        };
      });

      // Add Rows
      data.forEach(item => {
        const row = worksheet.addRow(item);
        row.eachCell(cell => {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
          };
        });
      });
      adjustColumnWidth(worksheet);
    };

    // Add new sheets with the changes
    if (deletedElements.length > 0) {
      addSheetWithData(workbook, 'Gastos eliminados', deletedElements);
    }

    if (newElements.length > 0) {
      addSheetWithData(workbook, 'Gastos nuevos', newElements);
    }

    // Save the changes
    await workbook.xlsx
      .writeFile(excelFilePath)
      .then(() => console.log('Archivo actualizado con éxito.'))
      .catch(error => console.error('Error al guardar el archivo:', error));
  } catch (error) {
    console.error('El archivo no se encontró, no se realizaron cambios.');
  }
}

export const findExcelFilePath = () => {
  // Get the current file path
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = path.dirname(__filename);

  // Navigate up two levels to reach the project's root directory
  const dirPath = path.join(__dirname, '..', '..');

  // Get the latest .xlsx file in the same directory
  const files = fs.readdirSync(dirPath);

  let latestFile;
  let latestTime = 0;

  files.forEach(file => {
    if (path.extname(file) === '.xlsx') {
      const stats = fs.statSync(path.join(dirPath, file));
      if (stats.mtimeMs > latestTime) {
        latestTime = stats.mtimeMs;
        latestFile = file;
      }
    }
  });

  return latestFile ? path.join(dirPath, latestFile) : null;
};

export const deleteExcelFile = filePath => {
  if (filePath) {
    try {
      fs.unlinkSync(filePath);
      console.log(`File deleted: ${filePath}`);
    } catch (error) {
      console.error(`Error deleting file: ${error.message}`);
    }
  } else {
    console.log('No Excel file found to delete.');
  }
};
