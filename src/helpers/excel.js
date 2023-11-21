import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';
import fs from 'fs';
import path from 'path';

import { getPreviousMonthAndYear } from './dates.js';

export async function exportToExcel(data, subscriptionName) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Datos');

  // Header style with RGB background color and white text
  const headerStyle = {
    font: { bold: true, color: { argb: 'FFFFFFFF' } },
    alignment: { horizontal: 'center' },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF366090' } },
    border: { bottom: { style: 'medium', color: { argb: '000000' } } },
  };

  const { month, year } = getPreviousMonthAndYear();
  const title = `Informe de costos: Suscripción ${subscriptionName} ${month} ${year}`;

  // Add a blank row to the beginning
  worksheet.addRow([]);

  // Add title row
  const titleRow = worksheet.addRow([title]);
  worksheet.mergeCells('A2:F2');
  titleRow.font = { size: 24, bold: true, color: { argb: 'FFFFFFFF' } };
  titleRow.alignment = { vertical: 'middle', horizontal: 'center' };

  titleRow.eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF16365C' },
    };
  });

  // Agrega otra fila en blanco después del título
  worksheet.addRow([]);

  // Border style for data cells
  const dataBorderStyle = {
    border: {
      top: { style: 'thin', color: { argb: '000000' } },
      left: { style: 'thin', color: { argb: '000000' } },
      bottom: { style: 'thin', color: { argb: '000000' } },
      right: { style: 'thin', color: { argb: '000000' } },
    },
  };

  // Add headers
  const headerRow = worksheet.addRow([
    'Suscripción',
    'Grupo de recursos',
    'Etiqueta(s)',
    'Tipo de recurso',
    'Nombre de recurso',
    'Costo',
  ]);
  headerRow.eachCell(cell => {
    cell.font = headerStyle.font;
    cell.alignment = headerStyle.alignment;
    cell.fill = headerStyle.fill;
    cell.border = headerStyle.border;
  });

  // Extract the last row data for 'TOTAL'
  const totalRowData = data.pop();

  // Sort the data by 'Etiqueta'
  data.sort((a, b) => a.Etiqueta.localeCompare(b.Etiqueta));

  let currentTag = null;
  let subtotal = 0;

  // Add data, apply styles, and calculate subtotals
  data.forEach(item => {
    // When the tag changes, insert a subtotal row
    if (currentTag && item.Etiqueta !== currentTag) {
      addSubtotalRow(worksheet, currentTag, subtotal, dataBorderStyle);
      subtotal = 0; // Reset subtotal for the next tag group
    }

    // Set the current tag for comparison in the next iteration
    currentTag = item.Etiqueta;
    subtotal += item.Costo;

    const row = worksheet.addRow([
      item.Suscripción,
      item['Grupo de recursos'],
      item.Etiqueta,
      item['Tipo de recurso'],
      item.Recurso,
      item.Costo,
    ]);

    row.eachCell((cell, colNumber) => {
      cell.border = dataBorderStyle.border;
      if (colNumber === 6) cell.numFmt = '"$"#,##0.00'; // Apply currency format
    });
  });

  // Insert the last subtotal row if any
  if (currentTag) {
    addSubtotalRow(worksheet, currentTag, subtotal, dataBorderStyle);
  }

  // Add the 'TOTAL' row at the end after all subtotals
  const totalRow = worksheet.addRow([`TOTAL SUSCRIPCIÓN: ${subscriptionName}`, '', '', '', '', totalRowData.Costo]);

  // Combine cells 'A' to 'E' for row 'TOTAL'
  const totalRowNumber = totalRow.number;
  worksheet.mergeCells(totalRowNumber, 1, totalRowNumber, 5); // De 'A' a 'E'

  // Applies styles to row 'TOTAL'
  totalRow.eachCell((cell, colNumber) => {
    if (colNumber >= 1 && colNumber <= 6) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF366092' },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      if (colNumber === 6) {
        cell.numFmt = '"$"#,##0.00';
      }
    }

    cell.border = {
      top: { style: 'medium', color: { argb: '000000' } },
      bottom: {},
      right: {},
    };

    if (colNumber >= 1 && colNumber <= 5) {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  });

  // Adjust column width
  adjustColumnWidth(worksheet);

  // Save the workbook to a file
  const fileName = `Costos-por-recursos-Suscripción-${subscriptionName}-periodo-${month}-${year}.xlsx`;
  await workbook.xlsx.writeFile(fileName);
}

function addSubtotalRow(worksheet, tag, subtotal, borderStyle) {
  // Inserta una fila para el subtotal
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

export const findLatestExcelFile = () => {
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
