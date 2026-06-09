const ExcelJS = require("exceljs");
const path = require("path");
const os = require("os");

const rowMap = {};

//const testFile = "C:\Users\Alexa\OneDrive\Desktop\Documents IT\GitHub\ar-formatter\test\HEADSTART - Weekly 7.xlsx";

/* Deep Clone Helper
    Helps copy style and formatting from the source document
*/
function deepClone(obj) {
    return obj ? JSON.parse(JSON.stringify(obj)) : obj;
}

/* Copy a cell exactly
function copyCell(sourceCell, targetCell) {

    targetCell.value = sourceCell.value;

    if (sourceCell.style) {
        targetCell.style = deepClone(sourceCell.style);
    }

    if (sourceCell.numFmt) {
        targetCell.numFmt = sourceCell.numFmt;
    }

    if (sourceCell.font) {
        targetCell.font = deepClone(sourceCell.font);
    }

    if (sourceCell.fill) {
        targetCell.fill = deepClone(sourceCell.fill);
    }

    if (sourceCell.border) {
        targetCell.border = deepClone(sourceCell.border);
    }

    if (sourceCell.alignment) {
        targetCell.alignment =
            deepClone(sourceCell.alignment);
    }

    if (sourceCell.protection) {
        targetCell.protection =
            deepClone(sourceCell.protection);
    }

    if (
        sourceCell.value &&
        sourceCell.value.formula
    ) {

    let formula =
          sourceCell.value.formula;

      formula =
          formula.replace(
              /\d+/g,
              match => {

                  const oldRow =
                      parseInt(match);

                  return rowMap[oldRow];
              }
          );

      targetCell.value = {
          formula
      };
  }
}
*/

function copyCell(sourceCell, targetCell) {

    targetCell.value = sourceCell.value;

    targetCell.style =
        JSON.parse(
            JSON.stringify(
                sourceCell.style || {}
            )
        );

    if (sourceCell.numFmt) {
        targetCell.numFmt =
            sourceCell.numFmt;
    }

    if (
        sourceCell.value &&
        sourceCell.value.formula
    ) {

    let formula =
          sourceCell.value.formula;

      formula =
          formula.replace(
              /\d+/g,
              match => {

                  const oldRow =
                      parseInt(match);

                  return rowMap[oldRow];
              }
          );

      targetCell.value = {
          formula
      };
  }
}

/* Parse list of supplied months 
  ex: 5/2026,6/2026,7/2026
*/
function parseMonthRanges(rangeString) {
    const allowed = new Set();

    rangeString
        .split(",")
        .map(x => x.trim())
        .forEach(item => {
            const [month, year] = item.split("/");

            if (month && year) {
                allowed.add(`${parseInt(month)}-${parseInt(year)}`);
            }
        });

    return allowed;
}

function excelDateToJSDate(value) {

    if (!value) return null;

    if (value instanceof Date) {
        return value;
    }

    if (typeof value === "object" && value.result instanceof Date) {
        return value.result;
    }

    const parsed = new Date(value);

    if (!isNaN(parsed)) {
        return parsed;
    }

    return null;
}

function isYellowColor(argb) {

    if (!argb || typeof argb !== "string") {
        return false;
    }

    argb = argb.replace(/^FF/i, "");

    if (argb.length !== 6) {
        return false;
    }

    const r = parseInt(argb.substring(0, 2), 16);
    const g = parseInt(argb.substring(2, 4), 16);
    const b = parseInt(argb.substring(4, 6), 16);

    /*
        Yellow detection logic:

        - Red high
        - Green high
        - Blue significantly lower

        Handles:
            FFFF00
            FFF200
            FFEA00
            FFD966
            etc.
    */

    return (
        r >= 180 &&
        g >= 180 &&
        b <= 150 &&
        Math.abs(r - g) <= 80
    );
}

function rowContainsYellow(row) {

    let coloredCells = 0;

    row.eachCell(cell => {

        const fill = cell.fill;

        if (!fill || !fill.fgColor) {
            return;
        }

        const rgb = fill.fgColor.argb;

        if (isYellowColor(rgb)) {
            coloredCells++;
        }
    });

    /*
        Prevent accidental matches from
        a single highlighted cell.
    */

    return coloredCells >= 3;
}


/*
async function processWorkbook(
  inputFile,
  dateRangeString
) {

  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile(inputFile);

  const sheet = workbook.getWorksheet(1);

  const outputWorkbook = new ExcelJS.Workbook();

  const outputSheet =
    outputWorkbook.addWorksheet("Filtered");

  // Parse:
  // 5/2026,6/2026,7/2026

  const allowedMonths =
    dateRangeString
      .split(",")
      .map(x => x.trim());

  // Copy headers

  const header = sheet.getRow(1);

  outputSheet.addRow(header.values);

  // Example:
  // Assume column A contains dates

  sheet.eachRow((row, rowNumber) => {

    if (rowNumber === 1) {
      return;
    }

    const dateValue = row.getCell(1).value;

    if (!dateValue) {
      return;
    }

    const date = new Date(dateValue);

    const monthYear =
      `${date.getMonth() + 1}/${date.getFullYear()}`;

    if (allowedMonths.includes(monthYear)) {
      outputSheet.addRow(row.values);
    }
  });

  const outputPath = path.join(
    os.tmpdir(),
    `processed-${Date.now()}.xlsx`
  );

  await outputWorkbook.xlsx.writeFile(outputPath);

  return outputPath;
}
*/

async function processWorkbook(inputFile, dateRangeString) {

    console.log("Starting processing...");

    const allowedMonths =
        parseMonthRanges(dateRangeString);

    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(inputFile);

    const outputWorkbook =
        new ExcelJS.Workbook();

    const outputSheet =
        outputWorkbook.addWorksheet("Filtered Results");

    let headerWritten = false;

    // Get the most recent sheet for the header
    let newestSheet = null;
    let newestYear = 0;

    workbook.eachSheet(sheet => {

        const match = sheet.name.match(/\b(20\d{2})\b/);

        if (!match) {
            return;
        }

        const year = parseInt(match[1]);

        if (year > newestYear) {
            newestYear = year;
            newestSheet = sheet;
        }
    });

    workbook.eachSheet(sheet => {

        if (!/\b20\d{2}\b/.test(sheet.name)) {
            return;
        }

        sheet.eachRow((row, rowNumber) => {

            if (rowNumber === 1) {
                /*
                if (!headerWritten) {
                    outputSheet.addRow(
                        row.values.slice(1)
                    );
                    headerWritten = true;
                }
                */

                if (!headerWritten) {
                    /*
                    const outputHeader =
                        outputSheet.getRow(1);

                    row.eachCell((cell, colNumber) => {

                        const targetCell =
                            outputHeader.getCell(colNumber);

                        copyCell(cell, targetCell);
                    });
                    */

                    //const outputHeader = newestSheet.getRow(1);
                    const headerRow = newestSheet.getRow(1);
                    const outputHeader = outputSheet.getRow(1);
                    /*
                    for (
                        let colNumber = 1;
                        colNumber <= newestSheet.columnCount;
                        colNumber++
                    ) {
                        const sourceCell =
                            row.getCell(colNumber);

                        const targetCell =
                            outputHeader.getCell(colNumber);

                        copyCell(sourceCell, targetCell);
                    }

                    outputHeader.height = row.height;
                    sheet.columns.forEach((column, index) => {
                        outputSheet.getColumn(index + 1).width =
                            column.width;
                    });
                    */

                    for (
                        let col = 1;
                        col <= newestSheet.columnCount;
                        col++
                    ) {

                        const sourceCell =
                            headerRow.getCell(col);

                        const targetCell =
                            outputHeader.getCell(col);

                        copyCell(sourceCell, targetCell);
                    }
                    outputHeader.height = headerRow.height;

                    newestSheet.columns.forEach(
                        (column, index) => {

                            outputSheet.getColumn(index + 1).width =
                                column.width;
                        }
                    );

                    rowMap[1] = 1;
                    headerWritten = true;
                }

                return;
            }

            const dateCell =
                row.getCell(1).value;

            const reportDate =
                excelDateToJSDate(dateCell);

            if (!reportDate) {
                return;
            }

            const monthKey =
                `${reportDate.getMonth() + 1}-${reportDate.getFullYear()}`;

            const inDateRange =
                allowedMonths.has(monthKey);

            if (!inDateRange) {
                return;
            }

            const yellow =
                rowContainsYellow(row);

            if (!yellow) {
                return;
            }
            
            /*
            outputSheet.addRow(
                row.values.slice(1)
            );
            */
           const newRow =
                outputSheet.getRow(
                    outputSheet.rowCount + 1
                );

            rowMap[row.number] = newRow.number;

            /*
            row.eachCell((cell, colNumber) => {

                const targetCell = newRow.getCell(colNumber);
                targetCell.fill = deepClone(cell.fill);
                targetCell.font = deepClone(cell.font);
                targetCell.numFmt = cell.numFmt;

                copyCell(cell, targetCell);
            });
            */

            for (
                let colNumber = 1;
                colNumber <= sheet.columnCount;
                colNumber++
            ) {
                const sourceCell =
                    row.getCell(colNumber);

                const targetCell =
                    newRow.getCell(colNumber);

                targetCell.fill = deepClone(sourceCell.fill);
                targetCell.font = deepClone(sourceCell.font);
                targetCell.numFmt = sourceCell.numFmt;

                copyCell(sourceCell, targetCell);
            }

            newRow.height = row.height;
        });
    });

    const outputPath = path.join(
        os.tmpdir(),
        `filtered-${Date.now()}.xlsx`
    );

    await outputWorkbook.xlsx.writeFile(
        outputPath
    );

    return outputPath;
}

module.exports = {
  processWorkbook
};