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

/* Pull Admin name from the file name */
function getAdminName(inputFile) {

    const fileName =
        path.basename(inputFile);

    const match =
        fileName.match(/^(.*?)\s*-\s*Weekly/i);

    if (match) {
        return match[1].trim();
    }

    return "UNKNOWN";
}

/* Determine last day in user range */
function getRangeEndingDate(dateRangeString) {

    const months = dateRangeString
        .split(",")
        .map(x => x.trim())
        .filter(Boolean)
        .map(item => {

            const [month, year] =
                item.split("/");

            return {
                month: parseInt(month),
                year: parseInt(year)
            };
        });

    if (!months.length) {
        return "";
    }

    months.sort((a, b) => {

        if (a.year !== b.year) {
            return a.year - b.year;
        }

        return a.month - b.month;
    });

    const last =
        months[months.length - 1];

    const lastDay =
        new Date(
            last.year,
            last.month,
            0
        );

    return (
        `${last.month}/${lastDay.getDate()}/${last.year}`
    );
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

    // Update formula 
    
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
 
  /*
  if (
      sourceCell.value &&
      typeof sourceCell.value === "object"
  ) {

      if (sourceCell.value.formula) {

          targetCell.value = {
              formula: sourceCell.value.formula
          };

          return;
      }

      if (sourceCell.value.sharedFormula) {

          targetCell.value =
              sourceCell.result ?? null;

          return;
      }
  }
  */
}

/* Parse list of supplied months 
  ex: 5/2026,6/2026,7/2026
*/
/*
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
*/

function parseMonthRanges(rangeString) {

    const allowed = new Set();

    rangeString
        .split(",")
        .map(x => x.trim())
        .forEach(item => {

            const [month, year] = item.split("/");

            if (!month || !year) {
                return;
            }

            allowed.add(
                `${parseInt(month)}/${parseInt(year)}`
            );
        });

    return allowed;
}

/*
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
    */

function excelDateToJSDate(value) {

    if (!value) return null;

    if (value instanceof Date) {
        return value;
    }

    if (typeof value === "number") {

        const excelEpoch =
            new Date(Date.UTC(1899, 11, 30));

        return new Date(
            excelEpoch.getTime() +
            value * 86400000
        );
    }

    if (
        typeof value === "object" &&
        value.result instanceof Date
    ) {
        return value.result;
    }

    const parsed = new Date(value);

    return isNaN(parsed)
        ? null
        : parsed;
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

            /*
            const monthKey =
                `${reportDate.getMonth() + 1}-${reportDate.getFullYear()}`;

            const inDateRange =
                allowedMonths.has(monthKey);
            */
            //const monthKey = `${reportDate.getMonth() + 1}/${reportDate.getFullYear()}`;
            const monthKey = `${reportDate.getUTCMonth() + 1}/${reportDate.getUTCFullYear()}`;

            const inDateRange = allowedMonths.has(monthKey);

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

    const adminName = getAdminName(inputFile);

    const rangeEndDate = getRangeEndingDate(dateRangeString);

    /*
    outputSheet.spliceRows(
        1,
        0,
        [`${adminName} A/R to ${rangeEndDate}`]
    );

    const titleRow = outputSheet.getRow(1);

    titleRow.font = {
        bold: true,
        size: 14
    };

    titleRow.alignment = {
        horizontal: "center"
    };
    */

    await outputWorkbook.xlsx.writeFile(
        outputPath
    );

    return outputPath;
}

module.exports = {
  processWorkbook
};