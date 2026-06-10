const ExcelJS = require("exceljs");
const path = require("path");
const os = require("os");

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

function copyCellStyle(sourceCell, targetCell) {

    targetCell.style =
        deepClone(
            sourceCell.style || {}
        );

    if (sourceCell.numFmt) {
        targetCell.numFmt =
            sourceCell.numFmt;
    }

    if (sourceCell.font) {
        targetCell.font =
            deepClone(sourceCell.font);
    }

    if (sourceCell.fill) {
        targetCell.fill =
            deepClone(sourceCell.fill);
    }

    if (sourceCell.border) {
        targetCell.border =
            deepClone(sourceCell.border);
    }

    if (sourceCell.alignment) {
        targetCell.alignment =
            deepClone(
                sourceCell.alignment
            );
    }

    if (sourceCell.protection) {
        targetCell.protection =
            deepClone(
                sourceCell.protection
            );
    }
}

function copyCell(sourceCell, targetCell) {

    copyCellStyle(
        sourceCell,
        targetCell
    );

    const value =
        sourceCell.value;

    // Shared Formula
    if (
        value &&
        typeof value === "object" &&
        value.sharedFormula
    ) {

        targetCell.value =
            value.result ?? 0;

        return;
    }

    // Normal Formula
    if (
        value &&
        typeof value === "object" &&
        typeof value.formula === "string"
    ) {
        /*
        targetCell.value = {
            formula: value.formula
        };
        */
        
        //targetCell.value = value;
        /*
        targetCell.value =
            value.result ?? 0;
        */
       targetCell.value =
        value.result !== undefined &&
        value.result !== null
        ? value.result
        : 0;

        return;
    }

    targetCell.value = value;
}

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

                if (!headerWritten) {
                    const headerRow = newestSheet.getRow(1);
                    const outputHeader = outputSheet.getRow(1);

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

                    headerWritten = true;
                }

                return;
            }

            const dateCell = row.getCell(1).value;
            const reportDate = excelDateToJSDate(dateCell);

            if (!reportDate) {
                return;
            }

            //const monthKey = `${reportDate.getMonth() + 1}/${reportDate.getFullYear()}`;
            // Using UTC time instead of local time to prevent selecting 5/1 contracts if the range is 1,2,3,4 (for example)
            const monthKey = `${reportDate.getUTCMonth() + 1}/${reportDate.getUTCFullYear()}`;

            const inDateRange = allowedMonths.has(monthKey);

            if (!inDateRange) {
                return;
            }

            const yellow = rowContainsYellow(row);

            if (!yellow) {
                return;
            }
            
           const newRow =
                outputSheet.getRow(
                    outputSheet.rowCount + 1
                );

            for (
                let colNumber = 1;
                colNumber <= sheet.columnCount;
                colNumber++
            ) {
                const sourceCell =
                    row.getCell(colNumber);

                const targetCell =
                    newRow.getCell(colNumber);

                // Main Copy Loop
                targetCell.fill = deepClone(sourceCell.fill);
                targetCell.font = deepClone(sourceCell.font);
                targetCell.numFmt = sourceCell.numFmt;

                copyCell(
                    sourceCell,
                    targetCell
                );
                
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

    outputWorkbook.calcProperties = {
        fullCalcOnLoad: true
    };

    await outputWorkbook.xlsx.writeFile(
        outputPath
    );

    return outputPath;
}

module.exports = {
  processWorkbook
};