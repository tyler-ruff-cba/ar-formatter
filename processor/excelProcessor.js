const ExcelJS = require("exceljs");
const path = require("path");
const os = require("os");

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

module.exports = {
  processWorkbook
};