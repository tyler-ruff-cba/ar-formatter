const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");

const { processWorkbook } = require("./processor/excelProcessor");

function createWindow() {
  const win = new BrowserWindow({
    width: 700,
    height: 500,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true
    }
  });

  win.loadFile("./renderer/index.html");
}

async function selectLocalFile() {
  const focusedWindow = BrowserWindow.getFocusedWindow();
  
  const result = await dialog.showOpenDialog(focusedWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Excel Spreadsheets', extensions: ['xlsx', 'csv', 'xlsm', 'xml'] }]
  });

  if (!result.canceled) {
    const filePath = result.filePaths[0]; // Contains the full absolute path string
    console.log('Selected file:', filePath);
    return filePath;
  }
}

app.whenReady().then(createWindow);

ipcMain.handle("upload-file", async (_) => {
    return await selectLocalFile();
})

ipcMain.handle("process-file", async (_, data) => {
  console.log(data);

  const outputFile = await processWorkbook(
    data.filePath,
    data.dateRanges
  );

  return outputFile;
});

ipcMain.handle("save-output", async (_, sourcePath) => {
  const result = await dialog.showSaveDialog({
    defaultPath: "processed.xlsx"
  });

  if (result.canceled) {
    return null;
  }

  const fs = require("fs");

  fs.copyFileSync(sourcePath, result.filePath);

  return result.filePath;
});