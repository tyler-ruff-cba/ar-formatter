let generatedFile = null;

async function uploadFile(){
  const file = await window.electronAPI.uploadFile();
  if(file === undefined){
      document.getElementById('uploadedFilePath').value = "Please select a file...";
      return;
  }
  console.log(`File uploaded:${file}`);
  document.getElementById('uploadedFilePath').value = file;
}

function isFilePath(value) {
  if (typeof value !== "string" || !value.trim()) {
    return false;
  }

  const str = value.trim();

  // Windows absolute paths (C:\folder\file.txt, C:/folder/file.txt)
  const windowsAbsolute = /^[a-zA-Z]:[\\/]/;
  // UNC paths (\\server\share\file.txt)
  const uncPath = /^\\\\[^\\]+\\[^\\]+/;
  // Unix absolute paths (/home/user/file.txt, /etc/hosts)
  const unixAbsolute = /^\/(?:[^/]+\/?)*$/;
  // Relative paths containing directory separators
  const relativePath = /^(?:\.{1,2}[\\/]|[^<>:"|?*\r\n]+[\\/])/;
  // Paths ending with a filename extension
  const fileExtension = /(?:^|[\\/])[^\\/]+\.[a-zA-Z0-9_-]+$/;

  return (
    windowsAbsolute.test(str) ||
    uncPath.test(str) ||
    unixAbsolute.test(str) ||
    relativePath.test(str) ||
    fileExtension.test(str)
  );
}

document.getElementById('fileUpload').addEventListener('click', uploadFile);
document.getElementById('uploadedFilePath').addEventListener('click', uploadFile);

document
  .getElementById('submitBtn')
  .addEventListener('click', async () => {

    const filePath = document.getElementById('uploadedFilePath').value;

    if (filePath === undefined || filePath === null || filePath === "" || !isFilePath(filePath)) {
        alert("Please select a valid file.");
        return;
    }

    const dateRanges = document.getElementById('dateRanges').value;

    if(dateRanges === undefined || dateRanges === null || dateRanges.length <= 0){
        alert("Please enter a date range.");
        return;
    }

    document.getElementById('status').innerText = "Processing...";

    generatedFile =
      await window.electronAPI.processFile({
        filePath: filePath,
        dateRanges
      });

    document.getElementById('status').innerText = "Finished.";

    document.getElementById('downloadBtn').style.display = "inline-block";
  });

document
  .getElementById('downloadBtn')
  .addEventListener('click', async () => {
    await window.electronAPI.saveOutput(
      generatedFile
    );
  });