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

document.getElementById('fileUpload').addEventListener('click', uploadFile);
document.getElementById('uploadedFilePath').addEventListener('click', uploadFile);

document
  .getElementById('submitBtn')
  .addEventListener('click', async () => {

    let absolutePath = null;

    const filePath = document.getElementById('uploadedFilePath').value;
    /*
    if (!file) {
      alert("Please select a file.");
      return;
    }
    */

    const dateRanges = document.getElementById('dateRanges').value;

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