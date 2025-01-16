document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('uploadForm');
    const fileInput = document.getElementById('formFile');
    const fileNameDisplay = document.getElementById('fileName');
    const resetButton = document.getElementById('resetButton');
  
    // Reset form and custom UI elements
    resetButton.addEventListener('click', function () {
      form.reset(); // Reset the form to its initial state
      fileNameDisplay.textContent = "No file chosen. Upload .xlsx, .xls"; // Reset file name display
    });
  
    // Update file name display on file selection
    fileInput.addEventListener('change', function () {
      if (fileInput.files.length > 0) {
        fileNameDisplay.textContent = fileInput.files[0].name;
      } else {
        fileNameDisplay.textContent = "No file chosen. Upload .xlsx, .xls";
      }
    });
  });