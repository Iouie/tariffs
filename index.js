const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// List of numbers to highlight (as strings for exact matching)
const numbersToHighlight = ["42", "123", "7"]; // Add your numbers here

// Function to highlight numbers in an Excel file
async function highlightNumbers(inputFilePath, outputFilePath) {
  const workbook = new ExcelJS.Workbook();

  try {
    // Load the input Excel file
    await workbook.xlsx.readFile(inputFilePath);

    // Iterate through each worksheet
    workbook.eachSheet((worksheet) => {
      // Iterate through each row
      worksheet.eachRow((row) => {
        // Iterate through each cell in the row
        row.eachCell((cell) => {
          // Convert cell value to a string for consistent comparison
          const cellValue = cell.value?.toString() || "";

          // Log the cell value for debugging
          console.log(
            `Checking cell at row ${row.number}, column ${cell.col}: ${cellValue}`
          );

          // Check if the cell value exactly matches any of the numbers to highlight
          if (numbersToHighlight.includes(cellValue)) {
            // Log the highlighted cell for debugging
            console.log(`Highlighted: ${cellValue}`);

            // Apply highlighting (background color)
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFADD8E6" }, // Blue background
            };
            cell.value = "9908125";
            cell.font = {
              name: "Arial",
              size: 14,
              bold: true,
              color: { argb: "00000000" }, // black text
            };
          }
        });
      });
    });

    // Save the modified workbook to a new file
    await workbook.xlsx.writeFile(outputFilePath);
    console.log(`Highlighted file saved to: ${outputFilePath}`);
  } catch (error) {
    console.error("Error processing the Excel file:", error);
  }
}

// Handle file drag-and-drop
function handleFileDrop(filePath) {
  const inputFilePath = filePath;
  const outputFilePath = path.join(
    path.dirname(inputFilePath),
    `highlighted_${path.basename(inputFilePath)}`
  );

  highlightNumbers(inputFilePath, outputFilePath);
}

// Check if a file path is provided as a command-line argument
if (process.argv[2]) {
  handleFileDrop(process.argv[2]);
} else {
  console.log("Drag and drop an Excel file onto this executable.");
}
