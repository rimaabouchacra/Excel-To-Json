const XLSX = require("xlsx");
const fs = require("fs");

try {
  // Load Excel file
  const workbook = XLSX.readFile("./T_All.xlsx"); // Replace with your file path

  // Choose a specific sheet
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Convert sheet to JSON
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  if (jsonData.length === 0) {
    throw new Error("No data found in the Excel sheet.");
  }

  // Get the header row to map column letters to names
  const headerRow = jsonData[0];
  const columnNames = headerRow.map((name) => name.trim());

  // Remove the header row from the data
  jsonData.shift();

  // Group data by ID
  const groupedData = jsonData.reduce((acc, row) => {
    const id = row[columnNames.indexOf("cr_rf")];
    const entry = {};

    columnNames.forEach((columnName, index) => {
      if (columnName === "cdate") {
        // Convert Excel date serial number to a formatted date string
        const dateValue = excelDateToJSDate(row[index]);
        entry[columnName] = dateValue
          ? dateValue.toISOString().split("T")[0]
          : "Invalid date";
      } else {
        entry[columnName] = row[index];
      }
    });

    if (!acc[id]) {
      acc[id] = [];
    }

    acc[id].push(entry);
    return acc;
  }, {});

  // Function to convert Excel date serial number to JavaScript Date object
  function excelDateToJSDate(serial) {
    if (serial > 0) {
      // Adjust for Excel's base date (January 1, 1900)
      const baseDate = new Date(Date.UTC(1900, 0, 0));
      return new Date(baseDate.getTime() + (serial - 1) * 24 * 60 * 60 * 1000);
    }
    return null; // Return null for invalid date values
  }

  // Save JSON to a file
  fs.writeFileSync("output.json", JSON.stringify(groupedData, null, 2));

  console.log("Conversion complete. Check output.json");
} catch (error) {
  console.error("Error:", error.message);
}


