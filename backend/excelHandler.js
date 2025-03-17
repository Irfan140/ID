const XLSX = require("xlsx");
const path = require("path");

const filePath = path.join(__dirname, "../data/users.xlsx");

function saveToExcel(data) {
    try {
        let workbook;

        // Check if file exists; if not, create a new workbook
        try {
            workbook = XLSX.readFile(filePath);
        } catch (error) {
            workbook = XLSX.utils.book_new();
        }

        // Check for existing 'Users' worksheet or create one
        let worksheet = workbook.Sheets["Users"];
        if (!worksheet) {
            worksheet = XLSX.utils.aoa_to_sheet([["Name", "Email", "ID Number"]]);
            XLSX.utils.book_append_sheet(workbook, worksheet, "Users");
        }

        // Append data to the worksheet
        const newRow = [data.name, data.email, data.idNumber];
        const existingData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        existingData.push(newRow);

        // Update worksheet with new data
        const updatedWorksheet = XLSX.utils.aoa_to_sheet(existingData);
        workbook.Sheets["Users"] = updatedWorksheet;

        // Save the file
        XLSX.writeFile(workbook, filePath);

    } catch (error) {
        console.error("Error saving to Excel:", error.message);
    }
}

module.exports = { saveToExcel };
