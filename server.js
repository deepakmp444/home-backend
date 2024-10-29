const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const bodyParser = require('body-parser');

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Path to the Excel file
const excelFilePath = path.join(__dirname, 'Pagination_data1.xlsx');

// Helper function to load data from the Excel file
const loadExcelData = () => {
    const workbook = XLSX.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
};

// GET: Fetch all data from the Excel file
app.get('/api/data', (req, res) => {
    try {
        const data = loadExcelData();
        res.json(data);
    } catch (error) {
        res.status(500).json({ message: 'Error reading Excel file', error });
    }
});

// POST: Update the Excel file with new data
app.post('/api/data', (req, res) => {
    const newData = req.body;
    try {
        const worksheet = XLSX.utils.json_to_sheet(newData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        XLSX.writeFile(workbook, excelFilePath);
        res.json({ message: 'Data updated successfully' });
    } catch (error) {
        res.status(500).json({ message: 'Error updating Excel file', error });
    }
});

// Start the server
const PORT = 5003;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
