const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

app.post('/merge', upload.array('files', 10), (req, res) => {
    const files = req.files;
    if (!files || files.length < 2) {
        return res.status(400).send('Please upload at least two Excel files.');
    }

    let workbooks = files.map(file => xlsx.readFile(file.path));
    let commonColumn = findCommonColumn(workbooks);
    let mergedData = mergeWorkbooks(workbooks, commonColumn);

    let newWorkbook = xlsx.utils.book_new();
    let newWorksheet = xlsx.utils.json_to_sheet(mergedData);
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'MergedData');

    const filePath = path.join(__dirname, 'uploads', 'merged_output.xlsx');
    xlsx.writeFile(newWorkbook, filePath);

    files.forEach(file => fs.unlinkSync(file.path));

    
    res.download(filePath, 'merged_output.xlsx', err => {
        if (err) {
            console.log('Error while sending file:', err);
        }
        fs.unlinkSync(filePath);
    });
});

//  function to find the common column across all files
function findCommonColumn(workbooks) {
    let columnsList = workbooks.map(workbook => {
        let firstSheetName = workbook.SheetNames[0];
        let firstSheet = workbook.Sheets[firstSheetName];
        let headers = xlsx.utils.sheet_to_json(firstSheet, { header: 1 })[0];
        return headers;
    });

    let commonColumns = columnsList.reduce((common, headers) => {
        return common.filter(col => headers.includes(col));
    });

    return commonColumns.length > 0 ? commonColumns[0] : null;
}

// function to merge data based on the common column
function mergeWorkbooks(workbooks, commonColumn) {
    let mergedData = {};
    
    workbooks.forEach(workbook => {
        let firstSheetName = workbook.SheetNames[0];
        let firstSheet = workbook.Sheets[firstSheetName];
        let jsonData = xlsx.utils.sheet_to_json(firstSheet);

        jsonData.forEach(row => {
            let key = row[commonColumn];
            if (!mergedData[key]) {
                mergedData[key] = { ...row }; 
            } else {
                for (let col in row) {
                    if (!mergedData[key][col]) {
                        mergedData[key][col] = row[col] || '0'; // Add missing values
                    }
                }
            }
        });
    });

    return Object.values(mergedData);
}

app.listen(300, () => {
    console.log('Server running on port 3000');
});
