import Excel from 'exceljs';
import mysql from 'mysql';
import dotenv from 'dotenv';
dotenv.config();

// MySQL connection setup
const connection = mysql.createConnection({
    host: process.env.HOST_NAME,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_DATABASE
});


connection.connect(err => {
    if (err) throw err;
    console.log('Connected to MySQL server.');
});

// Function to read and migrate Excel data
async function migrateData() {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile('./product_list.xlsx');

    let currentParentCategory = '';
    let currentSubcategory1 = '';
    let currentSubcategory2 = '';

    const sheet = workbook.getWorksheet(1); // Assuming data is in the first sheet
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 1) { // Skipping header row
            const cells = row.values;
            const parentCategory = cells[1] && cells[1] !== null ? cells[1].toString().trim() : currentParentCategory;
            const subcategory1 = cells[2] && cells[2] !== null ? cells[2].toString().trim() : currentSubcategory1;
            const subcategory2 = cells[3] && cells[3] !== null ? cells[3].toString().trim() : currentSubcategory2;

            // Update current categories if they have changed
            if (cells[1] && cells[1] !== null) currentParentCategory = parentCategory;
            if (cells[2] && cells[2] !== null) currentSubcategory1 = subcategory1;
            if (cells[3] && cells[3] !== null) currentSubcategory2 = subcategory2;
            const attributeHeader = row.getCell(4).value;
            const attributeValues = row.getCell(5).value.split(',');

            attributeValues.forEach(value => {
                const insertQuery = 'INSERT INTO MigCategoryInfo (parent_category, subcategory_1, subcategory_2, attribute, attribute_value) VALUES (?, ?, ?, ?, ?)';
                connection.query(insertQuery, [parentCategory, subcategory1, subcategory2, attributeHeader, value.trim()], (err, results, fields) => {
                    if (err) {
                        console.error('Error inserting data: ', err);
                        return;
                    }
                    console.log('Inserted row with ID: ', results.insertId);
                });
            });
        }
    });
}

migrateData().then(() => {
    console.log('Data migration complete.');
    connection.end();
}).catch(err => {
    console.error('Failed to migrate data: ', err);
    connection.end();
});