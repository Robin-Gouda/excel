import express from "express";
import xlsx from "xlsx";
import mysql from "mssql";
import fs from "fs";
import axios from "axios";
import "dotenv/config";

const app = express();
const port = 3000;

// Load Excel file
const workbook = xlsx.readFile("data.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert Excel data to JSON
const jsonData = xlsx.utils.sheet_to_json(worksheet);

// SQL Server configuration
const sqlConfig = {
  user: process.env.USER,
  password: process.env.PASSWORD,
  server: process.env.SERVER,
  database: process.env.DATABASE,
};

// Connect to SQL Server
mysql
  .connect(sqlConfig)
  .then(() => {
    // Define table schema
    const request = new sql.Request();
    request
      .query(
        `CREATE TABLE Student (
      Column1 VARCHAR(50),
      Column2 VARCHAR(50),
      Column3 VARCHAR(50),
      Column4 VARCHAR(50)
    )`
      )
      .then(() => {
        // Insert data into SQL table
        jsonData.forEach((row) => {
          request
            .query(
              `INSERT INTO SampleData (Column1, Column2, Column3, Column4)
                       VALUES ('${row.Column1}', '${row.Column2}', '${row.Column3}', '${row.Column4}')`
            )
            .catch((err) => console.error(err));
        });
      })
      .catch((err) => console.error(err));
  })
  .catch((err) => console.error(err));

// Format data for REST API
const information = jsonData.map((row) => ({
  column1: row.Column1,
  column2: row.Column2,
  column3: row.Column3,
  column4: row.Column4,
}));

// Send data to REST API endpoint
axios
  .post("/upload", information)
  .then((response) => console.log("Data sent successfully:", response.data))
  .catch((error) => console.error("Error:", error));

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
