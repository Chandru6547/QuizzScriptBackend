const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const cors = require("cors");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(cors());
app.use(express.json());

// Multer setup for file uploads
const upload = multer({ dest: "uploads/" });

app.post("/upload", upload.array("files", 10), (req, res) => {
    try {
        let studentData = {};
        let allTestNames = new Set();
        let testMarks = {}; // Store valid marks for each test
        let outputFileName = req.body.outputFileName || "output.xlsx"; // Get filename or set default

        req.files.forEach((file, index) => {
            const testName = `Test_${index + 1}`;
            const workbook = XLSX.readFile(file.path);
            const sheetName = "Participant Data";

            if (!workbook.Sheets[sheetName]) {
                console.log(`INFO: Sheet "${sheetName}" not found in ${file.originalname}`);
                return;
            }

            const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            allTestNames.add(testName);
            testMarks[testName] = [];

            sheet.forEach(row => {
                let rollNumber = row["First Name"] || row["Firstname"];
                if (!rollNumber) return;

                rollNumber = rollNumber.trim().toLowerCase();
                let percentage = row["Accuracy"] || "Absent";

                if (typeof percentage === "string") {
                    percentage = percentage.replace("%", "").trim();
                }

                if (percentage !== "Absent") {
                    testMarks[testName].push(percentage);
                }

                if (!studentData[rollNumber]) {
                    studentData[rollNumber] = { RollNumber: rollNumber };
                }

                studentData[rollNumber][testName] = percentage;
            });

            fs.unlinkSync(file.path); // Remove uploaded file after processing
        });

        // Fill absent marks with random valid values
        Object.values(studentData).forEach(student => {
            allTestNames.forEach(testName => {
                if (!student[testName] || student[testName] === "Absent") {
                    student[testName] = getRandomMark(testMarks[testName]);
                }
            });
        });

        const finalData = Object.values(studentData);

        // Convert JSON to Excel
        const newWorkbook = XLSX.utils.book_new();
        const newSheet = XLSX.utils.json_to_sheet(finalData);
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Consolidated Data");
        
        const outputFilePath = path.join(__dirname, outputFileName);
        XLSX.writeFile(newWorkbook, outputFilePath);

        res.json({ message: "File processed successfully", downloadLink: `/download?fileName=${outputFileName}` });

    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "Internal Server Error" });
    }
});

// Endpoint to download processed Excel file
app.get("/download", (req, res) => {
    console.log(req.query);
    
    const fileName = req.query.filename || "output.xlsx";
    const filePath = path.join(__dirname, fileName);

    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).json({ error: "File not found" });
    }
});

function getRandomMark(marks) {
    return marks.length > 0 ? marks[Math.floor(Math.random() * marks.length)] : "0";
}
