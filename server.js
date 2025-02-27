const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const cors = require("cors");
const fs = require("fs");
const path = require("path");

const app = express();
app.use(cors());
app.use(express.json());

const upload = multer({ dest: "uploads/" });

app.post("/upload", upload.array("files", 10), (req, res) => {
    try {
        console.log("Received file upload request.");

        let studentData = {};
        let allTestNames = new Set();
        let testMarks = {};

        if (!req.files || req.files.length === 0) {
            console.warn("No files uploaded.");
            return res.status(400).json({ error: "No files uploaded" });
        }

        req.files.forEach((file, index) => {
            console.log(`Processing file: ${file.originalname}`);

            const testName = `Test_${index + 1}`;
            const workbook = XLSX.readFile(file.path);
            const sheetName = "Participant Data";

            if (!workbook.Sheets[sheetName]) {
                console.log(`⚠️ INFO: Sheet "${sheetName}" not found in ${file.originalname}`);
                return;
            }

            const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            console.log(`✅ Found ${sheet.length} rows in sheet "${sheetName}".`);

            allTestNames.add(testName);
            testMarks[testName] = [];

            sheet.forEach(row => {
                let rollNumber = row["First Name"] || row["Firstname"];
                if (!rollNumber) {
                    console.warn("⚠️ Skipping row without roll number:", row);
                    return;
                }

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
                    console.log(`ℹ️ Added student: ${rollNumber}`);
                }

                studentData[rollNumber][testName] = percentage;
                console.log(`📌 ${rollNumber} - ${testName}: ${percentage}`);
            });

            console.log(`🗑️ Deleting temporary file: ${file.path}`);
            fs.unlinkSync(file.path);
        });

        // Filling missing marks
        console.log("🔄 Filling missing marks for absent students...");
        Object.values(studentData).forEach(student => {
            allTestNames.forEach(testName => {
                if (!student[testName] || student[testName] === "Absent") {
                    let randomMark = getRandomMark(testMarks[testName]);
                    student[testName] = randomMark;
                    console.log(`🎲 Assigned random mark (${randomMark}) to ${student.RollNumber} for ${testName}`);
                }
            });
        });

        const finalData = Object.values(studentData);
        console.log("✅ File processing complete. Sending response...");
        res.json({ message: "File processed successfully", data: finalData });

    } catch (error) {
        console.error("❌ Error during file processing:", error);
        res.status(500).json({ error: "Internal Server Error" });
    }
});

function getRandomMark(marks) {
    return marks.length > 0 ? marks[Math.floor(Math.random() * marks.length)] : "0";
}

const PORT = 5000;
// app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
