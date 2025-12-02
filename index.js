const express = require("express");
const multer = require("multer");
const path = require("path");
const { parseActionPlan } = require("./excelReader");

const app = express();
const upload = multer({ dest: "uploads/" });

app.post("/upload", upload.single("file"), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "Missing file upload" });
        }

        const json = await parseActionPlan(req.file.path);

        res.json({
            success: true,
            data: json
        });

    } catch (err) {
        console.error(err);
        res.status(500).json({ message: "Error reading Excel file" });
    }
});

app.listen(3000, () => console.log("Server running on http://localhost:3000"));
