const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const app = express();
const PORT = 3010;

const destinationFolder = path.join(__dirname, '../../Files/Contribution');
const excelFilePath = path.join(__dirname, '../../Files/Contribution/contribution.xlsx');

if (!fs.existsSync(destinationFolder)) {
  fs.mkdirSync(destinationFolder);
}

// Middleware to parse URL-encoded bodies
app.use(express.urlencoded({ extended: true }));

// Multer configuration for file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, destinationFolder);
  },
  filename: function (req, file, cb) {
    const nameInput = req.body.nameInput || "undefined";
    const batchInput = req.body.batchInput || "undefined";
    const resourceTypeInput = req.body.resourceTypeInput || "undefined";
    const courseNoInput = req.body.courseNoInput || "undefined";
    const teacherNameInput = req.body.teacherNameInput || "undefined";
    const filename = `${nameInput}-${batchInput}-${resourceTypeInput}-${courseNoInput}-${teacherNameInput}-${file.originalname}`;
    cb(null, filename);
  },
});

const upload = multer({ storage: storage });

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));

// Route for file upload
app.post("/upload", upload.single("fileInput"), async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    let worksheet;
    if (fs.existsSync(excelFilePath)) {
      await workbook.xlsx.readFile(excelFilePath);
      worksheet = workbook.getWorksheet("Contributions");
    } else {
      workbook.addWorksheet("Contributions");
      worksheet = workbook.getWorksheet("Contributions");
      worksheet.addRow([
        "Name",
        "Batch",
        "Resource Type",
        "Course No.",
        "Teacher's Name",
        "File",
      ]);
    }

    worksheet.addRow([
      req.body.nameInput,
      req.body.batchInput,
      req.body.resourceTypeInput,
      req.body.courseNoInput,
      req.body.teacherNameInput,
      req.file.filename,
    ]);

    await workbook.xlsx.writeFile(excelFilePath);

    // Redirect to success.html after successful upload
    res.sendFile(path.join(__dirname, 'public', 'success.html'));
  } catch (error) {
    console.error("Error uploading file:", error);
    res.status(500).send("Error uploading file.");
  }
});

// Start the server
app.listen(PORT, () => {
  console.log("Server running on port", PORT);
});
