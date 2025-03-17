require('dotenv').config();
const express = require("express");
const bodyParser = require("body-parser");
const path = require("path");
const { saveToExcel } = require("./excelHandler");

const app = express();
const PORT = 3000;

const cors = require("cors");

// Enable CORS for frontend requests
app.use(cors({
    origin: "http://localhost:3000", // Change this if your frontend URL differs
    methods: ["POST", "GET"],
    allowedHeaders: ["Content-Type"]
}));


app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, "../frontend")));

app.post("/api/save-data", (req, res) => {
    const { name, email, idNumber } = req.body;

    saveToExcel({ name, email, idNumber });

    res.status(200).json({ message: "Data saved successfully!" });
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});
