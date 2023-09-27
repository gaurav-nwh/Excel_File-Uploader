const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 5000;

const corsOptions = {
  origin: 'http://localhost:3000',
};

app.use(cors(corsOptions));

mongoose.connect('mongodb+srv://excel:12345@excel.qnqgxgd.mongodb.net/excel?retryWrites=true&w=majority', {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

const db = mongoose.connection;
db.on('error', console.error.bind(console, 'MongoDB connection error:'));
db.once('open', () => {
  console.log('Connected to MongoDB');
});

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const excelSchema = new mongoose.Schema({

  S_No: {
    type: Number
  },
  Name: {
    type: String
  },
  Email: {
    type: String
  },
  Mobile: {
    type: Number
  }
},{ collection: 'excel' });

const ExcelModel = mongoose.model('Excel', excelSchema);

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'No file uploaded.' });
    }

    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0]; 

    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    await ExcelModel.create(data);

    return res.status(200).json({ success: true, message: 'File uploaded successfully.' });
  } catch (error) {
    console.error('Error uploading file:', error);
    return res.status(500).json({ success: false, message: 'File upload failed.' });
  }
});

app.use(express.static(path.join(__dirname, 'frontend/build')));

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'frontend/build', 'index.html'));
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
