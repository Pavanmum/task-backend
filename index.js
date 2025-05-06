const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const csv = require('csv-parser');
const fs = require('fs');
const cors = require('cors');
const dotenv = require('dotenv');
const path = require('path');

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadPath = '/tmp/uploads';
    if (!fs.existsSync(uploadPath)) {
      fs.mkdirSync(uploadPath, { recursive: true });
    }
    cb(null, uploadPath);
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  },
});
const upload = multer({ storage });

const companySchema = new mongoose.Schema({
  name: String,
  industry: String,
  location: String,
  email: { type: String, unique: true, required: true },
  phone: String,
});
const Company = mongoose.model('Company', companySchema);

(async () => {
  try {
    await mongoose.connect(process.env.MONGO_URL, {
      useNewUrlParser: true,
      useUnifiedTopology: true,
    });
    console.log('Connected to MongoDB');
  } catch (error) {
    console.error('Error connecting to MongoDB:', error);
  }
})();

const processFile = (filePath, isExcel) => {
  return new Promise((resolve, reject) => {
    const companies = [];
    
    if (isExcel) {
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(sheet);
      companies.push(...data);
      resolve(companies);
    } else {
      fs.createReadStream(filePath)
        .pipe(csv())
        .on('data', (row) => companies.push(row))
        .on('end', () => resolve(companies))
        .on('error', reject);
    }
  });
};

const importCompanies = async (companies, importMode) => {
  let inserted = 0, updated = 0, skipped = 0;

  for (const company of companies) {
    if (!company.email) {
      skipped++;
      continue;
    }

    const existingCompany = await Company.findOne({ email: company.email });

    switch (importMode) {
      case 'create_new':
        if (!existingCompany) {
          await Company.create(company);
          inserted++;
        } else {
          skipped++;
        }
        break;

      case 'create_update_no_overwrite':
        if (!existingCompany) {
          await Company.create(company);
          inserted++;
        } else {
          const update = {};
          for (const key in company) {
            if (!existingCompany[key] && company[key]) {
              update[key] = company[key];
            }
          }
          if (Object.keys(update).length > 0) {
            await Company.updateOne({ email: company.email }, { $set: update });
            updated++;
          } else {
            skipped++;
          }
        }
        break;

      case 'create_update_overwrite':
        if (!existingCompany) {
          await Company.create(company);
          inserted++;
        } else {
          await Company.updateOne({ email: company.email }, { $set: company });
          updated++;
        }
        break;

      case 'update_no_overwrite':
        if (existingCompany) {
          const update = {};
          for (const key in company) {
            if (!existingCompany[key] && company[key]) {
              update[key] = company[key];
            }
          }
          if (Object.keys(update).length > 0) {
            await Company.updateOne({ email: company.email }, { $set: update });
            updated++;
          } else {
            skipped++;
          }
        } else {
          skipped++;
        }
        break;

      case 'update_overwrite':
        if (existingCompany) {
          await Company.updateOne({ email: company.email }, { $set: company });
          updated++;
        } else {
          skipped++;
        }
        break;
    }
  }

  return { inserted, updated, skipped };
};

app.post('/api/companies/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'No file uploaded' });
    }

    const { importMode } = req.body;
    if (!importMode) {
      return res.status(400).json({ message: 'Import mode is required' });
    }

    const isExcel = req.file.mimetype.includes('excel') || req.file.mimetype.includes('spreadsheet');
    const companies = await processFile(req.file.path, isExcel);
    
    const result = await importCompanies(companies, importMode);
    
    // Clean up uploaded file
    fs.unlinkSync(req.file.path);

    res.json({
      status: 'success',
      inserted: result.inserted,
      updated: result.updated,
      skipped: result.skipped,
    });
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

app.get('/api/companies', async (req, res) => {
  try {
    const companies = await Company.find({}, 'name industry email phone');
    res.json(companies);
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

app.delete('/api/companies/:email', async (req, res) => {
  try {
    const { email } = req.params;
    const result = await Company.deleteOne({ email });
    if (result.deletedCount === 0) {
      return res.status(404).json({ message: 'Company not found' });
    }
    res.json({ message: 'Company deleted successfully' });
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

app.listen(5000, () => console.log('Server running on port 5000'));