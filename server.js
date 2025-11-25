const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3001;

// Setup multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = './uploads';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== '.xlsx' && ext !== '.xls') {
      return cb(new Error('Doar fisiere Excel (.xlsx, .xls) sunt acceptate'));
    }
    cb(null, true);
  }
});

app.use(express.static('public'));
app.use(express.json());

// Serve main page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Helper function to create Oblio XLS from products
function createOblioWorkbook(products) {
  const oblioData = [];

  // Header row with 27 columns (7 with data, 20 empty)
  const headerRow = [
    'Denumire produs', 'Cod produs', 'U.M.', 'Cantitate', 'Pret achizitie', 'Cota TVA', 'TVA inclus'
  ];
  for (let i = 0; i < 20; i++) {
    headerRow.push(null);
  }
  oblioData.push(headerRow);

  // Product rows with 27 columns each
  products.forEach(p => {
    const row = [
      p.denumire,
      p.cod,
      p.um,
      p.cantitate,
      p.pret,
      p.cotaTVA,
      p.tvaInclus
    ];
    for (let i = 0; i < 20; i++) {
      row.push(null);
    }
    oblioData.push(row);
  });

  // Add empty row at the end (27 null columns)
  const emptyRow = [];
  for (let i = 0; i < 27; i++) {
    emptyRow.push(null);
  }
  oblioData.push(emptyRow);

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(oblioData);

  // Set column widths
  sheet['!cols'] = [
    { wch: 60 }, { wch: 15 }, { wch: 6 }, { wch: 10 },
    { wch: 15 }, { wch: 10 }, { wch: 10 }
  ];

  XLSX.utils.book_append_sheet(workbook, sheet, 'sheet 1');
  return workbook;
}

// Convert Qogita to Oblio format
app.post('/convert', upload.single('factura'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Nu a fost incarcat niciun fisier' });
    }

    // Get markup and exchange rate from request body
    const markupRON = parseFloat(req.body.markup) || 0;
    const exchangeRate = parseFloat(req.body.exchangeRate) || 5.00;

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Find the header row (contains "Name", "GTIN", "Price", "Quantity")
    let headerRowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row && row[0] === 'Name' && row.includes('GTIN') && row.includes('Price') && row.includes('Quantity')) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'Format invalid - nu am gasit header-ul cu produse (Name, GTIN, Price, Quantity)' });
    }

    const headers = data[headerRowIndex];
    const nameIdx = headers.indexOf('Name');
    const gtinIdx = headers.indexOf('GTIN');
    const priceIdx = headers.indexOf('Price');
    const quantityIdx = headers.indexOf('Quantity');
    const vatIdx = headers.indexOf('VAT');
    const rateIdx = headers.indexOf('Rate');

    // Extract invoice info
    let invoiceId = '';
    let invoiceDate = '';
    for (let i = 0; i < headerRowIndex; i++) {
      if (data[i] && data[i][0] === 'Invoice ID') {
        invoiceId = data[i][1] || '';
      }
      if (data[i] && data[i][0] === 'Date') {
        invoiceDate = data[i][1] || '';
      }
    }

    // Extract products
    const products = [];
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      if (!row || !row[nameIdx] || row[nameIdx] === '' || row[0] === '© 2025 Qogita.') {
        continue;
      }

      const name = row[nameIdx] || '';
      const gtin = row[gtinIdx] || '';
      const price = parseFloat(row[priceIdx]) || 0;
      const quantity = parseInt(row[quantityIdx]) || 0;

      // Rate is 0 for ABC transactions (reverse charge)
      const vatRate = parseFloat(row[rateIdx]) || 0;

      if (name && quantity > 0) {
        products.push({
          denumire: name,
          cod: gtin.toString(),
          um: 'buc',
          cantitate: quantity,
          pret: price,
          cotaTVA: vatRate,
          tvaInclus: 'NU'
        });
      }
    }

    if (products.length === 0) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'Nu am gasit produse in factura' });
    }

    // Calculate total value for markup distribution
    const totalValue = products.reduce((sum, p) => sum + (p.pret * p.cantitate), 0);

    // Generate file for Firma 1 (original prices)
    const workbook1 = createOblioWorkbook(products);
    const filename1 = `oblio_firma1_${invoiceId || Date.now()}.xls`;
    const filepath1 = path.join('./uploads', filename1);
    XLSX.writeFile(workbook1, filepath1, { bookType: 'xls' });

    let filename2 = null;
    let totalValue2 = totalValue;
    let products2Preview = [];

    // Generate file for Firma 2 (with markup) - prices in RON
    // Always generate if exchangeRate > 0, markup is optional
    if (exchangeRate > 0) {
      const products2 = products.map(p => {
        // Convert EUR to RON first
        const priceRON = p.pret * exchangeRate;
        const productValueRON = priceRON * p.cantitate;

        // Calculate markup distribution based on RON value
        const totalValueRON = totalValue * exchangeRate;
        const productWeight = productValueRON / totalValueRON;
        const productMarkup = markupRON * productWeight;

        // Final price in RON = converted price + distributed markup
        const newPriceRON = priceRON + (productMarkup / p.cantitate);

        return {
          ...p,
          pret: parseFloat(newPriceRON.toFixed(2))
        };
      });

      totalValue2 = products2.reduce((sum, p) => sum + (p.pret * p.cantitate), 0);

      const workbook2 = createOblioWorkbook(products2);
      filename2 = `oblio_firma2_${invoiceId || Date.now()}.xls`;
      const filepath2 = path.join('./uploads', filename2);
      XLSX.writeFile(workbook2, filepath2, { bookType: 'xls' });

      products2Preview = products2.map(p => ({
        denumire: p.denumire.substring(0, 50) + (p.denumire.length > 50 ? '...' : ''),
        cod: p.cod,
        cantitate: p.cantitate,
        pret: p.pret
      }));
    }

    // Clean up input file
    fs.unlinkSync(req.file.path);

    const response = {
      success: true,
      invoiceId,
      invoiceDate,
      productsCount: products.length,
      firma1: {
        totalValue: totalValue.toFixed(2),
        downloadUrl: `/download/${filename1}`,
        products: products.map(p => ({
          denumire: p.denumire.substring(0, 50) + (p.denumire.length > 50 ? '...' : ''),
          cod: p.cod,
          cantitate: p.cantitate,
          pret: p.pret
        }))
      }
    };

    if (filename2) {
      response.firma2 = {
        markup: markupRON,
        exchangeRate: exchangeRate,
        currency: 'RON',
        totalValue: totalValue2.toFixed(2),
        downloadUrl: `/download/${filename2}`,
        products: products2Preview
      };
    }

    res.json(response);

  } catch (error) {
    console.error('Conversion error:', error);
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).json({ error: 'Eroare la procesare: ' + error.message });
  }
});

// Download converted file
app.get('/download/:filename', (req, res) => {
  const filepath = path.join(__dirname, 'uploads', req.params.filename);
  if (fs.existsSync(filepath)) {
    res.download(filepath, req.params.filename, (err) => {
      if (!err) {
        // Delete file after download
        setTimeout(() => {
          if (fs.existsSync(filepath)) {
            fs.unlinkSync(filepath);
          }
        }, 5000);
      }
    });
  } else {
    res.status(404).json({ error: 'Fisierul nu a fost gasit' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Open http://localhost:${PORT} in your browser`);
});
