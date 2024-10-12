const express = require('express');
const multer = require('multer');
const { PDFDocument } = require('pdf-lib');
const sharp = require('sharp');
const fs = require('fs').promises;
const path = require('path');
const { fromPath } = require('pdf2pic');
const pdf = require('pdf-parse');
const Excel = require('exceljs');
const { Document, Packer, Paragraph } = require('docx');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Ensure the uploads and outputs directories exist
Promise.all([
  fs.mkdir('uploads', { recursive: true }),
  fs.mkdir('outputs', { recursive: true })
]).catch(err => console.error('Error creating directories:', err));

// Root route
app.get('/', (req, res) => {
  res.send('PDF Conversion API is running. Use POST /convert to convert files.');
});

app.post('/convert', upload.single('pdf'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('No file uploaded.');
    }

    const inputPath = req.file.path;
    const outputPath = path.join('outputs', `converted_${Date.now()}`);
    const targetFormat = req.body.format || 'pdf'; // Default to PDF if no format specified

    let result;
    switch (targetFormat.toLowerCase()) {
      case 'pdf':
        result = await convertPdfToPdf(inputPath, `${outputPath}.pdf`);
        break;
      case 'word':
        result = await convertPdfToWord(inputPath, `${outputPath}.docx`);
        break;
      case 'excel':
        result = await convertPdfToExcel(inputPath, `${outputPath}.xlsx`);
        break;
      case 'jpeg':
      case 'jpg':
      case 'png':
        result = await convertPdfToImage(inputPath, `${outputPath}.${targetFormat}`, targetFormat);
        break;
      case 'html':
        result = await convertPdfToHtml(inputPath, `${outputPath}.html`);
        break;
      default:
        return res.status(400).send('Unsupported format');
    }

    res.download(result, path.basename(result), (err) => {
      if (err) {
        console.error('Error sending file:', err);
        res.status(500).send('Error sending file');
      }
      // Clean up temporary files
      fs.unlink(inputPath).catch(err => console.error('Error deleting input file:', err));
      fs.unlink(result).catch(err => console.error('Error deleting output file:', err));
    });
  } catch (error) {
    console.error('Error in /convert route:', error);
    res.status(500).send('An error occurred during conversion.');
  }
});

async function convertPdfToPdf(input, output) {
  const pdfDoc = await PDFDocument.load(await fs.readFile(input));
  const pdfBytes = await pdfDoc.save();
  await fs.writeFile(output, pdfBytes);
  return output;
}

async function convertPdfToWord(input, output) {
  const data = await pdf(await fs.readFile(input));
  const doc = new Document({
    sections: [{
      properties: {},
      children: [new Paragraph(data.text)]
    }]
  });
  const buffer = await Packer.toBuffer(doc);
  await fs.writeFile(output, buffer);
  return output;
}

async function convertPdfToExcel(input, output) {
  const data = await pdf(await fs.readFile(input));
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.addRow([data.text]);
  await workbook.xlsx.writeFile(output);
  return output;
}

async function convertPdfToImage(input, output, format) {
  const options = {
    density: 100,
    saveFilename: path.basename(output),
    savePath: path.dirname(output),
    format: format.toUpperCase(),
    width: 600,
    height: 600
  };
  const convert = fromPath(input, options);
  await convert(1); // Convert first page
  return output;
}

async function convertPdfToHtml(input, output) {
  const data = await pdf(await fs.readFile(input));
  const html = `<html><body>${data.text}</body></html>`;
  await fs.writeFile(output, html);
  return output;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`PDF Conversion API running on port ${PORT}`);
});

// Error handling for uncaught exceptions
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  process.exit(1);
});

// Error handling for unhandled promise rejections
process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
  // Application specific logging, throwing an error, or other logic here
});