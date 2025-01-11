// Import dependencies
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = 5000;
app.use(cors());

// Configure Multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Endpoint to extract headers dynamically
app.post('/headers', upload.single('file'), (req, res) => {
  try {
    const filePath = req.file.path;

    // Read the uploaded file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Extract headers
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    const headers = data[0]; // First row contains headers

    if (!headers) {
      throw new Error('No headers found in the file.');
    }

    // Send headers back to the frontend
    res.json({ headers });

    // Clean up the uploaded file
    fs.unlinkSync(filePath);
  } catch (error) {
    console.error('Error extracting headers:', error);
    res.status(500).send('Error extracting headers');
  }
});

// Endpoint to process the Excel file
// Helper function to clean phone numbers
// Helper function to clean and format phone numbers
// Helper function to clean and format phone numbers
const cleanPhoneNumber = (phone) => {
    if (!phone) return null;
    return phone
      .replace(/\+91/g, '') // Remove +91
      .replace(/[^0-9]/g, '') // Remove non-numeric characters
      .slice(-10); // Keep only the last 10 digits
  };
  
  // Process endpoint
  app.post('/process', upload.single('file'), (req, res) => {
    try {
      const filePath = req.file.path;
      const selectedFields = JSON.parse(req.body.fields);
  
      // Read the uploaded file
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
  
      // Convert sheet to JSON
      const data = xlsx.utils.sheet_to_json(sheet);
  
      // To store unique numbers across all rows
      const allUniqueNumbers = new Set();
  
      // Filter data based on selected fields
     // Filter data based on selected fields
const filteredData = data.map((row) => {
    const filteredRow = {};
  
    // Retain selected fields
    selectedFields.forEach((field) => {
      if (row[field]) {
        filteredRow[field] = row[field];
      }
    });
  
    // Process 'Phone 1 - Value' field if it exists and is a string
    if (row['Phone 1 - Value'] && typeof row['Phone 1 - Value'] === 'string') {
      // Split multiple numbers by delimiter and clean each number
      const phoneNumbers = row['Phone 1 - Value']
        .split(':::') // Split numbers by delimiter
        .map(cleanPhoneNumber) // Clean each number
        .filter((num) => num); // Remove null or invalid numbers
  
      // Deduplicate numbers within the same cell
      const uniqueNumbers = [...new Set(phoneNumbers)];
  
      // Add to the global set of unique numbers
      uniqueNumbers.forEach((num) => allUniqueNumbers.add(num));
  
      // Combine cleaned, deduplicated numbers back into a single string
      filteredRow['Phone 1 - Value'] = uniqueNumbers.join(', ');
    } else {
      // Handle rows where 'Phone 1 - Value' is missing or invalid
      filteredRow['Phone 1 - Value'] = '';
    }
  
    return filteredRow;
  });
  
  
      // Remove duplicates across rows using global unique numbers
      const uniqueData = Array.from(
        new Map(
          filteredData.map((row) => [row['Phone 1 - Value'], row]) // Use 'Phone 1 - Value' as the unique key
        ).values()
      );
  
      // Write the filtered data back to a new Excel file
      const newSheet = xlsx.utils.json_to_sheet(uniqueData);
      const newWorkbook = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'FilteredData');
  
      const outputFileName = `processed_file_${Date.now()}.xlsx`; // Unique file name
      const outputFilePath = path.join(__dirname, 'uploads', outputFileName);
      xlsx.writeFile(newWorkbook, outputFilePath);
  
      // Send the download link back to the frontend
      res.json({ downloadLink: `https://08c1-123-201-174-31.ngrok-free.app/download/${outputFileName}` });
  
      // Clean up the uploaded file
      fs.unlinkSync(filePath);
    } catch (error) {
      console.error('Error processing file:', error);
      res.status(500).send('Error processing file');
  
      // Clean up the uploaded file in case of error
      if (req.file) {
        fs.unlinkSync(req.file.path);
      }
    }
  });
  
  
  
  

// Serve the processed file for download
app.get('/download/:filename', (req, res) => {
  const filePath = path.join(__dirname, 'uploads', req.params.filename);
  res.download(filePath, (err) => {
    if (err) {
      console.error('Error sending file:', err);
    } else {
      // Clean up the processed file after download
      fs.unlinkSync(filePath);
    }
  });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
