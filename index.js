const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const PORT = 3000;

app.set('view engine', 'ejs');

// Set up multer storage
const storage = multer.memoryStorage();
const upload = multer({ storage });

// Serve the upload.html file
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'upload.html'));
});

// Handle file upload and join data
app.post('/upload', upload.fields([{ name: 'file1' }, { name: 'file2' }]), (req, res) => {
  const file1 = req.files['file1'][0];
  const file2 = req.files['file2'][0];

  const workbook1 = xlsx.read(file1.buffer, { type: 'buffer' });
  const workbook2 = xlsx.read(file2.buffer, { type: 'buffer' });

  const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
  const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];

  const data1 = xlsx.utils.sheet_to_json(sheet1);
  const data2 = xlsx.utils.sheet_to_json(sheet2);

  const joinedData = data1.map(row1 => {
    const matchingRow = data2.find(row2 => row2['Phone'] === row1['Phone']);

    if (matchingRow) {
      return {
        Name: row1['Name'],
        Address: row1['Address'],
        Phone: row1['Phone'],
        City: matchingRow['City']
      };
    } else {
      return {
        Name: row1['Name'],
        Address: row1['Address'],
        Phone: row1['Phone'],
        City: 'N/A'
      };
    }
  });

  res.render('result.ejs', { joinedData });

});

// Download the joined data
// Download the joined data
app.get('/download', (req, res) => {
  const joinedData = JSON.parse(req.query.joinedData);

  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.json_to_sheet(joinedData);
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Joined Data');

  const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Disposition', 'attachment; filename=joined_data.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(Buffer.from(buffer));
});


// Start the server
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
