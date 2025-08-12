```js
// student-form.js — Single-file Node.js app serving the HTML form and saving to Excel

const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE = path.join(__dirname, 'students.xlsx');
const SHEET_NAME = 'Students';

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

function appendToExcel(obj) {
  let workbook = fs.existsSync(EXCEL_FILE) ? XLSX.readFile(EXCEL_FILE) : XLSX.utils.book_new();
  let worksheet = workbook.Sheets[SHEET_NAME];
  let data = worksheet ? XLSX.utils.sheet_to_json(worksheet, { defval: '' }) : [];
  data.push({ ...obj, timestamp: new Date().toISOString() });
  workbook.Sheets[SHEET_NAME] = XLSX.utils.json_to_sheet(data);
  if (!workbook.SheetNames.includes(SHEET_NAME)) workbook.SheetNames.push(SHEET_NAME);
  XLSX.writeFile(workbook, EXCEL_FILE);
}

app.get('/', (req, res) => {
  res.send(`<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Student Information Form</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container py-5">
  <div class="card shadow-sm">
    <div class="card-body">
      <h3 class="card-title mb-3">Student Information Form</h3>
      <form id="studentForm" class="row g-3">
        <div class="col-md-6"><label class="form-label">Full name</label><input type="text" name="name" class="form-control" required></div>
        <div class="col-md-6"><label class="form-label">Email</label><input type="email" name="email" class="form-control"></div>
        <div class="col-md-4"><label class="form-label">Class / Grade</label><input type="text" name="class" class="form-control"></div>
        <div class="col-md-4"><label class="form-label">Roll number</label><input type="text" name="roll" class="form-control"></div>
        <div class="col-md-4"><label class="form-label">Phone</label><input type="tel" name="phone" class="form-control"></div>
        <div class="col-12"><label class="form-label">Address</label><input type="text" name="address" class="form-control"></div>
        <div class="col-md-4"><label class="form-label">Date of birth</label><input type="date" name="dob" class="form-control"></div>
        <div class="col-md-4"><label class="form-label">Gender</label><select name="gender" class="form-select"><option value="">Select</option><option>Male</option><option>Female</option><option>Other</option></select></div>
        <div class="col-md-4"><label class="form-label">Parent / Guardian</label><input type="text" name="parent" class="form-control"></div>
        <div class="col-md-6"><label class="form-label">Parent Contact</label><input type="tel" name="parent_contact" class="form-control"></div>
        <div class="col-md-6"><label class="form-label">Notes</label><input type="text" name="notes" class="form-control"></div>
        <div class="col-12 mt-4 d-flex gap-2"><button class="btn btn-primary" type="submit">Submit</button><button class="btn btn-outline-secondary" id="clearBtn" type="button">Clear</button></div>
      </form>
      <div id="alertPlaceholder" class="mt-3"></div>
    </div>
  </div>
</div>
<script>
const form = document.getElementById('studentForm');
const alertPlaceholder = document.getElementById('alertPlaceholder');
const clearBtn = document.getElementById('clearBtn');
function showMessage(message, type = 'success') {
  alertPlaceholder.innerHTML = `<div class="alert alert-${type}" role="alert">${message}</div>`;
}
form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const fd = new FormData(form);
  const data = Object.fromEntries(fd.entries());
  try {
    const res = await fetch('/submit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    if (!res.ok) throw new Error('Server returned ' + res.status);
    const json = await res.json();
    showMessage(json.message || 'Saved successfully');
    form.reset();
  } catch (err) {
    showMessage('Error saving data: ' + err.message, 'danger');
  }
});
clearBtn.addEventListener('click', () => form.reset());
</script>
</body>
</html>`);
});

app.post('/submit', (req, res) => {
  try {
    if (!req.body.name) return res.status(400).json({ message: 'Name is required' });
    appendToExcel(req.body);
    res.json({ message: 'Student saved to Excel' });
  } catch (err) {
    res.status(500).json({ message: 'Internal server error' });
  }
});

app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
```

**How to run:**

1. Save as `student-form.js`.
2. `npm init -y && npm install express xlsx`
3. Run with `node student-form.js`
4. Open `http://localhost:3000` in your browser.
5. Submit the form — `students.xlsx` will be created/updated in the same folder.
