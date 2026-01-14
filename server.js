const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();  // New: Require SQLite

const app = express();
const upload = multer();

app.use(express.static('.'));  // Serve your HTML file
app.use(express.urlencoded({ extended: true }));

// New: SQLite database setup
const dbFile = 'rsvp_data.db';
const db = new sqlite3.Database(dbFile);

// New: Create table if it doesn't exist (run once on startup)
db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS rsvps (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        email TEXT NOT NULL,
        contact TEXT NOT NULL,
        guests INTEGER NOT NULL,
        attendance TEXT NOT NULL,
        guestNames TEXT
    )`);
});

app.post('/rsvp', upload.none(), (req, res) => {

    const { name, email, contact, guests, attendance, guestNames } = req.body;
    if (!name || !email || !contact || !guests || !attendance) {
        return res.status(400).send('Error: All required fields must be filled.');
}
    // New: Insert into SQLite instead of Excel
    const stmt = db.prepare('INSERT INTO rsvps (name, email, contact, guests, attendance, guestNames) VALUES (?, ?, ?, ?, ?, ?)');
    stmt.run(name, email, contact, parseInt(guests), attendance, guestNames || '', (err) => {
        if (err) {
            console.error('Error saving to SQLite:', err);
            return res.status(500).send('Error: Could not save RSVP. Please try again.');
        }
        stmt.finalize();
        res.send('RSVP submitted! Thank you.');
    });
});

app.get('/download', (req, res) =>  {
    // New: Query all data from SQLite and generate Excel on-the-fly
    db.all('SELECT name, email, contact, guests, attendance, guestNames FROM rsvps', [], (err, rows) => {
        if (err) {
            console.error('Error querying database:', err);
            return res.status(500).send('Error: Could not retrieve data.');
        }

        if (rows.length === 0) {
            return res.send('No RSVP data yet!');
        }

        // Convert rows to Excel format
        const data = [['Name', 'Email', 'Contact', 'Guests', 'Attendance', 'Guest Names'], ...rows.map(row => [
            row.name, row.email, row.contact, row.guests, row.attendance, row.guestNames || ''
        ])];
        const workbook = XLSX.utils.book_new();
        const sheet = XLSX.utils.aoa_to_sheet(data);
        sheet['!cols'] = [
            { wch: 15 },  // name
            { wch: 25 },  // email
            { wch: 15 },  // contact
            { wch: 10 },  // guests
            { wch: 12 },  // attendance
            { wch: 30 }   // guest names
        ];
        XLSX.utils.book_append_sheet(workbook, sheet, 'RSVPs');

        // Write to a temporary file and download
        const tempFile = 'temp_rsvp_data.xlsx';
        XLSX.writeFile(workbook, tempFile);
        res.download(tempFile, 'rsvp_data.xlsx', (err) => {
            if (err) console.error('Download error:', err);
            fs.unlinkSync(tempFile);  // Clean up temp file
        });
    });
});
    
app.listen(process.env.PORT || 3000, () => console.log('Server running'));
