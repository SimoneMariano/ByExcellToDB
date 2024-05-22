const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const mysql = require("mysql2/promise");
const fs = require("fs");

const app = express();
const PORT = 3000;

// Set up MySQL connection pool
const pool = mysql.createPool({
    host: "127.0.0.1",
    user: "root",
    password: "",
    database: "db_pallavolo",
});

// Configurazione multer per il caricamento dei file
const upload = multer({ dest: "uploads/" });

// Pagina HTML con un bottone per l'esecuzione
const htmlPage = `
<html>
<head>
    <title>Esegui comando e invia file</title>
</head>
<body>
    <h1>Esegui comando e invia file</h1>
    <form action="/execute" method="post" enctype="multipart/form-data">
        <input type="file" name="file">
        <button type="submit">Esegui</button>
    </form>
</body>
</html>
`;

// Endpoint per la visualizzazione della pagina HTML
app.get("/", (req, res) => {
    res.send(htmlPage);
});

// Endpoint per eseguire il caricamento del file e l'inserimento dei dati nel database
app.post("/execute", upload.single("file"), async (req, res) => {
    try {
        // Verifica se è stato caricato un file
        if (!req.file) {
            console.log("No file uploaded");
            return res.status(400).send("No file uploaded");
        }

        console.log(`File uploaded: ${req.file.path}`);

        // Leggi il file Excel
        const workbook = xlsx.readFile(req.file.path);
        console.log("Excel file read successfully");

        // Funzione per elaborare ciascun foglio
        const processSheet = async (sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

            // Estrai i nomi delle colonne dalla prima riga
            const columnNames = data[0];
            const rows = data.slice(1);

            // Incapsula i nomi delle colonne tra backticks
            const wrappedColumnNames = columnNames.map(name => `\`${name}\``);

            console.log(`Processing sheet: ${sheetName} with columns: ${wrappedColumnNames.join(', ')}`);

            const connection = await pool.getConnection();
            try {
                await connection.beginTransaction();

                // Elimina i dati esistenti nella tabella per un determinato 'codiceTorneo'
                const codiceTorneoIndex = columnNames.indexOf('codiceTorneo');
                if (codiceTorneoIndex === -1) {
                    throw new Error(`'codiceTorneo' column not found in sheet ${sheetName}`);
                }

                const codiceTorneo = rows[0][codiceTorneoIndex]; // Assuming the same 'codiceTorneo' for all rows in the sheet
                const deleteQuery = `DELETE FROM ?? WHERE codiceTorneo = ?`;
                console.log(`Executing query: ${deleteQuery} with values: [${sheetName}, ${codiceTorneo}]`);
                await connection.query(deleteQuery, [sheetName, codiceTorneo]);
                console.log(`Existing data deleted from table: ${sheetName}`);

                // Inserisci nuovi dati
                const insertQuery = `INSERT INTO ?? (${wrappedColumnNames.join(',')}) VALUES (${columnNames.map(() => '?').join(',')})`;
                for (let row of rows) {
                    const values = row.map((value) => value || null); // Handle empty values
                    console.log(`Executing query: ${insertQuery} with values: [${sheetName}, ${values.join(', ')}]`);
                    await connection.query(insertQuery, [sheetName, ...values]);
                }

                await connection.commit();
                console.log(`Data inserted into table: ${sheetName}`);
            } catch (error) {
                await connection.rollback();
                console.error(`Error processing sheet: ${sheetName}`, error);
                throw error; // Propagate the error
            } finally {
                connection.release();
            }
        };

        // Itera attraverso tutti i nomi dei fogli
        const sheetNames = workbook.SheetNames;
        for (const sheetName of sheetNames) {
            await processSheet(sheetName);
        }

        // Rimuovi il file caricato dopo l'elaborazione
        fs.unlinkSync(req.file.path);
        console.log(`Temporary file deleted: ${req.file.path}`);

        res.send("Data uploaded and updated successfully" + htmlPage);
        
    } catch (err) {
        console.error("Error occurred during data processing", err);
        res.status(500).send("Error occurred during data processing");
    }
});

// Avvia il server
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
