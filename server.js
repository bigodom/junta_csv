import express from 'express';
import multer from 'multer';
import { promises as fsPromises } from 'fs';
import ExcelJS from 'exceljs';
import path from 'path';

const { readFile, writeFile, unlink } = fsPromises;

const app = express();
const port = 3000;

const upload = multer({ dest: 'uploads/' });

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

app.post('/upload', upload.array('csvFiles'), async (req, res) => {
    try {

        const files = req.files;
        const fileExtension = path.extname(files[0].originalname).toLowerCase();
        const combinedData = [];

        let isFirstFile = true;

        for (const file of files) {
            if (fileExtension === '.csv') {
                console.log('juntando o ' + file.originalname);
                const content = await readFile(file.path, 'utf8');
                const lines = content.split('\n');

                lines.forEach((line, index) => {
                    if (line.trim() !== '') {
                        const record = line.split(';');
                        if (index === 0 && isFirstFile) {
                            combinedData.push(record);
                            isFirstFile = false;
                        } else if (index !== 0) {
                            combinedData.push(record);
                        }
                    }
                });
            } else if (fileExtension === '.xlsx') {
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.readFile(file.path);
                const sheet = workbook.worksheets[0];

                sheet.eachRow((row, rowNumber) => {
                    if (rowNumber === 1 && isFirstFile) {
                        combinedData.push(row.values);
                        isFirstFile = false;
                    } else if (rowNumber !== 1) {
                        combinedData.push(row.values);
                    }
                });
            }
        }

        if (fileExtension === '.xlsx') {
            const workbook = new ExcelJS.Workbook();
            const sheet = workbook.addWorksheet('Combined Data');

            combinedData.forEach(record => {
                sheet.addRow(record);
            });

            const fileName = 'combined.xlsx';

            console.log('Escrevendo o arquivo ' + fileName);
            await workbook.xlsx.writeFile(fileName);
            console.log('Arquivo ' + fileName + ' escrito com sucesso!');

            cleanupFiles();
        } else {
            const combinedContent = combinedData.map(record => record.join(';')).join('\n');
            const fileName = 'combined.csv';

            console.log('Escrevendo o arquivo ' + fileName);
            await writeFile(fileName, combinedContent, 'utf8');
            console.log('Arquivo ' + fileName + ' escrito com sucesso!');

            cleanupFiles();
        }

        // Function to clean up files after download
        async function cleanupFiles() {
            for (const file of files) {
                await unlink(file.path);
            }
        }
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Internal Server Error');
    }
});

app.listen(port, () => {
    console.log(`Server is running at http://localhost:${port}`);
});
