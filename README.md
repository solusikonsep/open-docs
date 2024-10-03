## OPEN DOCS

- BASIC EXAMPLE

```javascript

const exampleUsage = async () => {
    try {
        // Example 1: Get buffer
        const bufferResult = await BuatExcel({
            data: [
                { Name: "Alam Wibowo", Umur: 28, Kota: "Jakarta Pusat" },
                { Name: "Sherly Smith", Umur: 34, Kota: "Los Angeles" }
            ],
            columns: [
                { header: 'Nama', key: 'Nama', width: 30 },
                { header: 'Umur', key: 'Umur', width: 10 },
                { header: 'Kota', key: 'Kota', width: 30 }
            ]
        });
        console.log('Buffer created, size:', bufferResult.buffer.length);

        // Example 2: Save to file
        const fileResult = await BuatExcel({
              data: [
                { Name: "Alam Wibowo", Umur: 28, Kota: "Jakarta Pusat" },
                { Name: "Sherly Smith", Umur: 34, Kota: "Los Angeles" }
            ],
            columns: [
                { header: 'Nama', key: 'Nama', width: 30 },
                { header: 'Umur', key: 'Umur', width: 10 },
                { header: 'Kota', key: 'Kota', width: 30 }
            ],
            filename: 'example.xlsx',
            download: true
        });
        console.log(fileResult.message);

    } catch (error) {
        console.error('Failed to create Excel:', error.message);
    }
};

```

- API ENDPOINT

```javascript
app.get('/download-excel', async (req, res) => {
    try {
        const result = await CREATE_EXCEL({
            data: yourData,
            columns: yourColumns
        });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${result.filename}"`);
        res.send(result.buffer);
    } catch (error) {
        res.status(500).send('Error generating Excel file');
    }
});
```