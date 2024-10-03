import BuatExcel from "./index.js";

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

exampleUsage();