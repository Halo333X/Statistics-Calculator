const promptSync = require('prompt-sync');
const Excel = require('./Excel');
const fs = require('fs');
const prompt = promptSync();

try {
    const excel = new Excel('build/file.xlsx');
    const column = prompt("Ingresa la columna que quieres leer (por ejemplo, 'A'): ").toUpperCase();
    const startRow = parseInt(prompt("Ingresa la fila inicial (por ejemplo, '1'): "));
    const endRow = parseInt(prompt("Ingresa la fila final (por ejemplo, '5'): "));

    const data = excel.getData(column, startRow, endRow);

    if (data.length === 0) {
        console.log("No se encontraron datos en la columna y rango especificados.");
        process.exit(1);
    }

    excel.calculateStatistics(data);
} catch (error) {
    console.error('Error:', error);
    fs.writeFileSync('error.txt', JSON.stringify(error, null, 4), 'utf-8');
    process.exit(1);
}