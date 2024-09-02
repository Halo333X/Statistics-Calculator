const xlsx = require('xlsx');
const math = require('mathjs');
const fs = require('fs');
const promptSync = require('prompt-sync');
const { ChartJSNodeCanvas } = require('chartjs-node-canvas');
const path = require('path');

const graphicOutput = path.join(__dirname, '..', 'build', 'output', 'graphic.png');
const dataOutput = path.join(__dirname, '..', 'build', 'output', 'data.json');

const prompt = promptSync();

class Excel {
    constructor(fileName) {
        this.fileName = fileName;
        this.workbook = xlsx.readFile(fileName);
        this.sheetNameList = this.workbook.SheetNames;
        this.worksheet = this.workbook.Sheets[this.sheetNameList[0]];
        this.chartJSNodeCanvas = new ChartJSNodeCanvas({ width: 1600, height: 1200 });
    }

    getData(column, startRow, endRow) {
        let data = [];
        for (let i = startRow; i <= endRow; i++) {
            const cellAddress = column + i;
            const cell = this.worksheet[cellAddress];
            if (cell && typeof cell.v === 'number') {
                data.push(cell.v);
            }
        }
        if (data.length === 0) {
            console.log("No se encontraron datos en la columna y rango especificados.");
            return [];
        }
        return data;
    }

    calculateStatistics(data) {
        const fileData = {
            Sumatoria: math.sum(data),
            NumeroCasillas: data.length,
            Media: math.mean(data),
            Mediana: math.median(data),
            Moda: math.mode(data),
            DesviacionEstandar: math.std(data),
            Varianza: math.variance(data),
            Cuartiles: {
                q1: math.quantileSeq(data, 0.25, false),
                q2: math.quantileSeq(data, 0.50, false),
                q3: math.quantileSeq(data, 0.75, false),
                q4: math.quantileSeq(data, 1.00, false)
            },
            Percentiles: this.calculatePercentiles(data)
        };
        const json = JSON.stringify(fileData, null, 4);
        fs.writeFileSync(dataOutput, json, 'utf-8');
        this.generateCombinedChart(fileData);
        setTimeout(() => {
            process.exit();
        }, 1500);
    }

    calculatePercentiles(data) {
        let percentiles = {};
        for (let p = 10; p <= 100; p += 10) {
            percentiles[`p${p}`] = math.quantileSeq(data, p / 100, false);
        }
        return percentiles;
    }

    async generateCombinedChart(fileData) {
        const { Media, Mediana, Moda, DesviacionEstandar, Varianza, Cuartiles, Percentiles } = fileData;
        const chartData = {
            labels: [
                'Media', 'Mediana', 'Moda', 'Desviación Estándar', 'Varianza',
                'Cuartiles Q1', 'Cuartiles Q2', 'Cuartiles Q3', 'Cuartiles Q4',
                'Percentiles P10', 'Percentiles P20', 'Percentiles P30', 'Percentiles P40', 'Percentiles P50', 'Percentiles P60', 'Percentiles P70', 'Percentiles P80', 'Percentiles P90', 'Percentiles P100'
            ],
            datasets: [{
                label: 'Statistics',
                data: [
                    Media, Mediana, Moda, DesviacionEstandar, Varianza,
                    Cuartiles.q1, Cuartiles.q2, Cuartiles.q3, Cuartiles.q4,
                    Percentiles.p10, Percentiles.p20, Percentiles.p30, Percentiles.p40, Percentiles.p50,
                    Percentiles.p60, Percentiles.p70, Percentiles.p80, Percentiles.p90, Percentiles.p100
                ],
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1
            }]
        };
        const workName = prompt('Como se llama tu proyecto?: ');
        const configuration = {
            type: 'bar',
            data: chartData,
            options: {
                scales: {
                    x: {
                        ticks: {
                            maxRotation: 90,
                            minRotation: 45
                        }
                    },
                    y: {
                        beginAtZero: true
                    }
                },
                plugins: {
                    title: {
                        display: true,
                        text: workName ? workName : 'Exam'
                    }
                }
            }
        };
        const image = await this.chartJSNodeCanvas.renderToBuffer(configuration);
        fs.writeFileSync(graphicOutput, image);
        console.log('The file data.json and graphic were created successfully!');
    }
}

module.exports = Excel;
