import Excel from 'exceljs';
const fs = require('fs');

interface IData {
    index: number;
    datetime: string;
    weight: number;
    mark: string;
}

const load = async () => {
    const workbook = new Excel.Workbook();
    try {
        await workbook.xlsx.readFile("Kurz Tagebuch mit messwerten.xlsx");
    }
    catch(exception) {
        console.log(exception);
    }
    
    const worksheet = workbook.getWorksheet(1);
    const rows = worksheet.getRows(2, 500);

    const output: IData[] = [];
    let index = 0;
    rows?.forEach(row => {
        const datetime = row.getCell(1).value;
        const weight = row.getCell(3).value;
        const mark = row.getCell(9).value;

        if(datetime && weight) {
            console.log(datetime, weight)
            output.push({
                index: index++,
                datetime,
                weight,
                mark: mark ? mark : null
            } as IData)
        }

        let data = JSON.stringify(output);
        fs.writeFileSync('output.json', data);
    })
    
    
}

load();