import { equal } from 'assert';
import { Excel } from "./fable/src/Fable.exceljs/ExcelJs.js";
import { TestWorkbooks_Workbook1 } from "./fable/Xlsx.Tests.js";

describe('Mocha native', function () {
    describe('Native Mocha running', function () {
        it('should return -1 when the value is not present', function () {
            equal([1, 2, 3].indexOf(4), -1);
        });
    });
    describe('IO Read', function () {
        it('read from path', async () => { 
            const sheetname = "Tabelle1";
            const workbook = new Excel.Workbook();
            await workbook.xlsx.readFile("./tests/Fable.exceljs.JsNativeTests/MinimalTest.xlsx");
            const worksheet = workbook.getWorksheet(sheetname);
            equal(worksheet.name, sheetname);
        });
    });
    describe('IO roundtest', function () {
        it('write to path js native workbook', async () => { 
            const path = "./tests/Fable.exceljs.JsNativeTests/WriteTest.xlsx";
            const sheetname = "NextWorksheet";
            const wb = new Excel.Workbook();
            const ws = wb.addWorksheet(sheetname);
            let date = new Date("2011-10-10T14:48:00");
            ws.addRow([3, 'Sam', date]);
            const rowValues = [];
            rowValues[1] = 4;
            rowValues[5] = 'Kyle';
            rowValues[9] = date;
            ws.addRow(rowValues);
            await wb.xlsx.writeFile(path);
            const readback_wb = new Excel.Workbook();
            await readback_wb.xlsx.readFile(path);
            const readback_ws = readback_wb.getWorksheet(sheetname);
            equal(readback_ws.name, sheetname);
            equal(readback_ws.getRow(1).values[1], 3);
            equal(readback_ws.getRow(1).values[2], 'Sam');
            equal(readback_ws.getRow(1).values[3].getTime(), date.getTime());
            equal(readback_ws.getRow(2).values[1], 4);
            equal(readback_ws.getRow(2).values[5], 'Kyle');
            equal(readback_ws.getRow(2).values[9].getTime(), date.getTime());
        });
        it('write to path fable workbook', async () => { 
            const path = "./tests/Fable.exceljs.JsNativeTests/WriteTestFable.xlsx";
            const sheetname = "MySheet1";
            await TestWorkbooks_Workbook1.xlsx.writeFile(path);
            const readback_wb = new Excel.Workbook();
            await readback_wb.xlsx.readFile(path);
            const readback_ws = readback_wb.getWorksheet(sheetname);
            equal(readback_ws.name, sheetname);
        });
    });
    describe('IO read OpenXml', function () {
        it('read', async () => {
            const path = "./tests/Fable.exceljs.JsNativeTests/OpenXmlTest.xlsx";
            const sheetname = "Hello";
            const workbook = new Excel.Workbook();
            await workbook.xlsx.readFile(path);
            const worksheet = workbook.getWorksheet(sheetname);
            equal(worksheet.name, sheetname);
        });
        it('should return undefined', async () => {
            const path = "./tests/Fable.exceljs.JsNativeTests/OpenXmlTest.xlsx";
            const sheetname = "Hello2";
            const workbook = new Excel.Workbook();
            await workbook.xlsx.readFile(path);
            const worksheet = workbook.getWorksheet(sheetname);
            equal(worksheet, undefined)
        });
    });
});