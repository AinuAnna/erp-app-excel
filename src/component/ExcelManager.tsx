import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import ExcelEditor from './ExcelEditor';
import ExcelReaderSheet2 from './ExcelReaderSheet2';

const ExcelManager: React.FC = () => {
    const [data, setData] = useState<any[][]>([]);
    const [inputData, setInputData] = useState<{ [key: string]: string }>({});
    const [sheet2Data, setSheet2Data] = useState<any[][]>([]);
    const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
    const [fileName, setFileName] = useState<string>('');
    const [tables, setTables] = useState<{ [key: string]: { data: any[][], startIndex: number, endIndex: number } }>({});

    const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (file) {
            setFileName(file.name);

            const reader = new FileReader();
            reader.onload = (e) => {
                const binaryStr = e.target?.result;
                if (binaryStr) {
                    const wb = XLSX.read(binaryStr, { type: 'binary' });
                    setWorkbook(wb);

                    const sheet1Name = wb.SheetNames[0];
                    const sheet1 = wb.Sheets[sheet1Name];
                    let sheet1Data: any[][] = XLSX.utils.sheet_to_json(sheet1, { header: 1, blankrows: true, defval: '' }) as any[][];
                    setData(sheet1Data);

                    const sheet2Name = wb.SheetNames[1];
                    const sheet2 = wb.Sheets[sheet2Name];
                    const sheet2Data: any[][] = XLSX.utils.sheet_to_json(sheet2, { header: 1, blankrows: true, defval: '' }) as any[][];
                    setSheet2Data(sheet2Data);

                    let currentTable: string | null = null;

                    sheet2Data.forEach((row, rowIndex) => {
                        const firstCellValue = row[0];
                        if (typeof firstCellValue === 'number' && Number.isInteger(firstCellValue)) {
                            if (currentTable) {
                                tables[currentTable].endIndex = rowIndex - 1;
                            }
                            currentTable = `Table ${firstCellValue}`;
                            tables[currentTable] = { data: [], startIndex: rowIndex + 2, endIndex: rowIndex + 2 };
                        }
                        if (currentTable) {
                            tables[currentTable].data.push(row);
                        }
                    });

                    if (currentTable) {
                        tables[currentTable].endIndex = sheet2Data.length - 1;
                    }

                    setTables(tables);

                    const initialInputData: any = sheet1Data.reduce((acc: any, row: any, rowIndex: number) => {
                        sheet1Data[0].forEach((col: string, colIndex: number) => {
                            acc[`${col}-${rowIndex}`] = row[colIndex] || '';
                        });
                        return acc;
                    }, {});
                    setInputData(initialInputData);
                }
            };
            reader.readAsBinaryString(file);
        }
    };

    const handleSave = () => {
        if (!workbook) return;

        const sheet1Name = workbook.SheetNames[0];
        const originalSheet1 = workbook.Sheets[sheet1Name];
        const originalRange = XLSX.utils.decode_range(originalSheet1['!ref'] as string);
        const updatedSheetData = Array.from({ length: originalRange.e.r + 1 }, (_, rowIndex) =>
            Array.from({ length: originalRange.e.c + 1 }, (_, colIndex) => {
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
                return originalSheet1[cellAddress] ? originalSheet1[cellAddress].v : '';
            })
        );

        for (let key in inputData) {
            const [col, rowIndex] = key.split('-');
            const colIndex = data[0].indexOf(col);
            const cellValue = inputData[key];
            const numericValue = parseFloat(cellValue);
            if (!isNaN(numericValue)) {
                updatedSheetData[parseInt(rowIndex, 10) + 1][colIndex] = numericValue;
            } else {
                updatedSheetData[parseInt(rowIndex, 10) + 1][colIndex] = cellValue;
            }
        }

        const updatedSheet1 = XLSX.utils.aoa_to_sheet(updatedSheetData);
        workbook.Sheets[sheet1Name] = updatedSheet1;
        let updatedData: any[][] = XLSX.utils.sheet_to_json(updatedSheet1, { header: 1, blankrows: false, defval: '' }) as any[][];
        if (updatedData.length > 0 && updatedData[0].every((cell: any) => cell === '')) {
            updatedData = updatedData.slice(1);
        }
        setData(updatedData);

        const sheet2Name = workbook.SheetNames[1];
        const sheet2 = workbook.Sheets[sheet2Name];
        let sheet2D: any[][] = XLSX.utils.sheet_to_json(sheet2, { header: 1, blankrows: true, defval: '' }) as any[][];
        if (sheet2D.length > 0 && sheet2D[0].every((cell: any) => cell === '')) {
            sheet2D = sheet2D.slice(1);
        }
        setSheet2Data(sheet2D);

        recalculateSheet2(updatedData, sheet2D);

        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
        const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
    };

    const s2ab = (s: string) => {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    };

    const recalculateSheet2 = (data1: any[][], data2: any[][]) => {
        if (!workbook) return;

        const sheet2Name = workbook.SheetNames[1];
        const sheet2 = workbook.Sheets[sheet2Name];

        // Считываем формулы из второго листа
        const formulas: { [cell: string]: string } = {};
        for (const cell in sheet2) {
            if (sheet2.hasOwnProperty(cell) && sheet2[cell].f) {
                formulas[cell] = sheet2[cell].f;
            }
        }

        const calculatedValues = {};

        const processFormula = (formula: string, data1: any[][], data2: any[][], calculatedValues: { [x: string]: any; }) => {
            const formulaParts = formula.split(/([-+*/()])/);

            const updatedFormulaParts = formulaParts.map(part => {
                if (part.match(/(([^\s]+)?!)?([A-Z]+\d+)%?/)) {
                    return part.replace(/(([^\s]+)?!)?([A-Z]+\d+)%?/g, (match, sheetName, _, cellRef) => {
                        const ish = sheetName === 'ИсхДанные!';
                        const sheetData = ish ? data1 : data2;
                        const [col, row] = cellRef.match(/([A-Z]+)(\d+)/).slice(1);
                        const colIndex = XLSX.utils.decode_col(col);
                        const rowIndex = parseInt(row) - 2;
                        const cellKey = `${col}${row}`;

                        // Рекурсивно вычисляем значение, если оно еще не вычислено
                        if (!(cellKey in calculatedValues)) {
                            if (formulas[cellKey]) {
                                calculatedValues[cellKey] = processFormula(formulas[cellKey], data1, data2, calculatedValues);
                            } else {
                                let value = sheetData[rowIndex][colIndex];
                                if (typeof value === 'object' && value !== null && 'v' in value) {
                                    value = value.v;
                                }
                                calculatedValues[cellKey] = value || 0;
                            }
                        }

                        let value = calculatedValues[cellKey];

                        if (match.endsWith('%')) {
                            value = value * 0.01;
                        }

                        return value;
                    });
                } else {
                    return part;
                }
            });

            let updatedFormula = updatedFormulaParts.join('');
            updatedFormula = processPercentages(updatedFormula);

            try {
                const jsFunction = new Function(`return ${updatedFormula}`);
                return jsFunction();
            } catch (error) {
                console.error(`Error calculating formula "${formula}":`, error);
                return null;
            }
        };

        for (const cell in formulas) {
            // @ts-ignore
            const [col, row] = cell.match(/([A-Z]+)(\d+)/).slice(1);
            const colIndex = XLSX.utils.decode_col(col);
            const rowIndex = parseInt(row) - 2;

            // @ts-ignore
            const newValue = processFormula(formulas[cell], data1, data2, calculatedValues);
            if (newValue !== null) {
                // @ts-ignore
                calculatedValues[`${col}${row}`] = newValue;

                const tableEntry = Object.entries(tables).find(([_, table]) => rowIndex >= table.startIndex && rowIndex <= table.endIndex);
                if (tableEntry) {
                    const [tableNumber, tableData] = tableEntry;
                    const tableRowIndex = rowIndex - tableData.startIndex + 2;
                    tables[tableNumber].data[tableRowIndex][colIndex] = { v: newValue };
                }
            }
        }

        setTables(tables);
    };

    const processPercentages = (str: string) => str.replace(/(\d+(\.\d+)?|\([^()]+\))%/g, (_: string, expr: string): string => {
        const evaluated = new Function(`return ${expr}`)();
        return (evaluated * 0.01).toString();
    });

    return (
        <div>
            <h1>ERP-SYSTEM</h1>
            <label>
                Upload
                <input
                    type="file"
                    className="custom-file-input"
                    onChange={handleFileUpload}
                />
            </label>
            <span className="file-name">{fileName}</span>
            <ExcelEditor
                data={data}
                setData={setData}
                inputData={inputData}
                setInputData={setInputData}
                workbook={workbook}
                fileName={fileName}
                handleSave={handleSave}
            />
            <ExcelReaderSheet2 tables={tables} tableNumbersToDisplay={['Table 1', 'Table 2', 'Table 3', 'Table 4', 'Table 5', 'Table 6', 'Table 7', 'Table 8', 'Table 9', 'Table 10', 'Table 11', 'Table 12', 'Table 13', 'Table 14', 'Table 15']} />
        </div>
    );
};

export default ExcelManager;
