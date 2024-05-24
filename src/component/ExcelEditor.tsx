import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const ExcelEditor: React.FC = () => {
    const [data, setData] = useState<any[][]>([]);
    const [inputData, setInputData] = useState<{ [key: string]: string }>(() => {
        const initialInputData: { [key: string]: string } = {};
        data.forEach((row, rowIndex) => {
            row.forEach((col, colIndex) => {
                initialInputData[`${col}-${rowIndex}`] = col || '';
            });
        });
        return initialInputData;
    });
    const [sheet2Data, setSheet2Data] = useState<any[][]>([]);
    const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
    const [fileName, setFileName] = useState<string>('');

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
                    let sheet1Data: any[][] = XLSX.utils.sheet_to_json(sheet1, { header: 1, blankrows: false }) as any[][];

                    if (sheet1Data.length > 0 && sheet1Data[0].every((cell: any) => cell === '')) {
                        sheet1Data = sheet1Data.slice(1);
                    }

                    setData(sheet1Data);

                    const sheet2Name = wb.SheetNames[1];
                    const sheet2 = wb.Sheets[sheet2Name];
                    const sheet2Data: any[][] = XLSX.utils.sheet_to_json(sheet2, { header: 1, blankrows: true }) as any[][];
                    setSheet2Data(sheet2Data);

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

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const { name, value } = e.target;
        setInputData({ ...inputData, [name]: value });
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

        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
        const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);

        const fileReader = new FileReader();
        fileReader.onload = (e) => {
            const binaryStr = e.target?.result;
            if (binaryStr) {
                const updatedWb = XLSX.read(binaryStr, { type: 'binary' });
                const sheet2Name = updatedWb.SheetNames[1];
                const sheet2 = updatedWb.Sheets[sheet2Name];
                const updatedSheet2Data = XLSX.utils.sheet_to_json(sheet2, { header: 1, blankrows: true });
                setSheet2Data(updatedSheet2Data as any[][]);
            }
        };
        fileReader.readAsBinaryString(blob);
    };

    const s2ab = (s: string) => {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    };

    interface TableStyles {
        table: React.CSSProperties;
        th: React.CSSProperties;
        td: React.CSSProperties;
        input: React.CSSProperties;
    }

    const tableStyles: TableStyles = {
        table: {
            borderCollapse: 'collapse',
            width: '100%',
        },
        th: {
            border: '1px solid black',
            padding: '8px',
        },
        td: {
            border: '1px solid black',
            padding: '8px',
        },
        input: {
            width: 'calc(100% - 16px)',
            boxSizing: 'border-box',
        },
    };

    return (
        <div>
            <input type="file" onChange={handleFileUpload} />
            <table style={tableStyles.table}>
                <tbody>
                    {data.length > 0 && (
                        <table style={tableStyles.table}>
                            <thead>
                                <tr>
                                    {data[0].map((header: string, index: number) => (
                                        <th key={index} style={tableStyles.th}>
                                            {header}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody>
                                {data.map((row: any[], rowIndex: number) => (
                                    <tr key={rowIndex}>
                                        {data[0].map((header: string, colIndex: number) => (
                                            <td key={colIndex} style={tableStyles.td}>
                                                {typeof row[colIndex] === 'number' ? (
                                                    <input
                                                        type="number"
                                                        name={`${data[0][colIndex]}-${rowIndex}`}
                                                        value={inputData[`${data[0][colIndex]}-${rowIndex}`]}
                                                        onChange={handleChange}
                                                        style={tableStyles.input}
                                                    />
                                                ) : (
                                                    row[colIndex]
                                                )}
                                            </td>
                                        ))}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    )}
                </tbody>
            </table>
            <button onClick={handleSave}>Save</button>
        </div>
    );
};

export default ExcelEditor;
