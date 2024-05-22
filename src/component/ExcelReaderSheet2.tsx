import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const ExcelReaderSheet2: React.FC = () => {
    const [data, setData] = useState<any[]>([]);

    const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                const binaryStr = e.target?.result;
                if (binaryStr) {
                    const workbook = XLSX.read(binaryStr, { type: 'binary' });
                    const sheetName = workbook.SheetNames[1];
                    if (sheetName) {
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                        setData(jsonData);
                    } else {
                        console.error("Второй лист не найден");
                    }
                }
            };
            reader.readAsBinaryString(file);
        }
    };

    interface TableStyles {
        table: React.CSSProperties;
        th: React.CSSProperties;
        td: React.CSSProperties;
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
    };

    return (
        <div>
            <h1>Excel Reader - Sheet 2</h1>
            <input type="file" onChange={handleFileUpload} />
            <table style={tableStyles.table}>
                <tbody>
                    {data.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                            {row.map((cell: any, cellIndex: number) => (
                                <td key={cellIndex} style={tableStyles.td}>{cell}</td>
                            ))}
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};

export default ExcelReaderSheet2;
