import React, { useState } from 'react';
import * as XLSX from 'xlsx';

interface ExcelEditorProps {
    data: any[][];
    setData: React.Dispatch<React.SetStateAction<any[][]>>;
    inputData: { [key: string]: string };
    setInputData: React.Dispatch<React.SetStateAction<{ [key: string]: string }>>;
    workbook: XLSX.WorkBook | null;
    fileName: string;
    handleSave: () => void;
}

const ExcelEditor: React.FC<ExcelEditorProps> = ({ data, setData, inputData, setInputData, workbook, fileName, handleSave }) => {

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const { name, value } = e.target;
        setInputData({ ...inputData, [name]: value });
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