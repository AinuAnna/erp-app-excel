import React from 'react';

interface ExcelReaderSheet2Props {
    tables: { [key: string]: { data: any[][], startIndex: number, endIndex: number } };
    tableNumbersToDisplay?: string[];
}

const ExcelReaderSheet2: React.FC<ExcelReaderSheet2Props> = ({ tables, tableNumbersToDisplay }) => {

    interface TableStyles {
        table: React.CSSProperties;
        th: React.CSSProperties;
        td: React.CSSProperties;
    }

    const tableStyles: TableStyles = {
        table: {
            borderCollapse: 'collapse',
            width: '100%',
            marginBottom: '20px',
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

    const tablesToDisplay = tableNumbersToDisplay ?
        Object.entries(tables).filter(([tableNumber]) => tableNumbersToDisplay.includes(tableNumber)) :
        Object.entries(tables);

    return (
        <div>
            <h1>RESULT</h1>
            {tablesToDisplay.map(([tableNumber, tableData]) => (
                <div key={tableNumber}>
                    <h2>{tableNumber}</h2>
                    <table style={tableStyles.table}>
                        <tbody>
                        {tableData.data.map((row, rowIndex) => (
                            <tr key={rowIndex}>
                                {row.map((cell: any, cellIndex: number) => {
                                    const cellValue = (typeof cell === 'object' && cell !== null && 'v' in cell) ? cell.v : cell;
                                    return (
                                        <td key={cellIndex} style={tableStyles.td}>{cellValue}</td>
                                    );
                                })}
                            </tr>
                        ))}
                        </tbody>
                    </table>
                </div>
            ))}
        </div>
    );
};

export default ExcelReaderSheet2;