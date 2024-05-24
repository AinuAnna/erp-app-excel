import React from 'react';

interface ExcelReaderSheet2Props {
    sheet2Data: any[][];
}

const ExcelReaderSheet2: React.FC<ExcelReaderSheet2Props> = ({ sheet2Data }) => {

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
            <h1>RESULT</h1>
            <table style={tableStyles.table}>
                <tbody>
                    {sheet2Data.map((row, rowIndex) => (
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