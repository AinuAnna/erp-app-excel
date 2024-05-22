import React from 'react';
import './App.css';
import ExcelEditor from './component/ExcelEditor';
import ExcelReaderSheet2 from './component/ExcelReaderSheet2';

const App: React.FC = () => {
  return (
    <div className="App">
      <ExcelEditor />
      <ExcelReaderSheet2 />
    </div>
  );
};

export default App;
