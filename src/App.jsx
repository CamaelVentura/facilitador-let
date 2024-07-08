import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import { useState } from 'react'

function App() {
  const [fileData, setFileData] = useState([]);
  const [fileInformations, setFileInformations] = useState({});

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    console.log(file)
    const reader = new FileReader();

    reader.onload = (event) => {
      console.log(event.target, file.name)
      const workbook = XLSX.read(event.target.result, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(sheet);
      setFileData(sheetData);
      setFileInformations(oldInformations => ({
        ...oldInformations,
        fileName: file.name,
        sheetName: sheetName
      }));
    };

    reader.readAsBinaryString(file);
  }

  const handleResponsable = () => {
    const filtered = fileData.map(legalTerm => {
      if (!legalTerm['Fonte'] || legalTerm['Advogado']) {
        return legalTerm;
      }

      if (legalTerm['Fonte'].toLowerCase().includes('custas')) {
        legalTerm['Advogado'] = 'Camael';
      }

      return legalTerm;
    });

    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(filtered, {
      header: [
        'Inicio', 'Fim', 'Assunto', 'Título',
        'Conclusão', 'Finalizado', 'Area', 'Advogado',
        'Tipo', 'Fatalissimo', 'Observação', 'Status',
        'Fonte'
      ]
    });

    console.log(fileInformations)
    XLSX.utils.book_append_sheet(workbook, sheet, fileInformations.sheetName);

    const blob = XLSX.writeFile(workbook, fileInformations.fileName);

    saveAs(blob, fileInformations.fileName);
    console.log(filtered)
  }

  return (
    <>
      <input type='file' onChange={handleFileUpload} />
      <button type='button' onClick={handleResponsable}>
        Faça a mágica
      </button>
    </>
  )
}

export default App
