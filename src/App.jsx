import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import { useState } from 'react'

function App() {
  const [fileData, setFileData] = useState([]);
  const [fileInformations, setFileInformations] = useState({});

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (event) => {
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

      console.log(legalTerm)
      if (
        (legalTerm['Assunto']
          && typeof legalTerm['Assunto'] === 'string'
          && legalTerm['Assunto'].toUpperCase().includes('AJUIZAR AÇÃO')
        ) || (
          legalTerm['Fonte'].toUpperCase().includes('NPJ')
          && !legalTerm['Fonte'].toUpperCase().includes('CNPJ')
        )
      ) {
        legalTerm['Advogado'] = 'LETÍCIA';
      } else if (
        legalTerm['Fonte'].toUpperCase().includes('ALTINOPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('ALTINÓPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('AMERICO BRASILIENSE')
        || legalTerm['Fonte'].toUpperCase().includes('AMÉRICO BRASILIENSE')
        || legalTerm['Fonte'].toUpperCase().includes('ANDRADINA')
        || legalTerm['Fonte'].toUpperCase().includes('AURIFLAMA')
        || legalTerm['Fonte'].toUpperCase().includes('BARUERI')
        || legalTerm['Fonte'].toUpperCase().includes('BILAC')
        || legalTerm['Fonte'].toUpperCase().includes('BORBOREMA')
        || legalTerm['Fonte'].toUpperCase().includes('BURITAMA')
        || legalTerm['Fonte'].toUpperCase().includes('CACAPAVA')
        || legalTerm['Fonte'].toUpperCase().includes('CAÇAPAVA')
        || legalTerm['Fonte'].toUpperCase().includes('CACHOEIRA PAULISTA')
        || legalTerm['Fonte'].toUpperCase().includes('CAFELANDIA')
        || legalTerm['Fonte'].toUpperCase().includes('CAFELÂNDIA')
        || legalTerm['Fonte'].toUpperCase().includes('CAIEIRAS')
        || legalTerm['Fonte'].toUpperCase().includes('CAMPOS DO JORDAO')
        || legalTerm['Fonte'].toUpperCase().includes('CAMPOS DO JORDÃO')
        || legalTerm['Fonte'].toUpperCase().includes('CANDIDO MOTA')
        || legalTerm['Fonte'].toUpperCase().includes('CÂNDIDO MOTA')
        || legalTerm['Fonte'].toUpperCase().includes('CARAPICUIBA')
        || legalTerm['Fonte'].toUpperCase().includes('CARAPICUÍBA')
        || (
          legalTerm['Fonte'].toUpperCase().includes('CARDOSO')
          && !legalTerm['Fonte'].toUpperCase().includes('FERNANDA YANNA MOREIRA CARDOSO')
        )
        || legalTerm['Fonte'].toUpperCase().includes('COSMOPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('COSMÓPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('COTIA')
        || legalTerm['Fonte'].toUpperCase().includes('DESCALVADO')
        || legalTerm['Fonte'].toUpperCase().includes('DIADEMA')
        || legalTerm['Fonte'].toUpperCase().includes('DOIS CORREGOS')
        || legalTerm['Fonte'].toUpperCase().includes('DUARTINA')
        || legalTerm['Fonte'].toUpperCase().includes('EMBU')
        || legalTerm['Fonte'].toUpperCase().includes('EMBU DAS ARTES')
        || legalTerm['Fonte'].toUpperCase().includes('EMBU-GUAÇU')
        || legalTerm['Fonte'].toUpperCase().includes("ESTRELA D'OESTE")
        || legalTerm['Fonte'].toUpperCase().includes('FRANCO DA ROCHA')
        || legalTerm['Fonte'].toUpperCase().includes('GENERAL SALGADO')
        || legalTerm['Fonte'].toUpperCase().includes('GETULINA')
        || legalTerm['Fonte'].toUpperCase().includes('GUARA')
        || legalTerm['Fonte'].toUpperCase().includes('GUARÁ')
        || legalTerm['Fonte'].toUpperCase().includes('GUARATINGUETA')
        || legalTerm['Fonte'].toUpperCase().includes('GUARATINGUETÁ')
        || legalTerm['Fonte'].toUpperCase().includes('IBATE')
        || legalTerm['Fonte'].toUpperCase().includes('IBATÉ')
        || legalTerm['Fonte'].toUpperCase().includes('IEPE')
        || legalTerm['Fonte'].toUpperCase().includes('IEPÊ')
        || legalTerm['Fonte'].toUpperCase().includes('LAPA')
        || legalTerm['Fonte'].toUpperCase().includes('PINHEIROS')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 11ª')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 16ª')
      ) {
        legalTerm['Advogado'] = 'CAIO';
      } else if (
        legalTerm['Fonte'].toUpperCase().includes('ORLANDIA')
        || legalTerm['Fonte'].toUpperCase().includes('ORLÂNDIA')
        || legalTerm['Fonte'].toUpperCase().includes('PALMITAL')
        || legalTerm['Fonte'].toUpperCase().includes('SAO JOSE DO RIO PRETO')
        || legalTerm['Fonte'].toUpperCase().includes('SÃO JOSE DO RIO PRETO')
        || legalTerm['Fonte'].toUpperCase().includes('SAO JOSÉ DO RIO PRETO')
        || legalTerm['Fonte'].toUpperCase().includes('SÃO JOSÉ DO RIO PRETO')
        || legalTerm['Fonte'].toUpperCase().includes('SAO LUIS DO PARAITINGA')
        || legalTerm['Fonte'].toUpperCase().includes('SÃO LUIS DO PARAITINGA')
        || legalTerm['Fonte'].toUpperCase().includes('SAO LUÍS DO PARAITINGA')
        || legalTerm['Fonte'].toUpperCase().includes('SÃO LUÍS DO PARAITINGA')
        || legalTerm['Fonte'].toUpperCase().includes('SERRANA')
        || legalTerm['Fonte'].toUpperCase().includes('TANABI')
        || legalTerm['Fonte'].toUpperCase().includes('TAUBATE')
        || legalTerm['Fonte'].toUpperCase().includes('TAUBATÉ')
        || legalTerm['Fonte'].toUpperCase().includes('TREMEMBE')
        || legalTerm['Fonte'].toUpperCase().includes('TREMEMBÉ')
        || legalTerm['Fonte'].toUpperCase().includes('TUPA')
        || legalTerm['Fonte'].toUpperCase().includes('TUPÃ')
        || legalTerm['Fonte'].toUpperCase().includes('SAO BERNARDO DO CAMPO')
        || legalTerm['Fonte'].toUpperCase().includes('SÃO BERNARDO DO CAMPO')
        || legalTerm['Fonte'].toUpperCase().includes('SANTO AMARO')
        || legalTerm['Fonte'].toUpperCase().includes('BUTANTA')
        || legalTerm['Fonte'].toUpperCase().includes('BUTANTÃ')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 21ª')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 26ª')
      ) {
        legalTerm['Advogado'] = 'GISELE';
      } else if (
        legalTerm['Fonte'].toUpperCase().includes('ITAPECERICA DA SERRA')
        || legalTerm['Fonte'].toUpperCase().includes('ITAPEVI')
        || legalTerm['Fonte'].toUpperCase().includes('ITUVERAVA')
        || legalTerm['Fonte'].toUpperCase().includes('JACAREI')
        || legalTerm['Fonte'].toUpperCase().includes('JACAREÍ')
        || legalTerm['Fonte'].toUpperCase().includes('JANDIRA')
        || legalTerm['Fonte'].toUpperCase().includes('JARDINOPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('JARDINÓPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('JOSE BONIFACIO')
        || legalTerm['Fonte'].toUpperCase().includes('JOSÉ BONIFACIO')
        || legalTerm['Fonte'].toUpperCase().includes('JOSE BONIFÁCIO')
        || legalTerm['Fonte'].toUpperCase().includes('JOSÉ BONIFÁCIO')
        || legalTerm['Fonte'].toUpperCase().includes('JUNQUEIROPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('JUNQUEIRÓPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('LORENA')
        || legalTerm['Fonte'].toUpperCase().includes('LUCELIA')
        || legalTerm['Fonte'].toUpperCase().includes('LUCÉLIA')
        || legalTerm['Fonte'].toUpperCase().includes('MACATUBA')
        || legalTerm['Fonte'].toUpperCase().includes('MAIRIPORA')
        || legalTerm['Fonte'].toUpperCase().includes('MAIRIPORÃ')
        || legalTerm['Fonte'].toUpperCase().includes('MARACAI')
        || legalTerm['Fonte'].toUpperCase().includes('MARACAÍ')
        || legalTerm['Fonte'].toUpperCase().includes('MARTINOPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('MARTINÓPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('MAUA')
        || legalTerm['Fonte'].toUpperCase().includes('MAUÁ')
        || legalTerm['Fonte'].toUpperCase().includes('MIGUELOPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('MIGUELÓPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('MIRACATU')
        || legalTerm['Fonte'].toUpperCase().includes('MIRANDOPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('MIRANDÓPOLIS')
        || legalTerm['Fonte'].toUpperCase().includes('MIRANTE DO PARANAPANEMA')
        || legalTerm['Fonte'].toUpperCase().includes('MONTE APRAZIVEL')
        || legalTerm['Fonte'].toUpperCase().includes('MONTE APRAZÍVEL')
        || legalTerm['Fonte'].toUpperCase().includes('NEVES PAULISTA')
        || legalTerm['Fonte'].toUpperCase().includes('NOVA GRANADA')
        || legalTerm['Fonte'].toUpperCase().includes('NUPORANGA')
        || legalTerm['Fonte'].toUpperCase().includes('OLIMPIA')
        || legalTerm['Fonte'].toUpperCase().includes('OLÍMPIA')
        || legalTerm['Fonte'].toUpperCase().includes('OUROESTE')
        || legalTerm['Fonte'].toUpperCase().includes("PALMEIRA D'OESTE")
        || legalTerm['Fonte'].toUpperCase().includes('IPAUSSU')
        || legalTerm['Fonte'].toUpperCase().includes('JABAQUARA')
        || legalTerm['Fonte'].toUpperCase().includes('PENHA DE FRANCA')
        || legalTerm['Fonte'].toUpperCase().includes('PENHA DE FRANÇA')
        || legalTerm['Fonte'].toUpperCase().includes('PENHA')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 1ª')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 6ª')
      ) {
        legalTerm['Advogado'] = 'ISABELA';
      } else if (
        (
          legalTerm['Fonte'].toUpperCase().includes('SANTANA')
          && !legalTerm['Fonte'].toUpperCase().includes('LUCAS SOUZA SANTANA')
        )
        || legalTerm['Fonte'].toUpperCase().includes('VILA PRUDENTE')
        || legalTerm['Fonte'].toUpperCase().includes('TATUAPE')
        || legalTerm['Fonte'].toUpperCase().includes('TATUAPÉ')
        || legalTerm['Fonte'].toUpperCase().includes('IPIRANGA')
        || legalTerm['Fonte'].toUpperCase().includes('GUARULHOS')
        || legalTerm['Fonte'].toUpperCase().includes('NOSSA SENHORA DO O')
        || legalTerm['Fonte'].toUpperCase().includes('NOSSA SENHORA DO Ó')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 31ª')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 36ª')
        || legalTerm['Fonte'].toUpperCase().includes('UPJ 41ª')
      ) {
        legalTerm['Advogado'] = 'FERNANDA';
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
