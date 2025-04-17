import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

const ExcelProcessor = () => {
  const [file, setFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState('');
  const [combinazioni, setCombinazioni] = useState([]);
  const [datiProcessati, setDatiProcessati] = useState({});
  const [combSelezionata, setCombSelezionata] = useState('');
  const [classiMerch, setClassiMerch] = useState([]);
  const [classeMerchSelezionata, setClasseMerchSelezionata] = useState('');
  const [quantitaProposta, setQuantitaProposta] = useState('');
  
  // Stato per la gestione della mappatura delle colonne
  const [availableColumns, setAvailableColumns] = useState([]);
  const [showColumnMapping, setShowColumnMapping] = useState(false);
  const [columnMapping, setColumnMapping] = useState({
    Gender: '',
    Line: '',
    MerchandisingClass: '',
    SizeCode: '',
    OrderQty: '',
    SoldQty: ''
  });

  // Funzione per ordinare correttamente le taglie
  const ordenarTallas = (sizes) => {
    // Separare le taglie numeriche dalle alfanumeriche
    const numericas = [];
    const alfanumericas = [];

    sizes.forEach(size => {
      if (!isNaN(parseFloat(size))) {
        numericas.push(size);
      } else {
        alfanumericas.push(size);
      }
    });

    // Ordina le taglie numeriche
    numericas.sort((a, b) => parseFloat(a) - parseFloat(b));

    // Definisci l'ordine delle taglie alfanumeriche
    const ordenTallas = {
      "XXXS": 1, "XXS": 2, "XS": 3, "S": 4, "M": 5, "L": 6, "XL": 7, "XXL": 8, "XXXL": 9,
      "3XL": 10, "4XL": 11, "5XL": 12, "6XL": 13, 
      "ONE SIZE": 14, "U": 15, "OS": 16
    };
    
    // Funzione per ottenere il valore di ordinamento di una taglia
    const getValorOrden = (talla) => {
      const tallaMayusculas = talla.toUpperCase();
      if (tallaMayusculas in ordenTallas) {
        return ordenTallas[tallaMayusculas];
      }
      // Per taglie non riconosciute
      return 99; 
    };

    // Ordina le taglie alfanumeriche
    alfanumericas.sort((a, b) => getValorOrden(a) - getValorOrden(b));

    // Combina le taglie ordinate
    return [...alfanumericas, ...numericas];
  };

  // Gestisce il caricamento del file
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    setFile(file);
    setCombinazioni([]);
    setDatiProcessati({});
    setCombSelezionata('');
    setClassiMerch([]);
    setClasseMerchSelezionata('');
    setError('');
    
    if (file) {
      analyzeFile(file);
    }
  };

  // Analizza il file Excel per ottenere le colonne disponibili
  const analyzeFile = async (file) => {
    try {
      setProcessing(true);
      
      // Legge il file come ArrayBuffer
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);

      // Verifica che esista il foglio Database
      if (!workbook.SheetNames.includes('Database')) {
        throw new Error('Il foglio "Database" non è presente nel file Excel');
      }

      // Ottiene i dati dal foglio Database
      const databaseSheet = workbook.Sheets['Database'];
      const jsonData = XLSX.utils.sheet_to_json(databaseSheet);

      if (jsonData.length === 0) {
        throw new Error('Il foglio Database è vuoto');
      }

      // Estrae le colonne disponibili
      const firstRow = jsonData[0];
      const columns = Object.keys(firstRow);
      setAvailableColumns(columns);
      
      // Tenta di mappare automaticamente le colonne
      const newMapping = { ...columnMapping };
      
      // Funzione per trovare la migliore corrispondenza
      const findBestMatch = (columns, searchTerms) => {
        // Prima cerca corrispondenze esatte
        for (const term of searchTerms) {
          const exactMatch = columns.find(col => col.toLowerCase() === term.toLowerCase());
          if (exactMatch) return exactMatch;
        }
        
        // Poi cerca corrispondenze parziali
        for (const term of searchTerms) {
          const partialMatches = columns.filter(col => 
            col.toLowerCase().includes(term.toLowerCase())
          );
          if (partialMatches.length > 0) return partialMatches[0];
        }
        
        return '';
      };
      
      newMapping.Gender = findBestMatch(columns, ['Gender', 'Genere', 'Sesso']);
      newMapping.Line = findBestMatch(columns, ['Line', 'Linea']);
      newMapping.MerchandisingClass = findBestMatch(columns, ['Merchandising Class', 'Merch Class', 'Class', 'Classe']);
      newMapping.SizeCode = findBestMatch(columns, ['Size Code', 'Size', 'Taglia', 'Cod Taglia']);
      newMapping.OrderQty = findBestMatch(columns, ['ORDER QTY', 'Order Qty', 'Order Quantity', 'Qty Ordered', 'Quantità Ordinata']);
      newMapping.SoldQty = findBestMatch(columns, ['SOLD QTY', 'Sold Qty', 'Sold Quantity', 'Qty Sold', 'Quantità Venduta']);
      
      setColumnMapping(newMapping);
      
      // Mostra il pannello di mappatura se non siamo riusciti a mappare tutte le colonne
      const hasMissingMapping = Object.values(newMapping).some(value => !value);
      setShowColumnMapping(hasMissingMapping);
      
      setProcessing(false);
    } catch (err) {
      setError(`Errore durante l'analisi del file: ${err.message}`);
      setProcessing(false);
    }
  };

  // Gestisce il cambio di mapping delle colonne
  const handleColumnMappingChange = (columnKey, selectedColumn) => {
    setColumnMapping(prevMapping => ({
      ...prevMapping,
      [columnKey]: selectedColumn
    }));
  };

  // Processa il file Excel
  const processFile = async () => {
    if (!file) {
      setError('Seleziona un file Excel');
      return;
    }
    
    // Verifica che tutte le colonne siano mappate
    const missingMappings = Object.entries(columnMapping)
      .filter(([_, value]) => !value)
      .map(([key, _]) => key);
    
    if (missingMappings.length > 0) {
      setError(`Devi selezionare le colonne per: ${missingMappings.join(', ')}`);
      setShowColumnMapping(true);
      return;
    }

    setProcessing(true);
    setError('');
    setShowColumnMapping(false);

    try {
      // Legge il file come ArrayBuffer
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);

      // Ottiene i dati dal foglio Database
      const databaseSheet = workbook.Sheets['Database'];
      const jsonData = XLSX.utils.sheet_to_json(databaseSheet);

      // Trova le combinazioni uniche di Gender e Line
      const combs = [];
      const uniqueCombSet = new Set();
      
      jsonData.forEach(row => {
        const gender = row[columnMapping.Gender] || '';
        const line = row[columnMapping.Line] || '';
        const comb = `${gender} - ${line}`;
        
        if (gender && line && !uniqueCombSet.has(comb)) {
          uniqueCombSet.add(comb);
          combs.push(comb);
        }
      });

      // Ordina le combinazioni
      combs.sort();
      setCombinazioni(combs);

      // Processa i dati per ogni combinazione
      const datiPerCombinazione = {};
      
      combs.forEach(comb => {
        const [gender, line] = comb.split(' - ');
        
        // Filtra i dati per questa combinazione
        const datiComb = jsonData.filter(row => 
          (row[columnMapping.Gender] === gender) && (row[columnMapping.Line] === line)
        );
        
        // Trova le classi merchandising uniche per questa combinazione
        const classiSet = new Set();
        datiComb.forEach(row => {
          const classe = row[columnMapping.MerchandisingClass];
          if (classe) classiSet.add(classe);
        });
        
        const classi = Array.from(classiSet).sort();
        
        // Per ogni classe, calcola i dati per la tabella
        const datiPerClasse = {};
        
        classi.forEach(classe => {
          // Filtra i dati per questa classe
          const datiClasse = datiComb.filter(row => row[columnMapping.MerchandisingClass] === classe);
          
          // Calcola ORDER QTY e SOLD QTY totali
          let orderQtyTotal = 0;
          let soldQtyTotal = 0;
          
          datiClasse.forEach(row => {
            const orderQty = Number(row[columnMapping.OrderQty] || 0);
            const soldQty = Number(row[columnMapping.SoldQty] || 0);
            
            orderQtyTotal += orderQty;
            soldQtyTotal += soldQty;
          });
          
          // Calcola i dati per dimensione
          const datiPerSize = {};
          
          datiClasse.forEach(row => {
            const size = row[columnMapping.SizeCode];
            if (!size) return;
            
            if (!datiPerSize[size]) {
              datiPerSize[size] = {
                orderQty: 0,
                soldQty: 0
              };
            }
            
            datiPerSize[size].orderQty += Number(row[columnMapping.OrderQty] || 0);
            datiPerSize[size].soldQty += Number(row[columnMapping.SoldQty] || 0);
          });
          
          // Calcola le percentuali
          Object.keys(datiPerSize).forEach(size => {
            const sizeData = datiPerSize[size];
            sizeData.sellOutPct = soldQtyTotal > 0 ? (sizeData.soldQty / soldQtyTotal) * 100 : 0;
            sizeData.sellThroughPct = sizeData.orderQty > 0 ? (sizeData.soldQty / sizeData.orderQty) * 100 : 0;
          });
          
          datiPerClasse[classe] = {
            orderQtyTotal,
            soldQtyTotal,
            datiPerSize
          };
        });
        
        datiPerCombinazione[comb] = {
          classi,
          datiPerClasse
        };
      });
      
      setDatiProcessati(datiPerCombinazione);
      setProcessing(false);
      
    } catch (err) {
      setError(`Errore durante l'elaborazione: ${err.message}`);
      setProcessing(false);
    }
  };

  // Gestisce il cambio di combinazione selezionata
  const handleChangeCombinazione = (e) => {
    const comb = e.target.value;
    setCombSelezionata(comb);
    
    if (comb && datiProcessati[comb]) {
      setClassiMerch(datiProcessati[comb].classi);
      setClasseMerchSelezionata('');
    } else {
      setClassiMerch([]);
      setClasseMerchSelezionata('');
    }
  };

  // Formatta il numero con la virgola come separatore decimale
  const formatNumber = (num, decimals = 2) => {
    return num.toFixed(decimals).replace('.', ',');
  };

  // Calcola la proposta split basata sulla percentuale di vendita
  // Usa un valore totale predefinito di 100 se non è specificato altro
  const calcolaPropostaSplit = (sellOutPct, quantitaTotale = 100) => {
    return (quantitaTotale * sellOutPct) / 100;
  };

  // Scarica il report Excel
  // Scarica il report Excel
  const downloadReport = () => {
    if (Object.keys(datiProcessati).length === 0) {
      setError('Prima elabora il file Excel');
      return;
    }

    try {
      // Crea un nuovo workbook
      const wb = XLSX.utils.book_new();
      
      // Per ogni combinazione Gender-Line, crea un singolo foglio
      const datiCombinazioni = Object.keys(datiProcessati);
      
      datiCombinazioni.forEach(comb => {
        const datiComb = datiProcessati[comb];
        
        // Limita il nome del foglio a 31 caratteri (limite Excel)
        const sheetName = comb.length > 31 ? comb.substring(0, 31) : comb;
        
        // Dati per questo foglio
        const sheetData = [];
        
        // Per ogni classe merchandising
        datiComb.classi.forEach(classe => {
          // Aggiungi il titolo della classe
          sheetData.push([`Merchandising Class: ${classe}`]);
          sheetData.push([]);
          
          // Intestazioni della tabella
          sheetData.push([
            'Size Code',
            'ORDER QTY',
            'SOLD QTY',
            'SELL-OUT % by Size',
            'S/T %',
            'Proposta Split'
          ]);
          
          const datiClasse = datiComb.datiPerClasse[classe];
          const { orderQtyTotal, soldQtyTotal, datiPerSize } = datiClasse;
          
          // Ottieni e ordina le taglie
          const sortedSizes = ordenarTallas(Object.keys(datiPerSize));
          
          // Input per la quantità proposta (predefinito 100)
          const quantitaPropostaCell = 100;
          
          // Calcola la posizione della cella con il totale per le formule
          const currentRowIdx = sheetData.length; // Indice riga corrente (per intestazioni tabella)
          const totalRowIdx = currentRowIdx + sortedSizes.length + 1; // Indice per la riga del totale
          
          // Righe per ogni taglia
          sortedSizes.forEach((size, idx) => {
            const rowIdx = currentRowIdx + idx + 1; // Indice di riga corrente
            const sizeData = datiPerSize[size];
            
            sheetData.push([
              size,
              sizeData.orderQty,
              sizeData.soldQty,
              { t: 'n', v: sizeData.sellOutPct / 100, z: '0,00%' }, // Formato percentuale
              { t: 'n', v: sizeData.sellThroughPct / 100, z: '0,00%' }, // Formato percentuale
              { f: `F${totalRowIdx}*D${rowIdx}` } // Formula Excel: TotaleQuantità * SellOutPercentuale
            ]);
          });
          
          // Totale
          sheetData.push([
            'TOTALE',
            orderQtyTotal,
            soldQtyTotal,
            { t: 'n', v: 1, z: '0%' }, // 100%
            { t: 'n', v: (soldQtyTotal > 0 && orderQtyTotal > 0 ? (soldQtyTotal / orderQtyTotal) : 0), z: '0,00%' },
            quantitaPropostaCell // Valore di default per la proposta totale
          ]);
          
          // Aggiungi righe vuote tra le classi
          sheetData.push([]);
          sheetData.push([]);
        });
        
        // Crea il foglio
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        
        // Aggiungi il foglio al workbook
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      });
      
      // Scarica il file
      XLSX.writeFile(wb, 'Report_Gender_Line.xlsx');
      
    } catch (err) {
      setError(`Errore durante la creazione del report: ${err.message}`);
    }
  };

  return (
    <div className="p-6 max-w-4xl mx-auto bg-white shadow-md rounded-lg">
      <h1 className="text-2xl font-bold mb-6 text-center">Excel Report Generator per Gender &amp; Line</h1>
      
      {/* Caricamento file */}
      <div className="mb-6">
        <label className="block text-gray-700 font-semibold mb-2">Seleziona file Excel:</label>
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          className="w-full border border-gray-300 p-2 rounded"
        />
      </div>
      
      {/* Mappatura delle colonne */}
      {showColumnMapping && availableColumns.length > 0 && (
        <div className="mb-6 p-4 border border-blue-300 bg-blue-50 rounded">
          <h2 className="text-lg font-semibold mb-3">Mappatura delle Colonne:</h2>
          <p className="mb-3">Seleziona a quale colonna del tuo file Excel corrisponde ciascun campo richiesto:</p>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-gray-700 font-semibold mb-1">Gender:</label>
              <select
                value={columnMapping.Gender}
                onChange={(e) => handleColumnMappingChange('Gender', e.target.value)}
                className="w-full border border-gray-300 p-2 rounded mb-3"
              >
                <option value="">Seleziona...</option>
                {availableColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>
            
            <div>
              <label className="block text-gray-700 font-semibold mb-1">Line:</label>
              <select
                value={columnMapping.Line}
                onChange={(e) => handleColumnMappingChange('Line', e.target.value)}
                className="w-full border border-gray-300 p-2 rounded mb-3"
              >
                <option value="">Seleziona...</option>
                {availableColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>
            
            <div>
              <label className="block text-gray-700 font-semibold mb-1">Merchandising Class:</label>
              <select
                value={columnMapping.MerchandisingClass}
                onChange={(e) => handleColumnMappingChange('MerchandisingClass', e.target.value)}
                className="w-full border border-gray-300 p-2 rounded mb-3"
              >
                <option value="">Seleziona...</option>
                {availableColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>
            
            <div>
              <label className="block text-gray-700 font-semibold mb-1">Size Code:</label>
              <select
                value={columnMapping.SizeCode}
                onChange={(e) => handleColumnMappingChange('SizeCode', e.target.value)}
                className="w-full border border-gray-300 p-2 rounded mb-3"
              >
                <option value="">Seleziona...</option>
                {availableColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>
            
            <div>
              <label className="block text-gray-700 font-semibold mb-1">ORDER QTY:</label>
              <select
                value={columnMapping.OrderQty}
                onChange={(e) => handleColumnMappingChange('OrderQty', e.target.value)}
                className="w-full border border-gray-300 p-2 rounded mb-3"
              >
                <option value="">Seleziona...</option>
                {availableColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>
            
            <div>
              <label className="block text-gray-700 font-semibold mb-1">SOLD QTY:</label>
              <select
                value={columnMapping.SoldQty}
                onChange={(e) => handleColumnMappingChange('SoldQty', e.target.value)}
                className="w-full border border-gray-300 p-2 rounded mb-3"
              >
                <option value="">Seleziona...</option>
                {availableColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>
          </div>
        </div>
      )}
      
      <div className="mb-6">
        <button
          onClick={processFile}
          disabled={!file || processing}
          className="bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700 disabled:bg-gray-400"
        >
          {processing ? 'Elaborazione in corso...' : 'Elabora File'}
        </button>
        
        {availableColumns.length > 0 && !showColumnMapping && (
          <button
            onClick={() => setShowColumnMapping(true)}
            className="ml-4 bg-gray-200 text-gray-800 px-4 py-2 rounded hover:bg-gray-300"
          >
            Configura Mappatura Colonne
          </button>
        )}
      </div>
      
      {error && (
        <div className="bg-red-100 border border-red-400 text-red-700 p-3 rounded mb-4">
          {error}
        </div>
      )}
      
      {combinazioni.length > 0 && (
        <div className="mb-6">
          <h2 className="text-xl font-semibold mb-3">Combinazioni Gender-Line trovate:</h2>
          <select
            value={combSelezionata}
            onChange={handleChangeCombinazione}
            className="w-full border border-gray-300 p-2 rounded mb-4"
          >
            <option value="">Seleziona una combinazione...</option>
            {combinazioni.map((comb, index) => (
              <option key={index} value={comb}>{comb}</option>
            ))}
          </select>
          
          {combSelezionata && classiMerch.length > 0 && (
            <div>
              <h3 className="text-lg font-semibold mb-2">Classi Merchandising:</h3>
              <select
                value={classeMerchSelezionata}
                onChange={(e) => setClasseMerchSelezionata(e.target.value)}
                className="w-full border border-gray-300 p-2 rounded mb-4"
              >
                <option value="">Seleziona una classe...</option>
                {classiMerch.map((classe, index) => (
                  <option key={index} value={classe}>{classe}</option>
                ))}
              </select>
            </div>
          )}
          
          {classeMerchSelezionata && (
            <div className="mb-6">
              <h3 className="text-lg font-semibold mb-2">Dettagli per {classeMerchSelezionata}:</h3>
              <div className="overflow-x-auto">
                <table className="min-w-full bg-white border border-gray-300">
                  <thead>
                    <tr className="bg-gray-100">
                      <th className="py-2 px-4 border-b">Size Code</th>
                      <th className="py-2 px-4 border-b">ORDER QTY</th>
                      <th className="py-2 px-4 border-b">SOLD QTY</th>
                      <th className="py-2 px-4 border-b">SELL-OUT % by Size</th>
                      <th className="py-2 px-4 border-b">S/T %</th>
                      <th className="py-2 px-4 border-b">Proposta Split</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ordenarTallas(Object.keys(datiProcessati[combSelezionata].datiPerClasse[classeMerchSelezionata].datiPerSize)).map((size, index) => {
                      const sizeData = datiProcessati[combSelezionata].datiPerClasse[classeMerchSelezionata].datiPerSize[size];
                      const propostaTotal = quantitaProposta || 100; // Default a 100 se non specificato
                      const split = (propostaTotal * sizeData.sellOutPct) / 100;
                      
                      return (
                        <tr key={index} className={index % 2 === 0 ? 'bg-gray-50' : ''}>
                          <td className="py-2 px-4 border-b">{size}</td>
                          <td className="py-2 px-4 border-b text-right">{sizeData.orderQty}</td>
                          <td className="py-2 px-4 border-b text-right">{sizeData.soldQty}</td>
                          <td className="py-2 px-4 border-b text-right">{formatNumber(sizeData.sellOutPct)}%</td>
                          <td className="py-2 px-4 border-b text-right">{formatNumber(sizeData.sellThroughPct)}%</td>
                          <td className="py-2 px-4 border-b text-right">{formatNumber(split, 2)}</td>
                        </tr>
                      );
                    })}
                    <tr className="font-semibold bg-gray-100">
                      <td className="py-2 px-4 border-b">TOTALE</td>
                      <td className="py-2 px-4 border-b text-right">
                        {datiProcessati[combSelezionata].datiPerClasse[classeMerchSelezionata].orderQtyTotal}
                      </td>
                      <td className="py-2 px-4 border-b text-right">
                        {datiProcessati[combSelezionata].datiPerClasse[classeMerchSelezionata].soldQtyTotal}
                      </td>
                      <td className="py-2 px-4 border-b text-right">100%</td>
                      <td className="py-2 px-4 border-b text-right">
                        {formatNumber(datiProcessati[combSelezionata].datiPerClasse[classeMerchSelezionata].soldQtyTotal > 0 && 
                          datiProcessati[combSelezionata].datiPerClasse[classeMerchSelezionata].orderQtyTotal > 0 ? 
                          (datiProcessati[combSelezionata].datiPerClasse[classeMerchSelezionata].soldQtyTotal / 
                            datiProcessati[combSelezionata].datiPerClasse[classeMerchSelezionata].orderQtyTotal) * 100 : 0)}%
                      </td>
                      <td className="py-2 px-4 border-b text-right">
                        {quantitaProposta ? quantitaProposta : '100'}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
              
              <div className="mt-4">
                <label className="block text-gray-700 font-semibold mb-2">Quantità Proposta:</label>
                <input
                  type="number"
                  value={quantitaProposta}
                  onChange={(e) => setQuantitaProposta(e.target.value)}
                  placeholder="Inserisci la quantità proposta"
                  className="border border-gray-300 p-2 rounded w-full"
                />
              </div>
            </div>
          )}
          
          <div className="mt-6">
            <button
              onClick={downloadReport}
              disabled={processing}
              className="bg-green-600 text-white px-4 py-2 rounded hover:bg-green-700 disabled:bg-gray-400"
            >
              Scarica Report Excel Completo
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExcelProcessor;