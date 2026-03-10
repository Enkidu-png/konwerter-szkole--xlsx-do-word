import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileText, Download, Copy, Check, Search, Trash2, Info } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { saveAs } from 'file-saver';
import PizZip from 'pizzip';

interface TrainingProduct {
  id: string;
  code: string;
  identifier: string;
  title: string;
  name: string;
  description: string;
  scope: string;
  extraData: { label: string; value: string }[];
  rawData: any;
}

export default function App() {
  const [products, setProducts] = useState<TrainingProduct[]>([]);
  const [selectedProduct, setSelectedProduct] = useState<TrainingProduct | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [copied, setCopied] = useState(false);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement> | React.DragEvent) => {
    let file: File | undefined;
    if ('files' in e.target && e.target.files) {
      file = e.target.files[0];
    } else if ('dataTransfer' in e && e.dataTransfer.files) {
      file = e.dataTransfer.files[0];
    }

    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        
        // Pobieramy dane jako tablicę tablic (rows as arrays)
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        if (rows.length < 2) return;

        const headers = rows[0];
        const dataRows = rows.slice(1);

        const parsedProducts: TrainingProduct[] = dataRows.map((row, index) => {
          // Mapowanie kolumn na podstawie indeksów (A=0, B=1, C=2, D=3, E=4, ...)
          
          // Szukamy indeksów dla kluczowych pól (pomijając A i B)
          const searchHeaders = headers.map((h: any) => String(h || '').toLowerCase());
          
          // Tytuł/Produkt - kolumna F (indeks 5)
          const titleIndex = 5;

          // Opis i Zakres
          let descIndex = searchHeaders.findIndex((h, i) => i >= 2 && h.includes('opis'));
          let scopeIndex = searchHeaders.findIndex((h, i) => i >= 2 && h.includes('zakres'));
          
          // Fallbacki jeśli nie znaleziono
          if (descIndex === -1) descIndex = 5; 
          if (scopeIndex === -1) scopeIndex = 6;

          const title = String(row[titleIndex] || 'Bez nazwy').trim();
          const code = String(row[3] || '').trim();
          const identifier = String(row[4] || '').trim();
          const description = String(row[descIndex] || '').trim();
          const scope = String(row[scopeIndex] || '').trim();
          
          const displayName = identifier ? `[${identifier}] ${title}` : title;

          // Dodatkowe kolumny X do AC (indeksy 23 do 28)
          const extraData: { label: string; value: string }[] = [];
          for (let i = 23; i <= 28; i++) {
            const val = row[i];
            if (val !== undefined && val !== null) {
              const strVal = String(val).trim();
              if (strVal && strVal !== 'undefined' && strVal !== 'null') {
                // Używamy nagłówka z pierwszego wiersza jako etykiety
                const label = headers[i] || `Kolumna ${XLSX.utils.encode_col(i)}`;
                extraData.push({ label, value: strVal });
              }
            }
          }

          return {
            id: `prod-${index}`,
            code,
            identifier,
            title,
            name: displayName,
            description,
            scope,
            extraData,
            rawData: row
          };
        });

        setProducts(parsedProducts.filter(p => p.title !== 'Bez nazwy' || p.identifier !== ''));
      } catch (err) {
        console.error("Error parsing Excel:", err);
        alert("Błąd podczas odczytu pliku Excel. Upewnij się, że plik jest poprawny.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const downloadWord = async () => {
    if (!selectedProduct) return;

    const extraHtml = selectedProduct.extraData.map(item => `
      <h2>${item.label}</h2>
      <div class="content">${item.value}</div>
    `).join('');

    const htmlContent = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset="UTF-8">
        <title>${selectedProduct.name}</title>
        <style>
          body { font-family: 'Calibri', sans-serif; }
          h1 { color: #1e293b; border-bottom: 2px solid #6366f1; padding-bottom: 10px; }
          h2 { color: #475569; margin-top: 20px; font-size: 14pt; }
          .content { margin-bottom: 20px; }
        </style>
      </head>
      <body>
        <div style="font-size: 10pt; color: #6366f1; font-weight: bold;">ID: ${selectedProduct.identifier}</div>
        <h1>${selectedProduct.title}</h1>
        <h2>Opis produktu</h2>
        <div class="content">${selectedProduct.description}</div>
        <h2>Zakres szkolenia</h2>
        <div class="content">${selectedProduct.scope}</div>
        ${extraHtml}
      </body>
      </html>
    `;

    try {
      const response = await fetch('/SZABLON2.docx');
      if (!response.ok) throw new Error('Nie można załadować szablonu SZABLON2.docx z folderu public/');
      const arrayBuffer = await response.arrayBuffer();

      const zip = new PizZip(arrayBuffer);

      // Replace ###TRESC### paragraph with altChunk reference
      let docXml = zip.files['word/document.xml'].asText();
      docXml = docXml.replace(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?###TRESC###[\s\S]*?<\/w:p>/, '<w:altChunk r:id="htmlContent"/>');
      zip.file('word/document.xml', docXml);

      // Add HTML chunk as separate file in ZIP
      zip.file('word/afchunk.html', htmlContent);

      // Add relationship for the altChunk
      let relsXml = zip.files['word/_rels/document.xml.rels'].asText();
      relsXml = relsXml.replace(
        '</Relationships>',
        '<Relationship Id="htmlContent" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk.html"/></Relationships>'
      );
      zip.file('word/_rels/document.xml.rels', relsXml);

      // Register HTML content type
      let contentTypesXml = zip.files['[Content_Types].xml'].asText();
      contentTypesXml = contentTypesXml.replace(
        '</Types>',
        '<Override PartName="/word/afchunk.html" ContentType="text/html"/></Types>'
      );
      zip.file('[Content_Types].xml', contentTypesXml);

      const blob = zip.generate({ type: 'blob' });
      saveAs(blob, `${selectedProduct.name.replace(/[^a-z0-9]/gi, '_')}.docx`);
    } catch (err) {
      console.error('Error generating Word document:', err);
      alert(`Błąd podczas generowania dokumentu: ${err instanceof Error ? err.message : 'Nieznany błąd'}`);
    }
  };

  const copyToClipboard = () => {
    if (!selectedProduct) return;
    
    const extraHtml = selectedProduct.extraData.map(item => `
      <h3>${item.label}</h3>
      ${item.value}
    `).join('');
    
    const combinedHtml = `
      <h3>Opis produktu</h3>
      ${selectedProduct.description}
      <h3>Zakres szkolenia</h3>
      ${selectedProduct.scope}
      ${extraHtml}
    `;

    // Copy as text/html so Word/RTF editors can paste it with formatting
    const type = "text/html";
    const blob = new Blob([combinedHtml], { type });
    const data = [new ClipboardItem({ [type]: blob })];

    navigator.clipboard.write(data).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }).catch(err => {
      console.error("Clipboard error:", err);
      // Fallback to plain text if HTML copy fails
      const plainText = `Opis produktu:\n${selectedProduct.description.replace(/<[^>]*>/g, '')}\n\nZakres szkolenia:\n${selectedProduct.scope.replace(/<[^>]*>/g, '')}`;
      navigator.clipboard.writeText(plainText).then(() => {
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
      });
    });
  };

  const filteredProducts = products.filter(p => 
    p.name.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="min-h-screen bg-slate-50 font-sans p-4 md:p-8">
      <div className="max-w-6xl mx-auto">
        <header className="mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h1 className="text-3xl font-bold text-slate-900 tracking-tight">Konwerter Szkoleń</h1>
            <p className="text-slate-500 mt-1">Importuj bazę XLSX i eksportuj opisy do Worda.</p>
          </div>
          {products.length > 0 && (
            <button 
              onClick={() => { setProducts([]); setSelectedProduct(null); }}
              className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-red-600 hover:bg-red-50 rounded-lg transition-colors"
            >
              <Trash2 size={18} />
              Wyczyść bazę
            </button>
          )}
        </header>

        {products.length === 0 ? (
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className={`
              relative border-2 border-dashed rounded-2xl p-12 text-center transition-all
              ${isDragging ? 'border-indigo-500 bg-indigo-50' : 'border-slate-300 bg-white'}
            `}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(e) => { e.preventDefault(); setIsDragging(false); handleFileUpload(e); }}
          >
            <input 
              type="file" 
              accept=".xlsx, .xls" 
              onChange={handleFileUpload}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            />
            <div className="flex flex-col items-center">
              <div className="w-16 h-16 bg-indigo-100 text-indigo-600 rounded-full flex items-center justify-center mb-4">
                <Upload size={32} />
              </div>
              <h2 className="text-xl font-semibold text-slate-800">Prześlij plik bazy danych</h2>
              <p className="text-slate-500 mt-2 max-w-xs mx-auto">
                Przeciągnij i upuść plik Excel (.xlsx) tutaj lub kliknij, aby wybrać z dysku.
              </p>
              <div className="mt-6 flex flex-col items-center gap-2 text-xs text-slate-400">
                <div className="flex items-center gap-2">
                  <Info size={14} />
                  <span>Wymagane kolumny: Produkt, Opis, Zakres</span>
                </div>
                <div className="flex items-center gap-2">
                  <Check size={12} className="text-green-500" />
                  <span>Kolumny A i B są ignorowane</span>
                </div>
                <div className="flex items-center gap-2">
                  <Check size={12} className="text-green-500" />
                  <span>Identyfikator pobierany z kolumny E</span>
                </div>
              </div>
            </div>
          </motion.div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            <div className="lg:col-span-4 flex flex-col gap-4">
              <div className="relative">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                <input 
                  type="text"
                  placeholder="Szukaj produktu..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-transparent outline-none transition-all shadow-sm"
                />
              </div>
              
              <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden flex flex-col max-h-[600px]">
                <div className="p-4 border-b border-slate-100 bg-slate-50/50">
                  <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                    Produkty ({filteredProducts.length})
                  </span>
                </div>
                <div className="overflow-y-auto flex-1">
                  {filteredProducts.map(product => (
                    <button
                      key={product.id}
                      onClick={() => setSelectedProduct(product)}
                      className={`
                        w-full text-left p-4 border-b border-slate-50 transition-all hover:bg-slate-50
                        ${selectedProduct?.id === product.id ? 'bg-indigo-50 border-l-4 border-l-indigo-500' : ''}
                      `}
                    >
                      <div className="font-medium text-slate-800 line-clamp-2">{product.name}</div>
                    </button>
                  ))}
                  {filteredProducts.length === 0 && (
                    <div className="p-8 text-center text-slate-400 italic">
                      Nie znaleziono produktów.
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="lg:col-span-8">
              <AnimatePresence mode="wait">
                {selectedProduct ? (
                  <motion.div
                    key={selectedProduct.id}
                    initial={{ opacity: 0, x: 20 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, x: -20 }}
                    className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden"
                  >
                    <div className="p-4 border-b border-slate-100 bg-slate-50/50 flex flex-wrap items-center justify-between gap-4">
                      <div className="flex items-center gap-2 text-indigo-600 font-semibold">
                        <FileText size={20} />
                        <span>Podgląd treści</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={copyToClipboard}
                          className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 rounded-lg text-sm font-medium text-slate-700 hover:bg-slate-50 transition-all shadow-sm"
                        >
                          {copied ? <Check size={16} className="text-green-500" /> : <Copy size={16} />}
                          {copied ? 'Skopiowano!' : 'Kopiuj HTML'}
                        </button>
                        <button 
                          onClick={downloadWord}
                          className="flex items-center gap-2 px-4 py-2 bg-indigo-600 rounded-lg text-sm font-medium text-white hover:bg-indigo-700 transition-all shadow-sm"
                        >
                          <Download size={16} />
                          Pobierz Word (.docx)
                        </button>
                      </div>
                    </div>

                    <div className="p-6 md:p-8 space-y-8 max-h-[700px] overflow-y-auto">
                      <div>
                        {selectedProduct.identifier && (
                          <div className="text-xs font-bold text-indigo-500 uppercase tracking-widest mb-1">
                            ID: {selectedProduct.identifier}
                          </div>
                        )}
                        <h2 className="text-2xl font-bold text-slate-900 mb-2">{selectedProduct.title}</h2>
                        <div className="h-1 w-20 bg-indigo-500 rounded-full"></div>
                      </div>

                      <section>
                        <h3 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-4">Opis produktu</h3>
                        <div 
                          className="prose prose-slate max-w-none text-slate-700 bg-slate-50 p-4 rounded-xl border border-slate-100"
                          dangerouslySetInnerHTML={{ __html: selectedProduct.description || '<p class="italic text-slate-400">Brak opisu</p>' }}
                        />
                      </section>

                      <section>
                        <h3 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-4">Zakres szkolenia</h3>
                        <div 
                          className="prose prose-slate max-w-none text-slate-700 bg-slate-50 p-4 rounded-xl border border-slate-100"
                          dangerouslySetInnerHTML={{ __html: selectedProduct.scope || '<p class="italic text-slate-400">Brak zakresu</p>' }}
                        />
                      </section>

                      {selectedProduct.extraData.map((item, idx) => (
                        <section key={idx}>
                          <h3 className="text-sm font-bold text-slate-400 uppercase tracking-widest mb-4">{item.label}</h3>
                          <div 
                            className="prose prose-slate max-w-none text-slate-700 bg-slate-50 p-4 rounded-xl border border-slate-100"
                            dangerouslySetInnerHTML={{ __html: item.value }}
                          />
                        </section>
                      ))}
                    </div>
                  </motion.div>
                ) : (
                  <div className="h-full min-h-[400px] flex flex-col items-center justify-center text-center p-12 bg-white border border-slate-200 border-dashed rounded-2xl text-slate-400">
                    <div className="w-16 h-16 bg-slate-50 rounded-full flex items-center justify-center mb-4">
                      <FileText size={32} />
                    </div>
                    <p>Wybierz produkt z listy po lewej stronie,<br />aby zobaczyć szczegóły i wyeksportować treść.</p>
                  </div>
                )}
              </AnimatePresence>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
