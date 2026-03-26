import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileText, Download, Copy, Check, Search, Trash2, Info, UserCheck } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { saveAs } from 'file-saver';
import PizZip from 'pizzip';

type FooterKey = 'jagienka' | 'alicja' | 'agata';

const FOOTERS: { key: FooterKey; label: string }[] = [
  { key: 'jagienka', label: 'Jagienka' },
  { key: 'alicja', label: 'Alicja' },
  { key: 'agata', label: 'Agata' },
];

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

function buildProductHtml(product: TrainingProduct): string {
  const extraHtml = product.extraData.map(item => `
    ${String(item.label).toLowerCase().includes('cross') ? '' : `<h2>${item.label}</h2>`}
    <div class="content">${item.value}</div>
  `).join('');

  return `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta charset="UTF-8">
      <title>${product.name}</title>
      <style>
        body { font-family: 'Aptos', sans-serif; }
        h1 { color: #1e293b; border-bottom: 2px solid #6366f1; padding-bottom: 10px; }
        h2 { color: #475569; margin-top: 20px; font-size: 14pt; }
        .content { margin-bottom: 20px; color: #000000 !important; }
        .content * { color: #000000 !important; }
      </style>
    </head>
    <body>
      <div style="font-size: 10pt; color: #6366f1; font-weight: bold;">ID: ${product.identifier}</div>
      <h1>${product.title}</h1>
      <h2>Opis produktu</h2>
      <div class="content">${product.description}</div>
      <h2>Zakres szkolenia</h2>
      <div class="content">${product.scope}</div>
      ${extraHtml}
    </body>
    </html>
  `;
}

async function appendFooterToZip(mainZip: PizZip, footerBuffer: ArrayBuffer): Promise<void> {
  // Direct XML merge approach:
  // The footer is a full-page design — a PNG background image (page-anchored) with
  // floating tables (margin-anchored, default). Since the footer sectPr has zero margins,
  // margin == page edge, so all tblpX/tblpY coordinates are correct absolute positions.
  // We replace the main document's final <w:sectPr> with the footer's zero-margin sectPr
  // so the floating tables land exactly where they should.
  const footerZip = new PizZip(footerBuffer);

  const footerDoc = footerZip.files['word/document.xml'].asText();
  const bodyStart = footerDoc.indexOf('<w:body>') + '<w:body>'.length;
  const bodyEnd = footerDoc.lastIndexOf('</w:body>');
  const bodyContent = footerDoc.slice(bodyStart, bodyEnd);

  const lastSectPrStart = bodyContent.lastIndexOf('<w:sectPr');
  let footerBodyXml = bodyContent.slice(0, lastSectPrStart);
  // Use footer's own sectPr (zero margins) — strip its Word header/footer references
  // since those files live in the footer docx, not the main ZIP
  let footerSectPrXml = bodyContent.slice(lastSectPrStart)
    .replace(/<w:headerReference[^/]*\/>/g, '')
    .replace(/<w:footerReference[^/]*\/>/g, '');

  // Copy images and hyperlinks from footer rels into main document rels,
  // building a remapping table so we can update the body XML references.
  const footerRelsXml = footerZip.files['word/_rels/document.xml.rels'].asText();
  const relPattern = /<Relationship Id="([^"]+)" Type="([^"]+)" Target="([^"]+)"(?:[^/]*)\/>/g;
  let mainRels = mainZip.files['word/_rels/document.xml.rels'].asText();
  const idMap: Record<string, string> = {};
  let counter = 200;

  let m;
  while ((m = relPattern.exec(footerRelsXml)) !== null) {
    const [, oldId, type, target] = m;
    const relType = type.split('/').pop();
    if (relType === 'image') {
      const newId = `rIdFImg${counter++}`;
      const imgFileName = target.split('/').pop()!;
      const imgBytes = footerZip.files[`word/${target}`]?.asUint8Array();
      if (imgBytes) {
        mainZip.file(`word/media/footer_${imgFileName}`, imgBytes);
        mainRels = mainRels.replace('</Relationships>',
          `<Relationship Id="${newId}" Type="${type}" Target="media/footer_${imgFileName}"/></Relationships>`);
        idMap[oldId] = newId;
      }
    } else if (relType === 'hyperlink') {
      const newId = `rIdFHlnk${counter++}`;
      mainRels = mainRels.replace('</Relationships>',
        `<Relationship Id="${newId}" Type="${type}" Target="${target}" TargetMode="External"/></Relationships>`);
      idMap[oldId] = newId;
    }
  }

  // Remap all old IDs in the injected body XML
  for (const [oldId, newId] of Object.entries(idMap)) {
    footerBodyXml = footerBodyXml.split(`"${oldId}"`).join(`"${newId}"`);
  }

  // Replace main document's final standalone <w:sectPr> (and everything after it until
  // </w:body>) with: page break + footer body content + footer's zero-margin sectPr
  let mainDoc = mainZip.files['word/document.xml'].asText();
  const mainLastSectPr = mainDoc.lastIndexOf('<w:sectPr');
  mainDoc =
    mainDoc.slice(0, mainLastSectPr) +
    footerBodyXml +
    footerSectPrXml +
    '</w:body></w:document>';

  mainZip.file('word/document.xml', mainDoc);
  mainZip.file('word/_rels/document.xml.rels', mainRels);
}

async function generateDocxBlob(product: TrainingProduct, templateBuffer: ArrayBuffer, footerBuffer: ArrayBuffer): Promise<Blob> {
  const htmlContent = buildProductHtml(product);

  const zip = new PizZip(templateBuffer);

  let docXml = zip.files['word/document.xml'].asText();
  docXml = docXml.replace(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?###TRESC###[\s\S]*?<\/w:p>/, '<w:altChunk r:id="htmlContent"/>');
  const xmlTitle = product.title
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
  // {{Title}} is split across multiple XML runs by Word — match from <w:t>{{</w:t> through <w:t>}}</w:t></w:r>
  docXml = docXml.replace(/<w:t>\{\{<\/w:t><\/w:r>[\s\S]*?<w:t>\}\}<\/w:t><\/w:r>/g, `<w:t>${xmlTitle}</w:t></w:r>`);
  zip.file('word/document.xml', docXml);

  zip.file('word/afchunk.html', htmlContent);

  let relsXml = zip.files['word/_rels/document.xml.rels'].asText();
  relsXml = relsXml.replace(
    '</Relationships>',
    '<Relationship Id="htmlContent" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk.html"/></Relationships>'
  );
  zip.file('word/_rels/document.xml.rels', relsXml);

  let contentTypesXml = zip.files['[Content_Types].xml'].asText();
  contentTypesXml = contentTypesXml.replace(
    '</Types>',
    '<Override PartName="/word/afchunk.html" ContentType="text/html"/></Types>'
  );
  zip.file('[Content_Types].xml', contentTypesXml);

  await appendFooterToZip(zip, footerBuffer);

  return zip.generate({ type: 'blob' });
}

async function generateMergedOfferBlob(selectedProducts: TrainingProduct[], templateBuffer: ArrayBuffer, footerBuffer: ArrayBuffer): Promise<Blob> {
  const zip = new PizZip(templateBuffer);

  let docXml = zip.files['word/document.xml'].asText();

  // {{Title}} intentionally left as-is in the merged offer

  // Replace ###TRESC### paragraph with one altChunk per product, page breaks between them
  const PAGE_BREAK = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
  const altChunksXml = selectedProducts
    .map((_, i) => `<w:altChunk r:id="htmlContent${i}"/>`)
    .join(PAGE_BREAK);
  docXml = docXml.replace(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?###TRESC###[\s\S]*?<\/w:p>/, altChunksXml);
  zip.file('word/document.xml', docXml);

  let relsXml = zip.files['word/_rels/document.xml.rels'].asText();
  let contentTypesXml = zip.files['[Content_Types].xml'].asText();

  for (let i = 0; i < selectedProducts.length; i++) {
    const html = buildProductHtml(selectedProducts[i]);
    zip.file(`word/afchunk${i}.html`, html);
    relsXml = relsXml.replace(
      '</Relationships>',
      `<Relationship Id="htmlContent${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="afchunk${i}.html"/></Relationships>`
    );
    contentTypesXml = contentTypesXml.replace(
      '</Types>',
      `<Override PartName="/word/afchunk${i}.html" ContentType="text/html"/></Types>`
    );
  }

  zip.file('word/_rels/document.xml.rels', relsXml);
  zip.file('[Content_Types].xml', contentTypesXml);

  await appendFooterToZip(zip, footerBuffer);

  return zip.generate({ type: 'blob' });
}

export default function App() {
  const [products, setProducts] = useState<TrainingProduct[]>([]);
  const [selectedProduct, setSelectedProduct] = useState<TrainingProduct | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [copied, setCopied] = useState(false);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [batchProgress, setBatchProgress] = useState<{ current: number; total: number } | null>(null);
  const [isGeneratingOffer, setIsGeneratingOffer] = useState(false);
  const [selectedFooter, setSelectedFooter] = useState<FooterKey | null>(null);

  const loadFooterBuffer = async (key: FooterKey): Promise<ArrayBuffer> => {
    const response = await fetch(`/Stopka_${key}.docx`);
    if (!response.ok) throw new Error(`Nie można załadować stopki dla: ${key}`);
    return response.arrayBuffer();
  };

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

        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        if (rows.length < 2) return;

        const headers = rows[0];
        const dataRows = rows.slice(1);

        const parsedProducts: TrainingProduct[] = dataRows.map((row, index) => {
          const searchHeaders = headers.map((h: any) => String(h || '').toLowerCase());

          const titleIndex = 5;

          let descIndex = searchHeaders.findIndex((h, i) => i >= 2 && h.includes('opis'));
          let scopeIndex = searchHeaders.findIndex((h, i) => i >= 2 && h.includes('zakres'));

          if (descIndex === -1) descIndex = 5;
          if (scopeIndex === -1) scopeIndex = 6;

          const title = String(row[titleIndex] || 'Bez nazwy').trim();
          const code = String(row[3] || '').trim();
          const identifier = String(row[4] || '').trim();
          const description = String(row[descIndex] || '').trim();
          const scope = String(row[scopeIndex] || '').trim();

          const displayName = identifier ? `[${identifier}] ${title}` : title;

          const extraData: { label: string; value: string }[] = [];
          for (let i = 23; i <= 28; i++) {
            const val = row[i];
            if (val !== undefined && val !== null) {
              const strVal = String(val).trim();
              if (strVal && strVal !== 'undefined' && strVal !== 'null') {
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
        setSelectedIds(new Set());
      } catch (err) {
        console.error("Error parsing Excel:", err);
        alert("Błąd podczas odczytu pliku Excel. Upewnij się, że plik jest poprawny.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const downloadWord = async () => {
    if (!selectedProduct || !selectedFooter) return;

    try {
      const [templateResponse, footerBuffer] = await Promise.all([
        fetch('/SZABLON2.docx'),
        loadFooterBuffer(selectedFooter),
      ]);
      if (!templateResponse.ok) throw new Error('Nie można załadować szablonu SZABLON2.docx z folderu public/');
      const templateBuffer = await templateResponse.arrayBuffer();
      const blob = await generateDocxBlob(selectedProduct, templateBuffer, footerBuffer);
      saveAs(blob, `${selectedProduct.name.replace(/[^a-z0-9]/gi, '_')}.docx`);
    } catch (err) {
      console.error('Error generating Word document:', err);
      alert(`Błąd podczas generowania dokumentu: ${err instanceof Error ? err.message : 'Nieznany błąd'}`);
    }
  };

  const downloadBatch = async () => {
    const ids = Array.from(selectedIds);
    const total = ids.length;
    if (total === 0 || !selectedFooter) return;

    try {
      const [templateResponse, footerBuffer] = await Promise.all([
        fetch('/SZABLON2.docx'),
        loadFooterBuffer(selectedFooter),
      ]);
      if (!templateResponse.ok) throw new Error('Nie można załadować szablonu SZABLON2.docx z folderu public/');
      const templateBuffer = await templateResponse.arrayBuffer();

      const batchZip = new PizZip();

      for (let i = 0; i < ids.length; i++) {
        setBatchProgress({ current: i + 1, total });
        const product = products.find(p => p.id === ids[i]);
        if (!product) continue;

        const blob = await generateDocxBlob(product, templateBuffer, footerBuffer);
        const arrayBuffer = await blob.arrayBuffer();
        const safeName = product.name.replace(/[^a-z0-9]/gi, '_');
        const entryName = batchZip.files[`${safeName}.docx`]
          ? `${safeName}_${product.id}.docx`
          : `${safeName}.docx`;
        batchZip.file(entryName, arrayBuffer);
      }

      const zipBlob = batchZip.generate({ type: 'blob' });
      const today = new Date().toISOString().split('T')[0];
      saveAs(zipBlob, `eksport_${today}.zip`);

      setSelectedIds(new Set());
    } catch (err) {
      console.error('Error generating batch export:', err);
      alert(`Błąd podczas eksportu: ${err instanceof Error ? err.message : 'Nieznany błąd'}`);
    } finally {
      setBatchProgress(null);
    }
  };

  const toggleSelected = (id: string) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
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

    const type = "text/html";
    const blob = new Blob([combinedHtml], { type });
    const data = [new ClipboardItem({ [type]: blob })];

    navigator.clipboard.write(data).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }).catch(err => {
      console.error("Clipboard error:", err);
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

  const downloadMergedOffer = async () => {
    const ids = Array.from(selectedIds);
    if (ids.length === 0 || !selectedFooter) return;

    const selectedProducts = ids
      .map(id => products.find(p => p.id === id))
      .filter(Boolean) as TrainingProduct[];

    setIsGeneratingOffer(true);
    try {
      const [templateResponse, footerBuffer] = await Promise.all([
        fetch('/SZABLON2.docx'),
        loadFooterBuffer(selectedFooter),
      ]);
      if (!templateResponse.ok) throw new Error('Nie można załadować szablonu SZABLON2.docx z folderu public/');
      const templateBuffer = await templateResponse.arrayBuffer();
      const blob = await generateMergedOfferBlob(selectedProducts, templateBuffer, footerBuffer);
      const today = new Date().toISOString().split('T')[0];
      saveAs(blob, `Oferta_Asseco_Academy_${today}.docx`);
    } catch (err) {
      console.error('Error generating merged offer:', err);
      alert(`Błąd podczas generowania oferty: ${err instanceof Error ? err.message : 'Nieznany błąd'}`);
    } finally {
      setIsGeneratingOffer(false);
    }
  };

  const isBusy = batchProgress !== null || isGeneratingOffer;

  const FooterSelector = ({ compact = false }: { compact?: boolean }) => (
    <div className={`flex items-center gap-1.5 ${compact ? '' : 'gap-2'}`}>
      {!compact && (
        <span className="flex items-center gap-1 text-xs text-slate-400 font-medium">
          <UserCheck size={13} />
          Stopka:
        </span>
      )}
      {FOOTERS.map(f => (
        <button
          key={f.key}
          onClick={() => setSelectedFooter(prev => prev === f.key ? null : f.key)}
          disabled={isBusy}
          className={`
            px-3 py-1.5 rounded-lg text-xs font-semibold transition-all border
            ${selectedFooter === f.key
              ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm'
              : 'bg-white text-slate-600 border-slate-200 hover:border-indigo-300 hover:text-indigo-600'}
            disabled:opacity-50
          `}
        >
          {f.label}
        </button>
      ))}
    </div>
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
              onClick={() => { setProducts([]); setSelectedProduct(null); setSelectedIds(new Set()); }}
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
                <div className="p-4 border-b border-slate-100 bg-slate-50/50 flex flex-wrap items-center justify-between gap-2">
                  <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                    Produkty ({filteredProducts.length})
                    {selectedIds.size > 0 && (
                      <span className="text-indigo-500"> · Zaznaczono: {selectedIds.size}</span>
                    )}
                  </span>
                  {selectedIds.size > 0 && (
                    <div className="flex flex-wrap items-center gap-2">
                      <FooterSelector compact />
                      <div className="w-px h-4 bg-slate-200" />
                      <button
                        onClick={() => setSelectedIds(new Set())}
                        disabled={isBusy}
                        className="text-xs text-slate-500 hover:text-slate-700 disabled:opacity-40 transition-colors"
                      >
                        Odznacz
                      </button>
                      <button
                        onClick={downloadBatch}
                        disabled={isBusy || !selectedFooter}
                        title={!selectedFooter ? 'Wybierz stopkę przed eksportem' : undefined}
                        className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 disabled:opacity-40 disabled:cursor-not-allowed rounded-lg text-xs font-medium text-white transition-all shadow-sm"
                      >
                        <Download size={13} />
                        {batchProgress
                          ? `Generowanie (${batchProgress.current}/${batchProgress.total})…`
                          : `Osobno (${selectedIds.size})`}
                      </button>
                      <button
                        onClick={downloadMergedOffer}
                        disabled={isBusy || !selectedFooter}
                        title={!selectedFooter ? 'Wybierz stopkę przed eksportem' : undefined}
                        className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 disabled:cursor-not-allowed rounded-lg text-xs font-medium text-white transition-all shadow-sm"
                      >
                        <Download size={13} />
                        {isGeneratingOffer
                          ? `Generowanie oferty…`
                          : `Jako ofertę (${selectedIds.size})`}
                      </button>
                    </div>
                  )}
                </div>
                <div className="overflow-y-auto flex-1">
                  {filteredProducts.map(product => (
                    <div
                      key={product.id}
                      className={`
                        flex items-center border-b border-slate-50 transition-all hover:bg-slate-50
                        ${selectedProduct?.id === product.id ? 'bg-indigo-50 border-l-4 border-l-indigo-500' : ''}
                      `}
                    >
                      <label className="flex items-center pl-3 pr-2 py-4 cursor-pointer" onClick={e => e.stopPropagation()}>
                        <input
                          type="checkbox"
                          checked={selectedIds.has(product.id)}
                          onChange={() => toggleSelected(product.id)}
                          className="w-4 h-4 rounded border-slate-300 text-indigo-600 cursor-pointer accent-indigo-600"
                        />
                      </label>
                      <button
                        onClick={() => setSelectedProduct(product)}
                        className="flex-1 text-left px-2 py-4 min-w-0"
                      >
                        <div className="font-medium text-slate-800 line-clamp-2">{product.name}</div>
                      </button>
                    </div>
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
                      <div className="flex flex-wrap items-center gap-2">
                        <button
                          onClick={copyToClipboard}
                          className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 rounded-lg text-sm font-medium text-slate-700 hover:bg-slate-50 transition-all shadow-sm"
                        >
                          {copied ? <Check size={16} className="text-green-500" /> : <Copy size={16} />}
                          {copied ? 'Skopiowano!' : 'Kopiuj HTML'}
                        </button>
                          <FooterSelector />
                        <button
                          onClick={downloadWord}
                          disabled={!selectedFooter}
                          title={!selectedFooter ? 'Wybierz stopkę przed pobraniem' : undefined}
                          className="flex items-center gap-2 px-4 py-2 bg-indigo-600 rounded-lg text-sm font-medium text-white hover:bg-indigo-700 disabled:opacity-40 disabled:cursor-not-allowed transition-all shadow-sm"
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
