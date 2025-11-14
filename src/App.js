import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import bwipjs from 'bwip-js';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import './App.css';

function App() {
  const [data, setData] = useState([]);
  const [fileName, setFileName] = useState('');
  const [loading, setLoading] = useState(false);
  const [selectedColumn, setSelectedColumn] = useState(3);
  const [startRow, setStartRow] = useState(1);
  const [endRow, setEndRow] = useState(null); // YANGI: Tugash qatori
  const [generatedBarcodes, setGeneratedBarcodes] = useState([]);

  // Excel faylni yuklash
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      const binaryStr = event.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      
      setData(jsonData);
      setEndRow(jsonData.length - 1); // Default: oxirgi qatorgacha
      setGeneratedBarcodes([]);
    };

    reader.readAsBinaryString(file);
  };

  // Data Matrix yaratish (OQ FON, TINIQ VA KATTA)
  const generateDataMatrix = async (text) => {
    try {
      const innerCanvas = document.createElement('canvas');
      
      await bwipjs.toCanvas(innerCanvas, {
        bcid: 'datamatrix',
        text: text.toString(),
        scale: 12,
        height: 20,
        includetext: false,
        paddingwidth: 0,
        paddingheight: 0,
      });

      const padding = 80;
      const finalCanvas = document.createElement('canvas');
      const finalWidth = innerCanvas.width + (padding * 2);
      const finalHeight = innerCanvas.height + (padding * 2);
      
      finalCanvas.width = finalWidth;
      finalCanvas.height = finalHeight;
      
      const ctx = finalCanvas.getContext('2d');
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(0, 0, finalWidth, finalHeight);
      ctx.drawImage(innerCanvas, padding, padding);

      return finalCanvas.toDataURL('image/png', 1.0);
    } catch (error) {
      console.error('Data Matrix xatolik:', error);
      return null;
    }
  };

  // Excel yaratish va yuklab olish
  const processAndDownload = async () => {
    if (data.length === 0) return;
    
    // Qator oralig'ini tekshirish
    const actualEndRow = endRow || data.length - 1;
    
    if (startRow < 1 || startRow > data.length - 1) {
      alert('❌ Boshlang\'ich qator noto\'g\'ri!');
      return;
    }
    
    if (actualEndRow < startRow || actualEndRow > data.length - 1) {
      alert('❌ Tugash qatori noto\'g\'ri!');
      return;
    }

    setLoading(true);

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      const barcodes = [];

      // Sarlavhani qo'shish
      worksheet.addRow(data[0]);
      worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF667EEA' }
      };
      worksheet.getRow(1).height = 25;

      // Har bir qator uchun
      for (let i = 1; i < data.length; i++) {
        const row = [...data[i]];
        const columnValue = row[selectedColumn];

        const originalValue = row[selectedColumn];
        row[selectedColumn] = '';
        worksheet.addRow(row);

        // YANGI: Tanlangan oraliqda faqat Data Matrix yaratish
        if (i >= startRow && i <= actualEndRow && columnValue && columnValue.toString().trim()) {
          const base64Image = await generateDataMatrix(columnValue);

          if (base64Image) {
            barcodes.push({
              rowNumber: i,
              data: columnValue.toString(),
              base64: base64Image
            });

            const imageId = workbook.addImage({
              base64: base64Image,
              extension: 'png',
            });

            worksheet.addImage(imageId, {
              tl: { col: selectedColumn, row: i },
              ext: { width: 220, height: 220 }
            });

            worksheet.getRow(i + 1).height = 165;
          }
        } else {
          worksheet.getCell(i + 1, selectedColumn + 1).value = originalValue;
        }

        const columnLetter = String.fromCharCode(65 + selectedColumn);
        const cell = worksheet.getCell(`${columnLetter}${i + 1}`);
        
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFFFF' }
        };
      }

      worksheet.columns = data[0].map((_, index) => {
        if (index === selectedColumn) return { width: 32 };
        if (index === 0) return { width: 10 };
        if (index === 1) return { width: 35 };
        return { width: 15 };
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = fileName.replace(/\.[^/.]+$/, '') + '_datamatrix.xlsx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      setGeneratedBarcodes(barcodes);

      const columnLetter = String.fromCharCode(65 + selectedColumn);
      alert(`✅ Excel fayl tayyor!\n\n${columnLetter} ustundagi ${barcodes.length} ta ma'lumot Data Matrix kodlari bilan almashtirildi.\n\n📌 Qator oralig'i: ${startRow} dan ${actualEndRow} gacha\n\n✨ Xususiyatlar:\n• Oq fon (100% oq)\n• Chetlarda 80px bo'sh joy\n• 220x220 piksel o'lcham\n• Maksimal tiniqlik`);
    } catch (error) {
      console.error('Xatolik:', error);
      alert('❌ Xatolik yuz berdi: ' + error.message);
    }

    setLoading(false);
  };

  const downloadSingleImage = (barcode) => {
    const link = document.createElement('a');
    link.href = barcode.base64;
    const safeData = barcode.data.substring(0, 30).replace(/[^a-zA-Z0-9]/g, '_');
    link.download = `datamatrix_row${barcode.rowNumber}_${safeData}.png`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const downloadAllImagesAsZip = async () => {
    if (generatedBarcodes.length === 0) {
      alert('❌ Avval Data Matrix kodlarni yaratish kerak!');
      return;
    }

    setLoading(true);

    try {
      const zip = new JSZip();
      const imagesFolder = zip.folder('DataMatrix_Images');

      for (let i = 0; i < generatedBarcodes.length; i++) {
        const barcode = generatedBarcodes[i];
        const base64Data = barcode.base64.split(',')[1];
        const safeData = barcode.data.substring(0, 30).replace(/[^a-zA-Z0-9]/g, '_');
        const fileName = `row${barcode.rowNumber}_${safeData}.png`;
        imagesFolder.file(fileName, base64Data, { base64: true });
      }

      const readmeContent = `Data Matrix Rasmlar (Tiniq va Oq Fonli)
==========================================

Jami rasmlar: ${generatedBarcodes.length}
Asl Excel fayl: ${fileName}
Tanlangan ustun: ${String.fromCharCode(65 + selectedColumn)}
Qator oralig'i: ${startRow} dan ${endRow || data.length - 1} gacha
Yaratilgan sana: ${new Date().toLocaleString('uz-UZ')}

✅ Har bir rasm xususiyatlari:
   • 100% Oq fonli (#FFFFFF)
   • Chetlarida 80px bo'sh joy
   • Scale: 12 (maksimal tiniqlik)
   • 220x220 piksel o'lcham
   • PNG formatda (maksimal sifat)
   • O'qish uchun optimal

Ushbu rasmlar Data Matrix skanerlar bilan osongina o'qiladi.`;
      
      zip.file('README.txt', readmeContent);

      const zipBlob = await zip.generateAsync({ 
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 9 }
      });

      const zipFileName = fileName.replace(/\.[^/.]+$/, '') + '_datamatrix_images.zip';
      saveAs(zipBlob, zipFileName);

      setLoading(false);
      alert(`✅ ${generatedBarcodes.length} ta rasm ZIP faylda muvaffaqiyatli yuklandi!`);
    } catch (error) {
      console.error('ZIP yaratishda xatolik:', error);
      alert('❌ ZIP yaratishda xatolik: ' + error.message);
      setLoading(false);
    }
  };

  const getColumnLetter = (index) => {
    return String.fromCharCode(65 + index);
  };

  const getDataCount = () => {
    if (data.length === 0) return 0;
    const actualEndRow = endRow || data.length - 1;
    return data.slice(startRow, actualEndRow + 1).filter(row => 
      row[selectedColumn] && row[selectedColumn].toString().trim()
    ).length;
  };

  return (
    <div className="App">
      <div className="container">
        <header>
          <div className="header-icon">✨</div>
          <h1>Data Matrix Generator Pro</h1>
          <p>Professional darajadagi Data Matrix kodlar yaratish tizimi</p>
        </header>

        <div className="info-box">
          <div className="info-icon">💎</div>
          <h3>Premium Xususiyatlar</h3>
          <div className="features-grid">
            <div className="feature-item">
              <span className="feature-icon">⬜</span>
              <span>100% Oq fon</span>
            </div>
            <div className="feature-item">
              <span className="feature-icon">🔍</span>
              <span>Maksimal tiniqlik</span>
            </div>
            <div className="feature-item">
              <span className="feature-icon">📏</span>
              <span>220x220 piksel</span>
            </div>
            <div className="feature-item">
              <span className="feature-icon">🎯</span>
              <span>Qator oralig'i</span>
            </div>
          </div>
        </div>

        <div className="upload-section">
          <label htmlFor="file-upload" className="file-upload-btn">
            <span className="btn-icon">📂</span>
            <span>Excel faylni yuklash</span>
          </label>
          <input
            id="file-upload"
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            style={{ display: 'none' }}
          />
        </div>

        {data.length > 0 && (
          <div className="result-section">
            <div className="settings-box">
              <h3><span className="section-icon">⚙️</span> Sozlamalar</h3>
              
              <div className="setting-group">
                <label className="setting-label">
                  <span className="label-icon">📊</span>
                  Ustunni tanlang:
                </label>
                <div className="column-selector">
                  {data[0].map((header, index) => (
                    <button
                      key={index}
                      className={`column-btn ${selectedColumn === index ? 'active' : ''}`}
                      onClick={() => setSelectedColumn(index)}
                    >
                      <div className="column-letter">{getColumnLetter(index)}</div>
                      <div className="column-name">{header || 'Bo\'sh'}</div>
                    </button>
                  ))}
                </div>
              </div>

              <div className="row-range-section">
                <label className="setting-label">
                  <span className="label-icon">📍</span>
                  Qator oralig'ini tanlang:
                </label>
                
                <div className="range-inputs">
                  <div className="range-input-group">
                    <label className="range-label">Boshlanish</label>
                    <input
                      type="number"
                      min="1"
                      max={data.length - 1}
                      value={startRow}
                      onChange={(e) => setStartRow(parseInt(e.target.value) || 1)}
                      className="range-input"
                    />
                  </div>
                  
                  <div className="range-separator">→</div>
                  
                  <div className="range-input-group">
                    <label className="range-label">Tugash</label>
                    <input
                      type="number"
                      min={startRow}
                      max={data.length - 1}
                      value={endRow || data.length - 1}
                      onChange={(e) => setEndRow(parseInt(e.target.value) || data.length - 1)}
                      className="range-input"
                    />
                  </div>
                </div>
                
                <div className="range-hint">
                  Jami ma'lumotlar: {data.length - 1} ta qator
                </div>
              </div>
            </div>

            <div className="stats-box">
              <div className="stats-grid">
                <div className="stat-card">
                  <div className="stat-icon">📄</div>
                  <div className="stat-content">
                    <div className="stat-label">Fayl</div>
                    <div className="stat-value">{fileName}</div>
                  </div>
                </div>
                
                <div className="stat-card">
                  <div className="stat-icon">📊</div>
                  <div className="stat-content">
                    <div className="stat-label">Jami qatorlar</div>
                    <div className="stat-value">{data.length - 1}</div>
                  </div>
                </div>
                
                <div className="stat-card highlight">
                  <div className="stat-icon">🎯</div>
                  <div className="stat-content">
                    <div className="stat-label">Tanlangan ustun</div>
                    <div className="stat-value">
                      {getColumnLetter(selectedColumn)} ({data[0][selectedColumn] || 'Bo\'sh'})
                    </div>
                  </div>
                </div>
                
                <div className="stat-card highlight">
                  <div className="stat-icon">📍</div>
                  <div className="stat-content">
                    <div className="stat-label">Qator oralig'i</div>
                    <div className="stat-value">{startRow} - {endRow || data.length - 1}</div>
                  </div>
                </div>
                
                <div className="stat-card success">
                  <div className="stat-icon">✨</div>
                  <div className="stat-content">
                    <div className="stat-label">Yaratiladi</div>
                    <div className="stat-value">{getDataCount()} ta Data Matrix</div>
                  </div>
                </div>
              </div>
            </div>

            <div className="action-buttons">
              <button 
                onClick={processAndDownload}
                className="process-btn"
                disabled={loading}
              >
                {loading ? (
                  <>
                    <span className="spinner"></span>
                    <span>Ishlanmoqda...</span>
                  </>
                ) : (
                  <>
                    <span className="btn-icon">🚀</span>
                    <span>Data Matrix yaratish va Excel yuklab olish</span>
                  </>
                )}
              </button>

              {generatedBarcodes.length > 0 && (
                <button 
                  onClick={downloadAllImagesAsZip}
                  className="download-all-btn"
                  disabled={loading}
                >
                  {loading ? (
                    <>
                      <span className="spinner"></span>
                      <span>ZIP yaratilmoqda...</span>
                    </>
                  ) : (
                    <>
                      <span className="btn-icon">📦</span>
                      <span>Barcha rasmlarni ZIP da yuklab olish ({generatedBarcodes.length} ta)</span>
                    </>
                  )}
                </button>
              )}
            </div>

            {generatedBarcodes.length > 0 && (
              <div className="barcodes-gallery">
                <h3><span className="section-icon">🖼️</span> Yaratilgan Data Matrix Kodlar</h3>
                <p className="gallery-subtitle">
                  Professional sifatda yaratilgan - oq fonli, maksimal tiniqlik!
                </p>
                
                <div className="gallery-grid">
                  {generatedBarcodes.map((barcode, index) => (
                    <div key={index} className="gallery-item">
                      <div className="gallery-header">
                        <span className="gallery-row">Qator {barcode.rowNumber}</span>
                        <button
                          onClick={() => downloadSingleImage(barcode)}
                          className="download-single-btn"
                          title="Rasmni yuklab olish"
                        >
                          💾
                        </button>
                      </div>
                      <div className="gallery-image-container">
                        <img 
                          src={barcode.base64} 
                          alt={`Data Matrix ${barcode.rowNumber}`}
                          className="gallery-image"
                        />
                      </div>
                      <div className="gallery-data">
                        <strong>Ma'lumot:</strong>
                        <div className="gallery-text">
                          {barcode.data.substring(0, 50)}
                          {barcode.data.length > 50 ? '...' : ''}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="preview-section">
              <h3><span className="section-icon">📋</span> Tanlangan ustun ko'rinishi</h3>
              <div className="preview-info">
                {getColumnLetter(selectedColumn)} ustuni, {startRow}-{endRow || data.length - 1} qatorlar
              </div>
              <div className="preview-list">
                {data.slice(startRow, Math.min(startRow + 5, (endRow || data.length - 1) + 1)).map((row, index) => (
                  <div key={index} className="preview-item">
                    <span className="preview-number">Qator {startRow + index}</span>
                    <span className="preview-text">
                      {row[selectedColumn] ? row[selectedColumn] : '(Bo\'sh)'}
                    </span>
                  </div>
                ))}
                {(endRow || data.length - 1) - startRow > 5 && (
                  <div className="preview-more">
                    ... va yana {(endRow || data.length - 1) - startRow - 5} ta qator
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {data.length === 0 && (
          <div className="empty-state">
            <div className="empty-icon">📊</div>
            <h2>Excel fayl yuklanmagan</h2>
            <p>Markirovka ma'lumotlari bo'lgan Excel faylni yuklang</p>
            <div className="steps">
              <div className="step">
                <span className="step-number">1</span>
                <div className="step-content">
                  <div className="step-title">Excel faylni yuklang</div>
                  <div className="step-desc">XLSX yoki XLS formatdagi faylni tanlang</div>
                </div>
              </div>
              <div className="step">
                <span className="step-number">2</span>
                <div className="step-content">
                  <div className="step-title">Sozlamalarni kiriting</div>
                  <div className="step-desc">Ustun va qator oralig'ini tanlang</div>
                </div>
              </div>
              <div className="step">
                <span className="step-number">3</span>
                <div className="step-content">
                  <div className="step-title">Data Matrix yarating</div>
                  <div className="step-desc">Tugmani bosing va kutib turing</div>
                </div>
              </div>
              <div className="step">
                <span className="step-number">4</span>
                <div className="step-content">
                  <div className="step-title">Fayllarni yuklab oling</div>
                  <div className="step-desc">Excel va ZIP formatda oling</div>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;