import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';
import { Upload, Download, FileText, FileSpreadsheet, Check, AlertCircle, Eye, Trash2, File } from 'lucide-react';

interface CertificateData {
  [key: string]: string | number;
}

interface UploadStatus {
  docx: boolean;
  excel: boolean;
}

interface SavedTemplate {
  id: string;
  name: string;
  html: string;
  binary: ArrayBuffer;
  placeholders: string[];
  uploadDate: string;
}

const CertificateGenerator: React.FC = () => {
  const [data, setData] = useState<CertificateData[]>([]);
  const [docxTemplate, setDocxTemplate] = useState<string>('');
  const [docxHtml, setDocxHtml] = useState<string>('');
  const [docxBinary, setDocxBinary] = useState<ArrayBuffer | null>(null);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [uploadStatus, setUploadStatus] = useState<UploadStatus>({ docx: false, excel: false });
  const [placeholders, setPlaceholders] = useState<string[]>([]);
  const [excelColumns, setExcelColumns] = useState<string[]>([]);
  const [savedTemplates, setSavedTemplates] = useState<SavedTemplate[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState<string | null>(null);
  const [templateName, setTemplateName] = useState<string>('');
  const [showSaveDialog, setShowSaveDialog] = useState(false);
  const [showRangeDialog, setShowRangeDialog] = useState(false);
  const [rangeStart, setRangeStart] = useState(1);
  const [rangeEnd, setRangeEnd] = useState(1);
  const certRef = useRef<HTMLDivElement>(null);

  // Load saved templates on mount
  React.useEffect(() => {
    const saved = localStorage.getItem('certificateTemplates');
    if (saved) {
      const templates = JSON.parse(saved);
      // Convert base64 back to ArrayBuffer
      const templatesWithBinary = templates.map((t: any) => ({
        ...t,
        binary: base64ToArrayBuffer(t.binary)
      }));
      setSavedTemplates(templatesWithBinary);
    }
  }, []);

  const arrayBufferToBase64 = (buffer: ArrayBuffer): string => {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  };

  const base64ToArrayBuffer = (base64: string): ArrayBuffer => {
    const binaryString = atob(base64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  };

  // Convert Excel serial date to readable date
  const excelDateToJSDate = (serial: number): string => {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    
    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    let total_seconds = Math.floor(86400 * fractional_day);
    const seconds = total_seconds % 60;
    total_seconds -= seconds;
    const hours = Math.floor(total_seconds / (60 * 60));
    const minutes = Math.floor(total_seconds / 60) % 60;
    
    const date = new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
    
    return date.toLocaleDateString('en-US', { 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    });
  };

  const processExcelData = (jsonData: CertificateData[]): CertificateData[] => {
    return jsonData.map(row => {
      const processed: CertificateData = {};
      Object.keys(row).forEach(key => {
        const value = row[key];
        // Check if value is a number that looks like an Excel date (between 1900 and 2100)
        if (typeof value === 'number' && value > 1 && value < 73050) {
          processed[key] = excelDateToJSDate(value);
        } else {
          processed[key] = value;
        }
      });
      return processed;
    });
  };

  const handleDocxUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      
      // Store the binary for later use
      setDocxBinary(arrayBuffer);
      
      // Convert to HTML for preview only
      const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer.slice(0) });
      const html = result.value;
      setDocxHtml(html);
      setDocxTemplate(html);
      
      // Extract placeholders from the DOCX
      const zip = new PizZip(arrayBuffer);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });
      
      const text = doc.getFullText();
      const placeholderRegex = /\{(\w+)\}|\{\{(\w+)\}\}|\[(\w+)\]/g;
      const found = new Set<string>();
      let match;
      
      while ((match = placeholderRegex.exec(text)) !== null) {
        const placeholder = match[1] || match[2] || match[3];
        if (placeholder) found.add(placeholder);
      }
      
      const foundPlaceholders = Array.from(found);
      setPlaceholders(foundPlaceholders);
      setUploadStatus(prev => ({ ...prev, docx: true }));
      setSelectedTemplateId(null);
      
      setTemplateName(file.name.replace('.docx', ''));
      setShowSaveDialog(true);
    } catch (error) {
      console.error('Error reading DOCX:', error);
      alert('Error reading DOCX file. Please try again.');
    }
  };

  const saveTemplate = () => {
    if (!templateName.trim()) {
      alert('Please enter a template name');
      return;
    }

    if (!docxBinary) {
      alert('No template to save');
      return;
    }

    const newTemplate: SavedTemplate = {
      id: Date.now().toString(),
      name: templateName,
      html: docxHtml,
      binary: docxBinary,
      placeholders: placeholders,
      uploadDate: new Date().toLocaleDateString()
    };

    const updated = [...savedTemplates, newTemplate];
    setSavedTemplates(updated);
    
    // Convert binary to base64 for localStorage
    const templatesForStorage = updated.map(t => ({
      ...t,
      binary: arrayBufferToBase64(t.binary)
    }));
    localStorage.setItem('certificateTemplates', JSON.stringify(templatesForStorage));
    
    setShowSaveDialog(false);
    setSelectedTemplateId(newTemplate.id);
  };

  const loadTemplate = (template: SavedTemplate) => {
    setDocxHtml(template.html);
    setDocxTemplate(template.html);
    setDocxBinary(template.binary);
    setPlaceholders(template.placeholders);
    setUploadStatus(prev => ({ ...prev, docx: true }));
    setSelectedTemplateId(template.id);
  };

  const deleteTemplate = (id: string) => {
    if (!confirm('Are you sure you want to delete this template?')) return;
    
    const updated = savedTemplates.filter(t => t.id !== id);
    setSavedTemplates(updated);
    
    const templatesForStorage = updated.map(t => ({
      ...t,
      binary: arrayBufferToBase64(t.binary)
    }));
    localStorage.setItem('certificateTemplates', JSON.stringify(templatesForStorage));
    
    if (selectedTemplateId === id) {
      setDocxHtml('');
      setDocxTemplate('');
      setDocxBinary(null);
      setPlaceholders([]);
      setUploadStatus(prev => ({ ...prev, docx: false }));
      setSelectedTemplateId(null);
    }
  };

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json(ws) as CertificateData[];
        
        if (jsonData.length > 0) {
          const processedData = processExcelData(jsonData);
          setData(processedData);
          setExcelColumns(Object.keys(jsonData[0]));
          setCurrentIndex(0);
          setRangeEnd(processedData.length);
          setUploadStatus(prev => ({ ...prev, excel: true }));
        }
      } catch (error) {
        console.error('Error reading Excel:', error);
        alert('Error reading Excel file. Please try again.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const mergeCertificate = (template: string, record: CertificateData): string => {
    let merged = template;
    
    Object.keys(record).forEach(key => {
      const value = record[key]?.toString() || '';
      merged = merged.replace(new RegExp(`\\{${key}\\}`, 'gi'), value);
      merged = merged.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'gi'), value);
      merged = merged.replace(new RegExp(`\\[${key}\\]`, 'gi'), value);
    });
    
    return merged;
  };

  const generateDocx = (record: CertificateData): Blob => {
    if (!docxBinary) throw new Error('No template loaded');
    
    const zip = new PizZip(docxBinary);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => '', // Return empty string instead of undefined
    });

    // Prepare data - match placeholders case-insensitively
    const templateData: { [key: string]: string } = {};
    
    placeholders.forEach(placeholder => {
      // Try to find matching column (case-insensitive)
      const matchingKey = Object.keys(record).find(
        key => key.toLowerCase() === placeholder.toLowerCase()
      );
      
      if (matchingKey) {
        templateData[placeholder] = record[matchingKey]?.toString() || '';
      } else {
        templateData[placeholder] = ''; // Empty if no match found
      }
    });

    doc.render(templateData);

    const blob = doc.getZip().generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });

    return blob;
  };

  const downloadDocx = (blob: Blob, filename: string) => {
    saveAs(blob, `${filename}.docx`);
  };

  const handleDownloadCurrent = () => {
    try {
      const blob = generateDocx(data[currentIndex]);
      const name = data[currentIndex].name || data[currentIndex].Name || `certificate_${currentIndex + 1}`;
      downloadDocx(blob, name.toString());
    } catch (error) {
      console.error('Error generating DOCX:', error);
      alert('Error generating certificate. Please check your template and data.');
    }
  };

  const handleDownloadAll = () => {
    data.forEach((record, idx) => {
      setTimeout(() => {
        try {
          const blob = generateDocx(record);
          const name = record.name || record.Name || `certificate_${idx + 1}`;
          downloadDocx(blob, name.toString());
        } catch (error) {
          console.error(`Error generating certificate ${idx + 1}:`, error);
        }
      }, idx * 500);
    });
  };

  const handleDownloadRange = () => {
    if (rangeStart < 1 || rangeEnd > data.length || rangeStart > rangeEnd) {
      alert('Invalid range. Please check your input.');
      return;
    }

    for (let i = rangeStart - 1; i < rangeEnd; i++) {
      setTimeout(() => {
        try {
          const blob = generateDocx(data[i]);
          const name = data[i].name || data[i].Name || `certificate_${i + 1}`;
          downloadDocx(blob, name.toString());
        } catch (error) {
          console.error(`Error generating certificate ${i + 1}:`, error);
        }
      }, (i - rangeStart + 1) * 500);
    }
    
    setShowRangeDialog(false);
  };

  const handlePrint = () => {
    const printWindow = window.open('', '', 'width=800,height=600');
    if (!printWindow || !certRef.current) return;

    printWindow.document.write(`
      <html>
        <head>
          <title>Certificate</title>
          <style>
            body { 
              margin: 20px; 
              font-family: 'Times New Roman', serif; 
            }
            @media print {
              body { margin: 0; }
            }
          </style>
        </head>
        <body>
          ${certRef.current.innerHTML}
        </body>
      </html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  const handlePrintAll = () => {
    const printWindow = window.open('', '', 'width=800,height=600');
    if (!printWindow) return;

    const allCertificates = data.map((record, idx) => {
      const merged = mergeCertificate(docxHtml, record);
      return `
        <div style="page-break-after: ${idx < data.length - 1 ? 'always' : 'auto'};">
          ${merged}
        </div>
      `;
    }).join('');

    printWindow.document.write(`
      <html>
        <head>
          <title>All Certificates</title>
          <style>
            body { 
              margin: 20px; 
              font-family: 'Times New Roman', serif; 
            }
            @media print {
              body { margin: 0; }
            }
          </style>
        </head>
        <body>
          ${allCertificates}
        </body>
      </html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  const isReady = uploadStatus.docx && uploadStatus.excel;

  return (
    <div className="flex min-h-screen bg-gradient-to-br from-purple-50 to-blue-100">
      {/* Sidebar */}
      <div className="w-80 bg-white shadow-xl p-6 overflow-y-auto">
        <h2 className="text-2xl font-bold text-gray-800 mb-6">📁 Saved Templates</h2>
        
        {/* Saved Templates List */}
        {savedTemplates.length > 0 ? (
          <div className="space-y-3 mb-8">
            {savedTemplates.map(template => (
              <div 
                key={template.id}
                className={`p-4 rounded-lg border-2 cursor-pointer transition ${
                  selectedTemplateId === template.id 
                    ? 'bg-purple-50 border-purple-500' 
                    : 'bg-gray-50 border-gray-300 hover:border-purple-300'
                }`}
                onClick={() => loadTemplate(template)}
              >
                <div className="flex items-start justify-between mb-2">
                  <div className="flex-1">
                    <div className="font-semibold text-gray-800 flex items-center gap-2">
                      <File className="w-4 h-4" />
                      {template.name}
                    </div>
                    <div className="text-xs text-gray-500 mt-1">
                      {template.placeholders.length} fields • {template.uploadDate}
                    </div>
                  </div>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      deleteTemplate(template.id);
                    }}
                    className="text-red-500 hover:text-red-700 p-1"
                    title="Delete template"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
            ))}
          </div>
        ) : (
          <div className="bg-gray-50 p-4 rounded-lg mb-8 text-center text-gray-500 text-sm">
            No saved templates yet. Upload a DOCX to get started!
          </div>
        )}

        <div className="border-t pt-6 mb-6"></div>

        <h2 className="text-xl font-bold text-gray-800 mb-4">📊 Current Status</h2>
        
        {/* Upload Status */}
        <div className="mb-8">
          <div className={`p-5 rounded-lg border-2 ${isReady ? 'bg-green-50 border-green-500' : 'bg-gray-50 border-gray-300'}`}>
            <div className="flex items-center gap-2 mb-3">
              {isReady ? (
                <Check className="w-6 h-6 text-green-600" />
              ) : (
                <Upload className="w-6 h-6 text-gray-400" />
              )}
              <span className="font-semibold text-lg">Files Status</span>
            </div>
            <div className="space-y-2">
              <div className="flex items-center gap-2 text-sm">
                {uploadStatus.docx ? (
                  <Check className="w-4 h-4 text-green-600" />
                ) : (
                  <div className="w-4 h-4 border-2 border-gray-300 rounded" />
                )}
                <FileText className="w-4 h-4 text-gray-500" />
                <span className="text-gray-700">DOCX Template</span>
              </div>
              <div className="flex items-center gap-2 text-sm">
                {uploadStatus.excel ? (
                  <Check className="w-4 h-4 text-green-600" />
                ) : (
                  <div className="w-4 h-4 border-2 border-gray-300 rounded" />
                )}
                <FileSpreadsheet className="w-4 h-4 text-gray-500" />
                <span className="text-gray-700">Excel Data {uploadStatus.excel && `(${data.length} records)`}</span>
              </div>
            </div>
          </div>
        </div>

        {/* Placeholders Found */}
        {placeholders.length > 0 && (
          <div className="mb-8">
            <h3 className="font-bold text-gray-800 mb-3 flex items-center gap-2">
              <AlertCircle className="w-5 h-5" />
              Placeholders Found
            </h3>
            <div className="space-y-2">
              {placeholders.map(ph => (
                <div key={ph} className="bg-blue-50 px-3 py-2 rounded text-sm">
                  <code className="text-blue-700">{`{${ph}}`}</code>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Excel Columns */}
        {excelColumns.length > 0 && (
          <div className="mb-8">
            <h3 className="font-bold text-gray-800 mb-3 flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5" />
              Excel Columns
            </h3>
            <div className="space-y-2">
              {excelColumns.map(col => (
                <div key={col} className="bg-green-50 px-3 py-2 rounded text-sm">
                  <span className="text-green-700 font-medium">{col}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Field Mapping Info */}
        {isReady && (
          <div className="bg-purple-50 p-4 rounded-lg border border-purple-200">
            <h3 className="font-bold text-purple-900 mb-2">💡 Auto-Mapping</h3>
            <p className="text-sm text-purple-700">
              Placeholders are matched with Excel columns. Date serials are auto-converted!
            </p>
          </div>
        )}
      </div>

      {/* Main Content */}
      <div className="flex-1 p-8 overflow-y-auto">
        <div className="max-w-6xl mx-auto">
          {/* Header */}
          <div className="bg-white rounded-lg shadow-xl p-8 mb-8">
            <h1 className="text-3xl font-bold text-gray-800 mb-6">🎓 Certificate Generator</h1>
            
            {/* Upload Areas */}
            <div className="grid grid-cols-2 gap-6">
              {/* DOCX Upload */}
              <div>
                <label className="flex flex-col items-center justify-center h-40 px-4 transition bg-white border-2 border-dashed rounded-lg cursor-pointer hover:border-purple-500 border-gray-300">
                  <div className="flex flex-col items-center space-y-2">
                    <FileText className="w-10 h-10 text-gray-400" />
                    <span className="font-medium text-gray-600">Upload New DOCX Template</span>
                    <span className="text-xs text-gray-500">
                      {uploadStatus.docx ? '✓ Uploaded' : 'Click to select file'}
                    </span>
                  </div>
                  <input 
                    type="file" 
                    className="hidden" 
                    accept=".docx" 
                    onChange={handleDocxUpload} 
                  />
                </label>
              </div>

              {/* Excel Upload */}
              <div>
                <label className="flex flex-col items-center justify-center h-40 px-4 transition bg-white border-2 border-dashed rounded-lg cursor-pointer hover:border-blue-500 border-gray-300">
                  <div className="flex flex-col items-center space-y-2">
                    <Upload className="w-10 h-10 text-gray-400" />
                    <span className="font-medium text-gray-600">Upload Excel Data</span>
                    <span className="text-xs text-gray-500">
                      {uploadStatus.excel ? `✓ ${data.length} records` : 'Click to select file'}
                    </span>
                  </div>
                  <input 
                    type="file" 
                    className="hidden" 
                    accept=".xlsx,.xls" 
                    onChange={handleExcelUpload} 
                  />
                </label>
              </div>
            </div>
          </div>

          {/* Navigation & Preview */}
          {isReady && (
            <>
              <div className="bg-white rounded-lg shadow-xl p-6 mb-8">
                <div className="flex items-center justify-between mb-4">
                  <h2 className="text-xl font-semibold text-gray-800 flex items-center gap-2">
                    <Eye className="w-5 h-5" />
                    Preview
                  </h2>
                  <div className="flex gap-2 items-center">
                    <button
                      onClick={() => setCurrentIndex(Math.max(0, currentIndex - 1))}
                      disabled={currentIndex === 0}
                      className="px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 disabled:opacity-50 disabled:cursor-not-allowed transition"
                    >
                      ← Previous
                    </button>
                    <span className="px-4 py-2 bg-purple-100 text-purple-800 rounded font-medium">
                      {currentIndex + 1} / {data.length}
                    </span>
                    <button
                      onClick={() => setCurrentIndex(Math.min(data.length - 1, currentIndex + 1))}
                      disabled={currentIndex === data.length - 1}
                      className="px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 disabled:opacity-50 disabled:cursor-not-allowed transition"
                    >
                      Next →
                    </button>
                  </div>
                </div>

                {/* Certificate Preview */}
                <div className="bg-white border-4 border-gray-200 rounded-lg overflow-auto p-8 max-h-[600px]">
                  <div 
                    ref={certRef}
                    className="certificate-preview mx-auto"
                    dangerouslySetInnerHTML={{ 
                      __html: mergeCertificate(docxHtml, data[currentIndex]) 
                    }}
                  />
                </div>

                {/* Action Buttons */}
                <div className="mt-6 space-y-4">
                  {/* Download Buttons */}
                  <div className="grid grid-cols-3 gap-4">
                    <button
                      onClick={handleDownloadCurrent}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition shadow-lg"
                    >
                      <Download className="w-5 h-5" />
                      Download Current
                    </button>
                    <button
                      onClick={() => setShowRangeDialog(true)}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition shadow-lg"
                    >
                      <Download className="w-5 h-5" />
                      Download Range
                    </button>
                    <button
                      onClick={handleDownloadAll}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-green-600 text-white rounded-lg hover:bg-green-700 transition shadow-lg"
                    >
                      <Download className="w-5 h-5" />
                      Download All ({data.length})
                    </button>
                  </div>

                  {/* Print Buttons */}
                  <div className="grid grid-cols-2 gap-4">
                    <button
                      onClick={handlePrint}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition shadow-lg"
                    >
                      <Download className="w-5 h-5" />
                      Print Current
                    </button>
                    <button
                      onClick={handlePrintAll}
                      className="flex items-center justify-center gap-2 px-4 py-3 bg-purple-700 text-white rounded-lg hover:bg-purple-800 transition shadow-lg"
                    >
                      <Download className="w-5 h-5" />
                      Print All ({data.length})
                    </button>
                  </div>
                </div>
              </div>
            </>
          )}

          {/* Instructions */}
          {!isReady && (
            <div className="bg-white rounded-lg shadow-xl p-8">
              <h2 className="text-xl font-bold text-gray-800 mb-4">📝 How to Use</h2>
              <ol className="space-y-3 text-gray-700">
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">1.</span>
                  <span>Upload your DOCX template with placeholders like <code className="bg-gray-100 px-2 py-1 rounded">{`{name}`}</code>, <code className="bg-gray-100 px-2 py-1 rounded">{`{degree}`}</code></span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">2.</span>
                  <span>Upload your Excel file with matching column names (name, degree, etc.)</span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">3.</span>
                  <span>Preview the merged certificates and navigate through records</span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">4.</span>
                  <span>Download or print certificates individually, by range, or all at once</span>
                </li>
              </ol>
            </div>
          )}
        </div>
      </div>

      {/* Save Template Dialog */}
      {showSaveDialog && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-96 shadow-2xl">
            <h3 className="text-xl font-bold mb-4">Save Template</h3>
            <input
              type="text"
              value={templateName}
              onChange={(e) => setTemplateName(e.target.value)}
              placeholder="Enter template name"
              className="w-full px-4 py-2 border border-gray-300 rounded-lg mb-4 focus:outline-none focus:ring-2 focus:ring-purple-500"
              autoFocus
            />
            <div className="flex gap-3">
              <button
                onClick={saveTemplate}
                className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition"
              >
                Save
              </button>
              <button
                onClick={() => setShowSaveDialog(false)}
                className="flex-1 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Download Range Dialog */}
      {showRangeDialog && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-96 shadow-2xl">
            <h3 className="text-xl font-bold mb-4">Download Range</h3>
            <div className="space-y-4 mb-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">From Record:</label>
                <input
                  type="number"
                  min="1"
                  max={data.length}
                  value={rangeStart}
                  onChange={(e) => setRangeStart(parseInt(e.target.value) || 1)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">To Record:</label>
                <input
                  type="number"
                  min="1"
                  max={data.length}
                  value={rangeEnd}
                  onChange={(e) => setRangeEnd(parseInt(e.target.value) || data.length)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                />
              </div>
              <p className="text-sm text-gray-600">
                Total records: {data.length} | Selected: {Math.max(0, rangeEnd - rangeStart + 1)}
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={handleDownloadRange}
                className="flex-1 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition"
              >
                Download
              </button>
              <button
                onClick={() => setShowRangeDialog(false)}
                className="flex-1 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default CertificateGenerator;