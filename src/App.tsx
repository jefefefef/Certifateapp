import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';
import { Upload, Download, FileText, FileSpreadsheet, Check, AlertCircle, Eye, Trash2, File, FolderOpen } from 'lucide-react';

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
  fileName: string;
  html: string;
  binary: ArrayBuffer;
  placeholders: string[];
  uploadDate: string;
}

// IndexedDB for storing file handles
const DB_NAME = 'CertificateGeneratorDB';
const DB_VERSION = 2;
const HANDLES_STORE = 'fileHandles';

const openDB = (): Promise<IDBDatabase> => {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);
    
    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains(HANDLES_STORE)) {
        db.createObjectStore(HANDLES_STORE, { keyPath: 'id' });
      }
    };
  });
};

const saveFileHandle = async (id: string, handle: FileSystemFileHandle): Promise<void> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([HANDLES_STORE], 'readwrite');
    const store = transaction.objectStore(HANDLES_STORE);
    const request = store.put({ id, handle });
    
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
};

const getFileHandle = async (id: string): Promise<FileSystemFileHandle | null> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([HANDLES_STORE], 'readonly');
    const store = transaction.objectStore(HANDLES_STORE);
    const request = store.get(id);
    
    request.onsuccess = () => resolve(request.result?.handle || null);
    request.onerror = () => reject(request.error);
  });
};

const getAllFileHandles = async (): Promise<{ id: string; handle: FileSystemFileHandle }[]> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([HANDLES_STORE], 'readonly');
    const store = transaction.objectStore(HANDLES_STORE);
    const request = store.getAll();
    
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
};

const deleteFileHandle = async (id: string): Promise<void> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([HANDLES_STORE], 'readwrite');
    const store = transaction.objectStore(HANDLES_STORE);
    const request = store.delete(id);
    
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
};

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
  const [showRangeDialog, setShowRangeDialog] = useState(false);
  const [rangeStart, setRangeStart] = useState(1);
  const [rangeEnd, setRangeEnd] = useState(1);
  const [excelFileName, setExcelFileName] = useState<string>('');
  const [hasFileSystemSupport, setHasFileSystemSupport] = useState(false);
  const certRef = useRef<HTMLDivElement>(null);

  // Check for File System Access API support
  React.useEffect(() => {
    const supported = 'showOpenFilePicker' in window;
    setHasFileSystemSupport(supported);
    if (!supported) {
      console.warn('⚠️ File System Access API not supported in this browser');
    }
  }, []);

  // Load saved file handles on mount
  React.useEffect(() => {
    const loadSavedFiles = async () => {
      if (!hasFileSystemSupport) {
        console.log('⚠️ File System Access API not available');
        return;
      }

      try {
        console.log('🔍 Loading saved file handles...');
        const handles = await getAllFileHandles();
        console.log(`📁 Found ${handles.length} saved file(s)`);

        for (const { id, handle } of handles) {
          try {
            // Request permission to read the file
            const permission = await handle.queryPermission({ mode: 'read' });
            
            if (permission === 'granted' || permission === 'prompt') {
              if (id.startsWith('template_')) {
                await loadTemplateFromHandle(handle, id);
              } else if (id === 'excel') {
                await loadExcelFromHandle(handle);
              }
            } else {
              console.log(`⚠️ No permission for file: ${handle.name}`);
            }
          } catch (error) {
            console.error(`❌ Error loading file ${id}:`, error);
          }
        }
      } catch (error) {
        console.error('❌ Error loading saved files:', error);
      }
    };

    if (hasFileSystemSupport) {
      loadSavedFiles();
    }
  }, [hasFileSystemSupport]);

  const loadTemplateFromHandle = async (handle: FileSystemFileHandle, id: string) => {
    try {
      const file = await handle.getFile();
      const arrayBuffer = await file.arrayBuffer();
      
      // Convert to HTML for preview
      const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer.slice(0) });
      const html = result.value;
      
      // Extract placeholders
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
      
      const template: SavedTemplate = {
        id,
        name: file.name.replace('.docx', ''),
        fileName: file.name,
        html,
        binary: arrayBuffer,
        placeholders: Array.from(found),
        uploadDate: new Date(file.lastModified).toLocaleDateString()
      };

      setSavedTemplates(prev => {
        const filtered = prev.filter(t => t.id !== id);
        return [...filtered, template];
      });

      // Auto-load first template
      if (!docxBinary) {
        setDocxHtml(html);
        setDocxTemplate(html);
        setDocxBinary(arrayBuffer);
        setPlaceholders(Array.from(found));
        setUploadStatus(prev => ({ ...prev, docx: true }));
        setSelectedTemplateId(id);
        console.log(`✅ Loaded template: ${file.name}`);
      }
    } catch (error) {
      console.error('Error loading template from handle:', error);
    }
  };

  const loadExcelFromHandle = async (handle: FileSystemFileHandle) => {
    try {
      const file = await handle.getFile();
      const arrayBuffer = await file.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: 'array' });
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
        setExcelFileName(file.name);
        console.log(`✅ Loaded Excel: ${file.name} (${processedData.length} records)`);
      }
    } catch (error) {
      console.error('Error loading Excel from handle:', error);
    }
  };

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
        if (typeof value === 'number' && value > 1 && value < 73050) {
          processed[key] = excelDateToJSDate(value);
        } else {
          processed[key] = value;
        }
      });
      return processed;
    });
  };

  const handleDocxUpload = async () => {
    if (!hasFileSystemSupport) {
      alert('File System Access API is not supported in your browser. Please use Chrome, Edge, or another Chromium-based browser.');
      return;
    }

    try {
      const [fileHandle] = await (window as any).showOpenFilePicker({
        types: [{
          description: 'Word Documents',
          accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] }
        }],
        multiple: false
      });

      const id = `template_${Date.now()}`;
      await saveFileHandle(id, fileHandle);
      await loadTemplateFromHandle(fileHandle, id);
      
      console.log('✅ Template saved with file system access');
    } catch (error) {
      if ((error as Error).name !== 'AbortError') {
        console.error('Error selecting DOCX:', error);
        alert('Error selecting file. Please try again.');
      }
    }
  };

  const handleExcelUpload = async () => {
    if (!hasFileSystemSupport) {
      alert('File System Access API is not supported in your browser. Please use Chrome, Edge, or another Chromium-based browser.');
      return;
    }

    try {
      const [fileHandle] = await (window as any).showOpenFilePicker({
        types: [{
          description: 'Excel Files',
          accept: { 
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
            'application/vnd.ms-excel': ['.xls']
          }
        }],
        multiple: false
      });

      await saveFileHandle('excel', fileHandle);
      await loadExcelFromHandle(fileHandle);
      
      console.log('✅ Excel saved with file system access');
    } catch (error) {
      if ((error as Error).name !== 'AbortError') {
        console.error('Error selecting Excel:', error);
        alert('Error selecting file. Please try again.');
      }
    }
  };

  const loadTemplate = (template: SavedTemplate) => {
    setDocxHtml(template.html);
    setDocxTemplate(template.html);
    setDocxBinary(template.binary);
    setPlaceholders(template.placeholders);
    setUploadStatus(prev => ({ ...prev, docx: true }));
    setSelectedTemplateId(template.id);
  };

  const deleteTemplate = async (id: string) => {
    if (!confirm('Are you sure you want to delete this template?')) return;
    
    try {
      await deleteFileHandle(id);
      const updated = savedTemplates.filter(t => t.id !== id);
      setSavedTemplates(updated);
      
      if (selectedTemplateId === id) {
        setDocxHtml('');
        setDocxTemplate('');
        setDocxBinary(null);
        setPlaceholders([]);
        setUploadStatus(prev => ({ ...prev, docx: false }));
        setSelectedTemplateId(null);
      }
      console.log('✅ Template deleted');
    } catch (error) {
      console.error('❌ Error deleting template:', error);
      alert('Error deleting template. Please try again.');
    }
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
      nullGetter: () => '',
    });

    const templateData: { [key: string]: string } = {};
    
    placeholders.forEach(placeholder => {
      const matchingKey = Object.keys(record).find(
        key => key.toLowerCase() === placeholder.toLowerCase()
      );
      
      if (matchingKey) {
        templateData[placeholder] = record[matchingKey]?.toString() || '';
      } else {
        templateData[placeholder] = '';
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
        
        {!hasFileSystemSupport && (
          <div className="bg-yellow-50 border border-yellow-300 rounded-lg p-4 mb-6">
            <p className="text-sm text-yellow-800">
              ⚠️ File System Access API not supported. Please use Chrome, Edge, or another Chromium-based browser for persistent file access.
            </p>
          </div>
        )}

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
            No saved templates yet. Select a DOCX file to get started!
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
                <span className="text-gray-700">
                  {excelFileName || 'Excel Data'} {uploadStatus.excel && `(${data.length} records)`}
                </span>
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
                <button
                  onClick={handleDocxUpload}
                  className="flex flex-col items-center justify-center w-full h-40 px-4 transition bg-white border-2 border-dashed rounded-lg hover:border-purple-500 border-gray-300"
                >
                  <div className="flex flex-col items-center space-y-2">
                    <FolderOpen className="w-10 h-10 text-gray-400" />
                    <span className="font-medium text-gray-600">Select DOCX Template</span>
                    <span className="text-xs text-gray-500">
                      {uploadStatus.docx ? '✓ Template linked' : 'File will be remembered'}
                    </span>
                  </div>
                </button>
              </div>

              {/* Excel Upload */}
              <div>
                <button
                  onClick={handleExcelUpload}
                  className="flex flex-col items-center justify-center w-full h-40 px-4 transition bg-white border-2 border-dashed rounded-lg hover:border-blue-500 border-gray-300"
                >
                  <div className="flex flex-col items-center space-y-2">
                    <FolderOpen className="w-10 h-10 text-gray-400" />
                    <span className="font-medium text-gray-600">Select Excel Data</span>
                    <span className="text-xs text-gray-500">
                      {uploadStatus.excel ? `✓ ${data.length} records` : 'File will be remembered'}
                    </span>
                  </div>
                </button>
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
                  <span>Click to select your DOCX template with placeholders like <code className="bg-gray-100 px-2 py-1 rounded">{`{name}`}</code>, <code className="bg-gray-100 px-2 py-1 rounded">{`{degree}`}</code></span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">2.</span>
                  <span>Click to select your Excel file with matching column names</span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">3.</span>
                  <span>Grant permission when prompted - files will be remembered!</span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">4.</span>
                  <span>Next time you open the app, your files will load automatically</span>
                </li>
              </ol>
            </div>
          )}
        </div>
      </div>

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