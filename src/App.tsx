import React, { useState, useRef } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { saveAs } from "file-saver";
import {
  Upload,
  Download,
  FileText,
  FileSpreadsheet,
  Check,
  AlertCircle,
  Eye,
  Trash2,
  File,
  Filter,
} from "lucide-react";

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
  binary: string; // base64
  placeholders: string[];
  uploadDate: string;
}

interface FilterCondition {
  column: string;
  value: string;
}

// Simple IndexedDB operations
const DB_NAME = "CertGenDB";
const DB_VERSION = 2; // Increment version to trigger upgrade
const STORE_NAME = "templates";
const EXCEL_STORE = "excelData";

const openDB = (): Promise<IDBDatabase> => {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);
    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(EXCEL_STORE)) {
        db.createObjectStore(EXCEL_STORE, { keyPath: "id" });
      }
    };
  });
};

const saveTemplate = async (template: SavedTemplate): Promise<void> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    try {
      const transaction = db.transaction([STORE_NAME], "readwrite");
      const store = transaction.objectStore(STORE_NAME);

      // Check size before saving
      const estimatedSize =
        template.binary.length + template.html.length + template.name.length;
      console.log(
        `📊 Estimated template size: ${(estimatedSize / 1024 / 1024).toFixed(2)} MB`,
      );

      if (estimatedSize > 5 * 1024 * 1024) {
        // 5MB limit
        reject(
          new Error("Template is too large (>5MB). Try a simpler DOCX file."),
        );
        return;
      }

      const request = store.put(template);

      request.onsuccess = () => {
        console.log("✅ IndexedDB write successful");
        resolve();
      };

      request.onerror = () => {
        console.error("❌ IndexedDB write error:", request.error);
        reject(request.error);
      };

      transaction.oncomplete = () => {
        console.log("✅ Transaction complete");
      };

      transaction.onerror = () => {
        console.error("❌ Transaction error:", transaction.error);
        reject(transaction.error);
      };
    } catch (error) {
      console.error("❌ Exception in saveTemplate:", error);
      reject(error);
    }
  });
};

const getAllTemplates = async (): Promise<SavedTemplate[]> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([STORE_NAME], "readonly");
    const store = transaction.objectStore(STORE_NAME);
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
};

const deleteTemplate = async (id: string): Promise<void> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([STORE_NAME], "readwrite");
    const store = transaction.objectStore(STORE_NAME);
    const request = store.delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
};

const saveExcelData = async (
  data: CertificateData[],
  columns: string[],
): Promise<void> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([EXCEL_STORE], "readwrite");
    const store = transaction.objectStore(EXCEL_STORE);
    const request = store.put({ id: "excel", data, columns });
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
};

// remove stored excel record
const deleteExcelData = async (): Promise<void> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([EXCEL_STORE], "readwrite");
    const store = transaction.objectStore(EXCEL_STORE);
    const request = store.delete("excel");
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
};

const getExcelData = async (): Promise<{
  data: CertificateData[];
  columns: string[];
} | null> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([EXCEL_STORE], "readonly");
    const store = transaction.objectStore(EXCEL_STORE);
    const request = store.get("excel");
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
};

const CertificateGenerator: React.FC = () => {
  const [data, setData] = useState<CertificateData[]>([]);
  const [docxHtml, setDocxHtml] = useState<string>("");
  const [docxBinary, setDocxBinary] = useState<ArrayBuffer | null>(null);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [uploadStatus, setUploadStatus] = useState<UploadStatus>({
    docx: false,
    excel: false,
  });
  const [placeholders, setPlaceholders] = useState<string[]>([]);
  const [excelColumns, setExcelColumns] = useState<string[]>([]);
  const [savedTemplates, setSavedTemplates] = useState<SavedTemplate[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState<string | null>(
    null,
  );
  const [templateName, setTemplateName] = useState<string>("");
  const [showSaveDialog, setShowSaveDialog] = useState(false);
  const [showRangeDialog, setShowRangeDialog] = useState(false);
  const [rangeStart, setRangeStart] = useState(1);
  const [rangeEnd, setRangeEnd] = useState(1);
  const certRef = useRef<HTMLDivElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);
  
  // Filter state
  const [filterColumn, setFilterColumn] = useState<string>("");
  const [filterValue, setFilterValue] = useState<string>("");
  const [uniqueValues, setUniqueValues] = useState<string[]>([]);
  const [filteredData, setFilteredData] = useState<CertificateData[]>([]);
  const [isFiltered, setIsFiltered] = useState(false);
  const [filteredIndex, setFilteredIndex] = useState(0);
  const [showFilterModal, setShowFilterModal] = useState(false);
  const [filterConditions, setFilterConditions] = useState<FilterCondition[]>([]);
  const [tempFilterConditions, setTempFilterConditions] = useState<FilterCondition[]>([]);
  
  // Excel Editor state
  const [showExcelEditor, setShowExcelEditor] = useState(false);
  const [editableData, setEditableData] = useState<CertificateData[]>([]);
  const [newRowData, setNewRowData] = useState<CertificateData>({});

  // Load templates on mount
  React.useEffect(() => {
    const loadData = async () => {
      try {
        // Load templates
        const templates = await getAllTemplates();
        console.log(`✅ Loaded ${templates.length} template(s)`);

        setSavedTemplates(templates);

        // Auto-load first template
        if (templates.length > 0) {
          const first = templates[0];
          setDocxHtml(first.html);
          setDocxBinary(base64ToArrayBuffer(first.binary));
          setPlaceholders(first.placeholders);
          setUploadStatus((prev) => ({ ...prev, docx: true }));
          setSelectedTemplateId(first.id);
          console.log(`🎯 Auto-loaded template: ${first.name}`);
        }

        // Load Excel data
        const excelData = await getExcelData();
        if (excelData) {
          setData(excelData.data);
          setEditableData(excelData.data);
          setExcelColumns(excelData.columns);
          setCurrentIndex(0);
          setRangeEnd(excelData.data.length);
          setUploadStatus((prev) => ({ ...prev, excel: true }));
          console.log(`✅ Loaded Excel: ${excelData.data.length} records`);
        }
      } catch (error) {
        console.error("Error loading data:", error);
      }
    };
    loadData();
  }, []);

  const arrayBufferToBase64 = (buffer: ArrayBuffer): string => {
    const bytes = new Uint8Array(buffer);
    let binary = "";
    const chunkSize = 0x8000; // 32KB chunks
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
      binary += String.fromCharCode.apply(null, Array.from(chunk));
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

    const date = new Date(
      date_info.getFullYear(),
      date_info.getMonth(),
      date_info.getDate(),
      hours,
      minutes,
      seconds,
    );

    return date.toLocaleDateString("en-US", {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
  };

  const processExcelData = (jsonData: CertificateData[]): CertificateData[] => {
    return jsonData.map((row) => {
      const processed: CertificateData = {};
      Object.keys(row).forEach((key) => {
        const value = row[key];
        if (typeof value === "number" && value > 1 && value < 73050) {
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
      setDocxBinary(arrayBuffer);

      // Convert to HTML for preview
      const result = await mammoth.convertToHtml({
        arrayBuffer: arrayBuffer.slice(0),
      });
      const html = result.value;
      setDocxHtml(html);

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

      const foundPlaceholders = Array.from(found);
      setPlaceholders(foundPlaceholders);
      setUploadStatus((prev) => ({ ...prev, docx: true }));
      setSelectedTemplateId(null);

      setTemplateName(file.name.replace(".docx", ""));
      setShowSaveDialog(true);
    } catch (error) {
      console.error("Error reading DOCX:", error);
      alert("Error reading DOCX file. Please try again.");
    }
  };

  const handleSaveTemplate = async () => {
    if (!templateName.trim()) {
      alert("Please enter a template name");
      return;
    }
    if (!docxBinary) {
      alert("No template to save");
      return;
    }

    try {
      console.log("💾 Starting to save template...");
      console.log("Template name:", templateName);
      console.log("Binary size:", docxBinary.byteLength, "bytes");

      // Convert to base64
      console.log("🔄 Converting to base64...");
      const base64 = arrayBufferToBase64(docxBinary);
      console.log("✅ Base64 length:", base64.length);

      const newTemplate: SavedTemplate = {
        id: Date.now().toString(),
        name: templateName,
        html: docxHtml,
        binary: base64,
        placeholders: placeholders,
        uploadDate: new Date().toLocaleDateString(),
      };

      console.log("💾 Saving to IndexedDB...");
      await saveTemplate(newTemplate);
      console.log("✅ Saved to IndexedDB successfully");

      setSavedTemplates((prev) => [...prev, newTemplate]);
      setShowSaveDialog(false);
      setSelectedTemplateId(newTemplate.id);

      console.log("✅ Template saved completely!");
      alert("Template saved successfully!");
    } catch (error) {
      console.error("❌ Error saving template:", error);
      alert(`Error saving template: ${error}`);
    }
  };

  const handleLoadTemplate = (template: SavedTemplate) => {
    setDocxHtml(template.html);
    setDocxBinary(base64ToArrayBuffer(template.binary));
    setPlaceholders(template.placeholders);
    setUploadStatus((prev) => ({ ...prev, docx: true }));
    setSelectedTemplateId(template.id);
  };

  const handleDeleteTemplate = async (id: string) => {
    if (!confirm("Are you sure you want to delete this template?")) return;

    try {
      await deleteTemplate(id);
      setSavedTemplates((prev) => prev.filter((t) => t.id !== id));

      if (selectedTemplateId === id) {
        setDocxHtml("");
        setDocxBinary(null);
        setPlaceholders([]);
        setUploadStatus((prev) => ({ ...prev, docx: false }));
        setSelectedTemplateId(null);
      }
      console.log("✅ Template deleted");
    } catch (error) {
      console.error("Error deleting template:", error);
    }
  };

  const handleDeleteExcel = async () => {
    if (!confirm("Are you sure you want to remove the uploaded Excel data?"))
      return;

    try {
      await deleteExcelData();
      // reset related state
      setData([]);
      setExcelColumns([]);
      setUploadStatus((prev) => ({ ...prev, excel: false }));
      setFilterColumn("");
      setFilterValue("");
      setUniqueValues([]);
      setCurrentIndex(0);
      setRangeStart(1);
      setRangeEnd(1);
      // Reset filter state
      setIsFiltered(false);
      setFilterColumn("");
      setFilterValue("");
      setFilteredData([]);
      setFilterConditions([]);
      setTempFilterConditions([]);
      // Reset editor state
      setEditableData([]);
      setNewRowData({});
      if (excelInputRef.current) {
        excelInputRef.current.value = "";
      }
      console.log("✅ Excel data deleted");
    } catch (error) {
      console.error("Error deleting excel data:", error);
    }
  };

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json(ws) as CertificateData[];

        if (jsonData.length > 0) {
          const processedData = processExcelData(jsonData);
          const columns = Object.keys(jsonData[0]);

          setData(processedData);
          setEditableData(processedData); // Initialize editor data
          setExcelColumns(columns);
          setCurrentIndex(0);
          setRangeEnd(processedData.length);
          setUploadStatus((prev) => ({ ...prev, excel: true }));

          // Reset filter state when new Excel data is loaded
          setIsFiltered(false);
          setFilterColumn("");
          setFilterValue("");
          setFilteredData([]);
          setFilterConditions([]);
          setTempFilterConditions([]);

          // Save Excel data to IndexedDB
          await saveExcelData(processedData, columns);
          console.log(
            `✅ Loaded and saved Excel: ${processedData.length} records`,
          );
        }
      } catch (error) {
        console.error("Error reading Excel:", error);
        alert("Error reading Excel file. Please try again.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const mergeCertificate = (
    template: string,
    record: CertificateData,
  ): string => {
    let merged = template;
    Object.keys(record).forEach((key) => {
      const value = record[key]?.toString() || "";
      merged = merged.replace(new RegExp(`\\{${key}\\}`, "gi"), value);
      merged = merged.replace(new RegExp(`\\{\\{${key}\\}\\}`, "gi"), value);
      merged = merged.replace(new RegExp(`\\[${key}\\]`, "gi"), value);
    });
    return merged;
  };

  const generateDocx = (record: CertificateData): Blob => {
    if (!docxBinary) throw new Error("No template loaded");

    const zip = new PizZip(docxBinary);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => "",
    });

    const templateData: { [key: string]: string } = {};

    // Add special date placeholders
    const now = new Date();
    templateData["TODAY"] = now.toLocaleDateString("en-US", {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    templateData["DATE"] = now.toLocaleDateString("en-US", {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    templateData["DATE_SHORT"] = now.toLocaleDateString("en-US");
    templateData["DATE_ISO"] = now.toISOString().split("T")[0]; // 2025-12-03 format
    templateData["YEAR"] = now.getFullYear().toString();
    templateData["MONTH"] = (now.getMonth() + 1).toString().padStart(2, "0");
    templateData["DAY"] = now.getDate().toString().padStart(2, "0");

    // Process all placeholders
    placeholders.forEach((placeholder) => {
      // Check if placeholder has _UPPER modifier
      if (placeholder.endsWith("_UPPER")) {
        const baseName = placeholder.replace("_UPPER", "");
        const matchingKey = Object.keys(record).find(
          (key) => key.toLowerCase() === baseName.toLowerCase(),
        );
        if (matchingKey) {
          templateData[placeholder] =
            record[matchingKey]?.toString().toUpperCase() || "";
        }
      } else {
        // Normal placeholder
        const matchingKey = Object.keys(record).find(
          (key) => key.toLowerCase() === placeholder.toLowerCase(),
        );
        if (matchingKey) {
          templateData[placeholder] = record[matchingKey]?.toString() || "";
        }
      }
    });

    doc.render(templateData);
    return doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
  };

  const handleDownloadCurrent = () => {
    try {
      const record = isFiltered && filteredData.length > 0 
        ? filteredData[filteredIndex] 
        : data[currentIndex];
      const blob = generateDocx(record);
      const name =
        record.name ||
        record.Name ||
        `certificate_${isFiltered ? filteredIndex + 1 : currentIndex + 1}`;
      saveAs(blob, `${name}.docx`);
    } catch (error) {
      console.error("Error generating DOCX:", error);
      alert("Error generating certificate.");
    }
  };

  const handleDownloadAll = () => {
    const dataToDownload = isFiltered && filteredData.length > 0 ? filteredData : data;
    
    dataToDownload.forEach((record, idx) => {
      setTimeout(() => {
        try {
          const blob = generateDocx(record);
          const name = record.name || record.Name || `certificate_${idx + 1}`;
          saveAs(blob, `${name}.docx`);
        } catch (error) {
          console.error(`Error generating certificate ${idx + 1}:`, error);
        }
      }, idx * 500);
    });
  };

  const handleDownloadRange = () => {
    if (rangeStart < 1 || rangeEnd > data.length || rangeStart > rangeEnd) {
      alert("Invalid range.");
      return;
    }
    for (let i = rangeStart - 1; i < rangeEnd; i++) {
      setTimeout(
        () => {
          try {
            const blob = generateDocx(data[i]);
            const name = data[i].name || data[i].Name || `certificate_${i + 1}`;
            saveAs(blob, `${name}.docx`);
          } catch (error) {
            console.error(`Error generating certificate ${i + 1}:`, error);
          }
        },
        (i - rangeStart + 1) * 500,
      );
    }
    setShowRangeDialog(false);
  };

  const applyFilters = (conditions: FilterCondition[]): CertificateData[] => {
    if (conditions.length === 0) return data;
    
    return data.filter(record => {
      // All conditions must match (AND logic)
      return conditions.every(condition => {
        const recordValue = record[condition.column]?.toString() || '';
        return recordValue.toLowerCase() === condition.value.toLowerCase();
      });
    });
  };

  const handleAddFilterCondition = () => {
    setTempFilterConditions([
      ...tempFilterConditions,
      { column: '', value: '' }
    ]);
  };

  const handleRemoveFilterCondition = (index: number) => {
    const newConditions = tempFilterConditions.filter((_, i) => i !== index);
    setTempFilterConditions(newConditions);
    
    // Update filtered data preview
    const filtered = applyFilters(newConditions);
    setFilteredData(filtered);
  };

  const handleUpdateFilterCondition = (index: number, field: 'column' | 'value', newValue: string) => {
    const updatedConditions = [...tempFilterConditions];
    updatedConditions[index] = {
      ...updatedConditions[index],
      [field]: newValue
    };
    setTempFilterConditions(updatedConditions);
    
    // Update filtered data preview
    const filtered = applyFilters(updatedConditions);
    setFilteredData(filtered);
  };

  const handleApplyFilters = () => {
    // Remove empty conditions
    const validConditions = tempFilterConditions.filter(
      c => c.column && c.value
    );
    
    if (validConditions.length === 0) {
      alert('Please add at least one filter condition');
      return;
    }
    
    setFilterConditions(validConditions);
    const filtered = applyFilters(validConditions);
    setFilteredData(filtered);
    setIsFiltered(true);
    setFilteredIndex(0);
    setShowFilterModal(false);
  };

  const handleClearAllFilters = () => {
    setFilterConditions([]);
    setTempFilterConditions([]);
    setIsFiltered(false);
    setFilteredData([]);
    setFilterColumn('');
    setFilterValue('');
  };

  const handleExcelDatabaseHeadersAndQuery = (column: string) => {
    setFilterColumn(column);
    // clear any previous value/search
    setFilterValue("");
    setIsFiltered(false);
    setFilteredData([]);

    const values = Array.from(
      new Set(data.map((row) => row[column]?.toString()).filter(Boolean)),
    );

    setUniqueValues(values as string[]);
  };

  const handleDownloadBySelectedCriteria = () => {
    // Use filterConditions if available, otherwise use tempFilterConditions
    const conditions = filterConditions.length > 0 
      ? filterConditions 
      : tempFilterConditions.filter(c => c.column && c.value);
    
    if (conditions.length === 0) {
      alert("Please add at least one filter condition.");
      return;
    }

    // Apply all conditions
    const filtered = applyFilters(conditions);

    if (filtered.length === 0) {
      alert("No matching records found.");
      return;
    }

    filtered.forEach((record, idx) => {
      setTimeout(() => {
        try {
          const blob = generateDocx(record);
          // Create a descriptive filename from filter values
          const filterStr = conditions.map(c => c.value).join('_');
          const name = record.name || record.Name || `${filterStr}_${idx + 1}`;
          saveAs(blob, `${name}.docx`);
        } catch (error) {
          console.error(`Error generating certificate ${idx + 1}:`, error);
        }
      }, idx * 400);
    });
  };

  const handleClearFilter = () => {
    setFilterConditions([]);
    setTempFilterConditions([]);
    setIsFiltered(false);
    setFilteredData([]);
    setFilterColumn('');
    setFilterValue('');
  };

  const handlePrint = () => {
    const printWindow = window.open("", "", "width=800,height=600");
    if (!printWindow || !certRef.current) return;
    printWindow.document.write(`
      <html><head><title>Certificate</title><style>body{margin:20px;font-family:'Times New Roman',serif;}@media print{body{margin:0;}}</style></head><body>${certRef.current.innerHTML}</body></html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  const handlePrintAll = () => {
    const printWindow = window.open("", "", "width=800,height=600");
    if (!printWindow) return;
    
    const dataToPrint = isFiltered && filteredData.length > 0 ? filteredData : data;
    
    const allCertificates = dataToPrint
      .map(
        (record, idx) =>
          `<div style="page-break-after:${idx < dataToPrint.length - 1 ? "always" : "auto"};">${mergeCertificate(docxHtml, record)}</div>`,
      )
      .join("");
    printWindow.document.write(`
      <html><head><title>All Certificates</title><style>body{margin:20px;font-family:'Times New Roman',serif;}@media print{body{margin:0;}}</style></head><body>${allCertificates}</body></html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  const isReady = uploadStatus.docx && uploadStatus.excel;

  return (
    <div className="flex min-h-screen bg-gradient-to-br from-purple-50 to-blue-100">
      {/* Sidebar */}
      <div className="w-80 bg-white shadow-xl p-6 overflow-y-auto">
        <h2 className="text-2xl font-bold text-gray-800 mb-6">
          📁 Saved Templates
        </h2>

        {savedTemplates.length > 0 ? (
          <div className="space-y-3 mb-8">
            {savedTemplates.map((template) => (
              <div
                key={template.id}
                className={`p-4 rounded-lg border-2 cursor-pointer transition ${
                  selectedTemplateId === template.id
                    ? "bg-purple-50 border-purple-500"
                    : "bg-gray-50 border-gray-300 hover:border-purple-300"
                }`}
                onClick={() => handleLoadTemplate(template)}
              >
                <div className="flex items-start justify-between">
                  <div className="flex-1">
                    <div className="font-semibold text-gray-800 flex items-center gap-2">
                      <File className="w-4 h-4" />
                      {template.name}
                    </div>
                    <div className="text-xs text-gray-500 mt-1">
                      {template.placeholders.length} fields •{" "}
                      {template.uploadDate}
                    </div>
                  </div>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleDeleteTemplate(template.id);
                    }}
                    className="text-red-500 hover:text-red-700 p-1"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
            ))}
          </div>
        ) : (
          <div className="bg-gray-50 p-4 rounded-lg mb-8 text-center text-gray-500 text-sm">
            No saved templates. Upload a DOCX to get started!
          </div>
        )}

        <div className="border-t pt-6 mb-6"></div>

        <h2 className="text-xl font-bold text-gray-800 mb-4">📊 Status</h2>

        <div className="mb-8">
          <div
            className={`p-5 rounded-lg border-2 ${isReady ? "bg-green-50 border-green-500" : "bg-gray-50 border-gray-300"}`}
          >
            <div className="flex items-center gap-2 mb-3">
              {isReady ? (
                <Check className="w-6 h-6 text-green-600" />
              ) : (
                <Upload className="w-6 h-6 text-gray-400" />
              )}
              <span className="font-semibold text-lg">Files</span>
            </div>
            <div className="space-y-2">
              <div className="flex items-center gap-2 text-sm">
                {uploadStatus.docx ? (
                  <Check className="w-4 h-4 text-green-600" />
                ) : (
                  <div className="w-4 h-4 border-2 border-gray-300 rounded" />
                )}
                <FileText className="w-4 h-4" />
                <span>Template</span>
              </div>
              <div className="flex items-center gap-2 text-sm">
                {uploadStatus.excel ? (
                  <Check className="w-4 h-4 text-green-600" />
                ) : (
                  <div className="w-4 h-4 border-2 border-gray-300 rounded" />
                )}
                <FileSpreadsheet className="w-4 h-4" />
                <span>Excel {uploadStatus.excel && `(${data.length})`}</span>
                {uploadStatus.excel && (
                  <button
                    onClick={handleDeleteExcel}
                    className="text-red-500 hover:text-red-700 p-1 ml-1"
                    title="Delete Excel data"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>

        {placeholders.length > 0 && (
          <div className="mb-8">
            <h3 className="font-bold text-gray-800 mb-3 flex items-center gap-2">
              <AlertCircle className="w-5 h-5" />
              Placeholders
            </h3>
            <div className="space-y-2">
              {placeholders.map((ph) => (
                <div key={ph} className="bg-blue-50 px-3 py-2 rounded text-sm">
                  <code className="text-blue-700">{`{${ph}}`}</code>
                </div>
              ))}
            </div>
          </div>
        )}

        {excelColumns.length > 0 && (
          <div className="mb-8">
            <h3 className="font-bold text-gray-800 mb-3 flex items-center gap-2">
              <FileSpreadsheet className="w-5 h-5" />
              Excel Columns
            </h3>
            <div className="space-y-2">
              {excelColumns.map((col) => (
                <div
                  key={col}
                  className="bg-green-50 px-3 py-2 rounded text-sm"
                >
                  <span className="text-green-700 font-medium">{col}</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Main Content */}
      <div className="flex-1 p-8 overflow-y-auto">
        <div className="max-w-6xl mx-auto">
          <div className="bg-white rounded-lg shadow-xl p-8 mb-8">
            <h1 className="text-3xl font-bold text-gray-800 mb-6">
              🎓 Certificate Generator
            </h1>

            <div className="grid grid-cols-2 gap-6">
              <label className="flex flex-col items-center justify-center h-40 px-4 transition bg-white border-2 border-dashed rounded-lg cursor-pointer hover:border-purple-500 border-gray-300">
                <div className="flex flex-col items-center space-y-2">
                  <FileText className="w-10 h-10 text-gray-400" />
                  <span className="font-medium text-gray-600">
                    Upload DOCX Template
                  </span>
                  <span className="text-xs text-gray-500">
                    {uploadStatus.docx ? "✓ Loaded" : "Will be saved"}
                  </span>
                </div>
                <input
                  type="file"
                  className="hidden"
                  accept=".docx"
                  onChange={handleDocxUpload}
                />
              </label>

              <label className="relative flex flex-col items-center justify-center h-40 px-4 transition bg-white border-2 border-dashed rounded-lg cursor-pointer hover:border-blue-500 border-gray-300">
                {uploadStatus.excel && (
                  <button
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      handleDeleteExcel();
                    }}
                    className="absolute top-2 right-2 text-red-500 hover:text-red-700 p-1"
                    title="Delete Excel data"
                  >
                    <Trash2 className="w-5 h-5" />
                  </button>
                )}
                <div className="flex flex-col items-center space-y-2">
                  <Upload className="w-10 h-10 text-gray-400" />
                  <span className="font-medium text-gray-600">
                    Upload Excel Data
                  </span>
                  <span className="text-xs text-gray-500">
                    {uploadStatus.excel
                      ? `✓ ${data.length} records saved`
                      : "Will be saved"}
                  </span>
                </div>
                <input
                  ref={excelInputRef}
                  type="file"
                  className="hidden"
                  accept=".xlsx,.xls"
                  onChange={handleExcelUpload}
                />
              </label>
            </div>
          </div>

          {isReady && (
            <div className="bg-white rounded-lg shadow-xl p-6 mb-8">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xl font-semibold text-gray-800 flex items-center gap-2">
                  <Eye className="w-5 h-5" />
                  Preview {isFiltered && <span className="text-sm text-orange-600">(Filtered)</span>}
                </h2>
                <div className="flex gap-2 items-center">
                  <button
                    onClick={() => {
                      if (isFiltered) {
                        setFilteredIndex(Math.max(0, filteredIndex - 1));
                      } else {
                        setCurrentIndex(Math.max(0, currentIndex - 1));
                      }
                    }}
                    disabled={isFiltered ? filteredIndex === 0 : currentIndex === 0}
                    className="px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    ← Previous
                  </button>
                  <span className="px-4 py-2 bg-purple-100 text-purple-800 rounded font-medium">
                    {isFiltered && filteredData.length > 0
                      ? `${filteredIndex + 1} / ${filteredData.length} (filtered)` 
                      : `${currentIndex + 1} / ${data.length}`}
                  </span>
                  <button
                    onClick={() => {
                      if (isFiltered) {
                        setFilteredIndex(Math.min(filteredData.length - 1, filteredIndex + 1));
                      } else {
                        setCurrentIndex(Math.min(data.length - 1, currentIndex + 1));
                      }
                    }}
                    disabled={isFiltered 
                      ? filteredIndex === filteredData.length - 1 || filteredData.length === 0
                      : currentIndex === data.length - 1}
                    className="px-4 py-2 bg-gray-200 text-gray-700 rounded hover:bg-gray-300 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    Next →
                  </button>
                  
                  {/* Clear filter button */}
                  {isFiltered && (
                    <button
                      onClick={handleClearFilter}
                      className="px-3 py-2 bg-gray-500 text-white rounded hover:bg-gray-600 ml-2"
                      title="Clear filter"
                    >
                      ✕ Clear Filter
                    </button>
                  )}
                </div>
              </div>

              <div className="bg-white border-4 border-gray-200 rounded-lg overflow-auto p-8 max-h-[600px]">
                <div
                  ref={certRef}
                  className="certificate-preview mx-auto"
                  dangerouslySetInnerHTML={{
                    __html: mergeCertificate(
                      docxHtml, 
                      isFiltered && filteredData.length > 0 
                        ? filteredData[filteredIndex] 
                        : data[currentIndex]
                    ),
                  }}
                />
                {isFiltered && filteredData.length === 0 && (
                  <div className="text-center text-red-500 py-8">
                    No records match the selected filter
                  </div>
                )}
              </div>

              <div className="mt-6 space-y-4">
                {/* Action Buttons */}
                <div className="flex justify-end gap-3 mb-4">
                  <button
                    onClick={() => {
                      setEditableData([...data]);
                      setShowExcelEditor(true);
                    }}
                    className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
                  >
                    <FileSpreadsheet className="w-5 h-5" />
                    Edit Excel Data
                  </button>
                  
                  <button
                    onClick={() => {
                      setTempFilterConditions(filterConditions);
                      setShowFilterModal(true);
                    }}
                    className="flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700"
                  >
                    <Filter className="w-5 h-5" />
                    Filter Records
                    {isFiltered && (
                      <span className="bg-white text-purple-600 rounded-full px-2 py-0.5 text-xs font-bold">
                        {filteredData.length}
                      </span>
                    )}
                  </button>
                </div>

                {/* Download buttons */}
                <div className="grid grid-cols-3 gap-4">
                  <button 
                    onClick={handleDownloadCurrent} 
                    className="flex items-center justify-center gap-2 px-4 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
                  >
                    <Download className="w-5 h-5" />
                    Download Current
                  </button>
                  <button 
                    onClick={() => setShowRangeDialog(true)} 
                    className="flex items-center justify-center gap-2 px-4 py-3 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700"
                  >
                    <Download className="w-5 h-5" />
                    Download Range
                  </button>
                  <button 
                    onClick={handleDownloadAll} 
                    className="flex items-center justify-center gap-2 px-4 py-3 bg-green-600 text-white rounded-lg hover:bg-green-700"
                  >
                    <Download className="w-5 h-5" />
                    Download {isFiltered ? 'Filtered' : 'All'} ({isFiltered ? filteredData.length : data.length})
                  </button>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <button 
                    onClick={handlePrint} 
                    className="flex items-center justify-center gap-2 px-4 py-3 bg-purple-600 text-white rounded-lg hover:bg-purple-700"
                  >
                    <Download className="w-5 h-5" />
                    Print Current
                  </button>
                  <button 
                    onClick={handlePrintAll} 
                    className="flex items-center justify-center gap-2 px-4 py-3 bg-purple-700 text-white rounded-lg hover:bg-purple-800"
                  >
                    <Download className="w-5 h-5" />
                    Print {isFiltered ? 'Filtered' : 'All'} ({isFiltered ? filteredData.length : data.length})
                  </button>
                </div>
              </div>
            </div>
          )}

          {!isReady && (
            <div className="bg-white rounded-lg shadow-xl p-8">
              <h2 className="text-xl font-bold text-gray-800 mb-4">
                📝 How to Use
              </h2>
              <ol className="space-y-3 text-gray-700">
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">1.</span>
                  <span>
                    Upload DOCX template with placeholders -{" "}
                    <strong>saves forever</strong>
                  </span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">2.</span>
                  <span>
                    Upload Excel file - <strong>re-upload each session</strong>
                  </span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">3.</span>
                  <span>Generate and download certificates!</span>
                </li>
              </ol>
            </div>
          )}
        </div>
      </div>

      {/* Save Dialog */}
      {showSaveDialog && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-96 shadow-2xl">
            <h3 className="text-xl font-bold mb-4">Save Template</h3>
            <input
              type="text"
              value={templateName}
              onChange={(e) => setTemplateName(e.target.value)}
              placeholder="Enter template name"
              className="w-full px-4 py-2 border rounded-lg mb-4 focus:outline-none focus:ring-2 focus:ring-purple-500"
              autoFocus
            />
            <div className="flex gap-3">
              <button
                onClick={handleSaveTemplate}
                className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700"
              >
                Save
              </button>
              <button
                onClick={() => setShowSaveDialog(false)}
                className="flex-1 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Range Dialog */}
      {showRangeDialog && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-96 shadow-2xl">
            <h3 className="text-xl font-bold mb-4">Download Range</h3>
            <div className="space-y-4 mb-4">
              <div>
                <label className="block text-sm font-medium mb-1">From:</label>
                <input
                  type="number"
                  min="1"
                  max={data.length}
                  value={rangeStart}
                  onChange={(e) => setRangeStart(parseInt(e.target.value) || 1)}
                  className="w-full px-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                />
              </div>
              <div>
                <label className="block text-sm font-medium mb-1">To:</label>
                <input
                  type="number"
                  min="1"
                  max={data.length}
                  value={rangeEnd}
                  onChange={(e) =>
                    setRangeEnd(parseInt(e.target.value) || data.length)
                  }
                  className="w-full px-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                />
              </div>
              <p className="text-sm text-gray-600">
                Selected: {Math.max(0, rangeEnd - rangeStart + 1)}
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={handleDownloadRange}
                className="flex-1 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700"
              >
                Download
              </button>
              <button
                onClick={() => setShowRangeDialog(false)}
                className="flex-1 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Filter Modal */}
      {showFilterModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-full max-w-2xl md:w-[600px] lg:w-[800px] max-h-[80vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-xl font-bold">Filter Records</h3>
              <button
                onClick={() => {
                  setTempFilterConditions(filterConditions);
                  setShowFilterModal(false);
                }}
                className="text-gray-500 hover:text-gray-700"
              >
                ✕
              </button>
            </div>

            <div className="space-y-4">
              {/* Filter Conditions */}
              <div className="space-y-3">
                {tempFilterConditions.length === 0 ? (
                  <div className="text-center text-gray-500 py-4 bg-gray-50 rounded-lg">
                    No filters applied. Click "Add Filter" to get started.
                  </div>
                ) : (
                  tempFilterConditions.map((condition, index) => (
                    <div key={index} className="bg-gray-50 p-3 rounded-lg relative">
                      <button
                        onClick={() => handleRemoveFilterCondition(index)}
                        className="absolute top-2 right-2 text-red-500 hover:text-red-700"
                        title="Remove filter"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                      
                      <div className="grid grid-cols-2 gap-2 pr-8">
                        <div>
                          <label className="block text-xs font-medium text-gray-700 mb-1">
                            Column
                          </label>
                          <select
                            className="w-full border rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500"
                            value={condition.column}
                            onChange={(e) => handleUpdateFilterCondition(index, 'column', e.target.value)}
                          >
                            <option value="">Select column</option>
                            {excelColumns.map((col) => (
                              <option key={col} value={col}>
                                {col}
                              </option>
                            ))}
                          </select>
                        </div>
                        
                        <div>
                          <label className="block text-xs font-medium text-gray-700 mb-1">
                            Value
                          </label>
                          {condition.column ? (
                            <div className="relative">
                              <input
                                type="text"
                                className="w-full border rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500"
                                value={condition.value}
                                onChange={(e) => handleUpdateFilterCondition(index, 'value', e.target.value)}
                                placeholder="Type to search..."
                              />
                              
                              {/* Value suggestions */}
                              {condition.value && (
                                <ul className="absolute z-10 w-full mt-1 bg-white border rounded-lg shadow-lg max-h-40 overflow-auto">
                                  {Array.from(new Set(
                                    data
                                      .map(row => row[condition.column]?.toString() || '')
                                      .filter(v => v.toLowerCase().includes(condition.value.toLowerCase()))
                                  )).slice(0, 10)
                                    .map((v) => (
                                      <li
                                        key={v}
                                        className="px-3 py-1 hover:bg-purple-50 cursor-pointer text-sm"
                                        onClick={() => handleUpdateFilterCondition(index, 'value', v)}
                                      >
                                        {v}
                                      </li>
                                    ))}
                                </ul>
                              )}
                            </div>
                          ) : (
                            <input
                              type="text"
                              className="w-full border rounded-lg px-3 py-2 text-sm bg-gray-100"
                              value="Select a column first"
                              disabled
                            />
                          )}
                        </div>
                      </div>
                    </div>
                  ))
                )}
              </div>

              {/* Add Filter Button */}
              <button
                onClick={handleAddFilterCondition}
                className="w-full px-4 py-2 border-2 border-dashed border-purple-300 text-purple-600 rounded-lg hover:bg-purple-50 transition-colors"
              >
                + Add Another Filter
              </button>

              {/* Preview of matched records */}
              {tempFilterConditions.some(c => c.column && c.value) && (
                <div className="mt-4 p-4 bg-purple-50 rounded-lg">
                  <div className="flex justify-between items-center">
                    <div>
                      <p className="text-sm text-purple-700">
                        Found <span className="font-bold">{filteredData.length}</span> matching record(s)
                      </p>
                      {filteredData.length > 0 && (
                        <p className="text-xs text-purple-600 mt-1">
                          Showing all columns for these records
                        </p>
                      )}
                    </div>
                    {filteredData.length > 0 && (
                      <button
                        onClick={handleApplyFilters}
                        className="px-3 py-1 bg-purple-600 text-white rounded hover:bg-purple-700 text-sm"
                      >
                        Apply Filters
                      </button>
                    )}
                  </div>
                  
                  {/* Preview of first few filtered records */}
                  {filteredData.length > 0 && (
                    <div className="mt-3 max-h-40 overflow-auto bg-white rounded border">
                      <table className="min-w-full text-xs">
                        <thead className="bg-gray-50 sticky top-0">
                          <tr>
                            {excelColumns.slice(0, 4).map(col => (
                              <th key={col} className="px-2 py-1 text-left font-medium text-gray-600">
                                {col}
                              </th>
                            ))}
                            {excelColumns.length > 4 && (
                              <th className="px-2 py-1 text-left font-medium text-gray-600">
                                ...
                              </th>
                            )}
                          </tr>
                        </thead>
                        <tbody>
                          {filteredData.slice(0, 3).map((record, idx) => (
                            <tr key={idx} className="border-t">
                              {excelColumns.slice(0, 4).map(col => (
                                <td key={col} className="px-2 py-1 truncate max-w-[100px]">
                                  {record[col]?.toString() || '-'}
                                </td>
                              ))}
                              {excelColumns.length > 4 && (
                                <td className="px-2 py-1 text-gray-400">
                                  +{excelColumns.length - 4} more
                                </td>
                              )}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      {filteredData.length > 3 && (
                        <div className="text-center text-gray-500 text-xs py-1 border-t">
                          ... and {filteredData.length - 3} more
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}

              {/* Active Filters Indicator */}
              {filterConditions.length > 0 && (
                <div className="mt-4 p-3 bg-blue-50 rounded-lg">
                  <div className="flex justify-between items-center mb-2">
                    <span className="text-sm font-medium text-blue-700">Active Filters:</span>
                    <button
                      onClick={handleClearAllFilters}
                      className="text-blue-700 hover:text-blue-900 text-xs font-medium"
                    >
                      Clear All
                    </button>
                  </div>
                  <div className="space-y-1">
                    {filterConditions.map((condition, idx) => (
                      <div key={idx} className="text-sm text-blue-600 flex items-center gap-2">
                        <span className="bg-blue-100 px-2 py-0.5 rounded">
                          {condition.column} = {condition.value}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Modal Actions */}
            <div className="flex gap-3 mt-6">
              <button
                onClick={() => {
                  const validConditions = tempFilterConditions.filter(c => c.column && c.value);
                  if (validConditions.length === 0) {
                    alert('Please add at least one filter condition');
                    return;
                  }
                  handleDownloadBySelectedCriteria();
                  setShowFilterModal(false);
                }}
                disabled={!tempFilterConditions.some(c => c.column && c.value)}
                className={`flex-1 px-4 py-2 rounded-lg ${
                  !tempFilterConditions.some(c => c.column && c.value)
                    ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                    : "bg-orange-600 text-white hover:bg-orange-700"
                }`}
              >
                Download Filtered ({filteredData.length})
              </button>
              <button
                onClick={() => {
                  setTempFilterConditions(filterConditions);
                  setShowFilterModal(false);
                }}
                className="flex-1 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Excel Editor Modal */}
      {showExcelEditor && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-full max-w-6xl max-h-[90vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-xl font-bold">Edit Excel Data</h3>
              <button
                onClick={() => setShowExcelEditor(false)}
                className="text-gray-500 hover:text-gray-700"
              >
                ✕
              </button>
            </div>

            {/* Data Table */}
            <div className="overflow-x-auto mb-4">
              <table className="min-w-full border-collapse border border-gray-300">
                <thead className="bg-gray-100 sticky top-0">
                  <tr>
                    {excelColumns.map((col) => (
                      <th key={col} className="border border-gray-300 px-4 py-2 text-left text-sm font-semibold">
                        {col}
                        <button
                          onClick={() => {
                            // Add new column
                            const newCol = prompt("Enter new column name:");
                            if (newCol && !excelColumns.includes(newCol)) {
                              const updatedColumns = [...excelColumns, newCol];
                              const updatedData = editableData.map(row => ({
                                ...row,
                                [newCol]: ""
                              }));
                              setExcelColumns(updatedColumns);
                              setEditableData(updatedData);
                            }
                          }}
                          className="ml-2 text-green-600 hover:text-green-800"
                          title="Add column"
                        >
                          +
                        </button>
                      </th>
                    ))}
                    <th className="border border-gray-300 px-4 py-2 text-center">
                      Actions
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {/* Existing rows */}
                  {editableData.map((row, rowIndex) => (
                    <tr key={rowIndex} className="hover:bg-gray-50">
                      {excelColumns.map((col) => (
                        <td key={col} className="border border-gray-300 px-4 py-2">
                          <input
                            type="text"
                            value={row[col]?.toString() || ''}
                            onChange={(e) => {
                              const updatedData = [...editableData];
                              updatedData[rowIndex] = {
                                ...updatedData[rowIndex],
                                [col]: e.target.value
                              };
                              setEditableData(updatedData);
                            }}
                            className="w-full px-2 py-1 border rounded focus:outline-none focus:ring-2 focus:ring-purple-500"
                          />
                        </td>
                      ))}
                      <td className="border border-gray-300 px-4 py-2 text-center">
                        <button
                          onClick={() => {
                            const updatedData = editableData.filter((_, i) => i !== rowIndex);
                            setEditableData(updatedData);
                          }}
                          className="text-red-600 hover:text-red-800 mx-1"
                          title="Delete row"
                        >
                          <Trash2 className="w-4 h-4 inline" />
                        </button>
                      </td>
                    </tr>
                  ))}
                  
                  {/* New row input */}
                  <tr className="bg-blue-50">
                    {excelColumns.map((col) => (
                      <td key={col} className="border border-gray-300 px-4 py-2">
                        <input
                          type="text"
                          placeholder={`Enter ${col}`}
                          value={newRowData[col]?.toString() || ''}
                          onChange={(e) => {
                            setNewRowData({
                              ...newRowData,
                              [col]: e.target.value
                            });
                          }}
                          className="w-full px-2 py-1 border rounded focus:outline-none focus:ring-2 focus:ring-green-500"
                        />
                      </td>
                    ))}
                    <td className="border border-gray-300 px-4 py-2 text-center">
                      <button
                        onClick={() => {
                          // Check if any field is filled
                          if (Object.keys(newRowData).length > 0) {
                            setEditableData([...editableData, newRowData]);
                            setNewRowData({});
                          }
                        }}
                        className="text-green-600 hover:text-green-800"
                        title="Add row"
                      >
                        +
                      </button>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Bulk Add Options */}
            <div className="grid grid-cols-2 gap-4 mb-4">
              <div>
                <h4 className="font-semibold mb-2">Bulk Add Records</h4>
                <textarea
                  placeholder="Paste CSV data here (one row per line, comma-separated)"
                  className="w-full h-24 p-2 border rounded focus:outline-none focus:ring-2 focus:ring-purple-500"
                  onChange={(e) => {
                    const csvText = e.target.value;
                    if (csvText.trim()) {
                      const rows = csvText.split('\n');
                      const newRows = rows.map(row => {
                        const values = row.split(',').map(v => v.trim());
                        const newRow: CertificateData = {};
                        excelColumns.forEach((col, index) => {
                          if (values[index]) {
                            newRow[col] = values[index];
                          }
                        });
                        return newRow;
                      }).filter(row => Object.keys(row).length > 0);
                      
                      setEditableData([...editableData, ...newRows]);
                    }
                  }}
                />
                <p className="text-xs text-gray-500 mt-1">
                  Example: John,Doe,john@email.com,2024-01-01
                </p>
              </div>
              
              <div>
                <h4 className="font-semibold mb-2">Import from File</h4>
                <label className="flex items-center justify-center h-24 px-4 border-2 border-dashed border-purple-300 rounded-lg cursor-pointer hover:bg-purple-50">
                  <div className="text-center">
                    <Upload className="w-6 h-6 text-purple-600 mx-auto mb-1" />
                    <span className="text-sm text-purple-600">Click to upload CSV/Excel</span>
                  </div>
                  <input
                    type="file"
                    className="hidden"
                    accept=".csv,.xlsx,.xls"
                    onChange={(e) => {
                      const file = e.target.files?.[0];
                      if (file) {
                        const reader = new FileReader();
                        reader.onload = (evt) => {
                          try {
                            const bstr = evt.target?.result;
                            const wb = XLSX.read(bstr, { type: "binary" });
                            const wsname = wb.SheetNames[0];
                            const ws = wb.Sheets[wsname];
                            const jsonData = XLSX.utils.sheet_to_json(ws) as CertificateData[];
                            
                            if (jsonData.length > 0) {
                              const processedData = processExcelData(jsonData);
                              setEditableData([...editableData, ...processedData]);
                            }
                          } catch (error) {
                            console.error("Error importing file:", error);
                            alert("Error importing file. Please check the format.");
                          }
                        };
                        reader.readAsBinaryString(file);
                      }
                    }}
                  />
                </label>
              </div>
            </div>

            {/* Modal Actions */}
            <div className="flex gap-3 justify-end mt-6">
              <button
                onClick={() => {
                  // Save changes
                  setData(editableData);
                  setRangeEnd(editableData.length);
                  saveExcelData(editableData, excelColumns);
                  setShowExcelEditor(false);
                  alert("Excel data updated successfully!");
                }}
                className="px-6 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700"
              >
                Save Changes
              </button>
              <button
                onClick={() => {
                  if (confirm("Discard all changes?")) {
                    setEditableData(data);
                    setNewRowData({});
                    setShowExcelEditor(false);
                  }
                }}
                className="px-6 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300"
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