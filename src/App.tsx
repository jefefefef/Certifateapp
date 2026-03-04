import React, { useState, useRef, useEffect } from "react";
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
  Save,
  ChevronLeft,
  ChevronDown,
  FolderOpen,
  Search,
  X,
  Plus,
  Activity,
  Tag,
  Grid,
  Edit2,
  MoreVertical,
  Copy,
  ArrowUp,
  ArrowDown,
  RotateCcw,
  RotateCw
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

interface DownloadCounter {
  templateId: string;
  templateName: string;
  monthKey: string;
  count: number;
}

interface CellPosition {
  row: number;
  col: number;
}

interface EditHistory {
  past: ExcelFileSnapshot[];
  future: ExcelFileSnapshot[];
}

interface ExcelFileSnapshot {
  data: CertificateData[];
  columns: string[];
}

// Simple IndexedDB operations
const DB_NAME = "CertGenDB";
const DB_VERSION = 2;
const STORE_NAME = "templates";
const EXCEL_STORE = "excelData";

// Counter functions using LocalStorage with CSV
const COUNTERS_KEY = 'certificate_counters';

const loadCounters = (): DownloadCounter[] => {
  try {
    const csv = localStorage.getItem(COUNTERS_KEY);
    if (!csv) return [];
    
    const lines = csv.split('\n').slice(1);
    return lines
      .filter(line => line.trim())
      .map(line => {
        const [templateId, templateName, monthKey, count] = line.split(',');
        return {
          templateId,
          templateName: templateName || '',
          monthKey,
          count: parseInt(count, 10) || 0
        };
      });
  } catch (error) {
    console.error('Error loading counters:', error);
    return [];
  }
};

const saveCounters = (counters: DownloadCounter[]): void => {
  try {
    const header = 'templateId,templateName,monthKey,count\n';
    const csv = header + counters
      .map(c => `${c.templateId},${c.templateName},${c.monthKey},${c.count}`)
      .join('\n');
    
    localStorage.setItem(COUNTERS_KEY, csv);
    console.log('✅ Counters saved to LocalStorage');
  } catch (error) {
    console.error('Error saving counters:', error);
  }
};

const getNextCertificateNumber = (
    templateId: string,
    templateName: string,
    prefix: string = ''
  ): string => {
    const counters = loadCounters();
    const now = new Date();
    const monthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    
    const uniqueKey = prefix ? `${templateId}_${prefix}` : templateId;
    
    const existingCounter = counters.find(
      c => c.templateId === uniqueKey && c.monthKey === monthKey
    );
    
    let nextCount = 1;
    
    if (existingCounter) {
      nextCount = existingCounter.count + 1;
      existingCounter.count = nextCount;
    } else {
      counters.push({
        templateId: uniqueKey,
        templateName: prefix ? `${templateName}_${prefix}` : templateName,
        monthKey,
        count: 1
      });
    }
    
    const threeMonthsAgo = new Date();
    threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
    const threeMonthsAgoKey = `${threeMonthsAgo.getFullYear()}-${String(threeMonthsAgo.getMonth() + 1).padStart(2, '0')}`;
    
    const filteredCounters = counters.filter(c => 
      c.templateId === uniqueKey ? c.monthKey >= threeMonthsAgoKey : true
    );
    
    saveCounters(filteredCounters);
    
    return `${monthKey}-${String(nextCount).padStart(2, '0')}`;
  };

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

      const estimatedSize =
        template.binary.length + template.html.length + template.name.length;
      console.log(
        `📊 Estimated template size: ${(estimatedSize / 1024 / 1024).toFixed(2)} MB`,
      );

      if (estimatedSize > 5 * 1024 * 1024) {
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

// Custom hook for click outside
const useClickOutside = (ref: React.RefObject<HTMLElement>, handler: () => void) => {
  React.useEffect(() => {
    const listener = (event: MouseEvent | TouchEvent) => {
      if (!ref.current || ref.current.contains(event.target as Node)) {
        return;
      }
      handler();
    };

    document.addEventListener('mousedown', listener);
    document.addEventListener('touchstart', listener);

    return () => {
      document.removeEventListener('mousedown', listener);
      document.removeEventListener('touchstart', listener);
    };
  }, [ref, handler]);
};

// Column Header Component
const ColumnHeader: React.FC<{
  column: string;
  index: number;
  onEdit: () => void;
  onDelete: () => void;
  onInsertLeft: () => void;
  onInsertRight: () => void;
  menuOpen: boolean;
  onMenuToggle: (open: boolean) => void;
}> = ({ column, index, onEdit, onDelete, onInsertLeft, onInsertRight, menuOpen, onMenuToggle }) => {
  const menuRef = useRef<HTMLDivElement>(null);

  useClickOutside(menuRef, () => onMenuToggle(false));

  return (
    <th className="border border-gray-300 bg-gray-100 relative group">
      <div className="flex items-center justify-between px-3 py-2">
        <span className="font-medium text-sm">{column}</span>
        <button
          onClick={() => onMenuToggle(true)}
          className="p-1 hover:bg-gray-200 rounded transition"
          title="Column options"
        >
          <ChevronDown className="w-4 h-4" />
        </button>
      </div>

      {menuOpen && (
        <div ref={menuRef} className="absolute top-full left-0 mt-1 bg-white shadow-xl rounded-lg border py-1 z-50 min-w-[180px]">
          <button
            onClick={() => { onEdit(); onMenuToggle(false); }}
            className="w-full px-4 py-2 text-left hover:bg-gray-50 flex items-center gap-2 text-sm"
          >
            <Edit2 className="w-4 h-4" /> Edit Column
          </button>
          <button
            onClick={() => { onInsertLeft(); onMenuToggle(false); }}
            className="w-full px-4 py-2 text-left hover:bg-gray-50 flex items-center gap-2 text-sm"
          >
            <Plus className="w-4 h-4" /> Insert Left
          </button>
          <button
            onClick={() => { onInsertRight(); onMenuToggle(false); }}
            className="w-full px-4 py-2 text-left hover:bg-gray-50 flex items-center gap-2 text-sm"
          >
            <Plus className="w-4 h-4" /> Insert Right
          </button>
          <div className="border-t my-1" />
          <button
            onClick={() => { onDelete(); onMenuToggle(false); }}
            className="w-full px-4 py-2 text-left hover:bg-gray-50 text-red-600 flex items-center gap-2 text-sm"
          >
            <Trash2 className="w-4 h-4" /> Delete Column
          </button>
        </div>
      )}
    </th>
  );
};

// Row Edit Modal
const RowEditModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  row: CertificateData;
  columns: string[];
  onSave: (updatedRow: CertificateData) => void;
}> = ({ isOpen, onClose, row, columns, onSave }) => {
  const [editedRow, setEditedRow] = useState(row);
  const inputRefs = useRef<(HTMLInputElement | null)[]>([]);

  useEffect(() => {
    setEditedRow(row);
    if (isOpen) {
      setTimeout(() => inputRefs.current[0]?.focus(), 100);
    }
  }, [isOpen, row]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[60]">
      <div className="bg-white rounded-lg p-6 w-[500px] shadow-2xl">
        <div className="flex justify-between items-center mb-4">
          <h3 className="text-xl font-bold">Edit Row</h3>
          <button onClick={onClose} className="text-gray-500 hover:text-gray-700">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="space-y-4">
          {columns.map((col, index) => (
            <div key={col}>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                {col}
              </label>
              <input
                ref={el => inputRefs.current[index] = el}
                type="text"
                value={editedRow[col]?.toString() || ''}
                onChange={(e) => setEditedRow({ ...editedRow, [col]: e.target.value })}
                className="w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && index === columns.length - 1) {
                    onSave(editedRow);
                    onClose();
                  } else if (e.key === 'Enter') {
                    inputRefs.current[index + 1]?.focus();
                  } else if (e.key === 'Escape') {
                    onClose();
                  }
                }}
              />
            </div>
          ))}
        </div>

        <div className="flex gap-3 mt-6">
          <button
            onClick={() => {
              onSave(editedRow);
              onClose();
            }}
            className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700"
          >
            Save Changes
          </button>
          <button
            onClick={onClose}
            className="flex-1 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
};

// Column Edit Modal
const ColumnEditModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  columnName: string;
  onSave: (newName: string) => void;
}> = ({ isOpen, onClose, columnName, onSave }) => {
  const [newName, setNewName] = useState(columnName);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      setNewName(columnName);
      setTimeout(() => inputRef.current?.focus(), 100);
    }
  }, [isOpen, columnName]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[60]">
      <div className="bg-white rounded-lg p-6 w-[400px] shadow-2xl">
        <h3 className="text-xl font-bold mb-4">Edit Column</h3>
        
        <div className="mb-4">
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Column Name
          </label>
          <input
            ref={inputRef}
            type="text"
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            className="w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500"
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                onSave(newName);
                onClose();
              } else if (e.key === 'Escape') {
                onClose();
              }
            }}
          />
        </div>

        <div className="flex gap-3">
          <button
            onClick={() => {
              onSave(newName);
              onClose();
            }}
            className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700"
          >
            Save
          </button>
          <button
            onClick={onClose}
            className="flex-1 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300"
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
};

// Table Row Component
const TableRow: React.FC<{
  row: CertificateData;
  rowIndex: number;
  columns: string[];
  onEditRow: () => void;
  onDeleteRow: () => void;
  onDuplicateRow: () => void;
  onInsertAbove: () => void;
  onInsertBelow: () => void;
  menuOpen: boolean;
  onMenuToggle: (open: boolean) => void;
}> = ({
  row,
  rowIndex,
  columns,
  onEditRow,
  onDeleteRow,
  onDuplicateRow,
  onInsertAbove,
  onInsertBelow,
  menuOpen,
  onMenuToggle
}) => {
  const menuRef = useRef<HTMLDivElement>(null);

  useClickOutside(menuRef, () => onMenuToggle(false));

  return (
    <tr className="group hover:bg-gray-50">
      <td className="border border-gray-300 px-2 py-2 text-center text-sm text-gray-500 bg-gray-50 w-12">
        {rowIndex + 1}
      </td>
      
      {columns.map((col) => (
        <td key={col} className="border border-gray-300 px-3 py-2 text-sm">
          {row[col]?.toString() || <span className="text-gray-400">—</span>}
        </td>
      ))}

      <td className="border border-gray-300 px-2 py-1 text-center relative">
        <div className="flex items-center justify-center gap-1">
          <button
            onClick={onEditRow}
            className="p-1.5 text-blue-600 hover:bg-blue-50 rounded transition"
            title="Edit row"
          >
            <Edit2 className="w-4 h-4" />
          </button>
          <button
            onClick={() => onMenuToggle(!menuOpen)}
            className="p-1.5 text-gray-600 hover:bg-gray-200 rounded transition"
            title="Row options"
          >
            <MoreVertical className="w-4 h-4" />
          </button>
        </div>

        {menuOpen && (
          <div ref={menuRef} className="absolute right-0 mt-1 bg-white shadow-xl rounded-lg border py-1 z-50 min-w-[160px]">
            <button
              onClick={() => { onInsertAbove(); onMenuToggle(false); }}
              className="w-full px-4 py-2 text-left hover:bg-gray-50 text-sm flex items-center gap-2"
            >
              <ArrowUp className="w-4 h-4" /> Insert Above
            </button>
            <button
              onClick={() => { onInsertBelow(); onMenuToggle(false); }}
              className="w-full px-4 py-2 text-left hover:bg-gray-50 text-sm flex items-center gap-2"
            >
              <ArrowDown className="w-4 h-4" /> Insert Below
            </button>
            <button
              onClick={() => { onDuplicateRow(); onMenuToggle(false); }}
              className="w-full px-4 py-2 text-left hover:bg-gray-50 text-sm flex items-center gap-2"
            >
              <Copy className="w-4 h-4" /> Duplicate Row
            </button>
            <div className="border-t my-1" />
            <button
              onClick={() => { onDeleteRow(); onMenuToggle(false); }}
              className="w-full px-4 py-2 text-left hover:bg-gray-50 text-red-600 text-sm flex items-center gap-2"
            >
              <Trash2 className="w-4 h-4" /> Delete Row
            </button>
          </div>
        )}
      </td>
    </tr>
  );
};

// Empty Row Component
const EmptyRow: React.FC<{
  columns: string[];
  onAdd: (newRow: CertificateData) => void;
}> = ({ columns, onAdd }) => {
  const [newRow, setNewRow] = useState<CertificateData>({});
  const inputRefs = useRef<(HTMLInputElement | null)[]>([]);

  const handleAdd = () => {
    if (Object.keys(newRow).length > 0) {
      onAdd(newRow);
      setNewRow({});
      setTimeout(() => inputRefs.current[0]?.focus(), 0);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent, index: number) => {
    if (e.key === 'Enter') {
      if (index === columns.length - 1) {
        handleAdd();
      } else {
        inputRefs.current[index + 1]?.focus();
      }
    } else if (e.key === 'Escape') {
      setNewRow({});
      inputRefs.current[0]?.focus();
    }
  };

  return (
    <tr className="bg-blue-50 group">
      <td className="border border-gray-300 px-2 py-1 text-center text-sm text-gray-500 bg-gray-50">
        <Plus className="w-4 h-4 inline text-green-600" />
      </td>
      {columns.map((col, index) => (
        <td key={col} className="border border-gray-300 p-1">
          <input
            ref={el => inputRefs.current[index] = el}
            type="text"
            placeholder={`Enter ${col}`}
            value={newRow[col]?.toString() || ''}
            onChange={(e) => setNewRow({ ...newRow, [col]: e.target.value })}
            onKeyDown={(e) => handleKeyDown(e, index)}
            className="w-full px-2 py-1 border rounded focus:outline-none focus:ring-2 focus:ring-green-500 focus:border-transparent"
          />
        </td>
      ))}
      <td className="border border-gray-300 px-2 py-1 text-center">
        <button
          onClick={handleAdd}
          disabled={Object.keys(newRow).length === 0}
          className={`p-1.5 rounded-full transition ${
            Object.keys(newRow).length > 0
              ? 'text-green-600 hover:bg-green-100'
              : 'text-gray-400 cursor-not-allowed'
          }`}
          title="Add row"
        >
          <Check className="w-5 h-5" />
        </button>
      </td>
    </tr>
  );
};

// Main Excel Editor Modal
const ExcelEditorModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  data: CertificateData[];
  columns: string[];
  fileName: string;
  onSave: (data: CertificateData[], columns: string[]) => void;
}> = ({ isOpen, onClose, data, columns, fileName, onSave }) => {
  const [localData, setLocalData] = useState(data);
  const [localColumns, setLocalColumns] = useState(columns);
  const [history, setHistory] = useState<EditHistory>({ past: [], future: [] });
  
  // Modal states
  const [editingRow, setEditingRow] = useState<{ index: number; data: CertificateData } | null>(null);
  const [editingColumn, setEditingColumn] = useState<{ index: number; name: string } | null>(null);
  const [columnMenu, setColumnMenu] = useState<{ col: number; open: boolean } | null>(null);
  const [rowMenu, setRowMenu] = useState<{ row: number; open: boolean } | null>(null);

  // Reset state when modal opens
  useEffect(() => {
    if (isOpen) {
      setLocalData(data);
      setLocalColumns(columns);
      setHistory({ past: [], future: [] });
    }
  }, [isOpen, data, columns]);

  const pushToHistory = (newData: CertificateData[], newColumns: string[]) => {
    setHistory(prev => ({
      past: [...prev.past, { data: localData, columns: localColumns }].slice(-2),
      future: []
    }));
  };

  const undo = () => {
    if (history.past.length === 0) return;
    const previous = history.past[history.past.length - 1];
    setHistory(prev => ({
      past: prev.past.slice(0, -1),
      future: [{ data: localData, columns: localColumns }, ...prev.future],
    }));
    setLocalData(previous.data);
    setLocalColumns(previous.columns);
  };

  const redo = () => {
    if (history.future.length === 0) return;
    const next = history.future[0];
    setHistory(prev => ({
      past: [...prev.past, { data: localData, columns: localColumns }],
      future: prev.future.slice(1),
    }));
    setLocalData(next.data);
    setLocalColumns(next.columns);
  };

  const handleAddRow = () => {
    const newRow: CertificateData = {};
    localColumns.forEach(col => newRow[col] = '');
    const newData = [...localData, newRow];
    pushToHistory(newData, localColumns);
    setLocalData(newData);
  };

  const handleAddRowWithData = (newRow: CertificateData) => {
    const newData = [...localData, newRow];
    pushToHistory(newData, localColumns);
    setLocalData(newData);
  };

  const handleDeleteRow = (rowIndex: number) => {
    if (confirm('Delete this row?')) {
      const newData = localData.filter((_, i) => i !== rowIndex);
      pushToHistory(newData, localColumns);
      setLocalData(newData);
    }
  };

  const handleDuplicateRow = (rowIndex: number) => {
    const rowToDuplicate = { ...localData[rowIndex] };
    const newData = [
      ...localData.slice(0, rowIndex + 1),
      rowToDuplicate,
      ...localData.slice(rowIndex + 1)
    ];
    pushToHistory(newData, localColumns);
    setLocalData(newData);
  };

  const handleInsertRow = (rowIndex: number, position: 'above' | 'below') => {
    const newRow: CertificateData = {};
    localColumns.forEach(col => newRow[col] = '');
    const insertIndex = position === 'above' ? rowIndex : rowIndex + 1;
    const newData = [
      ...localData.slice(0, insertIndex),
      newRow,
      ...localData.slice(insertIndex)
    ];
    pushToHistory(newData, localColumns);
    setLocalData(newData);
  };

  const handleAddColumn = () => {
    const newColName = prompt('Enter new column name:');
    if (!newColName || localColumns.includes(newColName)) return;
    
    const newColumns = [...localColumns, newColName];
    const newData = localData.map(row => ({
      ...row,
      [newColName]: ''
    }));
    
    pushToHistory(newData, newColumns);
    setLocalColumns(newColumns);
    setLocalData(newData);
  };

  const handleDeleteColumn = (colIndex: number) => {
    if (!confirm(`Delete column "${localColumns[colIndex]}"?`)) return;
    
    const colToDelete = localColumns[colIndex];
    const newColumns = localColumns.filter((_, i) => i !== colIndex);
    const newData = localData.map(row => {
      const newRow = { ...row };
      delete newRow[colToDelete];
      return newRow;
    });
    
    pushToHistory(newData, newColumns);
    setLocalColumns(newColumns);
    setLocalData(newData);
  };

  const handleInsertColumn = (colIndex: number, position: 'left' | 'right') => {
    const newColName = prompt('Enter new column name:');
    if (!newColName || localColumns.includes(newColName)) return;
    
    const insertIndex = position === 'left' ? colIndex : colIndex + 1;
    const newColumns = [
      ...localColumns.slice(0, insertIndex),
      newColName,
      ...localColumns.slice(insertIndex)
    ];
    
    const newData = localData.map(row => ({
      ...row,
      [newColName]: ''
    }));
    
    pushToHistory(newData, newColumns);
    setLocalColumns(newColumns);
    setLocalData(newData);
  };

  const handleEditRow = (rowIndex: number) => {
    setEditingRow({ index: rowIndex, data: { ...localData[rowIndex] } });
  };

  const handleSaveRow = (updatedRow: CertificateData) => {
    const newData = [...localData];
    newData[editingRow!.index] = updatedRow;
    pushToHistory(newData, localColumns);
    setLocalData(newData);
    setEditingRow(null);
  };

  const handleEditColumn = (colIndex: number) => {
    setEditingColumn({ index: colIndex, name: localColumns[colIndex] });
  };

  const handleSaveColumn = (newName: string) => {
    if (!newName || newName === editingColumn!.name) {
      setEditingColumn(null);
      return;
    }

    const oldName = localColumns[editingColumn!.index];
    const newColumns = [...localColumns];
    newColumns[editingColumn!.index] = newName;
    
    const newData = localData.map(row => {
      const newRow = { ...row };
      newRow[newName] = newRow[oldName];
      delete newRow[oldName];
      return newRow;
    });
    
    pushToHistory(newData, newColumns);
    setLocalColumns(newColumns);
    setLocalData(newData);
    setEditingColumn(null);
  };

  const handleSave = () => {
    onSave(localData, localColumns);
    onClose();
  };

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!isOpen) return;
      if (e.ctrlKey || e.metaKey) {
        if (e.key === 'z' && !e.shiftKey) {
          e.preventDefault();
          undo();
        }
        if (e.key === 'y' || (e.key === 'z' && e.shiftKey)) {
          e.preventDefault();
          redo();
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, history]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg w-full max-w-6xl max-h-[90vh] flex flex-col shadow-2xl">
        {/* Header */}
        <div className="flex justify-between items-center p-4 border-b">
          <div className="flex items-center gap-3">
            <FileSpreadsheet className="w-6 h-6 text-blue-600" />
            <div>
              <h3 className="text-xl font-bold">Edit Excel Data</h3>
              <p className="text-sm text-gray-500">{fileName} • {localData.length} rows • {localColumns.length} columns</p>
            </div>
          </div>
          <button onClick={onClose} className="text-gray-500 hover:text-gray-700">
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* Toolbar */}
        <div className="flex items-center gap-2 p-2 border-b bg-gray-50">
          <button
            onClick={undo}
            disabled={history.past.length === 0}
            className="p-2 rounded hover:bg-gray-200 disabled:opacity-30 disabled:hover:bg-transparent"
            title="Undo (Ctrl+Z)"
          >
            <RotateCcw className="w-5 h-5" />
          </button>
          <button
            onClick={redo}
            disabled={history.future.length === 0}
            className="p-2 rounded hover:bg-gray-200 disabled:opacity-30 disabled:hover:bg-transparent"
            title="Redo (Ctrl+Y)"
          >
            <RotateCw className="w-5 h-5" />
          </button>
          <div className="w-px h-6 bg-gray-300 mx-1" />
          <button
            onClick={handleAddRow}
            className="flex items-center gap-1 px-3 py-1.5 bg-green-600 text-white rounded hover:bg-green-700 text-sm"
          >
            <Plus className="w-4 h-4" /> Add Row
          </button>
          <button
            onClick={handleAddColumn}
            className="flex items-center gap-1 px-3 py-1.5 bg-purple-600 text-white rounded hover:bg-purple-700 text-sm"
          >
            <Plus className="w-4 h-4" /> Add Column
          </button>
          <div className="flex-1" />
          <div className="text-xs text-gray-500 flex items-center gap-3">
            <span className="flex items-center gap-1">
              <Edit2 className="w-4 h-4" /> Click edit buttons
            </span>
            <span className="flex items-center gap-1">
              <kbd className="px-1.5 py-0.5 bg-gray-200 rounded text-xs">Ctrl+Z</kbd> undo
            </span>
            <span className="flex items-center gap-1">
              <kbd className="px-1.5 py-0.5 bg-gray-200 rounded text-xs">Ctrl+Y</kbd> redo
            </span>
          </div>
          <button
            onClick={handleSave}
            className="flex items-center gap-1 px-4 py-1.5 bg-blue-600 text-white rounded hover:bg-blue-700"
          >
            <Save className="w-4 h-4" /> Save
          </button>
        </div>

        {/* Excel Grid */}
        <div className="flex-1 overflow-auto p-4">
          <table className="border-collapse border border-gray-300 w-full">
            <thead>
              <tr>
                <th className="border border-gray-300 bg-gray-100 w-12 text-center">#</th>
                {localColumns.map((col, index) => (
                  <ColumnHeader
                    key={col}
                    column={col}
                    index={index}
                    onEdit={() => handleEditColumn(index)}
                    onDelete={() => handleDeleteColumn(index)}
                    onInsertLeft={() => handleInsertColumn(index, 'left')}
                    onInsertRight={() => handleInsertColumn(index, 'right')}
                    menuOpen={columnMenu?.col === index && columnMenu.open}
                    onMenuToggle={(open) => setColumnMenu({ col: index, open })}
                  />
                ))}
                <th className="border border-gray-300 bg-gray-100 w-20 text-center">Actions</th>
              </tr>
            </thead>
            <tbody>
              {localData.map((row, rowIndex) => (
                <TableRow
                  key={rowIndex}
                  row={row}
                  rowIndex={rowIndex}
                  columns={localColumns}
                  onEditRow={() => handleEditRow(rowIndex)}
                  onDeleteRow={() => handleDeleteRow(rowIndex)}
                  onDuplicateRow={() => handleDuplicateRow(rowIndex)}
                  onInsertAbove={() => handleInsertRow(rowIndex, 'above')}
                  onInsertBelow={() => handleInsertRow(rowIndex, 'below')}
                  menuOpen={rowMenu?.row === rowIndex && rowMenu.open}
                  onMenuToggle={(open) => setRowMenu({ row: rowIndex, open })}
                />
              ))}
              <EmptyRow columns={localColumns} onAdd={handleAddRowWithData} />
            </tbody>
          </table>
        </div>

        {/* Bottom Bar */}
        <div className="border-t p-3 bg-gray-50 flex justify-between items-center text-sm text-gray-600">
          <div className="flex items-center gap-4">
            <span className="flex items-center gap-1">
              <Edit2 className="w-4 h-4" /> Click edit buttons to modify
            </span>
          </div>
          <button onClick={onClose} className="px-4 py-1.5 border rounded hover:bg-gray-200">
            Cancel
          </button>
        </div>
      </div>

      {/* Modals */}
      <RowEditModal
        isOpen={editingRow !== null}
        onClose={() => setEditingRow(null)}
        row={editingRow?.data || {}}
        columns={localColumns}
        onSave={handleSaveRow}
      />

      <ColumnEditModal
        isOpen={editingColumn !== null}
        onClose={() => setEditingColumn(null)}
        columnName={editingColumn?.name || ''}
        onSave={handleSaveColumn}
      />
    </div>
  );
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
  
  // UI State - sidebar collapsed by default
  const [sidebarCollapsed, setSidebarCollapsed] = useState(true);
  const [templateSearch, setTemplateSearch] = useState("");
  const [expandedSections, setExpandedSections] = useState({
    templates: true,
    status: true,
    placeholders: true,
    columns: true
  });
  
  // File System Access API handle
  const [excelFileHandle, setExcelFileHandle] = useState<FileSystemFileHandle | null>(null);
  const [originalFileName, setOriginalFileName] = useState<string>("");
  
  const [filterColumn, setFilterColumn] = useState<string>("");
  const [filterValue, setFilterValue] = useState<string>("");
  const [uniqueValues, setUniqueValues] = useState<string[]>([]);
  const [filteredData, setFilteredData] = useState<CertificateData[]>([]);
  const [isFiltered, setIsFiltered] = useState(false);
  const [filteredIndex, setFilteredIndex] = useState(0);
  const [showFilterModal, setShowFilterModal] = useState(false);
  const [filterConditions, setFilterConditions] = useState<FilterCondition[]>([]);
  const [tempFilterConditions, setTempFilterConditions] = useState<FilterCondition[]>([]);
  
  const [showExcelEditor, setShowExcelEditor] = useState(false);
  const [editableData, setEditableData] = useState<CertificateData[]>([]);
  const [newRowData, setNewRowData] = useState<CertificateData>({});

  const toggleSection = (section: keyof typeof expandedSections) => {
    setExpandedSections(prev => ({
      ...prev,
      [section]: !prev[section]
    }));
  };

  React.useEffect(() => {
    const loadData = async () => {
      try {
        const templates = await getAllTemplates();
        console.log(`✅ Loaded ${templates.length} template(s)`);

        setSavedTemplates(templates);

        if (templates.length > 0) {
          const first = templates[0];
          setDocxHtml(first.html);
          setDocxBinary(base64ToArrayBuffer(first.binary));
          
          // Clean placeholders when loading
          const cleanedPlaceholders = first.placeholders
            .map(p => p.replace(/<[^>]*>/g, '').trim())
            .filter(p => p && p.length > 0 && p.length < 50)
            .filter(p => !p.match(/^[0-9a-f]{8}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{12}$/i));
          
          setPlaceholders(cleanedPlaceholders);
          setUploadStatus((prev) => ({ ...prev, docx: true }));
          setSelectedTemplateId(first.id);
          console.log(`🎯 Auto-loaded template: ${first.name}`);
        }

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
    const chunkSize = 0x8000;
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

    console.log("📄 Uploading DOCX file:", file.name, "Size:", file.size, "bytes");

    try {
      const arrayBuffer = await file.arrayBuffer();
      console.log("✅ File read successfully, buffer size:", arrayBuffer.byteLength);
      
      setDocxBinary(arrayBuffer);
      
      // Generate HTML preview with mammoth
      try {
        console.log("🔄 Converting to HTML with mammoth...");
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer.slice(0) });
        console.log("✅ Mammoth conversion successful");
        setDocxHtml(result.value);
      } catch (mammothError) {
        console.warn("⚠️ Mammoth conversion failed:", mammothError);
        setDocxHtml("<p>Preview not available</p>");
      }
      
      // Extract placeholders by scanning all XML files
      console.log("🔄 Scanning all XML files for placeholders...");
      const zip = new PizZip(arrayBuffer);
      const xmlFiles = Object.keys(zip.files).filter(f => f.endsWith('.xml'));
      
      const placeholderSet = new Set<string>();
      
      xmlFiles.forEach(filename => {
        const content = zip.files[filename].asText();
        
        // Find all {placeholder} patterns
        const matches = content.match(/\{([^}]+)\}/g) || [];
        matches.forEach(match => {
          let placeholder = match.replace(/{|}/g, '').trim();
          
          // Remove ALL XML tags
          placeholder = placeholder.replace(/<[^>]*>/g, '').trim();
          
          // Split by XML tag boundaries that might have been removed
          const parts = placeholder.split(/[\s<>]+/);
          parts.forEach(part => {
            // Clean each part
            const cleanPart = part.replace(/[^a-zA-Z0-9_\-]/g, '').trim();
            if (cleanPart && cleanPart.length > 1 && cleanPart.length < 50) {
              // Skip UUIDs and random strings
              if (!cleanPart.match(/^[0-9a-f]{8}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{12}$/i)) {
                placeholderSet.add(cleanPart);
              }
            }
          });
        });
        
        // Also look for patterns that might be split across XML elements
        const textParts = content.split(/<[^>]*>/);
        textParts.forEach(part => {
          // Look for DATE_ISO patterns
          if (part.includes('DATE_ISO') || part.includes('-DATE_ISO')) {
            const words = part.match(/[a-zA-Z0-9_\-]+/g) || [];
            words.forEach(word => {
              if (word.includes('DATE_ISO') || word.includes('-DATE_ISO')) {
                placeholderSet.add(word);
              }
            });
          }
          
          // Look for simple word placeholders
          const words = part.match(/[a-zA-Z_]+/g) || [];
          words.forEach(word => {
            if (word.length > 1 && word.length < 30 && 
                !word.match(/^[0-9]+$/) &&
                !word.match(/^(w:|r:|t:|p:|xml|rsid)/i)) {
              placeholderSet.add(word);
            }
          });
        });
      });
      
      // Convert Set to Array and do final cleanup
      let foundPlaceholders = Array.from(placeholderSet)
        .filter(p => p && p.length > 0)
        .filter(p => !p.match(/^(w:|r:|t:|p:|m:|v:|a:|o:|d:|wp|pic|rel)/i)) // Remove XML namespace prefixes
        .filter(p => !p.match(/^[0-9a-f]{8}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{12}$/i)); // Remove UUIDs
      
      // Sort by length (shorter first) to prioritize real placeholders
      foundPlaceholders.sort((a, b) => a.length - b.length);
      
      // Remove any that are substrings of others (e.g., "DATE" vs "DATE_ISO")
      const uniquePlaceholders: string[] = [];
      foundPlaceholders.forEach(p => {
        if (!uniquePlaceholders.some(existing => existing.includes(p) && existing !== p)) {
          uniquePlaceholders.push(p);
        }
      });
      
      console.log("📋 Final cleaned placeholders:", uniquePlaceholders);
      
      setPlaceholders(uniquePlaceholders);
      setUploadStatus(prev => ({ ...prev, docx: true }));
      setSelectedTemplateId(null);
      
      const baseName = file.name.replace(/\.docx$/i, '');
      setTemplateName(baseName);
      setShowSaveDialog(true);
      
    } catch (error) {
      console.error("❌ Fatal error reading DOCX:", error);
      alert(`Error reading DOCX file: ${error.message}. Please check the file format and try again.`);
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
    
    // Clean the placeholders when loading from saved template
    const cleanedPlaceholders = template.placeholders
      .map(p => p.replace(/<[^>]*>/g, '').trim())
      .filter(p => p && p.length > 0 && p.length < 50)
      .filter(p => !p.match(/^[0-9a-f]{8}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{4}[-]?[0-9a-f]{12}$/i));
    
    setPlaceholders(cleanedPlaceholders);
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
      setData([]);
      setExcelColumns([]);
      setUploadStatus((prev) => ({ ...prev, excel: false }));
      setFilterColumn("");
      setFilterValue("");
      setUniqueValues([]);
      setCurrentIndex(0);
      setRangeStart(1);
      setRangeEnd(1);
      setIsFiltered(false);
      setFilterColumn("");
      setFilterValue("");
      setFilteredData([]);
      setFilterConditions([]);
      setTempFilterConditions([]);
      setEditableData([]);
      setNewRowData({});
      setExcelFileHandle(null);
      setOriginalFileName("");
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

    try {
      // Store original filename
      setOriginalFileName(file.name);
      
      // Reset file handle when uploading new file
      setExcelFileHandle(null);

      // Parse the file as usual
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
            setEditableData(processedData);
            setExcelColumns(columns);
            setCurrentIndex(0);
            setRangeEnd(processedData.length);
            setUploadStatus((prev) => ({ ...prev, excel: true }));

            setIsFiltered(false);
            setFilterColumn("");
            setFilterValue("");
            setFilteredData([]);
            setFilterConditions([]);
            setTempFilterConditions([]);

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
    } catch (error) {
      console.error("Error in Excel upload:", error);
      alert("Error uploading Excel file.");
    }
  };

  // Request write permission for direct file saving
  const handleRequestWritePermission = async () => {
    try {
      if (!('showOpenFilePicker' in window)) {
        alert("Your browser doesn't support direct file saving. Please use the download option instead.");
        return;
      }

      // Ask user to select the same file again (this time for write permission)
      const [handle] = await window.showOpenFilePicker({
        startIn: 'documents',
        suggestedName: originalFileName,
        types: [{
          description: 'Excel Files',
          accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
            'application/vnd.ms-excel': ['.xls']
          }
        }]
      });
      
      // Request write permission
      const permission = await handle.requestPermission({ mode: 'readwrite' });
      
      if (permission === 'granted') {
        setExcelFileHandle(handle);
        alert(`✅ Write access granted for: ${originalFileName}`);
      } else {
        alert("Write permission denied. You can still download the file.");
      }
    } catch (error) {
      console.log("Permission request cancelled or failed:", error);
    }
  };

  // Helper function to convert string to array buffer
  const s2ab = (s: string) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  };

  // Save changes using File System Access API
  const handleSaveExcelChanges = async () => {
    try {
      // Create Excel file from edited data
      const ws = XLSX.utils.json_to_sheet(editableData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
      const buffer = s2ab(wbout);

      if (excelFileHandle && 'showSaveFilePicker' in window) {
        // We have a file handle - try to save directly to original file
        try {
          // Verify we still have permission
          const permission = await excelFileHandle.queryPermission({ mode: 'readwrite' });
          
          if (permission === 'granted') {
            // Write directly to the original file
            const writable = await excelFileHandle.createWritable();
            await writable.write(buffer);
            await writable.close();
            
            alert(`✅ File saved directly to: ${originalFileName}`);
          } else {
            // Need to request permission again
            const newPermission = await excelFileHandle.requestPermission({ mode: 'readwrite' });
            if (newPermission === 'granted') {
              const writable = await excelFileHandle.createWritable();
              await writable.write(buffer);
              await writable.close();
              alert(`✅ File saved directly to: ${originalFileName}`);
            } else {
              throw new Error('Permission denied');
            }
          }
        } catch (writeError) {
          console.error("Write error:", writeError);
          // Fallback to download
          XLSX.writeFile(wb, originalFileName || 'data.xlsx');
          alert(`⚠️ Could not save directly. File downloaded instead.`);
        }
      } else {
        // No file handle - use download
        XLSX.writeFile(wb, originalFileName || 'data.xlsx');
        alert(`✅ File downloaded as: ${originalFileName || 'data.xlsx'}`);
      }

      // Update app data regardless
      setData(editableData);
      await saveExcelData(editableData, excelColumns);
      setShowExcelEditor(false);

    } catch (error) {
      console.error("Error saving Excel:", error);
      alert("Error saving file. Please try again.");
    }
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

    console.log("🎯 Generating DOCX with placeholders:", placeholders);
    console.log("📊 Current record:", record);

    const zip = new PizZip(docxBinary);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => "",
    });

    const templateData: { [key: string]: string } = {};

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
    templateData["DATE_ISO"] = now.toISOString().split("T")[0];
    templateData["YEAR"] = now.getFullYear().toString();
    templateData["MONTH"] = (now.getMonth() + 1).toString().padStart(2, "0");
    templateData["DAY"] = now.getDate().toString().padStart(2, "0");
    
    const templateName = selectedTemplateId 
      ? savedTemplates.find(t => t.id === selectedTemplateId)?.name || 'Certificate'
      : 'Certificate';
    
    const certNumber = getNextCertificateNumber(
      selectedTemplateId || 'temp_' + Date.now(),
      templateName
    );
    
    console.log("🔢 Generated certificate number:", certNumber);

    // Process each placeholder
    placeholders.forEach((placeholder) => {
      console.log("🔍 Processing placeholder:", placeholder);
      
      // Handle _UPPER suffix
      if (placeholder.endsWith("_UPPER")) {
        const baseName = placeholder.replace("_UPPER", "");
        const matchingKey = Object.keys(record).find(
          (key) => key.toLowerCase() === baseName.toLowerCase(),
        );
        if (matchingKey) {
          templateData[placeholder] =
            record[matchingKey]?.toString().toUpperCase() || "";
          console.log(`✅ Mapped ${placeholder} to uppercase:`, templateData[placeholder]);
        }
      }
      // Handle ANY placeholder that contains DATE_ISO (with underscore)
      else if (placeholder.includes("DATE_ISO")) {
        console.log(`🎯 Found DATE_ISO placeholder: ${placeholder} -> using cert number: ${certNumber}`);
        templateData[placeholder] = certNumber;
      }
      // Regular placeholder from Excel data
      else {
        // Try exact match first
        let matchingKey = Object.keys(record).find(
          (key) => key === placeholder
        );
        
        // If no exact match, try case-insensitive match
        if (!matchingKey) {
          matchingKey = Object.keys(record).find(
            (key) => key.toLowerCase() === placeholder.toLowerCase()
          );
        }
        
        if (matchingKey) {
          templateData[placeholder] = record[matchingKey]?.toString() || "";
          console.log(`✅ Mapped ${placeholder} to:`, templateData[placeholder]);
        } else {
          console.log(`⚠️ No match found for placeholder: ${placeholder}`);
        }
      }
    });

    console.log("📦 Final template data:", templateData);
    
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
    
    const filtered = applyFilters(updatedConditions);
    setFilteredData(filtered);
  };

  const handleApplyFilters = () => {
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
    setFilterValue("");
    setIsFiltered(false);
    setFilteredData([]);

    const values = Array.from(
      new Set(data.map((row) => row[column]?.toString()).filter(Boolean)),
    );

    setUniqueValues(values as string[]);
  };

  const handleDownloadBySelectedCriteria = () => {
    const conditions = filterConditions.length > 0 
      ? filterConditions 
      : tempFilterConditions.filter(c => c.column && c.value);
    
    if (conditions.length === 0) {
      alert("Please add at least one filter condition.");
      return;
    }

    const filtered = applyFilters(conditions);

    if (filtered.length === 0) {
      alert("No matching records found.");
      return;
    }

    filtered.forEach((record, idx) => {
      setTimeout(() => {
        try {
          const blob = generateDocx(record);
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

  const handleViewCounters = () => {
    const counters = loadCounters();
    console.log('📊 Current Counters:', counters);
    
    const tableData = counters.map(c => ({
      Template: c.templateName,
      'Month': c.monthKey,
      'Downloads': c.count
    }));
    
    console.table(tableData);
    
    if (counters.length === 0) {
      alert('No download counters yet');
    } else {
      const summary = counters
        .map(c => `${c.templateName} (${c.monthKey}): ${c.count} downloads`)
        .join('\n');
      alert(`Download Counters:\n${summary}`);
    }
  };

  const handleResetCounters = () => {
    if (confirm('Are you sure you want to reset all download counters?')) {
      localStorage.removeItem(COUNTERS_KEY);
      alert('Counters reset successfully');
    }
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
    <div className="flex h-screen overflow-hidden bg-gradient-to-br from-purple-50 to-blue-100">
      {/* Sidebar */}
      <div className={`bg-white shadow-xl transition-all duration-300 ${
        sidebarCollapsed ? 'w-20' : 'w-80'
      } flex flex-col h-full overflow-hidden relative`}>
        
        {/* Collapse Toggle Button - stays fixed */}
        <button
          onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
          className="absolute -right-3 top-10 bg-white rounded-full p-1.5 shadow-md hover:shadow-lg transition border border-gray-200 z-10"
          title={sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar"}
        >
          <ChevronLeft className={`w-4 h-4 text-gray-600 transition-transform duration-300 ${
            sidebarCollapsed ? 'rotate-180' : ''
          }`} />
        </button>

        {/* Sidebar Content - scrollable */}
        <div className="flex-1 overflow-y-auto p-6">
          {/* App Title */}
          <div className="flex items-center gap-3 mb-6">
            <div className="p-2 bg-purple-100 rounded-lg flex-shrink-0">
              <FileText className="w-6 h-6 text-purple-600" />
            </div>
            {!sidebarCollapsed && (
              <h2 className="text-2xl font-bold text-gray-800">📁 Certificates</h2>
            )}
          </div>

          {/* Templates Section */}
          <div className="mb-6">
            {/* Section Header */}
            <div className="flex items-center justify-between mb-3">
              <button
                onClick={() => toggleSection('templates')}
                className="flex items-center gap-2 text-gray-700 hover:text-purple-600 transition"
              >
                <FolderOpen className="w-5 h-5 flex-shrink-0" />
                {!sidebarCollapsed && <span className="font-semibold">Templates</span>}
                {!sidebarCollapsed && (
                  <ChevronDown className={`w-4 h-4 transition-transform ${
                    expandedSections.templates ? 'rotate-180' : ''
                  }`} />
                )}
              </button>
              {!sidebarCollapsed && (
                <span className="text-sm text-gray-500 bg-gray-100 px-2 py-1 rounded-full">
                  {savedTemplates.length}
                </span>
              )}
            </div>

            {/* Search Bar - Only show when expanded */}
            {!sidebarCollapsed && expandedSections.templates && (
              <div className="mb-4 relative">
                <Search className="w-4 h-4 absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" />
                <input
                  type="text"
                  placeholder="Search templates..."
                  value={templateSearch}
                  onChange={(e) => setTemplateSearch(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-purple-500 focus:border-transparent"
                />
                {templateSearch && (
                  <button
                    onClick={() => setTemplateSearch("")}
                    className="absolute right-3 top-1/2 transform -translate-y-1/2 text-gray-400 hover:text-gray-600"
                  >
                    <X className="w-4 h-4" />
                  </button>
                )}
              </div>
            )}

            {/* Templates List */}
            {!sidebarCollapsed && expandedSections.templates && (
              <div className="space-y-3">
                {savedTemplates.length > 0 ? (
                  <>
                    {/* Filtered results count */}
                    {templateSearch && (
                      <p className="text-xs text-gray-500 mb-2">
                        Found {savedTemplates.filter(t => 
                          t.name.toLowerCase().includes(templateSearch.toLowerCase())
                        ).length} template(s)
                      </p>
                    )}
                    
                    {/* Template Grid/List */}
                    <div className="grid grid-cols-2 gap-2">
                      {savedTemplates
                        .filter(t => t.name.toLowerCase().includes(templateSearch.toLowerCase()))
                        .map((template) => (
                          <div
                            key={template.id}
                            className={`group relative rounded-lg border-2 cursor-pointer transition-all hover:shadow-md ${
                              selectedTemplateId === template.id
                                ? "border-purple-500 bg-purple-50"
                                : "border-gray-200 hover:border-purple-300 bg-white"
                            }`}
                            onClick={() => handleLoadTemplate(template)}
                          >
                            {/* Template Thumbnail */}
                            <div className="aspect-[3/4] bg-gradient-to-br from-gray-50 to-gray-100 rounded-t-lg p-2 relative overflow-hidden">
                              {/* Mini certificate preview */}
                              <div className="text-[6px] leading-tight">
                                <div className="font-bold truncate">{template.name}</div>
                                <div className="border-t border-gray-300 my-1"></div>
                                {template.placeholders.slice(0, 3).map((ph, i) => (
                                  <div key={i} className="text-gray-600 truncate">{"{" + ph + "}"}</div>
                                ))}
                                {template.placeholders.length > 3 && (
                                  <div className="text-gray-400">+{template.placeholders.length - 3}</div>
                                )}
                              </div>
                              
                              {/* Delete button overlay */}
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleDeleteTemplate(template.id);
                                }}
                                className="absolute top-1 right-1 opacity-0 group-hover:opacity-100 bg-red-500 text-white rounded-full p-1 hover:bg-red-600 transition shadow-sm"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </div>
                            
                            {/* Template Name */}
                            <div className="p-2 text-xs font-medium text-gray-700 truncate bg-white rounded-b-lg border-t">
                              {template.name}
                            </div>
                          </div>
                        ))}
                    </div>

                    {/* Upload new template card */}
                    <label className="block border-2 border-dashed border-gray-300 rounded-lg p-4 cursor-pointer hover:border-purple-500 hover:bg-purple-50 transition text-center">
                      <Plus className="w-6 h-6 text-gray-400 mx-auto mb-1" />
                      <span className="text-xs text-gray-600">Upload New</span>
                      <input
                        type="file"
                        className="hidden"
                        accept=".docx"
                        onChange={handleDocxUpload}
                      />
                    </label>
                  </>
                ) : (
                  <div className="text-center py-8">
                    <div className="bg-gray-100 rounded-full w-16 h-16 mx-auto mb-3 flex items-center justify-center">
                      <FileText className="w-8 h-8 text-gray-400" />
                    </div>
                    <p className="text-sm text-gray-600 mb-3">No templates yet</p>
                    <label className="inline-flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg cursor-pointer hover:bg-purple-700 transition text-sm">
                      <Upload className="w-4 h-4" />
                      Upload First Template
                      <input
                        type="file"
                        className="hidden"
                        accept=".docx"
                        onChange={handleDocxUpload}
                      />
                    </label>
                  </div>
                )}
              </div>
            )}

            {/* Collapsed View - Icons only */}
            {sidebarCollapsed && (
              <div className="space-y-3 mt-4">
                {savedTemplates.slice(0, 3).map((template) => (
                  <div
                    key={template.id}
                    className={`relative rounded-lg cursor-pointer ${
                      selectedTemplateId === template.id ? 'ring-2 ring-purple-500' : ''
                    }`}
                    onClick={() => handleLoadTemplate(template)}
                    title={template.name}
                  >
                    <div className="aspect-square bg-gradient-to-br from-gray-50 to-gray-100 rounded-lg p-1">
                      <FileText className="w-full h-full text-gray-600 p-1" />
                    </div>
                  </div>
                ))}
                <label className="block aspect-square border-2 border-dashed border-gray-300 rounded-lg cursor-pointer hover:border-purple-500">
                  <Plus className="w-full h-full text-gray-400 p-2" />
                  <input
                    type="file"
                    className="hidden"
                    accept=".docx"
                    onChange={handleDocxUpload}
                  />
                </label>
              </div>
            )}
          </div>

          {/* Status Section */}
          <div className="mb-6 border-t pt-4">
            <button
              onClick={() => toggleSection('status')}
              className="flex items-center gap-2 text-gray-700 hover:text-purple-600 transition w-full"
            >
              <Activity className="w-5 h-5 flex-shrink-0" />
              {!sidebarCollapsed && (
                <>
                  <span className="font-semibold">Status</span>
                  <ChevronDown className={`w-4 h-4 ml-auto transition-transform ${
                    expandedSections.status ? 'rotate-180' : ''
                  }`} />
                </>
              )}
            </button>

            {!sidebarCollapsed && expandedSections.status && (
              <div className="mt-3 space-y-2">
                <div className="flex items-center gap-2 p-2 bg-gray-50 rounded-lg">
                  <div className={`w-2 h-2 rounded-full ${uploadStatus.docx ? 'bg-green-500' : 'bg-gray-300'}`} />
                  <FileText className="w-4 h-4 text-gray-600" />
                  <span className="text-sm flex-1">Template</span>
                  {uploadStatus.docx && <Check className="w-4 h-4 text-green-500" />}
                </div>
                <div className="flex items-center gap-2 p-2 bg-gray-50 rounded-lg">
                  <div className={`w-2 h-2 rounded-full ${uploadStatus.excel ? 'bg-green-500' : 'bg-gray-300'}`} />
                  <FileSpreadsheet className="w-4 h-4 text-gray-600" />
                  <span className="text-sm flex-1">Data</span>
                  {uploadStatus.excel && (
                    <>
                      <span className="text-xs text-gray-500">{data.length} rows</span>
                      <Check className="w-4 h-4 text-green-500" />
                    </>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Placeholders Section */}
          {placeholders.length > 0 && (
            <div className="mb-6 border-t pt-4">
              <button
                onClick={() => toggleSection('placeholders')}
                className="flex items-center gap-2 text-gray-700 hover:text-purple-600 transition w-full"
              >
                <Tag className="w-5 h-5 flex-shrink-0" />
                {!sidebarCollapsed && (
                  <>
                    <span className="font-semibold">Placeholders</span>
                    <span className="ml-auto text-xs bg-gray-200 px-2 py-0.5 rounded-full">
                      {placeholders.length}
                    </span>
                    <ChevronDown className={`w-4 h-4 transition-transform ${
                      expandedSections.placeholders ? 'rotate-180' : ''
                    }`} />
                  </>
                )}
              </button>

              {!sidebarCollapsed && expandedSections.placeholders && (
                <div className="mt-3 space-y-1 max-h-40 overflow-y-auto">
                  {placeholders.map((ph) => (
                    <div key={ph} className="flex items-center gap-2 p-1.5 bg-blue-50 rounded text-xs">
                      <code className="text-blue-700 flex-1">{`{${ph}}`}</code>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Excel Columns Section */}
          {excelColumns.length > 0 && (
            <div className="mb-6 border-t pt-4">
              <button
                onClick={() => toggleSection('columns')}
                className="flex items-center gap-2 text-gray-700 hover:text-purple-600 transition w-full"
              >
                <Grid className="w-5 h-5 flex-shrink-0" />
                {!sidebarCollapsed && (
                  <>
                    <span className="font-semibold">Columns</span>
                    <span className="ml-auto text-xs bg-gray-200 px-2 py-0.5 rounded-full">
                      {excelColumns.length}
                    </span>
                    <ChevronDown className={`w-4 h-4 transition-transform ${
                      expandedSections.columns ? 'rotate-180' : ''
                    }`} />
                  </>
                )}
              </button>

              {!sidebarCollapsed && expandedSections.columns && (
                <div className="mt-3 space-y-1 max-h-40 overflow-y-auto">
                  {excelColumns.map((col) => (
                    <div key={col} className="flex items-center gap-2 p-1.5 bg-green-50 rounded text-xs">
                      <span className="text-green-700 flex-1">{col}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Main Content - independently scrollable */}
      <div className="flex-1 overflow-y-auto p-8">
        <div className="max-w-6xl mx-auto">
          {/* Upload panels removed - now starts directly with preview */}

          {isReady ? (
            <div className="bg-white rounded-lg shadow-xl p-6">
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
                <div className="flex justify-end gap-3 mb-4">
                  <button
                    onClick={handleViewCounters}
                    className="flex items-center gap-2 px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 text-sm"
                    title="View download counters"
                  >
                    📊 View Counters
                  </button>
                  <button
                    onClick={handleResetCounters}
                    className="flex items-center gap-2 px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 text-sm"
                    title="Reset all counters"
                  >
                    🔄 Reset Counters
                  </button>
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
          ) : (
            <div className="bg-white rounded-lg shadow-xl p-8">
              <h2 className="text-xl font-bold text-gray-800 mb-4">
                📝 Get Started
              </h2>
              <ol className="space-y-3 text-gray-700">
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">1.</span>
                  <span>
                    Upload a DOCX template using the <strong>"+" button</strong> in the sidebar
                  </span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">2.</span>
                  <span>
                    Upload an Excel file with your data using the <strong>Upload Excel Data</strong> button above
                  </span>
                </li>
                <li className="flex gap-3">
                  <span className="font-bold text-purple-600">3.</span>
                  <span>Generate and download your certificates!</span>
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
          <div className="bg-white rounded-lg p-6 w-[500px] max-h-[80vh] overflow-y-auto shadow-2xl">
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

              <button
                onClick={handleAddFilterCondition}
                className="w-full px-4 py-2 border-2 border-dashed border-purple-300 text-purple-600 rounded-lg hover:bg-purple-50 transition-colors"
              >
                + Add Another Filter
              </button>

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
      <ExcelEditorModal
        isOpen={showExcelEditor}
        onClose={() => setShowExcelEditor(false)}
        data={editableData}
        columns={excelColumns}
        fileName={originalFileName}
        onSave={(newData, newColumns) => {
          setData(newData);
          setEditableData(newData);
          setExcelColumns(newColumns);
          saveExcelData(newData, newColumns);
        }}
      />

      {/* File info with permission button - only show when needed */}
      {excelFileHandle && (
        <div className="fixed bottom-4 right-4 bg-green-50 border border-green-200 rounded-lg p-3 shadow-lg z-40">
          <div className="flex items-center gap-2 text-green-700">
            <Check className="w-4 h-4" />
            <span className="text-sm">Write access: {originalFileName}</span>
          </div>
        </div>
      )}
    </div>
  );
};

export default CertificateGenerator;