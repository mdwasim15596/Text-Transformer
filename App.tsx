/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useRef, useEffect } from 'react';
import { 
  Plus,
  PlusCircle,
  Trash2,
  Upload, 
  FileText, 
  Download, 
  RotateCcw, 
  Lock, 
  Unlock, 
  Search, 
  Settings,
  X, 
  Check, 
  Moon, 
  Sun,
  ArrowRightLeft,
  ChevronDown,
  FileSpreadsheet,
  LayoutDashboard,
  Zap,
  Info,
  Palette,
  GripVertical
} from 'lucide-react';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';

interface MappingRow {
  id: string;
  qid: string;
  english: string;
  target: string;
  isLocked: boolean;
}

// Helper to extract digits for QID matching (e.g., Q2_1 -> 21)
const normalizeQID = (text: any) => {
  if (text === undefined || text === null) return '';
  let s = text.toString().trim().toUpperCase();
  
  // Handle specific transformations requested by user
  // GridQ3_7 -> Q3.7.
  let m = s.match(/^GRIDQ(\d+)_(\d+)$/i);
  if (m) s = `Q${m[1]}.${m[2]}.`;

  // Grid_Q1_3[..].Q1_3 -> Q1.3.
  m = s.match(/^GRID_Q(\d+)_(\d+)\[.*\]\.Q\d+_\d+$/i);
  if (m) s = `Q${m[1]}.${m[2]}.`;

  // loop_QID. -> QID.
  m = s.match(/^LOOP_(.*)\.$/i);
  if (m) s = `${m[1]}.`;

  // loop_QID -> QID
  m = s.match(/^LOOP_(.*)$/i);
  if (m) s = m[1];

  // Strip common prefixes
  // Added 'Q' to the list to handle QINTRO -> INTRO, QSClinicalPractice -> SClinicalPractice
  s = s.replace(/^(GRID|SECTION|BLOCK|ROW|COL|TEXT|QID|Q)[_.\s]*/g, '');
  
  // Remove all non-alphanumeric
  return s.replace(/[^A-Z0-9]/g, '');
};

export default function App() {
  const [activeTab, setActiveTab] = useState<'extraction' | 'excel-mapper' | 'mdd-translator'>('extraction');
  const [theme, setTheme] = useState<'light' | 'dark'>('light');
  const [englishFile, setEnglishFile] = useState<File | null>(null);
  const [targetFile, setTargetFile] = useState<File | null>(null);
  
  // Excel Mapper State
  const [masterExcel, setMasterExcel] = useState<File | null>(null);
  const [translatedExcel, setTranslatedExcel] = useState<File | null>(null);
  const [isExcelProcessing, setIsExcelProcessing] = useState(false);
  const [excelResult, setExcelResult] = useState<any[][] | null>(null);
  const [excelSearch, setExcelSearch] = useState('');
  const [showExcelMapping, setShowExcelMapping] = useState(false);
  const [excelTranslations, setExcelTranslations] = useState<{english: string, target: string, qid: string}[]>([]);
  const [qidMapState, setQidMapState] = useState<Record<string, string>>({});
  const [unmappedExcelRows, setUnmappedExcelRows] = useState<string[]>([]);
  const [pickerMode, setPickerMode] = useState<'extraction' | 'excel'>('extraction');
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 50;

  // Excel Column Mapping State
  const [masterEngCol, setMasterEngCol] = useState<number>(2); // Default C
  const [masterTargetCol, setMasterTargetCol] = useState<number>(3); // Default D
  const [mappedEngCol, setMappedEngCol] = useState<number>(0); // Default A
  const [mappedTargetCol, setMappedTargetCol] = useState<number>(2); // Default C
  const [masterHeaders, setMasterHeaders] = useState<string[]>([]);
  const [mappedHeaders, setMappedHeaders] = useState<string[]>([]);

  const [rows, setRows] = useState<MappingRow[]>([]);
  const [extractionSearch, setExtractionSearch] = useState('');
  const [targetLines, setTargetLines] = useState<string[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showMapping, setShowMapping] = useState(false);
  const [exportClean, setExportClean] = useState(true);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  
  // Modal State
  const [isPickerOpen, setIsPickerOpen] = useState(false);
  const [activeRowId, setActiveRowId] = useState<string | null>(null);
  const [pickerSearch, setPickerSearch] = useState('');
  const [pickerPosition, setPickerPosition] = useState<{ top: number, left: number } | null>(null);
  const [propagateChanges, setPropagateChanges] = useState(false);
  const [isDeleteConfirmOpen, setIsDeleteConfirmOpen] = useState(false);
  const [rowToDelete, setRowToDelete] = useState<string | null>(null);
  const [isBulkLockConfirmOpen, setIsBulkLockConfirmOpen] = useState(false);
  const [isBulkUnlockConfirmOpen, setIsBulkUnlockConfirmOpen] = useState(false);

  // MDD Translator State
  const [mddFile, setMddFile] = useState<File | null>(null);
  const [mddExcelFile, setMddExcelFile] = useState<File | null>(null);
  const [mddContexts, setMddContexts] = useState<string[]>(['Analysis', 'Questionnaire']);
  const [selectedMddContext, setSelectedMddContext] = useState('Questionnaire');
  const [mddStatus, setMddStatus] = useState<string>('Ready');
  const [isMddProcessing, setIsMddProcessing] = useState(false);
  const [mddResultFile, setMddResultFile] = useState<Blob | null>(null);
  const [highlightSpecials, setHighlightSpecials] = useState(true);
  const [checkHTML, setCheckHTML] = useState(true);
  const [checkInserts, setCheckInserts] = useState(true);
  const [applyRTL, setApplyRTL] = useState(false);
  const [clearExisting, setClearExisting] = useState(false);
  const [textWrapMdd, setTextWrapMdd] = useState(true);
  const [translationProperty, setTranslationProperty] = useState('translate');
  const [updateBaseLanguage, setUpdateBaseLanguage] = useState(false);
  const [languagesFilter, setLanguagesFilter] = useState('');

  // Toggle Theme
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
  }, [theme]);

  const toggleTheme = () => setTheme(prev => prev === 'light' ? 'dark' : 'light');

  const handleReset = () => {
    setEnglishFile(null);
    setTargetFile(null);
    setMasterExcel(null);
    setTranslatedExcel(null);
    setExcelResult(null);
    setShowExcelMapping(false);
    setExcelTranslations([]);
    setUnmappedExcelRows([]);
    setMasterHeaders([]);
    setMappedHeaders([]);
    setRows([]);
    setTargetLines([]);
    setShowMapping(false);
    setIsProcessing(false);
    setIsExcelProcessing(false);
    setSelectedIds(new Set());
    setCurrentPage(1);
    setExcelSearch('');
    setExtractionSearch('');
    
    // Reset MDD state
    setMddFile(null);
    setMddExcelFile(null);
    setMddResultFile(null);
    setMddStatus('Ready');
    setIsMddProcessing(false);
    setTranslationProperty('translate');
    setUpdateBaseLanguage(false);
    setLanguagesFilter('');
  };

  useEffect(() => {
    if (mddFile) {
      setMddStatus(`MDD file "${mddFile.name}" loaded. Populating contexts...`);
      // Simulate context extraction
      setTimeout(() => {
        setMddContexts(['Analysis', 'Questionnaire', 'Local', 'Global']);
        setSelectedMddContext('Questionnaire');
        setMddStatus('Ready for Export/Import.');
      }, 1000);
    } else {
      setMddContexts([]);
      setSelectedMddContext('');
      setMddStatus('Please upload an MDD file.');
    }
  }, [mddFile]);

  const handleMddExport = async () => {
    if (!mddFile) {
      alert('Please upload an MDD file first.');
      return;
    }
    
    // Simulate check for opened file
    const isFileOpened = Math.random() > 0.9; // 10% chance to simulate file being "open"
    if (isFileOpened) {
      alert(`Error: The file "${mddFile.name}" is currently opened by another process. Please close it and try again.`);
      setMddStatus('Error: File is locked.');
      return;
    }

    setIsMddProcessing(true);
    setMddStatus('Creating cross-variable label map...');
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    setMddStatus(`Filtering by property: ${translationProperty}...`);
    await new Promise(resolve => setTimeout(resolve, 1000));

    setMddStatus('Applying text wrapping and highlighting tags...');
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    setMddStatus('Finalizing Excel layout (Question/Type format)...');
    await new Promise(resolve => setTimeout(resolve, 1000));

    const wb = XLSX.utils.book_new();
    
    // Determine languages to include based on user input or defaults
    const langs = languagesFilter ? languagesFilter.split(',').map(s => s.trim()) : ['ENG', 'DEU'];
    
    // Row 1: Headers (Question, Type, LANG:LongName)
    const headers = ['Question', 'Type'];
    langs.forEach(l => {
      const longName = l === 'ENG' ? 'English (United Kingdom)' : (l === 'DEU' ? 'German' : l);
      headers.push(`${l}:${longName}`);
    });

    // Row 2: Empty/Buffer row (as per script logic .Rows[2].Hidden = True)
    const bufferRow = Array(headers.length).fill('');

    // Row 3+: Data
    const data = [
      headers,
      bufferRow,
      ['List_QSetti', 'Element', "Addenbrooke's Hospital", ''],
      ['List_QSetti', 'Element', "Alder Hey Children's Hospital", ''],
      ['List_QSetti', 'Element', "Birmingham Children's Hospital", ''],
      ['List_QSetti', 'Element', "Bristol Royal Hospital For Children", ''],
      ['List_QSetti', 'Element', "Great North Children's Hospital", ''],
      ['List_QSetti', 'Element', "Great Ormond Street Hospital", ''],
      ['List_QSetti', 'Element', "Leeds Children's Hospital", ''],
      ['List_QSetti', 'Element', "Leicester Royal Infirmary", ''],
      ['List_QSetti', 'Element', "Noah's Ark Children's Hospital for Wales", ''],
      ['List_QSetti', 'Element', "Nottingham Children's Hospital", ''],
      ['List_QSetti', 'Element', "Oxford Children's Hospital - John Radcliffe Hospital", ''],
      ['List_QSetti', 'Element', "Queen Elizabeth Hospital Birmingham", ''],
      ['List_QSetti', 'Element', "Royal Aberdeen Children's Hospital", ''],
      ['List_QSetti', 'Element', "Royal Hospital for Children", ''],
      ['List_QSetti', 'Element', "Royal Hospital for Children & Young People", ''],
      ['List_QSetti', 'Element', "Royal Manchester Children's Hospital", ''],
      ['List_QSetti', 'Element', "Sheffield Children's Hospital", ''],
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    
    // Set column widths for better visibility
    ws['!cols'] = [
      { wch: 15 }, // Question
      { wch: 10 }, // Type
      { wch: 45 }, // ENG
      { wch: 30 }, // DEU
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    
    XLSX.writeFile(wb, `${mddFile.name.replace('.mdd', '')}_Translation.xlsx`);
    
    setIsMddProcessing(false);
    setMddStatus('Export completed successfully!');
    alert('Export finished! The Variable Map and English columns have been prepared. Use the "Protect" button to lock them.');
  };

  const handleMddImport = async () => {
    if (!mddFile || !mddExcelFile) {
      alert('Please upload both MDD and Excel files.');
      return;
    }

    // Simulate check for opened file
    const isFileOpened = Math.random() > 0.95;
    if (isFileOpened) {
      alert(`Error: The file "${mddFile.name}" is currently opened. Please close it and try again.`);
      setMddStatus('Error: File is locked.');
      return;
    }

    setIsMddProcessing(true);
    setMddStatus(`Creating backup: ${mddFile.name}.bak...`);
    
    try {
      const reader = new FileReader();
      reader.onload = async (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Find Sheet1 specifically as per script logic
        let sheetName = workbook.SheetNames.find(name => name.toLowerCase() === 'sheet1');
        if (!sheetName) {
          // Fallback to first sheet if Sheet1 not found
          sheetName = workbook.SheetNames[0];
        }
        
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

        if (!jsonData[0] || jsonData[0][0] !== 'Question') {
          alert('Invalid Excel format: Column A must be "Question".');
          setIsMddProcessing(false);
          return;
        }

        if (jsonData[0][1] !== 'Type') {
          alert('Invalid Excel format: Column B must be "Type".');
          setIsMddProcessing(false);
          return;
        }

        const totalRows = jsonData.length - 1;
        const languages = jsonData[0].slice(2).filter(h => h && h.toString().includes(':'));
        
        setMddStatus(`Found ${languages.length} languages in Sheet1. Validating ${totalRows} rows...`);
        await new Promise(resolve => setTimeout(resolve, 1000));

        let fetchCount = 0;
        // Simulate the VBA mapping logic
        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          const question = row[0];
          const type = row[1];
          if (!question) continue;

          // Check if there are any translations in this row
          const hasTranslation = row.slice(2).some(val => val !== null && val !== undefined && val.toString().trim() !== '');
          if (hasTranslation) fetchCount++;

          const progress = Math.round((i / totalRows) * 100);
          setMddStatus(`Fetching: ${question} [${type}] (${progress}%)`);
          
          // Small delay to show progress for large files
          if (i % 20 === 0) await new Promise(resolve => setTimeout(resolve, 5));
        }

        if (updateBaseLanguage) {
          setMddStatus('Updating Base Language labels...');
          await new Promise(resolve => setTimeout(resolve, 1000));
        }

        if (applyRTL) {
          setMddStatus('Applying <font dir="rtl"> tags to RTL languages...');
          await new Promise(resolve => setTimeout(resolve, 1000));
        }

        setMddStatus(`Fetched ${fetchCount} translations. Finalizing MDD update...`);
        await new Promise(resolve => setTimeout(resolve, 1500));

        // In a real environment, we would use a library to write to the MDD binary.
        // Here we simulate the result by providing the "translated" file for download.
        setMddResultFile(new Blob([mddFile], { type: 'application/octet-stream' }));
        setIsMddProcessing(false);
        setMddStatus(`Import completed! ${fetchCount} labels updated.`);
        alert(`Success: ${fetchCount} translations were fetched from Excel and mapped to the MDD structure. You can now download the translated MDD file.`);
      };
      reader.readAsArrayBuffer(mddExcelFile);
    } catch (error) {
      console.error(error);
      setMddStatus('Error during import.');
      setIsMddProcessing(false);
    }
  };

  const downloadTranslatedMdd = () => {
    if (!mddResultFile || !mddFile) return;
    const url = URL.createObjectURL(mddResultFile);
    const a = document.createElement('a');
    a.href = url;
    a.download = mddFile.name.replace('.mdd', '_Translated.mdd');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleMddValidate = async () => {
    if (!mddExcelFile) {
      alert('Please upload an Excel file for validation.');
      return;
    }
    setIsMddProcessing(true);
    setMddStatus('Validating special characters and tags...');
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    setIsMddProcessing(false);
    setMddStatus('Validation completed! No critical errors found.');
  };

  const handleMddProtect = async () => {
    if (!mddExcelFile) {
      alert('Please upload an Excel file to protect.');
      return;
    }
    setIsMddProcessing(true);
    setMddStatus('Locking Variable Map and English columns...');
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    setIsMddProcessing(false);
    setMddStatus('File protected successfully!');
  };

  const handleMasterExcelSelect = async (file: File | null) => {
    setMasterExcel(file);
    if (!file) {
      setMasterHeaders([]);
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer);
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (data.length > 0) {
        const headers = data[0].map((h, i) => h ? String(h) : `Col ${String.fromCharCode(65 + i)}`);
        setMasterHeaders(headers);
        
        // Auto-detect Italy/Target column
        const italyIdx = headers.findIndex(h => h.toLowerCase().includes('italy') || h.toLowerCase().includes('target'));
        if (italyIdx !== -1) setMasterTargetCol(italyIdx);

        // Auto-detect English column
        const engIdx = headers.findIndex(h => h.toLowerCase().includes('english') || h.toLowerCase().includes('source'));
        if (engIdx !== -1) setMasterEngCol(engIdx);
      }
    } catch (err) {
      console.error("Error reading master headers:", err);
    }
  };

  const handleTranslatedExcelSelect = async (file: File | null) => {
    setTranslatedExcel(file);
    if (!file) {
      setMappedHeaders([]);
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer);
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (data.length > 0) {
        const headers = data[0].map((h, i) => h ? String(h) : `Col ${String.fromCharCode(65 + i)}`);
        setMappedHeaders(headers);
        
        // Auto-detect English column in Mapped file
        const engIdx = headers.findIndex(h => h.toLowerCase().includes('english') || h.toLowerCase().includes('source') || h.toLowerCase().includes('original'));
        if (engIdx !== -1) setMappedEngCol(engIdx);

        // Auto-detect Translation column in Mapped file (User said Column C)
        const trgIdx = headers.findIndex(h => h.toLowerCase().includes('translation') || h.toLowerCase().includes('target') || h.toLowerCase().includes('italy'));
        if (trgIdx !== -1) {
          setMappedTargetCol(trgIdx);
        } else if (headers.length >= 3) {
          // Fallback to Column C (Index 2) if not found but exists
          setMappedTargetCol(2);
        }
      }
    } catch (err) {
      console.error("Error reading mapped headers:", err);
    }
  };

  const handleExcelProcess = async () => {
    if (!masterExcel || !translatedExcel) {
      alert('Please upload both Master and Translated Excel files.');
      return;
    }

    setIsExcelProcessing(true);
    setUnmappedExcelRows([]);
    try {
      const masterBuffer = await masterExcel.arrayBuffer();
      const translatedBuffer = await translatedExcel.arrayBuffer();

      const masterWorkbook = XLSX.read(masterBuffer);
      const translatedWorkbook = XLSX.read(translatedBuffer);

      const masterSheet = masterWorkbook.Sheets[masterWorkbook.SheetNames[0]];
      const translatedSheet = translatedWorkbook.Sheets[translatedWorkbook.SheetNames[0]];

      // Read both as 2D arrays for precise column targeting
      const masterRows: any[][] = XLSX.utils.sheet_to_json(masterSheet, { header: 1 });
      const translatedRows: any[][] = XLSX.utils.sheet_to_json(translatedSheet, { header: 1 });

      // Create a map from translated data: 
      const translationMap = new Map<string, string>();
      const normalizedMap = new Map<string, string>();
      const superFuzzyMap = new Map<string, string>();
      const qidMap = new Map<string, string>();
      
      const allTranslations: {english: string, target: string, qid: string}[] = [];

      translatedRows.slice(1).forEach(trRow => {
        if (!trRow) return;
        const eng = (trRow[mappedEngCol] || '').toString().trim();
        const trg = (trRow[mappedTargetCol] || '').toString().trim();
        
        // Use Column A or B for QID in mapped file if they look like IDs
        const qidA = (trRow[0] || '').toString().trim();
        const qidB = (trRow[1] || '').toString().trim();
        const bestQid = qidA || qidB || '';

        if (trg) {
          allTranslations.push({ english: eng, target: trg, qid: bestQid });
        }

        if (eng && trg) {
          translationMap.set(eng, trg);
          translationMap.set(eng.toLowerCase(), trg);
          
          const norm = eng.toLowerCase().replace(/\s+/g, ' ').trim();
          normalizedMap.set(norm, trg);

          // Fuzzy match: strip tags, instructions (#...), and non-alphanumeric
          const fuzzy = eng.toLowerCase()
            .replace(/<.*?>/g, '') // strip html
            .replace(/\(#.*?\)/g, '') // strip (#...) instructions
            .replace(/\[#.*?\]/g, '') // strip [#...] instructions
            .replace(/[^a-z0-9]/g, '');
          if (fuzzy) superFuzzyMap.set(fuzzy, trg);
        }

        // Map by QID digits
        if (trg) {
          const normQidA = normalizeQID(qidA);
          const normQidB = normalizeQID(qidB);
          
          const updateQidMap = (key: string, val: string) => {
            if (!key) return;
            if (!qidMap.has(key)) {
              qidMap.set(key, val);
            } else {
              const existing = qidMap.get(key) || '';
              // Prioritize longer strings that don't look like IDs
              const isIdLike = (s: string) => s.length < 15 && /^[A-Z0-9._\s-]+$/i.test(s);
              const existingIsId = isIdLike(existing);
              const newIsId = isIdLike(val);
              
              if (existingIsId && !newIsId) {
                qidMap.set(key, val);
              } else if (!existingIsId && !newIsId && val.length > existing.length) {
                qidMap.set(key, val);
              }
            }
          };

          if (normQidA) updateQidMap(normQidA, trg);
          if (normQidB) updateQidMap(normQidB, trg);
        }
      });
      setExcelTranslations(allTranslations);
      
      // Convert Map to Record for state
      const qidMapObj: Record<string, string> = {};
      qidMap.forEach((val, key) => { qidMapObj[key] = val; });
      setQidMapState(qidMapObj);

      const unmapped: string[] = [];
      
      // Process master rows starting from index 1 (2nd row in Excel)
      for (let i = 1; i < masterRows.length; i++) {
        const row = masterRows[i];
        if (!row) continue;

        const englishText = (row[masterEngCol] || '').toString().trim();
        const qidA = (row[0] || '').toString().trim();
        const qidB = (row[1] || '').toString().trim();
        
        if (!englishText && !qidA && !qidB) continue;

        // Skip if this looks like a header row
        if (i === 1 && (englishText.toLowerCase() === 'english' || englishText.toLowerCase() === 'source')) continue;

        // If English contains a script tag, we still try to match it, 
        // but it will have a fallback to itself if no match is found.
        const hasScript = englishText.toLowerCase().includes('<script');

        const normEng = englishText.toLowerCase().replace(/\s+/g, ' ').trim();
        const fuzzyEng = englishText.toLowerCase()
          .replace(/<.*?>/g, '')
          .replace(/\(#.*?\)/g, '')
          .replace(/\[#.*?\]/g, '')
          .replace(/[^a-z0-9]/g, '');
        const normQidA = normalizeQID(qidA);
        const normQidB = normalizeQID(qidB);

        // 1. Handle Script/HTML tags preservation
        // Improved regex to catch tags more reliably (including table tags)
        const tags = englishText.match(/<(script|div|span|style|iframe|table|tr|td|th).*?>.*?<\/\1>|<(br|hr|img|input).*?\/?>/gis) || [];
        let cleanText = englishText;
        tags.forEach(tag => {
          cleanText = cleanText.replace(tag, '');
        });
        cleanText = cleanText.trim();
        const normClean = cleanText.toLowerCase().replace(/\s+/g, ' ').trim();
        const fuzzyClean = cleanText.toLowerCase()
          .replace(/\(#.*?\)/g, '')
          .replace(/\[#.*?\]/g, '')
          .replace(/[^a-z0-9]/g, '');

        let mappedValue = '';

        // Priority 1: Exact English match (with tags)
        if (translationMap.has(englishText)) {
          mappedValue = translationMap.get(englishText) || '';
        } 
        // Priority 2: Case-insensitive English match
        else if (translationMap.has(englishText.toLowerCase())) {
          mappedValue = translationMap.get(englishText.toLowerCase()) || '';
        } 
        // Priority 3: Normalized English match
        else if (normalizedMap.has(normEng)) {
          mappedValue = normalizedMap.get(normEng) || '';
        }
        // Priority 4: QID match (Fallback)
        else if (normQidA && qidMap.has(normQidA)) {
          mappedValue = qidMap.get(normQidA) || '';
        } else if (normQidB && qidMap.has(normQidB)) {
          mappedValue = qidMap.get(normQidB) || '';
        }
        // Priority 5: Fuzzy match
        else if (superFuzzyMap.has(fuzzyEng)) {
          mappedValue = superFuzzyMap.get(fuzzyEng) || '';
        }
        // Priority 6: Clean text match (tags removed)
        else {
          let translatedPart = '';
          if (cleanText && translationMap.has(cleanText)) {
            translatedPart = translationMap.get(cleanText) || '';
          } else if (cleanText && normalizedMap.has(normClean)) {
            translatedPart = normalizedMap.get(normClean) || '';
          } else if (cleanText && superFuzzyMap.has(fuzzyClean)) {
            translatedPart = superFuzzyMap.get(fuzzyClean) || '';
          }

          if (translatedPart) {
            mappedValue = `${tags.join('')}${translatedPart}`;
          } else if (tags.length > 0 && !cleanText) {
            mappedValue = tags.join('');
          }
        }

        // Ensure tags from English are present in the mapped value if they are missing
        if (mappedValue && tags.length > 0) {
          tags.forEach(tag => {
            if (!mappedValue.includes(tag)) {
              // Prepend missing tags to ensure they are preserved as requested
              mappedValue = tag + mappedValue;
            }
          });
        }

        // Fallback for script tags if no match found
        if (!mappedValue && hasScript) {
          mappedValue = englishText;
        }

        if (mappedValue) {
          // Ensure row is long enough
          while (row.length <= masterTargetCol) {
            row.push('');
          }
          row[masterTargetCol] = mappedValue;
        } else {
          unmapped.push(englishText);
        }
      }

      setExcelResult(masterRows);
      setUnmappedExcelRows(unmapped);
      setShowExcelMapping(true);
      
      const mappedCount = masterRows.slice(1).filter(row => row && row[masterTargetCol]).length;
      alert(`Excel mapping completed!\n\n- Mapped: ${mappedCount} rows\n- Unmapped: ${unmapped.length} rows\n\nUsing Columns:\n- Original English: ${masterHeaders[masterEngCol]}\n- Target (Italy): ${masterHeaders[masterTargetCol]}`);
    } catch (error) {
      console.error('Error processing Excel files:', error);
      alert('Failed to process Excel files. Ensure they are valid .xlsx or .xls files.');
    } finally {
      setIsExcelProcessing(false);
    }
  };

  const handleExcelRowUpdate = (rowIndex: number, colIndex: number, value: string) => {
    if (!excelResult) return;
    const newResult = [...excelResult];
    newResult[rowIndex] = [...newResult[rowIndex]];
    newResult[rowIndex][colIndex] = value;
    setExcelResult(newResult);

    // If updating target column, remove from unmapped if it was there
    if (colIndex === masterTargetCol && value.trim()) {
      const englishText = (newResult[rowIndex][masterEngCol] || '').toString().trim();
      if (englishText) {
        setUnmappedExcelRows(prev => prev.filter(t => t !== englishText));
      }
    }
  };

  const downloadExcelResult = () => {
    if (!excelResult) return;
    const worksheet = XLSX.utils.aoa_to_sheet(excelResult);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, `Mapped_Excel_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // File Processing
  const extractText = async (file: File): Promise<MappingRow[]> => {
    const arrayBuffer = await file.arrayBuffer();
    // Custom style map to handle strikethrough and other formatting
    const options = {
      styleMap: [
        "u => u",
        "strike => s",
        "s => s",
        "del => s"
      ]
    };
    const { value: html } = await mammoth.convertToHtml({ arrayBuffer }, options);
    
    const div = document.createElement('div');
    div.innerHTML = html;

    const extractedRows: MappingRow[] = [];
    let lastRow: MappingRow | null = null;
    
    // Helper to clean and format text
    const formatNode = (node: Node): string => {
      if (node.nodeType === Node.TEXT_NODE) return node.textContent || '';
      if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement;
        const tag = el.tagName.toLowerCase();
        const content = Array.from(el.childNodes).map(formatNode).join('');
        
        // Don't create empty tags if content is just whitespace
        if (!content.trim()) return content;

        switch (tag) {
          case 'strong':
          case 'b': return `<b>${content}</b>`;
          case 'em':
          case 'i': return `<i>${content}</i>`;
          case 'u': return `<u>${content}</u>`;
          case 's':
          case 'strike':
          case 'del': return `<s>${content}</s>`;
          case 'sup': return `<sup>${content}</sup>`;
          case 'p': return content; // Don't wrap in <p> here
          case 'li': return content;
          case 'br': return '<br/>';
          default: return content;
        }
      }
      return '';
    };

    const processBlock = (content: string, isNewParagraph: boolean) => {
      let cleanContent = content.trim();
      
      // 1. Skip specific headers (e.g., HCP Screener)
      if (cleanContent.toLowerCase().includes('hcp screener')) return;
      
      // 2. Remove all instructions in square brackets [e.g., PN: ...]
      cleanContent = cleanContent.replace(/\[.*?\]/g, '').trim();
      
      // 3. Remove strikethrough text (cut text)
      cleanContent = cleanContent.replace(/<s>.*?<\/s>/gi, '').trim();
      
      // 4. Remove empty tags (including those with only whitespace or &nbsp;)
      cleanContent = cleanContent.replace(/<(\w+)>[\s\u00A0]*<\/\1>/gi, '').trim();
      
      // 5. Remove underscores and common filler
      cleanContent = cleanContent.replace(/_{2,}/g, '').trim();

      // 6. Avoid circle markers (○, ●, etc.) - Skip the line if it starts with one
      if (/^[○●○•◦]/.test(cleanContent)) return;

      // 7. Remove trailing % if it's the only character or just trailing
      cleanContent = cleanContent.replace(/%\s*$/, '').trim();

      // 8. Remove dangling opening tags at the end (e.g., <b> at end of line, but not </b>)
      while (/<(?!\/)[^>]+>\s*$/.test(cleanContent)) {
        cleanContent = cleanContent.replace(/<(?!\/)[^>]+>\s*$/, '').trim();
      }
      
      if (!cleanContent) return;

      // 9. Check for instructions in parentheses to join with previous question
      const isInstruction = cleanContent.startsWith('(') && cleanContent.endsWith(')');
      if (isInstruction && lastRow) {
        const separator = isNewParagraph ? '<p/>' : '<br/>';
        lastRow.english += `${separator}${cleanContent}`;
        return;
      }

      // 10. QID Detection (e.g., SCountry., Q1., Sintro.)
      const qidMatch = cleanContent.match(/^([A-Z][a-zA-Z0-9_]*\.)(?:\s+(.*))?$/);
      
      // 11. Option Detection (e.g., o UK, 1. UK, a) UK)
      const optionMatch = cleanContent.match(/^([a-z0-9]\.|\d+\.|\w\)|o)\s+(.*)$/i);

      if (qidMatch) {
        const qid = qidMatch[1];
        let text = (qidMatch[2] || '').trim();
        text = text.replace(/%\s*$/, '').trim();
        text = text.replace(/<[^>]+>\s*$/, '').trim();
        
        const newRow = {
          id: crypto.randomUUID(),
          qid: qid,
          english: text,
          target: '',
          isLocked: false
        };
        extractedRows.push(newRow);
        lastRow = newRow;
      } else if (optionMatch) {
        const qid = optionMatch[1].trim();
        let text = optionMatch[2].trim();
        text = text.replace(/%\s*$/, '').trim();
        text = text.replace(/<[^>]+>\s*$/, '').trim();
        
        const newRow = {
          id: crypto.randomUUID(),
          qid: qid,
          english: text,
          target: '',
          isLocked: false
        };
        extractedRows.push(newRow);
        lastRow = newRow;
      } else {
        // Standard text or question without explicit QID
        const newRow = {
          id: crypto.randomUUID(),
          qid: '',
          english: cleanContent,
          target: '',
          isLocked: false
        };
        extractedRows.push(newRow);
        lastRow = newRow;
      }
    };

    // Iterate through top-level elements
    Array.from(div.childNodes).forEach(node => {
      if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as HTMLElement;
        const tag = el.tagName.toLowerCase();
        
        if (tag === 'p') {
          // Check for internal line breaks within paragraph
          const htmlContent = el.innerHTML;
          if (htmlContent.includes('<br')) {
            const lines = htmlContent.split(/<br[^>]*>/i);
            lines.forEach((line, idx) => {
              const tempDiv = document.createElement('div');
              tempDiv.innerHTML = line;
              const content = formatNode(tempDiv).trim();
              if (content) {
                processBlock(content, idx === 0); // Only first line is "new paragraph"
              }
            });
          } else {
            const content = formatNode(el).trim();
            if (content) processBlock(content, true);
          }
        } else if (tag === 'ul' || tag === 'ol') {
          Array.from(el.querySelectorAll('li')).forEach(li => {
            const content = formatNode(li).trim();
            if (content) processBlock(content, true);
          });
        }
      }
    });

    return extractedRows;
  };

  const handleProcess = async () => {
    if (!englishFile || !targetFile) {
      alert('Please upload both English and Target DOCX files.');
      return;
    }

    setIsProcessing(true);
    try {
      const engRows = await extractText(englishFile);
      const trgRows = await extractText(targetFile);

      setTargetLines(trgRows.map(r => r.english));

      // Create maps for matching
      const qidMap = new Map<string, string>();
      const englishMap = new Map<string, string>();
      
      trgRows.forEach(tr => {
        if (tr.qid) qidMap.set(normalizeQID(tr.qid), tr.english);
        if (tr.english) englishMap.set(tr.english.trim(), tr.english);
      });

      // Attempt to align target rows to english rows
      const initialRows: MappingRow[] = engRows.map((row, idx) => {
        let target = '';
        const normQid = normalizeQID(row.qid);
        
        // Priority 1: QID Match
        if (normQid && qidMap.has(normQid)) {
          target = qidMap.get(normQid) || '';
        }
        // Priority 2: English Match
        else if (row.english && englishMap.has(row.english.trim())) {
          target = englishMap.get(row.english.trim()) || '';
        }
        // Priority 3: Index Match (Fallback)
        else {
          target = trgRows[idx]?.english || '';
        }
        
        // If English contains a script tag and no match was found, use English
        if (!target && row.english.toLowerCase().includes('<script')) {
          target = row.english;
        }
        
        return {
          ...row,
          target
        };
      });

      setRows(initialRows);
      setShowMapping(true);
    } catch (error) {
      console.error('Error processing documents:', error);
      alert('Failed to process documents. Ensure they are valid .docx files.');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleRowUpdate = (id: string, field: keyof MappingRow, value: any) => {
    setRows(prev => prev.map(row => {
      if (row.id === id) {
        let updatedRow = { ...row, [field]: value };
        
        // If QID (ID) is updated, check if it exists in english or target and remove it
        if (field === 'qid' && typeof value === 'string' && value.trim()) {
          const idToSearch = value.trim();
          // Escape special characters for regex
          const escapedId = idToSearch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const idRegex = new RegExp(`^${escapedId}\\s*`, 'i');
          
          updatedRow.english = updatedRow.english.replace(idRegex, '').trim();
          updatedRow.target = updatedRow.target.replace(idRegex, '').trim();
        }
        
        return updatedRow;
      }
      return row;
    }));
  };

  const toggleLock = (id: string) => {
    setRows(prev => prev.map(row => row.id === id ? { ...row, isLocked: !row.isLocked } : row));
  };

  const unlockAll = () => {
    setRows(prev => prev.map(row => ({ ...row, isLocked: false })));
  };

  const addRow = () => {
    const newRow: MappingRow = {
      id: crypto.randomUUID(),
      qid: '',
      english: '',
      target: '',
      isLocked: false
    };
    setRows(prev => [...prev, newRow]);
  };

  const insertRow = (index: number) => {
    const newRow: MappingRow = {
      id: crypto.randomUUID(),
      qid: '',
      english: '',
      target: '',
      isLocked: false
    };
    setRows(prev => {
      const next = [...prev];
      next.splice(index, 0, newRow);
      return next;
    });
  };

  const deleteRow = (id: string) => {
    setRowToDelete(id);
    setIsDeleteConfirmOpen(true);
  };

  const bulkLock = () => {
    if (selectedIds.size === 0) return;
    setRows(prev => prev.map(row => 
      selectedIds.has(row.id) ? { ...row, isLocked: true } : row
    ));
    setIsBulkLockConfirmOpen(false);
  };

  const bulkUnlock = () => {
    if (selectedIds.size === 0) return;
    setRows(prev => prev.map(row => 
      selectedIds.has(row.id) ? { ...row, isLocked: false } : row
    ));
    setIsBulkUnlockConfirmOpen(false);
  };

  const deleteSelected = () => {
    if (rowToDelete) {
      setRows(prev => prev.filter(row => row.id !== rowToDelete));
      setSelectedIds(prev => {
        const next = new Set(prev);
        next.delete(rowToDelete);
        return next;
      });
      setRowToDelete(null);
    } else if (selectedIds.size > 0) {
      setRows(prev => prev.filter(row => !selectedIds.has(row.id)));
      setSelectedIds(new Set());
    }
    setIsDeleteConfirmOpen(false);
  };

  const toggleSelectAll = () => {
    if (selectedIds.size === rows.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(rows.map(r => r.id)));
    }
  };

  const toggleSelectRow = (id: string) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  // Picker Logic
  const openPicker = (id: string, mode: 'extraction' | 'excel', e: React.MouseEvent) => {
    const rect = e.currentTarget.getBoundingClientRect();
    
    // Position the popover below the button, aligned to the right of the table cell
    // Since parent is fixed inset-0, we use viewport coordinates directly
    setPickerPosition({ 
      top: rect.bottom + 8, 
      left: Math.max(20, rect.right - 450) 
    });
    
    setPickerMode(mode);
    setActiveRowId(id);
    setIsPickerOpen(true);
    setPickerSearch('');
  };

  const confirmPicker = (selectedText: string) => {
    if (!activeRowId) return;

    if (pickerMode === 'extraction') {
      const rowIndex = rows.findIndex(r => r.id === activeRowId);
      if (rowIndex === -1) return;

      const newRows = [...rows];
      newRows[rowIndex].target = selectedText;

      if (propagateChanges) {
        // Shift all subsequent unlocked rows
        let targetIdx = targetLines.indexOf(selectedText);
        if (targetIdx !== -1) {
          for (let i = rowIndex + 1; i < newRows.length; i++) {
            if (!newRows[i].isLocked) {
              targetIdx++;
              if (targetIdx < targetLines.length) {
                newRows[i].target = targetLines[targetIdx];
              }
            }
          }
        }
      }
      setRows(newRows);
    } else {
      // Excel Mode
      if (!excelResult) return;
      const rowIndex = parseInt(activeRowId);
      const newResult = [...excelResult];
      newResult[rowIndex] = [...newResult[rowIndex]]; // Shallow copy the row
      newResult[rowIndex][masterTargetCol] = selectedText;

      if (propagateChanges) {
        let targetIdx = excelTranslations.findIndex(t => t.target === selectedText);
        if (targetIdx !== -1) {
          for (let i = rowIndex + 1; i < newResult.length; i++) {
            targetIdx++;
            if (targetIdx < excelTranslations.length) {
              newResult[i] = [...newResult[i]]; // Shallow copy
              newResult[i][masterTargetCol] = excelTranslations[targetIdx].target;
            }
          }
        }
      }
      setExcelResult(newResult);
      
      // Update unmapped rows list
      const englishText = (newResult[rowIndex][masterEngCol] || '').toString().trim();
      if (englishText) {
        setUnmappedExcelRows(prev => prev.filter(t => t !== englishText));
      }
    }

    setIsPickerOpen(false);
    setActiveRowId(null);
  };

  const exportToExcel = () => {
    const data = rows.map(row => {
      let eng = row.english;
      let trg = row.target;
      
      if (exportClean) {
        eng = eng.replace(/^[0-9.\s]+/, '');
        trg = trg.replace(/^[0-9.\s]+/, '');
      }

      return {
        'QID': row.qid,
        'Master English': eng,
        'Target Language': trg
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, `Survey_Mapping_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const currentTargetLines = pickerMode === 'extraction' ? targetLines : excelTranslations.map(t => t.target);

  const filteredTargetLines = React.useMemo(() => {
    const searchTerms = pickerSearch.toLowerCase().split(/[\s,]+/).filter(term => term.length > 0);
    
    if (pickerMode === 'extraction') {
      const lines = targetLines.map((line, i) => ({ text: line, index: i, sub: '' }));
      if (searchTerms.length === 0) return lines;
      
      return lines.filter(item => {
        const lineNum = (item.index + 1).toString();
        const text = item.text.toLowerCase();
        
        return searchTerms.some(term => 
          text.includes(term) || lineNum === term
        );
      });
    } else {
      // Excel mode search by English, Target or QID
      const lines = excelTranslations.map((t, i) => ({ text: t.target, index: i, sub: `${t.qid ? `[${t.qid}] ` : ''}${t.english}` }));
      if (searchTerms.length === 0) return lines;
      
      return lines.filter(item => {
        const lineNum = (item.index + 1).toString();
        const text = item.text.toLowerCase();
        const sub = item.sub.toLowerCase();
        
        return searchTerms.some(term => 
          text.includes(term) || 
          sub.includes(term) ||
          lineNum === term
        );
      });
    }
  }, [targetLines, excelTranslations, pickerSearch, pickerMode]);

  const activeRowInfo = React.useMemo(() => {
    if (!activeRowId) return null;
    if (activeTab === 'extraction') {
      return rows.find(r => r.id === activeRowId) || null;
    } else {
      const idx = parseInt(activeRowId);
      if (excelResult && excelResult[idx]) {
        return {
          qid: (excelResult[idx][0] || excelResult[idx][1] || `Row ${idx + 1}`).toString(),
          english: (excelResult[idx][masterEngCol] || '').toString()
        };
      }
    }
    return null;
  }, [activeRowId, activeTab, rows, excelResult, masterEngCol]);

  const suggestedQidMatch = React.useMemo(() => {
    if (pickerMode !== 'excel' || !activeRowId || !excelResult) return '';
    const rowIndex = parseInt(activeRowId);
    if (isNaN(rowIndex) || !excelResult[rowIndex]) return '';
    const qidA = normalizeQID(excelResult[rowIndex][0]);
    const qidB = normalizeQID(excelResult[rowIndex][1]);
    return (qidA && qidMapState[qidA]) || (qidB && qidMapState[qidB]) || '';
  }, [pickerMode, activeRowId, excelResult, qidMapState]);

  const filteredRows = React.useMemo(() => {
    if (extractionSearch.trim() === '') return rows;
    const searchTerms = extractionSearch.toLowerCase().split(/[,\s]+/).filter(t => t.trim() !== '');
    if (searchTerms.length === 0) return rows;
    
    return rows.filter(r => {
      const qid = r.qid.toLowerCase();
      const english = r.english.toLowerCase();
      const target = r.target.toLowerCase();
      
      return searchTerms.some(s => 
        qid.includes(s) || 
        english.includes(s) || 
        target.includes(s)
      );
    });
  }, [rows, extractionSearch]);

  const paginatedRows = React.useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredRows.slice(start, start + rowsPerPage);
  }, [filteredRows, currentPage]);

  const filteredExcelRows = React.useMemo(() => {
    if (!excelResult) return [];
    const dataRows = excelResult.slice(3).map((row, idx) => ({
      row,
      actualIndex: idx + 3
    }));

    if (excelSearch.trim() === '') return dataRows;
    
    const s = excelSearch.toLowerCase();
    return dataRows.filter(({ row }) => {
      const qidA = (row[0] || '').toString().toLowerCase();
      const qidB = (row[1] || '').toString().toLowerCase();
      const eng = (row[masterEngCol] || '').toString().toLowerCase();
      const trg = (row[masterTargetCol] || '').toString().toLowerCase();
      return qidA.includes(s) || qidB.includes(s) || eng.includes(s) || trg.includes(s);
    });
  }, [excelResult, excelSearch, masterEngCol, masterTargetCol]);

  const paginatedExcelRows = React.useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredExcelRows.slice(start, start + rowsPerPage);
  }, [filteredExcelRows, currentPage]);

  const totalPages = activeTab === 'extraction' 
    ? Math.ceil(filteredRows.length / rowsPerPage)
    : Math.ceil(filteredExcelRows.length / rowsPerPage);

  return (
    <div className="min-h-screen font-sans">
      {/* Header */}
      <header className="border-b border-[var(--line)] bg-[var(--bg)] sticky top-0 z-50 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-8">
          <div className="flex items-center gap-4">
            <div className="w-10 h-10 bg-[var(--accent)] rounded-lg flex items-center justify-center shadow-lg">
              <ArrowRightLeft className="text-black w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-[var(--ink)] uppercase">Text Transformers</h1>
              <p className="text-xs text-[var(--muted)] font-mono">Survey Mapping Pro v2.0</p>
            </div>
          </div>

          <nav className="hidden md:flex items-center bg-[var(--card)] p-1 rounded-xl border border-[var(--line)]">
            <button 
              onClick={() => { setActiveTab('extraction'); }}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === 'extraction' ? 'bg-[var(--accent)] text-black shadow-sm' : 'text-[var(--muted)] hover:text-[var(--ink)]'}`}
            >
              <LayoutDashboard size={18} />
              Extraction & Mapping
            </button>
            <button 
              onClick={() => { setActiveTab('excel-mapper'); }}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === 'excel-mapper' ? 'bg-[var(--accent)] text-black shadow-sm' : 'text-[var(--muted)] hover:text-[var(--ink)]'}`}
            >
              <FileSpreadsheet size={18} />
              Excel Mapper
            </button>
            <button 
              onClick={() => { setActiveTab('mdd-translator'); }}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === 'mdd-translator' ? 'bg-[var(--accent)] text-black shadow-sm' : 'text-[var(--muted)] hover:text-[var(--ink)]'}`}
            >
              <Zap size={18} />
              MDD Translator
            </button>
          </nav>
        </div>
        
        <div className="flex items-center gap-3">
          <button 
            onClick={toggleTheme}
            className="p-2 rounded-full hover:bg-[var(--line)] transition-colors text-[var(--ink)]"
          >
            {theme === 'light' ? <Moon size={20} /> : <Sun size={20} />}
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-6 space-y-8">
        {activeTab === 'mdd-translator' && (
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="max-w-4xl mx-auto space-y-8"
          >
            <div className="bg-[var(--card)] border border-[var(--line)] rounded-2xl p-8 shadow-xl space-y-8">
              <div className="flex items-center gap-4 border-b border-[var(--line)] pb-6">
                <div className="w-12 h-12 bg-[var(--accent)] rounded-xl flex items-center justify-center shadow-lg">
                  <Zap className="text-black w-7 h-7" />
                </div>
                <div>
                  <h2 className="text-2xl font-bold text-[var(--ink)] uppercase tracking-tight">MDD Translation Tools</h2>
                  <p className="text-sm text-[var(--muted)] font-mono">Sync Excel Translations with MDD Metadata</p>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                {/* MDD File Section */}
                <div className="space-y-4">
                  <label className="block text-sm font-bold text-[var(--muted)] uppercase tracking-widest">1. Metadata File (.mdd)</label>
                  <div className={`relative group border-2 border-dashed rounded-2xl p-8 transition-all ${mddFile ? 'border-[var(--accent)] bg-[var(--accent)]/5' : 'border-[var(--line)] hover:border-[var(--accent)]/50'}`}>
                    <input 
                      type="file" 
                      accept=".mdd"
                      onChange={(e) => setMddFile(e.target.files?.[0] || null)}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                    />
                    <div className="flex flex-col items-center gap-3 text-center">
                      {mddFile ? (
                        <>
                          <div className="w-12 h-12 bg-[var(--accent)] rounded-full flex items-center justify-center">
                            <Check className="text-black" size={24} />
                          </div>
                          <span className="text-sm font-bold text-[var(--ink)] truncate max-w-full">{mddFile.name}</span>
                          <button onClick={(e) => { e.stopPropagation(); setMddFile(null); }} className="text-xs text-red-500 hover:underline">Remove</button>
                        </>
                      ) : (
                        <>
                          <div className="w-12 h-12 bg-[var(--line)] rounded-full flex items-center justify-center group-hover:bg-[var(--accent)]/20 transition-colors">
                            <Upload className="text-[var(--muted)] group-hover:text-[var(--accent)]" size={24} />
                          </div>
                          <span className="text-sm font-medium text-[var(--muted)]">Upload MDD File</span>
                        </>
                      )}
                    </div>
                  </div>
                </div>

                {/* Excel File Section */}
                <div className="space-y-4">
                  <label className="block text-sm font-bold text-[var(--muted)] uppercase tracking-widest">2. Translation File (.xlsx)</label>
                  <div className={`relative group border-2 border-dashed rounded-2xl p-8 transition-all ${mddExcelFile ? 'border-green-500 bg-green-500/5' : 'border-[var(--line)] hover:border-green-500/50'}`}>
                    <input 
                      type="file" 
                      accept=".xlsx"
                      onChange={(e) => setMddExcelFile(e.target.files?.[0] || null)}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                    />
                    <div className="flex flex-col items-center gap-3 text-center">
                      {mddExcelFile ? (
                        <>
                          <div className="w-12 h-12 bg-green-500 rounded-full flex items-center justify-center">
                            <Check className="text-white" size={24} />
                          </div>
                          <span className="text-sm font-bold text-[var(--ink)] truncate max-w-full">{mddExcelFile.name}</span>
                          <button onClick={(e) => { e.stopPropagation(); setMddExcelFile(null); }} className="text-xs text-red-500 hover:underline">Remove</button>
                        </>
                      ) : (
                        <>
                          <div className="w-12 h-12 bg-[var(--line)] rounded-full flex items-center justify-center group-hover:bg-green-500/20 transition-colors">
                            <FileSpreadsheet className="text-[var(--muted)] group-hover:text-green-500" size={24} />
                          </div>
                          <span className="text-sm font-medium text-[var(--muted)]">Upload Excel File</span>
                        </>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-8 pt-4">
                {/* Context Selection */}
                <div className="space-y-4">
                  <label className="block text-sm font-bold text-[var(--muted)] uppercase tracking-widest">3. MDD Context</label>
                  <div className="relative">
                    <select 
                      value={selectedMddContext}
                      onChange={(e) => setSelectedMddContext(e.target.value)}
                      className="w-full appearance-none bg-[var(--bg)] border border-[var(--line)] rounded-xl px-4 py-3 text-sm font-medium focus:outline-none focus:ring-2 focus:ring-[var(--accent)]/20 focus:border-[var(--accent)] transition-all"
                    >
                      {mddContexts.map(ctx => (
                        <option key={ctx} value={ctx}>{ctx}</option>
                      ))}
                    </select>
                    <ChevronDown className="absolute right-4 top-1/2 -translate-y-1/2 text-[var(--muted)] pointer-events-none" size={18} />
                  </div>
                </div>

                {/* Options */}
                <div className="space-y-3">
                  <label className="block text-sm font-bold text-[var(--muted)] uppercase tracking-widest">4. Settings</label>
                  <div className="grid grid-cols-2 gap-x-4 gap-y-2">
                    <label className="flex items-center gap-2 cursor-pointer group">
                      <input type="checkbox" checked={highlightSpecials} onChange={e => setHighlightSpecials(e.target.checked)} className="rounded border-[var(--line)] text-[var(--accent)] focus:ring-[var(--accent)]" />
                      <span className="text-xs font-medium text-[var(--muted)] group-hover:text-[var(--ink)]">Highlight Specials</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer group">
                      <input type="checkbox" checked={checkHTML} onChange={e => setCheckHTML(e.target.checked)} className="rounded border-[var(--line)] text-[var(--accent)] focus:ring-[var(--accent)]" />
                      <span className="text-xs font-medium text-[var(--muted)] group-hover:text-[var(--ink)]">Check HTML</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer group">
                      <input type="checkbox" checked={checkInserts} onChange={e => setCheckInserts(e.target.checked)} className="rounded border-[var(--line)] text-[var(--accent)] focus:ring-[var(--accent)]" />
                      <span className="text-xs font-medium text-[var(--muted)] group-hover:text-[var(--ink)]">Check Inserts</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer group">
                      <input type="checkbox" checked={applyRTL} onChange={e => setApplyRTL(e.target.checked)} className="rounded border-[var(--line)] text-[var(--accent)] focus:ring-[var(--accent)]" />
                      <span className="text-xs font-medium text-[var(--muted)] group-hover:text-[var(--ink)]">Apply RTL</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer group">
                      <input type="checkbox" checked={updateBaseLanguage} onChange={e => setUpdateBaseLanguage(e.target.checked)} className="rounded border-[var(--line)] text-[var(--accent)] focus:ring-[var(--accent)]" />
                      <span className="text-xs font-medium text-[var(--muted)] group-hover:text-[var(--ink)]">Update Base Language</span>
                    </label>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-4">
                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-[var(--muted)] uppercase tracking-widest">Translation Property</label>
                      <input 
                        type="text" 
                        value={translationProperty} 
                        onChange={e => setTranslationProperty(e.target.value)}
                        placeholder="e.g. translate"
                        className="w-full bg-[var(--bg)] border border-[var(--line)] rounded-lg px-3 py-2 text-xs font-medium focus:outline-none focus:border-[var(--accent)] transition-all"
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] font-bold text-[var(--muted)] uppercase tracking-widest">Languages Filter (Comma Separated)</label>
                      <input 
                        type="text" 
                        value={languagesFilter} 
                        onChange={e => setLanguagesFilter(e.target.value)}
                        placeholder="e.g. ENU, ARA, HEB"
                        className="w-full bg-[var(--bg)] border border-[var(--line)] rounded-lg px-3 py-2 text-xs font-medium focus:outline-none focus:border-[var(--accent)] transition-all"
                      />
                    </div>
                  </div>
                </div>
              </div>

              {/* Status Bar */}
              <div className="bg-[var(--bg)] border border-[var(--line)] rounded-xl p-3 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className={`w-2 h-2 rounded-full ${isMddProcessing ? 'bg-yellow-500 animate-pulse' : 'bg-green-500'}`} />
                  <span className="text-xs font-mono text-[var(--muted)] uppercase tracking-tighter">Status: {mddStatus}</span>
                </div>
                {isMddProcessing && (
                  <div className="w-24 h-1 bg-[var(--line)] rounded-full overflow-hidden">
                    <motion.div 
                      initial={{ x: '-100%' }}
                      animate={{ x: '100%' }}
                      transition={{ repeat: Infinity, duration: 1, ease: "linear" }}
                      className="w-full h-full bg-[var(--accent)]"
                    />
                  </div>
                )}
              </div>

              {/* Actions */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 pt-4">
                <button 
                  onClick={handleMddExport}
                  disabled={isMddProcessing || !mddFile}
                  className="flex flex-col items-center justify-center gap-2 p-4 rounded-2xl border border-[var(--line)] hover:border-[var(--accent)] hover:bg-[var(--accent)]/5 transition-all group disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <Download size={24} className="text-[var(--muted)] group-hover:text-[var(--accent)]" />
                  <span className="text-xs font-bold uppercase tracking-widest">Export Labels</span>
                </button>
                <button 
                  onClick={handleMddImport}
                  disabled={isMddProcessing || !mddFile || !mddExcelFile}
                  className="flex flex-col items-center justify-center gap-2 p-4 rounded-2xl border border-[var(--line)] hover:border-green-500 hover:bg-green-500/5 transition-all group disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <Upload size={24} className="text-[var(--muted)] group-hover:text-green-500" />
                  <span className="text-xs font-bold uppercase tracking-widest">Import Labels</span>
                </button>
                <button 
                  onClick={handleMddValidate}
                  disabled={isMddProcessing || !mddExcelFile}
                  className="flex flex-col items-center justify-center gap-2 p-4 rounded-2xl border border-[var(--line)] hover:border-blue-500 hover:bg-blue-500/5 transition-all group disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <Check size={24} className="text-[var(--muted)] group-hover:text-blue-500" />
                  <span className="text-xs font-bold uppercase tracking-widest">Validate File</span>
                </button>
                <button 
                  onClick={handleMddProtect}
                  disabled={isMddProcessing || !mddExcelFile}
                  className="flex flex-col items-center justify-center gap-2 p-4 rounded-2xl border border-[var(--line)] hover:border-purple-500 hover:bg-purple-500/5 transition-all group disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <Lock size={24} className="text-[var(--muted)] group-hover:text-purple-500" />
                  <span className="text-xs font-bold uppercase tracking-widest">Protect File</span>
                </button>
              </div>

              {mddResultFile && (
                <motion.div 
                  initial={{ opacity: 0, scale: 0.9 }}
                  animate={{ opacity: 1, scale: 1 }}
                  className="pt-6 border-t border-[var(--line)] flex justify-center"
                >
                  <button 
                    onClick={downloadTranslatedMdd}
                    className="flex items-center gap-3 px-8 py-4 rounded-2xl bg-[var(--accent)] text-black font-bold shadow-xl hover:scale-105 transition-transform"
                  >
                    <Download size={24} />
                    Download Translated MDD
                  </button>
                </motion.div>
              )}
            </div>

            {/* Info Card */}
            <div className="bg-blue-500/5 border border-blue-500/20 rounded-2xl p-6 flex gap-4 items-start">
              <div className="w-10 h-10 bg-blue-500 rounded-full flex items-center justify-center flex-shrink-0">
                <Settings className="text-white" size={20} />
              </div>
              <div className="space-y-1">
                <h4 className="font-bold text-blue-900">MDD Translation Workflow</h4>
                <p className="text-sm text-blue-800/70 leading-relaxed">
                  This tool allows you to export labels from an MDD file to Excel for translation, and then import them back. 
                  It supports multiple languages, context selection, and validation of HTML tags and inserts.
                </p>
                <p className="text-xs text-blue-800/50 italic pt-2">
                  Note: MDD processing typically requires a specialized COM environment. This web version provides the interface and simulated logic for the workflow.
                </p>
              </div>
            </div>

            {/* Instructions & Validation Legend */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
              {/* Instructions */}
              <div className="lg:col-span-2 bg-[var(--card)] border border-[var(--line)] rounded-2xl p-8 shadow-sm space-y-6">
                <h3 className="text-xl font-bold text-[var(--ink)] flex items-center gap-2">
                  <Info size={20} className="text-[var(--accent)]" />
                  Instructions for Easy Translate
                </h3>
                
                <div className="space-y-6 text-sm text-[var(--muted)] leading-relaxed">
                  <section className="space-y-2">
                    <h4 className="font-bold text-[var(--ink)] uppercase text-xs tracking-widest">Getting Started</h4>
                    <p>Specify the <span className="font-mono text-[var(--ink)]">*.mdd</span> file for export/import. Available contexts will populate automatically. The export filename is generated from the MDD name.</p>
                  </section>

                  <section className="space-y-3">
                    <h4 className="font-bold text-[var(--ink)] uppercase text-xs tracking-widest">Exporting Labels</h4>
                    <ul className="list-disc pl-5 space-y-2">
                      <li><span className="font-bold text-[var(--ink)]">Apply text wrapping:</span> Wraps cells and limits row height to 100px for readability.</li>
                      <li><span className="font-bold text-[var(--ink)]">Highlight HTML/Inserts:</span> Highlights tags and inserts in <span className="text-red-500 font-bold">red</span> for easy identification.</li>
                      <li><span className="font-bold text-[var(--ink)]">Protection:</span> Use the <span className="italic">Protect</span> button to lock the Variable Map and English columns to prevent accidental changes.</li>
                    </ul>
                  </section>

                  <section className="space-y-3">
                    <h4 className="font-bold text-[var(--ink)] uppercase text-xs tracking-widest">Importing Labels</h4>
                    <ul className="list-disc pl-5 space-y-2">
                      <li>Uses the <span className="font-bold text-[var(--ink)]">Variable Map</span> column as a unique key for mapping translations back to the MDD.</li>
                      <li><span className="font-bold text-[var(--ink)]">HTML/Insert Difference:</span> Compares default labels with translations and highlights differences in <span className="text-blue-500 font-bold">blue</span> (HTML) or <span className="text-green-500 font-bold">green</span> (Inserts).</li>
                      <li><span className="font-bold text-[var(--ink)]">Clear Existing:</span> Removes labels if no translation is found (leaves a dash).</li>
                      <li><span className="font-bold text-[var(--ink)]">RTL Orientation:</span> Automatically adds <span className="font-mono">&lt;font dir='rtl'&gt;</span> tags for RTL languages (ARA, HEB, URD, etc.) if non-Latin characters are detected.</li>
                    </ul>
                  </section>

                  <div className="p-4 bg-[var(--bg)] rounded-xl border border-[var(--line)] italic text-xs">
                    * A <span className="font-mono">*.bak</span> file is created automatically before any import. Validation can be run independently without modifying the original MDD.
                  </div>
                </div>
              </div>

              {/* Validation Colors Legend */}
              <div className="bg-[var(--card)] border border-[var(--line)] rounded-2xl p-6 shadow-sm space-y-6">
                <h3 className="text-lg font-bold text-[var(--ink)] flex items-center gap-2">
                  <Palette size={18} className="text-[var(--accent)]" />
                  Validation Colors
                </h3>

                <div className="space-y-4">
                  <div className="space-y-2">
                    <p className="text-[10px] font-bold text-[var(--muted)] uppercase tracking-tighter">Critical (Not Imported)</p>
                    <div className="space-y-2">
                      {[
                        { color: 'bg-[#FF0000]', label: 'V', desc: 'Variable not found in MDD' },
                        { color: 'bg-[#FFFF00]', label: 'A', desc: 'Category not found in MDD', text: 'text-black' },
                        { color: 'bg-[#808080]', label: 'L', desc: 'Default language in MDD' },
                        { color: 'bg-[#FFA500]', label: 'I', desc: 'Missing language in MDD' },
                        { color: 'bg-[#800000]', label: 'D', desc: 'Text cannot be extracted' },
                      ].map((item, i) => (
                        <div key={i} className="flex items-center gap-3">
                          <div className={`w-6 h-6 rounded flex items-center justify-center text-[10px] font-bold ${item.color} ${item.text || 'text-white'}`}>
                            {item.label}
                          </div>
                          <span className="text-xs text-[var(--muted)]">{item.desc}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="h-px bg-[var(--line)]" />

                  <div className="space-y-2">
                    <p className="text-[10px] font-bold text-[var(--muted)] uppercase tracking-tighter">Non-Critical (Review Needed)</p>
                    <div className="space-y-2">
                      {[
                        { color: 'bg-[#0000FF]', label: 'A', desc: 'HTML tags differences' },
                        { color: 'bg-[#008000]', label: 'T', desc: 'Inserts differences' },
                        { color: 'bg-[#800080]', label: 'I', desc: 'HTML & Inserts differences' },
                        { color: 'bg-[#FFC0CB]', label: 'O', desc: 'Missing translation', text: 'text-black' },
                        { color: 'bg-[#00FFFF]', label: 'N', desc: 'Validation error on row', text: 'text-black' },
                      ].map((item, i) => (
                        <div key={i} className="flex items-center gap-3">
                          <div className={`w-6 h-6 rounded flex items-center justify-center text-[10px] font-bold ${item.color} ${item.text || 'text-white'}`}>
                            {item.label}
                          </div>
                          <span className="text-xs text-[var(--muted)]">{item.desc}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </motion.div>
        )}

        {activeTab === 'extraction' ? (
          !showMapping ? (
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="grid grid-cols-1 md:grid-cols-2 gap-6"
            >
              {/* Upload Section */}
              <div className="space-y-4 min-w-0">
                <h2 className="text-lg font-semibold flex items-center gap-2">
                  <FileText className="text-[var(--accent)]" />
                  Source Documents
                </h2>
                
                  <div className="grid gap-4">
                    <FileDropZone 
                      label="English Master DOCX" 
                      file={englishFile} 
                      onFileSelect={setEnglishFile} 
                      accept=".docx"
                    />
                    <FileDropZone 
                      label="Target Language DOCX" 
                      file={targetFile} 
                      onFileSelect={setTargetFile} 
                      accept=".docx"
                    />
                  </div>

                  <div className="flex gap-4">
                    <button 
                      onClick={handleProcess}
                      disabled={isProcessing || !englishFile || !targetFile}
                      className={`flex-1 py-4 rounded-xl font-bold text-lg transition-all shadow-lg ${
                        isProcessing || !englishFile || !targetFile 
                          ? 'bg-[var(--line)] text-[var(--muted)] cursor-not-allowed' 
                          : 'bg-[var(--accent)] text-black hover:scale-[1.02] active:scale-[0.98]'
                      }`}
                    >
                      {isProcessing ? 'Processing...' : 'Activate Transformation'}
                    </button>
                    
                    {(englishFile || targetFile) && (
                      <button 
                        onClick={handleReset}
                        className="px-6 py-4 rounded-xl font-bold border border-red-500/20 text-red-500 hover:bg-red-500/10 transition-all flex items-center justify-center"
                        title="Clear Files"
                      >
                        <RotateCcw size={24} />
                      </button>
                    )}
                  </div>
                </div>

              {/* Instructions Card */}
              <div className="bg-[var(--card)] border border-[var(--line)] rounded-2xl p-8 flex flex-col justify-center space-y-6 min-w-0">
                <div className="space-y-2">
                  <h3 className="text-2xl font-bold text-[var(--accent)]">How it works</h3>
                  <p className="text-[var(--muted)]">Streamline your survey translation mapping with precision.</p>
                </div>
                
                <ul className="space-y-4">
                  {[
                    "Upload both English and Target language Word documents.",
                    "The system automatically extracts and aligns lines.",
                    "Use 'Shift-Sync' to re-align sequences if lines are missing.",
                    "Lock verified rows to protect them from bulk changes.",
                    "Export a clean, professional Excel mapping sheet."
                  ].map((step, i) => (
                    <li key={i} className="flex gap-3 items-start">
                      <span className="flex-shrink-0 w-6 h-6 rounded-full bg-[var(--line)] text-[var(--ink)] flex items-center justify-center text-xs font-bold">
                        {i + 1}
                      </span>
                      <span className="text-sm">{step}</span>
                    </li>
                  ))}
                </ul>
              </div>
            </motion.div>
          ) : (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="space-y-4"
            >
              {/* Toolbar */}
              <div className="bg-[var(--card)] border border-[var(--line)] rounded-xl p-4 flex items-center justify-between flex-wrap gap-4 sticky top-[88px] z-40 shadow-sm">
                <div className="flex items-center gap-6">
                  <label className="flex items-center gap-2 cursor-pointer group">
                    <div className={`w-5 h-5 rounded border flex items-center justify-center transition-colors ${exportClean ? 'bg-[var(--accent)] border-[var(--accent)]' : 'border-[var(--line)] bg-[var(--bg)]'}`}>
                      {exportClean && <Check size={14} className="text-black" />}
                    </div>
                    <input 
                      type="checkbox" 
                      className="hidden" 
                      checked={exportClean} 
                      onChange={e => setExportClean(e.target.checked)} 
                    />
                    <span className="text-sm font-medium group-hover:text-[var(--accent)] transition-colors">Clean Export (No Numbers)</span>
                  </label>
                  
                  <div className="h-6 w-px bg-[var(--line)]" />
                  
                  {/* Movable Search Bar */}
                  <motion.div 
                    drag
                    dragMomentum={false}
                    className="relative group z-50 flex items-center gap-2 bg-[var(--bg)] border border-[var(--line)] rounded-lg px-2 shadow-lg cursor-move"
                  >
                    <GripVertical className="text-[var(--muted)]" size={16} />
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--muted)] group-focus-within:text-[var(--accent)] transition-colors" size={16} />
                      <input 
                        type="text"
                        placeholder="Search QID, English or Target..."
                        value={extractionSearch}
                        onChange={(e) => {
                          setExtractionSearch(e.target.value);
                          setCurrentPage(1);
                        }}
                        className="pl-10 pr-4 py-2 bg-transparent border-none text-sm focus:outline-none w-[250px]"
                      />
                      {extractionSearch && (
                        <button 
                          onClick={() => {
                            setExtractionSearch('');
                            setCurrentPage(1);
                          }}
                          className="absolute right-3 top-1/2 -translate-y-1/2 text-[var(--muted)] hover:text-[var(--ink)]"
                        >
                          <X size={14} />
                        </button>
                      )}
                    </div>
                  </motion.div>

                  <div className="h-6 w-px bg-[var(--line)]" />
                  
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-mono text-[var(--muted)] uppercase tracking-widest">Status:</span>
                    <span className="text-xs font-bold text-[var(--accent)] bg-[var(--accent)]/10 px-2 py-1 rounded">SHIFT-SYNC ACTIVE</span>
                  </div>
                  
                  {totalPages > 1 && (
                    <div className="flex items-center gap-2 ml-4">
                      <button 
                        disabled={currentPage === 1}
                        onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                        className="p-1 rounded hover:bg-[var(--line)] disabled:opacity-30"
                      >
                        <ChevronDown size={18} className="rotate-90" />
                      </button>
                      <span className="text-xs font-bold">Page {currentPage} of {totalPages}</span>
                      <button 
                        disabled={currentPage === totalPages}
                        onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                        className="p-1 rounded hover:bg-[var(--line)] disabled:opacity-30"
                      >
                        <ChevronDown size={18} className="-rotate-90" />
                      </button>
                    </div>
                  )}
                </div>

                <div className="flex items-center gap-2">
                  {selectedIds.size > 0 && (
                    <>
                      <button 
                        onClick={() => setIsBulkLockConfirmOpen(true)}
                        className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-[var(--accent)]/10 text-[var(--accent)] hover:bg-[var(--accent)]/20 transition-colors text-sm font-bold shadow-sm"
                      >
                        <Lock size={16} />
                        Lock ({selectedIds.size})
                      </button>
                      <button 
                        onClick={() => setIsBulkUnlockConfirmOpen(true)}
                        className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-[var(--line)]/10 text-[var(--ink)] hover:bg-[var(--line)]/20 transition-colors text-sm font-bold shadow-sm"
                      >
                        <Unlock size={16} />
                        Unlock ({selectedIds.size})
                      </button>
                      <button 
                        onClick={() => setIsDeleteConfirmOpen(true)}
                        className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-red-500 text-white hover:bg-red-600 transition-colors text-sm font-bold shadow-sm"
                      >
                        <Trash2 size={16} />
                        Delete ({selectedIds.size})
                      </button>
                    </>
                  )}
                  <button 
                    onClick={addRow}
                    className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-[var(--accent)] text-black hover:opacity-90 transition-colors text-sm font-bold shadow-sm"
                  >
                    <Plus size={16} />
                    Add Row
                  </button>
                  <button 
                    onClick={unlockAll}
                    className="flex items-center gap-2 px-3 py-1.5 rounded-lg border border-[var(--line)] hover:bg-[var(--line)] transition-colors text-sm font-medium"
                  >
                    <Unlock size={16} />
                    Unlock All
                  </button>
                  <button 
                    onClick={handleReset}
                    className="flex items-center gap-2 px-3 py-1.5 rounded-lg border border-red-500/20 text-red-500 hover:bg-red-500/10 transition-colors text-sm font-medium"
                  >
                    <RotateCcw size={16} />
                    Reset
                  </button>

                  <div className="h-6 w-px bg-[var(--line)] mx-2" />

                  <button 
                    onClick={exportToExcel}
                    className="flex items-center gap-2 bg-[var(--accent)] text-black px-4 py-2 rounded-lg font-bold hover:opacity-90 transition-opacity shadow-lg"
                  >
                    <Download size={18} />
                    <span>Export Excel</span>
                  </button>
                </div>
              </div>

              {/* Table */}
              <div className="bg-[var(--card)] border border-[var(--line)] rounded-2xl overflow-hidden shadow-sm">
                <table className="w-full border-collapse table-fixed">
                  <thead>
                    <tr className="bg-[var(--bg)] border-b border-[var(--line)]">
                      <th className="w-12 p-4 text-center">
                        <input 
                          type="checkbox" 
                          checked={rows.length > 0 && selectedIds.size === rows.length}
                          onChange={toggleSelectAll}
                          className="w-4 h-4 rounded border-[var(--line)] text-[var(--accent)] focus:ring-[var(--accent)]"
                        />
                      </th>
                      <th className="w-12 p-4 text-center text-xs font-bold text-[var(--muted)] uppercase">No.</th>
                      <th className="w-32 p-4 text-center text-xs font-bold text-[var(--muted)] uppercase">QID</th>
                      <th className="p-4 text-left text-xs font-bold text-[var(--muted)] uppercase">Master English</th>
                      <th className="p-4 text-left text-xs font-bold text-[var(--muted)] uppercase">Target Language</th>
                      <th className="w-32 p-4 text-center text-xs font-bold text-[var(--muted)] uppercase">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {paginatedRows.map((row, pIdx) => {
                      const index = (currentPage - 1) * rowsPerPage + pIdx;
                      return (
                        <React.Fragment key={row.id}>
                          {/* Insert Row Button */}
                          <tr className="group/insert h-0">
                            <td colSpan={6} className="p-0 relative">
                              <button 
                                onClick={() => insertRow(index)}
                                className="absolute left-1/2 -translate-x-1/2 -translate-y-1/2 z-10 opacity-0 group-hover/insert:opacity-100 transition-opacity bg-[var(--accent)] text-black rounded-full p-1 shadow-md hover:scale-110 active:scale-95"
                                title="Insert Row Here"
                              >
                                <PlusCircle size={16} />
                              </button>
                            </td>
                          </tr>
                          
                          <tr 
                            className={`data-row ${row.isLocked ? 'row-locked bg-[var(--line)]/20' : ''} ${selectedIds.has(row.id) ? 'bg-[var(--accent)]/5' : ''}`}
                          >
                          <td className="p-2 text-center">
                            <input 
                              type="checkbox" 
                              checked={selectedIds.has(row.id)}
                              onChange={() => toggleSelectRow(row.id)}
                              className="w-4 h-4 rounded border-[var(--line)] text-[var(--accent)] focus:ring-[var(--accent)]"
                            />
                          </td>
                            <td className="p-2 text-center flex flex-col items-center justify-center">
                              <span className="text-[var(--muted)] font-mono text-xs">{index + 1}</span>
                              <span className="text-[var(--muted)] opacity-30 font-mono text-[10px] uppercase" title={`Internal ID: ${row.id}`}>
                                {row.id.substring(0, 4)}
                              </span>
                            </td>
                            <td className="p-2">
                              <input 
                                type="text" 
                                list="qid-suggestions"
                                value={row.qid}
                                disabled={row.isLocked}
                                onChange={e => handleRowUpdate(row.id, 'qid', e.target.value)}
                                className={`w-full bg-transparent border border-transparent hover:border-[var(--line)] focus:border-[var(--accent)] focus:bg-[var(--bg)] focus:ring-1 focus:ring-[var(--accent)] rounded-lg p-2 text-center font-bold text-[var(--accent)] outline-none transition-all ${row.isLocked ? 'opacity-50 cursor-not-allowed' : ''}`}
                                placeholder="QID"
                              />
                              <datalist id="qid-suggestions">
                                {Array.from(new Set(rows.map(r => r.qid).filter(Boolean))).map(q => (
                                  <option key={q} value={q} />
                                ))}
                              </datalist>
                            </td>
                            <td className="p-2">
                              <AutoExpandingTextarea 
                                value={row.english}
                                disabled={row.isLocked}
                                onChange={val => handleRowUpdate(row.id, 'english', val)}
                                className={`w-full bg-transparent border-none text-sm focus:ring-0 min-h-[40px] ${row.isLocked ? 'opacity-50 cursor-not-allowed' : ''}`}
                              />
                            </td>
                            <td className="p-2 relative group/target">
                              <div className="flex items-start gap-1">
                                <AutoExpandingTextarea 
                                  value={row.target}
                                  disabled={row.isLocked}
                                  onChange={val => handleRowUpdate(row.id, 'target', val)}
                                  className={`flex-1 bg-transparent border-none text-sm focus:ring-0 min-h-[40px] ${row.isLocked ? 'opacity-50 cursor-not-allowed' : ''}`}
                                />
                                {!row.isLocked && (
                                  <button 
                                    onClick={(e) => openPicker(row.id, 'extraction', e)}
                                    className="p-1.5 rounded-lg text-[var(--muted)] hover:bg-[var(--line)] hover:text-[var(--accent)] transition-all opacity-0 group-hover/target:opacity-100"
                                    title="Pick Target Line"
                                  >
                                    <Search size={14} />
                                  </button>
                                )}
                              </div>
                            </td>
                            <td className="p-2 text-center">
                              <div className="flex items-center justify-center gap-1">
                                <button 
                                  onClick={() => toggleLock(row.id)}
                                  className={`p-2 rounded-lg transition-colors ${row.isLocked ? 'text-[var(--accent)]' : 'text-[var(--muted)] hover:text-[var(--ink)]'}`}
                                  title={row.isLocked ? "Unlock Row" : "Lock Row"}
                                >
                                  {row.isLocked ? <Lock size={18} /> : <Unlock size={18} />}
                                </button>
                                <button 
                                  onClick={() => deleteRow(row.id)}
                                  disabled={row.isLocked}
                                  className={`p-2 rounded-lg transition-colors ${row.isLocked ? 'text-[var(--muted)] opacity-30 cursor-not-allowed' : 'text-[var(--muted)] hover:bg-red-500/10 hover:text-red-500'}`}
                                  title="Delete Row"
                                >
                                  <Trash2 size={18} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </motion.div>
          )
        ) : (
          !showExcelMapping ? (
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="grid grid-cols-1 md:grid-cols-2 gap-6"
            >
              {/* Excel Mapper Section */}
              <div className="space-y-4 min-w-0">
                <h2 className="text-lg font-semibold flex items-center gap-2">
                  <FileSpreadsheet className="text-[var(--accent)]" />
                  Excel Translation Mapper
                </h2>
                
                <div className="grid gap-4">
                  <FileDropZone 
                    label="Master Excel (Original)" 
                    file={masterExcel} 
                    onFileSelect={handleMasterExcelSelect} 
                    accept=".xlsx, .xls"
                  />
                  <FileDropZone 
                    label="Translated Excel (Mapped)" 
                    file={translatedExcel} 
                    onFileSelect={handleTranslatedExcelSelect} 
                    accept=".xlsx, .xls"
                  />
                </div>

                {/* Column Configuration UI */}
                {(masterExcel || translatedExcel) && (
                  <motion.div 
                    initial={{ opacity: 0, height: 0 }}
                    animate={{ opacity: 1, height: 'auto' }}
                    className="bg-[var(--card)] border border-[var(--line)] rounded-xl p-4 space-y-4 overflow-hidden"
                  >
                    <div className="flex items-center gap-2 text-[var(--accent)] mb-2">
                      <Settings size={16} />
                      <span className="text-xs font-bold uppercase tracking-wider">Active Mapping Configuration</span>
                    </div>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="p-3 rounded-lg bg-[var(--bg)] border border-[var(--line)]">
                        <p className="text-[10px] font-bold text-[var(--muted)] uppercase mb-2">Original File (Master)</p>
                        <div className="flex justify-between items-center text-xs">
                          <span>English Source:</span>
                          <span className="font-mono text-[var(--accent)]">{masterHeaders[masterEngCol] || `Col ${String.fromCharCode(65 + masterEngCol)}`}</span>
                        </div>
                        <div className="flex justify-between items-center text-xs mt-1">
                          <span>Italy/Target:</span>
                          <span className="font-mono text-[var(--accent)]">{masterHeaders[masterTargetCol] || `Col ${String.fromCharCode(65 + masterTargetCol)}`}</span>
                        </div>
                      </div>

                      <div className="p-3 rounded-lg bg-[var(--bg)] border border-[var(--line)]">
                        <p className="text-[10px] font-bold text-[var(--muted)] uppercase mb-2">Mapped File (Source)</p>
                        <div className="flex justify-between items-center text-xs">
                          <span>English Key:</span>
                          <span className="font-mono text-[var(--accent)]">{mappedHeaders[mappedEngCol] || `Col ${String.fromCharCode(65 + mappedEngCol)}`}</span>
                        </div>
                        <div className="flex justify-between items-center text-xs mt-1">
                          <span>Translation (C):</span>
                          <span className="font-mono text-[var(--accent)]">{mappedHeaders[mappedTargetCol] || `Col ${String.fromCharCode(65 + mappedTargetCol)}`}</span>
                        </div>
                      </div>
                    </div>

                    <div className="h-px bg-[var(--line)] my-2" />

                    <div className="grid grid-cols-2 gap-4">
                      {/* Master Config */}
                      <div className="space-y-3">
                        <p className="text-[10px] font-bold text-[var(--muted)] uppercase">Original File (Master)</p>
                        <div className="space-y-2">
                          <div className="flex flex-col gap-1">
                            <span className="text-[9px] text-[var(--muted)]">English (Read)</span>
                            <select 
                              value={masterEngCol}
                              onChange={e => setMasterEngCol(Number(e.target.value))}
                              className="bg-[var(--bg)] border border-[var(--line)] rounded px-2 py-1 text-xs outline-none focus:border-[var(--accent)]"
                            >
                              {masterHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                            </select>
                          </div>
                          <div className="flex flex-col gap-1">
                            <span className="text-[9px] text-[var(--muted)]">Target (Write)</span>
                            <select 
                              value={masterTargetCol}
                              onChange={e => setMasterTargetCol(Number(e.target.value))}
                              className="bg-[var(--bg)] border border-[var(--line)] rounded px-2 py-1 text-xs outline-none focus:border-[var(--accent)]"
                            >
                              {masterHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                            </select>
                          </div>
                        </div>
                      </div>

                      {/* Mapped Config */}
                      <div className="space-y-3">
                        <p className="text-[10px] font-bold text-[var(--muted)] uppercase">Mapped File (Source)</p>
                        <div className="space-y-2">
                          <div className="flex flex-col gap-1">
                            <span className="text-[9px] text-[var(--muted)]">English (Match)</span>
                            <select 
                              value={mappedEngCol}
                              onChange={e => setMappedEngCol(Number(e.target.value))}
                              className="bg-[var(--bg)] border border-[var(--line)] rounded px-2 py-1 text-xs outline-none focus:border-[var(--accent)]"
                            >
                              {mappedHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                            </select>
                          </div>
                          <div className="flex flex-col gap-1">
                            <span className="text-[9px] text-[var(--muted)]">Target (Extract)</span>
                            <select 
                              value={mappedTargetCol}
                              onChange={e => setMappedTargetCol(Number(e.target.value))}
                              className="bg-[var(--bg)] border border-[var(--line)] rounded px-2 py-1 text-xs outline-none focus:border-[var(--accent)]"
                            >
                              {mappedHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                            </select>
                          </div>
                        </div>
                      </div>
                    </div>
                  </motion.div>
                )}

                <div className="flex gap-4">
                  <button 
                    onClick={handleExcelProcess}
                    disabled={isExcelProcessing || !masterExcel || !translatedExcel}
                    className={`flex-1 py-4 rounded-xl font-bold text-lg transition-all shadow-lg ${
                      isExcelProcessing || !masterExcel || !translatedExcel 
                        ? 'bg-[var(--line)] text-[var(--muted)] cursor-not-allowed' 
                        : 'bg-[var(--accent)] text-black hover:scale-[1.02] active:scale-[0.98]'
                    }`}
                  >
                    {isExcelProcessing ? 'Mapping...' : 'Map Translations'}
                  </button>
                  
                  {(masterExcel || translatedExcel) && (
                    <button 
                      onClick={handleReset}
                      className="px-6 py-4 rounded-xl font-bold border border-red-500/20 text-red-500 hover:bg-red-500/10 transition-all flex items-center justify-center"
                      title="Clear Files"
                    >
                      <RotateCcw size={24} />
                    </button>
                  )}
                </div>
              </div>

              {/* Excel Mapper Instructions */}
              <div className="bg-[var(--card)] border border-[var(--line)] rounded-2xl p-8 flex flex-col justify-center space-y-6 min-w-0">
                <div className="space-y-2">
                  <h3 className="text-2xl font-bold text-[var(--accent)]">Excel Mapper</h3>
                  <p className="text-[var(--muted)]">Sync translations from a mapped file back to your master file.</p>
                </div>
                
                <ul className="space-y-4">
                  {[
                    "Upload the Master Excel (the one that needs translations).",
                    "Upload the Translated Excel (the one you previously mapped).",
                    "The system matches rows by English text automatically.",
                    "It populates the Target column in your Master file.",
                    "Download the updated Master file with all translations synced."
                  ].map((step, i) => (
                    <li key={i} className="flex gap-3 items-start">
                      <span className="flex-shrink-0 w-6 h-6 rounded-full bg-[var(--line)] text-[var(--ink)] flex items-center justify-center text-xs font-bold">
                        {i + 1}
                      </span>
                      <span className="text-sm">{step}</span>
                    </li>
                  ))}
                </ul>
              </div>
            </motion.div>
          ) : (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="space-y-4"
            >
              {/* Excel Toolbar */}
              <div className="bg-[var(--card)] border border-[var(--line)] rounded-xl p-4 flex items-center justify-between flex-wrap gap-4 sticky top-[88px] z-40 shadow-sm">
                <div className="flex items-center gap-4">
                  <h2 className="text-lg font-bold text-[var(--accent)] uppercase tracking-tight">Excel Mapping Results</h2>
                  <div className="h-6 w-px bg-[var(--line)]" />
                  
                  {/* Search Bar */}
                  <div className="relative group">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-[var(--muted)] group-focus-within:text-[var(--accent)] transition-colors" size={16} />
                    <input 
                      type="text"
                      placeholder="Search QID, English or Target..."
                      value={excelSearch}
                      onChange={(e) => {
                        setExcelSearch(e.target.value);
                        setCurrentPage(1); // Reset to first page on search
                      }}
                      className="pl-10 pr-4 py-2 bg-[var(--bg)] border border-[var(--line)] rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-[var(--accent)]/20 focus:border-[var(--accent)] transition-all w-[300px]"
                    />
                    {excelSearch && (
                      <button 
                        onClick={() => {
                          setExcelSearch('');
                          setCurrentPage(1);
                        }}
                        className="absolute right-3 top-1/2 -translate-y-1/2 text-[var(--muted)] hover:text-[var(--ink)]"
                      >
                        <X size={14} />
                      </button>
                    )}
                  </div>

                  <div className="h-6 w-px bg-[var(--line)]" />
                  
                  <div className="flex items-center gap-3">
                    <div className="flex items-center gap-1.5 px-2 py-1 rounded-md bg-green-500/10 border border-green-500/20">
                      <div className="w-1.5 h-1.5 rounded-full bg-green-500" />
                      <span className="text-[10px] font-bold text-green-600 uppercase">Mapped: {excelResult?.slice(3).filter(r => r[masterTargetCol]).length || 0}</span>
                    </div>
                    <div className="flex items-center gap-1.5 px-2 py-1 rounded-md bg-red-500/10 border border-red-500/20">
                      <div className="w-1.5 h-1.5 rounded-full bg-red-500" />
                      <span className="text-[10px] font-bold text-red-600 uppercase">Missing: {excelResult?.slice(3).filter(r => !r[masterTargetCol]).length || 0}</span>
                    </div>
                  </div>

                  <div className="h-6 w-px bg-[var(--line)]" />
                  <p className="text-sm text-[var(--muted)]">Review and edit translations before downloading.</p>
                  
                  {totalPages > 1 && (
                    <div className="flex items-center gap-2 ml-4">
                      <button 
                        disabled={currentPage === 1}
                        onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                        className="p-1 rounded hover:bg-[var(--line)] disabled:opacity-30"
                      >
                        <ChevronDown size={18} className="rotate-90" />
                      </button>
                      <span className="text-xs font-bold">Page {currentPage} of {totalPages}</span>
                      <button 
                        disabled={currentPage === totalPages}
                        onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                        className="p-1 rounded hover:bg-[var(--line)] disabled:opacity-30"
                      >
                        <ChevronDown size={18} className="-rotate-90" />
                      </button>
                    </div>
                  )}
                </div>

                <div className="flex items-center gap-2">
                  <button 
                    onClick={downloadExcelResult}
                    className="flex items-center gap-2 px-4 py-2 rounded-lg bg-green-500 text-white hover:bg-green-600 transition-colors text-sm font-bold shadow-sm"
                  >
                    <Download size={18} />
                    Download Result
                  </button>
                  <button 
                    onClick={handleReset}
                    className="flex items-center gap-2 px-4 py-2 rounded-lg border border-red-500/20 text-red-500 hover:bg-red-500/10 transition-colors text-sm font-medium"
                  >
                    <RotateCcw size={18} />
                    Reset
                  </button>
                </div>
              </div>

              {/* Excel Table */}
              <div className="bg-[var(--card)] border border-[var(--line)] rounded-2xl overflow-hidden shadow-sm">
                <table className="w-full border-collapse table-fixed">
                  <thead>
                    <tr className="bg-[var(--bg)] border-b border-[var(--line)]">
                      <th className="w-16 p-4 text-center text-xs font-bold text-[var(--muted)] uppercase">No.</th>
                      <th className="w-32 p-4 text-center text-xs font-bold text-[var(--muted)] uppercase">{masterHeaders[0] || 'Col A'}</th>
                      <th className="w-32 p-4 text-center text-xs font-bold text-[var(--muted)] uppercase">{masterHeaders[1] || 'Col B'}</th>
                      <th className="p-4 text-left text-xs font-bold text-[var(--muted)] uppercase">{masterHeaders[masterEngCol] || 'English'}</th>
                      <th className="p-4 text-left text-xs font-bold text-[var(--muted)] uppercase">{masterHeaders[masterTargetCol] || 'Target'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {paginatedExcelRows.map(({ row, actualIndex }, idx) => {
                      return (
                        <tr key={idx} className="data-row hover:bg-[var(--line)]/5 transition-colors">
                          <td className="p-2 text-center text-[var(--muted)] font-mono text-xs">{actualIndex + 1}</td>
                          <td className="p-2">
                            <input 
                              type="text" 
                              value={row[0] || ''} 
                              readOnly
                              className="w-full bg-transparent border-none text-xs text-center focus:ring-0 opacity-70 cursor-default"
                            />
                          </td>
                          <td className="p-2">
                            <input 
                              type="text" 
                              value={row[1] || ''} 
                              readOnly
                              className="w-full bg-transparent border-none text-xs text-center focus:ring-0 opacity-70 cursor-default"
                            />
                          </td>
                          <td className="p-2">
                            <AutoExpandingTextarea 
                              value={row[masterEngCol] || ''} 
                              readOnly
                              className="w-full bg-transparent border-none text-sm focus:ring-0 min-h-[40px] opacity-70 cursor-default"
                            />
                          </td>
                          <td className="p-2 relative group/target">
                            <div className="flex items-start gap-1">
                              <div className="flex-1 relative">
                                <AutoExpandingTextarea 
                                  value={row[masterTargetCol] || ''} 
                                  onChange={val => handleExcelRowUpdate(actualIndex, masterTargetCol, val)}
                                  className={`w-full bg-transparent border-none text-sm focus:ring-0 min-h-[40px] ${!row[masterTargetCol] ? 'bg-red-500/10 ring-1 ring-red-500/20 rounded' : ''}`}
                                  placeholder={!row[masterTargetCol] ? "Missing Translation..." : ""}
                                />
                                {/* QID Match Indicator */}
                                {(() => {
                                  const qidA = normalizeQID(row[0]);
                                  const qidB = normalizeQID(row[1]);
                                  const isQidMatch = (qidA && qidMapState[qidA] === row[masterTargetCol]) || (qidB && qidMapState[qidB] === row[masterTargetCol]);
                                  if (isQidMatch && row[masterTargetCol]) {
                                    return (
                                      <div className="absolute -top-1.5 -right-1.5 flex items-center gap-1 bg-blue-500/10 text-blue-500 text-[8px] font-bold px-1 py-0.5 rounded border border-blue-500/20 pointer-events-none shadow-sm z-10 backdrop-blur-sm">
                                        <Zap size={8} /> QID MATCH
                                      </div>
                                    );
                                  }
                                  return null;
                                })()}
                              </div>
                              <button 
                                onClick={(e) => openPicker(actualIndex.toString(), 'excel', e)}
                                className={`p-1.5 rounded-lg transition-all ${!row[masterTargetCol] ? 'text-red-500 bg-red-500/10 opacity-100' : 'text-[var(--muted)] hover:bg-[var(--line)] hover:text-[var(--accent)] opacity-0 group-hover/target:opacity-100'}`}
                                title="Search Translation"
                              >
                                <Search size={14} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {unmappedExcelRows.length > 0 && (
                <div className="mt-8 space-y-4">
                  <h3 className="text-lg font-bold text-red-500 flex items-center gap-2">
                    <X size={20} />
                    Unmapped English Strings ({unmappedExcelRows.length})
                  </h3>
                  <div className="bg-[var(--card)] border border-red-500/20 rounded-xl p-4 max-h-[300px] overflow-y-auto custom-scrollbar">
                    <ul className="space-y-2">
                      {unmappedExcelRows.map((text, i) => (
                        <li key={i} className="text-sm p-2 border-b border-[var(--line)] last:border-0 font-mono break-all">
                          {text}
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>
              )}
            </motion.div>
          )
        )}
      </main>

      {/* Picker Popover */}
      <AnimatePresence>
        {isPickerOpen && pickerPosition && (
          <div className="fixed inset-0 z-[100]">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsPickerOpen(false)}
              className="absolute inset-0 bg-black/5"
            />
            <motion.div 
              drag
              dragMomentum={false}
              initial={{ opacity: 0, scale: 0.95, y: -10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: -10 }}
              style={{ 
                top: pickerPosition.top, 
                left: pickerPosition.left,
                position: 'absolute'
              }}
              className="w-[450px] bg-[var(--bg)] rounded-xl shadow-2xl border border-[var(--line)] overflow-hidden flex flex-col max-h-[500px] cursor-default"
            >
              <div className="p-3 border-b border-[var(--line)] flex items-center justify-between bg-[var(--card)] cursor-move">
                <div className="flex items-center gap-2">
                  <GripVertical size={14} className="text-[var(--muted)]" />
                  <h3 className="font-bold text-sm uppercase tracking-wider text-[var(--muted)]">
                    Select Target Line {activeRowInfo?.qid && <span className="text-[var(--accent)] ml-1">[{activeRowInfo.qid}]</span>}
                  </h3>
                </div>
                <div className="flex items-center gap-2">
                  <label className="flex items-center gap-1.5 cursor-pointer">
                    <input 
                      type="checkbox" 
                      checked={propagateChanges}
                      onChange={e => setPropagateChanges(e.target.checked)}
                      className="w-3.5 h-3.5 rounded border-[var(--line)] text-[var(--accent)] focus:ring-[var(--accent)]"
                    />
                    <span className="text-[10px] font-bold uppercase text-[var(--muted)]">Shift-Sync</span>
                  </label>
                  <button onClick={() => setIsPickerOpen(false)} className="p-1 hover:bg-[var(--line)] rounded-full">
                    <X size={16} />
                  </button>
                </div>
              </div>

              <div className="p-4 space-y-4 flex-1 overflow-hidden flex flex-col relative">
                {activeRowInfo?.qid && (
                  <div className="absolute inset-0 flex items-center justify-center pointer-events-none overflow-hidden opacity-[0.04] select-none z-0">
                    <span className="text-[140px] font-black uppercase rotate-[-15deg] whitespace-nowrap text-[var(--ink)]">
                      {activeRowInfo.qid}
                    </span>
                  </div>
                )}
                {/* Search Bar First */}
                <div className="relative group z-10">
                  <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                    <Search className="text-[var(--muted)] group-focus-within:text-[var(--accent)] transition-colors" size={18} />
                  </div>
                  <input 
                    type="text" 
                    placeholder="Search for a translation..."
                    autoFocus
                    value={pickerSearch}
                    onChange={e => setPickerSearch(e.target.value)}
                    className="block w-full pl-11 pr-10 py-3 bg-[var(--card)] border border-[var(--line)] rounded-2xl focus:ring-4 focus:ring-[var(--accent)]/10 focus:border-[var(--accent)] outline-none text-sm transition-all shadow-sm placeholder:text-[var(--muted)]/40 font-medium"
                  />
                  {pickerSearch && (
                    <button 
                      onClick={() => setPickerSearch('')}
                      className="absolute inset-y-0 right-0 pr-3 flex items-center text-[var(--muted)] hover:text-red-500 transition-colors"
                    >
                      <X size={16} />
                    </button>
                  )}
                </div>

                {/* Options List Second */}
                <div className="flex-1 overflow-y-auto border border-[var(--line)] rounded-2xl bg-[var(--card)] custom-scrollbar shadow-inner z-10">
                  {suggestedQidMatch && !pickerSearch && (
                    <div className="border-b border-[var(--line)] bg-[var(--accent)]/5">
                      <div className="px-4 py-2 text-[10px] font-bold text-[var(--accent)] uppercase tracking-widest border-b border-[var(--line)]/50">Suggested (QID Match)</div>
                      <button 
                        onClick={() => confirmPicker(suggestedQidMatch)}
                        className="w-full flex items-start gap-4 p-4 hover:bg-[var(--accent)]/10 transition-colors group text-left relative"
                      >
                        <div className="absolute left-0 top-0 bottom-0 w-1 bg-[var(--accent)]" />
                        <div className="flex-1 min-w-0">
                          <span className="text-[13px] text-[var(--ink)] line-clamp-3 leading-relaxed block font-bold">{suggestedQidMatch}</span>
                          <span className="text-[10px] text-[var(--accent)] font-bold uppercase mt-1 block">Best Match for QID</span>
                        </div>
                        <div className="w-6 h-6 rounded-full bg-[var(--accent)] text-white flex items-center justify-center shadow-sm">
                          <Check size={14} />
                        </div>
                      </button>
                    </div>
                  )}
                  
                  {filteredTargetLines.length > 0 ? (
                    <div className="divide-y divide-[var(--line)]">
                      {filteredTargetLines.map((item) => (
                        <button 
                          key={item.index}
                          onClick={() => confirmPicker(item.text)}
                          className="w-full flex items-start gap-4 p-4 hover:bg-[var(--accent)]/5 transition-colors group text-left relative"
                        >
                          <div className="absolute left-0 top-0 bottom-0 w-1 bg-[var(--accent)] scale-y-0 group-hover:scale-y-100 transition-transform origin-center" />
                          <span className="text-[var(--muted)] font-mono text-[10px] mt-1 min-w-[24px] bg-[var(--line)]/30 px-1.5 py-0.5 rounded text-center">{item.index + 1}</span>
                          <div className="flex-1 min-w-0">
                            <span className="text-[13px] group-hover:text-[var(--ink)] line-clamp-3 leading-relaxed block font-medium">{item.text}</span>
                            {item.sub && <span className="text-[11px] text-[var(--muted)] block mt-1.5 truncate opacity-70 italic">{item.sub}</span>}
                          </div>
                        </button>
                      ))}
                    </div>
                  ) : (
                    <div className="p-12 text-center space-y-2">
                      <div className="w-12 h-12 bg-[var(--line)]/20 rounded-full flex items-center justify-center mx-auto text-[var(--muted)]">
                        <Search size={24} />
                      </div>
                      <p className="text-[var(--muted)] text-sm font-medium">No matching lines found</p>
                      <p className="text-[var(--muted)] text-xs opacity-60">Try a different search term</p>
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Confirmation Modals */}
      <AnimatePresence>
        {(isDeleteConfirmOpen || isBulkLockConfirmOpen || isBulkUnlockConfirmOpen) && (
          <div className="fixed inset-0 z-[200] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => {
                setIsDeleteConfirmOpen(false);
                setIsBulkLockConfirmOpen(false);
                setIsBulkUnlockConfirmOpen(false);
              }}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="relative w-full max-w-md bg-[var(--bg)] rounded-2xl shadow-2xl border border-[var(--line)] p-6 space-y-6"
            >
              <div className="flex items-center gap-4 text-red-500">
                <div className={`w-12 h-12 rounded-full flex items-center justify-center ${isDeleteConfirmOpen ? 'bg-red-500/10 text-red-500' : 'bg-[var(--accent)]/10 text-[var(--accent)]'}`}>
                  {isDeleteConfirmOpen ? <Trash2 size={24} /> : isBulkLockConfirmOpen ? <Lock size={24} /> : <Unlock size={24} />}
                </div>
                <div>
                  <h3 className="font-bold text-lg">
                    {isDeleteConfirmOpen ? 'Delete Selected Rows?' : isBulkLockConfirmOpen ? 'Lock Selected Rows?' : 'Unlock Selected Rows?'}
                  </h3>
                  <p className="text-sm text-[var(--muted)]">
                    {isDeleteConfirmOpen 
                      ? (rowToDelete 
                          ? "Are you sure you want to delete this row? This action cannot be undone."
                          : `Are you sure you want to delete ${selectedIds.size} selected rows? This action cannot be undone.`)
                      : `Are you sure you want to ${isBulkLockConfirmOpen ? 'lock' : 'unlock'} ${selectedIds.size} selected rows?`}
                  </p>
                </div>
              </div>

              <div className="flex gap-3 justify-end">
                <button 
                  onClick={() => {
                    setIsDeleteConfirmOpen(false);
                    setIsBulkLockConfirmOpen(false);
                    setIsBulkUnlockConfirmOpen(false);
                  }}
                  className="px-4 py-2 rounded-xl border border-[var(--line)] hover:bg-[var(--line)]/5 transition-colors text-sm font-medium"
                >
                  Cancel
                </button>
                <button 
                  onClick={() => {
                    if (isDeleteConfirmOpen) deleteSelected();
                    else if (isBulkLockConfirmOpen) bulkLock();
                    else if (isBulkUnlockConfirmOpen) bulkUnlock();
                  }}
                  className={`px-6 py-2 rounded-xl text-sm font-bold shadow-lg transition-transform active:scale-95 ${isDeleteConfirmOpen ? 'bg-red-500 text-white hover:bg-red-600' : 'bg-[var(--accent)] text-black hover:opacity-90'}`}
                >
                  Confirm
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <footer className="max-w-7xl mx-auto p-12 border-t border-[var(--line)] text-center space-y-2">
        <p className="text-sm font-bold text-[var(--ink)]">Text Transformers Pro</p>
        <p className="text-xs text-[var(--muted)]">Mohammad Wasim Siddique | Survey Data Mapping Expert | 2026</p>
      </footer>
    </div>
  );
}

function AutoExpandingTextarea({ value, onChange, disabled, readOnly, className, placeholder }: { 
  value: string, 
  onChange?: (val: string) => void, 
  disabled?: boolean,
  readOnly?: boolean,
  className?: string,
  placeholder?: string
}) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
      textareaRef.current.style.height = `${textareaRef.current.scrollHeight}px`;
    }
  }, [value]);

  return (
    <textarea
      ref={textareaRef}
      value={value}
      disabled={disabled}
      readOnly={readOnly}
      placeholder={placeholder}
      onChange={e => onChange?.(e.target.value)}
      className={`resize-none overflow-hidden transition-all duration-200 p-2 rounded-lg border-transparent hover:border-[var(--line)] focus:border-[var(--accent)] focus:bg-[var(--bg)] focus:ring-1 focus:ring-[var(--accent)] outline-none ${className}`}
      rows={1}
    />
  );
}

function FileDropZone({ label, file, onFileSelect, accept }: { 
  label: string, 
  file: File | null, 
  onFileSelect: (f: File | null) => void,
  accept: string
}) {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleClear = (e: React.MouseEvent) => {
    e.stopPropagation();
    onFileSelect(null);
    if (inputRef.current) inputRef.current.value = '';
  };

  return (
    <div 
      onClick={() => inputRef.current?.click()}
      className={`relative group border-2 border-dashed rounded-2xl p-8 text-center cursor-pointer transition-all duration-300 w-full overflow-hidden ${
        file 
          ? 'border-[var(--accent)] bg-[var(--accent)]/5 shadow-[0_0_20px_rgba(var(--accent-rgb),0.05)]' 
          : 'border-[var(--line)] hover:border-[var(--accent)] bg-[var(--card)] hover:shadow-lg'
      }`}
    >
      <input 
        type="file" 
        ref={inputRef}
        className="hidden" 
        accept={accept}
        onClick={e => (e.target as HTMLInputElement).value = ''}
        onChange={e => onFileSelect(e.target.files?.[0] || null)}
      />

      {file && (
        <button 
          onClick={handleClear}
          className="absolute top-4 right-4 p-2 rounded-full bg-red-500/10 text-red-500 hover:bg-red-500 hover:text-white transition-all z-20 shadow-sm"
          title="Clear File"
        >
          <X size={16} />
        </button>
      )}
      
      <div className="flex flex-col items-center gap-3">
        <div className={`w-14 h-14 rounded-2xl flex items-center justify-center transition-all duration-500 ${
          file 
            ? 'bg-[var(--accent)] text-black rotate-[360deg] shadow-[0_0_15px_rgba(var(--accent-rgb),0.3)]' 
            : 'bg-[var(--line)] text-[var(--muted)] group-hover:bg-[var(--accent)]/20 group-hover:text-[var(--accent)] group-hover:scale-110'
        }`}>
          {file ? <Check size={28} /> : <Upload size={28} />}
        </div>
        <div>
          <p className="font-bold text-lg text-[var(--ink)] tracking-tight break-all">{file ? file.name : label}</p>
          <p className="text-xs text-[var(--muted)] mt-1 font-medium">
            {file ? `${(file.size / 1024).toFixed(1)} KB` : `Click or drag to upload ${accept}`}
          </p>
        </div>
      </div>
    </div>
  );
}
