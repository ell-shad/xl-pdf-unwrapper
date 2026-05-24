/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useRef, useEffect } from 'react';
import JSZip from 'jszip';
import * as CFB from 'cfb';
import { saveAs } from 'file-saver';
import { FileUp, FileText, Download, CheckCircle2, AlertCircle, Loader2, Trash2, Eye, X, Pencil } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface ExtractedPDF {
  id: string;
  name: string;
  originalBinName: string;
  data: Uint8Array;
  size: number;
  previewUrl?: string;
}

const SYSTEM_STREAMS = [
  'Workbook', 'Book', 'WordDocument', 'PowerPoint Document',
  '\x05DocumentSummaryInformation', '\x05SummaryInformation', '\x05PropertySet',
  '\x01CompObj', '\x01Ole', '\x03ObjInfo', '\x01Ole10Native',
  '1Table', '0Table', 'Data', 'Current User', 'Root Entry', 'macros',
  'PROJECT', 'PROJECTwm', '_VBA_PROJECT_CUR', 'VBA', 'dir', 'UserForm'
];

const guessExtensionFromBytes = (data: Uint8Array): string => {
  if (data.length >= 4) {
    if (data[0] === 0x25 && data[1] === 0x50 && data[2] === 0x44 && data[3] === 0x46) return '.pdf'; // %PDF
    if (data[0] === 0x50 && data[1] === 0x4B && data[2] === 0x03 && data[3] === 0x04) return '.docx'; // PK zip
    if (data[0] === 0xD0 && data[1] === 0xCF && data[2] === 0x11 && data[3] === 0xE0) return '.doc'; // legacy CFB doc/xls
  }
  return '.bin';
};

const isAllowedExtension = (nameOrPath: string): boolean => {
  const parts = nameOrPath.split('.');
  if (parts.length <= 1) return false;
  const ext = parts.pop()?.toLowerCase() || '';
  const allowed = [
    'pdf', 
    'docx', 'doc', 'docm', 'dotx', 'dotm',
    'xlsx', 'xls', 'xlsb', 'xlsm', 'xltx', 'xltm',
    'pptx', 'ppt', 'pptm', 'potx', 'potm'
  ];
  return allowed.includes(ext);
};

const identifyOfficeDocFromCfb = (fileData: Uint8Array): string | null => {
  try {
    const cfb = CFB.read(fileData, { type: 'buffer' });
    for (const path of cfb.FullPaths) {
      const lowerPath = path.toLowerCase();
      if (lowerPath.includes('workbook')) return '.xls';
      if (lowerPath.includes('worddocument')) return '.doc';
      if (lowerPath.includes('powerpoint document')) return '.ppt';
    }
  } catch (e) {}
  return null;
};

const identifyOfficeDocFromZip = async (fileData: Uint8Array): Promise<string | null> => {
  try {
    const zip = new JSZip();
    const zipContent = await zip.loadAsync(fileData);
    const files = Object.keys(zipContent.files);
    if (files.some(f => f.includes('word/document.xml'))) return '.docx';
    if (files.some(f => f.includes('xl/workbook.xml'))) return '.xlsx';
    if (files.some(f => f.includes('ppt/presentation.xml'))) return '.pptx';
  } catch (e) {}
  return null;
};

const addExtractedFile = (fileData: Uint8Array, name: string, originalPath: string, results: ExtractedPDF[]) => {
  if (results.some(r => r.size === fileData.length && r.data.every((val, idx) => val === fileData[idx]))) {
    return;
  }

  const cleanedName = name.replace(/^[^\x20-\x7E]+/, '').trim() || 'extracted_file';
  const extension = cleanedName.split('.').pop()?.toLowerCase();

  // Enforce allowed extensions (only office docs and pdfs)
  if (!isAllowedExtension(cleanedName)) {
    return;
  }

  // Reject VBA objects, macros, or .bin system streams
  const nameL = cleanedName.toLowerCase();
  if (nameL.includes('vba') || nameL.includes('vbe') || nameL.includes('macro') || nameL.endsWith('.bin')) {
    return;
  }

  const isPdf = extension === 'pdf';
  let previewUrl = undefined;
  if (isPdf) {
    const blob = new Blob([fileData], { type: 'application/pdf' });
    previewUrl = URL.createObjectURL(blob);
  }

  results.push({
    id: Math.random().toString(36).substr(2, 9),
    name: cleanedName,
    originalBinName: originalPath,
    data: fileData,
    size: fileData.length,
    previewUrl
  });
};

const parseOle10Native = (data: Uint8Array): { fileName: string, fileData: Uint8Array } | null => {
  if (data.length < 16) return null;
  
  let cursor = 0;
  const totalSize = data[0] | (data[1] << 8) | (data[2] << 16) | (data[3] << 24);
  cursor += 4;
  
  let sig = data[cursor] | (data[cursor+1] << 8);
  if (sig !== 0x0002 && sig !== 0x0003) {
    if (data[0] === 0x02 && data[1] === 0x00) {
      cursor = 0;
      sig = data[cursor] | (data[cursor+1] << 8);
    } else {
      return null;
    }
  }
  cursor += 2;
  
  const readNullTerminatedString = () => {
    let s = "";
    while (cursor < data.length && data[cursor] !== 0) {
      s += String.fromCharCode(data[cursor]);
      cursor++;
    }
    cursor++; // skip null byte
    return s;
  };
  
  const label = readNullTerminatedString();
  const originalPath = readNullTerminatedString();
  
  cursor += 4; 
  
  const tempPath = readNullTerminatedString();
  
  if (cursor + 4 > data.length) return null;
  
  const payloadSize = data[cursor] | (data[cursor+1] << 8) | (data[cursor+2] << 16) | (data[cursor+3] << 24);
  cursor += 4;
  
  if (payloadSize <= 0 || cursor + payloadSize > data.length) {
    return null;
  }
  
  const payload = data.slice(cursor, cursor + payloadSize);
  
  let name = label || originalPath.split(/[/\\]/).pop() || tempPath.split(/[/\\]/).pop() || "embedded_file";
  name = name.replace(/[^\x20-\x7E]+/g, '').trim();
  
  return {
    fileName: name,
    fileData: payload
  };
};

export default function App() {
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [extractedFiles, setExtractedFiles] = useState<ExtractedPDF[]>([]);
  const [extractAll, setExtractAll] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingName, setEditingName] = useState("");
  const [previewId, setPreviewId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Cleanup object URLs on unmount
  useEffect(() => {
    return () => {
      extractedFiles.forEach(file => {
        if (file.previewUrl) URL.revokeObjectURL(file.previewUrl);
      });
    };
  }, [extractedFiles]);

  const processFile = async (file: File) => {
    setIsProcessing(true);
    setError(null);
    
    // Revoke old URLs
    extractedFiles.forEach(f => {
      if (f.previewUrl) URL.revokeObjectURL(f.previewUrl);
    });
    setExtractedFiles([]);

    const results: ExtractedPDF[] = [];

    try {
      const arrayBuffer = await file.arrayBuffer();
      const isZip = /\.(xlsx|xlsm|xlsb|docx|docm|dotx)$/i.test(file.name);
      
      if (isZip) {
        let zipContent;
        try {
          const zip = new JSZip();
          zipContent = await zip.loadAsync(arrayBuffer);
        } catch (e) {
          throw new Error("This file appears to be encrypted or corrupted.");
        }

        const filesInZip = Object.keys(zipContent.files);
        const embeddedCandidates = filesInZip.filter(path => {
          const lowerPath = path.toLowerCase();
          return lowerPath.includes("embeddings/") || lowerPath.endsWith(".bin");
        });

        if (embeddedCandidates.length === 0) {
          throw new Error("No embedded objects were found in this document.");
        }

        for (const candidatePath of embeddedCandidates) {
          const fileData = await zipContent.files[candidatePath].async("uint8array");
          
          // First, check if this file contains a PDF directly (the same robust PDF extraction logic)!
          const wasPdfExtracted = await extractPdfFromStream(fileData, candidatePath, results);
          
          if (!extractAll) {
            // PDF Only Mode (default): We only care about extracting PDFs, so we are done with this candidate
            continue;
          }
          
          if (wasPdfExtracted) {
            // In All Embeds Mode, since the PDF was successfully found and pulled out, we are also done
            continue;
          }
          
          // Fallback to other allowed types if extractAll is True:
          const isCfb = fileData[0] === 0xD0 && fileData[1] === 0xCF && fileData[2] === 0x11 && fileData[3] === 0xE0;
          const streamName = candidatePath.split('/').pop() || "embedded_file";
          
          if (isCfb) {
            try {
              const cfb = CFB.read(fileData, { type: 'buffer' });
              let processedAny = false;
              
              // Find \x01Ole10Native streams
              for (const path of cfb.FullPaths) {
                if (path.includes("\x01Ole10Native")) {
                  const entry = CFB.find(cfb, path);
                  if (entry && entry.content) {
                    const oleData = entry.content as Uint8Array;
                    const parsed = parseOle10Native(oleData);
                    if (parsed) {
                      // Check if the inner OleData is a PDF
                      const innerWasPdf = await extractPdfFromStream(parsed.fileData, parsed.fileName, results);
                      if (innerWasPdf) {
                        processedAny = true;
                        continue;
                      }
                      
                      addExtractedFile(parsed.fileData, parsed.fileName, candidatePath, results);
                      processedAny = true;
                    }
                  }
                }
              }
              
              if (!processedAny) {
                for (const path of cfb.FullPaths) {
                  const sn = path.split('/').pop() || "";
                  const isSystemStream = SYSTEM_STREAMS.some(sys => sn.includes(sys) || sys.includes(sn));
                  if (!isSystemStream) {
                    const entry = CFB.find(cfb, path);
                    if (entry && entry.content && (entry.content as Uint8Array).length > 20) {
                      const contentBytes = entry.content as Uint8Array;
                      
                      // Check if the inner content is a PDF
                      const innerWasPdf = await extractPdfFromStream(contentBytes, sn, results);
                      if (innerWasPdf) {
                        processedAny = true;
                        continue;
                      }
                      
                      const guessedExt = guessExtensionFromBytes(contentBytes);
                      let name = sn.replace(/[^\x20-\x7E]+/g, '').trim() || "embedded_object";
                      if (!name.includes('.')) name += guessedExt;
                      addExtractedFile(contentBytes, name, candidatePath, results);
                      processedAny = true;
                    }
                  }
                }
              }
              
              if (!processedAny) {
                const guessedExt = guessExtensionFromBytes(fileData);
                let name = streamName;
                if (!name.includes('.')) name += guessedExt;
                addExtractedFile(fileData, name, candidatePath, results);
              }
            } catch (e) {
              const guessedExt = guessExtensionFromBytes(fileData);
              let name = streamName;
              if (!name.includes('.')) name += guessedExt;
              addExtractedFile(fileData, name, candidatePath, results);
            }
          } else {
            const guessedExt = guessExtensionFromBytes(fileData);
            let name = streamName;
            if (candidatePath.toLowerCase().endsWith('.bin') && guessedExt !== '.bin') {
              name = name.replace(/\.bin$/i, guessedExt);
            }
            if (!name.includes('.')) name += guessedExt;
            addExtractedFile(fileData, name, candidatePath, results);
          }
        }
      } else if (/\.(xls|doc)$/i.test(file.name)) {
        try {
          const cfb = CFB.read(new Uint8Array(arrayBuffer), { type: 'buffer' });
          
          for (const path of cfb.FullPaths) {
            const entry = CFB.find(cfb, path);
            if (entry && entry.content && (entry.content as Uint8Array).length > 20) {
              const contentBytes = entry.content as Uint8Array;
              
              // First, check if this contains a PDF (PDF-only logic)
              const wasPdfExtracted = await extractPdfFromStream(contentBytes, path, results);
              
              if (!extractAll) {
                // PDF Only Mode (default): we are done for this nested path
                continue;
              }
              
              if (wasPdfExtracted) {
                // Already extracted as a PDF, so we do not extract standard office doc
                continue;
              }
              
              // Extract All Mode: process other office doc formats
              const streamName = path.split('/').pop() || "";
              const isSystemStream = SYSTEM_STREAMS.some(sys => streamName.includes(sys) || sys.includes(streamName));
              
              if (!isSystemStream && !path.includes("\x01Ole10Native")) {
                const guessedExt = guessExtensionFromBytes(contentBytes);
                let name = streamName.replace(/[^\x20-\x7E]+/g, '').trim() || "embedded_object";
                if (!name.includes('.')) name += guessedExt;
                addExtractedFile(contentBytes, name, path, results);
              }
            }
          }
          
          // Let's also look at \x01Ole10Native streams in legacy docs
          if (extractAll) {
            for (const path of cfb.FullPaths) {
              if (path.includes("\x01Ole10Native")) {
                const entry = CFB.find(cfb, path);
                if (entry && entry.content) {
                  const parsed = parseOle10Native(entry.content as Uint8Array);
                  if (parsed) {
                    const innerWasPdf = await extractPdfFromStream(parsed.fileData, parsed.fileName, results);
                    if (!innerWasPdf) {
                      addExtractedFile(parsed.fileData, parsed.fileName, path, results);
                    }
                  }
                }
              }
            }
          }
        } catch (e) {
          throw new Error("Failed to parse the legacy document stream.");
        }
      } else {
        throw new Error("Unsupported file format.");
      }

      if (results.length === 0) {
        throw new Error(extractAll ? "No embedded attachments found." : "No PDF attachments identified.");
      }

      setExtractedFiles(results);
    } catch (err: any) {
      setError(err.message || "An error occurred.");
    } finally {
      setIsProcessing(false);
    }
  };

  const extractPdfFromStream = async (streamData: Uint8Array, sourcePath: string, results: ExtractedPDF[]): Promise<boolean> => {
    const pdfHeader = new Uint8Array([0x25, 0x50, 0x44, 0x46]); // %PDF
    const eofMarker = new Uint8Array([0x25, 0x25, 0x45, 0x4F, 0x46]); // %%EOF
    
    let startIndex = -1;
    for (let i = 0; i < streamData.length - 4; i++) {
      if (streamData[i] === pdfHeader[0] && 
          streamData[i+1] === pdfHeader[1] && 
          streamData[i+2] === pdfHeader[2] && 
          streamData[i+3] === pdfHeader[3]) {
        startIndex = i;
        break;
      }
    }

    if (startIndex === -1) return false;

    let endIndex = -1;
    for (let j = streamData.length - 5; j >= startIndex; j--) {
      if (streamData[j] === eofMarker[0] && 
          streamData[j+1] === eofMarker[1] && 
          streamData[j+2] === eofMarker[2] && 
          streamData[j+3] === eofMarker[3] && 
          streamData[j+4] === eofMarker[4]) {
        endIndex = j + 5; 
        break;
      }
    }

    if (endIndex === -1) {
      endIndex = streamData.length;
    }

    const pdfData = streamData.slice(startIndex, endIndex);

    if (results.some(r => r.size === pdfData.length && r.data.every((val, idx) => val === pdfData[idx]))) {
      return true;
    }

    let fileName = sourcePath.split('/').pop()?.replace('.bin', '.pdf').replace('\x01', '') || 'extracted.pdf';
    fileName = fileName.replace(/^oleObject\d+/, 'attachment').replace(/^[^\x20-\x7E]+/, '');
    if (!fileName.toLowerCase().endsWith('.pdf')) fileName += '.pdf';

    const blob = new Blob([pdfData], { type: 'application/pdf' });
    const previewUrl = URL.createObjectURL(blob);

    results.push({
      id: Math.random().toString(36).substr(2, 9),
      name: fileName,
      originalBinName: sourcePath,
      data: pdfData,
      size: pdfData.length,
      previewUrl
    });

    return true;
  };

  const handleDragOver = (e: React.DragEvent) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = () => { setIsDragging(false); };
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.match(/\.(xlsx?|xlsm|xlsb|docx?|docm|dotx)$/i))) processFile(file);
    else setError("Please upload a valid spreadsheet or document.");
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file);
    }
    e.target.value = "";
  };

  const downloadFile = (file: ExtractedPDF) => {
    const ext = file.name.split('.').pop()?.toLowerCase();
    let type = 'application/octet-stream';
    if (ext === 'pdf') type = 'application/pdf';
    else if (ext === 'jpeg' || ext === 'jpg') type = 'image/jpeg';
    else if (ext === 'png') type = 'image/png';
    else if (ext === 'gif') type = 'image/gif';
    else if (ext === 'docx') type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    else if (ext === 'doc') type = 'application/msword';
    else if (ext === 'xlsx') type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    else if (ext === 'xls') type = 'application/vnd.ms-excel';
    
    const blob = new Blob([file.data], { type });
    saveAs(blob, file.name);
  };

  const downloadAll = async () => {
    const zip = new JSZip();
    extractedFiles.forEach(file => {
      zip.file(file.name, file.data);
    });
    const content = await zip.generateAsync({ type: "blob" });
    saveAs(content, extractAll ? "extracted_attachments.zip" : "extracted_pdfs.zip");
  };

  const startEditing = (e: React.MouseEvent, file: ExtractedPDF) => {
    e.stopPropagation();
    setEditingId(file.id);
    setEditingName(file.name);
  };

  const saveName = () => {
    if (!editingId) return;
    let newName = editingName.trim();
    if (newName) {
      const currentFile = extractedFiles.find(f => f.id === editingId);
      if (currentFile) {
        const originalExt = currentFile.name.split('.').pop();
        if (originalExt && !newName.toLowerCase().endsWith('.' + originalExt.toLowerCase())) {
          newName += '.' + originalExt;
        }
      }
      setExtractedFiles(prev => prev.map(f => f.id === editingId ? { ...f, name: newName } : f));
    }
    setEditingId(null);
  };

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const activePreview = previewId ? extractedFiles.find(f => f.id === previewId) : null;

  return (
    <div id="app-root" className="min-h-screen bg-slate-50 dark:bg-[#0a0a0b] text-slate-900 dark:text-slate-100 font-sans selection:bg-orange-100 dark:selection:bg-orange-900/30 p-4 md:p-8 transition-colors duration-300">
      <div id="container" className="max-w-4xl mx-auto space-y-12">
        
        {/* Header */}
        <header id="header" className="space-y-4 text-center md:text-left">
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="inline-flex items-center gap-3"
          >
            <div className="p-3 bg-orange-500 rounded-2xl text-white shadow-lg shadow-orange-500/20">
              <FileText size={36} />
            </div>
            <h1 id="title" className="text-4xl md:text-5xl font-bold tracking-tight bg-clip-text text-transparent bg-gradient-to-r from-slate-900 to-slate-600 dark:from-white dark:to-slate-400">
              XL-PDF Unwrapper
            </h1>
          </motion.div>
          <p id="description" className="text-lg text-slate-500 dark:text-slate-400 max-w-3xl font-light">
            Recover embedded attachments and PDFs from spreadsheets and documents securely in your browser.
          </p>
        </header>

        <main id="main" className="space-y-8 pb-20">
          
          {/* Settings Segmented Option bar */}
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-4 rounded-3xl flex flex-col md:flex-row items-center justify-between gap-4 shadow-sm">
            <div className="space-y-1 text-center md:text-left">
              <span className="text-xs uppercase font-extrabold tracking-wider text-orange-500">Extraction Mode</span>
              <p className="text-sm font-semibold text-slate-700 dark:text-slate-300">
                {extractAll 
                  ? "Extract all embedded attachments (doc, xls, ppt, images, zip) in their native format." 
                  : "Detect, decode and pull clean PDF attachments out of file streams."}
              </p>
            </div>
            
            <div className="flex bg-slate-100 dark:bg-slate-800/80 p-1 rounded-2xl w-full md:w-auto self-stretch md:self-auto">
              <button
                type="button"
                onClick={() => {
                  setExtractAll(false);
                  setExtractedFiles([]);
                  setError(null);
                }}
                className={`flex-1 md:flex-none px-5 py-2.5 rounded-xl text-xs font-bold transition-all whitespace-nowrap ${!extractAll ? 'bg-white dark:bg-slate-700 text-orange-600 dark:text-orange-400 shadow-sm' : 'text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200'}`}
              >
                PDF Only
              </button>
              <button
                type="button"
                onClick={() => {
                  setExtractAll(true);
                  setExtractedFiles([]);
                  setError(null);
                }}
                className={`flex-1 md:flex-none px-5 py-2.5 rounded-xl text-xs font-bold transition-all whitespace-nowrap ${extractAll ? 'bg-white dark:bg-slate-700 text-orange-600 dark:text-orange-400 shadow-sm' : 'text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200'}`}
              >
                All Embeds
              </button>
            </div>
          </div>

          {/* Dropzone */}
          <div
            id="dropzone"
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`
              relative group cursor-pointer
              border-2 border-dashed rounded-[2.5rem] p-12 transition-all duration-500
              flex flex-col items-center justify-center space-y-6
              ${isDragging 
                ? 'border-orange-500 bg-orange-50/50 dark:bg-orange-500/5 scale-[0.98] shadow-2xl shadow-orange-500/10' 
                : 'border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900/50 hover:border-orange-400 dark:hover:border-orange-500 hover:bg-slate-50 dark:hover:bg-slate-800/50 shadow-xl shadow-slate-200/50 dark:shadow-none'
              }
            `}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileInput}
              accept=".xlsx,.xlsm,.xlsb,.xls,.docx,.docm,.dotx,.doc"
              className="hidden"
            />
            
            <div id="icon-container" className={`p-8 rounded-full transition-all duration-500 transform group-hover:scale-110 ${isDragging ? 'bg-orange-100 dark:bg-orange-500/20' : 'bg-slate-100 dark:bg-slate-800 group-hover:bg-orange-50 dark:group-hover:bg-orange-500/10'}`}>
              <FileUp 
                size={56} 
                className={`transition-colors duration-500 ${isDragging ? 'text-orange-600' : 'text-slate-400 group-hover:text-orange-500'}`} 
              />
            </div>
            
            <div id="text-container" className="text-center space-y-2">
              <h3 className="text-2xl font-semibold">
                {isProcessing ? 'Analyzing Document...' : 'Upload Document'}
              </h3>
              <p className="text-slate-500 dark:text-slate-400 font-light">
                {isProcessing ? 'Decompressing and parsing layout objects' : 'Drag & drop Excel or Word files here'}
              </p>
            </div>

            {isProcessing && (
              <motion.div 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="absolute inset-0 bg-white/80 dark:bg-[#0a0a0b]/80 backdrop-blur-md rounded-[2.5rem] flex items-center justify-center z-10"
              >
                <div className="flex flex-col items-center space-y-4">
                  <div className="relative">
                    <Loader2 className="animate-spin text-orange-500" size={48} />
                    <div className="absolute inset-0 blur-xl bg-orange-500/20 animate-pulse"></div>
                  </div>
                  <span className="font-medium text-slate-700 dark:text-slate-300 tracking-wide uppercase text-xs">Processing Stream...</span>
                </div>
              </motion.div>
            )}
          </div>

          {/* Error Message */}
          <AnimatePresence>
            {error && (
              <motion.div
                initial={{ opacity: 0, height: 0 }}
                animate={{ opacity: 1, height: 'auto' }}
                exit={{ opacity: 0, height: 0 }}
                className="overflow-hidden"
              >
                <div className="bg-red-50 dark:bg-red-500/10 border border-red-100 dark:border-red-500/20 p-5 rounded-3xl flex items-start gap-4 text-red-700 dark:text-red-400">
                  <div className="p-2 bg-red-100 dark:bg-red-500/20 rounded-xl">
                    <AlertCircle size={20} />
                  </div>
                  <div className="flex-1 space-y-1">
                    <p className="font-semibold">Extraction Failed</p>
                    <p className="text-sm opacity-80 leading-relaxed">{error}</p>
                  </div>
                  <button onClick={() => setError(null)} className="p-2 hover:bg-red-200/50 dark:hover:bg-red-500/20 rounded-xl transition-colors">
                    <Trash2 size={18} />
                  </button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Results Area */}
          <AnimatePresence>
            {extractedFiles.length > 0 && (
              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="space-y-6"
              >
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 px-2">
                  <h2 className="text-xl font-bold flex items-center gap-3">
                    <div className="p-1.5 bg-green-500 rounded-full text-white">
                      <CheckCircle2 size={16} />
                    </div>
                    {extractedFiles.length} {extractedFiles.length === 1 
                      ? (extractAll ? 'Attachment Found' : 'PDF Found') 
                      : (extractAll ? 'Attachments Found' : 'PDFs Found')
                    }
                  </h2>
                  <div className="flex items-center gap-3">
                    {extractedFiles.length > 1 && (
                      <button 
                        onClick={downloadAll}
                        className="flex items-center gap-2 px-4 py-2 bg-orange-100 dark:bg-orange-500/10 text-orange-600 dark:text-orange-400 rounded-xl font-bold text-sm hover:bg-orange-500 hover:text-white transition-all"
                      >
                        <Download size={16} />
                        Download All
                      </button>
                    )}
                    <button 
                      onClick={() => setExtractedFiles([])}
                      className="text-sm font-medium text-slate-400 hover:text-orange-500 transition-colors"
                    >
                      Clear All
                    </button>
                  </div>
                </div>

                <div className="grid gap-4">
                  {extractedFiles.map((file, index) => {
                    const ext = file.name.split('.').pop()?.toLowerCase() || '';
                    const isPdf = ext === 'pdf';
                    const isImage = ['png', 'jpg', 'jpeg', 'gif'].includes(ext);
                    const isPreviewable = isPdf || isImage;
                    
                    let tagText = ext.toUpperCase();
                    let iconBg = 'bg-slate-100 dark:bg-slate-800 text-slate-500';
                    if (isPdf) {
                      iconBg = 'bg-red-50 dark:bg-red-500/10 text-red-500 dark:text-red-400 group-hover:bg-red-500';
                    } else if (['docx', 'doc', 'docm'].includes(ext)) {
                      iconBg = 'bg-blue-50 dark:bg-blue-500/10 text-blue-500 dark:text-blue-400 group-hover:bg-blue-500';
                    } else if (['xlsx', 'xls', 'xlsb', 'xlsm'].includes(ext)) {
                      iconBg = 'bg-emerald-50 dark:bg-emerald-500/10 text-emerald-500 dark:text-emerald-400 group-hover:bg-emerald-500';
                    } else if (['pptx', 'ppt'].includes(ext)) {
                      iconBg = 'bg-orange-50 dark:bg-orange-500/10 text-orange-500 dark:text-orange-400 group-hover:bg-orange-500';
                    } else if (isImage) {
                      iconBg = 'bg-purple-50 dark:bg-purple-500/10 text-purple-500 dark:text-purple-400 group-hover:bg-purple-500';
                    }

                    return (
                      <motion.div
                        key={file.id}
                        initial={{ opacity: 0, scale: 0.95 }}
                        animate={{ opacity: 1, scale: 1 }}
                        transition={{ delay: index * 0.05 }}
                        className="group bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-5 rounded-[2rem] flex flex-col sm:flex-row items-center gap-5 hover:border-orange-300 dark:hover:border-orange-500/50 hover:shadow-2xl hover:shadow-orange-500/5 transition-all duration-300"
                      >
                        <div 
                          className={`p-4 rounded-2xl group-hover:text-white transition-all duration-300 cursor-pointer ${iconBg}`}
                          onClick={() => setPreviewId(file.id)}
                        >
                          {isPreviewable ? <Eye size={32} /> : <FileText size={32} />}
                        </div>
                        
                        <div className="flex-1 min-w-0 w-full">
                          {editingId === file.id ? (
                            <div className="flex items-center gap-2">
                              <input
                                type="text"
                                value={editingName}
                                onChange={(e) => setEditingName(e.target.value)}
                                onBlur={saveName}
                                onKeyDown={(e) => e.key === 'Enter' && saveName()}
                                autoFocus
                                className="flex-1 bg-slate-50 dark:bg-slate-800 border-2 border-orange-500 rounded-xl px-3 py-1.5 font-bold outline-none"
                              />
                              <button onClick={saveName} className="p-2 bg-green-500 text-white rounded-xl">
                                <CheckCircle2 size={18} />
                              </button>
                            </div>
                          ) : (
                            <div className="space-y-1">
                              <div className="flex items-center gap-2 group/title">
                                <h4 className="font-bold text-lg truncate group-hover:text-orange-600 dark:group-hover:text-orange-400 transition-colors">
                                  {file.name}
                                </h4>
                                <button 
                                  onClick={(e) => startEditing(e, file)}
                                  className="opacity-0 group-hover/title:opacity-100 p-1.5 hover:bg-slate-100 dark:hover:bg-slate-800 rounded-lg transition-all"
                                  title="Rename"
                                >
                                  <Pencil size={14} className="text-slate-400" />
                                </button>
                              </div>
                              <div className="flex items-center gap-x-3 text-sm text-slate-400">
                                <span className="font-medium text-slate-600 dark:text-slate-300">{formatSize(file.size)}</span>
                                <span className="opacity-30">•</span>
                                <span className="px-2 py-0.5 rounded-md bg-slate-100 dark:bg-slate-800 text-[10px] uppercase font-bold text-slate-600 dark:text-slate-400">{tagText}</span>
                                <span className="opacity-30">•</span>
                                <span className="truncate italic max-w-[200px] inline-block">ID: {file.originalBinName.split('/').pop()}</span>
                              </div>
                            </div>
                          )}
                        </div>

                        <div className="flex items-center gap-3 w-full sm:w-auto">
                          <button
                            onClick={() => setPreviewId(file.id)}
                            className="flex-1 sm:w-auto px-5 py-4 bg-orange-100 dark:bg-orange-500/10 text-orange-600 dark:text-orange-400 rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-orange-500 hover:text-white transition-all duration-300"
                          >
                            <Eye size={20} />
                            {isPreviewable ? 'Preview' : 'Info'}
                          </button>
                          <button
                            onClick={() => downloadFile(file)}
                            className="flex-1 sm:w-auto px-6 py-4 bg-slate-900 dark:bg-white text-white dark:text-slate-900 group-hover:bg-orange-500 group-hover:text-white rounded-2xl transition-all duration-300 flex items-center justify-center gap-3 font-bold"
                          >
                            <Download size={20} />
                            Download
                          </button>
                        </div>
                      </motion.div>
                    );
                  })}
                </div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Tips Section */}
          {!isProcessing && extractedFiles.length === 0 && !error && (
            <section id="tips" className="bg-white/50 dark:bg-slate-900/30 border border-slate-200 dark:border-slate-800/50 rounded-[2.5rem] p-10 space-y-8">
              <h3 className="text-xl font-bold text-slate-800 dark:text-slate-200">How it works</h3>
              <div className="grid sm:grid-cols-3 gap-8">
                <div className="space-y-4">
                  <div className="w-10 h-10 rounded-full bg-orange-100 dark:bg-orange-500/20 text-orange-600 dark:text-orange-400 flex items-center justify-center font-bold text-lg">1</div>
                  <p className="text-sm text-slate-600 dark:text-slate-400 leading-relaxed font-light">
                    Upload a spreadsheet or document (<span className="font-medium text-slate-800 dark:text-slate-200">.xlsx, .xls, .docx, .doc</span>).
                  </p>
                </div>
                <div className="space-y-4">
                  <div className="w-10 h-10 rounded-full bg-orange-100 dark:bg-orange-500/20 text-orange-600 dark:text-orange-400 flex items-center justify-center font-bold text-lg">2</div>
                  <p className="text-sm text-slate-600 dark:text-slate-400 leading-relaxed font-light">
                    The app scans zip/compound paths and decompresses embedded OLE streams safely.
                  </p>
                </div>
                <div className="space-y-4">
                  <div className="w-10 h-10 rounded-full bg-orange-100 dark:bg-orange-500/20 text-orange-600 dark:text-orange-400 flex items-center justify-center font-bold text-lg">3</div>
                  <p className="text-sm text-slate-600 dark:text-slate-400 leading-relaxed font-light">
                    Preview, rename and download. <span className="font-medium text-slate-800 dark:text-slate-200">100% standard static client-side execution</span>. No remote APIs are requested.
                  </p>
                </div>
              </div>
            </section>
          )}

        </main>

        <footer id="footer" className="pt-12 border-t border-slate-200 dark:border-slate-800 text-center space-y-4">
          <p className="text-sm font-medium text-slate-500 dark:text-slate-400">
            © {new Date().getFullYear()} • <a href="https://az.linkedin.com/in/elshad-guliyev" target="_blank" rel="noopener noreferrer" className="text-orange-500 hover:underline transition-all">Elshad Guliyev</a>
          </p>
        </footer>

        {/* Preview Modal */}
        <AnimatePresence>
          {activePreview && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 z-50 flex items-center justify-center p-4 md:p-8 bg-slate-900/80 backdrop-blur-sm"
              onClick={() => setPreviewId(null)}
            >
              <motion.div
                initial={{ scale: 0.95, opacity: 0, y: 20 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                exit={{ scale: 0.95, opacity: 0, y: 20 }}
                className="bg-white dark:bg-slate-900 w-full h-full max-w-5xl rounded-[2rem] overflow-hidden flex flex-col shadow-2xl shadow-black/50"
                onClick={e => e.stopPropagation()}
              >
                <div className="flex items-center justify-between p-6 border-b border-slate-100 dark:border-slate-800">
                  <div className="flex items-center gap-4">
                    <div className="p-2 bg-orange-50 dark:bg-orange-500/10 text-orange-500 dark:text-orange-400 rounded-xl">
                      <FileText size={24} />
                    </div>
                    <div>
                      <h3 className="font-bold text-lg leading-none">{activePreview.name}</h3>
                      <p className="text-xs text-slate-400 mt-1 uppercase tracking-wider">{formatSize(activePreview.size)}</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-3">
                    {activePreview.previewUrl && (
                      <button 
                        onClick={() => window.open(activePreview.previewUrl, '_blank')}
                        className="p-3 bg-slate-100 dark:bg-slate-800 hover:bg-orange-500 hover:text-white rounded-xl transition-all"
                        title="Open in New Tab"
                      >
                        <Eye size={20} />
                      </button>
                    )}
                    <button 
                      onClick={() => downloadFile(activePreview)}
                      className="p-3 bg-slate-100 dark:bg-slate-800 hover:bg-orange-500 hover:text-white rounded-xl transition-all"
                      title="Download"
                    >
                      <Download size={20} />
                    </button>
                    <button 
                      onClick={() => setPreviewId(null)}
                      className="p-3 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-xl transition-all"
                    >
                      <X size={20} />
                    </button>
                  </div>
                </div>
                <div className="flex-1 bg-slate-100 dark:bg-[#050506] relative">
                  {activePreview.previewUrl ? (
                    <object
                      data={activePreview.previewUrl}
                      type={activePreview.name.toLowerCase().endsWith('.pdf') ? "application/pdf" : undefined}
                      className="w-full h-full border-none"
                    >
                      <div className="absolute inset-0 flex flex-col items-center justify-center p-8 text-center space-y-4">
                        <div className="p-4 bg-slate-200 dark:bg-slate-800 rounded-full">
                          <AlertCircle className="text-slate-400" size={48} />
                        </div>
                        <div className="space-y-2">
                          <p className="font-bold text-lg text-slate-700 dark:text-slate-300">Inline Viewer Blocked</p>
                          <p className="text-slate-500 dark:text-slate-400 max-w-xs text-sm">Chrome's security settings often restrict local previews in nested sandboxes.</p>
                        </div>
                        <button 
                           onClick={() => window.open(activePreview.previewUrl, '_blank')}
                           className="px-8 py-4 bg-orange-500 text-white rounded-2xl font-bold shadow-lg shadow-orange-500/20 hover:scale-105 transition-all"
                        >
                          Open Preview in New Tab
                        </button>
                      </div>
                    </object>
                  ) : (
                    <div className="absolute inset-0 flex flex-col items-center justify-center p-8 text-center space-y-4">
                      <div className="p-4 bg-orange-50 dark:bg-orange-950/20 rounded-full text-orange-500">
                        <FileText size={48} />
                      </div>
                      <div className="space-y-2">
                        <p className="font-bold text-lg text-slate-700 dark:text-slate-300">Alternative File Format</p>
                        <p className="text-slate-500 dark:text-slate-400 max-w-sm text-sm">
                          Extracted file <span className="font-bold text-slate-600 dark:text-slate-300">"{activePreview.name}"</span> is a native attachment that cannot be viewed directly inside the browser.
                        </p>
                      </div>
                      <button 
                         onClick={() => downloadFile(activePreview)}
                         className="px-8 py-4 bg-orange-500 text-white rounded-2xl font-bold shadow-lg shadow-orange-500/20 hover:scale-[1.02] transition-all"
                      >
                        Download Extracted File
                      </button>
                    </div>
                  )}
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

      </div>
    </div>
  );
}

