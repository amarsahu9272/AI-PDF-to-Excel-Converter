
import React, { useState, useCallback, useEffect, useMemo, useRef } from 'react';
import FileUpload from './components/FileUpload';
import DataTable from './components/DataTable';
import { ExcelIcon, ProcessingIcon, ErrorIcon, SunIcon, MoonIcon, CheckCircleIcon, XCircleIcon, QueueListIcon, RetryIcon, DocumentIcon, SearchIcon, TrashIcon, DownloadIcon, BroomIcon, BoltIcon, TurtleIcon, ChevronLeftIcon, ChevronRightIcon, SaveIcon, Cog6ToothIcon, DocumentIconPortrait, DocumentIconLandscape, PencilIcon } from './components/icons';
import { extractDataFromPdfImages } from './services/geminiService';
import type { ProcessJob, ProcessJobStatus, SheetData, ConversionMode, PdfOptions, CellData, CellStyle } from './types';

const MAX_CONCURRENT_JOBS = 3;
const ITEMS_PER_PAGE = 10;
const DEFAULT_PDF_OPTIONS: PdfOptions = { orientation: 'p', fontSize: 10, autoWidth: true };

type Selection = {
  start: { row: number, col: number };
  end: { row: number, col: number };
} | null;


const PdfOptionsModal: React.FC<{ job: ProcessJob; onSave: (jobId: string, options: PdfOptions) => void; onClose: () => void; }> = ({ job, onSave, onClose }) => {
    const [options, setOptions] = useState<PdfOptions>(job.pdfOptions || DEFAULT_PDF_OPTIONS);

    useEffect(() => {
        setOptions(job.pdfOptions || DEFAULT_PDF_OPTIONS);
    }, [job]);

    const handleSave = () => {
        onSave(job.id, options);
    };

    return (
        <div className="fixed inset-0 bg-background/80 backdrop-blur-sm z-50 flex items-center justify-center p-4" onClick={onClose}>
            <div className="w-full max-w-md bg-secondary rounded-xl border border-border-color shadow-2xl p-6" onClick={e => e.stopPropagation()}>
                <div className="flex justify-between items-center mb-4">
                    <h3 className="text-lg font-semibold text-text-main">PDF Output Options</h3>
                    <button onClick={onClose} className="p-1 rounded-full hover:bg-border-color text-text-secondary text-2xl leading-none">&times;</button>
                </div>
                <p className="text-sm text-text-secondary mb-6 truncate" title={job.fileName}>File: {job.fileName}</p>
                <div className="space-y-4">
                    <div>
                        <label htmlFor="orientation" className="block text-sm font-medium text-text-main mb-1">Page Orientation</label>
                        <select
                            id="orientation"
                            value={options.orientation}
                            onChange={e => setOptions(o => ({ ...o, orientation: e.target.value as 'p' | 'l' }))}
                            className="w-full px-3 py-2 bg-background border border-border-color rounded-lg focus:ring-1 focus:ring-primary focus:outline-none"
                        >
                            <option value="p">Portrait</option>
                            <option value="l">Landscape</option>
                        </select>
                    </div>
                    <div>
                        <label htmlFor="fontSize" className="block text-sm font-medium text-text-main mb-1">Font Size</label>
                        <select
                            id="fontSize"
                            value={options.fontSize}
                            onChange={e => setOptions(o => ({ ...o, fontSize: parseInt(e.target.value, 10) as 8 | 10 | 12 }))}
                            className="w-full px-3 py-2 bg-background border border-border-color rounded-lg focus:ring-1 focus:ring-primary focus:outline-none"
                        >
                            <option value="8">Small</option>
                            <option value="10">Medium</option>
                            <option value="12">Large</option>
                        </select>
                    </div>
                    <div className="flex items-center" title="Automatically adjust column widths to prevent content from being cut off.">
                        <input
                            id="autoWidth"
                            type="checkbox"
                            checked={options.autoWidth}
                            onChange={e => setOptions(o => ({ ...o, autoWidth: e.target.checked }))}
                            className="h-4 w-4 rounded border-gray-300 text-primary focus:ring-primary"
                        />
                        <label htmlFor="autoWidth" className="ml-2 block text-sm text-text-main cursor-pointer">Scale columns to fit content</label>
                    </div>
                </div>
                <div className="mt-6 flex justify-end gap-4">
                    <button type="button" onClick={onClose} className="px-6 py-2 text-sm font-semibold rounded-lg bg-secondary text-text-secondary hover:bg-border-color border border-border-color transition-colors">Cancel</button>
                    <button type="button" onClick={handleSave} className="px-6 py-2 text-sm font-semibold rounded-lg bg-primary text-white hover:bg-primary-hover transition-colors">Save</button>
                </div>
            </div>
        </div>
    );
};

const App: React.FC = () => {
  const [conversionMode, setConversionMode] = useState<ConversionMode>('pdf-to-excel');
  const [jobs, setJobs] = useState<ProcessJob[]>(() => {
    try {
      const savedJobsJSON = localStorage.getItem('pdf-to-excel-queue');
      if (savedJobsJSON) {
        const savedJobs: ProcessJob[] = JSON.parse(savedJobsJSON);
        return savedJobs.map(job => {
          if (job.status === 'processing' || job.status === 'queued') {
            return {
              ...job,
              status: 'error',
              errorMessage: 'Processing was interrupted. Please re-upload to retry.',
              file: undefined
            };
          }
          return { ...job, file: undefined };
        });
      }
    } catch (error) {
      console.error("Failed to load saved queue from localStorage:", error);
      localStorage.removeItem('pdf-to-excel-queue');
    }
    return [];
  });
  
  const [activeJobsCount, setActiveJobsCount] = useState(0);
  const [currentlyViewing, setCurrentlyViewing] = useState<ProcessJob | null>(null);
  const [isEditingData, setIsEditingData] = useState(false);
  const [editedData, setEditedData] = useState<SheetData[] | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [showDownloadConfirm, setShowDownloadConfirm] = useState(false);
  const [selectedJobIds, setSelectedJobIds] = useState<Set<string>>(new Set());
  const [currentPage, setCurrentPage] = useState(1);
  const [editingOptionsForJobId, setEditingOptionsForJobId] = useState<string | null>(null);
  const selectAllCheckboxRef = useRef<HTMLInputElement>(null);
  
  const [theme, setTheme] = useState(() => {
    if (typeof window === 'undefined') return 'light';
    if (localStorage.theme === 'dark' || (!('theme' in localStorage) && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
      return 'dark';
    }
    return 'light';
  });

  const successfulJobsCount = useMemo(() => jobs.filter(j => j.status === 'success' && j.conversionMode === 'pdf-to-excel').length, [jobs]);

  useEffect(() => {
    try {
      const savableJobs = jobs.map(({ file, ...rest }) => rest);
      localStorage.setItem('pdf-to-excel-queue', JSON.stringify(savableJobs));
    } catch (error) {
      console.error("Failed to save queue to localStorage:", error);
    }
  }, [jobs]);

  useEffect(() => {
    const root = window.document.documentElement;
    if (theme === 'dark') {
      root.classList.add('dark');
      localStorage.setItem('theme', 'dark');
    } else {
      root.classList.remove('dark');
      localStorage.setItem('theme', 'light');
    }
  }, [theme]);

  const toggleTheme = useCallback(() => {
    setTheme(prevTheme => (prevTheme === 'dark' ? 'light' : 'dark'));
  }, []);

  const handleClearAll = useCallback(() => {
    setJobs([]);
    setSearchQuery('');
    setCurrentlyViewing(null);
    setSelectedJobIds(new Set());
    setCurrentPage(1);
  }, []);

  const switchConversionMode = useCallback((mode: ConversionMode) => {
    if (mode !== conversionMode) {
      handleClearAll();
      setConversionMode(mode);
    }
  }, [conversionMode, handleClearAll]);
  
  const generatePdfThumbnail = useCallback(async (file: File): Promise<{ urls: string[] | null; pageCount: number }> => {
    const fileReader = new FileReader();
    return new Promise((resolve) => {
      fileReader.onload = async (event) => {
        if (!event.target?.result) {
          return resolve({ urls: null, pageCount: 0 });
        }
        try {
          // @ts-ignore
          const pdf = await pdfjsLib.getDocument({ data: event.target.result }).promise;
          const pageCount = pdf.numPages;
          const urls: string[] = [];
          const MAX_THUMBNAILS = Math.min(5, pageCount);

          for (let i = 1; i <= MAX_THUMBNAILS; i++) {
            try {
                const page = await pdf.getPage(i);
                const viewport = page.getViewport({ scale: 0.5 });
                const canvas = document.createElement('canvas');
                const context = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;

                if (!context) throw new Error("Could not get canvas context");
                
                await page.render({ canvasContext: context, viewport: viewport }).promise;
                urls.push(canvas.toDataURL('image/jpeg', 0.8));
            } catch (renderError) {
                console.warn(`Could not render page ${i} of ${file.name}. Retrying in safe mode.`, renderError);
                try {
                    const page = await pdf.getPage(i);
                    const viewport = page.getViewport({ scale: 0.5 });
                    const canvas = document.createElement('canvas');
                    const context = canvas.getContext('2d');
                    canvas.height = viewport.height;
                    canvas.width = viewport.width;
                    if (!context) throw new Error("Could not get canvas context on retry");
                    
                    await page.render({
                        canvasContext: context,
                        viewport: viewport,
                        disableFontFace: true,
                        useSystemFonts: true
                    }).promise;
                    urls.push(canvas.toDataURL('image/jpeg', 0.8));
                } catch (safeRenderError) {
                    console.error(`Safe mode render also failed for page ${i} of ${file.name}.`, safeRenderError);
                }
            }
          }
          resolve({ urls: urls.length > 0 ? urls : null, pageCount });
        } catch (error) {
          if (error instanceof Error && error.name === 'PasswordException') {
            console.warn(`Could not process ${file.name} because it is password-protected.`);
          } else {
            console.error(`Error processing thumbnail for ${file.name}:`, error);
          }
          resolve({ urls: null, pageCount: 0 });
        }
      };
      fileReader.onerror = (error) => {
        console.error("FileReader error on thumbnail generation:", error);
        resolve({ urls: null, pageCount: 0 });
      };
      fileReader.readAsArrayBuffer(file);
    });
  }, []);

  const handleFileSelect = useCallback((files: File[]) => {
    const newJobs: ProcessJob[] = files.map(file => ({
        id: `${file.name}-${file.lastModified}-${file.size}`,
        file,
        fileName: file.name,
        fileSize: file.size,
        status: 'queued',
        conversionMode,
        pageRange: '',
        progressMessage: '',
        currentThumbnailIndex: 0,
        ...(conversionMode === 'excel-to-pdf' && { pdfOptions: { ...DEFAULT_PDF_OPTIONS } })
    }));

    setJobs(prevJobs => {
        const activeJobs = prevJobs.filter(job => job.status === 'queued' || job.status === 'processing');
        return [...activeJobs, ...newJobs];
    });
    
    setSelectedJobIds(new Set());

    if (conversionMode === 'pdf-to-excel') {
        newJobs.forEach(async (job) => {
            if (!job.file) return;
            const { urls, pageCount } = await generatePdfThumbnail(job.file);
            setJobs(prev => prev.map(j => 
                j.id === job.id ? { ...j, thumbnailUrls: urls, pdfPageCount: pageCount } : j
            ));
        });
    }
  }, [conversionMode, generatePdfThumbnail]);
  
  const handleClearCompleted = useCallback(() => {
    setJobs(prevJobs => prevJobs.filter(job => job.status === 'queued' || job.status === 'processing'));
    setSelectedJobIds(new Set());
  }, []);
  
  const parsePageRange = useCallback((rangeStr: string, maxPages: number): number[] => {
    if (!rangeStr.trim()) {
      return Array.from({ length: maxPages }, (_, i) => i + 1);
    }
    const pages = new Set<number>();
    const parts = rangeStr.split(',');
    for (const part of parts) {
      if (part.includes('-')) {
        const [start, end] = part.split('-').map(s => parseInt(s.trim(), 10));
        if (!isNaN(start) && !isNaN(end)) {
          for (let i = start; i <= end; i++) {
            if (i > 0 && i <= maxPages) pages.add(i);
          }
        }
      } else {
        const page = parseInt(part.trim(), 10);
        if (!isNaN(page) && page > 0 && page <= maxPages) {
          pages.add(page);
        }
      }
    }
    return Array.from(pages).sort((a, b) => a - b);
  }, []);

  const convertPdfToImages = useCallback(async (file: File, pageRange: string, onProgress: (message: string) => void): Promise<string[]> => {
    const fileReader = new FileReader();
    return new Promise((resolve, reject) => {
        fileReader.onload = async (event) => {
            if (!event.target?.result) {
                return reject(new Error("Failed to read file."));
            }
            try {
                // @ts-ignore
                const pdf = await pdfjsLib.getDocument({ data: event.target.result }).promise;
                const pagesToConvert = parsePageRange(pageRange, pdf.numPages);
                
                if (pagesToConvert.length === 0) {
                    onProgress("No pages selected for conversion.");
                    return resolve([]);
                }

                const images: string[] = [];
                for (let i = 0; i < pagesToConvert.length; i++) {
                    const pageNum = pagesToConvert[i];
                    onProgress(`Converting page ${i + 1} of ${pagesToConvert.length}...`);
                    const page = await pdf.getPage(pageNum);
                    const viewport = page.getViewport({ scale: 2.0 });
                    const canvas = document.createElement('canvas');
                    const context = canvas.getContext('2d');
                    canvas.height = viewport.height;
                    canvas.width = viewport.width;

                    if (!context) throw new Error("Could not get canvas context");
                    
                    await page.render({ canvasContext: context, viewport: viewport }).promise;
                    const dataUrl = canvas.toDataURL('image/jpeg');
                    images.push(dataUrl.split(',')[1]);
                }
                resolve(images);
            } catch (error) {
                console.error("Error converting PDF to images:", error);
                if (error instanceof Error && error.name === 'PasswordException') {
                    reject(new Error("The PDF is password-protected and cannot be processed."));
                } else {
                    reject(new Error("Failed to convert PDF to images. The file might be corrupted."));
                }
            }
        };
        fileReader.onerror = () => reject(new Error("Failed to read the PDF file."));
        fileReader.readAsArrayBuffer(file);
    });
  }, [parsePageRange]);

  const readExcelData = useCallback(async (file: File): Promise<{ sheetName: string, data: string[][] }[]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          // @ts-ignore
          const workbook = XLSX.read(data, { type: 'array' });
          const sheets: { sheetName: string, data: string[][] }[] = workbook.SheetNames.map((sheetName: string) => {
            const worksheet = workbook.Sheets[sheetName];
            // @ts-ignore
            const jsonData: string[][] = XLSX.utils.sheet_to_aoa(worksheet);
            return { sheetName, data: jsonData };
          });
          resolve(sheets.filter(s => s.data.length > 0));
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  }, []);

  const convertRawSheetsToRich = (rawSheets: {sheetName: string; data: string[][]}[]): SheetData[] => {
    return rawSheets.map(sheet => ({
        ...sheet,
        data: sheet.data.map(row => row.map(cellValue => ({ value: cellValue, style: {} })))
    }));
  };

  const processJob = useCallback(async (jobToProcess: ProcessJob) => {
    if (!jobToProcess || !jobToProcess.file) {
      setJobs(prev => prev.map(j => j.id === jobToProcess.id ? { ...j, status: 'error', errorMessage: 'File not available. Please re-upload to process.' } : j));
      return;
    }

    const jobId = jobToProcess.id;
    setActiveJobsCount(prev => prev + 1);

    try {
      if (jobToProcess.conversionMode === 'pdf-to-excel') {
          const images = await convertPdfToImages(jobToProcess.file, jobToProcess.pageRange || '', (message) => {
              setJobs(prev => prev.map(j => j.id === jobId ? { ...j, progressMessage: message } : j));
          });
          setJobs(prev => prev.map(j => j.id === jobId ? { ...j, progressMessage: `AI is analyzing ${images.length} page(s)...` } : j));
          const rawSheets = await extractDataFromPdfImages(images);
          const sheets = convertRawSheetsToRich(rawSheets);
          setJobs(prev => prev.map(j => j.id === jobId ? { ...j, status: 'success', extractedSheets: sheets, progressMessage: undefined } : j));
      } else { // excel-to-pdf
          setJobs(prev => prev.map(j => j.id === jobId ? { ...j, progressMessage: 'Reading spreadsheet...' } : j));
          const rawSheets = await readExcelData(jobToProcess.file);
          if (rawSheets.length === 0) throw new Error("Excel file contains no data.");
          const sheets = convertRawSheetsToRich(rawSheets);
          setJobs(prev => prev.map(j => j.id === jobId ? { ...j, status: 'success', extractedSheets: sheets, progressMessage: undefined } : j));
      }
    } catch (error) {
        const message = error instanceof Error ? error.message : "An unexpected error occurred.";
        setJobs(prev => prev.map(j => j.id === jobId ? { ...j, status: 'error', errorMessage: message, progressMessage: undefined } : j));
    } finally {
        setActiveJobsCount(prev => prev - 1);
    }
  }, [convertPdfToImages, readExcelData]);

  useEffect(() => {
    const availableSlots = MAX_CONCURRENT_JOBS - activeJobsCount;
    if (availableSlots <= 0) return;
    
    const jobsToStart = jobs.filter((job) => job.status === 'queued').slice(0, availableSlots);
    
    if (jobsToStart.length > 0) {
      setJobs(prev => prev.map(j => jobsToStart.some(js => js.id === j.id) ? { ...j, status: 'processing' } : j));
      jobsToStart.forEach(job => processJob(job));
    }
  }, [jobs, activeJobsCount, processJob]);
  
  const downloadPdfFromSheets = useCallback((sheets: SheetData[], fileName: string, options: PdfOptions, pageRange?: string) => {
    const totalSheets = sheets.length;
    const selectedSheetNumbers = parsePageRange(pageRange || '', totalSheets);
    const sheetsToProcess = sheets.filter((_, index) => selectedSheetNumbers.includes(index + 1));
    
    if (sheetsToProcess.length === 0) {
        throw new Error(`No sheets match the specified range: "${pageRange}". Please check your input.`);
    }

    // @ts-ignore
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: options.orientation });

    sheetsToProcess.forEach((sheet, index) => {
        if (index > 0) {
            doc.addPage();
        }
        const body = sheet.data.slice(1).map(row => row.map(cell => cell.value));
        const head = sheet.data[0].map(cell => cell.value);

        doc.text(sheet.sheetName, 14, 16);
        // @ts-ignore
        doc.autoTable({
            head: [head],
            body: body,
            startY: 20,
            theme: 'striped',
            styles: { fontSize: options.fontSize },
            headStyles: { fillColor: [79, 70, 229] },
            tableWidth: options.autoWidth ? 'auto' : 'wrap',
        });
    });
    
    doc.save(fileName);
  }, [parsePageRange]);

    const createStyledWorkbook = (sheets: SheetData[]) => {
      // @ts-ignore
      const wb = XLSX.utils.book_new();
      sheets.forEach(sheet => {
          const sanitizedSheetName = sheet.sheetName.replace(/[/\\?*:[\]]/g, '').substring(0, 31);
          // @ts-ignore
          const ws = {};
          let maxCols = 0;
          sheet.data.forEach((row, R) => {
              if (row.length > maxCols) maxCols = row.length;
              row.forEach((cell, C) => {
                  // @ts-ignore
                  const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
                  const cellObject: any = { v: cell.value, t: 's' };

                  const num = Number(cell.value);
                  if (!isNaN(num) && cell.value.trim() !== '') {
                      cellObject.v = num;
                      cellObject.t = 'n';
                  }

                  const sheetjsStyle: any = {};
                  const fontStyle: any = {};
                  const alignmentStyle: any = {};

                  if (cell.style?.bold) fontStyle.bold = true;
                  if (cell.style?.italic) fontStyle.italic = true;
                  if (cell.style?.underline) fontStyle.underline = true;
                  if (cell.style?.strikethrough) fontStyle.strike = true;
                  if (cell.style?.color) {
                      const rgb = cell.style.color.substring(1).toUpperCase();
                      fontStyle.color = { rgb: "FF" + rgb };
                  }
                  
                  if (cell.style?.align) alignmentStyle.horizontal = cell.style.align;

                  if (Object.keys(fontStyle).length > 0) sheetjsStyle.font = fontStyle;
                  if (Object.keys(alignmentStyle).length > 0) sheetjsStyle.alignment = alignmentStyle;

                  if (Object.keys(sheetjsStyle).length > 0) {
                      cellObject.s = sheetjsStyle;
                  }
                  
                  ws[cellRef] = cellObject;
              });
          });
          // @ts-ignore
          ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: Math.max(0, maxCols - 1), r: sheet.data.length - 1 } });
          // @ts-ignore
          XLSX.utils.book_append_sheet(wb, ws, sanitizedSheetName);
      });
      return wb;
    }


  const handleDownload = useCallback((job: ProcessJob) => {
    if (!job.extractedSheets) return;
    try {
      if (job.conversionMode === 'pdf-to-excel') {
          const wb = createStyledWorkbook(job.extractedSheets);
          const fileName = job.fileName.replace(/\.pdf$/i, '') + '.xlsx';
          // @ts-ignore
          XLSX.writeFile(wb, fileName);
      } else { // excel-to-pdf
          const fileName = job.fileName.replace(/\.(xlsx|xls)$/i, '') + '.pdf';
          downloadPdfFromSheets(job.extractedSheets, fileName, job.pdfOptions || DEFAULT_PDF_OPTIONS, job.pageRange);
      }
    } catch(error) {
        console.error("Failed to generate output file:", error);
        const message = error instanceof Error ? error.message : 'Failed to generate the output file.';
        setJobs(prev => prev.map(j => j.id === job.id ? { ...j, status: 'error', errorMessage: message } : j));
    }
  }, [downloadPdfFromSheets]);
  
  const executeDownloadAll = useCallback(async () => {
    setShowDownloadConfirm(false);
    // @ts-ignore
    const zip = new JSZip();
    const successfulJobs = jobs.filter(j => j.status === 'success' && j.extractedSheets && j.conversionMode === 'pdf-to-excel');
    if (successfulJobs.length === 0) return;

    for (const job of successfulJobs) {
        const wb = createStyledWorkbook(job.extractedSheets!);
        const fileName = job.fileName.replace(/\.pdf$/i, '') + '.xlsx';
        // @ts-ignore
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        zip.file(fileName, wbout, { compression: "DEFLATE", compressionOptions: { level: 9 } });
    }
    
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = 'converted_files.zip';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }, [jobs]);

  const handleRetry = useCallback((jobId: string) => {
    setJobs(prevJobs =>
        prevJobs.map(job =>
            job.id === jobId
            ? { ...job, status: 'queued', errorMessage: null, extractedSheets: null }
            : job
        )
    );
  }, []);

  const handlePageRangeChange = useCallback((jobId: string, range: string) => {
    setJobs(prevJobs =>
        prevJobs.map(job =>
            job.id === jobId ? { ...job, pageRange: range } : job
        )
    );
  }, []);

  const handleThumbnailCycle = useCallback((jobId: string, direction: 'next' | 'prev') => {
    setJobs(prevJobs => {
      return prevJobs.map(job => {
        if (job.id === jobId && job.thumbnailUrls && job.thumbnailUrls.length > 1) {
          const newIndex = direction === 'next'
            ? (job.currentThumbnailIndex! + 1) % job.thumbnailUrls.length
            : (job.currentThumbnailIndex! - 1 + job.thumbnailUrls.length) % job.thumbnailUrls.length;
          return { ...job, currentThumbnailIndex: newIndex };
        }
        return job;
      });
    });
  }, []);
  
  const filteredJobs = useMemo(() =>
    jobs.filter(job => job.fileName.toLowerCase().includes(searchQuery.toLowerCase()) && job.conversionMode === conversionMode),
    [jobs, searchQuery, conversionMode]
  );
  
  const totalPages = useMemo(() => Math.ceil(filteredJobs.length / ITEMS_PER_PAGE), [filteredJobs]);

  const paginatedJobs = useMemo(() => {
    const startIndex = (currentPage - 1) * ITEMS_PER_PAGE;
    return filteredJobs.slice(startIndex, startIndex + ITEMS_PER_PAGE);
  }, [filteredJobs, currentPage]);

  useEffect(() => {
    if (currentPage > 1 && paginatedJobs.length === 0) {
      setCurrentPage(prev => Math.max(1, prev - 1));
    }
  }, [currentPage, paginatedJobs]);

  useEffect(() => {
    if (currentPage !== 1) {
      setCurrentPage(1);
    }
  }, [searchQuery]);
  
  useEffect(() => {
      const jobIds = new Set(jobs.map(j => j.id));
      const newSelectedIds = new Set<string>();
      selectedJobIds.forEach(id => {
          if (jobIds.has(id)) {
              newSelectedIds.add(id);
          }
      });
      if (newSelectedIds.size !== selectedJobIds.size) {
          setSelectedJobIds(newSelectedIds);
      }
  }, [jobs, selectedJobIds]);

  const handleToggleSelection = useCallback((jobId: string) => {
    setSelectedJobIds(prev => {
        const newSet = new Set(prev);
        if (newSet.has(jobId)) {
            newSet.delete(jobId);
        } else {
            newSet.add(jobId);
        }
        return newSet;
    });
  }, []);

  const handleToggleSelectAll = useCallback(() => {
      const paginatedJobIds = new Set(paginatedJobs.map(j => j.id));
      const selectedOnPage = Array.from(selectedJobIds).filter(id => paginatedJobIds.has(id));

      if (selectedOnPage.length === paginatedJobs.length && paginatedJobs.length > 0) {
          setSelectedJobIds(prev => {
              const newSet = new Set(prev);
              paginatedJobIds.forEach(id => newSet.delete(id));
              return newSet;
          });
      } else {
          setSelectedJobIds(prev => {
              const newSet = new Set(prev);
              paginatedJobIds.forEach(id => newSet.add(id));
              return newSet;
          });
      }
  }, [paginatedJobs, selectedJobIds]);

  useEffect(() => {
    if (!selectAllCheckboxRef.current) return;
    const paginatedJobIds = new Set(paginatedJobs.map(j => j.id));
    if (paginatedJobIds.size === 0) {
        selectAllCheckboxRef.current.checked = false;
        selectAllCheckboxRef.current.indeterminate = false;
        return;
    }
    const selectedOnPageCount = Array.from(selectedJobIds).filter(id => paginatedJobIds.has(id)).length;
    
    if (selectedOnPageCount === 0) {
        selectAllCheckboxRef.current.checked = false;
        selectAllCheckboxRef.current.indeterminate = false;
    } else if (selectedOnPageCount === paginatedJobs.length) {
        selectAllCheckboxRef.current.checked = true;
        selectAllCheckboxRef.current.indeterminate = false;
    } else {
        selectAllCheckboxRef.current.checked = false;
        selectAllCheckboxRef.current.indeterminate = true;
    }
}, [selectedJobIds, paginatedJobs]);

  const handleDeleteSelected = useCallback(() => {
    setJobs(prev => prev.filter(job => !selectedJobIds.has(job.id)));
    setSelectedJobIds(new Set());
  }, [selectedJobIds]);

  const handleRetrySelected = useCallback(() => {
    const jobsToRetryIds = Array.from(selectedJobIds).filter(id => {
        const job = jobs.find(j => j.id === id);
        return job?.status === 'error' && job?.file;
    });

    if (jobsToRetryIds.length > 0) {
        setJobs(prevJobs =>
            prevJobs.map(job =>
                jobsToRetryIds.includes(job.id)
                ? { ...job, status: 'queued', errorMessage: null, extractedSheets: null }
                : job
            )
        );
    }
    setSelectedJobIds(new Set());
  }, [jobs, selectedJobIds]);

  const handleDownloadSelected = useCallback(async () => {
    const jobsToDownload = jobs.filter(j => selectedJobIds.has(j.id) && j.status === 'success' && j.extractedSheets);
    if (jobsToDownload.length === 0) {
        setSelectedJobIds(new Set());
        return;
    }

    // @ts-ignore
    const zip = new JSZip();
    let containsPdfToExcel = false;
    let containsExcelToPdf = false;
    
    for(const job of jobsToDownload) {
      if(job.conversionMode === 'pdf-to-excel') containsPdfToExcel = true;
      if(job.conversionMode === 'excel-to-pdf') containsExcelToPdf = true;
    }

    if (containsPdfToExcel) {
      const excelJobs = jobsToDownload.filter(j => j.conversionMode === 'pdf-to-excel');
      for (const job of excelJobs) {
          const wb = createStyledWorkbook(job.extractedSheets!);
          const fileName = job.fileName.replace(/\.pdf$/i, '') + '.xlsx';
          // @ts-ignore
          const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          zip.file(fileName, wbout, { compression: "DEFLATE", compressionOptions: { level: 9 } });
      }
    }
    
    if (containsExcelToPdf) {
      const pdfJobs = jobsToDownload.filter(j => j.conversionMode === 'excel-to-pdf');
      for (const job of pdfJobs) {
        const fileName = job.fileName.replace(/\.(xlsx|xls)$/i, '') + '.pdf';
        const sheets = job.extractedSheets!;
        const options = job.pdfOptions || DEFAULT_PDF_OPTIONS;
        const pageRange = job.pageRange;

        const totalSheets = sheets.length;
        const selectedSheetNumbers = parsePageRange(pageRange || '', totalSheets);
        const sheetsToProcess = sheets.filter((_, index) => selectedSheetNumbers.includes(index + 1));

        if (sheetsToProcess.length > 0) {
            // @ts-ignore
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF({ orientation: options.orientation });
            sheetsToProcess.forEach((sheet, index) => {
                if (index > 0) doc.addPage();
                const body = sheet.data.slice(1).map(row => row.map(cell => cell.value));
                const head = sheet.data[0].map(cell => cell.value);
                doc.text(sheet.sheetName, 14, 16);
                // @ts-ignore
                doc.autoTable({
                    head: [head], body: body, startY: 20, theme: 'striped',
                    styles: { fontSize: options.fontSize }, headStyles: { fillColor: [79, 70, 229] },
                    tableWidth: options.autoWidth ? 'auto' : 'wrap',
                });
            });
            const pdfBlob = doc.output('blob');
            zip.file(fileName, pdfBlob);
        }
      }
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = `converted_${jobsToDownload.length}_files.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    setSelectedJobIds(new Set());
  }, [jobs, selectedJobIds, parsePageRange]);

  const formatBytes = useCallback((bytes: number, decimals = 2) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  }, []);

  const getStatusIcon = useCallback((status: ProcessJobStatus) => {
    switch(status) {
        case 'processing': return <div className="flex-shrink-0 w-8 h-8 flex items-center justify-center" title="Processing..."><ProcessingIcon className="w-6 h-6 text-primary animate-spin" /></div>;
        case 'success': return <div className="flex-shrink-0 w-8 h-8 flex items-center justify-center" title="Success"><CheckCircleIcon className="w-6 h-6 text-green-500" /></div>;
        case 'error': return <div className="flex-shrink-0 w-8 h-8 flex items-center justify-center" title="Error"><XCircleIcon className="w-6 h-6 text-red-500" /></div>;
        case 'queued':
        default: return <div className="flex-shrink-0 w-8 h-8 flex items-center justify-center" title="Queued"><QueueListIcon className="w-6 h-6 text-text-secondary" /></div>;
    }
  }, []);

  const renderThumbnail = useCallback((job: ProcessJob) => {
    if (job.conversionMode === 'excel-to-pdf') {
       const options = job.pdfOptions || DEFAULT_PDF_OPTIONS;
       const fontSizeIndicator = { 8: 'S', 10: 'M', 12: 'L' }[options.fontSize];

       return (
           <div className="group relative flex-shrink-0 w-12 h-16 flex items-center justify-center bg-green-50 dark:bg-green-900/20 rounded-md overflow-hidden">
               <ExcelIcon className="w-8 h-8 text-green-600 dark:text-green-400"/>
               
               <div className="absolute inset-0 bg-black/10 dark:bg-black/30 p-1 flex flex-col justify-between text-white text-[10px] font-semibold pointer-events-none">
                   <div className="flex justify-start">
                       {job.pageRange && job.pageRange.trim() !== '' && (
                           <span className="bg-black/50 px-1.5 py-0.5 rounded" title={`Sheets: ${job.pageRange}`}>
                               {job.pageRange}
                           </span>
                       )}
                   </div>

                   <div className="flex justify-between items-center">
                       {options.orientation === 'p' 
                           ? <DocumentIconPortrait className="w-4 h-4" title="Portrait"/> 
                           : <DocumentIconLandscape className="w-4 h-4" title="Landscape"/>
                       }
                       <span className="bg-black/50 w-4 h-4 flex items-center justify-center rounded-full" title={`Font size: ${options.fontSize}pt`}>
                           {fontSizeIndicator}
                       </span>
                   </div>
               </div>
               
               <div className="absolute inset-0 bg-black/70 opacity-0 group-hover:opacity-100 transition-opacity p-1.5 text-white text-xs flex flex-col justify-center pointer-events-none">
                   <p className="font-bold mb-1 text-center">Settings</p>
                   <ul className="space-y-0.5 text-[10px]">
                       <li><span className="font-medium">Orientation:</span> {options.orientation === 'p' ? 'Portrait' : 'Landscape'}</li>
                       <li><span className="font-medium">Font:</span> {fontSizeIndicator} ({options.fontSize}pt)</li>
                       <li><span className="font-medium">Sheets:</span> {job.pageRange || 'All'}</li>
                       <li><span className="font-medium">Auto Width:</span> {options.autoWidth ? 'On' : 'Off'}</li>
                   </ul>
               </div>
           </div>
       );
    }
    
    if (!job.thumbnailUrls) {
      return (
          <div className="flex-shrink-0 w-12 h-16 flex items-center justify-center bg-border-color/50 rounded-md text-text-secondary">
              {job.thumbnailUrls === null 
                  ? <DocumentIcon className="w-8 h-8" />
                  : <ProcessingIcon className="w-6 h-6 animate-spin"/>
              }
          </div>
      );
    }

    return (
        <div className="group relative flex-shrink-0 w-12 h-16 bg-border-color rounded-md overflow-hidden">
            <img src={job.thumbnailUrls[job.currentThumbnailIndex || 0]} alt={`Preview of ${job.fileName}`} className="w-full h-full object-cover" />
            <div className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-between px-1">
                <button onClick={(e) => { e.stopPropagation(); handleThumbnailCycle(job.id, 'prev'); }} className="p-1 rounded-full bg-white/20 hover:bg-white/40 text-white">&lt;</button>
                <button onClick={(e) => { e.stopPropagation(); handleThumbnailCycle(job.id, 'next'); }} className="p-1 rounded-full bg-white/20 hover:bg-white/40 text-white">&gt;</button>
            </div>
            {job.pdfPageCount && job.pdfPageCount > 1 && (
                <div className="absolute bottom-1 right-1 bg-black/60 text-white text-[10px] px-1.5 py-0.5 rounded-full font-semibold">
                    {job.currentThumbnailIndex! + 1}/{job.thumbnailUrls.length}
                </div>
            )}
            {job.pdfPageCount && job.pdfPageCount > 5 && (
                <div className="absolute top-1 left-1 bg-primary/80 text-white text-[10px] px-1.5 py-0.5 rounded-full font-semibold">
                    {job.pdfPageCount}p
                </div>
            )}
        </div>
    );
  }, [handleThumbnailCycle]);
  
  const selectedJobs = useMemo(() => jobs.filter(j => selectedJobIds.has(j.id)), [jobs, selectedJobIds]);
  const canRetrySelected = useMemo(() => selectedJobs.some(j => j.status === 'error' && j.file), [selectedJobs]);
  const canDownloadSelected = useMemo(() => selectedJobs.some(j => j.status === 'success'), [selectedJobs]);

  const handleViewData = useCallback((job: ProcessJob) => {
    setCurrentlyViewing(job);
    setIsEditingData(false);
    if (job.extractedSheets) {
        setEditedData(JSON.parse(JSON.stringify(job.extractedSheets)));
    } else {
        setEditedData(null);
    }
  }, []);

  const handleCellChange = useCallback((sheetIndex: number, rowIndex: number, cellIndex: number, value: string) => {
    setEditedData(prevData => {
        if (!prevData) return null;
        const newSheets = JSON.parse(JSON.stringify(prevData));
        newSheets[sheetIndex].data[rowIndex][cellIndex].value = value;
        return newSheets;
    });
  }, []);
  
  const handleStyleChange = useCallback((sheetIndex: number, selection: Selection, style: Partial<CellStyle>) => {
    if (!selection) return;
    setEditedData(prevData => {
        if (!prevData) return null;
        const newSheets = JSON.parse(JSON.stringify(prevData));
        for (let r = selection.start.row; r <= selection.end.row; r++) {
            for (let c = selection.start.col; c <= selection.end.col; c++) {
                const currentStyle = newSheets[sheetIndex].data[r][c].style || {};
                newSheets[sheetIndex].data[r][c].style = { ...currentStyle, ...style };
            }
        }
        return newSheets;
    });
  }, []);

  const handleSaveChanges = useCallback(() => {
    if (!currentlyViewing || !editedData) return;
    setJobs(prevJobs =>
        prevJobs.map(job =>
            job.id === currentlyViewing.id
                ? { ...job, extractedSheets: editedData }
                : job
        )
    );
    setIsEditingData(false);
  }, [currentlyViewing, editedData]);
  
  const closeViewer = useCallback(() => {
    setCurrentlyViewing(null);
    setEditedData(null);
    setIsEditingData(false);
  }, []);

  const handleSavePdfOptions = useCallback((jobId: string, newOptions: PdfOptions) => {
    setJobs(prevJobs => prevJobs.map(j => j.id === jobId ? { ...j, pdfOptions: newOptions } : j));
    setEditingOptionsForJobId(null);
  }, []);

  const fileUploadConfig = {
    'pdf-to-excel': { accept: 'application/pdf', description: 'PDF' },
    'excel-to-pdf': { accept: '.xlsx, .xls, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel', description: 'Excel' },
  }[conversionMode];

  const renderContent = () => {
    if (jobs.length === 0 && searchQuery === '') {
        return (
          <div className="w-full flex-grow flex items-center justify-center">
            <FileUpload 
              onFileSelect={handleFileSelect} 
              disabled={activeJobsCount > 0} 
              accept={fileUploadConfig.accept}
              fileTypeDescription={fileUploadConfig.description}
            />
          </div>
        )
    }

    const hasCompletedJobs = jobs.some(j => j.status === 'success' || j.status === 'error');
    
    return (
        <div className="w-full max-w-4xl mx-auto flex flex-col space-y-4 flex-grow">
            <div className="space-y-4">
                <div className="relative">
                    <input 
                        type="text"
                        placeholder="Search files..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="w-full pl-10 pr-4 py-2 bg-secondary border border-border-color rounded-lg focus:ring-2 focus:ring-primary focus:outline-none transition-shadow"
                    />
                    <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                        <SearchIcon className="w-5 h-5 text-text-secondary" />
                    </div>
                </div>

                <div className="flex flex-col sm:flex-row items-center justify-between gap-2">
                    <div className="flex items-center gap-4">
                        <h2 className="text-xl font-bold text-text-main">Conversion Queue</h2>
                        <span className="text-sm font-medium text-text-secondary bg-border-color/50 px-2.5 py-1 rounded-full">{filteredJobs.length} files</span>
                    </div>
                    <div className="flex items-center gap-2 flex-wrap justify-center">
                        {jobs.length > 0 && (
                            <button onClick={handleClearAll} className="flex items-center justify-center gap-2 px-3 py-1.5 text-sm bg-secondary text-text-secondary font-semibold rounded-lg border border-border-color hover:bg-border-color hover:text-text-main transition-colors duration-200">
                                <BroomIcon className="w-4 h-4" />
                                Clear All
                            </button>
                        )}
                        {hasCompletedJobs && (
                             <button onClick={handleClearCompleted} className="flex items-center justify-center gap-2 px-3 py-1.5 text-sm bg-secondary text-text-secondary font-semibold rounded-lg border border-border-color hover:bg-border-color hover:text-text-main transition-colors duration-200">
                                <TrashIcon className="w-4 h-4" />
                                Clear Completed
                            </button>
                        )}
                        {successfulJobsCount > 1 && (
                            <button onClick={() => setShowDownloadConfirm(true)} className="flex items-center justify-center gap-2 px-4 py-1.5 text-sm bg-primary text-white font-semibold rounded-lg hover:bg-primary-hover transition-colors duration-200">
                                <DownloadIcon className="w-4 h-4" />
                                <span>Download All (.zip)</span>
                            </button>
                        )}
                    </div>
                </div>
            </div>
            
            <div className="space-y-3 flex-grow overflow-y-auto pr-2 -mr-2 flex flex-col">
              {filteredJobs.length > 0 ? (
                <>
                  {selectedJobIds.size > 0 && (
                    <div className="sticky top-0 z-20 bg-background/80 dark:bg-background/90 backdrop-blur-sm p-3 rounded-lg border border-border-color mb-3 shadow-md">
                      <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
                        <div className="flex items-center gap-3">
                          <input
                            type="checkbox"
                            ref={selectAllCheckboxRef}
                            onChange={handleToggleSelectAll}
                            className="h-5 w-5 rounded border-gray-300 text-primary focus:ring-primary"
                            aria-label="Select all on page"
                          />
                          <span className="font-bold text-lg text-text-main">{selectedJobIds.size} selected</span>
                           <button onClick={() => setSelectedJobIds(new Set())} className="flex items-center gap-2 px-3 py-1.5 text-sm font-semibold rounded-lg bg-secondary text-text-secondary hover:bg-border-color border border-border-color transition-colors">
                             <XCircleIcon className="w-5 h-5" />
                             <span>Clear</span>
                           </button>
                        </div>
                        <div className="flex items-center gap-2 flex-wrap justify-center">
                          <button
                            onClick={handleRetrySelected}
                            disabled={!canRetrySelected}
                            className="flex items-center gap-2 px-3 py-1.5 text-sm font-semibold rounded-lg bg-secondary text-text-secondary hover:bg-border-color border border-border-color transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                          >
                            <RetryIcon className="w-4 h-4" />
                            <span>Retry Selected</span>
                          </button>
                          <button
                            onClick={handleDeleteSelected}
                            className="flex items-center gap-2 px-3 py-1.5 text-sm font-semibold rounded-lg bg-red-600 text-white hover:bg-red-700 border border-transparent transition-colors"
                          >
                            <TrashIcon className="w-4 h-4" />
                            <span>Delete Selected</span>
                          </button>
                          <button
                            onClick={handleDownloadSelected}
                            disabled={!canDownloadSelected}
                            className="flex items-center gap-2 px-3 py-1.5 text-sm font-semibold rounded-lg bg-primary text-white hover:bg-primary-hover transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                          >
                            <DownloadIcon className="w-4 h-4" />
                            <span>Download Selected</span>
                          </button>
                        </div>
                      </div>
                    </div>
                  )}
                  {paginatedJobs.map(job => (
                    <div key={job.id} className={`p-4 rounded-lg border flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 transition-colors ${selectedJobIds.has(job.id) ? 'bg-primary/10 border-primary/50' : 'bg-secondary border-border-color'}`}>
                        <div className="flex items-center gap-4 w-full sm:w-auto flex-grow">
                            <input
                                type="checkbox"
                                className="h-5 w-5 rounded border-gray-300 text-primary focus:ring-primary flex-shrink-0 mt-1"
                                checked={selectedJobIds.has(job.id)}
                                onChange={() => handleToggleSelection(job.id)}
                                aria-labelledby={`job-name-${job.id}`}
                            />
                            {renderThumbnail(job)}
                            {getStatusIcon(job.status)}
                            <div className="flex-grow min-w-0">
                                <p id={`job-name-${job.id}`} className="font-semibold text-text-main truncate" title={job.fileName}>{job.fileName}</p>
                                <p className="text-sm text-text-secondary">
                                  {formatBytes(job.fileSize)} - 
                                  {job.status === 'processing' 
                                      ? <span className="text-primary font-medium">{job.progressMessage || 'Starting...'}</span> 
                                      : <span className="capitalize">{job.status}</span>
                                  }
                                </p>
                                {(job.conversionMode === 'pdf-to-excel' || job.conversionMode === 'excel-to-pdf') && (job.status === 'queued' || (job.status === 'error' && job.file)) ? (
                                    <div className="mt-1.5">
                                      <label htmlFor={`page-range-${job.id}`} className="text-xs font-medium text-text-secondary mr-2">{job.conversionMode === 'pdf-to-excel' ? 'Pages:' : 'Sheets:'}</label>
                                      <input 
                                        id={`page-range-${job.id}`}
                                        type="text" 
                                        placeholder="All" 
                                        value={job.pageRange || ''} 
                                        onChange={(e) => handlePageRangeChange(job.id, e.target.value)}
                                        className="w-32 px-2 py-0.5 text-sm bg-secondary border border-border-color rounded-md focus:ring-1 focus:ring-primary focus:outline-none"
                                      />
                                    </div>
                                ) : (job.conversionMode === 'pdf-to-excel' || job.conversionMode === 'excel-to-pdf') && job.status !== 'queued' && job.pageRange ? (
                                    <p className="text-xs text-text-secondary mt-1">{job.conversionMode === 'pdf-to-excel' ? 'Pages processed:' : 'Sheets processed:'} {job.pageRange}</p>
                                ) : null}
                            </div>
                        </div>
                        <div className="flex items-center gap-2 flex-shrink-0 self-end sm:self-center">
                            {job.conversionMode === 'excel-to-pdf' && (
                                <button
                                    onClick={() => setEditingOptionsForJobId(job.id)}
                                    className="p-1.5 rounded-md hover:bg-border-color text-text-secondary hover:text-text-main transition-colors"
                                    aria-label="PDF Output Settings"
                                >
                                    <Cog6ToothIcon className="w-5 h-5" />
                                </button>
                            )}

                            {job.status === 'success' && job.conversionMode === 'pdf-to-excel' && (
                                <>
                                    <button onClick={() => handleViewData(job)} className="px-3 py-1 text-sm font-medium text-primary hover:underline">View Data</button>
                                    <button onClick={() => handleDownload(job)} className="px-3 py-1 text-sm font-medium rounded-md bg-primary/10 text-primary hover:bg-primary/20">Download Excel</button>
                                </>
                            )}
                            {job.status === 'success' && job.conversionMode === 'excel-to-pdf' && (
                                <button onClick={() => handleDownload(job)} className="px-3 py-1 text-sm font-medium rounded-md bg-primary/10 text-primary hover:bg-primary/20">Download PDF</button>
                            )}
                             {job.status === 'error' && (
                                <div className="flex items-center gap-2">
                                    <p className="text-sm text-red-500" title={job.errorMessage ?? 'An unknown error occurred'}>Conversion Failed</p>
                                    <button onClick={() => handleRetry(job.id)} disabled={!job.file} className="flex items-center gap-1 px-3 py-1 text-sm font-medium rounded-md bg-secondary text-text-secondary hover:bg-border-color border border-border-color transition-colors disabled:opacity-50 disabled:cursor-not-allowed" title={!job.file ? "Re-upload this file to retry" : "Retry this file"} aria-label={`Retry conversion for ${job.fileName}`}><RetryIcon className="w-4 h-4" /><span>Retry</span></button>
                                </div>
                            )}
                        </div>
                    </div>
                  ))}
                </>
              ) : (
                <div className="text-center py-10 flex-grow flex items-center justify-center">
                  <p className="text-text-secondary">No files match your search.</p>
                </div>
              )}
              {paginatedJobs.length > 0 && totalPages > 1 && (
                  <div className="flex justify-center items-center gap-4 pt-4">
                      <button onClick={() => setCurrentPage(p => p-1)} disabled={currentPage === 1} className="px-3 py-1 rounded-md bg-secondary border border-border-color disabled:opacity-50 flex items-center gap-2"><ChevronLeftIcon className="w-4 h-4"/> Prev</button>
                      <span className="text-sm font-medium text-text-secondary">Page {currentPage} of {totalPages}</span>
                      <button onClick={() => setCurrentPage(p => p+1)} disabled={currentPage === totalPages} className="px-3 py-1 rounded-md bg-secondary border border-border-color disabled:opacity-50 flex items-center gap-2">Next <ChevronRightIcon className="w-4 h-4"/></button>
                  </div>
              )}
            </div>
             <div className="pt-4">
                <FileUpload 
                  onFileSelect={handleFileSelect} 
                  disabled={activeJobsCount > 0} 
                  accept={fileUploadConfig.accept}
                  fileTypeDescription={fileUploadConfig.description}
                />
            </div>
        </div>
    );
  };
  
  const TABS: { id: ConversionMode, name: string }[] = [
      { id: 'pdf-to-excel', name: 'PDF to Excel' },
      { id: 'excel-to-pdf', name: 'Excel to PDF' },
  ];

  const headerConfig = {
      'pdf-to-excel': {
          title: "AI PDF to Excel Converter",
          description: "Pull data straight from PDFs into Excel spreadsheets in seconds. Powered by Gemini."
      },
      'excel-to-pdf': {
          title: "Excel to PDF Converter",
          description: "Make EXCEL spreadsheets easy to read by converting them to PDF."
      }
  }[conversionMode];
  
  const jobToEditOptions = useMemo(() => jobs.find(j => j.id === editingOptionsForJobId), [jobs, editingOptionsForJobId]);

  return (
    <div className="min-h-screen w-full bg-background text-text-main flex flex-col items-center p-4 sm:p-6 lg:p-8 transition-colors duration-300">
       <header className="w-full max-w-6xl mx-auto flex justify-between items-start mb-8">
        <div className="text-left">
            <h1 className="text-4xl sm:text-5xl font-extrabold tracking-tight">
              {headerConfig.title}
            </h1>
            <p className="mt-4 max-w-2xl text-lg text-text-secondary">
              {headerConfig.description}
            </p>
        </div>
        <button
          onClick={toggleTheme}
          className="flex-shrink-0 p-2 rounded-full bg-secondary text-text-secondary hover:text-text-main hover:bg-border-color transition-colors"
          aria-label="Toggle dark mode"
        >
          {theme === 'dark' ? <SunIcon className="w-6 h-6" /> : <MoonIcon className="w-6 h-6" />}
        </button>
      </header>
      
      <div className="w-full max-w-4xl mx-auto mb-8">
        <div className="border-b border-border-color">
            <nav className="-mb-px flex space-x-6" aria-label="Tabs">
                {TABS.map((tab) => (
                    <button
                        key={tab.id}
                        onClick={() => switchConversionMode(tab.id)}
                        className={`${
                            conversionMode === tab.id
                                ? 'border-primary text-primary'
                                : 'border-transparent text-text-secondary hover:text-text-main hover:border-gray-300 dark:hover:border-gray-600'
                        } whitespace-nowrap py-4 px-1 border-b-2 font-medium text-lg transition-colors`}
                    >
                        {tab.name}
                    </button>
                ))}
            </nav>
        </div>
      </div>


      <main className="w-full flex-grow flex items-center justify-center">
        {renderContent()}
      </main>
      <footer className="text-center py-4 mt-8">
          <p className="text-sm text-text-secondary">A world-class app created by a senior frontend React engineer.</p>
      </footer>
      
      {/* Modals */}
      {showDownloadConfirm && (
        <div className="fixed inset-0 bg-background/80 backdrop-blur-sm z-50 flex items-center justify-center p-4" onClick={() => setShowDownloadConfirm(false)}>
          <div className="w-full max-w-md bg-secondary rounded-xl border border-border-color shadow-2xl p-6 text-center" onClick={e => e.stopPropagation()}>
             <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-full bg-primary/10">
                <DownloadIcon className="h-6 w-6 text-primary" aria-hidden="true" />
             </div>
             <h3 className="mt-4 text-lg font-semibold text-text-main">Confirm Download</h3>
             <p className="mt-2 text-sm text-text-secondary">
               You are about to download a ZIP archive containing {successfulJobsCount} converted files.
             </p>
             <div className="mt-6 flex justify-center gap-4">
               <button type="button" onClick={() => setShowDownloadConfirm(false)} className="px-6 py-2 text-sm font-semibold rounded-lg bg-secondary text-text-secondary hover:bg-border-color border border-border-color transition-colors">Cancel</button>
               <button type="button" onClick={executeDownloadAll} className="px-6 py-2 text-sm font-semibold rounded-lg bg-primary text-white hover:bg-primary-hover transition-colors">Download</button>
             </div>
          </div>
        </div>
      )}
      {currentlyViewing && (
          <div className="fixed inset-0 bg-background/80 backdrop-blur-sm z-50 flex items-center justify-center p-4" onClick={closeViewer}>
              <div className="w-full max-w-5xl h-[90vh] bg-secondary rounded-xl border border-border-color shadow-2xl flex flex-col" onClick={e => e.stopPropagation()}>
                  <div className="flex items-center justify-between p-4 border-b border-border-color">
                      <h3 className="font-bold text-text-main truncate" title={currentlyViewing.fileName}>
                        {isEditingData ? "Editing" : "Viewing"}: {currentlyViewing.fileName}
                      </h3>
                        <div className="flex items-center gap-4">
                          {isEditingData ? (
                            <>
                                <button onClick={() => setIsEditingData(false)} className="px-4 py-1.5 text-sm bg-secondary text-text-secondary font-semibold rounded-lg border border-border-color hover:bg-border-color hover:text-text-main transition-colors duration-200">Cancel</button>
                                <button onClick={handleSaveChanges} className="flex items-center gap-2 px-4 py-1.5 text-sm bg-primary text-white font-semibold rounded-lg hover:bg-primary-hover transition-colors duration-200"><SaveIcon className="w-4 h-4" />Save Changes</button>
                            </>
                          ) : (
                            <button onClick={() => setIsEditingData(true)} className="flex items-center gap-2 px-4 py-1.5 text-sm bg-secondary text-text-secondary font-semibold rounded-lg border border-border-color hover:bg-border-color hover:text-text-main transition-colors duration-200"><PencilIcon className="w-4 h-4" />Edit Data</button>
                          )}
                          <button onClick={closeViewer} className="p-1 rounded-full hover:bg-border-color text-text-secondary text-2xl leading-none">&times;</button>
                      </div>
                  </div>
                  <div className="flex-grow overflow-hidden">
                      {editedData ? (<DataTable sheets={editedData} editable={isEditingData} onCellChange={handleCellChange} onStyleChange={handleStyleChange} />) : (<p className="text-center p-8 text-text-secondary">No data to display.</p>)}
                  </div>
              </div>
          </div>
      )}
      {jobToEditOptions && <PdfOptionsModal job={jobToEditOptions} onSave={handleSavePdfOptions} onClose={() => setEditingOptionsForJobId(null)} />}
    </div>
  );
};

export default App;