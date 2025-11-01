
export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  align?: 'left' | 'center' | 'right';
}

export interface CellData {
  value: string;
  style?: CellStyle;
}

export interface SheetData {
  sheetName: string;
  data: CellData[][];
}

export type ProcessJobStatus = 'queued' | 'processing' | 'success' | 'error';

export type ConversionMode = 'pdf-to-excel' | 'excel-to-pdf';

export interface PdfOptions {
  orientation: 'p' | 'l';
  fontSize: 8 | 10 | 12;
  autoWidth: boolean;
}

export interface ProcessJob {
  id: string;
  file?: File;
  fileName: string;
  fileSize: number;
  status: ProcessJobStatus;
  conversionMode: ConversionMode;
  extractedSheets?: SheetData[] | null;
  errorMessage?: string | null;
  thumbnailUrls?: string[] | null;
  currentThumbnailIndex?: number;
  pdfPageCount?: number;
  pageRange?: string;
  progressMessage?: string;
  pdfOptions?: PdfOptions;
}