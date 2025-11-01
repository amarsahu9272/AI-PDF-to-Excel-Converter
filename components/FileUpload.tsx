import React, { useState, useCallback, useRef } from 'react';
import { UploadIcon, ErrorIcon, XMarkIcon } from './icons';

interface FileUploadProps {
  onFileSelect: (files: File[]) => void;
  disabled: boolean;
  accept: string;
  fileTypeDescription: string;
}

const FileUpload: React.FC<FileUploadProps> = ({ onFileSelect, disabled, accept, fileTypeDescription }) => {
  const [isDragging, setIsDragging] = useState(false);
  const [validationError, setValidationError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const acceptedTypes = accept.split(',').map(t => t.trim());

  const processFiles = (fileList: FileList | null) => {
    if (!fileList) return;
    
    const files = Array.from(fileList);
    const validFiles: File[] = [];
    const invalidFiles: File[] = [];

    const isFileValid = (file: File) => {
      // Check MIME type and file extension
      const extension = '.' + file.name.split('.').pop()?.toLowerCase();
      return acceptedTypes.includes(file.type) || acceptedTypes.includes(extension);
    }
    
    for (const file of files) {
        if (isFileValid(file)) {
            validFiles.push(file);
        } else {
            invalidFiles.push(file);
        }
    }
    
    if (invalidFiles.length > 0) {
      const invalidFileNames = invalidFiles.map(f => f.name).slice(0, 3).join(', ');
      const additionalFilesCount = invalidFiles.length > 3 ? ` and ${invalidFiles.length - 3} more` : '';
      setValidationError(`Ignored non-${fileTypeDescription} files: ${invalidFileNames}${additionalFilesCount}`);
    } else {
      setValidationError(null);
    }

    if (validFiles.length > 0) {
      onFileSelect(validFiles);
    }
  };

  const handleDragEnter = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (!disabled) setIsDragging(true);
  }, [disabled]);

  const handleDragLeave = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const handleDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    if (disabled) return;
    setValidationError(null);
    processFiles(e.dataTransfer.files);
  }, [disabled, onFileSelect]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setValidationError(null);
    processFiles(e.target.files);
    e.target.value = '';
  };
  
  const handleClick = () => {
    fileInputRef.current?.click();
  };

  const dragDropClasses = isDragging 
    ? 'border-primary bg-secondary' 
    : validationError 
    ? 'border-red-500 hover:border-red-600'
    : 'border-border-color hover:border-primary';

  return (
    <div className="w-full max-w-lg mx-auto">
      <div
        className={`p-8 text-center border-2 border-dashed rounded-xl transition-all duration-300 cursor-pointer ${dragDropClasses}`}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
        onClick={handleClick}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept={accept}
          className="hidden"
          onChange={handleFileChange}
          disabled={disabled}
          multiple
        />
        <div className="flex flex-col items-center justify-center space-y-4 text-text-secondary">
          <UploadIcon className="w-12 h-12" />
          <p className="text-lg font-semibold text-text-main">
            Drag & drop your {fileTypeDescription} files here
          </p>
          <p>or <span className="font-semibold text-primary">click to browse</span></p>
          <p className="text-xs">Multiple {fileTypeDescription} files are supported</p>
        </div>
      </div>

      {validationError && (
        <div className="mt-4 p-3 bg-red-500/10 border border-red-500/30 rounded-lg text-sm flex items-start justify-between gap-4">
            <div className="flex items-center gap-2">
                <ErrorIcon className="w-5 h-5 text-red-500 flex-shrink-0 mt-0.5"/>
                <p className="text-red-700 dark:text-red-300">{validationError}</p>
            </div>
            <button 
              onClick={() => setValidationError(null)} 
              className="p-1 rounded-full text-red-600 hover:bg-red-500/20 transition-colors flex-shrink-0"
              aria-label="Dismiss"
            >
                <XMarkIcon className="w-5 h-5" />
            </button>
        </div>
      )}
    </div>
  );
};

export default FileUpload;