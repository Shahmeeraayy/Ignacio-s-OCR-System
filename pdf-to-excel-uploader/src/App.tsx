import React, { useState, useEffect, useRef } from 'react';

interface FileItem {
  id: string;
  file: File;
  size: string;
}

interface StatusState {
  type: 'idle' | 'loading' | 'success' | 'error';
  message: string;
  details?: string;
}

interface PricingInputs {
  euroRate: string;
  marginPercent: string;
}

interface VendorOption {
  id: string;
  label: string;
}

interface VendorsResponse {
  ok?: boolean;
  default_vendor?: string;
  vendors?: VendorOption[];
}

const App: React.FC = () => {
  const [files, setFiles] = useState<FileItem[]>([]);
  const [isServerHealthy, setIsServerHealthy] = useState<boolean>(true);
  const [status, setStatus] = useState<StatusState>({ type: 'idle', message: '' });
  const [isDragOver, setIsDragOver] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [pricing, setPricing] = useState<PricingInputs>({ euroRate: '', marginPercent: '' });
  const [vendors, setVendors] = useState<VendorOption[]>([{ id: 'netskope', label: 'Netskope' }]);
  const [selectedVendor, setSelectedVendor] = useState<string>('netskope');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isVisible, setIsVisible] = useState(false);

  // Page load animation
  useEffect(() => {
    setTimeout(() => setIsVisible(true), 100);
    checkServerHealth();
    loadVendors();
  }, []);

  const checkServerHealth = async () => {
    try {
      const response = await fetch('/api/health');
      setIsServerHealthy(response.ok);
    } catch (error) {
      setIsServerHealthy(false);
    }
  };

  const loadVendors = async () => {
    try {
      const response = await fetch('/api/vendors');
      if (!response.ok) {
        return;
      }
      const payload = (await response.json()) as VendorsResponse;
      const availableVendors = Array.isArray(payload.vendors) ? payload.vendors : [];
      if (payload.ok && availableVendors.length > 0) {
        setVendors(availableVendors);
        const hasDefaultVendor = availableVendors.some(
          (vendor) => vendor.id === payload.default_vendor
        );
        setSelectedVendor(
          hasDefaultVendor && payload.default_vendor
            ? payload.default_vendor
            : availableVendors[0].id
        );
      }
    } catch (error) {
      // Keep local fallback vendor list when the endpoint is unavailable.
    }
  };

  const parseLocalizedNumber = (raw: string): number => {
    const compact = raw.trim().replace(/\s+/g, '');
    if (!compact) {
      return Number.NaN;
    }

    const hasComma = compact.includes(',');
    const hasDot = compact.includes('.');
    if (hasComma && hasDot) {
      if (compact.lastIndexOf(',') > compact.lastIndexOf('.')) {
        return Number(compact.replace(/\./g, '').replace(',', '.'));
      }
      return Number(compact.replace(/,/g, ''));
    }
    if (hasComma) {
      const parts = compact.split(',');
      if (parts.length > 2) {
        return Number(parts.join(''));
      }
      const [head, tail] = parts;
      if (head && tail && tail.length === 3 && head.length <= 3 && !head.startsWith('0')) {
        return Number(`${head}${tail}`);
      }
      return Number(compact.replace(',', '.'));
    }
    return Number(compact);
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(event.target.files || []);
    addFiles(selectedFiles);
    event.target.value = '';
  };

  const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragOver(false);
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragOver(false);
    const droppedFiles = Array.from(event.dataTransfer.files || []);
    addFiles(droppedFiles);
  };

  const addFiles = (newFiles: File[]) => {
    const validPdfFiles = newFiles.filter(file => 
      file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')
    );

    if (validPdfFiles.length !== newFiles.length) {
      setStatus({
        type: 'error',
        message: 'Some files were skipped. Only PDF files are allowed.'
      });
      setTimeout(() => {
        if (status.type === 'error') {
          setStatus({ type: 'idle', message: '' });
        }
      }, 3000);
    }

    const fileItems: FileItem[] = validPdfFiles.map(file => ({
      id: `${file.name}-${file.size}-${file.lastModified}`,
      file,
      size: formatFileSize(file.size)
    }));

    setFiles(prev => {
      const existingIds = new Set(prev.map(file => file.id));
      const uniqueNewItems = fileItems.filter(file => !existingIds.has(file.id));
      const skippedCount = fileItems.length - uniqueNewItems.length;
      if (skippedCount > 0) {
        setStatus({
          type: 'error',
          message: `${skippedCount} duplicate file(s) were skipped.`
        });
      }
      return [...prev, ...uniqueNewItems];
    });
  };

  const removeFile = (id: string) => {
    setFiles(prev => prev.filter(file => file.id !== id));
  };

  const handleSubmit = async () => {
    if (files.length === 0 || !isServerHealthy) return;
    const euroInput = pricing.euroRate.trim();
    const marginInput = pricing.marginPercent.trim();
    if (!euroInput) {
      setStatus({
        type: 'error',
        message: 'Euro Rate is required.'
      });
      return;
    }
    if (!marginInput) {
      setStatus({
        type: 'error',
        message: 'Margin (%) is required.'
      });
      return;
    }
    const parsedEuroRate = parseLocalizedNumber(pricing.euroRate);
    const parsedMarginPercent = parseLocalizedNumber(pricing.marginPercent);
    if (!Number.isFinite(parsedEuroRate) || parsedEuroRate <= 0) {
      setStatus({
        type: 'error',
        message: 'Please provide a valid Euro Rate greater than 0 (use . or , as decimal separator).'
      });
      return;
    }
    if (!Number.isFinite(parsedMarginPercent)) {
      setStatus({
        type: 'error',
        message: 'Please provide a valid Margin (%) using . or , as decimal separator.'
      });
      return;
    }

    setIsProcessing(true);
    setStatus({ type: 'loading', message: '' });

    try {
      const formData = new FormData();
      files.forEach(fileItem => {
        formData.append('pdf', fileItem.file);
      });
      formData.append('strict', 'true');
      formData.append('template_only', 'true');
      formData.append('ocr_mode', 'off');
      formData.append('euro_rate', String(parsedEuroRate));
      formData.append('margin_percent', String(parsedMarginPercent));
      formData.append('vendor', selectedVendor);

      const response = await fetch('/api/extract-template', {
        method: 'POST',
        body: formData,
      });

      if (response.ok) {
        const processedFiles = response.headers.get('X-Processed-Files') || '0';
        const rowsWritten = response.headers.get('X-Rows-Written') || '0';

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'filled_template.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        setStatus({
          type: 'success',
          message: `Successfully processed ${processedFiles} files`,
          details: `${rowsWritten} rows written`
        });

        setFiles([]);
      } else {
        let errorMessage = 'Processing failed';
        try {
          const errorData = await response.json();
          errorMessage = errorData.error || errorData.message || errorMessage;
        } catch {}
        
        setStatus({
          type: 'error',
          message: errorMessage
        });
      }
    } catch (error) {
      setStatus({
        type: 'error',
        message: 'Network error. Please check your connection and try again.'
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const isSubmitDisabled = files.length === 0 || !isServerHealthy || isProcessing;

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-900 via-gray-800 to-gray-900 flex items-center justify-center p-4">
      <div className={`w-full max-w-2xl transform transition-all duration-500 ${
        isVisible ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-5'
      }`}>
        
        {/* Title Section */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-white mb-3 tracking-tight">
            PDF to Excel Extractor
          </h1>
          <p className="text-gray-400 text-lg">
            Upload multiple PDF files and generate a structured Excel sheet instantly.
          </p>
        </div>

        {/* Main Card */}
        <div className="bg-gray-800/60 backdrop-blur-xl rounded-3xl shadow-2xl border border-gray-700 p-8 transition-all duration-300 hover:shadow-cyan-500/10 hover:border-gray-600">
          
          {/* Server Health Warning */}
          {!isServerHealthy && (
            <div className="mb-6 p-4 bg-red-500/10 border border-red-500/30 rounded-xl">
              <div className="flex items-center justify-center space-x-2">
                <svg className="h-5 w-5 text-red-400" fill="currentColor" viewBox="0 0 20 20">
                  <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
                </svg>
                <p className="text-sm text-red-400 font-medium">
                  Server is currently unavailable. Please try again later.
                </p>
              </div>
            </div>
          )}

          {/* Upload Area */}
          <div className="mb-6">
            <div
            className={`relative transition-all duration-200 ${
                !isServerHealthy ? 'opacity-50 pointer-events-none' : ''
              } ${isProcessing ? 'opacity-75' : ''}`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <div
                className={`
                  border-2 border-dashed rounded-2xl p-12 text-center transition-all duration-300 cursor-pointer
                  ${isDragOver 
                    ? 'border-cyan-400 bg-cyan-500/10 shadow-lg shadow-cyan-500/20' 
                    : 'border-gray-600 bg-gray-900/40 hover:border-gray-500 hover:bg-gray-900/60 hover:shadow-md'
                  }
                `}
                onClick={() => fileInputRef.current?.click()}
              >
                <div className="mb-6">
                  <div className="mx-auto h-16 w-16 text-gray-500 transition-transform duration-200 group-hover:scale-110">
                    <svg className="w-full h-full" fill="none" stroke="currentColor" viewBox="0 0 48 48">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M28 8H12a2 2 0 00-2 2v28a2 2 0 002 2h24a2 2 0 002-2V16L28 8z" />
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M28 8v8h8" />
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M16 24v8m8-8v8m8-8v8" />
                    </svg>
                  </div>
                </div>
                <div className="text-gray-300 text-lg mb-1 font-medium">
                  {isDragOver ? 'Drop files here' : 'Drag & Drop PDF files here'}
                </div>
                <div className="text-gray-500 text-sm">
                  or <span className="text-cyan-400 font-medium hover:text-cyan-300 transition-colors">click to browse</span>
                </div>
                <div className="text-xs text-gray-600 mt-2">
                  PDF files only • Multiple files supported
                </div>
              </div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".pdf"
                multiple
                onChange={handleFileSelect}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                disabled={!isServerHealthy}
              />
            </div>
          </div>

          {/* Extraction Inputs */}
          <div className="mb-6 grid grid-cols-1 md:grid-cols-3 gap-4">
            <div>
              <label className="block text-sm font-semibold text-gray-300 mb-2">
                Vendor
              </label>
              <select
                value={selectedVendor}
                onChange={(event) => setSelectedVendor(event.target.value)}
                disabled={!isServerHealthy || isProcessing}
                className="w-full rounded-xl bg-gray-900/50 border border-gray-700 text-gray-100 px-4 py-3 focus:outline-none focus:ring-2 focus:ring-cyan-500 focus:border-transparent"
              >
                {vendors.map((vendor) => (
                  <option key={vendor.id} value={vendor.id} className="bg-gray-900 text-gray-100">
                    {vendor.label}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-300 mb-2">
                Euro Rate
              </label>
              <input
                type="text"
                inputMode="decimal"
                placeholder="e.g. 1.17 or 1,17"
                value={pricing.euroRate}
                onChange={(event) =>
                  setPricing((prev) => ({ ...prev, euroRate: event.target.value }))
                }
                disabled={!isServerHealthy || isProcessing}
                className="w-full rounded-xl bg-gray-900/50 border border-gray-700 text-gray-100 px-4 py-3 focus:outline-none focus:ring-2 focus:ring-cyan-500 focus:border-transparent"
              />
            </div>
            <div>
              <label className="block text-sm font-semibold text-gray-300 mb-2">
                Margin (%)
              </label>
              <input
                type="text"
                inputMode="decimal"
                placeholder="e.g. 10 or 10,5"
                value={pricing.marginPercent}
                onChange={(event) =>
                  setPricing((prev) => ({ ...prev, marginPercent: event.target.value }))
                }
                disabled={!isServerHealthy || isProcessing}
                className="w-full rounded-xl bg-gray-900/50 border border-gray-700 text-gray-100 px-4 py-3 focus:outline-none focus:ring-2 focus:ring-cyan-500 focus:border-transparent"
              />
            </div>
          </div>

          {/* File List */}
          {files.length > 0 && (
            <div className="mb-6">
              <div className="flex items-center justify-between mb-4">
                <h3 className="text-sm font-semibold text-gray-300">
                  Selected Files ({files.length})
                </h3>
                <div className="text-xs text-gray-500 bg-gray-900/50 px-2 py-1 rounded-full">
                  {files.reduce((acc, f) => acc + f.file.size, 0) > 1024 * 1024 
                    ? `${(files.reduce((acc, f) => acc + f.file.size, 0) / (1024 * 1024)).toFixed(1)} MB total`
                    : `${(files.reduce((acc, f) => acc + f.file.size, 0) / 1024).toFixed(0)} KB total`
                  }
                </div>
              </div>
              <div className="space-y-2 max-h-72 overflow-y-auto custom-scrollbar">
                {files.map((fileItem) => (
                  <div
                    key={fileItem.id}
                    className="flex items-center justify-between p-3 bg-gray-900/50 rounded-xl border border-gray-700 hover:border-gray-600 transition-all duration-200 hover:shadow-md"
                  >
                    <div className="flex items-center min-w-0 flex-1">
                      <svg className="h-5 w-5 text-red-400 mr-3 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M4 4a2 2 0 012-2h4.586A2 2 0 0112 2.586L15.414 6A2 2 0 0116 7.414V16a2 2 0 01-2 2H6a2 2 0 01-2-2V4z" clipRule="evenodd" />
                      </svg>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-medium text-gray-200 truncate">
                          {fileItem.file.name}
                        </p>
                        <p className="text-xs text-gray-500">
                          {fileItem.size}
                        </p>
                      </div>
                    </div>
                    <button
                      onClick={() => removeFile(fileItem.id)}
                      className="ml-3 text-gray-500 hover:text-red-400 transition-all duration-200 hover:scale-110"
                    >
                      <svg className="h-5 w-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                      </svg>
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Submit Button */}
          <button
            onClick={handleSubmit}
            disabled={isSubmitDisabled}
            className={`
              w-full py-4 px-6 rounded-xl font-semibold text-lg transition-all duration-300 relative overflow-hidden
              ${isSubmitDisabled
                ? 'bg-gray-700 text-gray-500 cursor-not-allowed'
                : 'bg-gradient-to-r from-blue-600 to-cyan-600 text-white hover:from-blue-500 hover:to-cyan-500 hover:shadow-lg hover:shadow-blue-500/25 transform hover:-translate-y-0.5 active:translate-y-0'
              }
              ${isProcessing ? 'animate-pulse' : ''}
            `}
          >
            <div className="flex items-center justify-center space-x-2">
              {isProcessing && (
                <svg className="animate-spin h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
              )}
              <span>{isProcessing ? 'Processing...' : 'Generate Excel'}</span>
            </div>
          </button>

          {/* Status Messages */}
          {status.type !== 'idle' && (
            <div className="mt-6 space-y-3">
              {status.type === 'loading' && (
                <div className="p-5 bg-cyan-500/10 border border-cyan-500/30 rounded-xl backdrop-blur-sm">
                  <div className="flex items-center space-x-3">
                    <svg className="animate-spin h-6 w-6 text-cyan-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                    </svg>
                    <div>
                      <p className="text-sm font-medium text-cyan-300">Processing files. Please wait...</p>
                      <p className="text-xs text-gray-500 mt-1">This may take a moment depending on file size</p>
                    </div>
                  </div>
                </div>
              )}

              {status.type === 'success' && (
                <div className="p-5 bg-green-500/10 border border-green-500/30 rounded-xl backdrop-blur-sm">
                  <div className="flex items-start space-x-3">
                    <svg className="h-6 w-6 text-green-400 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    <div className="flex-1">
                      <p className="text-sm font-medium text-green-300 mb-1">✔ {status.message}</p>
                      {status.details && (
                        <p className="text-sm text-green-300 mb-2">✔ {status.details}</p>
                      )}
                      <p className="text-xs text-gray-400">Your Excel file has been downloaded.</p>
                    </div>
                  </div>
                </div>
              )}

              {status.type === 'error' && (
                <div className="p-5 bg-red-500/10 border border-red-500/30 rounded-xl backdrop-blur-sm">
                  <div className="flex items-start space-x-3">
                    <svg className="h-6 w-6 text-red-400 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                      <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
                    </svg>
                    <div className="flex-1">
                      <p className="text-sm font-medium text-red-300">
                        {status.message.startsWith('Error:') ? status.message : `Error: ${status.message}`}
                      </p>
                      <p className="text-xs text-gray-500 mt-1">Please try again or contact support if the issue persists.</p>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="mt-8 text-center">
          <p className="text-xs text-gray-600">
            Your files are processed securely and never stored on our servers
          </p>
        </div>
      </div>

                <style>{`
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: rgba(31, 41, 55, 0.5);
          border-radius: 3px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: rgba(107, 114, 128, 0.5);
          border-radius: 3px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: rgba(156, 163, 175, 0.7);
        }
      `}</style>
    </div>
  );
};

export default App;
