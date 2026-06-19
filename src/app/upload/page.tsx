'use client';

import React, { useState, useRef } from 'react';
import { useRouter } from 'next/navigation';
import type { OfficerRecord } from '../../types';

export default function UploadPage() {
  const [files, setFiles] = useState<File[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadingStep, setLoadingStep] = useState('');
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const router = useRouter();

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const selectedFiles = Array.from(e.target.files);
      setFiles((prev) => [...prev, ...selectedFiles]);
      setError(null);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files) {
      const droppedFiles = Array.from(e.dataTransfer.files).filter((file) => {
        const ext = file.name.split('.').pop()?.toLowerCase();
        return file.type.startsWith('image/') || ext === 'xlsx' || ext === 'xls';
      });
      setFiles((prev) => [...prev, ...droppedFiles]);
      setError(null);
    }
  };

  const removeFile = (index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const handleUpload = async () => {
    if (files.length === 0) {
      setError('कृपया कम से कम एक Excel (.xlsx) फाइल अपलोड करें।');
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const allRecords: OfficerRecord[] = [];
      let meta = { eventName: '', district: '', dutyDateFrom: '', dutyDateTo: '' };

      for (const file of files) {
        const ext = file.name.split('.').pop()?.toLowerCase();

        if (ext === 'xlsx' || ext === 'xls') {
          // ✅ Excel path — no API needed, instant, 100% free
          setLoadingStep('Excel फाइल पढ़ी जा रही है...');
          const { parseExcelFile } = await import('@/lib/parseExcel');
          const result = await parseExcelFile(file);
          if (!meta.eventName) {
            meta = {
              eventName: result.eventName,
              district: result.district,
              dutyDateFrom: result.dutyDateFrom,
              dutyDateTo: result.dutyDateTo,
            };
          }
          allRecords.push(...result.records);
        } else {
          // 📸 Image files are no longer supported — show error
          setError(
            `फाइल "${file.name}" एक छवि है। कृपया Excel (.xlsx) फाइल अपलोड करें। ` +
            `AI/API आधारित छवि पार्सिंग हटा दी गई है।`
          );
          setLoading(false);
          return;
        }
      }

      if (allRecords.length === 0) {
        throw new Error('कोई ड्यूटी रिकॉर्ड नहीं मिला। कृपया सही Excel फाइल अपलोड करें।');
      }

      // Save to localStorage and navigate
      setLoadingStep('डेटा व्यवस्थित किया जा रहा है...');
      localStorage.setItem(
        'duty_cards_result',
        JSON.stringify({ ...meta, records: allRecords })
      );

      setLoadingStep('सफलता! रिडायरेक्ट किया जा रहा है...');
      router.push('/cards');
    } catch (err: any) {
      console.error(err);
      setError(
        err.message || 'सॉफ्टवेयर डेटा निकालने में असमर्थ रहा। कृपया पुनः प्रयास करें।'
      );
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-900 via-indigo-950 to-slate-900 text-white flex flex-col items-center justify-center p-4">
      {/* Header Info */}
      <div className="text-center max-w-2xl mb-8 space-y-2">
        <span className="px-3 py-1 bg-indigo-500/20 text-indigo-400 border border-indigo-500/30 rounded-full text-xs font-bold uppercase tracking-wider">
          उत्तर प्रदेश पुलिस बुलन्दशहर
        </span>
        <h1 className="text-4xl font-extrabold tracking-tight bg-clip-text text-transparent bg-gradient-to-r from-white via-indigo-200 to-indigo-400">
          Police Duty Card Generator
        </h1>
        <p className="text-sm text-slate-400">
          काँवड़ यात्रा-2025 या अन्य पुलिस ड्यूटी चार्ट की Excel (.xlsx) फाइल अपलोड करें और तुरंत प्रिंट-रेडी ड्यूटी कार्ड्स प्राप्त करें।
        </p>
      </div>

      {/* Upload Container */}
      <div className="w-full max-w-2xl bg-slate-900/60 backdrop-blur-md border border-slate-800 rounded-2xl shadow-xl p-8 space-y-6">
        {/* Excel Recommendation Banner */}
        <div className="bg-green-900/30 border border-green-500 rounded-lg p-3 mb-4 text-sm">
          <span className="text-green-400 font-bold">✅ Excel फाइल अपलोड करें (अनुशंसित)</span>
          <p className="text-green-300 mt-1">
            Excel (.xlsx) फाइल अपलोड करने पर तुरंत और सटीक कार्ड बनते हैं —
            कोई इंटरनेट या AI की जरूरत नहीं।
          </p>
        </div>

        {error && (
          <div className="bg-red-500/10 border border-red-500/20 text-red-200 px-4 py-3 rounded-lg text-sm">
            {error}
          </div>
        )}

        {/* Drag & Drop Area */}
        <div
          onDragOver={handleDragOver}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className="border-2 border-dashed border-slate-700 hover:border-indigo-500 hover:bg-slate-800/20 rounded-xl p-8 text-center cursor-pointer transition-all flex flex-col items-center justify-center space-y-3 group"
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            multiple
            accept=".xlsx,.xls"
            className="hidden"
          />
          <div className="w-16 h-16 bg-slate-800/80 rounded-full flex items-center justify-center group-hover:scale-110 transition-transform shadow-inner">
            <svg
              className="w-8 h-8 text-slate-400 group-hover:text-indigo-400 transition-colors"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
              xmlns="http://www.w3.org/2000/svg"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth="2"
                d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"
              ></path>
            </svg>
          </div>
          <div className="space-y-1">
            <p className="font-semibold text-slate-200">ड्रैग और ड्रॉप करें या फ़ाइलें चुनें</p>
            <p className="text-xs text-slate-500">समर्थित फाइल प्रारूप: Excel (.xlsx, .xls)</p>
          </div>
        </div>

        {/* Selected Files */}
        {files.length > 0 && (
          <div className="space-y-3">
            <h3 className="text-xs font-bold uppercase text-slate-400 tracking-wider">
              चयनित फाइलें ({files.length})
            </h3>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
              {files.map((file, index) => (
                <div
                  key={index}
                  className="bg-slate-850 border border-slate-800 rounded-lg p-2 flex items-center justify-between gap-2 text-xs text-slate-300"
                >
                  <span className="truncate flex-1">{file.name}</span>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      removeFile(index);
                    }}
                    className="text-slate-500 hover:text-red-400 w-5 h-5 flex items-center justify-center hover:bg-slate-800 rounded"
                  >
                    &times;
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Action Button / Loading State */}
        {loading ? (
          <div className="bg-slate-950/40 border border-slate-800 rounded-xl p-6 flex flex-col items-center justify-center space-y-4">
            <div className="w-10 h-10 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin"></div>
            <div className="text-center space-y-1">
              <p className="font-semibold text-slate-200">डेटा प्रोसेस किया जा रहा है...</p>
              <p className="text-xs text-slate-400 animate-pulse">{loadingStep}</p>
            </div>
          </div>
        ) : (
          <button
            onClick={handleUpload}
            disabled={files.length === 0}
            className={`w-full py-3.5 rounded-xl font-bold text-sm tracking-wide shadow-lg transition-all ${
              files.length === 0
                ? 'bg-slate-800 text-slate-500 cursor-not-allowed border border-slate-850'
                : 'bg-indigo-600 hover:bg-indigo-500 text-white cursor-pointer active:scale-[0.99]'
            }`}
          >
            डेटा निकालें और कार्ड जनरेट करें (Extract & Generate)
          </button>
        )}
      </div>
    </main>
  );
}
