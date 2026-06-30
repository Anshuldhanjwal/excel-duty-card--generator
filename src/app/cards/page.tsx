'use client';

import React, { useState, useEffect, useRef } from 'react';
import { useRouter } from 'next/navigation';
import { ExtractionResult, OfficerRecord } from '../../types';
import { DutyCard } from '../../components/DutyCard';
import { EditCardModal } from '../../components/EditCardModal';
import { krutidevToUnicode } from '../../lib/krutidevToUnicode';

function cleanAndConvert(val: any): string {
  const str = String(val === null || val === undefined ? '' : val).trim();
  if (!str) return '';
  if (/[\u0900-\u097F]/.test(str)) {
    return str;
  }
  // Protect English codes, abbreviations, numbers from Krutidev conversion
  if (/[a-zA-Z]/.test(str)) {
    const clean = str.trim();
    if (/^[A-Z0-9\s\-_./()&]+$/.test(clean) || 
        clean.toLowerCase().includes('prv') || 
        clean.toLowerCase().includes('up') || 
        clean.toLowerCase().includes('sector') ||
        clean.toLowerCase().includes('zone')) {
      return clean;
    }
    return krutidevToUnicode(str);
  }
  return str;
}

function convertResult(res: ExtractionResult): ExtractionResult {
  return {
    eventName: cleanAndConvert(res.eventName),
    district: cleanAndConvert(res.district),
    dutyDateFrom: cleanAndConvert(res.dutyDateFrom),
    dutyDateTo: cleanAndConvert(res.dutyDateTo),
    records: (res.records || []).map(r => ({
      ...r,
      dutyType: cleanAndConvert(r.dutyType),
      mainOfficerName: cleanAndConvert(r.mainOfficerName),
      mainOfficerMobile: cleanAndConvert(r.mainOfficerMobile),
      mainOfficerPostingPlace: cleanAndConvert(r.mainOfficerPostingPlace),
      supportingOfficers: (r.supportingOfficers || []).map(s => ({
        name: cleanAndConvert(s.name),
        postingPlace: cleanAndConvert(s.postingPlace),
        mobile: cleanAndConvert(s.mobile)
      })),
      dutyPlace: cleanAndConvert(r.dutyPlace),
      thanaArea: cleanAndConvert(r.thanaArea),
      dutyTime: cleanAndConvert(r.dutyTime),
      zonalMagistrate: cleanAndConvert(r.zonalMagistrate),
      zonalPoliceOfficer: cleanAndConvert(r.zonalPoliceOfficer),
      sectorMagistrate: cleanAndConvert(r.sectorMagistrate),
      sectorPoliceOfficer: cleanAndConvert(r.sectorPoliceOfficer)
    }))
  };
}

export default function CardsPage() {
  const [data, setData] = useState<ExtractionResult | null>(null);
  const [editingRecord, setEditingRecord] = useState<OfficerRecord | null>(null);
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [printLayout, setPrintLayout] = useState<'a5' | 'a4' | 'a4-triple'>('a5');
  const [isPrinting, setIsPrinting] = useState(false);
  const printContainerRef = useRef<HTMLDivElement>(null);
  
  const cardsPerPage = 50;
  const router = useRouter();

  useEffect(() => {
    const saved = localStorage.getItem('duty_cards_result');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        const converted = convertResult(parsed);
        setData(converted);
        // Save back the converted Unicode data so it persists in Unicode
        localStorage.setItem('duty_cards_result', JSON.stringify(converted));
      } catch (err) {
        console.error('Failed to parse saved data:', err);
      }
    }
  }, []);

  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm]);

  useEffect(() => {
    const handleAfterPrint = () => {
      // Reset zoom on all cards before removing from DOM
      document.querySelectorAll('.duty-card-container').forEach(el => {
        (el as HTMLElement).style.zoom = '';
      });
      setIsPrinting(false);
      document.body.classList.remove('printing-all', 'printing-single', 'print-layout-a4-double', 'print-layout-a4-triple');
      const styleEl = document.getElementById('dynamic-print-page-style');
      if (styleEl) styleEl.remove();
    };
    window.addEventListener('afterprint', handleAfterPrint);
    return () => window.removeEventListener('afterprint', handleAfterPrint);
  }, []);

  if (!data) {
    return (
      <main className="min-h-screen bg-slate-900 text-white flex flex-col items-center justify-center p-4">
        <div className="w-8 h-8 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin mb-4"></div>
        <p className="text-sm text-slate-400">ड्यूटी कार्ड डेटा लोड हो रहा है...</p>
      </main>
    );
  }

  const updateEventField = (field: keyof ExtractionResult, value: string) => {
    setData((prev) => {
      if (!prev) return null;
      const updated = { ...prev, [field]: value };
      localStorage.setItem('duty_cards_result', JSON.stringify(updated));
      return updated;
    });
  };

  const handleSaveRecord = (updatedRecord: OfficerRecord) => {
    setData((prev) => {
      if (!prev) return null;
      const updatedRecords = prev.records.map((r) =>
        r.id === updatedRecord.id ? updatedRecord : r
      );
      const updated = { ...prev, records: updatedRecords };
      localStorage.setItem('duty_cards_result', JSON.stringify(updated));
      return updated;
    });
  };

  const handleDeleteRecord = (id: string) => {
    if (window.confirm('क्या आप इस ड्यूटी कार्ड को हटाना चाहते हैं?')) {
      setData((prev) => {
        if (!prev) return null;
        const updatedRecords = prev.records.filter((r) => r.id !== id);
        const updated = { ...prev, records: updatedRecords };
        localStorage.setItem('duty_cards_result', JSON.stringify(updated));
        return updated;
      });
    }
  };

  // Scale each duty card to fit exactly in half an A4 page (portrait)
  const scaleCardsForA4 = () => {
    const container = printContainerRef.current;
    if (!container) return;

    // Temporarily make the print container measurable (it's display:none on screen)
    container.style.display = 'block';
    container.style.position = 'absolute';
    container.style.left = '-9999px';
    container.style.width = '200mm'; // A4 width (210mm) minus 5mm margins each side
    container.style.visibility = 'hidden';

    // Create a temporary ruler to accurately convert mm to px at the current DPI
    // Target height = (297mm page - 10mm margins) / 2 cards - 4mm padding = 139.5mm
    const ruler = document.createElement('div');
    ruler.style.cssText = 'position:absolute;visibility:hidden;height:139.5mm;';
    document.body.appendChild(ruler);
    const targetHeightPx = ruler.offsetHeight;
    document.body.removeChild(ruler);

    // Measure and scale each card
    const cards = container.querySelectorAll('.duty-card-container') as NodeListOf<HTMLElement>;
    cards.forEach(card => {
      card.style.zoom = '';
      card.style.maxWidth = '100%';
      const naturalHeight = card.scrollHeight;
      if (naturalHeight > targetHeightPx) {
        const scale = targetHeightPx / naturalHeight;
        card.style.zoom = String(Math.min(scale, 1));
      }
    });

    // Reset container — print media CSS will handle visibility
    container.style.display = '';
    container.style.position = '';
    container.style.left = '';
    container.style.width = '';
    container.style.visibility = '';
  };

  const handlePrintAll = () => {
    setIsPrinting(true);
    // Let the DOM update with all cards before calling print dialog
    setTimeout(() => {
      const styleId = 'dynamic-print-page-style';
      let styleEl = document.getElementById(styleId);
      if (!styleEl) {
        styleEl = document.createElement('style');
        styleEl.id = styleId;
        document.head.appendChild(styleEl);
      }

      document.body.classList.remove('printing-single');
      document.body.classList.add('printing-all');

      if (printLayout === 'a4') {
        styleEl.innerHTML = `@page { size: A4 portrait; margin: 5mm; }`;
        document.body.classList.add('print-layout-a4-double');
        document.body.classList.remove('print-layout-a4-triple');

        // Measure cards and apply zoom scaling, then print
        requestAnimationFrame(() => {
          scaleCardsForA4();
          setTimeout(() => window.print(), 200);
        });
      } else if (printLayout === 'a4-triple') {
        styleEl.innerHTML = `@page { size: A4 portrait; margin: 8mm; }`;
        document.body.classList.add('print-layout-a4-triple');
        document.body.classList.remove('print-layout-a4-double');
        setTimeout(() => window.print(), 150);
      } else {
        styleEl.innerHTML = `@page { size: A5 landscape; margin: 8mm; }`;
        document.body.classList.remove('print-layout-a4-double', 'print-layout-a4-triple');
        setTimeout(() => window.print(), 150);
      }
    }, 300);
  };

  const handlePrintSingle = (id: string) => {
    // Inject style tag to force A5 landscape for single cards
    const styleId = 'dynamic-print-page-style';
    let styleEl = document.getElementById(styleId);
    if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.id = styleId;
      document.head.appendChild(styleEl);
    }
    styleEl.innerHTML = `@page { size: A5 landscape; margin: 8mm; }`;

    document.querySelectorAll('.printing-now').forEach((el) => {
      el.classList.remove('printing-now');
    });

    const target = document.getElementById(`card-${id}`);
    if (target) {
      target.classList.add('printing-now');
    }

    document.body.classList.remove('printing-all', 'print-layout-a4-double', 'print-layout-a4-triple');
    document.body.classList.add('printing-single');
    
    setTimeout(() => {
      window.print();
    }, 150);
  };

  // Filter records based on search term (match thana, place, officer name, or type)
  const filteredRecords = data.records.filter((r) => {
    const term = searchTerm.toLowerCase();
    return (
      r.mainOfficerName.toLowerCase().includes(term) ||
      r.dutyPlace.toLowerCase().includes(term) ||
      r.thanaArea.toLowerCase().includes(term) ||
      r.dutyType.toLowerCase().includes(term)
    );
  });

  const totalPages = Math.ceil(filteredRecords.length / cardsPerPage);
  const startIndex = (currentPage - 1) * cardsPerPage;
  const paginatedRecords = filteredRecords.slice(startIndex, startIndex + cardsPerPage);

  return (
    <main className="min-h-screen bg-slate-900 text-white p-6 print:bg-white print:text-black print:p-0">
      {/* HEADER CONTROLS (Hidden in Print) */}
      <div className="max-w-6xl mx-auto mb-8 bg-slate-950/40 border border-slate-800 rounded-2xl p-6 space-y-6 print:hidden">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          <div>
            <h1 className="text-2xl font-bold tracking-tight">जनरेट किए गए ड्यूटी कार्ड्स</h1>
            <p className="text-sm text-slate-400 mt-1">
              कुल {data.records.length} ड्यूटी समूह पाए गए ({filteredRecords.length} प्रदर्शित)
            </p>
          </div>
          <div className="flex flex-wrap gap-3 items-center">
            {/* Print Layout Selector */}
            <div className="flex items-center gap-2 bg-slate-900/60 px-3 py-2 rounded-xl border border-slate-800">
              <span className="text-xs text-slate-400 font-semibold uppercase">प्रिंट लेआउट:</span>
              <select
                value={printLayout}
                onChange={(e) => setPrintLayout(e.target.value as 'a5' | 'a4' | 'a4-triple')}
                className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer"
              >
                <option value="a5" className="bg-slate-950 text-white">1 कार्ड/पेज (A5 Landscape)</option>
                <option value="a4" className="bg-slate-950 text-white">2 कार्ड/पेज (A4 Portrait)</option>
                <option value="a4-triple" className="bg-slate-950 text-white">3 कार्ड/पेज (A4 Portrait)</option>
              </select>
            </div>
            
            <button
              onClick={() => router.push('/upload')}
              className="px-4 py-2 border border-slate-700 hover:bg-slate-800 rounded-xl text-sm font-semibold transition-colors cursor-pointer"
            >
              &larr; नया अपलोड
            </button>
            
            <button
              onClick={handlePrintAll}
              disabled={filteredRecords.length === 0}
              className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 rounded-xl text-sm font-bold shadow-lg transition-colors flex items-center gap-2 cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed"
            >
              <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path>
              </svg>
              सभी प्रिंट करें (Print All)
            </button>
          </div>
        </div>

        {/* Search Input */}
        <div className="relative">
          <input
            type="text"
            placeholder="अधिकारी का नाम, ड्यूटी स्थान, थाना क्षेत्र या ड्यूटी प्रकार खोजें..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-10 pr-4 py-3 text-sm text-white focus:outline-none focus:border-indigo-500 placeholder-slate-500"
          />
          <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
            <svg className="w-4 h-4 text-slate-500" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path>
            </svg>
          </div>
        </div>
      </div>

      {/* CARDS LIST CONTAINER (Screen View: Paginated) */}
      <div className="max-w-6xl mx-auto space-y-12 pb-8 print:hidden">
        {filteredRecords.length === 0 ? (
          <div className="text-center py-12 text-slate-500">
            कोई कार्ड मेल नहीं खाता।
          </div>
        ) : (
          paginatedRecords.map((record) => (
            <div
              key={record.id}
              className="card-container-wrapper bg-slate-950/20 border border-slate-800/40 p-6 rounded-2xl space-y-4 print:p-0 print:border-none print:bg-transparent"
            >
              {/* Card Title & Action Buttons (Hidden in Print) */}
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 bg-slate-900/40 px-4 py-3 rounded-xl border border-slate-800 print:hidden">
                <div className="text-xs text-slate-400">
                  <span className="font-semibold text-slate-200">स्थान:</span> {record.dutyPlace} |{' '}
                  <span className="font-semibold text-slate-200">थाना:</span> {record.thanaArea} |{' '}
                  <span className="font-semibold text-slate-200">श्रेणी:</span> {record.dutyType}
                </div>
                <div className="flex gap-2 w-full sm:w-auto">
                  <button
                    onClick={() => {
                      setEditingRecord(record);
                      setIsEditModalOpen(true);
                    }}
                    className="flex-1 sm:flex-initial px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-xs font-semibold rounded-lg transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
                    </svg>
                    सुधारें (Edit)
                  </button>
                  <button
                    onClick={() => handlePrintSingle(record.id)}
                    className="flex-1 sm:flex-initial px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-xs font-bold rounded-lg transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path>
                    </svg>
                    प्रिंट (Print)
                  </button>
                  <button
                    onClick={() => handleDeleteRecord(record.id)}
                    className="px-2 py-1.5 text-slate-400 hover:text-red-400 hover:bg-slate-800/60 rounded-lg transition-colors flex items-center justify-center cursor-pointer"
                    title="हटाएं"
                  >
                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
                    </svg>
                  </button>
                </div>
              </div>

              {/* RENDERED PIXEL-PERFECT CARD */}
              <div id={`card-${record.id}`} className="duty-card-wrapper bg-white rounded-lg p-2 print:p-0 print:rounded-none">
                <DutyCard
                  record={record}
                  eventName={data.eventName}
                  district={data.district}
                  dutyDateFrom={data.dutyDateFrom}
                  dutyDateTo={data.dutyDateTo}
                />
              </div>
            </div>
          ))
        )}
      </div>

      {/* PAGINATION NAVIGATION CONTROL (Hidden in Print) */}
      <div className="max-w-6xl mx-auto mb-16 print:hidden">
        {totalPages > 1 && (
          <div className="flex flex-col sm:flex-row justify-between items-center gap-4 bg-slate-950/40 border border-slate-800 p-4 rounded-2xl">
            <button
              onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
              disabled={currentPage === 1}
              className="w-full sm:w-auto px-4 py-2 bg-slate-800 hover:bg-slate-700 disabled:opacity-40 disabled:hover:bg-slate-800 rounded-xl text-xs font-semibold transition-colors cursor-pointer disabled:cursor-not-allowed"
            >
              &larr; पिछला (Prev)
            </button>
            <span className="text-sm text-slate-400 text-center">
              पृष्ठ <span className="font-semibold text-white">{currentPage}</span> / {totalPages} (कुल {filteredRecords.length} कार्ड्स)
            </span>
            <button
              onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
              disabled={currentPage === totalPages}
              className="w-full sm:w-auto px-4 py-2 bg-slate-800 hover:bg-slate-700 disabled:opacity-40 disabled:hover:bg-slate-800 rounded-xl text-xs font-semibold transition-colors cursor-pointer disabled:cursor-not-allowed"
            >
              अगला (Next) &rarr;
            </button>
          </div>
        )}
      </div>

      {/* PRINT-ONLY CONTAINER (Renders all cards on demand during window print, completely hidden on screen) */}
      {isPrinting && (
        <div ref={printContainerRef} className="hidden print:block print:space-y-0">
          {filteredRecords.map((record) => (
            <div
              key={record.id}
              className="card-container-wrapper print:block"
            >
              <div id={`card-${record.id}`} className="duty-card-wrapper bg-white p-0">
                <DutyCard
                  record={record}
                  eventName={data.eventName}
                  district={data.district}
                  dutyDateFrom={data.dutyDateFrom}
                  dutyDateTo={data.dutyDateTo}
                />
              </div>
            </div>
          ))}
        </div>
      )}

      {/* EDIT MODAL */}
      {editingRecord && (
        <EditCardModal
          isOpen={isEditModalOpen}
          onClose={() => {
            setIsEditModalOpen(false);
            setEditingRecord(null);
          }}
          record={editingRecord}
          onSave={handleSaveRecord}
        />
      )}
    </main>
  );
}
