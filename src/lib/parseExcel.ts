import * as XLSX from 'xlsx';
import type { OfficerRecord, ExtractionResult } from '../types';

export async function parseExcelFile(file: File): Promise<ExtractionResult> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  
  // Get the first sheet
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  
  // Convert to array of arrays (raw rows)
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { 
    header: 1, 
    defval: '',
    blankrows: false 
  });

  // Extract event metadata from header rows
  // Row 0 typically has the event name (e.g. "काँवड़ यात्रा-2025")
  // Row 1 typically has "ड्यूटी कार्ड जनपद बुलन्दशहर"  
  // Row 2 typically has "ड्यूटी दिनांक 11.07.2025 से 24.07.2025 तक"
  
  let eventName = '';
  let district = '';
  let dutyDateFrom = '';
  let dutyDateTo = '';
  
  // Search first 5 rows for metadata
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const rowText = rows[i].join(' ').trim();
    
    // Event name — look for यात्रा or शिवरात्रि or similar
    if (!eventName && (rowText.includes('यात्रा') || rowText.includes('शिवरात्रि') || rowText.includes('मेला'))) {
      eventName = rowText.split('\n')[0].trim();
    }
    
    // District — look for बुलन्दशहर or जनपद
    if (!district && rowText.includes('जनपद')) {
      const match = rowText.match(/जनपद\s+(\S+)/);
      if (match) district = match[1];
    }
    
    // Duty dates — look for दिनांक with date pattern
    if (!dutyDateFrom && rowText.includes('दिनांक')) {
      const datePattern = /(\d{1,2}[./]\d{1,2}[./]\d{4})/g;
      const dates = rowText.match(datePattern);
      if (dates && dates.length >= 1) dutyDateFrom = dates[0];
      if (dates && dates.length >= 2) dutyDateTo = dates[1];
    }
  }

  // Now parse the data rows
  // Find where the actual data table starts (look for column header row)
  let dataStartRow = -1;
  for (let i = 0; i < rows.length; i++) {
    const rowText = rows[i].join(' ');
    if (rowText.includes('ड्यूटी का प्रकार') || rowText.includes('अधिकारी') || rowText.includes('कर्मचारी')) {
      dataStartRow = i + 1; // Data starts after header row
      break;
    }
  }
  
  if (dataStartRow === -1) dataStartRow = 3; // Default fallback

  // Parse records
  const records: OfficerRecord[] = [];
  
  let currentDutyType = '';
  let currentMainOfficer = '';
  let currentMainMobile = '';
  let currentSupportingOfficers: { name: string; mobile: string }[] = [];
  let currentDutyPlace = '';
  let currentThanaArea = '';
  let currentDutyTime = '';
  let inSupportingSection = false;
  let currentDutyTimeLabel = false;
  let currentDutyPlaceLabel = false;
  let currentThanaLabel = false;

  const saveRecord = () => {
    if (currentMainOfficer || currentDutyPlace) {
      records.push({
        id: `${Date.now()}-${records.length}`,
        dutyType: currentDutyType || 'ड्यूटी',
        mainOfficerName: currentMainOfficer,
        mainOfficerMobile: currentMainMobile,
        supportingOfficers: [...currentSupportingOfficers],
        dutyPlace: currentDutyPlace,
        thanaArea: currentThanaArea,
        dutyTime: currentDutyTime || 'प्रातः 08:00 बजे से रात्रि 20:00 बजे तक',
        zonalMagistrate: '',
        zonalPoliceOfficer: '',
        sectorMagistrate: '',
        sectorPoliceOfficer: '',
      });
    }
  };

  for (let i = dataStartRow; i < rows.length; i++) {
    const row = rows[i];
    const colA = String(row[0] || '').trim();
    const colB = String(row[1] || '').trim();
    const colC = String(row[2] || '').trim();

    // Skip completely empty rows
    if (!colA && !colB && !colC) continue;

    // Detect section labels in column A
    if (colA.includes('ड्यूटी का प्रकार')) continue; // skip header
    
    if (colA.includes('ड्यूटी का स्थान')) {
      currentDutyPlaceLabel = true;
      currentThanaLabel = false;
      currentDutyTimeLabel = false;
      continue;
    }
    
    if (colA.includes('थाना क्षेत्र')) {
      currentThanaLabel = true;
      currentDutyPlaceLabel = false;
      currentDutyTimeLabel = false;
      continue;
    }
    
    if (colA.includes('ड्यूटी का समय')) {
      currentDutyTimeLabel = true;
      currentDutyPlaceLabel = false;
      currentThanaLabel = false;
      continue;
    }

    // If colA has a duty type (not a label), it's a new record starting
    if (colA && !colA.includes('ड्यूटी का') && !colA.includes('थाना क्षेत्र')) {
      // Check if it looks like a duty type (contains ड्यूटी or is short label)
      if (colA.includes('ड्यूटी') || (colA.length < 30 && colB && !currentDutyType)) {
        // Save previous record if exists
        if (currentMainOfficer || currentDutyPlace) saveRecord();
        
        // Start new record
        currentDutyType = colA;
        currentMainOfficer = colB;
        currentMainMobile = colC;
        currentSupportingOfficers = [];
        currentDutyPlace = '';
        currentThanaArea = '';
        currentDutyTime = '';
        inSupportingSection = false;
        currentDutyPlaceLabel = false;
        currentThanaLabel = false;
        currentDutyTimeLabel = false;
        continue;
      }
    }

    // "सहयोगी पुलिसकमियों के नाम" header row
    if (colB.includes('सहयोगी') || colB.includes('पुलिसकमियों')) {
      inSupportingSection = true;
      continue;
    }

    // Supporting officer row
    if (inSupportingSection && colB && !colB.includes('सहयोगी')) {
      currentSupportingOfficers.push({ name: colB, mobile: colC });
      continue;
    }

    // Value rows after labels
    if (currentDutyPlaceLabel && colA) {
      currentDutyPlace = colA;
      currentDutyPlaceLabel = false;
      continue;
    }
    
    if (currentThanaLabel && colA) {
      currentThanaArea = colA;
      currentThanaLabel = false;
      continue;
    }
    
    if (currentDutyTimeLabel && colA) {
      currentDutyTime = colA;
      currentDutyTimeLabel = false;
      continue;
    }
  }
  
  // Save the last record
  saveRecord();

  return {
    eventName: eventName || 'ड्यूटी कार्ड',
    district: district || 'बुलन्दशहर',
    dutyDateFrom,
    dutyDateTo,
    records,
  };
}
