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

  if (rows.length === 0) {
    throw new Error('Excel फाइल खाली है।');
  }

  // 1. Detect if it is the Tabular/Shift format or Standard format
  let isTabularShift = false;
  let hasMorning = false;
  let hasNight = false;
  
  for (let r = 0; r < Math.min(5, rows.length); r++) {
    const rowText = rows[r].join(' ');
    if (rowText.includes('सुबह') || rowText.includes('दिन') || rowText.includes('08.00') || rowText.includes('08:00')) {
      hasMorning = true;
    }
    if (rowText.includes('रात्रि') || rowText.includes('रात्री') || rowText.includes('रात') || rowText.includes('20.00') || rowText.includes('20:00')) {
      hasNight = true;
    }
  }

  if (hasMorning || hasNight || (rows[0] && rows[0].length >= 5)) {
    isTabularShift = true;
  }

  // Extract metadata
  let eventName = '';
  let district = 'बुलन्दशहर';
  let dutyDateFrom = '';
  let dutyDateTo = '';

  // Dates search
  const datePattern = /(\d{1,2}[./]\d{1,2}[./]\d{4})/g;
  for (let r = 0; r < rows.length; r++) {
    const rowText = rows[r].join(' ');
    const dates = rowText.match(datePattern);
    if (dates && dates.length >= 1) {
      dutyDateFrom = dates[0];
      if (dates.length >= 2) dutyDateTo = dates[1];
      break;
    }
  }

  // District search
  for (let r = 0; r < Math.min(5, rows.length); r++) {
    const rowText = rows[r].join(' ');
    if (rowText.includes('जनपद')) {
      const match = rowText.match(/जनपद\s+(\S+)/);
      if (match) {
        district = match[1].replace(/[^\u0900-\u097F]/g, '').trim(); // Keep only Hindi characters
      }
    }
  }

  if (isTabularShift) {
    // --- TABULAR SHIFT FORMAT ---
    if (rows[0] && rows[0][0]) {
      eventName = String(rows[0][0]).trim();
    } else {
      eventName = 'ड्यूटी कार्ड';
    }

    let thanaColIndex = 0;
    let serialColIndex = 1;
    let placeColIndex = 2;
    let dayOfficerColIndex = 3;
    let dayMobileColIndex = 4;
    let nightOfficerColIndex = 5;
    let nightMobileColIndex = 6;

    let dayShiftTime = 'सुबह 08.00 बजे से 20.00 बजे तक';
    let nightShiftTime = 'रात्री 20.00 बजे से 08.00 बजे तक';

    // Find start of data first so we know where the headers end
    let dataStartRow = 3;
    for (let r = 0; r < rows.length; r++) {
      const rowText = rows[r].join(' ');
      if (rowText.includes('कर्मचारी') || rowText.includes('मो0नं0') || rowText.includes('मोबाइल')) {
        dataStartRow = r + 1;
      }
    }

    // Dynamic headers identification - strictly scan ONLY header rows (before dataStartRow)
    for (let r = 0; r < dataStartRow; r++) {
      const row = rows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const val = String(row[c] || '').trim();
        if (val.includes('थाना')) {
          thanaColIndex = c;
        }
        if (val.includes('सं') || val.includes('क्र')) {
          serialColIndex = c;
        }
        if (val.includes('स्थान')) {
          placeColIndex = c;
        }
        if (val.includes('सुबह') || val.includes('दिन') || val.includes('08.00') || val.includes('08:00')) {
          dayOfficerColIndex = c;
          dayShiftTime = val;
          if (c + 1 < row.length) dayMobileColIndex = c + 1;
        }
        if (val.includes('रात्रि') || val.includes('रात्री') || val.includes('रात') || val.includes('20.00') || val.includes('20:00')) {
          nightOfficerColIndex = c;
          nightShiftTime = val;
          if (c + 1 < row.length) nightMobileColIndex = c + 1;
        }
      }
    }

    // Determine dutyType from eventName
    let dutyType = 'बैरियर ड्यूटी';
    if (eventName.includes('ड्यूटी')) {
      const match = eventName.match(/(.*?ड्यूटी)/);
      if (match) dutyType = match[1].trim();
    }

    interface RowBlock {
      thanaArea: string;
      serialNum: string;
      dutyPlace: string;
      dayMainName: string;
      dayMainMobile: string;
      nightMainName: string;
      nightMainMobile: string;
      daySupporting: { name: string; mobile: string }[];
      nightSupporting: { name: string; mobile: string }[];
    }

    const blocks: RowBlock[] = [];
    let currentBlock: RowBlock | null = null;
    let lastThana = '';

    for (let i = dataStartRow; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const colA = String(row[thanaColIndex] || '').trim();
      const colB = String(row[serialColIndex] || '').trim();
      const colC = String(row[placeColIndex] || '').trim();

      if (colA) {
        lastThana = colA;
      }

      // A new block starts if we have a non-empty serial number
      const isNewBlock = colB !== '' && !isNaN(Number(colB));

      if (isNewBlock) {
        currentBlock = {
          thanaArea: lastThana,
          serialNum: colB,
          dutyPlace: colC || (currentBlock ? currentBlock.dutyPlace : ''),
          dayMainName: String(row[dayOfficerColIndex] || '').trim(),
          dayMainMobile: String(row[dayMobileColIndex] || '').trim(),
          nightMainName: String(row[nightOfficerColIndex] || '').trim(),
          nightMainMobile: String(row[nightMobileColIndex] || '').trim(),
          daySupporting: [],
          nightSupporting: [],
        };
        blocks.push(currentBlock);
      } else if (currentBlock) {
        // Supporting staff rows
        const dayName = String(row[dayOfficerColIndex] || '').trim();
        const dayMobile = String(row[dayMobileColIndex] || '').trim();
        const nightName = String(row[nightOfficerColIndex] || '').trim();
        const nightMobile = String(row[nightMobileColIndex] || '').trim();

        if (dayName && !dayName.includes('कर्मचारी') && !dayName.includes('नाम')) {
          currentBlock.daySupporting.push({ name: dayName, mobile: dayMobile });
        }
        if (nightName && !nightName.includes('कर्मचारी') && !nightName.includes('नाम')) {
          currentBlock.nightSupporting.push({ name: nightName, mobile: nightMobile });
        }
      }
    }

    // Build final records
    const records: OfficerRecord[] = [];
    blocks.forEach((block) => {
      // 1. Day Shift Record
      if (block.dayMainName) {
        records.push({
          id: `${Date.now()}-${records.length}-day`,
          dutyType: dutyType,
          mainOfficerName: block.dayMainName,
          mainOfficerMobile: block.dayMainMobile,
          supportingOfficers: block.daySupporting,
          dutyPlace: block.dutyPlace,
          thanaArea: block.thanaArea,
          dutyTime: dayShiftTime,
          zonalMagistrate: '',
          zonalPoliceOfficer: '',
          sectorMagistrate: '',
          sectorPoliceOfficer: '',
        });
      }

      // 2. Night Shift Record
      if (block.nightMainName) {
        records.push({
          id: `${Date.now()}-${records.length}-night`,
          dutyType: dutyType,
          mainOfficerName: block.nightMainName,
          mainOfficerMobile: block.nightMainMobile,
          supportingOfficers: block.nightSupporting,
          dutyPlace: block.dutyPlace,
          thanaArea: block.thanaArea,
          dutyTime: nightShiftTime,
          zonalMagistrate: '',
          zonalPoliceOfficer: '',
          sectorMagistrate: '',
          sectorPoliceOfficer: '',
        });
      }
    });

    return {
      eventName,
      district,
      dutyDateFrom,
      dutyDateTo,
      records,
    };

  } else {
    // --- STANDARD ORIGINAL FORMAT ---
    // Search first 5 rows for metadata
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const rowText = rows[i].join(' ').trim();
      
      // Event name
      if (!eventName && (rowText.includes('यात्रा') || rowText.includes('शिवरात्रि') || rowText.includes('मेला') || rowText.includes('ड्यूटी'))) {
        eventName = rowText.split('\n')[0].trim();
      }
    }

    let dataStartRow = -1;
    for (let i = 0; i < rows.length; i++) {
      const rowText = rows[i].join(' ');
      if (rowText.includes('ड्यूटी का प्रकार') || rowText.includes('अधिकारी') || rowText.includes('कर्मचारी')) {
        dataStartRow = i + 1;
        break;
      }
    }
    
    if (dataStartRow === -1) dataStartRow = 3;

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

      if (!colA && !colB && !colC) continue;

      if (colA.includes('ड्यूटी का प्रकार')) continue; 
      
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

      if (colA && !colA.includes('ड्यूटी का') && !colA.includes('थाना क्षेत्र')) {
        if (colA.includes('ड्यूटी') || (colA.length < 30 && colB && !currentDutyType)) {
          if (currentMainOfficer || currentDutyPlace) saveRecord();
          
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

      if (colB.includes('सहयोगी') || colB.includes('पुलिसकमियों')) {
        inSupportingSection = true;
        continue;
      }

      if (inSupportingSection && colB && !colB.includes('सहयोगी')) {
        currentSupportingOfficers.push({ name: colB, mobile: colC });
        continue;
      }

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
    
    saveRecord();

    return {
      eventName: eventName || 'ड्यूटी कार्ड',
      district: district || 'बुलन्दशहर',
      dutyDateFrom,
      dutyDateTo,
      records,
    };
  }
}
