import * as XLSX from 'xlsx';
import type { OfficerRecord, ExtractionResult } from '../types';
import { krutidevToUnicode } from './krutidevToUnicode';

/**
 * Normalizes text. If it is Krutidev (ASCII layout), converts to Unicode.
 * If it is already Unicode Devanagari, keeps it untouched.
 */
function cleanAndConvert(val: any): string {
  const str = String(val === null || val === undefined ? '' : val).trim();
  if (!str) return '';
  // If it already has Devanagari characters, it is Unicode. Do not convert.
  if (/[\u0900-\u097F]/.test(str)) {
    return str;
  }
  // If it has English alphabetic characters, it is likely Krutidev legacy text.
  if (/[a-zA-Z]/.test(str)) {
    return krutidevToUnicode(str);
  }
  return str;
}

/**
 * Universal Excel parser for Police Duty Card generation.
 * 
 * Automatically detects whether the file is:
 * - FORMAT A (Barrier-style, side-by-side day/night shifts)
 * - FORMAT B (Zonal/Sector-style, vertically interleaved shifts)
 * 
 * Auto-detects column indices based on header keywords.
 */
export async function parseExcelFile(file: File): Promise<ExtractionResult> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    blankrows: false,
  });

  if (rawRows.length === 0) {
    throw new Error('Excel फाइल खाली है।');
  }

  // Convert the entire sheet to Unicode first
  const rows = rawRows.map(row => row.map(cleanAndConvert));

  // Determine file format by column count
  const maxCols = Math.max(...rows.map((r) => r.length));

  // Extract event name (from row 0, col 0 or first non-empty cell in first few rows)
  let eventName = 'ड्यूटी कार्ड';
  for (let r = 0; r < Math.min(3, rows.length); r++) {
    const firstCell = String(rows[r]?.[0] || '').trim();
    if (firstCell && firstCell.length > 5 && !firstCell.includes("थाना") && !firstCell.includes("क्र")) {
      eventName = firstCell.split('\n')[0].trim();
      break;
    }
  }

  // Extract dates
  let dutyDateFrom = '';
  let dutyDateTo = '';
  const dateRe = /(\d{1,2}[-./]\d{1,2}[-./]\d{4})/g;
  for (const row of rows) {
    const text = row.join(' ');
    const dates = text.match(dateRe);
    if (dates && dates.length >= 1) {
      dutyDateFrom = dates[0];
      if (dates.length >= 2) dutyDateTo = dates[1];
      break;
    }
  }

  // Extract district
  let district = 'बुलन्दशहर';
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const text = rows[i].join(' ');
    if (text.includes('जनपद')) {
      const m = text.match(/जनपद\s*[-—–]*\s*(\S+)/) || text.match(/जनपद\s+(\S+)/);
      if (m) {
        district = m[1].replace(/[^\u0900-\u097F]/g, '').trim() || district;
      }
    }
  }

  if (maxCols >= 9) {
    // ========== FORMAT B: Zonal/Sector (11-col) ==========
    return parseFormatB(rows, eventName, district, dutyDateFrom, dutyDateTo);
  } else {
    // ========== FORMAT A: Barrier (7-col) ==========
    return parseFormatA(rows, eventName, district, dutyDateFrom, dutyDateTo);
  }
}

// ────────────────────────────────────────────────────────────
//  FORMAT A: Barrier-style (7 columns)
//  Cols: [thana, serial, place, dayOfficer, dayMobile, nightOfficer, nightMobile]
// ────────────────────────────────────────────────────────────
function parseFormatA(
  rows: any[][],
  eventName: string,
  district: string,
  dutyDateFrom: string,
  dutyDateTo: string
): ExtractionResult {
  // Find header row dynamically
  let headerIndex = -1;
  let bestHeaderScore = 0;

  for (let r = 0; r < Math.min(10, rows.length); r++) {
    const row = rows[r];
    let score = 0;
    for (const cell of row) {
      if (cell.includes("थाना")) score += 2;
      if (cell.includes("बैरियर") || cell.includes("स्थान")) score += 2;
      if (cell.includes("कर्मचारी") || cell.includes("नाम")) score += 2;
      if (cell.includes("मो0") || cell.includes("मोबाइल") || cell.includes("मो०")) score += 2;
      if (cell.includes("क्र0") || cell.includes("सं0") || cell.includes("क0सं0") || cell.includes("क्र०")) score += 1;
      if (cell.includes("सुबह") || cell.includes("रात्री") || cell.includes("रात") || cell.includes("दिन")) score += 1;
    }
    if (score > bestHeaderScore) {
      bestHeaderScore = score;
      headerIndex = r;
    }
  }

  // Fallback defaults
  const cols = {
    thana: 0,
    serial: 1,
    place: 2,
    dayOfficerName: 3,
    dayOfficerMobile: 4,
    nightOfficerName: 5,
    nightOfficerMobile: 6
  };

  if (headerIndex !== -1) {
    const row1 = rows[headerIndex] || [];
    const row2 = rows[headerIndex + 1] || [];

    for (let idx = 0; idx < Math.max(row1.length, row2.length); idx++) {
      const val1 = row1[idx] || "";
      const val2 = row2[idx] || "";
      if (val1.includes("थाना") || val2.includes("थाना")) cols.thana = idx;
      if (val1.includes("सं") || val1.includes("क्र") || val2.includes("सं") || val2.includes("क्र")) cols.serial = idx;
      if (val1.includes("स्थान") || val1.includes("बैरियर") || val2.includes("स्थान") || val2.includes("बैरियर")) cols.place = idx;
    }

    const namesFound: number[] = [];
    const mobilesFound: number[] = [];
    for (let idx = 3; idx < Math.max(row1.length, row2.length); idx++) {
      const val1 = row1[idx] || "";
      const val2 = row2[idx] || "";
      if (val1.includes("कर्मचारी") || val1.includes("नाम") || val2.includes("कर्मचारी") || val2.includes("नाम")) {
        namesFound.push(idx);
      }
      if (val1.includes("मो0") || val1.includes("मोबाइल") || val1.includes("मो०") || val2.includes("मो0") || val2.includes("मोबाइल") || val2.includes("मो०")) {
        mobilesFound.push(idx);
      }
    }
    cols.dayOfficerName = namesFound[0] !== undefined ? namesFound[0] : 3;
    cols.dayOfficerMobile = mobilesFound[0] !== undefined ? mobilesFound[0] : 4;
    cols.nightOfficerName = namesFound[1] !== undefined ? namesFound[1] : 5;
    cols.nightOfficerMobile = mobilesFound[1] !== undefined ? mobilesFound[1] : 6;
  }

  // Find shift time labels from header row
  let dayShiftTime = 'सुबह 08.00 बजे से 20.00 बजे तक';
  let nightShiftTime = 'रात्री 20.00 बजे से 08.00 बजे तक';
  if (headerIndex !== -1 && rows[headerIndex]) {
    const row = rows[headerIndex];
    for (const cell of row) {
      if (cell.includes("सुबह") || cell.includes("दिन")) dayShiftTime = cell;
      if (cell.includes("रात्री") || cell.includes("रात्रि") || cell.includes("रात")) nightShiftTime = cell;
    }
  }

  let dataStart = headerIndex !== -1 ? headerIndex + 2 : 3;
  // Find first data row (where cols.serial is numeric)
  for (let i = dataStart; i < rows.length; i++) {
    const v = rows[i]?.[cols.serial];
    if (v !== '' && v !== undefined && !isNaN(Number(v))) {
      dataStart = i;
      break;
    }
  }

  interface Block {
    thana: string;
    place: string;
    dayMain: { name: string; mobile: string };
    nightMain: { name: string; mobile: string };
    daySupporting: { name: string; mobile: string }[];
    nightSupporting: { name: string; mobile: string }[];
  }

  const blocks: Block[] = [];
  let cur: Block | null = null;
  let lastThana = '';

  for (let i = dataStart; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const colThana = String(row[cols.thana] || '').trim();
    const colSerial = String(row[cols.serial] || '').trim();
    const colPlace = String(row[cols.place] || '').trim();
    const colDayName = String(row[cols.dayOfficerName] || '').trim();
    const colDayMob = String(row[cols.dayOfficerMobile] || '').trim();
    const colNightName = String(row[cols.nightOfficerName] || '').trim();
    const colNightMob = String(row[cols.nightOfficerMobile] || '').trim();

    if (colThana) lastThana = colThana;

    const isNew = colSerial !== '' && !isNaN(Number(colSerial));

    if (isNew) {
      cur = {
        thana: lastThana,
        place: colPlace,
        dayMain: { name: colDayName, mobile: colDayMob },
        nightMain: { name: colNightName, mobile: colNightMob },
        daySupporting: [],
        nightSupporting: [],
      };
      blocks.push(cur);
    } else if (cur) {
      if (colDayName) {
        cur.daySupporting.push({ name: colDayName, mobile: colDayMob });
      }
      if (colNightName) {
        cur.nightSupporting.push({ name: colNightName, mobile: colNightMob });
      }
    }
  }

  // Build OfficerRecords
  const records: OfficerRecord[] = [];
  let idCounter = 0;

  for (const b of blocks) {
    if (b.dayMain.name) {
      records.push({
        id: `${Date.now()}-${idCounter++}-day`,
        dutyType: 'बैरियर ड्यूटी',
        mainOfficerName: b.dayMain.name,
        mainOfficerMobile: b.dayMain.mobile,
        supportingOfficers: b.daySupporting,
        dutyPlace: b.place,
        thanaArea: b.thana,
        dutyTime: dayShiftTime,
        zonalMagistrate: '',
        zonalPoliceOfficer: '',
        sectorMagistrate: '',
        sectorPoliceOfficer: '',
      });
    }
    if (b.nightMain.name) {
      records.push({
        id: `${Date.now()}-${idCounter++}-night`,
        dutyType: 'बैरियर ड्यूटी',
        mainOfficerName: b.nightMain.name,
        mainOfficerMobile: b.nightMain.mobile,
        supportingOfficers: b.nightSupporting,
        dutyPlace: b.place,
        thanaArea: b.thana,
        dutyTime: nightShiftTime,
        zonalMagistrate: '',
        zonalPoliceOfficer: '',
        sectorMagistrate: '',
        sectorPoliceOfficer: '',
      });
    }
  }

  return { eventName, district, dutyDateFrom, dutyDateTo, records };
}

// ────────────────────────────────────────────────────────────
//  FORMAT B: Zonal/Sector-style (11 columns)
//  Cols: [superzone, zone, zonalMag, thana, sectorSerial, sectorHQ,
//         sectorName, sectorMag, shiftTime, officerName, mobile]
// ────────────────────────────────────────────────────────────
function parseFormatB(
  rows: any[][],
  eventName: string,
  district: string,
  dutyDateFrom: string,
  dutyDateTo: string
): ExtractionResult {
  // Find header row dynamically
  let headerIndex = -1;
  let bestHeaderScore = 0;

  for (let r = 0; r < Math.min(10, rows.length); r++) {
    const row = rows[r];
    let score = 0;
    for (const cell of row) {
      if (cell.includes("थाना")) score += 2;
      if (cell.includes("जोन")) score += 2;
      if (cell.includes("मजिस्ट्रेट")) score += 2;
      if (cell.includes("सेक्टर")) score += 2;
      if (cell.includes("मुख्यालय")) score += 2;
      if (cell.includes("मो0") || cell.includes("मोबाइल") || cell.includes("मो०")) score += 2;
      if (cell.includes("क्र0") || cell.includes("सं0") || cell.includes("क0सं0") || cell.includes("क्र०")) score += 1;
      if (cell.includes("समय")) score += 1;
      if (cell.includes("अधिकारी") || cell.includes("कर्मचारी")) score += 1;
    }
    if (score > bestHeaderScore) {
      bestHeaderScore = score;
      headerIndex = r;
    }
  }

  // Fallback defaults
  const cols = {
    zone: 1,
    zonalMag: 2,
    thana: 3,
    sectorSerial: 4,
    sectorHQ: 5,
    sectorName: 6,
    sectorMag: 7,
    shiftTime: 8,
    officer: 9,
    mobile: 10
  };

  if (headerIndex !== -1) {
    const headerRow = rows[headerIndex];
    headerRow.forEach((cell, idx) => {
      if (cell.includes("जोन") && !cell.includes("जोनल") && !cell.includes("मजिस्ट्रेट")) cols.zone = idx;
      if (cell.includes("जोनल")) cols.zonalMag = idx;
      if (cell.includes("थाना")) cols.thana = idx;
      if (cell.includes("मुख्यालय") || cell.includes("स्थान")) cols.sectorHQ = idx;
      if (cell.includes("सेक्टर") && cell.includes("नाम") && !cell.includes("अधिकारी") && !cell.includes("मजिस्ट्रेट")) cols.sectorName = idx;
      if (cell.includes("सेक्टर") && cell.includes("मजिस्ट्रेट")) cols.sectorMag = idx;
      if (cell.includes("समय")) cols.shiftTime = idx;
      if ((cell.includes("अधिकारी") || cell.includes("कर्मचारी") || cell.includes("नाम")) && 
          !cell.includes("थाना") && !cell.includes("मजिस्ट्रेट") && !cell.includes("मुख्यालय") && !cell.includes("सेक्टर का नाम")) {
        cols.officer = idx;
      }
      if (cell.includes("मो0") || cell.includes("मोबाइल") || cell.includes("मो०")) cols.mobile = idx;
      if (cell.includes("सेक्टर") && (cell.includes("संख्या") || cell.includes("सं0") || cell.includes("क्र0"))) cols.sectorSerial = idx;
    });
  }

  let dataStart = headerIndex !== -1 ? headerIndex + 1 : 2;
  // Find first data row (where cols.sectorSerial is a number)
  for (let i = dataStart; i < rows.length; i++) {
    const v = rows[i]?.[cols.sectorSerial];
    if (v !== '' && v !== undefined && !isNaN(Number(v))) {
      dataStart = i;
      break;
    }
  }

  interface Sector {
    thana: string;
    zone: string;
    zonalMagistrate: string;
    zonalPoliceOfficer: string;
    sectorHQ: string;
    sectorName: string;
    dayMagistrate: string;
    nightMagistrate: string;
    dayMain: { name: string; mobile: string };
    nightMain: { name: string; mobile: string };
    daySupporting: { name: string; mobile: string }[];
    nightSupporting: { name: string; mobile: string }[];
    dayTime: string;
    nightTime: string;
  }

  const sectors: Sector[] = [];
  let cur: Sector | null = null;
  let lastThana = '';
  let lastZone = '';
  let lastZonalMagRaw = '';
  let lastZonalMag = '';
  let lastZonalPolice = '';

  let phase: 'day' | 'night' = 'day';

  for (let i = dataStart; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const zone = String(row[cols.zone] || '').trim();
    const zonalMagRaw = String(row[cols.zonalMag] || '').trim();
    const thana = String(row[cols.thana] || '').trim();
    const sectorSerial = String(row[cols.sectorSerial] || '').trim();
    const sectorHQ = String(row[cols.sectorHQ] || '').trim();
    const sectorName = String(row[cols.sectorName] || '').trim();
    const sectorMag = String(row[cols.sectorMag] || '').trim();
    const shiftTime = String(row[cols.shiftTime] || '').trim();
    const officerName = String(row[cols.officer] || '').trim();
    const mobile = String(row[cols.mobile] || '').trim();

    // Track zone/thana carry-forward
    if (zone) lastZone = zone;
    if (thana) lastThana = thana;
    if (zonalMagRaw && zonalMagRaw !== lastZonalMagRaw) {
      lastZonalMagRaw = zonalMagRaw;
      // Split and extract Zonal Magistrate / Zonal Police Officer
      const lines = zonalMagRaw.split('\n').map(l => l.trim()).filter(Boolean);
      lastZonalMag = '';
      lastZonalPolice = '';
      for (const line of lines) {
        if (line.includes("मजिस्ट्रेट")) {
          lastZonalMag = line;
        } else if (line.includes("पुलिस") || line.includes("अधिकारी")) {
          lastZonalPolice = line;
        } else if (!lastZonalMag) {
          lastZonalMag = line;
        } else if (!lastZonalPolice) {
          lastZonalPolice = line;
        }
      }
    }

    // Skip super-zone or empty header rows
    if (!sectorSerial && !officerName && !sectorMag) continue;

    const isNewSector = sectorSerial !== '' && !isNaN(Number(sectorSerial));

    if (isNewSector) {
      phase = 'day';
      cur = {
        thana: lastThana,
        zone: lastZone,
        zonalMagistrate: lastZonalMag,
        zonalPoliceOfficer: lastZonalPolice,
        sectorHQ,
        sectorName,
        dayMagistrate: sectorMag,
        nightMagistrate: '',
        dayMain: { name: officerName, mobile },
        nightMain: { name: '', mobile: '' },
        daySupporting: [],
        nightSupporting: [],
        dayTime: shiftTime || 'प्रातः 08:00 बजे से 20:00 बजे तक',
        nightTime: '',
      };
      sectors.push(cur);
    } else if (cur) {
      // Detect night phase transition
      if (shiftTime && shiftTime !== cur.dayTime) {
        phase = 'night';
        cur.nightTime = shiftTime;
      }

      if (sectorMag && phase === 'night') {
        cur.nightMagistrate = sectorMag;
      } else if (sectorMag && phase === 'day' && !cur.dayMagistrate) {
        cur.dayMagistrate = sectorMag;
      }

      if (officerName) {
        if (phase === 'day') {
          if (!cur.dayMain.name) {
            cur.dayMain = { name: officerName, mobile };
          } else {
            cur.daySupporting.push({ name: officerName, mobile });
          }
        } else {
          if (!cur.nightMain.name) {
            cur.nightMain = { name: officerName, mobile };
          } else {
            cur.nightSupporting.push({ name: officerName, mobile });
          }
        }
      }
    }
  }

  // Build records
  const records: OfficerRecord[] = [];
  let idCounter = 0;

  for (const s of sectors) {
    // Day shift
    if (s.dayMain.name) {
      const sectorPoliceOfficer = s.dayMain.name + (s.dayMain.mobile ? ` मो0नं0 ${s.dayMain.mobile}` : '');
      records.push({
        id: `${Date.now()}-${idCounter++}-day`,
        dutyType: 'सेक्टर ड्यूटी',
        mainOfficerName: s.dayMain.name,
        mainOfficerMobile: s.dayMain.mobile,
        supportingOfficers: s.daySupporting,
        dutyPlace: s.sectorHQ,
        thanaArea: s.thana,
        dutyTime: s.dayTime,
        zonalMagistrate: s.zonalMagistrate,
        zonalPoliceOfficer: s.zonalPoliceOfficer,
        sectorMagistrate: s.dayMagistrate,
        sectorPoliceOfficer: sectorPoliceOfficer,
      });
    }

    // Night shift
    if (s.nightMain.name) {
      const sectorPoliceOfficer = s.nightMain.name + (s.nightMain.mobile ? ` मो0नं0 ${s.nightMain.mobile}` : '');
      records.push({
        id: `${Date.now()}-${idCounter++}-night`,
        dutyType: 'सेक्टर ड्यूटी',
        mainOfficerName: s.nightMain.name,
        mainOfficerMobile: s.nightMain.mobile,
        supportingOfficers: s.nightSupporting,
        dutyPlace: s.sectorHQ,
        thanaArea: s.thana,
        dutyTime: s.nightTime || 'रात्रि 20:00 बजे से प्रातः 08:00 बजे तक',
        zonalMagistrate: s.zonalMagistrate,
        zonalPoliceOfficer: s.zonalPoliceOfficer,
        sectorMagistrate: s.nightMagistrate || s.dayMagistrate,
        sectorPoliceOfficer: sectorPoliceOfficer,
      });
    }
  }

  return { eventName, district, dutyDateFrom, dutyDateTo, records };
}
