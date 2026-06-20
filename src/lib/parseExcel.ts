import * as XLSX from 'xlsx';
import type { OfficerRecord, ExtractionResult } from '../types';
import { krutidevToUnicode } from './krutidevToUnicode';

/**
 * Checks if a string is likely standard English text or abbreviations (like PRV, UP 32, QRT, etc.)
 * to protect it from being parsed as legacy Krutidev font.
 */
function isLikelyEnglish(str: string): boolean {
  const clean = str.trim();
  if (!clean) return true;

  // 1. Pure numbers and punctuation
  if (/^[0-9\s\-_./()+,:]*$/.test(clean)) return true;

  // 2. Pure uppercase English codes, numbers, abbreviations, spaces
  if (/^[A-Z0-9\s\-_./()&]+$/.test(clean)) return true;

  // 3. Check for specific known English words/phrases (case insensitive)
  const words = clean.toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
  if (words.length > 0) {
    const englishWords = new Set([
      'sector', 'zone', 'duty', 'card', 'police', 'mobile', 'name', 'time', 
      'shift', 'date', 'from', 'to', 'fire', 'checking', 'team', 'route', 
      'ahar', 'mandir', 'picket', 'cut', 'dress', 'temp', 'op', 'extra', 
      'fantom', 'booklet', 'final', 'prv', 'qrt', 'prv112', 'up112', 'gps',
      'sheet', 'page', 'id', 'record', 'officer', 'supporting', 'magistrate'
    ]);
    
    const allEnglish = words.every(w => {
      // If it is a number
      if (/^\d+$/.test(w)) return true;
      // Very short helper words and common acronyms
      if (w.length <= 3 && /^[a-z]+$/.test(w)) {
        if (new Set(['in', 'at', 'on', 'of', 'up', 'by', 'co', 'si', 'hc', 'pc', 'gp', 'hg', 'rt', 'nh', 'sh']).has(w)) return true;
      }
      return englishWords.has(w);
    });
    if (allEnglish) return true;
  }

  return false;
}

/**
 * Normalizes text. If it is already Unicode Devanagari (Mangal), keeps it untouched.
 * If it is English text or abbreviations, keeps it untouched.
 * Otherwise, translates from legacy Krutidev layout to Unicode.
 */
function cleanAndConvert(val: any): string {
  const str = String(val === null || val === undefined ? '' : val).trim();
  if (!str) return '';
  // If it already has Devanagari characters, it is Unicode. Do not convert.
  if (/[\u0900-\u097F]/.test(str)) {
    return str;
  }
  // If it is standard English words/abbreviations, do not convert.
  if (isLikelyEnglish(str)) {
    return str;
  }
  // Convert Krutidev layout text.
  return krutidevToUnicode(str);
}

/**
 * Removes administrative suffix text from event name to keep it clean.
 */
function cleanEventName(name: string): string {
  if (!name) return 'काँवड़ यात्रा-2026';
  let cleaned = name.trim();
  
  // Truncate from keywords
  cleaned = cleaned.replace(/हेतु.*/g, '');
  cleaned = cleaned.replace(/के\s+लिए.*/g, '');
  cleaned = cleaned.replace(/का\s+विवरण.*/g, '');
  cleaned = cleaned.replace(/बनाये\s+गये.*/g, '');
  cleaned = cleaned.replace(/बनाए\s+गए.*/g, '');
  cleaned = cleaned.replace(/जनपद.*/g, '');
  
  // Clean up leading/trailing punctuation and spaces
  cleaned = cleaned.replace(/^[-\s,/—–]+|[-\s,/—–]+$/g, '');

  return cleaned.trim() || 'काँवड़ यात्रा-2026';
}

/**
 * Universal Excel parser that loops through all sheets in a workbook
 * and dynamically extracts police duty cards from each sheet.
 */
export async function parseExcelFile(file: File): Promise<ExtractionResult> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  const allRecords: OfficerRecord[] = [];

  // Default global metadata
  let eventName = 'काँवड़ यात्रा-2026';
  let district = 'बुलन्दशहर';
  let dutyDateFrom = '01.0.2026';
  let dutyDateTo = '13.08.2026';
  
  let foundMeta = false;

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rawRows: any[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });

    if (rawRows.length === 0) continue;

    // Convert the entire sheet to Unicode first (translating Krutidev where appropriate)
    const rows = rawRows.map(row => row.map(cleanAndConvert));

    // Try to extract event metadata from the first sheet containing typical title text
    if (!foundMeta) {
      for (let r = 0; r < Math.min(3, rows.length); r++) {
        const firstCell = String(rows[r]?.[0] || '').trim();
        if (firstCell && firstCell.length > 5 && !firstCell.includes("थाना") && !firstCell.includes("क्र")) {
          eventName = firstCell.split('\n')[0].trim();
          foundMeta = true;
          break;
        }
      }
    }

    // Try to extract dates if not already found
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

    // Try to extract district if not already found
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const text = rows[i].join(' ');
      if (text.includes('जनपद')) {
        const m = text.match(/जनपद\s*[-—–]*\s*(\S+)/) || text.match(/जनपद\s+(\S+)/);
        if (m) {
          district = m[1].replace(/[^\u0900-\u097F]/g, '').trim() || district;
        }
      }
    }

    // Dynamically parse the sheet based on detected layout
    const sheetRecords = parseSheetRows(rows, sheetName);
    allRecords.push(...sheetRecords);
  }

  // Final cleanup of global metadata
  eventName = cleanEventName(eventName);
  district = district.replace(/जनपद/g, '').replace(/[-—–\s]+/g, '').trim();
  dutyDateFrom = (dutyDateFrom || '').replace(/[-/]/g, '.');
  dutyDateTo = (dutyDateTo || '').replace(/[-/]/g, '.');

  return { eventName, district, dutyDateFrom, dutyDateTo, records: allRecords };
}

/**
 * Route sheet parsing to appropriate layout parser based on cell keyword detection.
 */
function parseSheetRows(rows: any[][], sheetName: string): OfficerRecord[] {
  const maxCols = Math.max(...rows.map(r => r.length));
  if (maxCols === 0) return [];

  // 1. Detect Day/Night split layout (Type 1)
  let hasDay = false;
  let hasNight = false;
  for (let r = 0; r < Math.min(3, rows.length); r++) {
    const text = rows[r].join(' ').toLowerCase();
    if (text.includes('सुबह') || text.includes('lqcg') || text.includes(' दिन ') || text.includes('izkr%') || text.includes('पाली 1') || text.includes('izFke') || text.includes('पाली-1') || text.includes('पाली—1')) {
      hasDay = true;
    }
    if (text.includes('रात्रि') || text.includes('raatri') || text.includes('jk=h') || text.includes(' रात ') || text.includes('f}rh;') || text.includes('पाली 2') || text.includes('पाली-2') || text.includes('पाली—2')) {
      hasNight = true;
    }
  }

  if (hasDay && hasNight) {
    return parseDayNightSheet(rows, sheetName);
  }

  // 2. Detect Zonal/Sector layout (Type 2)
  let hasZonal = false;
  for (let r = 0; r < Math.min(3, rows.length); r++) {
    const text = rows[r].join(' ').toLowerCase();
    if (text.includes('जोन') || text.includes('tksu') || text.includes('मजिस्ट्रेट') || text.includes('eftLVªsV') || text.includes('सेक्टर') || text.includes('lsDVj')) {
      hasZonal = true;
      break;
    }
  }

  if (hasZonal && maxCols >= 9) {
    return parseZonalSectorSheet(rows, sheetName);
  }

  // 3. Fallback to Simple List layout (Type 3)
  return parseSimpleListSheet(rows, sheetName);
}

/**
 * Type 1: Day/Night Split Sheet Parser
 */
function parseDayNightSheet(rows: any[][], sheetName: string): OfficerRecord[] {
  let headerIndex = -1;
  let bestHeaderScore = 0;

  for (let r = 0; r < Math.min(8, rows.length); r++) {
    const row = rows[r];
    let score = 0;
    for (const cell of row) {
      const c = cell.toLowerCase();
      if (c.includes("थाना") || c.includes("Fkkuk")) score += 2;
      if (c.includes("स्थान") || c.includes("LFkku") || c.includes("बैरियर") || c.includes("cSfj;j") || c.includes("fidsV") || c.includes("पिकेट") || c.includes("मार्ग") || c.includes("ekxZ")) score += 2;
      if (c.includes("नाम") || c.includes("uke") || c.includes("deZpkjh")) score += 2;
      if (c.includes("मो0") || c.includes("eks0") || c.includes("मोबाइल") || c.includes("eksckbZy")) score += 2;
      if (c.includes("सं0") || c.includes("dz0") || c.includes("सं") || c.includes("dz")) score += 1;
    }
    if (score > bestHeaderScore) {
      bestHeaderScore = score;
      headerIndex = r;
    }
  }

  let thanaCol = -1;
  let placeCol = -1;
  let serialCol = -1;
  const nameCols: number[] = [];
  const mobileCols: number[] = [];

  const row1 = headerIndex !== -1 ? rows[headerIndex] : [];
  const row2 = headerIndex !== -1 && rows[headerIndex + 1] ? rows[headerIndex + 1] : [];

  for (let idx = 0; idx < Math.max(row1.length, row2.length); idx++) {
    const val1 = String(row1[idx] || '').toLowerCase();
    const val2 = String(row2[idx] || '').toLowerCase();

    if (val1.includes("थाना") || val2.includes("थाना") || val1.includes("Fkkuk") || val2.includes("Fkkuk")) {
      thanaCol = idx;
    }
    if (val1.includes("स्थान") || val2.includes("स्थान") || val1.includes("LFkku") || val2.includes("LFkku") ||
        val1.includes("बैरियर") || val2.includes("बैरियर") || val1.includes("cSfj;j") || val2.includes("cSfj;j") ||
        val1.includes("fidsV") || val2.includes("fidsV") || val1.includes("पिकेट") || val2.includes("पिकेट") ||
        val1.includes("मार्ग") || val2.includes("मार्ग") || val1.includes("ekxZ") || val2.includes("ekxZ")) {
      placeCol = idx;
    }
    if (val1.includes("सं") || val2.includes("सं") || val1.includes("dz") || val2.includes("dz") ||
        val1.includes("क्र") || val2.includes("क्र") || val1.includes("Ø") || val2.includes("Ø")) {
      serialCol = idx;
    }
    if (val1.includes("नाम") || val2.includes("नाम") || val1.includes("uke") || val2.includes("uke") ||
        val1.includes("deZpkjh") || val2.includes("deZpkjh") || val1.includes("vf/kdkjh") || val2.includes("vf/kdkjh")) {
      nameCols.push(idx);
    }
    if (val1.includes("मो") || val2.includes("मो") || val1.includes("eks") || val2.includes("eks") ||
        val1.includes("मोबाइल") || val2.includes("मोबाइल") || val1.includes("eksckbZy") || val2.includes("eksckbZy")) {
      mobileCols.push(idx);
    }
  }

  if (serialCol === -1) serialCol = 0;
  if (placeCol === -1) placeCol = Math.min(2, Math.max(0, rows[0]?.length - 1));
  if (thanaCol === -1) thanaCol = 0;

  const dayNameCol = nameCols[0] !== undefined ? nameCols[0] : 3;
  const dayMobileCol = mobileCols[0] !== undefined ? mobileCols[0] : 4;
  const nightNameCol = nameCols[1] !== undefined ? nameCols[1] : 5;
  const nightMobileCol = mobileCols[1] !== undefined ? mobileCols[1] : 6;

  let dayTime = 'सुबह 08.00 बजे से 20.00 बजे तक';
  let nightTime = 'रात्रि 20.00 बजे से 08.00 बजे तक';

  if (headerIndex !== -1) {
    const row = rows[headerIndex];
    for (const cell of row) {
      const c = cell.toLowerCase();
      if (c.includes("सुबह") || c.includes("lqcg") || c.includes("दिन") || c.includes("izkr%")) dayTime = cell;
      if (c.includes("रात्रि") || c.includes("raatri") || c.includes("jk=h") || c.includes("रात")) nightTime = cell;
    }
  }

  const dataStart = headerIndex !== -1 ? headerIndex + 2 : 2;
  
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

    const colThana = String(row[thanaCol] || '').trim();
    const colSerial = String(row[serialCol] || '').trim();
    const colPlace = String(row[placeCol] || '').trim();
    const colDayName = String(row[dayNameCol] || '').trim();
    const colDayMob = String(row[dayMobileCol] || '').trim();
    const colNightName = String(row[nightNameCol] || '').trim();
    const colNightMob = String(row[nightMobileCol] || '').trim();

    if (colThana) lastThana = colThana;

    const isNew = colSerial !== '' && !isNaN(Number(colSerial.replace(/[^\d]/g, '')));

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

  const records: OfficerRecord[] = [];
  let idCounter = 0;

  for (const b of blocks) {
    if (b.dayMain.name) {
      records.push({
        id: `${sheetName}-${Date.now()}-${idCounter++}-day`,
        dutyType: sheetName || 'पुलिस ड्यूटी',
        mainOfficerName: b.dayMain.name,
        mainOfficerMobile: b.dayMain.mobile,
        supportingOfficers: b.daySupporting,
        dutyPlace: b.place,
        thanaArea: b.thana,
        dutyTime: dayTime,
        zonalMagistrate: '',
        zonalPoliceOfficer: '',
        sectorMagistrate: '',
        sectorPoliceOfficer: '',
      });
    }
    if (b.nightMain.name) {
      records.push({
        id: `${sheetName}-${Date.now()}-${idCounter++}-night`,
        dutyType: sheetName || 'पुलिस ड्यूटी',
        mainOfficerName: b.nightMain.name,
        mainOfficerMobile: b.nightMain.mobile,
        supportingOfficers: b.nightSupporting,
        dutyPlace: b.place,
        thanaArea: b.thana,
        dutyTime: nightTime,
        zonalMagistrate: '',
        zonalPoliceOfficer: '',
        sectorMagistrate: '',
        sectorPoliceOfficer: '',
      });
    }
  }

  return records;
}

/**
 * Type 2: Zonal/Sector Sheet Parser
 */
function parseZonalSectorSheet(rows: any[][], sheetName: string): OfficerRecord[] {
  let headerIndex = -1;
  let bestHeaderScore = 0;

  for (let r = 0; r < Math.min(10, rows.length); r++) {
    const row = rows[r];
    let score = 0;
    for (const cell of row) {
      const c = cell.toLowerCase();
      if (c.includes("थाना") || c.includes("Fkkuk")) score += 2;
      if (c.includes("जोन") || c.includes("tksu")) score += 2;
      if (c.includes("मजिस्ट्रेट") || c.includes("eftLVªsV")) score += 2;
      if (c.includes("सेक्टर") || c.includes("lsDVj")) score += 2;
      if (c.includes("मुख्यालय") || c.includes("eq[;ky;")) score += 2;
      if (c.includes("मो0") || c.includes("eks0") || c.includes("मो०") || c.includes("eks")) score += 2;
      if (c.includes("सं") || c.includes("dz") || c.includes("सं0") || c.includes("dz0")) score += 1;
    }
    if (score > bestHeaderScore) {
      bestHeaderScore = score;
      headerIndex = r;
    }
  }

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
      const c = cell.toLowerCase();
      if (c.includes("जोन") && !c.includes("जोनल") && !c.includes("मजिस्ट्रेट") && !c.includes("tksu") && !c.includes("eftLVªsV")) cols.zone = idx;
      if (c.includes("जोनल") || c.includes("tksuy")) cols.zonalMag = idx;
      if (c.includes("थाना") || c.includes("Fkkuk")) cols.thana = idx;
      if (c.includes("मुख्यालय") || c.includes("eq[;ky;") || c.includes("स्थान") || c.includes("LFkku")) cols.sectorHQ = idx;
      if (c.includes("सेक्टर") && c.includes("नाम") && !c.includes("अधिकारी") && !c.includes("मजिस्ट्रेट") && !c.includes("lsDVj")) cols.sectorName = idx;
      if (c.includes("सेक्टर") && (c.includes("मजिस्ट्रेट") || c.includes("eftLVªsV"))) cols.sectorMag = idx;
      if (c.includes("समय") || c.includes("le;")) cols.shiftTime = idx;
      if ((c.includes("अधिकारी") || c.includes("deZpkjh") || c.includes("deZpkjhx.k") || c.includes("नाम") || c.includes("uke")) && 
          !c.includes("थाना") && !c.includes("मजिस्ट्रेट") && !c.includes("मुख्यालय") && !c.includes("सेक्टर का नाम") && !c.includes("Fkkuk") && !c.includes("eftLVªsV")) {
        cols.officer = idx;
      }
      if (c.includes("मो0") || c.includes("eks0") || c.includes("मोबाइल") || c.includes("eksckbZy") || c.includes("मो०") || c.includes("eks")) cols.mobile = idx;
      if (c.includes("सेक्टर") && (c.includes("संख्या") || c.includes("सं0") || c.includes("dz0") || c.includes("la0"))) cols.sectorSerial = idx;
    });
  }

  let dataStart = headerIndex !== -1 ? headerIndex + 1 : 2;
  for (let i = dataStart; i < rows.length; i++) {
    const v = String(rows[i]?.[cols.sectorSerial] || '').trim();
    if (v !== '' && !isNaN(Number(v.replace(/[^\d]/g, '')))) {
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

    if (zone) lastZone = zone;
    if (thana) lastThana = thana;
    if (zonalMagRaw && zonalMagRaw !== lastZonalMagRaw) {
      lastZonalMagRaw = zonalMagRaw;
      const lines = zonalMagRaw.split('\n').map(l => l.trim()).filter(Boolean);
      lastZonalMag = '';
      lastZonalPolice = '';
      for (const line of lines) {
        if (line.includes("मजिस्ट्रेट") || line.includes("eftLVªsV")) {
          lastZonalMag = line;
        } else if (line.includes("पुलिस") || line.includes("iqfyl") || line.includes("अधिकारी") || line.includes("vf/kdkjh")) {
          lastZonalPolice = line;
        } else if (!lastZonalMag) {
          lastZonalMag = line;
        } else if (!lastZonalPolice) {
          lastZonalPolice = line;
        }
      }
    }

    if (!sectorSerial && !officerName && !sectorMag) continue;

    const isNewSector = sectorSerial !== '' && !isNaN(Number(sectorSerial.replace(/[^\d]/g, '')));

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

  const records: OfficerRecord[] = [];
  let idCounter = 0;

  for (const s of sectors) {
    if (s.dayMain.name) {
      const sectorPoliceOfficer = s.dayMain.name + (s.dayMain.mobile ? ` मो0नं0 ${s.dayMain.mobile}` : '');
      records.push({
        id: `${sheetName}-${Date.now()}-${idCounter++}-day`,
        dutyType: 'सेक्टर ड्यूटी',
        mainOfficerName: s.dayMain.name,
        mainOfficerMobile: s.dayMain.mobile,
        supportingOfficers: s.daySupporting,
        dutyPlace: s.sectorHQ || s.sectorName,
        thanaArea: s.thana,
        dutyTime: s.dayTime,
        zonalMagistrate: s.zonalMagistrate,
        zonalPoliceOfficer: s.zonalPoliceOfficer,
        sectorMagistrate: s.dayMagistrate,
        sectorPoliceOfficer: sectorPoliceOfficer,
      });
    }

    if (s.nightMain.name) {
      const sectorPoliceOfficer = s.nightMain.name + (s.nightMain.mobile ? ` मो0नं0 ${s.nightMain.mobile}` : '');
      records.push({
        id: `${sheetName}-${Date.now()}-${idCounter++}-night`,
        dutyType: 'सेक्टर ड्यूटी',
        mainOfficerName: s.nightMain.name,
        mainOfficerMobile: s.nightMain.mobile,
        supportingOfficers: s.nightSupporting,
        dutyPlace: s.sectorHQ || s.sectorName,
        thanaArea: s.thana,
        dutyTime: s.nightTime || 'रात्रि 20:00 बजे से प्रातः 08:00 बजे तक',
        zonalMagistrate: s.zonalMagistrate,
        zonalPoliceOfficer: s.zonalPoliceOfficer,
        sectorMagistrate: s.nightMagistrate || s.dayMagistrate,
        sectorPoliceOfficer: sectorPoliceOfficer,
      });
    }
  }

  return records;
}

/**
 * Type 3: Simple List Sheet Parser
 */
function parseSimpleListSheet(rows: any[][], sheetName: string): OfficerRecord[] {
  let headerIndex = -1;
  let bestHeaderScore = 0;

  for (let r = 0; r < Math.min(8, rows.length); r++) {
    const row = rows[r];
    let score = 0;
    for (const cell of row) {
      const c = cell.toLowerCase();
      if (c.includes("थाना") || c.includes("Fkkuk")) score += 2;
      if (c.includes("स्थान") || c.includes("LFkku") || c.includes("मार्ग") || c.includes("ekxZ") || c.includes("cSfj;j") || c.includes("fidsV") || c.includes("eafnj")) score += 2;
      if (c.includes("नाम") || c.includes("deZpkjh") || c.includes("vf/kdkjh") || c.includes("uke")) score += 2;
      if (c.includes("मो0") || c.includes("eks0") || c.includes("मोबाइल") || c.includes("eksckbZy") || c.includes("मो०") || c.includes("eks")) score += 2;
      if (c.includes("सं0") || c.includes("dz0") || c.includes("सं") || c.includes("dz") || c.includes("Ø")) score += 1;
    }
    if (score > bestHeaderScore) {
      bestHeaderScore = score;
      headerIndex = r;
    }
  }

  const cols = {
    serial: 0,
    thana: -1,
    place: -1,
    officer: -1,
    mobile: -1,
    time: -1,
    magistrate: -1,
  };

  const row1 = headerIndex !== -1 ? rows[headerIndex] : [];
  const row2 = headerIndex !== -1 && rows[headerIndex + 1] ? rows[headerIndex + 1] : [];

  for (let idx = 0; idx < Math.max(row1.length, row2.length); idx++) {
    const val1 = String(row1[idx] || '').toLowerCase();
    const val2 = String(row2[idx] || '').toLowerCase();

    if (val1.includes("सं") || val2.includes("सं") || val1.includes("dz") || val2.includes("dz") ||
        val1.includes("क्र") || val2.includes("क्र") || val1.includes("Ø") || val2.includes("Ø")) {
      cols.serial = idx;
    }
    if (val1.includes("थाना") || val2.includes("थाना") || val1.includes("Fkkuk") || val2.includes("Fkkuk")) {
      cols.thana = idx;
    }
    if (val1.includes("स्थान") || val2.includes("स्थान") || val1.includes("LFkku") || val2.includes("LFkku") ||
        val1.includes("मार्ग") || val2.includes("मार्ग") || val1.includes("ekxZ") || val2.includes("ekxZ") ||
        val1.includes("eafnj") || val2.includes("eafnj") || val1.includes("मन्दिर") || val2.includes("मन्दिर")) {
      cols.place = idx;
    }
    if (val1.includes("मजिस्ट्रेट") || val2.includes("मजिस्ट्रेट") || val1.includes("eftLVªsV") || val2.includes("eftLVªsV")) {
      cols.magistrate = idx;
    }
    if (((val1.includes("नाम") || val2.includes("नाम") || val1.includes("uke") || val2.includes("uke") ||
          val1.includes("deZpkjh") || val2.includes("deZpkjh") || val1.includes("vf/kdkjh") || val2.includes("vf/kdkjh") ||
          val1.includes("deZpkjhx.k") || val2.includes("deZpkjhx.k")) && 
         !val1.includes("मजिस्ट्रेट") && !val2.includes("मजिस्ट्रेट") && !val1.includes("eftLVªsV") && !val2.includes("eftLVªsV")) ||
        (val1.includes("fu;qDr") || val2.includes("fu;qDr"))) {
      if (cols.officer === -1) cols.officer = idx;
    }
    if (val1.includes("मो") || val2.includes("मो") || val1.includes("eks") || val2.includes("eks") ||
        val1.includes("मोबाइल") || val2.includes("मोबाइल") || val1.includes("eksckbZy") || val2.includes("eksckbZy")) {
      cols.mobile = idx;
    }
    if (val1.includes("समय") || val2.includes("समय") || val1.includes("le;") || val2.includes("le;")) {
      cols.time = idx;
    }
  }

  if (cols.place === -1) cols.place = Math.min(2, Math.max(0, rows[0]?.length - 1));
  if (cols.officer === -1) cols.officer = Math.min(3, Math.max(0, rows[0]?.length - 1));
  if (cols.mobile === -1) cols.mobile = Math.min(4, Math.max(0, rows[0]?.length - 1));
  if (cols.thana === -1) cols.thana = 0;

  const dataStart = headerIndex !== -1 ? headerIndex + 1 : 1;
  const records: OfficerRecord[] = [];
  let idCounter = 0;

  let curRecord: OfficerRecord | null = null;
  let lastThana = '';
  let lastPlace = '';
  let lastTime = 'सुबह 08.00 बजे से 20.00 बजे तक';

  for (let i = dataStart; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const colSerial = String(row[cols.serial] || '').trim();
    const colThana = cols.thana !== -1 ? String(row[cols.thana] || '').trim() : '';
    const colPlace = cols.place !== -1 ? String(row[cols.place] || '').trim() : '';
    const colName = String(row[cols.officer] || '').trim();
    const colMob = cols.mobile !== -1 ? String(row[cols.mobile] || '').trim() : '';
    const colTime = cols.time !== -1 ? String(row[cols.time] || '').trim() : '';
    const colMag = cols.magistrate !== -1 ? String(row[cols.magistrate] || '').trim() : '';

    if (colThana) lastThana = colThana;
    if (colPlace) lastPlace = colPlace;
    if (colTime) lastTime = colTime;

    const isNew = colSerial !== '' && !isNaN(Number(colSerial.replace(/[^\d]/g, '')));

    if (isNew || (!curRecord && colName)) {
      curRecord = {
        id: `${sheetName}-${Date.now()}-${idCounter++}`,
        dutyType: sheetName || 'पुलिस ड्यूटी',
        mainOfficerName: colName,
        mainOfficerMobile: colMob,
        supportingOfficers: [],
        dutyPlace: lastPlace,
        thanaArea: lastThana,
        dutyTime: lastTime,
        zonalMagistrate: '',
        zonalPoliceOfficer: '',
        sectorMagistrate: colMag,
        sectorPoliceOfficer: '',
      };
      records.push(curRecord);
    } else if (curRecord && colName) {
      curRecord.supportingOfficers.push({ name: colName, mobile: colMob });
    }
  }

  return records;
}
