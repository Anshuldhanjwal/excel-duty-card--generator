import * as XLSX from 'xlsx';
import type { OfficerRecord, ExtractionResult } from '../types';

/**
 * Universal Excel parser for Police Duty Card generation.
 * 
 * Supports two layouts:
 * 
 * FORMAT A — "Barrier-style" (7 cols, Hindi Unicode):
 *   [thana, serial, place, dayOfficerName, dayMobile, nightOfficerName, nightMobile]
 *   Header rows 0-2, data from row 3+. Serial number in col 1 marks new blocks.
 * 
 * FORMAT B — "Zonal/Sector-style" (11 cols, often Krutidev):
 *   [superzone, zone, zonalMagistrate, thana, sectorSerial, sectorHQ, sectorName,
 *    sectorMagistrate, shiftTime, officerName, mobile]
 *   Day and night rows are interleaved. Sector serial in col 4 marks new blocks.
 */
export async function parseExcelFile(file: File): Promise<ExtractionResult> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    blankrows: false,
  });

  if (rows.length === 0) {
    throw new Error('Excel फाइल खाली है।');
  }

  // --- Determine file format by column count ---
  const maxCols = Math.max(...rows.map((r) => r.length));

  // Extract dates if present
  let dutyDateFrom = '';
  let dutyDateTo = '';
  const dateRe = /(\d{1,2}[./]\d{1,2}[./]\d{4})/g;
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
      const m = text.match(/जनपद\s+(\S+)/);
      if (m) district = m[1].replace(/[^\u0900-\u097F]/g, '').trim() || district;
    }
  }

  if (maxCols >= 9) {
    // ========== FORMAT B: Zonal/Sector (11-col) ==========
    return parseFormatB(rows, district, dutyDateFrom, dutyDateTo);
  } else {
    // ========== FORMAT A: Barrier (7-col) ==========
    return parseFormatA(rows, district, dutyDateFrom, dutyDateTo);
  }
}

// ────────────────────────────────────────────────────────────
//  FORMAT A: Barrier-style (7 columns)
//  Cols: [thana, serial, place, dayOfficer, dayMobile, nightOfficer, nightMobile]
// ────────────────────────────────────────────────────────────
function parseFormatA(
  rows: any[][],
  district: string,
  dutyDateFrom: string,
  dutyDateTo: string
): ExtractionResult {
  // Row 0 = title, Row 1-2 = headers, Row 3+ = data
  const eventName = String(rows[0]?.[0] || '').trim() || 'ड्यूटी कार्ड';

  // Determine duty type from event name
  let dutyType = 'बैरियर ड्यूटी';
  if (eventName.includes('ड्यूटी')) {
    const m = eventName.match(/(.*?ड्यूटी)/);
    if (m) dutyType = m[1].trim();
  }

  // Find shift time labels from header row 1
  let dayShiftTime = 'सुबह 08.00 बजे से 20.00 बजे तक';
  let nightShiftTime = 'रात्री 20.00 बजे से 08.00 बजे तक';
  if (rows[1]) {
    const dayLabel = String(rows[1][3] || '').trim();
    const nightLabel = String(rows[1][5] || '').trim();
    if (dayLabel) dayShiftTime = dayLabel;
    if (nightLabel) nightShiftTime = nightLabel;
  }

  // Find data start (first row where col[1] is a number)
  let dataStart = 3;
  for (let i = 2; i < rows.length; i++) {
    const v = rows[i][1];
    if (v !== '' && !isNaN(Number(v))) {
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
    const colThana = String(row[0] || '').trim();
    const colSerial = String(row[1] || '').trim();
    const colPlace = String(row[2] || '').trim();
    const colDayName = String(row[3] || '').trim();
    const colDayMob = String(row[4] || '').trim();
    const colNightName = String(row[5] || '').trim();
    const colNightMob = String(row[6] || '').trim();

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
      // Supporting staff rows
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
        dutyType,
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
        dutyType,
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
  district: string,
  dutyDateFrom: string,
  dutyDateTo: string
): ExtractionResult {
  const eventName = String(rows[0]?.[0] || '').trim() || 'ड्यूटी कार्ड';

  // Column indices (fixed for this format)
  const COL_ZONE = 1;
  const COL_ZONAL_MAG = 2;
  const COL_THANA = 3;
  const COL_SECTOR_SERIAL = 4;
  const COL_SECTOR_HQ = 5;
  const COL_SECTOR_NAME = 6;
  const COL_SECTOR_MAG = 7;
  const COL_SHIFT_TIME = 8;
  const COL_OFFICER = 9;
  const COL_MOBILE = 10;

  // Find header row (contains something like "सं" or "eks0" or sector keywords)
  let dataStart = 2;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const text = rows[i].join(' ');
    if (text.includes('eks0ua0') || text.includes('मो0नं0') || text.includes('lsDVj') || text.includes('सेक्टर')) {
      dataStart = i + 1;
      break;
    }
  }

  // Find first data row (where COL_SECTOR_SERIAL is numeric)
  for (let i = dataStart; i < rows.length; i++) {
    const v = rows[i]?.[COL_SECTOR_SERIAL];
    if (v !== '' && v !== undefined && !isNaN(Number(v))) {
      dataStart = i;
      break;
    }
  }

  interface Sector {
    thana: string;
    zone: string;
    zonalMagistrate: string;
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
  let lastZonalMag = '';

  // Phase: 'day' or 'night'. The first officer row after a sector serial is day.
  // When COL_SHIFT_TIME changes, we switch to night.
  let phase: 'day' | 'night' = 'day';

  for (let i = dataStart; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const zone = String(row[COL_ZONE] || '').trim();
    const zonalMag = String(row[COL_ZONAL_MAG] || '').trim();
    const thana = String(row[COL_THANA] || '').trim();
    const sectorSerial = String(row[COL_SECTOR_SERIAL] || '').trim();
    const sectorHQ = String(row[COL_SECTOR_HQ] || '').trim();
    const sectorName = String(row[COL_SECTOR_NAME] || '').trim();
    const sectorMag = String(row[COL_SECTOR_MAG] || '').trim();
    const shiftTime = String(row[COL_SHIFT_TIME] || '').trim();
    const officerName = String(row[COL_OFFICER] || '').trim();
    const mobile = String(row[COL_MOBILE] || '').trim();

    // Track zone/thana carry-forward
    if (zone) lastZone = zone;
    if (zonalMag) lastZonalMag = zonalMag;
    if (thana) lastThana = thana;

    // Skip super-zone header rows (col 0 has text but no sector serial or officer)
    if (!sectorSerial && !officerName && !sectorMag) continue;

    const isNewSector = sectorSerial !== '' && !isNaN(Number(sectorSerial));

    if (isNewSector) {
      phase = 'day';
      cur = {
        thana: lastThana,
        zone: lastZone,
        zonalMagistrate: lastZonalMag,
        sectorHQ,
        sectorName,
        dayMagistrate: sectorMag,
        nightMagistrate: '',
        dayMain: { name: officerName, mobile },
        nightMain: { name: '', mobile: '' },
        daySupporting: [],
        nightSupporting: [],
        dayTime: shiftTime || 'izkr% 08%00 cts ls 20%00 cts rd',
        nightTime: '',
      };
      sectors.push(cur);
    } else if (cur) {
      // Detect night phase transition: if shiftTime is present and different from dayTime
      if (shiftTime && shiftTime !== cur.dayTime) {
        phase = 'night';
        cur.nightTime = shiftTime;
      }

      // If sector magistrate changes in a non-new row, it's the night magistrate
      if (sectorMag && phase === 'night') {
        cur.nightMagistrate = sectorMag;
      } else if (sectorMag && phase === 'day' && !cur.dayMagistrate) {
        cur.dayMagistrate = sectorMag;
      }

      if (officerName) {
        if (phase === 'day') {
          if (!cur.dayMain.name) {
            // This shouldn't happen but handle gracefully
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
      records.push({
        id: `${Date.now()}-${idCounter++}-day`,
        dutyType: 'बैरियर ड्यूटी',
        mainOfficerName: s.dayMain.name,
        mainOfficerMobile: s.dayMain.mobile,
        supportingOfficers: s.daySupporting,
        dutyPlace: s.sectorHQ,
        thanaArea: s.thana,
        dutyTime: s.dayTime,
        zonalMagistrate: '',
        zonalPoliceOfficer: '',
        sectorMagistrate: s.dayMagistrate,
        sectorPoliceOfficer: '',
      });
    }

    // Night shift
    if (s.nightMain.name) {
      records.push({
        id: `${Date.now()}-${idCounter++}-night`,
        dutyType: 'बैरियर ड्यूटी',
        mainOfficerName: s.nightMain.name,
        mainOfficerMobile: s.nightMain.mobile,
        supportingOfficers: s.nightSupporting,
        dutyPlace: s.sectorHQ,
        thanaArea: s.thana,
        dutyTime: s.nightTime,
        zonalMagistrate: '',
        zonalPoliceOfficer: '',
        sectorMagistrate: s.nightMagistrate,
        sectorPoliceOfficer: '',
      });
    }
  }

  return { eventName, district, dutyDateFrom, dutyDateTo, records };
}
