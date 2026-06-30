import React from 'react';
import { OfficerRecord } from '../types';

interface DutyCardProps {
  record: OfficerRecord;
  eventName: string;
  district: string;
  dutyDateFrom: string;
  dutyDateTo: string;
}

export const DutyCard: React.FC<DutyCardProps> = ({
  record,
  eventName,
  district,
  dutyDateFrom,
  dutyDateTo,
}) => {
  // Normalize header text and date formats to use dots and clean prefixes
  const cleanEvent = (eventName || '').trim();
  const cleanDistrict = (district || '').replace(/जनपद/g, '').replace(/[-—–\s]+/g, '').trim();
  const formattedDateFrom = (dutyDateFrom || '').replace(/[-/]/g, '.');
  const formattedDateTo = (dutyDateTo || '').replace(/[-/]/g, '.');

  // Ensure we have at least 2 supporting rows (matching original Excel card)
  const displaySupporting = [...record.supportingOfficers];
  while (displaySupporting.length < 2) {
    displaySupporting.push({ name: '', mobile: '', postingPlace: '' });
  }

  const N = displaySupporting.length;

  // Column A "ड्यूटी का प्रकार" spans: main officer + supporting header + all supporting officers
  const dutyTypeRowSpan = 1 + 1 + N; // = N + 2

  // Column E zonal magistrate VALUE spans: main officer + supporting header + (N-1) supporting
  // The last supporting row gets a new E cell (the "जोनल पुलिस अधिकारी" label)
  const zonalMagRowSpan = 1 + 1 + (N - 1); // = N + 1

  return (
    <div
      className="duty-card-container w-full max-w-[850px] bg-white border-2 border-black text-black mx-auto select-none print:max-w-full print:h-full"
      style={{ fontFamily: "'Noto Sans Devanagari', 'Mangal', sans-serif" }}
    >
      {/* SINGLE TABLE — mirrors the Excel 5-column layout exactly */}
      <table className="w-full h-full border-collapse" style={{ tableLayout: 'fixed' }}>
        <colgroup>
          <col style={{ width: '15%' }} />
          <col style={{ width: '25%' }} />
          <col style={{ width: '15%' }} />
          <col style={{ width: '15%' }} />
          <col style={{ width: '30%' }} />
        </colgroup>
        <tbody>

          {/* ============================================ */}
          {/* HEADER ROW: Title text (A-D merged) + Logo (E) */}
          {/* ============================================ */}
          <tr>
            <td
              colSpan={4}
              className="text-center px-3 py-2.5 align-middle border-r border-b border-black"
            >
              <div className="text-[29px] font-extrabold leading-tight tracking-wide">
                {cleanEvent || 'काँवड़ यात्रा-2026'}
              </div>
              <div className="text-[23px] font-bold leading-tight mt-1">
                ड्यूटी कार्ड जनपद {cleanDistrict || 'बुलन्दशहर'}
              </div>
              <div className="text-[21px] font-bold leading-tight mt-1">
                ड्यूटी दिनांक {formattedDateFrom || '05.08.2026'} से {formattedDateTo || '12.08.2026'} तक
              </div>
            </td>
            <td
              className="text-center align-middle border-b border-black p-2"
            >
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img
                src="/up_police_logo.png"
                alt="UP Police Logo"
                className="mx-auto"
                style={{ width: '105px', height: 'auto' }}
              />
            </td>
          </tr>

          {/* ============================================ */}
          {/* TABLE ROW 1: Column headers (Bold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-2 text-center font-bold align-middle text-[18px]">
              ड्यूटी का प्रकार
            </td>
            <td className="border-b border-r border-black px-2 py-2 text-center font-bold align-middle text-[18px]">
              ड्यूटी अधिकारी/कर्मचारी का नाम
            </td>
            <td className="border-b border-r border-black px-2 py-2 text-center font-bold align-middle text-[18px]">
              नियुक्ति स्थान
            </td>
            <td className="border-b border-r border-black px-2 py-2 text-center font-bold align-middle text-[18px]">
              मो0नं0
            </td>
            <td className="border-b border-l-2 border-black px-2 py-2 text-center font-bold align-middle text-[18px]">
              जोनल मजि0 का नाम व मो0नं0
            </td>
          </tr>

          {/* ============================================ */}
          {/* TABLE ROW 2: Main duty officer details */}
          {/* ============================================ */}
          <tr>
            {/* Column A: Duty type value (Unbold) */}
            <td
              rowSpan={dutyTypeRowSpan}
              className="border-b border-r border-black px-2 py-2 font-normal text-center align-middle text-[18px]"
            >
              {record.dutyType}
            </td>
            {/* Column B: Officer name (Bold for visibility) */}
            <td className="border-b border-r border-black px-2 py-2 font-bold align-middle text-[20px]">
              {record.mainOfficerName}
            </td>
            {/* Column C: Posting place (Unbold) */}
            <td className="border-b border-r border-black px-2 py-2 font-normal text-center align-middle text-[18px]">
              {record.mainOfficerPostingPlace || '\u00A0'}
            </td>
            {/* Column D: Mobile (Unbold) */}
            <td className="border-b border-r border-black px-2 py-2 font-normal text-center align-middle text-[18px]">
              {record.mainOfficerMobile}
            </td>
            {/* Column E: Zonal magistrate value (Unbold) */}
            <td
              rowSpan={zonalMagRowSpan}
              className="border-b border-l-2 border-black px-3 py-2 font-normal text-center align-middle text-[18px]"
            >
              {record.zonalMagistrate || '\u00A0'}
            </td>
          </tr>

          {/* ============================================ */}
          {/* TABLE ROW 3: Supporting staff sub-headers (Bold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-1 font-bold text-center align-middle text-[17px]">
              सहयोगी पुलिसकर्मियों के नाम
            </td>
            <td className="border-b border-r border-black px-2 py-1 font-bold text-center align-middle text-[17px]">
              नियुक्ति स्थान
            </td>
            <td className="border-b border-r border-black px-2 py-1 font-bold text-center align-middle text-[17px]">
              मो0नं0
            </td>
            {/* Column E covered by zonal magistrate rowSpan */}
          </tr>

          {/* ============================================ */}
          {/* TABLE ROWS 4+: Supporting officers */}
          {/* ============================================ */}
          {displaySupporting.map((officer, index) => {
            const isLast = index === N - 1;
            return (
              <tr key={index}>
                {/* Column B: Supporting Officer Name (Unbold as requested) */}
                <td className="border-b border-r border-black px-2 py-2 font-normal align-middle text-[18px]">
                  {officer.name || '\u00A0'}
                </td>
                {/* Column C: Posting Place (Unbold) */}
                <td className="border-b border-r border-black px-2 py-2 font-normal text-center align-middle text-[18px]">
                  {officer.postingPlace || '\u00A0'}
                </td>
                {/* Column D: Mobile (Unbold) */}
                <td className="border-b border-r border-black px-2 py-2 font-normal text-center align-middle text-[18px]">
                  {officer.mobile || '\u00A0'}
                </td>
                {/* Column E: Only the LAST supporting row gets a new E cell */}
                {isLast && (
                  <td className="border-b border-l-2 border-black px-2 py-2 text-center align-middle font-bold text-[16px]">
                    जोनल पुलिस अधिकारी नाम व मो0नं0
                  </td>
                )}
                {/* All other rows: Column E is covered by the zonal magistrate rowSpan */}
              </tr>
            );
          })}

          {/* ============================================ */}
          {/* ड्यूटी का स्थान — Label row (Bold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-2 font-bold align-middle text-[18px]">
              ड्यूटी का स्थान
            </td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            {/* Column E: Zonal police officer value (Unbold) */}
            <td
              rowSpan={2}
              className="border-b border-l-2 border-black px-3 py-2 font-normal text-center align-middle text-[18px]"
            >
              {record.zonalPoliceOfficer || '\u00A0'}
            </td>
          </tr>

          {/* ============================================ */}
          {/* ड्यूटी का स्थान — Value row (Unbold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-2 font-normal align-middle text-[18px]">
              {record.dutyPlace}
            </td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
          </tr>

          {/* ============================================ */}
          {/* थाना क्षेत्र — Label row (Bold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-2 font-bold align-middle text-[18px]">
              थाना क्षेत्र
            </td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            {/* Column E: Sector magistrate LABEL (Bold) */}
            <td className="border-b border-l-2 border-black px-2 py-2 text-center align-middle font-bold text-[16px]">
              सेक्टर मजि0 का नाम व मो0नं0
            </td>
          </tr>

          {/* ============================================ */}
          {/* थाना क्षेत्र — Value row (Unbold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-2 font-normal align-middle text-[18px]">
              {record.thanaArea}
            </td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            {/* Column E: Sector magistrate value (Unbold) */}
            <td
              rowSpan={2}
              className="border-b border-l-2 border-black px-3 py-2 font-normal text-center align-middle text-[18px]"
            >
              {record.sectorMagistrate || '\u00A0'}
            </td>
          </tr>

          {/* ============================================ */}
          {/* ड्यूटी का समय — Label row (Bold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-2 font-bold align-middle text-[18px]">
              ड्यूटी का समय
            </td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
          </tr>

          {/* ============================================ */}
          {/* ड्यूटी का समय — Value row (Unbold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-2 font-normal align-middle text-[18px] leading-snug">
              {record.dutyTime}
            </td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-2">{'\u00A0'}</td>
            {/* Column E: Sector police officer LABEL (Bold) */}
            <td className="border-b border-l-2 border-black px-2 py-2 text-center align-middle font-bold text-[16px]">
              सेक्टर पुलिस अधिकारी का नाम/मो0नं0
            </td>
          </tr>

          {/* ============================================ */}
          {/* Spacer row + Sector police officer value (Unbold) */}
          {/* ============================================ */}
          <tr>
            <td className="border-b border-r border-black px-2 py-3">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-3">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-3">{'\u00A0'}</td>
            <td className="border-b border-r border-black px-2 py-3">{'\u00A0'}</td>
            <td className="border-b border-l-2 border-black px-3 py-3 font-normal text-center align-middle text-[18px]">
              {record.sectorPoliceOfficer || '\u00A0'}
            </td>
          </tr>

          {/* ============================================ */}
          {/* FOOTER ROW */}
          {/* ============================================ */}
          <tr>
            <td colSpan={5} className="px-4 pt-3 pb-2 text-right align-bottom">
              <span className="font-bold text-[18px] leading-snug">
                वरिष्ठ पुलिस अधीक्षक,
              </span>
              <br />
              <span className="font-bold text-[18px] leading-snug">
                जनपद—बुलन्दशहर
              </span>
            </td>
          </tr>

        </tbody>
      </table>
    </div>
  );
};
