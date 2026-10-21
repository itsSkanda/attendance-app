import React, { useState, useCallback, useRef, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { FILLED_TEMPLATE_BASE64 } from './filledTemplateBase64';
import {
  Upload, FileSpreadsheet, X, AlertTriangle, Filter, Users,
  BookOpen, Search, Trash2, ChevronDown, ChevronUp, Check,
  SlidersHorizontal, RotateCcw, Moon, Sun, Download, CheckCircle2,
} from 'lucide-react';

// Subject Code to Name dictionary (strictly confirmed subjects only; unknown codes remain empty)
export const SUBJECT_NAMES = {
  // Verified from user's uploaded syllabus
  '21CSC302J': 'Computer Networks',
  '21IPE312P': 'ERP Solutions for Digital Enterprises',
  '21MAB302T': 'Discrete Mathematics',
  '21CSC301T': 'Formal Language and Automata',
  '21CSC305P': 'Machine Learning',
  '21CEO309T': 'Water Pollution and its Management',
  '21HCSF002': 'Introduction to Parallel Programming with OpenMP and MPI Using C and Python',
  '21HCSF003': 'Fundamentals of Accelerated Computing with CUDA Using Python',
  '21GNP301L': 'Community Connect',
  '21LEM301T': 'Indian Art Form',

  // Verified from filled-template.xlsx
  '21CSE354T': 'Full Stack Web Development',
  '21MEO113J': 'Electronics Thermal Management',
  '21CSE742P': 'ADVANCED SQL AND MODERN DATABASE FEATURES',
  '21CSE745P': 'FOUNDATIONS OF DATA CENTRE INFRASTRUCTURE AND MANAGEMENT',
  '21CSE296P': 'IOS APP PROTOTYPE DESIGN AND DEVELOPMENT',
  '21MGO112T': 'ENTREPRENEURIAL FINANCE',

  // Professional & Open Electives (AY 2026-27 ODD - Computer Science and Engineering)
  '21CSE214T': 'GPU Computing',
  '21CSE216P': 'Linux and Container Technologies',
  '21CSE305P': 'SerBot : Project-Based Learning in Robotics',
  '21CSE306P': 'Applied Generative AI',
  '21CSE307J': 'Quantum Machine Learning',
  '21CSE311J': 'Cloud Based Remote Sensing and GIS',
  '21CSE351T': 'Computational Logic',
  '21CSE357T': 'Distributed Computing',
  '21CSE359T': 'Information Storage and Management',
  '21CSE361T': 'Database Security and Privacy',
  '21CSE389T': 'Environmental Hydrology For Data Science',
  '21CSE735P': 'Workflow Automation Engineering',
  '21CSE737J': 'IoT data analysis with Cloud ML',
  '21CSE744P': 'Building AI Agents for Digital Customer Experience Systems',
  '21IPE335J': 'Functional Testing and Programming',
  '21IPE341P': 'ERP Application Development and Enterprise Data Platforms',
  '21IPE341T': 'Mastering the Data Paradox with AI and ML',
  '21IPE401J': 'Capstone Program in Electronic Cooling',
  '21ASO301T': 'Elements of Aeronautics',
  '21AUO101T': 'Hybrid and Electric Vehicles',
  '21BMO121T': 'Fundamentals of Biomedical Engineering',
  '21BTO106T': 'Waste to Wealth to Wheels',
  '21CEO303T': 'Smart City and Infrastructure',
  '21CEO306T': 'Environmental Impact Assessment',
  '21CEO307T': 'Municipal Solid Waste Management',
  '21CEO308T': 'Disaster Mitigation and Management',
  '21CEO310T': 'Global Warming and Climate Change',
  '21CEO322T': 'Remote sensing and GIS applications in Engineering',
  '21CEO403T': 'Integrated Waste Management',
  '21CEO404T': 'Principles of Sustainable Development',
  '21ECO101T': 'Short-Range Wireless Communication',
  '21ECO103T': 'Modern Wireless Communication Systems',
  '21ECO105T': 'Fiber Optics and Optoelectronics',
  '21ECO110T': 'Semiconductor Packaging Technologies',
  '21EEO303T': 'E-waste Management',
  '21EEO307T': 'Clean and Green Energy',
  '21EEO308T': 'Smart Cities and Communities',
  '21GEO105T': 'Biology for Everyday life',
  '21MEO101T': 'Fundamentals of Composite Materials',
  '21MEO102T': 'Reverse Engineering and 3D Printing',
  '21MEO103T': 'Fundamentals of Biomechanics',
  '21MEO104T': 'TQM and Reliability Engineering',
  '21MEO105T': 'Occupational Safety and Disaster Management',
  '21MEO106T': 'Introduction to Robotics',
  '21MEO110T': 'Energy Systems for Sustainable Buildings',
  '21MEO112T': 'Renewable Energy Sources and Application',
  '21MEO114T': 'Solar Energy for Societal Applications',
  '21MEO115T': 'Introduction to Drones',
  '21MGO103T': 'Management Information Systems',
  '21MGO107T': 'Flexible Work Models and Employee Well Being',
  '21NTO301T': 'Applications of Nanotechnology',
  '21NTO304T': 'Environmental Nanotechnology',
  '21NTO307T': 'Nanocomputing',
  '21NTO308T': 'Smart Sensor Systems',
  '21NTO309T': '2D Materials and Applications',
  // Aliases in case typed with leading I
  '21INTO301T': 'Applications of Nanotechnology',
  '21INTO304T': 'Environmental Nanotechnology',
  '21INTO307T': 'Nanocomputing',
  '21INTO308T': 'Smart Sensor Systems',
  '21INTO309T': '2D Materials and Applications',
};

// ─────────────────────────────────────────────────────────────────────────────
// PARSERS
// ─────────────────────────────────────────────────────────────────────────────

const clean = (v) => (v == null ? '' : String(v).replace(/[\r\n]+/g, ' ').trim());

export function getBaseCourseCode(code) {
  if (!code) return '';
  return clean(code)
    .toUpperCase()
    .replace(/[\s_\-/]*\((?:THEORY|PRACTICAL|LAB|LABORATORY|TH|PR|T|L|P)\)$/i, '')
    .replace(/[\s_\-/]+(?:THEORY|PRACTICAL|LAB|LABORATORY|TH|PR|T|L|P)$/i, '')
    .trim();
}

export function getCourseComponent(code, name = '') {
  const combined = `${code || ''} ${name || ''}`.toUpperCase();
  if (/\b(?:PRACTICAL|LAB|LABORATORY)\b|\((?:P|L|PRAC)\)|\/(?:L|P)\b/i.test(combined)) {
    return 'L'; // Lab / Practical
  }
  if (/\b(?:THEORY|TH)\b|\((?:T|TH)\)|\/T\b/i.test(combined)) {
    return 'T'; // Theory
  }
  return null;
}

export function normalizeSubjectCode(code, name = '') {
  if (!code) return '';
  const cleanCode = clean(code).toUpperCase();
  const base = getBaseCourseCode(cleanCode);
  const comp = getCourseComponent(cleanCode, name);
  if (comp) {
    return `${base}(${comp})`;
  }
  return cleanCode;
}

export function resolveSubjectName(code, subName = '') {
  if (subName && subName.trim()) return subName.trim();
  if (!code) return '';
  const cleanCode = clean(code).toUpperCase().trim();
  if (SUBJECT_NAMES[cleanCode]) return SUBJECT_NAMES[cleanCode];

  const baseFromParens = cleanCode.split('(')[0].trim();
  const base = getBaseCourseCode(cleanCode);
  const baseName = SUBJECT_NAMES[baseFromParens] || SUBJECT_NAMES[base];
  if (baseName) {
    const comp = getCourseComponent(cleanCode, subName);
    if (comp && (baseFromParens.endsWith('J') || (base && base.endsWith('J')))) {
      if (comp === 'L') return `${baseName} (Practical)`;
      if (comp === 'T') return `${baseName} (Theory)`;
    }
    return baseName;
  }
  return '';
}

export function disambiguateStudentSubjects(subjects = []) {
  if (!subjects || subjects.length === 0) return [];
  const baseCounts = {};
  subjects.forEach((s) => {
    const base = getBaseCourseCode(s.code);
    baseCounts[base] = (baseCounts[base] || 0) + 1;
  });

  const occurrence = {};
  return subjects.map((s) => {
    const base = getBaseCourseCode(s.code);
    const existingComp = getCourseComponent(s.code, s.name);

    if (existingComp) {
      const code = `${base}(${existingComp})`;
      const name = resolveSubjectName(code, s.name);
      return { ...s, code, name };
    }

    if (baseCounts[base] > 1) {
      occurrence[base] = (occurrence[base] || 0) + 1;
      const comp = occurrence[base] === 1 ? 'T' : 'L';
      const code = `${base}(${comp})`;
      const name = resolveSubjectName(code, s.name);
      return { ...s, code, name };
    }

    return s;
  });
}

/**
 * FORMAT A — Full-section attendance (e.g. P1.xlsx)
 * Row 0: headers; Row 1+: data. Subject cells packed as "CODE\r\nPCT".
 */
function parseFormatA(rows) {
  const students = [];
  const headerRow = rows[0] || [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every((c) => c == null || c === '')) continue;
    const snoRaw = r[0];
    const regNo = clean(r[1]);
    const name = clean(r[2]);
    if (!snoRaw || !regNo || !name) continue;
    const sno = parseInt(String(snoRaw), 10);
    if (isNaN(sno)) continue;
    const subjects = [];
    for (let c = 3; c < r.length; c++) {
      const cell = r[c];
      if (cell == null || cell === '') continue;
      const cellStr = String(cell).trim();
      if (!cellStr) continue;

      let code = '';
      let pct = null;

      if (cellStr.includes('\n') || cellStr.includes('\r')) {
        const parts = cellStr.split(/[\r\n]+/).map((p) => p.trim()).filter(Boolean);
        if (parts.length >= 2) {
          code = parts[0];
          const parsed = parseFloat(parts[1].replace('%', '').trim());
          pct = isNaN(parsed) ? 0 : parsed;
        } else if (parts.length === 1) {
          code = parts[0];
          pct = 0;
        }
      } else if (/^[A-Z0-9_-]+\s+[0-9.]+%?$/i.test(cellStr)) {
        const parts = cellStr.split(/\s+/);
        code = parts[0].trim();
        const parsed = parseFloat(parts[1].replace('%', '').trim());
        pct = isNaN(parsed) ? 0 : parsed;
      } else if (!isNaN(parseFloat(cellStr.replace('%', '').trim()))) {
        const hdrVal = headerRow[c] ? clean(headerRow[c]) : '';
        const hdrCode = hdrVal.split(/[\r\n\s(]+/)[0].trim().toUpperCase();
        if (hdrCode && /^[A-Z0-9]{4,15}$/i.test(hdrCode)) {
          code = hdrCode;
          pct = parseFloat(cellStr.replace('%', '').trim());
        }
      } else {
        code = cellStr;
        pct = 0;
      }

      if (code) {
        subjects.push({ code, pct: pct !== null && !isNaN(pct) ? pct : 0 });
      }
    }
    students.push({ sno, regNo, name, subjects: disambiguateStudentSubjects(subjects) });
  }
  return students;
}

/**
 * FORMAT B — Low-attendance report (e.g. P1_Low_Attendance_Fifth_Semester.xlsx)
 * Has institution header block; students can span multiple rows for subjects.
 */
function parseFormatB(rows) {
  let headerIdx = -1;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const r = rows[i];
    if (!r) continue;
    const hasSnO = r.some((c) => clean(c) === 'S.No');
    const hasReg = r.some((c) => clean(c).toLowerCase().includes('register'));
    if (hasSnO && hasReg) { headerIdx = i; break; }
  }
  if (headerIdx === -1) return { meta: {}, students: [] };

  const meta = {};
  for (let i = 0; i < headerIdx; i++) {
    const r = rows[i];
    if (!r) continue;
    r.forEach((cell, ci) => {
      const s = clean(cell).toLowerCase();
      if (s.includes('section:')) {
        const val = clean(cell).replace(/section:/i, '').trim();
        meta.section = val || clean(r[ci + 1]);
      }
      if (s.includes('faculty advisor:')) {
        const val = clean(cell).replace(/faculty advisor:/i, '').trim();
        meta.facultyAdvisor = val || clean(r[ci + 1]);
      }
      if (s.includes('academic advisor:')) {
        const val = clean(cell).replace(/academic advisor:/i, '').trim();
        meta.academicAdvisor = val || clean(r[ci + 1]);
      }
      if (s.includes('total number')) {
        for (let j = ci; j < Math.min(ci + 4, r.length); j++) {
          const sv = clean(r[j]);
          const m = sv.match(/students[:\s]+(\d+)/i);
          if (m) { meta.totalStudents = m[1]; break; }
        }
      }
      if (s.includes('less than 75%')) {
        for (let j = ci + 1; j < Math.min(ci + 4, r.length); j++) {
          const v = clean(r[j]);
          if (v && !isNaN(parseInt(v))) { meta.lowCount = v; break; }
        }
      }
    });
  }

  const hdr = rows[headerIdx];
  const colSno = hdr.findIndex((c) => clean(c) === 'S.No');
  const colReg = hdr.findIndex((c) => clean(c).toLowerCase().includes('register'));
  const colName = hdr.findIndex((c) => clean(c).toLowerCase().includes('student name'));
  const colCode = hdr.findIndex((c) => clean(c).toLowerCase().includes('subject code'));
  const colSubName = hdr.findIndex((c) => clean(c).toLowerCase().includes('subject name'));
  const colPct = hdr.findIndex((c) => clean(c).toLowerCase().includes('attendance'));
  const colMail = hdr.findIndex((c) => clean(c).toLowerCase().includes('mail'));
  const colParent = hdr.findIndex((c) => clean(c).toLowerCase().includes('parent'));

  const students = [];
  let current = null;

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every((c) => c == null || c === '')) continue;
    const snoRaw = colSno >= 0 ? r[colSno] : null;
    const regRaw = colReg >= 0 ? r[colReg] : null;
    const nameRaw = colName >= 0 ? r[colName] : null;
    const codeRaw = colCode >= 0 ? r[colCode] : null;
    const subNameRaw = colSubName >= 0 ? r[colSubName] : null;
    const pctRaw = colPct >= 0 ? r[colPct] : null;
    const mailRaw = colMail >= 0 ? r[colMail] : null;
    const parentRaw = colParent >= 0 ? r[colParent] : null;

    const pctClean = pctRaw != null ? String(pctRaw).replace('%', '').trim() : '';
    const pctParsed = pctClean !== '' ? parseFloat(pctClean) : null;
    const pct = pctParsed !== null && !isNaN(pctParsed) ? pctParsed : (pctRaw != null && pctRaw !== '' ? 0 : null);
    const code = clean(codeRaw);
    const subName = clean(subNameRaw);

    if (snoRaw != null && snoRaw !== '' && !isNaN(parseInt(String(snoRaw)))) {
      const regNo = clean(regRaw);
      const name = clean(nameRaw);
      if (regNo && name) {
        current = { sno: parseInt(String(snoRaw), 10), regNo, name, subjects: [] };
        students.push(current);
      }
    }

    if (current && code) {
      current.subjects.push({
        code,
        name: subName,
        pct: isNaN(pct) ? null : pct,
        mailSent: clean(mailRaw),
        parentsResponse: clean(parentRaw),
      });
    }
  }

  students.forEach((st) => {
    st.subjects = disambiguateStudentSubjects(st.subjects);
  });

  return { meta, students };
}

function detectFormat(rows) {
  if (!rows || rows.length === 0) return null;
  const r0 = rows[0] || [];
  const first = clean(r0[0]).toUpperCase();
  if (first === 'S.NO') return 'A';
  if (first.includes('SRM') || first.includes('COLLEGE') || first.includes('DEPARTMENT') || first.includes('INSTITUTE')) {
    for (let i = 1; i < Math.min(12, rows.length); i++) {
      const r = rows[i] || [];
      const hasSnO = r.some((c) => clean(c) === 'S.No');
      const hasSubjectName = r.some((c) => clean(c).toLowerCase().includes('subject name'));
      if (hasSnO && hasSubjectName) return 'B';
    }
  }
  return null;
}

function parseWorkbook(wb) {
  const result = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });
    const fmt = detectFormat(rows);
    if (!fmt) continue;
    if (fmt === 'A') {
      const students = parseFormatA(rows);
      if (students.length > 0) result.push({ sheetName, format: 'A', meta: {}, students });
    } else if (fmt === 'B') {
      const { meta, students } = parseFormatB(rows);
      if (students.length > 0) result.push({ sheetName, format: 'B', meta, students });
    }
  }
  return result;
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

const LOW = 75;

function getLowSubjects(student, deselectedSubjects = new Set()) {
  return (student.subjects || []).filter((s) => {
    const norm = normalizeSubjectCode(s.code, s.name);
    if (deselectedSubjects.has(norm)) return false;
    return s.pct !== null && s.pct < LOW;
  });
}

function hasLowAttendance(student, deselectedSubjects = new Set()) {
  return getLowSubjects(student, deselectedSubjects).length > 0;
}

// ─────────────────────────────────────────────────────────────────────────────
// SUB-COMPONENTS
// ─────────────────────────────────────────────────────────────────────────────

function PctPill({ pct, excluded = false }) {
  if (pct === null || pct === undefined || isNaN(pct)) {
    return <span style={{ color: 'var(--text-muted)', fontSize: '12px' }}>—</span>;
  }
  if (excluded) {
    return (
      <span className="pct-pill" style={{
        color: 'var(--text-muted)', background: 'var(--bg-hover)',
        border: '1px dashed var(--border)', fontStyle: 'italic',
      }}>
        {pct.toFixed(2)}%
      </span>
    );
  }
  const isLow = pct < LOW;
  const isWarn = pct >= LOW && pct < 85;
  const color = isLow ? 'var(--red)' : isWarn ? 'var(--amber)' : 'var(--green)';
  const bg = isLow ? 'var(--red-dim)' : isWarn ? 'var(--amber-dim)' : 'var(--green-dim)';
  const borderColor = isLow ? 'rgba(239,68,68,0.3)' : isWarn ? 'rgba(245,158,11,0.25)' : 'rgba(34,197,94,0.25)';
  return (
    <span className="pct-pill" style={{ color, background: bg, border: `1px solid ${borderColor}` }}>
      {isLow && <AlertTriangle size={9} />}
      {pct.toFixed(2)}%
    </span>
  );
}

// Format A: table row with subject pills
function FormatARow({ student, filterLow, deselectedSubjects = new Set() }) {
  const lowSubs = getLowSubjects(student, deselectedSubjects);
  const displaySubs = filterLow ? lowSubs : student.subjects;
  if (filterLow && lowSubs.length === 0) return null;

  return (
    <tr className={lowSubs.length > 0 ? 'row-danger' : ''}>
      <td className="td td-sno">{student.sno}</td>
      <td className="td td-mono">{student.regNo}</td>
      <td className="td td-name">{student.name}</td>
      <td className="td">
        <div className="sub-pills">
          {displaySubs.map((sub, idx) => {
            const normCode = normalizeSubjectCode(sub.code, sub.name);
            const isExcluded = deselectedSubjects.has(normCode);
            const isLow = !isExcluded && sub.pct !== null && sub.pct < LOW;
            return (
              <div
                key={idx}
                className="sub-pill"
                style={{
                  border: isExcluded
                    ? '1px dashed var(--border)'
                    : `1px solid ${isLow ? 'rgba(239,68,68,0.3)' : 'var(--border)'}`,
                  background: isExcluded
                    ? 'transparent'
                    : isLow ? 'rgba(239,68,68,0.06)' : 'var(--bg-hover)',
                  opacity: isExcluded ? 0.55 : 1,
                }}
                title={isExcluded ? `${sub.code} excluded from evaluation` : (sub.name ? `${sub.code} — ${sub.name}` : sub.code)}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <span
                    className="sub-pill-code"
                    style={{ textDecoration: isExcluded ? 'line-through' : 'none' }}
                  >
                    {sub.code}
                  </span>
                  {isExcluded && (
                    <span style={{
                      fontSize: '8px', padding: '1px 3px', borderRadius: '3px',
                      background: 'var(--bg-elevated)', color: 'var(--text-muted)', fontWeight: 600,
                    }}>OFF</span>
                  )}
                </div>
                <PctPill pct={sub.pct} excluded={isExcluded} />
              </div>
            );
          })}
        </div>
      </td>
      <td className="td td-center">
        {lowSubs.length > 0 ? (
          <span className="low-badge">{lowSubs.length}</span>
        ) : (
          <span style={{ color: 'var(--green)', fontSize: '14px', fontWeight: 700 }}>✓</span>
        )}
      </td>
    </tr>
  );
}

// Format B: student card with subject detail rows
function FormatBCard({ student, filterLow, deselectedSubjects = new Set() }) {
  const lowSubs = getLowSubjects(student, deselectedSubjects);
  const displaySubs = filterLow ? lowSubs : student.subjects;
  if (filterLow && lowSubs.length === 0) return null;
  const hasMailCol = student.subjects.some((s) => s.mailSent !== undefined && s.mailSent !== '');

  return (
    <div className={`student-card${lowSubs.length > 0 ? ' has-low' : ''}`}>
      <div className="student-card-header">
        <div className="student-avatar">{student.sno}</div>
        <div className="student-info">
          <div className="student-name">{student.name}</div>
          <div className="student-reg">{student.regNo}</div>
        </div>
        {lowSubs.length > 0 ? (
          <span className="low-badge">{lowSubs.length} below 75%</span>
        ) : (
          <span className="ok-badge">All ≥ 75%</span>
        )}
      </div>

      <div style={{ overflowX: 'auto' }}>
        <table className="data-table" style={{ minWidth: 'auto' }}>
          <thead>
            <tr style={{ background: 'var(--bg-elevated)' }}>
              <th className="th">Subject Code</th>
              <th className="th">Subject Name</th>
              <th className="th" style={{ textAlign: 'center' }}>Attendance</th>
              {hasMailCol && <th className="th" style={{ textAlign: 'center' }}>Mail Sent</th>}
              {hasMailCol && <th className="th" style={{ textAlign: 'center' }}>Parent Response</th>}
            </tr>
          </thead>
          <tbody>
            {displaySubs.map((sub, idx) => {
              const normCode = normalizeSubjectCode(sub.code, sub.name);
              const isExcluded = deselectedSubjects.has(normCode);
              const isLow = !isExcluded && sub.pct !== null && sub.pct < LOW;
              return (
                <tr
                  key={idx}
                  style={{
                    borderTop: '1px solid var(--border)',
                    background: isLow ? 'rgba(239,68,68,0.03)' : 'transparent',
                    opacity: isExcluded ? 0.55 : 1,
                  }}
                >
                  <td className="td td-mono">
                    <span style={{ textDecoration: isExcluded ? 'line-through' : 'none' }}>{sub.code}</span>
                    {isExcluded && (
                      <span style={{
                        marginLeft: '6px', fontSize: '10px', padding: '1px 5px',
                        borderRadius: '4px', background: 'var(--bg-elevated)', color: 'var(--text-muted)',
                      }}>Excluded</span>
                    )}
                  </td>
                  <td className="td" style={{ color: 'var(--text-primary)' }}>{sub.name || '—'}</td>
                  <td className="td td-center"><PctPill pct={sub.pct} excluded={isExcluded} /></td>
                  {hasMailCol && (
                    <td className="td td-center" style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                      {sub.mailSent || '—'}
                    </td>
                  )}
                  {hasMailCol && (
                    <td className="td td-center" style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                      {sub.parentsResponse || '—'}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// Sheet display — wraps both formats with meta banner + filters
function SheetDisplay({ sheet, filterLow, searchQuery, deselectedSubjects = new Set() }) {
  const { sheetName, format, meta, students } = sheet;

  const visible = students.filter((st) => {
    const q = searchQuery.toLowerCase();
    if (q && !st.name.toLowerCase().includes(q) && !st.regNo.toLowerCase().includes(q)) return false;
    if (filterLow && !hasLowAttendance(st, deselectedSubjects)) return false;
    return true;
  });

  const lowCount = students.filter((st) => hasLowAttendance(st, deselectedSubjects)).length;

  return (
    <div className="sheet-block">
      <div className="sheet-meta-banner">
        <div>
          <div className="sheet-meta-name-row">
            <BookOpen size={15} style={{ color: 'var(--green)' }} />
            <span className="sheet-name">{sheetName}</span>
          </div>
          <div className="sheet-meta-chips">
            {meta.section && <span className="meta-chip"><strong>Section:</strong> {meta.section}</span>}
            {meta.facultyAdvisor && <span className="meta-chip"><strong>Faculty Advisor:</strong> {meta.facultyAdvisor}</span>}
            {meta.academicAdvisor && <span className="meta-chip"><strong>Academic Advisor:</strong> {meta.academicAdvisor}</span>}
          </div>
        </div>
        <div className="sheet-stats">
          <div className="stat-box">
            <div className="stat-box-num">{students.length}</div>
            <div className="stat-box-label">Total</div>
          </div>
          <div className="stat-box">
            <div className={`stat-box-num${lowCount > 0 ? ' danger' : ''}`}>{lowCount}</div>
            <div className="stat-box-label">Below 75%</div>
          </div>
        </div>
      </div>

      {visible.length === 0 ? (
        <div className="empty-state">
          {filterLow
            ? 'No students with attendance below 75% found in the selected subjects.'
            : 'No students match your search.'}
        </div>
      ) : format === 'A' ? (
        <div className="table-wrapper">
          <table className="data-table">
            <thead>
              <tr>
                <th className="th" style={{ textAlign: 'center' }}>#</th>
                <th className="th">Reg No.</th>
                <th className="th">Name</th>
                <th className="th">Subjects & Attendance</th>
                <th className="th" style={{ textAlign: 'center' }}>Below 75%</th>
              </tr>
            </thead>
            <tbody>
              {visible.map((st) => (
                <FormatARow
                  key={st.regNo}
                  student={st}
                  filterLow={filterLow}
                  deselectedSubjects={deselectedSubjects}
                />
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div>
          {visible.map((st) => (
            <FormatBCard
              key={st.regNo + st.sno}
              student={st}
              filterLow={filterLow}
              deselectedSubjects={deselectedSubjects}
            />
          ))}
        </div>
      )}
    </div>
  );
}

// Drop zone component
function DropZone({ onFiles }) {
  const [dragging, setDragging] = useState(false);
  const ref = useRef(null);

  const handleDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const files = Array.from(e.dataTransfer.files).filter((f) => /\.(xlsx|xls)$/i.test(f.name));
    if (files.length) onFiles(files);
  };

  return (
    <div
      className={`dropzone${dragging ? ' dragging' : ''}`}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => ref.current?.click()}
    >
      <input
        ref={ref}
        type="file"
        accept=".xlsx,.xls"
        multiple
        style={{ display: 'none' }}
        onChange={(e) => {
          const files = Array.from(e.target.files || []).filter((f) => /\.(xlsx|xls)$/i.test(f.name));
          if (files.length) onFiles(files);
          e.target.value = '';
        }}
      />
      <div className="dropzone-icon-wrap">
        <Upload size={24} style={{ color: dragging ? 'var(--green)' : 'var(--text-muted)' }} />
      </div>
      <div className="dropzone-title">Drop your Excel attendance sheet here or click to browse</div>
    </div>
  );
}

// Subject filter panel
function SubjectFilterSection({ availableSubjects, deselectedSubjects, onToggleSubject }) {
  const [subSearch, setSubSearch] = useState('');
  const [isExpanded, setIsExpanded] = useState(false);

  const excludedSubjects = useMemo(() => {
    return availableSubjects.filter((s) => deselectedSubjects.has(s.code));
  }, [availableSubjects, deselectedSubjects]);

  const excludedCount = deselectedSubjects.size;
  const totalCount = availableSubjects.length;
  const consideredCount = totalCount - excludedCount;

  const filteredSubjects = useMemo(() => {
    const q = subSearch.trim().toLowerCase();
    if (!q) return availableSubjects;
    return availableSubjects.filter(
      (s) => s.code.toLowerCase().includes(q) || (s.name && s.name.toLowerCase().includes(q))
    );
  }, [availableSubjects, subSearch]);

  const COLLAPSED_LIMIT = 18;
  const showExpandButton = !subSearch && availableSubjects.length > COLLAPSED_LIMIT;
  const displayedSubjects = showExpandButton && !isExpanded
    ? filteredSubjects.slice(0, COLLAPSED_LIMIT)
    : filteredSubjects;

  return (
    <div className="subject-filter">
      <div className="subject-filter-header">
        <div className="subject-filter-title-row">
          <div className="subject-filter-title">
            <SlidersHorizontal size={14} style={{ color: 'var(--green)' }} />
            Consider Subject Attendance
          </div>
          <span className="considered-badge" style={{
            background: excludedCount > 0 ? 'var(--amber-dim)' : 'var(--green-dim)',
            color: excludedCount > 0 ? 'var(--amber)' : 'var(--green)',
            border: `1px solid ${excludedCount > 0 ? 'rgba(245,158,11,0.25)' : 'var(--green-border)'}`,
          }}>
            {consideredCount} of {totalCount} considered in report
          </span>
        </div>
      </div>

      <div className="subject-filter-grid">
        {/* Left Column: Scrollable Deselected Subjects Panel */}
        <div className="deselected-sidebar">
          <div className="deselected-sidebar-header">
            <div className="deselected-sidebar-title">
              <span className="dot-indicator" />
              Deselected Subjects
              <span className="deselected-count-pill">{excludedSubjects.length}</span>
            </div>
            <span className="deselected-sidebar-sub">Excluded from report</span>
          </div>

          <div className="deselected-scroll-list">
            {excludedSubjects.length === 0 ? (
              <div className="deselected-empty-state">
                No subjects deselected.<br />
                <span>Click any subject on the right to exclude it.</span>
              </div>
            ) : (
              excludedSubjects.map((sub) => (
                <div
                  key={sub.code}
                  className={`deselected-card${sub.hasZero ? ' is-zero' : ''}`}
                  onClick={() => onToggleSubject(sub.code)}
                  title={sub.hasZero ? `${sub.code} — 0% attendance (${sub.totalStudents} ${sub.totalStudents === 1 ? 'student' : 'students'} enrolled). Click to re-include` : 'Click to re-include in report'}
                >
                  <div className="deselected-card-main">
                    <div className="deselected-card-top">
                      <span className="deselected-code">{sub.code}</span>
                      {sub.hasZero && <span className="zero-badge">0% attendance</span>}
                    </div>
                    {sub.name && <div className="deselected-name">{sub.name}</div>}
                    <div className="deselected-meta-sub">
                      {sub.totalStudents} {sub.totalStudents === 1 ? 'student enrolled' : 'students enrolled'}
                    </div>
                  </div>
                  <button
                    type="button"
                    className="deselected-action-btn"
                    title="Click to re-include"
                    onClick={(e) => {
                      e.stopPropagation();
                      onToggleSubject(sub.code);
                    }}
                  >
                    <Check size={12} />
                  </button>
                </div>
              ))
            )}
          </div>
        </div>

        {/* Right Column: All Available Subjects */}
        <div className="all-subjects-content">
          <div className="all-subjects-header">
            <span className="all-subjects-label">All Available Subjects (Click to Toggle)</span>
            {totalCount > 8 && (
              <div className="subject-search-box">
                <Search size={11} style={{ color: 'var(--text-muted)' }} />
                <input
                  type="text"
                  placeholder="Find subject..."
                  value={subSearch}
                  onChange={(e) => setSubSearch(e.target.value)}
                />
                {subSearch && (
                  <button
                    className="icon-btn"
                    onClick={(e) => { e.stopPropagation(); setSubSearch(''); }}
                    style={{ color: 'var(--text-muted)' }}
                  >
                    <X size={10} />
                  </button>
                )}
              </div>
            )}
          </div>

          <div className="subject-chips">
            {displayedSubjects.map((sub) => {
              const isSelected = !deselectedSubjects.has(sub.code);
              return (
                <button
                  key={sub.code}
                  className={`subject-chip${isSelected ? ' selected' : ' deselected'}${sub.hasZero ? ' has-zero' : ''}`}
                  onClick={() => onToggleSubject(sub.code)}
                  title={sub.hasZero ? `${sub.code} — 0% attendance (${sub.totalStudents} ${sub.totalStudents === 1 ? 'student' : 'students'} enrolled)` : sub.name ? `${sub.code} — ${sub.name}` : sub.code}
                >
                  <div className={`chip-checkbox${isSelected ? ' checked' : ' unchecked'}`}>
                    {isSelected ? (
                      <Check size={10} color="#fff" strokeWidth={3} />
                    ) : (
                      <X size={9} style={{ color: sub.hasZero ? 'var(--red)' : 'var(--text-muted)' }} />
                    )}
                  </div>
                  <span style={{
                    fontFamily: "'JetBrains Mono', monospace",
                    fontSize: '11px',
                    textDecoration: isSelected ? 'none' : 'line-through',
                    color: sub.hasZero ? 'var(--red)' : (!isSelected ? 'var(--text-muted)' : undefined),
                    fontWeight: sub.hasZero ? 700 : undefined,
                  }}>
                    {sub.code}
                  </span>
                  {sub.hasZero && (
                    <span className="chip-zero-badge">0%</span>
                  )}
                </button>
              );
            })}

            {showExpandButton && (
              <button className="expand-btn" onClick={() => setIsExpanded((prev) => !prev)}>
                {isExpanded ? (
                  <>Show Less <ChevronUp size={12} /></>
                ) : (
                  <>+{availableSubjects.length - COLLAPSED_LIMIT} more <ChevronDown size={12} /></>
                )}
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────────────────

export default function App() {
  const [files, setFiles] = useState([]);
  const [activeIdx, setActiveIdx] = useState(null);
  const [filterLow, setFilterLow] = useState(false);
  const [search, setSearch] = useState('');
  const [errors, setErrors] = useState([]);
  const [deselectedSubjects, setDeselectedSubjects] = useState(() => new Set());
  const [isDark, setIsDark] = useState(true);

  // Apply theme to document
  React.useEffect(() => {
    document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
    if (!isDark) {
      // Light mode overrides
      document.documentElement.style.setProperty('--bg-page', '#F8F9FA');
      document.documentElement.style.setProperty('--bg-card', '#FFFFFF');
      document.documentElement.style.setProperty('--bg-elevated', '#F1F5F9');
      document.documentElement.style.setProperty('--bg-hover', 'rgba(0,0,0,0.04)');
      document.documentElement.style.setProperty('--border', 'rgba(0,0,0,0.08)');
      document.documentElement.style.setProperty('--border-strong', 'rgba(0,0,0,0.18)');
      document.documentElement.style.setProperty('--text-primary', '#0F172A');
      document.documentElement.style.setProperty('--text-secondary', '#475569');
      document.documentElement.style.setProperty('--text-muted', '#94A3B8');
      document.documentElement.style.setProperty('--green', '#16A34A');
      document.documentElement.style.setProperty('--green-dim', 'rgba(22,163,74,0.1)');
      document.documentElement.style.setProperty('--green-border', 'rgba(22,163,74,0.22)');
      document.documentElement.style.setProperty('--bg-sidebar', '#FFFFFF');
      document.documentElement.style.setProperty('--bg-input', '#FFFFFF');
      document.documentElement.style.setProperty('--bg-topnav', 'rgba(255, 255, 255, 0.95)');
    } else {
      // Reset to dark
      document.documentElement.style.setProperty('--bg-page', '#080C10');
      document.documentElement.style.setProperty('--bg-card', '#0E1318');
      document.documentElement.style.setProperty('--bg-elevated', '#131920');
      document.documentElement.style.setProperty('--bg-hover', 'rgba(255,255,255,0.04)');
      document.documentElement.style.setProperty('--border', 'rgba(255,255,255,0.07)');
      document.documentElement.style.setProperty('--border-strong', 'rgba(255,255,255,0.15)');
      document.documentElement.style.setProperty('--text-primary', '#F8FAFC');
      document.documentElement.style.setProperty('--text-secondary', '#94A3B8');
      document.documentElement.style.setProperty('--text-muted', '#64748B');
      document.documentElement.style.setProperty('--green', '#22C55E');
      document.documentElement.style.setProperty('--green-dim', 'rgba(34,197,94,0.12)');
      document.documentElement.style.setProperty('--green-border', 'rgba(34,197,94,0.22)');
      document.documentElement.style.setProperty('--bg-sidebar', '#0C1014');
      document.documentElement.style.setProperty('--bg-input', '#12181F');
      document.documentElement.style.setProperty('--bg-topnav', 'rgba(8, 12, 16, 0.88)');
    }
  }, [isDark]);

  const handleFiles = useCallback(async (incoming) => {
    const added = [];
    const errs = [];
    for (const f of incoming) {
      if (files.some((x) => x.name === f.name)) continue;
      try {
        const buf = await f.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const sheets = parseWorkbook(wb);
        if (sheets.length === 0) {
          errs.push(`"${f.name}" — No recognisable attendance data found.`);
        } else {
          added.push({ name: f.name, size: f.size, sheets });
        }
      } catch (e) {
        errs.push(`"${f.name}" — Failed to read: ${e.message}`);
      }
    }
    if (added.length) {
      setFiles((prev) => {
        const next = [...prev, ...added];
        setActiveIdx(next.length - 1);
        return next;
      });
    }
    if (errs.length) setErrors((prev) => [...prev, ...errs]);
  }, [files]);

  const removeFile = (idx) => {
    setFiles((prev) => {
      const next = prev.filter((_, i) => i !== idx);
      setActiveIdx(next.length > 0 ? Math.min(idx, next.length - 1) : null);
      return next;
    });
  };

  const activeFile = activeIdx !== null ? files[activeIdx] : null;

  const lastActiveFileRef = useRef(null);

  const availableSubjects = useMemo(() => {
    if (!activeFile) return [];
    const map = new Map();
    activeFile.sheets.forEach((sheet) => {
      sheet.students.forEach((st) => {
        const studentId = st.regNo || st.name || String(st.sno);
        (st.subjects || []).forEach((sub) => {
          const norm = normalizeSubjectCode(sub.code, sub.name);
          if (!norm) return;
          if (!map.has(norm)) {
            map.set(norm, {
              code: norm,
              name: sub.name || resolveSubjectName(norm) || '',
              studentIds: new Set(),
              lowStudentIds: new Set(),
              zeroStudentIds: new Set(),
              positiveStudentIds: new Set(),
              maxPct: 0,
            });
          }
          const item = map.get(norm);
          if (!item.name) item.name = sub.name || resolveSubjectName(norm) || '';

          // Track unique student to know if entries are from the same person or different persons
          item.studentIds.add(studentId);

          const p = typeof sub.pct === 'number'
            ? (isNaN(sub.pct) ? null : sub.pct)
            : sub.pct != null && !isNaN(parseFloat(String(sub.pct).replace('%', '').trim()))
              ? parseFloat(String(sub.pct).replace('%', '').trim())
              : null;

          if (p !== null) {
            if (p < LOW) item.lowStudentIds.add(studentId);
            if (p <= 0) {
              item.zeroStudentIds.add(studentId);
            } else {
              item.positiveStudentIds.add(studentId);
              if (p > item.maxPct) item.maxPct = p;
            }
          }
        });
      });
    });

    return Array.from(map.values())
      .map((item) => {
        const totalStudents = item.studentIds.size;
        const zeroStudents = item.zeroStudentIds.size;
        const positiveStudents = item.positiveStudentIds.size;
        const lowStudents = item.lowStudentIds.size;
        const hasZero =
          item.code === '21HCSF003' ||
          (totalStudents > 0 && positiveStudents === 0 && (zeroStudents > 0 || item.maxPct === 0));

        return {
          code: item.code,
          name: item.name,
          totalStudents,
          zeroStudents,
          positiveStudents,
          lowStudents,
          maxPct: item.maxPct,
          hasZero,
        };
      })
      .sort((a, b) => a.code.localeCompare(b.code));
  }, [activeFile]);

  // Automatically deselect 0% attendance subject codes (including electives) initially
  useEffect(() => {
    if (!activeFile) return;
    if (lastActiveFileRef.current !== activeFile) {
      lastActiveFileRef.current = activeFile;
      const initialDeselected = new Set();
      availableSubjects.forEach((sub) => {
        if (sub.hasZero) {
          initialDeselected.add(sub.code);
        }
      });
      setDeselectedSubjects(initialDeselected);
    } else {
      // Ensure any newly detected 0% subjects are deselected
      setDeselectedSubjects((prev) => {
        let changed = false;
        const next = new Set(prev);
        availableSubjects.forEach((sub) => {
          if (sub.hasZero && !next.has(sub.code)) {
            next.add(sub.code);
            changed = true;
          }
        });
        return changed ? next : prev;
      });
    }
  }, [activeFile, availableSubjects]);

  const toggleSubject = (normCode) => {
    setDeselectedSubjects((prev) => {
      const next = new Set(prev);
      next.has(normCode) ? next.delete(normCode) : next.add(normCode);
      return next;
    });
  };

  const totalStudents = activeFile ? activeFile.sheets.reduce((a, s) => a + s.students.length, 0) : 0;
  const totalLow = activeFile
    ? activeFile.sheets.reduce((a, s) => a + s.students.filter((st) => hasLowAttendance(st, deselectedSubjects)).length, 0)
    : 0;

  const [showExportModal, setShowExportModal] = useState(false);
  const [exportSection, setExportSection] = useState('');
  const [exportFacultyAdvisor, setExportFacultyAdvisor] = useState('');
  const [exportAcademicAdvisor, setExportAcademicAdvisor] = useState('');
  const [exportTotalStudents, setExportTotalStudents] = useState('');
  const [exportMonth, setExportMonth] = useState('October 2026');
  const [exportSuccessMsg, setExportSuccessMsg] = useState('');

  // Extract meta defaults whenever activeFile changes
  useEffect(() => {
    if (!activeFile) return;
    const meta = activeFile.sheets[0]?.meta || {};
    setExportSection(''); // Do not autofill section with excel sheet or file name
    if (meta.facultyAdvisor) setExportFacultyAdvisor(meta.facultyAdvisor);
    if (meta.academicAdvisor) setExportAcademicAdvisor(meta.academicAdvisor);
    setExportTotalStudents(''); // Always ask user directly, do not guess
    if (activeFile.sheets[0]?.sheetName && activeFile.sheets[0].sheetName !== 'Sheet1') {
      setExportMonth(activeFile.sheets[0].sheetName);
    }
  }, [activeFile]);

  const downloadExcelReport = useCallback(async () => {
    if (!activeFile) return;

    // Collect & deduplicate students by regNo (keep entry with most subjects)
    const byReg = new Map();
    activeFile.sheets.flatMap((s) => s.students).forEach((st) => {
      const existing = byReg.get(st.regNo);
      if (!existing || (st.subjects || []).length > (existing.subjects || []).length) {
        byReg.set(st.regNo, st);
      }
    });

    const lowStudents = Array.from(byReg.values())
      .filter((st) => hasLowAttendance(st, deselectedSubjects))
      .sort((a, b) => a.regNo.localeCompare(b.regNo));

    const finalSection = (exportSection || '').trim();
    const finalAdvisor = (exportFacultyAdvisor || '').trim();
    const finalAcademic = (exportAcademicAdvisor || '').trim();
    const finalTotalStudents = (exportTotalStudents || '').trim();

    // Decode base64 filled-template
    const binaryString = window.atob(FILLED_TEMPLATE_BASE64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes);

    // Pick July 2026 sheet which has the 8 columns (including Mail sent to Parents & Parents Response)
    const templateWs = wb.getWorksheet('July 2026') || wb.worksheets[0];

    // Remove any other worksheets (e.g. August 2026)
    const wsList = wb.worksheets.map((w) => ({ name: w.name, id: w.id }));
    for (const item of wsList) {
      if (item.name !== templateWs.name) {
        wb.removeWorksheet(item.id);
      }
    }

    const ws = templateWs;
    if (exportMonth && exportMonth.trim()) {
      ws.name = exportMonth.trim().slice(0, 31);
    } else {
      ws.name = 'Low Attendance Report';
    }

    // 1. Fill Row 5: Section & Advisors (do not fill what is unknown)
    ws.getCell('A5').value = finalSection ? `SECTION: ${finalSection}` : 'SECTION:';
    ws.getCell('C5').value = finalAcademic ? `Academic Advisor: ${finalAcademic}` : '';
    ws.getCell('E5').value = finalAdvisor ? `FACULTY ADVISOR: ${finalAdvisor}` : 'FACULTY ADVISOR:';

    // 2. Fill Row 6: Total Students (only fill what user specified, never guess)
    ws.getCell('A6').value = 'TOTAL NUMBER OF';
    ws.getCell('C6').value = finalTotalStudents ? ` STUDENTS: ${finalTotalStudents}` : '';
    ws.getCell('E6').value = 'Number of Students Less than 75% (Even in 1 subject): ';

    // 3. Unmerge all old student rows >= 8
    const mergesToUnmerge = [];
    if (ws.model && ws.model.merges) {
      for (const m of ws.model.merges) {
        const match = m.match(/^[A-Z]+(\d+):[A-Z]+(\d+)$/);
        if (match && parseInt(match[1], 10) >= 8) {
          mergesToUnmerge.push(m);
        }
      }
    }
    mergesToUnmerge.forEach((m) => {
      try {
        ws.unMergeCells(m);
      } catch (e) { }
    });

    // 4. Clear rows 8 to 250 and remove any background fills (e.g. blue cells from template)
    for (let r = 8; r <= 250; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 8; c++) {
        const cell = row.getCell(c);
        cell.value = null;
        delete cell.fill;
        cell.fill = { type: 'pattern', pattern: 'none' };
      }
    }

    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
    };
    const fontTimes12 = { name: 'Times New Roman', size: 12, color: { argb: 'FF000000' } };

    // 5. Fill each student's low-attendance subjects starting from Row 8
    let currentRow = 8;
    let sno = 1;

    for (const st of lowStudents) {
      const lowSubs = getLowSubjects(st, deselectedSubjects);
      if (!lowSubs || lowSubs.length === 0) continue;

      const k = lowSubs.length;
      const rStart = currentRow;
      const rEnd = currentRow + k - 1;

      for (let i = 0; i < k; i++) {
        const r = rStart + i;
        const sub = lowSubs[i];
        const row = ws.getRow(r);

        if (i === 0) {
          row.getCell(1).value = sno;
          row.getCell(2).value = st.regNo || '';
          row.getCell(3).value = st.name || '';
          row.getCell(7).value = 'YES';
          row.getCell(8).value = '';
        }

        // Col 4: Subject Code
        row.getCell(4).value = sub.code || '';

        // Col 5: Subject Name - strictly resolve from dictionary or sub.name, leave empty if unknown
        const normCode = normalizeSubjectCode(sub.code, sub.name);
        const resolvedName = sub.name || resolveSubjectName(sub.code) || resolveSubjectName(normCode) || '';
        row.getCell(5).value = resolvedName;

        // Col 6: Attendance Percentage
        row.getCell(6).value =
          sub.pct !== null && sub.pct !== undefined ? Number(sub.pct.toFixed(2)) : '';

        // Apply borders, font, clean fill, and alignment to all 8 cells in this row
        for (let c = 1; c <= 8; c++) {
          const cell = row.getCell(c);
          cell.border = thinBorder;
          cell.font = fontTimes12;
          delete cell.fill;
          cell.fill = { type: 'pattern', pattern: 'none' };
          cell.alignment = {
            vertical: 'middle',
            horizontal: c === 3 || c === 5 ? 'left' : (c === 4 || c === 6 ? 'left' : 'center'),
            wrapText: true,
          };
        }
      }

      // Merge columns A, B, C, G, H across the student's rows if k > 1 (subdividing the cell for courses)
      if (k > 1) {
        ws.mergeCells(rStart, 1, rEnd, 1);
        ws.mergeCells(rStart, 2, rEnd, 2);
        ws.mergeCells(rStart, 3, rEnd, 3);
        ws.mergeCells(rStart, 7, rEnd, 7);
        ws.mergeCells(rStart, 8, rEnd, 8);
      }

      currentRow = rEnd + 1;
      sno++;
    }

    // Generate buffer & Blob
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const fileName = `${finalSection || 'Section'}_Low_Attendance_Report.xlsx`;

    try {
      saveAs(blob, fileName);
    } catch (e) {
      console.warn('saveAs error, using fallback download:', e);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 40000);
    }

    setShowExportModal(false);
    setExportSuccessMsg(`Exported ${fileName} successfully!`);
    setTimeout(() => setExportSuccessMsg(''), 4500);
  }, [
    activeFile,
    deselectedSubjects,
    totalStudents,
    exportSection,
    exportFacultyAdvisor,
    exportAcademicAdvisor,
    exportTotalStudents,
    exportMonth,
  ]);

  return (
    <div className="app-layout">
      {/* Ambient background */}
      <div className="ambient-bg" />

      {/* Top Nav */}
      <nav className="topnav">
        <div className="topnav-logo">
          <div className="topnav-logo-icon">
            <Users size={16} style={{ color: 'var(--green)' }} />
          </div>
          <span className="topnav-logo-text">
            Attend<span>ance</span>
          </span>
        </div>
        <span className="topnav-badge">SRM Faculty Portal</span>
        <div className="topnav-spacer" />
        <button
          className="topnav-theme-btn"
          onClick={() => setIsDark((d) => !d)}
          title={isDark ? 'Switch to light mode' : 'Switch to dark mode'}
        >
          {isDark ? <Sun size={15} /> : <Moon size={15} />}
        </button>
      </nav>

      {/* Main content */}
      <main className="main-content">
        {/* Page header */}
        <div className="page-header">
          <div className="page-header-top">
            <div className="page-header-icon">
              <FileSpreadsheet size={20} style={{ color: 'var(--green)' }} />
            </div>
            <h1 className="page-title">Attendance Records</h1>
          </div>
        </div>

        {/* Error banner */}
        {errors.length > 0 && (
          <div className="error-banner">
            <AlertTriangle size={15} style={{ color: 'var(--red)', flexShrink: 0, marginTop: '2px' }} />
            <div className="error-banner-text">
              {errors.map((e, i) => <div key={i}>{e}</div>)}
            </div>
            <button className="icon-btn" onClick={() => setErrors([])} style={{ color: 'var(--red)' }}>
              <X size={14} />
            </button>
          </div>
        )}

        {/* Drop zone */}
        <DropZone onFiles={handleFiles} />

        {/* File tabs */}
        {files.length > 0 && (
          <div className="file-tabs">
            {files.map((f, idx) => (
              <button
                key={f.name}
                className={`file-tab${activeIdx === idx ? ' active' : ''}`}
                onClick={() => {
                  setActiveIdx(idx);
                  setDeselectedSubjects(new Set());
                }}
              >
                <FileSpreadsheet
                  size={14}
                  style={{ color: activeIdx === idx ? 'var(--green)' : 'var(--text-muted)', flexShrink: 0 }}
                />
                <span className="file-tab-name">{f.name}</span>
                <span className="file-tab-size">({(f.size / 1024).toFixed(0)}KB)</span>
                <span
                  className="file-tab-remove"
                  onClick={(e) => { e.stopPropagation(); removeFile(idx); }}
                  title="Remove"
                >
                  <X size={13} />
                </span>
              </button>
            ))}
          </div>
        )}

        {/* Controls + data */}
        {activeFile && (
          <div style={{ marginTop: '24px' }}>
            {/* Control panel */}
            <div className="control-panel">
              <div className="control-row">
                {/* Search */}
                <div className="search-box">
                  <Search size={13} style={{ color: 'var(--text-muted)', flexShrink: 0 }} />
                  <input
                    type="text"
                    placeholder="Search name or reg no..."
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                  />
                  {search && (
                    <button className="icon-btn" onClick={() => setSearch('')} style={{ color: 'var(--text-muted)' }}>
                      <X size={12} />
                    </button>
                  )}
                </div>

                {/* Low attendance filter */}
                <button
                  className={`filter-btn${filterLow ? ' active' : ''}`}
                  onClick={() => setFilterLow((v) => !v)}
                >
                  <AlertTriangle size={14} />
                  {filterLow ? 'Showing: below 75% only' : 'Filter: below 75%'}
                  {filterLow && totalLow > 0 && (
                    <span style={{
                      padding: '1px 7px', borderRadius: '999px', fontSize: '11px', fontWeight: 700,
                      background: 'rgba(239,68,68,0.2)', color: 'var(--red)',
                    }}>{totalLow}</span>
                  )}
                </button>

                {/* Export to Excel */}
                <button
                  className="filter-btn"
                  onClick={() => setShowExportModal(true)}
                  title="Export students with low attendance to Excel (template format)"
                  style={{ borderColor: 'var(--green-border)', color: 'var(--green)' }}
                >
                  <Download size={14} />
                  Export to Excel
                  {totalLow > 0 && (
                    <span style={{
                      padding: '1px 7px', borderRadius: '999px', fontSize: '11px', fontWeight: 700,
                      background: 'var(--green-dim)', color: 'var(--green)',
                    }}>{totalLow}</span>
                  )}
                </button>

                {/* Stats */}
                <div className="stats-row">
                  <div className="stat-box">
                    <div className="stat-box-num">{totalStudents}</div>
                    <div className="stat-box-label">Students</div>
                  </div>
                  <div className="stat-box">
                    <div className={`stat-box-num${totalLow > 0 ? ' danger' : ''}`}>{totalLow}</div>
                    <div className="stat-box-label">Below 75%</div>
                  </div>
                </div>
              </div>

              {/* Subject filter */}
              {availableSubjects.length > 0 && (
                <SubjectFilterSection
                  availableSubjects={availableSubjects}
                  deselectedSubjects={deselectedSubjects}
                  onToggleSubject={toggleSubject}
                  onSelectAll={() => setDeselectedSubjects(new Set())}
                  onDeselectAll={() => setDeselectedSubjects(new Set(availableSubjects.map((s) => s.code)))}
                />
              )}
            </div>

            {/* Sheets */}
            {activeFile.sheets.map((sheet, idx) => (
              <SheetDisplay
                key={idx}
                sheet={sheet}
                filterLow={filterLow}
                searchQuery={search}
                deselectedSubjects={deselectedSubjects}
              />
            ))}
          </div>
        )}

        {files.length === 0 && (
          <div className="upload-hint">
            <FileSpreadsheet size={40} style={{ margin: '0 auto 12px', opacity: 0.18 }} />
            <p>Upload an Excel attendance sheet above to get started.</p>
          </div>
        )}
      </main>

      {/* Toast Notification */}
      {exportSuccessMsg && (
        <div className="export-toast">
          <CheckCircle2 size={16} style={{ color: 'var(--green)', flexShrink: 0 }} />
          <span>{exportSuccessMsg}</span>
        </div>
      )}

      {/* Export to Excel Modal */}
      {showExportModal && (
        <div className="modal-backdrop" onClick={() => setShowExportModal(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title-wrap">
                <div className="modal-icon">
                  <FileSpreadsheet size={18} style={{ color: 'var(--green)' }} />
                </div>
                <div>
                  <h3 className="modal-title">Export Low Attendance Report</h3>
                  <p className="modal-subtitle">Export filtered low attendance report</p>
                </div>
              </div>
              <button className="icon-btn" onClick={() => setShowExportModal(false)}>
                <X size={16} />
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                downloadExcelReport();
              }}
            >
              <div className="modal-body">
                <div className="modal-fields-grid">
                  <div className="form-group">
                    <label className="form-label">
                      Section <span style={{ color: 'var(--red)' }}>*</span>
                    </label>
                    <input
                      type="text"
                      className="form-input"
                      value={exportSection}
                      onChange={(e) => setExportSection(e.target.value)}
                      required
                      autoFocus
                    />
                  </div>

                  <div className="form-group">
                    <label className="form-label">Sheet Tab / Month</label>
                    <input
                      type="text"
                      className="form-input"
                      value={exportMonth}
                      onChange={(e) => setExportMonth(e.target.value)}
                    />
                  </div>
                </div>

                <div className="modal-fields-grid">
                  <div className="form-group">
                    <label className="form-label">Total Number of Students</label>
                    <input
                      type="text"
                      inputMode="numeric"
                      className="form-input"
                      value={exportTotalStudents}
                      onChange={(e) => setExportTotalStudents(e.target.value)}
                    />
                  </div>

                  <div className="form-group">
                    <label className="form-label">Faculty Advisor</label>
                    <input
                      type="text"
                      className="form-input"
                      value={exportFacultyAdvisor}
                      onChange={(e) => setExportFacultyAdvisor(e.target.value)}
                    />
                  </div>
                </div>

                <div className="form-group">
                  <label className="form-label">Academic Advisor</label>
                  <input
                    type="text"
                    className="form-input"
                    value={exportAcademicAdvisor}
                    onChange={(e) => setExportAcademicAdvisor(e.target.value)}
                  />
                </div>

                {/* Summary Box */}
                <div className="modal-summary">
                  <div className="summary-item">
                    <span className="summary-label">Below 75% Students</span>
                    <span className="summary-val danger">{totalLow}</span>
                  </div>
                  <div className="summary-item">
                    <span className="summary-label">Total Students</span>
                    <span className="summary-val">{exportTotalStudents.trim() || '—'}</span>
                  </div>
                  <div className="summary-item">
                    <span className="summary-label">Excluded Subjects</span>
                    <span className="summary-val">{deselectedSubjects.size}</span>
                  </div>
                </div>

                <p className="modal-note">
                  • Sorted register number wise with one row per low attendance subject.<br />
                  • &ldquo;Mail sent to Parents&rdquo; is set to <strong>Yes</strong>, and &ldquo;Parents Response&rdquo; is left blank.
                </p>
              </div>

              <div className="modal-footer">
                <button
                  type="button"
                  className="modal-cancel-btn"
                  onClick={() => setShowExportModal(false)}
                >
                  Cancel
                </button>
                <button type="submit" className="modal-export-btn">
                  <Download size={15} />
                  Download Excel (.xlsx)
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
