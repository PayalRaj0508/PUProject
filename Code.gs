/*************************************************************
 *  STUDENT DASHBOARD — Code.gs
 *  Created by Parekh Payal
 *  Spreadsheet: 1chDeFYe7OnNWiAe-1F6h9I9jMfMwDyyW2qx7WV6zC7s
 *************************************************************/

/* ============================================================
 *  CONFIG
 * ============================================================ */
const SPREADSHEET_ID = '1chDeFYe7OnNWiAe-1F6h9I9jMfMwDyyW2qx7WV6zC7s';

const SHEET_UDP      = 'Faculty Remark_UDP Project';
const SHEET_COUNT    = 'Project_Count';
const SHEET_NOT_SUB  = 'Faculty Remark_not submitted st';
const SHEET_CONSULT  = 'Faculty Remark for Consultancy';
const SHEET_IDP      = 'Faculty Remark_IDP Project';

const REVIEW_MAX = { r1: 5, r2: 10, r3: 15, r4: 20, r5: 20, r6: 30 };

const REVIEW_LABELS = {
  r1: 'Review 1 · Definition',
  r2: 'Review 2 · Feasibility',
  r3: 'Review 3 · SRS & Diagrams',
  r4: 'Review 4 · Dev Phase 1',
  r5: 'Review 5 · Dev Phase 2/3',
  r6: 'Review 6 · Completion'
};

const CACHE_SECONDS = 120;
const CACHE_KEY     = 'student_dashboard_v9';

/* ============================================================
 *  WEB APP ENTRY POINT
 * ============================================================ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student Dashboard — by Parekh Payal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================================================
 *  MAIN DATA API (called from the HTML via google.script.run)
 * ============================================================ */
function getDashboardData(force) {
  try {
    const cache = CacheService.getScriptCache();
    if (!force) {
      const hit = cache.get(CACHE_KEY);
      if (hit) {
        const o = JSON.parse(hit);
        o.cached = true;
        return o;
      }
    }
    const p = buildPayload();
    try { cache.put(CACHE_KEY, JSON.stringify(p), CACHE_SECONDS); } catch (e) {}
    return p;
  } catch (err) {
    return { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
}

/* ============================================================
 *  HELPERS
 * ============================================================ */
function findSheet_(ss, name) {
  const clean = s => String(s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  const w = clean(name);
  const all = ss.getSheets();

  // exact match first
  for (let i = 0; i < all.length; i++) {
    if (clean(all[i].getName()) === w) return all[i];
  }
  // partial match fallback
  for (let i = 0; i < all.length; i++) {
    if (clean(all[i].getName()).indexOf(w) !== -1) return all[i];
  }
  return null;
}

function mentorKey(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^(prof\.?|dr\.?|mr\.?|mrs\.?|ms\.?)\s*/i, '')
    .replace(/\./g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalize(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/\.0+$/, '').replace(/\s+/g, '').trim();
}

function parseMark(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (!s || s === '#N/A' || s === 'AB' || s === '-' || s === '—') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/* ============================================================
 *  BUILD THE FULL PAYLOAD
 * ============================================================ */
function buildPayload() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  /* ----------------------------------------------------------
   *  1. Resolve sheets + build sheetStatus
   * -------------------------------------------------------- */
  const sheetStatus = {}, resolved = {};
  const wanted = {
    udp:     SHEET_UDP,
    count:   SHEET_COUNT,
    notSub:  SHEET_NOT_SUB,
    consult: SHEET_CONSULT,
    idp:     SHEET_IDP
  };

  Object.keys(wanted).forEach(function (k) {
    const s = findSheet_(ss, wanted[k]);
    resolved[k] = s;
    sheetStatus[k] = {
      wanted: wanted[k],
      found:  !!s,
      actual: s ? s.getName() : '❌ NOT FOUND',
      rows:   s ? s.getLastRow() : 0
    };
  });

  /* ----------------------------------------------------------
   *  2. Preload IDP lookup (used to classify UDP vs IDP)
   * -------------------------------------------------------- */
  const idpSet = new Set(), idpLookup = {};
  if (resolved.idp) {
    const d = resolved.idp.getDataRange().getDisplayValues();
    for (let i = 1; i < d.length; i++) {
      const e = normalize(d[i][2]);
      if (!e || !/^\d{8,}$/.test(e)) continue;
      idpSet.add(e);
      idpLookup[e] = {
        company:      String(d[i][6]  || '').trim(),
        companyAddr:  String(d[i][7]  || '').trim(),
        companyPhone: String(d[i][8]  || '').trim(),
        companyEmail: String(d[i][9]  || '').trim(),
        jobRole:      String(d[i][10] || '').trim()
      };
    }
  }

  /* ----------------------------------------------------------
   *  3. Main UDP sheet → student map
   * -------------------------------------------------------- */
  const map = new Map();
  if (resolved.udp) {
    const data = resolved.udp.getDataRange().getDisplayValues();
    let curGroup = null, groupSerial = 0, orphan = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawEnr = String(row[1] || '').trim();
      const name   = String(row[2] || '').trim();
      const enr    = /^#VALUE!?$/i.test(rawEnr) ? '' : normalize(rawEnr);

      // group serial detection
      const sRaw = String(row[0] || '').trim();
      const sMatch = sRaw.match(/^(\d+)(\.0*)?$/);
      if (sMatch) {
        groupSerial = parseInt(sMatch[1], 10);
        curGroup = 'G' + groupSerial;
      } else if (!curGroup && (enr || name)) {
        orphan++;
        curGroup = 'X' + orphan;
      }

      if (!enr && !name) continue;
      if (/^(Total|HOLD|Semester\s*Exchange|NOT\s*SUBMITTED|Sr\.?\s*No)/i.test(name)) continue;
      if (/^(Total|HOLD|Semester\s*Exchange|NOT\s*SUBMITTED)/i.test(rawEnr)) continue;
      if (/^#VALUE!?$/i.test(name)) continue;

      const key  = enr || ('NAME:' + name);
      const jVal = String(row[9] || '').trim();

      let ptype;
      if (jVal.toUpperCase() === 'UDP')             ptype = 'UDP';
      else if (jVal.toUpperCase() === 'INTERNSHIP') ptype = 'IDP';
      else if (jVal.startsWith('=') || !jVal)       ptype = idpSet.has(enr) ? 'IDP' : 'UDP';
      else                                          ptype = 'UDP';

      const topic  = String(row[10] || '').trim();
      const mentor = String(row[11] || '').trim();

      const st = {
        enrollment:  enr,
        name:        name,
        contact:     String(row[3]  || '').trim(),
        email:       String(row[4]  || '').trim(),
        specialization: String(row[5] || '').trim(),
        course:      String(row[6]  || '').trim(),
        division:    String(row[7]  || '').trim(),
        technology:  String(row[8]  || '').trim(),
        projectType: ptype,
        topic:       topic,
        mentor:      mentor,
        groupId:     curGroup,
        groupSerial: groupSerial || null,
        reviews: {
          r1: parseMark(row[12]),
          r2: parseMark(row[13]),
          r3: parseMark(row[14]),
          r4: parseMark(row[15]),
          r5: parseMark(row[16]),
          r6: parseMark(row[17])
        },
        flags: {
          udp:          ptype === 'UDP',
          idp:          ptype === 'IDP',
          consultancy:  false,
          hold:         false,
          exchange:     false,
          notSubmitted: false
        }
      };

      if (map.has(key)) {
        const ex = map.get(key);
        if (!ex.topic  && topic)  ex.topic  = topic;
        if (!ex.mentor && mentor) ex.mentor = mentor;
        Object.assign(ex, st);
      } else {
        map.set(key, st);
      }
    }
  }

  /* ----------------------------------------------------------
   *  4. Project_Count sheet → hold / exchange / notSubmitted
   * -------------------------------------------------------- */
  const sections = parseProjectCountSheet(resolved.count);
  sections.hold.forEach(s => applyFlag(map, s, 'hold'));
  sections.exchange.forEach(s => applyFlag(map, s, 'exchange'));
  sections.notSubmitted.forEach(s => applyFlag(map, s, 'notSubmitted'));

  /* ----------------------------------------------------------
   *  5. Consultancy sheet
   * -------------------------------------------------------- */
  if (resolved.consult) {
    const data = resolved.consult.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawEnr = String(row[1] || '').trim();
      const name   = String(row[2] || '').trim();
      if (!rawEnr && !name) continue;
      if (/^Sr\.?\s*No/i.test(rawEnr) || /^StudentName/i.test(name)) continue;

      const enr = /^#VALUE!?$/i.test(rawEnr) ? '' : normalize(rawEnr);
      if (!enr && !name) continue;

      const d = {
        enrollment: enr,
        name:       name,
        program:    String(row[3] || '').trim(),
        topic:      String(row[4] || '').trim(),
        mentor:     String(row[5] || '').trim(),
        remark:     String(row[6] || '').trim(),
        callRemark: String(row[7] || '').trim(),
        misRemark:  String(row[8] || '').trim()
      };

      const key = enr || ('NAME:' + name);
      if (map.has(key)) {
        const ex = map.get(key);
        ex.flags.consultancy  = true;
        ex.consultancyRemark  = d.remark || d.callRemark || d.misRemark;
        ex.callRemark         = d.callRemark;
        ex.misRemark          = d.misRemark;
        if (!ex.mentor) ex.mentor = d.mentor;
      } else {
        map.set(key, {
          enrollment: enr,
          name:       name,
          contact: '', email: '',
          specialization: '', course: d.program, division: '', technology: '',
          projectType: 'CONSULTANCY',
          topic:  d.topic,
          mentor: d.mentor,
          consultancyRemark: d.remark || d.callRemark || d.misRemark,
          callRemark: d.callRemark,
          misRemark:  d.misRemark,
          reviews: {},
          flags: {
            udp: false, idp: false, consultancy: true,
            hold: false, exchange: false, notSubmitted: false
          }
        });
      }
    }
  }

  /* ----------------------------------------------------------
   *  6. Not‑Submitted sheet
   * -------------------------------------------------------- */
  if (resolved.notSub) {
    const data = resolved.notSub.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawEnr = String(row[1] || '').trim();
      const name   = String(row[2] || '').trim();
      if (!rawEnr && !name) continue;
      if (/^Sr\.?\s*No/i.test(rawEnr)) continue;

      const enr = /^#VALUE!?$/i.test(rawEnr) ? '' : normalize(rawEnr);
      if (!enr && !name) continue;

      const key = enr || ('NAME:' + name);
      if (map.has(key)) {
        map.get(key).flags.notSubmitted = true;
        const rmk = String(row[6] || '').trim();
        if (rmk) map.get(key).notSubRemark = rmk;
      } else {
        map.set(key, {
          enrollment: enr,
          name:       name,
          contact: '', email: '',
          specialization: '', course: String(row[3] || '').trim(),
          division: '', technology: '',
          projectType: 'NOT SUBMITTED',
          topic:  String(row[4] || '').trim(),
          mentor: String(row[5] || '').trim(),
          notSubRemark: String(row[6] || '').trim(),
          reviews: {},
          flags: {
            udp: false, idp: false, consultancy: false,
            hold: false, exchange: false, notSubmitted: true
          }
        });
      }
    }
  }

  /* ----------------------------------------------------------
   *  7. Clean students list
   * -------------------------------------------------------- */
  const students = Array.from(map.values()).filter(function (s) {
    if (!s.name || s.name.length < 2) return false;
    if (/^#VALUE!?$/i.test(s.name)) return false;
    if (/^(Total|HOLD|Semester|NOT SUBMITTED|Sr\.?\s*No|StudentName)/i.test(s.name)) return false;
    if (s.name === s.enrollment) return false;
    return true;
  });

  students.forEach(function (s) {
    if (s.flags.idp && idpLookup[s.enrollment]) {
      const i = idpLookup[s.enrollment];
      s.company      = i.company;
      s.companyAddr  = i.companyAddr;
      s.companyPhone = i.companyPhone;
      s.companyEmail = i.companyEmail;
      s.jobRole      = i.jobRole;
    }
  });

  /* ----------------------------------------------------------
   *  8. Counts
   * -------------------------------------------------------- */
  const counts = {
    total: students.length,
    udp: 0, idp: 0, consultancy: 0,
    hold: 0, exchange: 0, notSubmitted: 0,
    totalGroups: 0
  };

  students.forEach(function (s) {
    if (s.flags.udp)          counts.udp++;
    if (s.flags.idp)          counts.idp++;
    if (s.flags.consultancy)  counts.consultancy++;
    if (s.flags.hold)         counts.hold++;
    if (s.flags.exchange)     counts.exchange++;
    if (s.flags.notSubmitted) counts.notSubmitted++;
  });

  /* ----------------------------------------------------------
   *  9. Groups
   * -------------------------------------------------------- */
  const groupInfo = {};
  students.forEach(function (s) {
    if (!s.groupId) return;
    if (!groupInfo[s.groupId]) groupInfo[s.groupId] = { topic: '', mentor: '' };
    if (s.topic  && !groupInfo[s.groupId].topic)  groupInfo[s.groupId].topic  = s.topic;
    if (s.mentor && !groupInfo[s.groupId].mentor) groupInfo[s.groupId].mentor = s.mentor;
  });

  const groupMap = new Map();
  students.forEach(function (s) {
    if (!s.groupId) return;
    const info   = groupInfo[s.groupId] || {};
    const topic  = info.topic  || s.topic  || 'Untitled Project';
    const mentor = info.mentor || s.mentor || '';

    if (!groupMap.has(s.groupId)) {
      groupMap.set(s.groupId, {
        id: s.groupId,
        serial: s.groupSerial,
        topic: topic,
        mentor: mentor,
        members: [],
        categories: new Set(),
        courses: new Set(),
        divisions: new Set(),
        reviewSums:   { r1: 0, r2: 0, r3: 0, r4: 0, r5: 0, r6: 0 },
        reviewCounts: { r1: 0, r2: 0, r3: 0, r4: 0, r5: 0, r6: 0 },
        avgTotal: null
      });
    }
    const g = groupMap.get(s.groupId);

    if (!g.topic || g.topic === 'Untitled Project') g.topic = topic;
    if (!g.mentor) g.mentor = mentor;

    g.members.push({
      enrollment: s.enrollment,
      name:       s.name,
      course:     s.course,
      division:   s.division,
      technology: s.technology,
      email:      s.email,
      contact:    s.contact,
      reviews:    s.reviews,
      flags:      s.flags
    });

    ['udp', 'idp', 'consultancy', 'hold', 'exchange', 'notSubmitted'].forEach(function (k) {
      if (s.flags[k]) g.categories.add(k);
    });

    if (s.course)   g.courses.add(s.course);
    if (s.division) g.divisions.add(s.division);

    ['r1', 'r2', 'r3', 'r4', 'r5', 'r6'].forEach(function (k) {
      const v = s.reviews ? s.reviews[k] : null;
      if (v !== null && v !== undefined) {
        g.reviewSums[k] += v;
        g.reviewCounts[k]++;
      }
    });
  });

  const groups = Array.from(groupMap.values()).map(function (g) {
    const avgR = {};
    let sumT = 0, cntS = 0;

    ['r1', 'r2', 'r3', 'r4', 'r5', 'r6'].forEach(function (k) {
      avgR[k] = g.reviewCounts[k]
        ? +(g.reviewSums[k] / g.reviewCounts[k]).toFixed(2)
        : null;
    });

    g.members.forEach(function (m) {
      const r = m.reviews || {};
      const vals = ['r1', 'r2', 'r3', 'r4', 'r5', 'r6'].map(k => r[k]);
      if (vals.some(v => v !== null && v !== undefined)) {
        sumT += vals.reduce((a, v) => a + (v || 0), 0);
        cntS++;
      }
    });

    return {
      id: g.id,
      serial: g.serial,
      topic: g.topic,
      mentor: g.mentor,
      memberCount: g.members.length,
      members: g.members,
      categories: Array.from(g.categories),
      courses:    Array.from(g.courses),
      divisions:  Array.from(g.divisions),
      avgR1: avgR.r1, avgR2: avgR.r2, avgR3: avgR.r3,
      avgR4: avgR.r4, avgR5: avgR.r5, avgR6: avgR.r6,
      avgTotal: cntS ? +(sumT / cntS).toFixed(2) : null
    };
  });

  groups.sort(function (a, b) {
    if (a.serial && b.serial) return a.serial - b.serial;
    if (a.serial) return -1;
    if (b.serial) return 1;
    return a.topic.localeCompare(b.topic);
  });
  counts.totalGroups = groups.length;

  /* ----------------------------------------------------------
   * 10. Faculty
   * -------------------------------------------------------- */
  const facMap = new Map();
  students.forEach(function (s) {
    if (!s.mentor || /^#N\/A$/i.test(s.mentor) || s.mentor === '#N/A') return;
    const k = mentorKey(s.mentor);
    if (!k) return;

    if (!facMap.has(k)) {
      facMap.set(k, {
        key: k,
        names: {},
        students: [],
        groups: new Set(),
        udp: 0, idp: 0, consultancy: 0,
        reviewSums:   { r1: 0, r2: 0, r3: 0, r4: 0, r5: 0, r6: 0 },
        reviewCounts: { r1: 0, r2: 0, r3: 0, r4: 0, r5: 0, r6: 0 }
      });
    }
    const f = facMap.get(k);
    f.students.push(s);
    f.names[s.mentor] = (f.names[s.mentor] || 0) + 1;
    if (s.groupId) f.groups.add(s.groupId);
    if (s.flags.udp)         f.udp++;
    if (s.flags.idp)         f.idp++;
    if (s.flags.consultancy) f.consultancy++;

    ['r1', 'r2', 'r3', 'r4', 'r5', 'r6'].forEach(function (kk) {
      const v = s.reviews ? s.reviews[kk] : null;
      if (v !== null && v !== undefined) {
        f.reviewSums[kk] += v;
        f.reviewCounts[kk]++;
      }
    });
  });

  const faculty = [];
  facMap.forEach(function (f) {
    const display = Object.keys(f.names).sort((a, b) => f.names[b] - f.names[a])[0];

    const sums = { r1: 0, r2: 0, r3: 0, r4: 0, r5: 0, r6: 0, total: 0 };
    const cnts = { r1: 0, r2: 0, r3: 0, r4: 0, r5: 0, r6: 0, total: 0 };
    let completed = 0, withAny = 0, zeroReviews = 0;

    f.students.forEach(function (s) {
      const r = s.reviews || {};
      const vals = [r.r1, r.r2, r.r3, r.r4, r.r5, r.r6];
      const filled = vals.filter(v => v !== null && v !== undefined).length;

      if (filled === 0) { zeroReviews++; return; }
      withAny++;
      if (filled === 6) completed++;

      let st = 0;
      ['r1', 'r2', 'r3', 'r4', 'r5', 'r6'].forEach(function (kk) {
        if (r[kk] !== null && r[kk] !== undefined) {
          sums[kk] += r[kk];
          cnts[kk]++;
          st += r[kk];
        }
      });
      sums.total += st;
      cnts.total++;
    });

    const groupList = groups
      .filter(g => mentorKey(g.mentor) === f.key)
      .map(g => ({
        id: g.id,
        serial: g.serial,
        topic: g.topic,
        memberCount: g.memberCount,
        avgTotal: g.avgTotal
      }));

    faculty.push({
      name: display,
      key: f.key,
      totalStudents: f.students.length,
      totalGroups: f.groups.size,
      udp: f.udp,
      idp: f.idp,
      consultancy: f.consultancy,
      withData: withAny,
      zeroReviews: zeroReviews,
      completed: completed,
      completionPct: withAny ? Math.round((completed / withAny) * 100) : 0,
      avgR1: cnts.r1 ? +(sums.r1 / cnts.r1).toFixed(2) : null,
      avgR2: cnts.r2 ? +(sums.r2 / cnts.r2).toFixed(2) : null,
      avgR3: cnts.r3 ? +(sums.r3 / cnts.r3).toFixed(2) : null,
      avgR4: cnts.r4 ? +(sums.r4 / cnts.r4).toFixed(2) : null,
      avgR5: cnts.r5 ? +(sums.r5 / cnts.r5).toFixed(2) : null,
      avgR6: cnts.r6 ? +(sums.r6 / cnts.r6).toFixed(2) : null,
      avgTotal: cnts.total ? +(sums.total / cnts.total).toFixed(2) : null,
      groupList: groupList
    });
  });
  faculty.sort((a, b) => b.totalStudents - a.totalStudents);

  /* ----------------------------------------------------------
   * 11. Review stats (global)
   * -------------------------------------------------------- */
  const reviewStats = {};
  Object.keys(REVIEW_MAX).forEach(function (k) {
    reviewStats[k] = {
      key: k,
      max: REVIEW_MAX[k],
      label: REVIEW_LABELS[k],
      filled: 0,
      missing: 0,
      avg: null,
      sum: 0
    };
  });

  students.forEach(function (s) {
    Object.keys(REVIEW_MAX).forEach(function (k) {
      const v = s.reviews ? s.reviews[k] : null;
      if (v !== null && v !== undefined) {
        reviewStats[k].filled++;
        reviewStats[k].sum += v;
      } else {
        reviewStats[k].missing++;
      }
    });
  });

  Object.keys(reviewStats).forEach(function (k) {
    const r = reviewStats[k];
    r.avg = r.filled ? +(r.sum / r.filled).toFixed(2) : null;
    r.pct = r.filled ? Math.round((r.avg / r.max) * 100) : 0;
    delete r.sum;
  });

  /* ----------------------------------------------------------
   * 12. Filter helper lists (unique values)
   * -------------------------------------------------------- */
  const courses = {}, faculties = {}, topics = {};

  students.forEach(function (s) {
    if (s.course) courses[s.course] = (courses[s.course] || 0) + 1;
    if (s.mentor && !/^#N\/A$/i.test(s.mentor)) {
      faculties[s.mentor] = (faculties[s.mentor] || 0) + 1;
    }
    if (s.topic) topics[s.topic] = (topics[s.topic] || 0) + 1;
  });

  const courseList  = Object.keys(courses).sort().map(k => ({ name: k, count: courses[k] }));
  const facultyList = Object.keys(faculties).sort().map(k => ({ name: k, count: faculties[k] }));
  const topicList   = Object.keys(topics).sort().map(k => ({ name: k, count: topics[k] }));

  /* ----------------------------------------------------------
   * 13. Pre‑computed category lists
   * -------------------------------------------------------- */
  const lists = {
    hold:         students.filter(s => s.flags.hold),
    exchange:     students.filter(s => s.flags.exchange),
    notSubmitted: students.filter(s => s.flags.notSubmitted),
    consultancy:  students.filter(s => s.flags.consultancy),
    idp:          students.filter(s => s.flags.idp)
  };

  /* ----------------------------------------------------------
   * 14. Return payload
   * -------------------------------------------------------- */
  return {
    ok: true,
    total: students.length,
    updatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd MMM yyyy, HH:mm:ss'
    ),
    counts: counts,
    students: students,
    groups: groups,
    faculty: faculty,
    reviewStats: reviewStats,
    lists: lists,
    sheetStatus: sheetStatus,
    courseList: courseList,
    facultyList: facultyList,
    topicList: topicList,
    author: 'Parekh Payal'
  };
}

/* ============================================================
 *  applyFlag — set a category flag on a student (or create one)
 * ============================================================ */
function applyFlag(map, s, flag) {
  if (!s.enrollment && !s.name) return;
  const key = s.enrollment || ('NAME:' + s.name);

  if (map.has(key)) {
    map.get(key).flags[flag] = true;
    if (!map.get(key).mentor && s.mentor) map.get(key).mentor = s.mentor;
  } else {
    const flags = {
      udp: false, idp: false, consultancy: false,
      hold: false, exchange: false, notSubmitted: false
    };
    flags[flag] = true;

    map.set(key, {
      enrollment: s.enrollment,
      name: s.name,
      course: s.program || '',
      mentor: s.mentor || '',
      remark: s.remark || '',
      projectType: flag.toUpperCase(),
      reviews: {},
      flags: flags
    });
  }
}

/* ============================================================
 *  parseProjectCountSheet — read hold / exchange / not‑submitted
 * ============================================================ */
function parseProjectCountSheet(sheet) {
  const result = { hold: [], exchange: [], notSubmitted: [] };
  if (!sheet) return result;

  const data = sheet.getDataRange().getDisplayValues();
  let section = null, started = false;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const a = String(row[0] || '').trim();
    const b = String(row[1] || '').trim();
    const c = String(row[2] || '').trim();

    if (/^semester\s*exchange$/i.test(a)) {
      section = 'exchange'; started = true; continue;
    }
    if (/^not\s*submitted\s*mini\s*project/i.test(a)) {
      section = 'notSubmitted'; started = true; continue;
    }
    if (/^enrollment\s*no/i.test(b)) {
      started = true;
      if (!section) section = 'hold';
      continue;
    }

    if (!started || !section || !b || !c) continue;

    const enr = normalize(b);
    if (!enr || !/^\d{8,}$/.test(enr)) continue;

    result[section].push({
      enrollment: enr,
      name: c,
      program: String(row[3] || '').trim(),
      mentor:  String(row[5] || '').trim(),
      remark:  String(row[6] || '').trim()
    });
  }

  return result;
}

/* ============================================================
 *  Manual test (run from the Apps Script editor if needed)
 * ============================================================ */
function testDashboard() {
  const p = getDashboardData(true);
  Logger.log('Total: ' + p.total +
             ' | Groups: ' + p.groups.length +
             ' | Faculty: ' + p.faculty.length);
  Logger.log('Courses: ' + p.courseList.length +
             ' | Topics: ' + p.topicList.length);
  Logger.log('Counts: ' + JSON.stringify(p.counts));
}
