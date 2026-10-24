"""
================================================================================
 TEST SCHOOL — TIME TABLE PLANNER  (single-file, runs in VS Code)
================================================================================
 UPDATED: ingest_all() now scans sub-folders recursively via os.walk()
================================================================================
 HOW TO RUN:
   1. Put this file (app.py) in a folder. Put your .xlsx files in ./data/
      (sub-folders allowed: data/class8/, data/class9/, etc.)
   2. pip install fastapi uvicorn openpyxl pandas python-multipart
   3. python app.py
   4. Open http://127.0.0.1:8000
================================================================================
"""

import os, re, json, sqlite3, datetime as dt, calendar, random
from contextlib import contextmanager
from typing import Optional, List
from fastapi import FastAPI, HTTPException, UploadFile, File, Body
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

DB_PATH = "timetable.db"
DATA_DIR = "data"
ACADEMIC_YEAR = "2026-27"
YEAR_START = dt.date(2026, 6, 1)
YEAR_END   = dt.date(2027, 4, 30)

# ------------------------------------------------------------------ DB LAYER
@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()

SCHEMA = """
CREATE TABLE IF NOT EXISTS academic_year(
  id INTEGER PRIMARY KEY, name TEXT, start_date TEXT, end_date TEXT, status TEXT DEFAULT 'active');
CREATE TABLE IF NOT EXISTS grade(
  id INTEGER PRIMARY KEY, name TEXT UNIQUE, level_group TEXT,
  day_start TEXT, day_end TEXT);
CREATE TABLE IF NOT EXISTS section(
  id INTEGER PRIMARY KEY, grade_id INTEGER REFERENCES grade(id), label TEXT, strength INTEGER,
  UNIQUE(grade_id,label));
CREATE TABLE IF NOT EXISTS break_period(
  id INTEGER PRIMARY KEY, label TEXT, start_time TEXT, end_time TEXT, applies_to TEXT);
CREATE TABLE IF NOT EXISTS teacher(
  id INTEGER PRIMARY KEY, name TEXT, employment_status TEXT, shift_start TEXT, shift_end TEXT,
  is_flexible INTEGER DEFAULT 0, max_periods_per_day INTEGER DEFAULT 6, max_periods_per_week INTEGER DEFAULT 30);
CREATE TABLE IF NOT EXISTS subject(id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT, category TEXT);
CREATE TABLE IF NOT EXISTS teacher_capability(
  teacher_id INTEGER REFERENCES teacher(id), subject_id INTEGER REFERENCES subject(id),
  proficiency_rank INTEGER DEFAULT 1, PRIMARY KEY(teacher_id, subject_id));
CREATE TABLE IF NOT EXISTS syllabus_document(
  id INTEGER PRIMARY KEY, grade_id INTEGER, subject_id INTEGER, semester INTEGER,
  source_file TEXT, ingested_at TEXT);
CREATE TABLE IF NOT EXISTS syllabus_chapter(
  id INTEGER PRIMARY KEY, document_id INTEGER REFERENCES syllabus_document(id),
  chapter_no TEXT, chapter_name TEXT, target_month TEXT, page_range TEXT);
CREATE TABLE IF NOT EXISTS syllabus_topic(
  id INTEGER PRIMARY KEY, chapter_id INTEGER REFERENCES syllabus_chapter(id),
  sub_topic_no TEXT, sub_topic_name TEXT, learning_focus TEXT,
  target_month_start TEXT, target_month_end TEXT, estimated_periods INTEGER DEFAULT 1,
  granularity TEXT DEFAULT 'SUB_TOPIC');
CREATE TABLE IF NOT EXISTS timetable_version(
  id INTEGER PRIMARY KEY AUTOINCREMENT, academic_year_id INTEGER, version_no INTEGER,
  generated_by TEXT, status TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS schedule_slot(
  id INTEGER PRIMARY KEY AUTOINCREMENT, version_id INTEGER, section_id INTEGER,
  slot_date TEXT, start_time TEXT, end_time TEXT, slot_type TEXT,
  subject_id INTEGER, teacher_id INTEGER, syllabus_topic_id INTEGER,
  source TEXT DEFAULT 'AUTO', override_reason TEXT);
CREATE INDEX IF NOT EXISTS idx_slot_lookup ON schedule_slot(version_id, slot_date, section_id);
CREATE INDEX IF NOT EXISTS idx_slot_teacher ON schedule_slot(version_id, slot_date, teacher_id);
CREATE TABLE IF NOT EXISTS exam_event(
  id INTEGER PRIMARY KEY AUTOINCREMENT, grade_id INTEGER, assessment_code TEXT,
  subject_id INTEGER, revision_start TEXT, revision_end TEXT, exam_date TEXT,
  session TEXT, start_time TEXT, end_time TEXT);
CREATE TABLE IF NOT EXISTS generation_run(
  id INTEGER PRIMARY KEY AUTOINCREMENT, version_id INTEGER, status TEXT,
  seed INTEGER, solver_stats TEXT, conflict_report TEXT, started_at TEXT, finished_at TEXT);
CREATE TABLE IF NOT EXISTS conflict_log(
  id INTEGER PRIMARY KEY AUTOINCREMENT, version_id INTEGER, severity TEXT, rule_code TEXT,
  detail TEXT, resolved INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS audit_log(
  id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, slot_id INTEGER, action TEXT,
  before TEXT, after TEXT, created_at TEXT);
"""

def init_db():
    with db() as c:
        c.executescript(SCHEMA)
        c.execute("INSERT OR IGNORE INTO academic_year(id,name,start_date,end_date,status) VALUES(1,?,?,?,'active')",
                  (ACADEMIC_YEAR, YEAR_START.isoformat(), YEAR_END.isoformat()))
        grades = [
            ("Nursery","PRE_PRIMARY","08:40","12:30"),("LKG","PRE_PRIMARY","08:40","12:30"),
            ("UKG","PRE_PRIMARY","08:40","12:30"),
        ] + [(str(i),"MIDDLE" if i<=8 else "SECONDARY","09:00","15:45") for i in range(1,10)] + [
            ("10","BOARD","09:00","17:30")]
        for g in grades:
            c.execute("INSERT OR IGNORE INTO grade(name,level_group,day_start,day_end) VALUES(?,?,?,?)", g)
        for row in c.execute("SELECT id,name FROM grade").fetchall():
            n_sec = 4 if (row["name"] in ("Nursery","LKG","UKG") or row["name"].isdigit() and int(row["name"])<=7) else 2
            strength = 35 if n_sec==4 else 40
            for lbl in "ABCD"[:n_sec]:
                c.execute("INSERT OR IGNORE INTO section(grade_id,label,strength) VALUES(?,?,?)",
                          (row["id"], lbl, strength))
        for b in [("First Break","10:20","10:30"),("Lunch Break","12:25","12:55"),("Third Break","14:15","14:25")]:
            c.execute("INSERT OR IGNORE INTO break_period(label,start_time,end_time,applies_to) VALUES(?,?,?,'ALL')", b)
        subs = [("TELUGU","Telugu","LANGUAGE"),("ENGLISH","English","LANGUAGE"),("HINDI","Hindi","LANGUAGE"),
                ("MATHS","Mathematics","CORE"),("PHYSICS","Physics","CORE"),("BIOLOGY","Biology","CORE"),
                ("SOCIAL","Social","CORE"),("ICT","ICT","ACTIVITY"),("PET","PET","ACTIVITY"),
                ("TAEKWONDO","Taekwondo","ACTIVITY"),("LIBRARY","Library","ACTIVITY"),
                ("DRAWING","Drawing","ACTIVITY"),("DANCE","Dance","ACTIVITY")]
        for s in subs:
            c.execute("INSERT OR IGNORE INTO subject(code,name,category) VALUES(?,?,?)", s)

# ----------------------------------------------------------- MONTH PARSING
MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}
MONTH_ABBR = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}
def parse_target_month(text, year_start=YEAR_START):
    if not text: return None, None
    t = re.sub(r"[–—−]", "-", str(text)).lower()
    parts = [p.strip() for p in t.split("-") if p.strip()]
    def month_num(p):
        p = p.split()[0]
        return MONTHS.get(p) or MONTH_ABBR.get(p)
    if not parts: return None, None
    m1 = month_num(parts[0])
    m2 = month_num(parts[-1]) if len(parts) > 1 else m1
    if not m1 or not m2: return None, None
    def date_for(m, first):
        y = year_start.year
        if m < year_start.month: y += 1
        last = calendar.monthrange(y, m)[1]
        return dt.date(y, m, 1 if first else last)
    start = date_for(m1, True)
    end   = date_for(m2, False)
    if end < start: end = date_for(m2, False)
    return start, end

# ---------------------------------------------------------- EXCEL INGESTION
SUBJECT_HINTS = {
    "maths":"MATHS","mat":"MATHS","mathematics":"MATHS",
    "social":"SOCIAL","s.s":"SOCIAL","social_science":"SOCIAL","social science":"SOCIAL",
    "science":"SCIENCE","physics":"PHYSICS","biology":"BIOLOGY",
    "telugu":"TELUGU","english":"ENGLISH","hindi":"HINDI","languages":"LANGUAGES",
}
def guess_subject(filename):
    fn = filename.lower()
    for k, v in SUBJECT_HINTS.items():
        if k in fn: return v
    return "UNKNOWN"

def guess_grade(filename):
    m = re.search(r"class[_\s]*(\d+)", filename, re.I)
    return m.group(1) if m else None

def parse_syllabus_sheet(ws, filename):
    rows = []
    header = None
    header_row_idx = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if not any(row): continue
        cells = [str(c).strip() if c is not None else "" for c in row]
        joined = " ".join(cells).lower()
        if "chapter" in joined and ("sub-topic" in joined or "subtopic" in joined or "learning" in joined):
            header = cells; header_row_idx = i; break
    if not header: return rows
    def col(name):
        for i, h in enumerate(header):
            if name in h.lower(): return i
        return None
    idx = {k: col(v) for k, v in {
        "chapter_no":"chapter no","chapter_name":"chapter name","target_month":"target month",
        "page_range":"page","sub_topic_no":"sub-topic no","sub_topic_name":"sub-topic name",
        "learning_focus":"learning focus",
    }.items()}
    for row in ws.iter_rows(min_row=header_row_idx+2, values_only=True):
        if not any(row): continue
        def get(k):
            j = idx[k]
            return str(row[j]).strip() if j is not None and j < len(row) and row[j] is not None else ""
        ch_no = get("chapter_no") or get("sub_topic_no")
        ch_name = get("chapter_name")
        if not ch_no and not ch_name: continue
        rows.append({
            "chapter_no": ch_no, "chapter_name": ch_name,
            "target_month": get("target_month"), "page_range": get("page_range"),
            "sub_topic_no": get("sub_topic_no"), "sub_topic_name": get("sub_topic_name"),
            "learning_focus": get("learning_focus"),
        })
    return rows

def ingest_syllabus_file(path):
    filename = os.path.basename(path)
    grade_name = guess_grade(filename)
    subject_code = guess_subject(filename)
    if not grade_name:
        return {"file": filename, "skipped": "no grade in filename"}
    wb = openpyxl.load_workbook(path, data_only=True)
    inserted = 0
    replaced_docs = 0
    for si, sheet_name in enumerate(wb.sheetnames):
        ws = wb[sheet_name]
        sem = 2 if re.search(r"sem[\s_-]*2|semester[\s_-]*2", sheet_name, re.I) else 1
        rows = parse_syllabus_sheet(ws, filename)
        if not rows:
            continue
        with db() as c:
            g = c.execute("SELECT id FROM grade WHERE name=?", (grade_name,)).fetchone()
            s = c.execute("SELECT id FROM subject WHERE code=?", (subject_code,)).fetchone()
            if not g:
                return {"file": filename, "error": f"grade {grade_name} not found"}
            grade_id = g["id"]
            subject_id = s["id"] if s else None
            if not subject_id:
                c.execute("INSERT OR IGNORE INTO subject(code,name,category) VALUES(?,?,?)",
                          (subject_code, subject_code.title(), "CORE"))
                subject_id = c.execute("SELECT id FROM subject WHERE code=?", (subject_code,)).fetchone()["id"]

            # --- NEW: purge any prior doc for this (grade, subject, semester) ---
            old = c.execute(
                "SELECT id FROM syllabus_document WHERE grade_id=? AND subject_id=? AND semester=?",
                (grade_id, subject_id, sem)).fetchall()
            for od in old:
                c.execute("""DELETE FROM syllabus_topic WHERE chapter_id IN
                             (SELECT id FROM syllabus_chapter WHERE document_id=?)""", (od["id"],))
                c.execute("DELETE FROM syllabus_chapter WHERE document_id=?", (od["id"],))
                c.execute("DELETE FROM syllabus_document WHERE id=?", (od["id"],))
                replaced_docs += 1
            # -----------------------------------------------------------------

            cur = c.execute("INSERT INTO syllabus_document(grade_id,subject_id,semester,source_file,ingested_at) VALUES(?,?,?,?,?)",
                            (grade_id, subject_id, sem, filename, dt.datetime.now().isoformat()))
            doc_id = cur.lastrowid
            chapter_cache = {}
            for r in rows:
                key = (r["chapter_no"], r["chapter_name"])
                if key not in chapter_cache:
                    cur = c.execute("INSERT INTO syllabus_chapter(document_id,chapter_no,chapter_name,target_month,page_range) VALUES(?,?,?,?,?)",
                                    (doc_id, r["chapter_no"], r["chapter_name"], r["target_month"], r["page_range"]))
                    chapter_cache[key] = cur.lastrowid
                current_chapter_id = chapter_cache[key]
                ms, me = parse_target_month(r["target_month"])
                c.execute("""INSERT INTO syllabus_topic(chapter_id,sub_topic_no,sub_topic_name,learning_focus,
                             target_month_start,target_month_end,estimated_periods)
                             VALUES(?,?,?,?,?,?,1)""",
                          (current_chapter_id, r["sub_topic_no"], r["sub_topic_name"],
                           r["learning_focus"], ms.isoformat() if ms else None,
                           me.isoformat() if me else None))
                inserted += 1
    return {"file": filename, "topics_inserted": inserted, "documents_replaced": replaced_docs}

def ingest_teachers_file(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    inserted = 0; skipped = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0: continue
        if not row or not row[0]: continue
        name = str(row[0]).strip()
        cap = str(row[1]).strip() if row[1] else ""
        status = str(row[2]).strip().upper() if row[2] else "FULL_TIME"
        shift = str(row[3]).strip() if row[3] else ""
        if "flexible" in shift.lower() or "part" in status.lower():
            s_start, s_end, flex = None, None, 1
        else:
            m = re.match(r"(\d{1,2}):(\d{2})\s*(AM|PM)", shift, re.I)
            m2 = re.search(r"to\s+(\d{1,2}):(\d{2})\s*(AM|PM)", shift, re.I)
            def to24(m):
                if not m: return None
                h, mi, ap = int(m.group(1)), m.group(2), m.group(3).upper()
                if ap == "PM" and h != 12: h += 12
                if ap == "AM" and h == 12: h = 0
                return f"{h:02d}:{mi}"
            s_start = to24(m); s_end = to24(m2)
            flex = 0
        cap_code = None
        for k, v in SUBJECT_HINTS.items():
            if k in cap.lower(): cap_code = v; break
        if not cap_code:
            cap_code = cap.upper().replace(" ", "_")
        with db() as c:
            cur = c.execute("""INSERT INTO teacher(name,employment_status,shift_start,shift_end,is_flexible)
                               VALUES(?,?,?,?,?)""", (name, status, s_start, s_end, flex))
            tid = cur.lastrowid
            subj = c.execute("SELECT id FROM subject WHERE code=? OR name=?", (cap_code, cap)).fetchone()
            if not subj:
                c.execute("INSERT OR IGNORE INTO subject(code,name,category) VALUES(?,?,?)",
                          (cap_code, cap, "CORE"))
                subj = c.execute("SELECT id FROM subject WHERE code=?", (cap_code,)).fetchone()
            c.execute("INSERT OR IGNORE INTO teacher_capability(teacher_id,subject_id) VALUES(?,?)",
                      (tid, subj["id"]))
            inserted += 1
    return {"file": os.path.basename(path), "teachers_inserted": inserted, "skipped": skipped}

def ingest_exams_file(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    inserted = 0
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        grade_name = None
        m = re.search(r"class\s*(\d+)", sheet_name, re.I)
        if m: grade_name = m.group(1)
        rows = list(ws.iter_rows(values_only=True))
        for i, row in enumerate(rows):
            cells = [str(c).strip() if c is not None else "" for c in row]
            line = " ".join(cells).lower()
            if not grade_name: continue
            if re.search(r"class\s*\d+\s*-\s*(fa-\d|summative|sa-\d)", line):
                hdr_idx = None
                for j in range(i, min(i+3, len(rows))):
                    h = " ".join(str(x).lower() for x in rows[j] if x)
                    if "assessment" in h and "exam date" in h:
                        hdr_idx = j; break
                if hdr_idx is None: continue
                header = [str(x).strip().lower() if x else "" for x in rows[hdr_idx]]
                am = re.search(r"(fa-\d|sa-\d|summative assessment-\d)", line)
                assessment = am.group(1).upper() if am else "EXAM"
                for j in range(hdr_idx+1, len(rows)):
                    r = rows[j]
                    if not r or not r[0]:
                        if j > hdr_idx+1 and not any(r): break
                        continue
                    vals = [str(x).strip() if x is not None else "" for x in r]
                    if "assessment" in vals[0].lower(): break
                    if len(vals) < 5: continue
                    exam_date_raw = vals[2]
                    sess = vals[3]
                    subj_name = vals[4]
                    dm = re.search(r"(\d{2})/(\d{2})/(\d{2})", exam_date_raw)
                    if not dm: continue
                    exam_date = dt.date(2000+int(dm.group(3)), int(dm.group(2)), int(dm.group(1)))
                    rev_start = rev_end = None
                    rm = re.search(r"(\d{2})/(\d{2})/(\d{2})\s*to\s*(\d{2})/(\d{2})/(\d{2})", vals[1])
                    if rm:
                        rev_start = dt.date(2000+int(rm.group(3)), int(rm.group(2)), int(rm.group(1)))
                        rev_end = dt.date(2000+int(rm.group(6)), int(rm.group(5)), int(rm.group(4)))
                    session = "FN" if "f.n" in sess.lower() or "forenoon" in sess.lower() else \
                              "AN" if "a.n" in sess.lower() or "afternoon" in sess.lower() else "FULL"
                    tm = re.search(r"(\d{1,2})[.:](\d{2})\s*(a\.?m|p\.?m)", sess, re.I)
                    start_time = end_time = None
                    if tm:
                        h = int(tm.group(1)); mi = tm.group(2); ap = tm.group(3).lower()
                        if "p" in ap and h != 12: h += 12
                        start_time = f"{h:02d}:{mi}"
                        em = re.search(r"(\d{1,2})[.:](\d{2})\s*(a\.?m|p\.?m)", sess[tm.end():], re.I)
                        if em:
                            h2 = int(em.group(1)); mi2 = em.group(2); ap2 = em.group(3).lower()
                            if "p" in ap2 and h2 != 12: h2 += 12
                            end_time = f"{h2:02d}:{mi2}"
                    with db() as c:
                        g = c.execute("SELECT id FROM grade WHERE name=?", (grade_name,)).fetchone()
                        if not g: continue
                        subj_code = None
                        for k, v in SUBJECT_HINTS.items():
                            if k in subj_name.lower(): subj_code = v; break
                        sid = None
                        if subj_code:
                            s = c.execute("SELECT id FROM subject WHERE code=?", (subj_code,)).fetchone()
                            sid = s["id"] if s else None
                        c.execute("""INSERT INTO exam_event(grade_id,assessment_code,subject_id,
                                     revision_start,revision_end,exam_date,session,start_time,end_time)
                                     VALUES(?,?,?,?,?,?,?,?,?)""",
                                  (g["id"], assessment, sid,
                                   rev_start.isoformat() if rev_start else None,
                                   rev_end.isoformat() if rev_end else None,
                                   exam_date.isoformat(), session, start_time, end_time))
                        inserted += 1
    return {"file": os.path.basename(path), "exam_events_inserted": inserted}

# ============================================================
# *** UPDATED: RECURSIVE FOLDER SCAN via os.walk ***
# ============================================================
def ingest_all(data_dir=DATA_DIR):
    """Recursively walk data_dir and ingest every .xlsx/.xlsm found.
       Sub-folders (class8/, class9/, etc.) are supported."""
    results = []
    if not os.path.isdir(data_dir):
        return {"error": f"{data_dir}/ not found"}
    with db() as c:
        c.execute("DELETE FROM teacher_capability")
        c.execute("DELETE FROM teacher")
        c.execute("DELETE FROM syllabus_topic")
        c.execute("DELETE FROM syllabus_chapter")
        c.execute("DELETE FROM syllabus_document")
        c.execute("DELETE FROM exam_event")
    for root, _dirs, files in os.walk(data_dir):
        for fn in sorted(files):
            if not fn.lower().endswith((".xlsx", ".xlsm")):
                continue
            path = os.path.join(root, fn)
            # Relative path helps you see which sub-folder a file came from
            rel = os.path.relpath(path, data_dir)
            try:
                low = fn.lower()
                if "teacher" in low:
                    r = ingest_teachers_file(path); r["source"] = rel; results.append(r)
                elif "exam" in low or "time table" in low:
                    r = ingest_exams_file(path); r["source"] = rel; results.append(r)
                else:
                    r = ingest_syllabus_file(path); r["source"] = rel; results.append(r)
            except Exception as e:
                results.append({"file": rel, "error": str(e)})
    return {"ingested": results, "total_files": len(results),
            "scanned_root": os.path.abspath(data_dir)}

# ---------------------------------------------------------------- SCHEDULING
def build_day_slot_template(grade_row, breaks):
    def t2m(t):
        h, m = t.split(":"); return int(h)*60 + int(m)
    def m2t(x): return f"{x//60:02d}:{x%60:02d}"
    start = t2m(grade_row["day_start"]); end = t2m(grade_row["day_end"])
    bset = sorted([(t2m(b["start_time"]), t2m(b["end_time"])) for b in breaks])
    slots = []
    cur = start
    for bs, be in bset:
        if bs < start or be > end: continue
        if cur < bs:
            length = bs - cur
            if length <= 0: continue
            n = max(1, round(length / 45))
            per = length // n
            for k in range(n):
                slots.append((m2t(cur + k*per), m2t(cur + (k+1)*per)))
        cur = max(cur, be)
    if cur < end:
        length = end - cur
        if length > 0:
            n = max(1, round(length / 45))
            per = length // n
            for k in range(n):
                slots.append((m2t(cur + k*per), m2t(cur + (k+1)*per)))
    return slots

def get_teachers_for_subject(c, subject_id):
    return c.execute("""SELECT t.* FROM teacher t
        JOIN teacher_capability tc ON tc.teacher_id=t.id
        WHERE tc.subject_id=? ORDER BY tc.proficiency_rank, t.id""", (subject_id,)).fetchall()

def teacher_free(c, version_id, date, start, end, teacher_id):
    row = c.execute("""SELECT 1 FROM schedule_slot WHERE version_id=? AND slot_date=? AND teacher_id=?
                       AND slot_type='TEACHING' AND NOT (end_time<=? OR start_time>=?)""",
                    (version_id, date, teacher_id, start, end)).fetchone()
    return row is None

def section_free(c, version_id, date, start, end, section_id):
    row = c.execute("""SELECT 1 FROM schedule_slot WHERE version_id=? AND slot_date=? AND section_id=?
                       AND NOT (end_time<=? OR start_time>=?)""",
                    (version_id, date, section_id, start, end)).fetchone()
    return row is None

def get_exam_block(c, grade_id, date):
    r = c.execute("SELECT * FROM exam_event WHERE grade_id=? AND exam_date=?", (grade_id, date)).fetchall()
    if r: return {"type":"EXAM","events":[dict(x) for x in r]}
    r = c.execute("""SELECT * FROM exam_event WHERE grade_id=? AND revision_start<=? AND revision_end>=?""",
                  (grade_id, date, date)).fetchall()
    if r: return {"type":"REVISION","events":[dict(x) for x in r]}
    return None

def generate_timetable(sections=None, grade_filter=None, seed=42, version_label="AUTO"):
    random.seed(seed)
    today = dt.date.today()
    effective_start = max(YEAR_START, today)
    if effective_start > YEAR_END:
        raise HTTPException(400, f"No schedulable days: effective_start {effective_start} > YEAR_END {YEAR_END}")

    with db() as c:
        maxv = c.execute("SELECT COALESCE(MAX(version_no),0) v FROM timetable_version").fetchone()["v"]
        cur = c.execute("INSERT INTO timetable_version(academic_year_id,version_no,generated_by,status,created_at) VALUES(1,?,?,?,?)",
                        (maxv+1, version_label, "draft", dt.datetime.now().isoformat()))
        version_id = cur.lastrowid
        run = c.execute("INSERT INTO generation_run(version_id,status,seed,started_at) VALUES(?,?,?,?)",
                        (version_id, "running", seed, dt.datetime.now().isoformat()))
        run_id = run.lastrowid

    stats = {"sections": 0, "slots_inserted": 0, "skipped_no_teacher": 0,
             "skipped_no_syllabus": 0, "warnings": []}

    try:
        with db() as c:
            q = ("SELECT s.*, g.name gname, g.level_group, g.day_start, g.day_end "
                 "FROM section s JOIN grade g ON g.id=s.grade_id")
            params = []
            if grade_filter:
                q += " WHERE g.name=?"; params.append(grade_filter)
            if sections:
                q += " AND s.id IN (%s)" % ",".join("?" * len(sections)); params.extend(sections)
            sections_rows = c.execute(q, params).fetchall()

            breaks = [dict(r) for r in c.execute("SELECT * FROM break_period").fetchall()]
            days = [effective_start + dt.timedelta(days=i) for i in range((YEAR_END - effective_start).days + 1)]

            for srow in sections_rows:
                stats["sections"] += 1
                subjects_for_grade = [dict(r) for r in c.execute("""
                    SELECT DISTINCT sd.subject_id, sub.code, sub.name FROM syllabus_document sd
                    JOIN subject sub ON sub.id=sd.subject_id WHERE sd.grade_id=?""",
                    (srow["grade_id"],)).fetchall()]
                
                if not subjects_for_grade:
                    stats["skipped_no_syllabus"] += 1
                    continue

                exams = c.execute("SELECT exam_date FROM exam_event WHERE grade_id=?", (srow["grade_id"],)).fetchall()
                exam_dates = [dt.date.fromisoformat(e["exam_date"]) for e in exams if e["exam_date"]]
                first_exam = min(exam_dates) if exam_dates else dt.date(2027, 3, 1)

                queues = {}
                for sub in subjects_for_grade:
                    topics = [dict(r) for r in c.execute("""
                        SELECT st.*, sc.chapter_no, sc.chapter_name FROM syllabus_topic st
                        JOIN syllabus_chapter sc ON sc.id=st.chapter_id
                        JOIN syllabus_document sd ON sd.id=sc.document_id
                        WHERE sd.grade_id=? AND sd.subject_id=?
                        ORDER BY COALESCE(st.target_month_start,'9999'), st.id""",
                        (srow["grade_id"], sub["subject_id"])).fetchall()]
                    
                    if not topics: continue
                    
                    teachers = get_teachers_for_subject(c, sub["subject_id"])
                    if not teachers:
                        stats["skipped_no_teacher"] += 1
                        continue
                    
                    queues[sub["subject_id"]] = {
                        "topics": topics, 
                        "teachers": teachers, 
                        "name": sub["name"],
                        "first_exam": first_exam
                    }

                if not queues: continue

                tmpl = build_day_slot_template(srow, breaks)
                
                for d in days:
                    if d.weekday() >= 6: continue
                    
                    block = get_exam_block(c, srow["grade_id"], d)
                    if block and block["type"] == "EXAM":
                        for (st, en) in tmpl:
                            c.execute("""INSERT INTO schedule_slot(version_id,section_id,slot_date,start_time,end_time,
                                         slot_type,source) VALUES(?,?,?,?,?,'EXAM','AUTO')""",
                                      (version_id, srow["id"], d.isoformat(), st, en))
                            stats["slots_inserted"] += 1
                        continue

                    revision_subjects = set()
                    if block and block["type"] == "REVISION":
                        for ev in block["events"]:
                            if ev["subject_id"]: revision_subjects.add(ev["subject_id"])

                    for (st, en) in tmpl:
                        if not any(q["topics"] for q in queues.values()):
                            c.execute("""INSERT INTO schedule_slot(version_id,section_id,slot_date,start_time,end_time,
                                         slot_type,source) VALUES(?,?,?,?,?,'FREE','AUTO')""",
                                      (version_id, srow["id"], d.isoformat(), st, en))
                            stats["slots_inserted"] += 1
                            continue

                        candidates = []
                        for subj_id, q in queues.items():
                            if not q["topics"]: continue
                            
                            next_topic = q["topics"][0]
                            deadline = next_topic.get("target_month_end") or "9999"
                            
                            if d > (q["first_exam"] - dt.timedelta(days=14)):
                                deadline = "URGENT"
                            elif next_topic.get("target_month_start"):
                                target_start = dt.date.fromisoformat(next_topic["target_month_start"])
                                if d >= (target_start - dt.timedelta(days=15)):
                                    deadline = target_start.isoformat()
                                    
                            candidates.append((deadline, subj_id, next_topic))
                        
                        candidates.sort()

                        placed = False
                        for deadline, subj_id, topic in candidates:
                            q = queues[subj_id]
                            
                            # Allow max 2 periods per subject per day, but try all subjects
                            today_count = c.execute("""SELECT COUNT(*) n FROM schedule_slot
                                WHERE version_id=? AND section_id=? AND slot_date=? AND subject_id=?""",
                                (version_id, srow["id"], d.isoformat(), subj_id)).fetchone()["n"]
                            
                            if today_count >= 2: 
                                continue
                            
                            chosen_teacher = None
                            for t in q["teachers"]:
                                if not t["is_flexible"]:
                                    if t["shift_start"] and t["shift_start"] > st: continue
                                    if t["shift_end"] and t["shift_end"] < en: continue
                                
                                wk_start = (d - dt.timedelta(days=d.weekday())).isoformat()
                                wk_end = (d + dt.timedelta(days=6 - d.weekday())).isoformat()
                                wk = c.execute("""SELECT COUNT(*) n FROM schedule_slot WHERE version_id=? AND teacher_id=?
                                                  AND slot_date BETWEEN ? AND ? AND slot_type='TEACHING'""",
                                               (version_id, t["id"], wk_start, wk_end)).fetchone()["n"]
                                if wk >= t["max_periods_per_week"]: continue
                                
                                dy = c.execute("""SELECT COUNT(*) n FROM schedule_slot WHERE version_id=? AND teacher_id=?
                                                  AND slot_date=? AND slot_type='TEACHING'""",
                                               (version_id, t["id"], d.isoformat())).fetchone()["n"]
                                if dy >= t["max_periods_per_day"]: continue
                                
                                if teacher_free(c, version_id, d.isoformat(), st, en, t["id"]):
                                    chosen_teacher = t; break
                            
                            if not chosen_teacher: continue
                            if not section_free(c, version_id, d.isoformat(), st, en, srow["id"]): continue
                            
                            slot_type = "REVISION" if subj_id in revision_subjects else "TEACHING"
                            c.execute("""INSERT INTO schedule_slot(version_id,section_id,slot_date,start_time,end_time,
                                         slot_type,subject_id,teacher_id,syllabus_topic_id,source)
                                         VALUES(?,?,?,?,?,?,?,?,?,'AUTO')""",
                                      (version_id, srow["id"], d.isoformat(), st, en, slot_type,
                                       subj_id, chosen_teacher["id"], topic["id"]))
                            q["topics"].pop(0)
                            stats["slots_inserted"] += 1
                            placed = True
                            break
                        
                        if not placed:
                            c.execute("""INSERT INTO schedule_slot(version_id,section_id,slot_date,start_time,end_time,
                                         slot_type,source) VALUES(?,?,?,?,?,'FREE','AUTO')""",
                                      (version_id, srow["id"], d.isoformat(), st, en))

            stats["completed"] = True

        with db() as c:
            c.execute("UPDATE generation_run SET status='succeeded', finished_at=?, solver_stats=? WHERE id=?",
                      (dt.datetime.now().isoformat(), json.dumps(stats), run_id))
            c.execute("UPDATE timetable_version SET status='active' WHERE id=?", (version_id,))
    except Exception as e:
        with db() as c:
            c.execute("UPDATE generation_run SET status='failed', finished_at=?, solver_stats=? WHERE id=?",
                      (dt.datetime.now().isoformat(), json.dumps({**stats, "error": str(e)}), run_id))
        raise
    return {"version_id": version_id, "run_id": run_id, "stats": stats}

def detect_conflicts(version_id):
    conflicts = []
    with db() as c:
        # 1. Teacher double-booking
        rows = c.execute("""SELECT slot_date,start_time,end_time,teacher_id,COUNT(*) n FROM schedule_slot
                            WHERE version_id=? AND slot_type='TEACHING' AND teacher_id IS NOT NULL
                            GROUP BY slot_date,start_time,end_time,teacher_id HAVING n>1""",
                         (version_id,)).fetchall()
        for r in rows:
            conflicts.append({"severity": "BLOCKING", "rule": "TEACHER_CLASH", "detail": dict(r)})

        # 2. Syllabus unplaced
        today = dt.date.today().isoformat()
        
        # Find the first exam date for the grade of this version
        # (Simplification: we check if a topic is unplaced AND its deadline is close)
        rows = c.execute("""SELECT st.id, st.sub_topic_name, st.target_month_end 
                            FROM syllabus_topic st
                            WHERE st.target_month_end IS NOT NULL
                              AND st.target_month_end >= ? 
                              AND NOT EXISTS (SELECT 1 FROM schedule_slot ss
                                              WHERE ss.version_id=? AND ss.syllabus_topic_id=st.id)""",
                         (today, version_id)).fetchall()
        
        for r in rows:
            # Only show a warning if the deadline is within the next 30 days
            target_date = dt.date.fromisoformat(r["target_month_end"])
            if (target_date - dt.date.today()).days <= 30:
                conflicts.append({"severity": "WARNING", "rule": "SYLLABUS_UNPLACED", "detail": dict(r)})

        # 3. Informational: expired topics
        expired = c.execute("""SELECT COUNT(*) n FROM syllabus_topic
                               WHERE target_month_end IS NOT NULL AND target_month_end < ?""",
                            (today,)).fetchone()["n"]
        if expired:
            conflicts.append({"severity": "INFO", "rule": "EXPIRED_TOPICS_IGNORED",
                              "detail": {"count": expired,
                                         "note": "These topics are past their target month and are excluded from generation."}})
    return conflicts

# ------------------------------------------------------------------- API APP
app = FastAPI(title="Test School Timetable Planner", version="1.1")

@app.on_event("startup")
def _startup():
    init_db()

@app.post("/api/ingest/all")
def api_ingest_all():
    return ingest_all()

@app.post("/api/ingest/upload")
async def api_upload(file: UploadFile = File(...)):
    os.makedirs(DATA_DIR, exist_ok=True)
    path = os.path.join(DATA_DIR, file.filename)
    with open(path, "wb") as f: f.write(await file.read())
    low = file.filename.lower()
    try:
        if "teacher" in low: return ingest_teachers_file(path)
        if "exam" in low or "time table" in low: return ingest_exams_file(path)
        return ingest_syllabus_file(path)
    except Exception as e:
        raise HTTPException(500, str(e))

@app.post("/api/generate")
def api_generate(payload: dict = Body(default={})):
    grade = payload.get("grade")
    seed = int(payload.get("seed", 42))
    return generate_timetable(grade_filter=grade, seed=seed)

@app.get("/api/versions")
def api_versions():
    with db() as c:
        return [dict(r) for r in c.execute("SELECT * FROM timetable_version ORDER BY id DESC").fetchall()]

@app.get("/api/versions/{vid}/sections/{sid}/day/{date}")
def api_day(vid: int, sid: int, date: str):
    with db() as c:
        rows = c.execute("""
            SELECT ss.*, sub.name subject_name, sub.code subject_code,
                   t.name teacher_name, st.sub_topic_name topic
            FROM schedule_slot ss
            LEFT JOIN subject sub ON sub.id=ss.subject_id
            LEFT JOIN teacher t ON t.id=ss.teacher_id
            LEFT JOIN syllabus_topic st ON st.id=ss.syllabus_topic_id
            WHERE ss.version_id=? AND ss.section_id=? AND ss.slot_date=?
            ORDER BY ss.start_time""", (vid, sid, date)).fetchall()
        return [dict(r) for r in rows]

@app.get("/api/versions/{vid}/sections/{sid}/week/{date}")
def api_week(vid: int, sid: int, date: str):
    d = dt.date.fromisoformat(date)
    start = d - dt.timedelta(days=d.weekday())
    end = start + dt.timedelta(days=6)
    with db() as c:
        rows = c.execute("""
            SELECT ss.*, sub.name subject_name, t.name teacher_name, st.sub_topic_name topic
            FROM schedule_slot ss
            LEFT JOIN subject sub ON sub.id=ss.subject_id
            LEFT JOIN teacher t ON t.id=ss.teacher_id
            LEFT JOIN syllabus_topic st ON st.id=ss.syllabus_topic_id
            WHERE ss.version_id=? AND ss.section_id=? AND ss.slot_date BETWEEN ? AND ?
            ORDER BY ss.slot_date, ss.start_time""", (vid, sid, start.isoformat(), end.isoformat())).fetchall()
        return [dict(r) for r in rows]
# --------------------------------------------------------- WEEK EXPORT HELPERS
def _week_bounds(date_str):
    d = dt.date.fromisoformat(date_str)
    start = d - dt.timedelta(days=d.weekday())          # Monday
    end   = start + dt.timedelta(days=5)                # Saturday
    return start, end

def _fetch_week_rows(vid, sid, date_str):
    start, end = _week_bounds(date_str)
    with db() as c:
        sec = c.execute("""SELECT s.label, g.name gname FROM section s
                           JOIN grade g ON g.id=s.grade_id WHERE s.id=?""", (sid,)).fetchone()
        rows = c.execute("""
            SELECT ss.slot_date, ss.start_time, ss.end_time, ss.slot_type, ss.source,
                   sub.name subject_name, sub.code subject_code,
                   t.name teacher_name, st.sub_topic_name topic,
                   sc.chapter_no
            FROM schedule_slot ss
            LEFT JOIN subject sub ON sub.id=ss.subject_id
            LEFT JOIN teacher t ON t.id=ss.teacher_id
            LEFT JOIN syllabus_topic st ON st.id=ss.syllabus_topic_id
            LEFT JOIN syllabus_chapter sc ON sc.id=st.chapter_id
            WHERE ss.version_id=? AND ss.section_id=? AND ss.slot_date BETWEEN ? AND ?
            ORDER BY ss.slot_date, ss.start_time""",
            (vid, sid, start.isoformat(), end.isoformat())).fetchall()
    return sec, [dict(r) for r in rows], start, end

# -------------------------------------------------------------- EXCEL EXPORT
@app.get("/api/export/week.xlsx")
def export_week_xlsx(version_id: int, section_id: int, date: str):
    sec, rows, start, end = _fetch_week_rows(version_id, section_id, date)
    if not sec:
        raise HTTPException(404, "section not found")

    # Build grid: {time_label: {iso_date: slot_dict}}
    by_time, times = {}, []
    for r in rows:
        key = f"{r['start_time']}–{r['end_time']}"
        if key not in by_time:
            by_time[key] = {}
            times.append(key)
        by_time[key][r["slot_date"]] = r
    times.sort()

    days = [start + dt.timedelta(days=i) for i in range(6)]  # Mon..Sat

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Timetable"

    # ---- Title ----
    title = f"Timetable  •  {sec['gname']}-{sec['label']}  •  {start.isoformat()} to {end.isoformat()}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    tc = ws.cell(row=1, column=1, value=title)
    tc.font = Font(bold=True, size=14, color="FFFFFF")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    tc.fill = PatternFill("solid", fgColor="1E293B")
    ws.row_dimensions[1].height = 28

    # ---- Header row ----
    header = ["Time"] + [d.strftime("%a %d-%b-%Y") for d in days]
    for j, h in enumerate(header, start=1):
        cell = ws.cell(row=2, column=j, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="475569")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ---- Fill grid ----
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    row_idx = 3
    for t in times:
        c1 = ws.cell(row=row_idx, column=1, value=t)
        c1.font = Font(bold=True)
        c1.fill = PatternFill("solid", fgColor="F1F5F9")
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = border

        for col_off, d in enumerate(days, start=2):
            iso = d.isoformat()
            s = by_time.get(t, {}).get(iso)
            cell = ws.cell(row=row_idx, column=col_off)
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            if not s:
                continue
            stype = s["slot_type"] or ""
            if stype == "EXAM":
                cell.value = "EXAM"
                cell.fill = PatternFill("solid", fgColor="FEE2E2")
            elif stype == "REVISION":
                cell.value = f"REVISION\n{s['subject_name'] or ''}\n{s['teacher_name'] or ''}"
                cell.fill = PatternFill("solid", fgColor="EDE9FE")
            elif stype == "FREE":
                cell.value = "FREE"
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
            else:
                parts = [s["subject_name"] or ""]
                if s["teacher_name"]: parts.append(s["teacher_name"])
                if s["topic"]:        parts.append(s["topic"][:40])
                cell.value = "\n".join(parts)
                cell.fill = PatternFill("solid", fgColor="EFF6FF")
        ws.row_dimensions[row_idx].height = 42
        row_idx += 1

    # Column widths
    ws.column_dimensions[get_column_letter(1)].width = 16
    for j in range(2, 8):
        ws.column_dimensions[get_column_letter(j)].width = 24

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"timetable_{sec['gname']}-{sec['label']}_week_{start.isoformat()}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )
@app.get("/api/sections")
def api_sections():
    with db() as c:
        return [dict(r) for r in c.execute("""
            SELECT s.id, s.label, g.name gname FROM section s JOIN grade g ON g.id=s.grade_id
            ORDER BY g.id, s.label""").fetchall()]

@app.get("/api/teachers")
def api_teachers():
    with db() as c:
        return [dict(r) for r in c.execute("SELECT * FROM teacher ORDER BY id").fetchall()]

@app.get("/api/subjects")
def api_subjects():
    with db() as c:
        return [dict(r) for r in c.execute("SELECT * FROM subject ORDER BY id").fetchall()]

@app.get("/api/syllabus/topics")
def api_topics(grade: str, subject_code: str):
    with db() as c:
        g = c.execute("SELECT id FROM grade WHERE name=?", (grade,)).fetchone()
        s = c.execute("SELECT id FROM subject WHERE code=? OR name=?", (subject_code, subject_code)).fetchone()
        if not g or not s: return []
        rows = c.execute("""
            SELECT st.id, st.sub_topic_no, st.sub_topic_name, st.target_month_start, st.target_month_end,
                   sc.chapter_no, sc.chapter_name
            FROM syllabus_topic st JOIN syllabus_chapter sc ON sc.id=st.chapter_id
            JOIN syllabus_document sd ON sd.id=sc.document_id
            WHERE sd.grade_id=? AND sd.subject_id=?
            ORDER BY CAST(sc.chapter_no AS INTEGER), st.id""", (g["id"], s["id"])).fetchall()
        return [dict(r) for r in rows]

@app.get("/api/teachers/eligible")
def api_eligible(grade: str, subject_code: str, date: str, start: str, end: str, version_id: int):
    with db() as c:
        s = c.execute("SELECT id FROM subject WHERE code=? OR name=?", (subject_code, subject_code)).fetchone()
        if not s: return []
        teachers = get_teachers_for_subject(c, s["id"])
        out = []
        for t in teachers:
            ok = True
            reason = None
            if not t["is_flexible"]:
                if t["shift_start"] and t["shift_start"] > start: ok, reason = False, "outside shift"
                if t["shift_end"] and t["shift_end"] < end: ok, reason = False, "outside shift"
            if ok and not teacher_free(c, version_id, date, start, end, t["id"]):
                ok, reason = False, "busy at this time"
            out.append({**dict(t), "eligible": ok, "reason": reason})
        return out

@app.patch("/api/slots/{slot_id}")
def api_override(slot_id: int, payload: dict = Body(...)):
    with db() as c:
        old = c.execute("SELECT * FROM schedule_slot WHERE id=?", (slot_id,)).fetchone()
        if not old: raise HTTPException(404, "slot not found")
        before = dict(old)
        new_subject = payload.get("subject_id", old["subject_id"])
        new_teacher = payload.get("teacher_id", old["teacher_id"])
        new_topic = payload.get("syllabus_topic_id", old["syllabus_topic_id"])
        reason = payload.get("reason", "manual override")
        if new_teacher and old["slot_type"] in ("TEACHING", "FREE"):
            if not teacher_free(c, old["version_id"], old["slot_date"], old["start_time"], old["end_time"], new_teacher) \
               and new_teacher != old["teacher_id"]:
                raise HTTPException(409, "teacher busy at that time")
        c.execute("""UPDATE schedule_slot SET subject_id=?, teacher_id=?, syllabus_topic_id=?,
                     slot_type='TEACHING', source='MANUAL', override_reason=? WHERE id=?""",
                  (new_subject, new_teacher, new_topic, reason, slot_id))
        c.execute("INSERT INTO audit_log(slot_id,action,before,after,created_at) VALUES(?,?,?,?,?)",
                  (slot_id, "OVERRIDE", json.dumps(before),
                   json.dumps({**before, "subject_id":new_subject,"teacher_id":new_teacher,
                               "syllabus_topic_id":new_topic,"source":"MANUAL"}),
                   dt.datetime.now().isoformat()))
        return {"ok": True, "slot": dict(c.execute("SELECT * FROM schedule_slot WHERE id=?", (slot_id,)).fetchone())}

@app.get("/api/conflicts")
def api_conflicts(version_id: int):
    return detect_conflicts(version_id)

@app.get("/api/stats")
def api_stats():
    with db() as c:
        return {
            "grades": c.execute("SELECT COUNT(*) n FROM grade").fetchone()["n"],
            "sections": c.execute("SELECT COUNT(*) n FROM section").fetchone()["n"],
            "teachers": c.execute("SELECT COUNT(*) n FROM teacher").fetchone()["n"],
            "topics": c.execute("SELECT COUNT(*) n FROM syllabus_topic").fetchone()["n"],
            "exam_events": c.execute("SELECT COUNT(*) n FROM exam_event").fetchone()["n"],
            "versions": c.execute("SELECT COUNT(*) n FROM timetable_version").fetchone()["n"],
            "slots": c.execute("SELECT COUNT(*) n FROM schedule_slot").fetchone()["n"],
        }

# --------------------------------------------------------------- FRONTEND UI
HTML = r"""
<!DOCTYPE html><html><head><meta charset="utf-8"><title>Test School · Timetable Planner</title>
<style>
*{box-sizing:border-box} body{margin:0;font-family:system-ui,Segoe UI,Arial;background:#f4f6fa;color:#1e293b}
header{background:#1e293b;color:#fff;padding:12px 20px;display:flex;gap:16px;align-items:center}
header h1{font-size:16px;margin:0;font-weight:600}
.tabs{display:flex;gap:4px;margin-left:auto}
.tab{padding:6px 14px;border-radius:6px;cursor:pointer;font-size:13px;background:#334155}
.tab.on{background:#3b82f6}
main{padding:16px;max-width:1500px;margin:0 auto}
.card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px;margin-bottom:14px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px}
.kpi{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px}
.kpi .v{font-size:22px;font-weight:700} .kpi .l{font-size:11px;color:#64748b;text-transform:uppercase}
button{background:#3b82f6;color:#fff;border:0;padding:7px 14px;border-radius:6px;cursor:pointer;font-size:13px}
button.sec{background:#e2e8f0;color:#1e293b} button:disabled{opacity:.5;cursor:not-allowed}
select,input{padding:6px 8px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px;background:#fff}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{border:1px solid #e2e8f0;padding:6px 8px;text-align:left}
th{background:#f1f5f9;font-weight:600}
.grid{display:grid;grid-template-columns:70px repeat(6,1fr);gap:2px}
.slot{background:#eff6ff;border:1px solid #bfdbfe;border-radius:6px;padding:4px 6px;font-size:11px;min-height:48px}
.slot.break{background:#fef3c7;border-color:#fcd34d;text-align:center;font-weight:600;font-size:10px}
.slot.exam{background:#fee2e2;border-color:#fca5a5;font-weight:600}
.slot.revision{background:#ede9fe;border-color:#c4b5fd}
.slot.manual{box-shadow:inset 3px 0 #f59e0b}
.slot .s{font-weight:600;font-size:12px} .slot .t{color:#64748b}
.hdr{background:#1e293b;color:#fff;padding:6px;text-align:center;border-radius:4px;font-size:12px;font-weight:600}
.hdr.time{background:#64748b}
.banner{padding:10px 14px;border-radius:8px;margin-bottom:10px;font-size:13px}
.warn{background:#fef3c7;border:1px solid #fcd34d} .ok{background:#dcfce7;border:1px solid #86efac}
.err{background:#fee2e2;border:1px solid #fca5a5}
.valid{font-size:12px;margin-top:6px} .valid .ok{color:#15803d} .valid .no{color:#b91c1c}
.hidden{display:none}
</style></head><body>
<header>
  <h1>🏫 Test School · Timetable Planner</h1>
  <div class="tabs">
    <div class="tab on" data-tab="dash">Dashboard</div>
    <div class="tab" data-tab="grid">Day Grid</div>
    <div class="tab" data-tab="manual">Manual Override</div>
  </div>
</header>
<main>
<section id="dash">
  <div class="card">
    <div class="row" style="margin-bottom:10px">
      <b>Setup</b>
      <button onclick="ingestAll()">📥 Ingest /data/ files (recursive)</button>
      <span id="ingestMsg" class="valid"></span>
    </div>
    <div class="row">
      <b>Generate</b>
      <select id="genGrade"><option value="">All grades</option></select>
      <input id="genSeed" type="number" value="42" style="width:80px">
      <button onclick="generate()">⚙️ Generate timetable</button>
      <span id="genMsg" class="valid"></span>
    </div>
  </div>
  <div class="card"><div class="kpis" id="kpis"></div></div>
  <div class="card"><b>Versions</b><div id="versions"></div></div>
  <div class="card"><b>Conflicts</b><div id="conflicts"></div></div>
</section>
<section id="grid" class="hidden">
  <div class="card">
    <div class="row">
      <select id="gridVersion"></select>
      <select id="gridSection"></select>
      <input type="date" id="gridDate">
      <button onclick="loadWeek()">Load week</button>
      <button class="sec" onclick="downloadWeekExcel()">⬇️ Download Excel</button>
      <span id="dlMsg" class="valid"></span>
    </div>
  </div>
  <div class="card"><div id="weekGrid"></div></div>
</section>
<section id="manual" class="hidden">
  <div class="card">
    <b>Find slot to override</b>
    <div class="row" style="margin-top:8px">
      <select id="manVersion"></select>
      <select id="manSection"></select>
      <input type="date" id="manDate">
      <button onclick="loadSlots()">Load slots</button>
    </div>
  </div>
  <div class="card"><div id="slotList"></div></div>
  <div class="card hidden" id="editPanel">
    <b>Edit slot</b>
    <div class="row" style="margin-top:8px">
      <label>Subject <select id="edSubject"></select></label>
      <label>Teacher <select id="edTeacher"></select></label>
      <label>Topic <select id="edTopic"></select></label>
    </div>
    <div class="row" style="margin-top:8px">
      <input id="edReason" placeholder="Reason for override" style="flex:1">
      <button onclick="saveOverride()">💾 Save</button>
      <button class="sec" onclick="closeEdit()">Cancel</button>
    </div>
    <div id="edValid" class="valid"></div>
  </div>
</section>
</main>
<script>
const $=id=>document.getElementById(id);
const API='/api';
let state={sections:[],subjects:[],teachers:[],versions:[],editingSlot:null};
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('on')); t.classList.add('on');
  ['dash','grid','manual'].forEach(id=>$(id).classList.add('hidden'));
  $(t.dataset.tab).classList.remove('hidden');
});
async function j(url, opts){const r=await fetch(url,opts); if(!r.ok) throw new Error(await r.text()); return r.json();}
async function boot(){
  state.sections=await j(API+'/sections');
  state.subjects=await j(API+'/subjects');
  state.teachers=await j(API+'/teachers');
  const grades=[...new Set(state.sections.map(s=>s.gname))];
  $('genGrade').innerHTML='<option value="">All grades</option>'+grades.map(g=>`<option>${g}</option>`).join('');
  const secOpts=state.sections.map(s=>`<option value="${s.id}">${s.gname}-${s.label}</option>`).join('');
  $('gridSection').innerHTML=secOpts; $('manSection').innerHTML=secOpts;
  $('edSubject').innerHTML=state.subjects.map(s=>`<option value="${s.id}" data-code="${s.code}">${s.name}</option>`).join('');
  $('edTeacher').innerHTML=state.teachers.map(t=>`<option value="${t.id}">${t.name}</option>`).join('');
  const today=new Date().toISOString().slice(0,10);
  $('gridDate').value=today; $('manDate').value=today;
  await refreshStats(); await refreshVersions();
}
async function refreshStats(){
  const s=await j(API+'/stats');
  $('kpis').innerHTML=Object.entries(s).map(([k,v])=>`<div class="kpi"><div class="v">${v}</div><div class="l">${k}</div></div>`).join('');
}
async function refreshVersions(){
  state.versions=await j(API+'/versions');
  const opts=state.versions.map(v=>`<option value="${v.id}">v${v.version_no} · ${v.status} · ${v.created_at.slice(0,10)}</option>`).join('');
  $('gridVersion').innerHTML=opts; $('manVersion').innerHTML=opts;
}
async function ingestAll(){
  $('ingestMsg').textContent='Ingesting (recursive)…';
  try{
    const r=await j(API+'/ingest/all',{method:'POST'});
    const ok=r.ingested?r.ingested.filter(x=>!x.error).length:0;
    const err=r.ingested?r.ingested.filter(x=>x.error).length:0;
    $('ingestMsg').innerHTML=`<span class="ok">✓ ${ok} files ingested</span>${err?' · <span class="no">'+err+' errors</span>':''} · scanned: ${r.scanned_root||'?'}`;
    console.log(r);
    await refreshStats();
  }catch(e){$('ingestMsg').textContent='Error: '+e.message;}
}
async function generate(){
  $('genMsg').textContent='Generating…';
  try{
    const body={grade:$('genGrade').value||null, seed:parseInt($('genSeed').value)||42};
    const r=await j(API+'/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    $('genMsg').textContent=`Done: v${r.version_id} · ${r.stats.slots_inserted} slots · warnings: ${r.stats.warnings.length}`;
    await refreshStats(); await refreshVersions();
  }catch(e){$('genMsg').textContent='Error: '+e.message;}
}
async function loadWeek(){
  const vid=$('gridVersion').value, sid=$('gridSection').value, d=$('gridDate').value;
  if(!vid||!sid||!d){$('weekGrid').innerHTML='<div class="banner warn">Generate a version first</div>';return;}
  const rows=await j(`${API}/versions/${vid}/sections/${sid}/week/${d}`);
  const start=new Date(d); start.setDate(start.getDate()-start.getDay()+1);
  const days=[]; for(let i=0;i<6;i++){const x=new Date(start);x.setDate(x.getDate()+i);days.push(x);}
  const byTime={};
  rows.forEach(r=>{const k=r.start_time+'-'+r.end_time; (byTime[k]=byTime[k]||{})[r.slot_date]=r;});
  let html='<div class="grid"><div class="hdr time">Time</div>'+days.map(d=>`<div class="hdr">${d.toDateString().slice(0,10)}</div>`).join('');
  const times=Object.keys(byTime).sort();
  if(!times.length) html+='<div style="grid-column:1/-1;padding:20px;text-align:center;color:#64748b">No slots</div>';
  for(const t of times){
    html+=`<div class="hdr time">${t}</div>`;
    for(const d of days){
      const iso=d.toISOString().slice(0,10);
      const s=byTime[t][iso];
      if(!s){html+='<div class="slot"></div>';continue;}
      let cls='slot '+(s.slot_type||'').toLowerCase()+(s.source==='MANUAL'?' manual':'');
      const label=s.slot_type==='EXAM'?'EXAM':s.slot_type==='REVISION'?'REVISION':(s.subject_name||'Free');
      html+=`<div class="${cls}"><div class="s">${label}</div><div class="t">${s.teacher_name||''}</div><div class="t">${(s.topic||'').slice(0,26)}</div></div>`;
    }
  }
  html+='</div>';
  $('weekGrid').innerHTML=html;
}
function downloadWeekExcel(){
  const vid=$('gridVersion').value, sid=$('gridSection').value, d=$('gridDate').value;
  if(!vid||!sid||!d){ $('dlMsg').innerHTML='<span class="no">Generate a version first</span>'; return; }
  $('dlMsg').textContent='Preparing Excel…';
  const url=`${API}/export/week.xlsx?version_id=${vid}&section_id=${sid}&date=${d}`;
  const a=document.createElement('a');
  a.href=url;
  a.download='';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(()=>{ $('dlMsg').innerHTML='<span class="ok">✓ Download started</span>'; }, 600);
}
async function loadSlots(){
  const vid=$('manVersion').value, sid=$('manSection').value, d=$('manDate').value;
  if(!vid){$('slotList').innerHTML='<div class="banner warn">Generate first</div>';return;}
  const rows=await j(`${API}/versions/${vid}/sections/${sid}/day/${d}`);
  if(!rows.length){$('slotList').innerHTML='<div class="banner warn">No slots for this date</div>';return;}
  $('slotList').innerHTML='<table><tr><th>Time</th><th>Type</th><th>Subject</th><th>Teacher</th><th>Topic</th><th></th></tr>'+
    rows.map(r=>{
      // Show Edit button for both TEACHING and FREE slots
      const canEdit = r.slot_type === 'TEACHING' || r.slot_type === 'FREE';
      const btn = canEdit ? `<button onclick='startEdit(${JSON.stringify(r)})'>Edit</button>` : '';
      return `<tr><td>${r.start_time}–${r.end_time}</td><td>${r.slot_type}</td><td>${r.subject_name||''}</td><td>${r.teacher_name||''}</td><td>${r.topic||''}</td><td>${btn}</td></tr>`;
    }).join('')+'</table>';
}
async function startEdit(slot){
  state.editingSlot=slot;
  $('editPanel').classList.remove('hidden');
  // If it's a FREE slot, we might not have subject/teacher data
  if(slot.subject_id) $('edSubject').value=slot.subject_id;
  if(slot.teacher_id) $('edTeacher').value=slot.teacher_id;
  else $('edTeacher').value = ""; // reset teacher if free
  await loadTopics();
  if(slot.syllabus_topic_id) $('edTopic').value=slot.syllabus_topic_id;
  $('edReason').value='';
  await checkValid();
}
async function loadTopics(){
  const sid=$('edSubject').value;
  const sub=state.subjects.find(s=>s.id==sid);
  if(!sub) return;
  const sec=state.sections.find(s=>s.id==$('manSection').value);
  const grade=sec?sec.gname:'8';
  try{
    const topics=await j(`${API}/syllabus/topics?grade=${encodeURIComponent(grade)}&subject_code=${encodeURIComponent(sub.code)}`);
    $('edTopic').innerHTML=topics.map(t=>`<option value="${t.id}">${t.chapter_no} · ${t.sub_topic_no||''} ${t.sub_topic_name||''}</option>`).join('')||'<option value="">(no topics)</option>';
  }catch(e){$('edTopic').innerHTML='<option value="">(no topics)</option>';}
}
async function checkValid(){
  const slot=state.editingSlot; if(!slot) return;
  try{
    const vid=$('manVersion').value;
    const sub=state.subjects.find(s=>s.id==$('edSubject').value);
    const eligible=await j(`${API}/teachers/eligible?grade=${encodeURIComponent(state.sections.find(s=>s.id==$('manSection').value).gname)}&subject_code=${encodeURIComponent(sub.code)}&date=${slot.slot_date}&start=${slot.start_time}&end=${slot.end_time}&version_id=${vid}`);
    $('edTeacher').innerHTML=eligible.map(t=>`<option value="${t.id}" ${!t.eligible?'disabled':''}>${t.name}${t.eligible?' ✓':' ✗ '+t.reason}</option>`).join('');
    if(slot.teacher_id) $('edTeacher').value=slot.teacher_id;
    $('edValid').innerHTML='<span class="ok">✓ Eligible teachers marked</span>';
  }catch(e){$('edValid').innerHTML='<span class="no">'+e.message+'</span>';}
}
async function saveOverride(){
  const slot=state.editingSlot; if(!slot) return;
  const body={subject_id:parseInt($('edSubject').value),
              teacher_id:parseInt($('edTeacher').value)||null,
              syllabus_topic_id:parseInt($('edTopic').value)||null,
              reason:$('edReason').value||'manual override'};
  try{ await j(`${API}/slots/${slot.id}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    alert('Saved'); closeEdit(); loadSlots();
  }catch(e){alert('Error: '+e.message);}
}
function closeEdit(){ $('editPanel').classList.add('hidden'); state.editingSlot=null; }
$('edSubject').addEventListener('change',loadTopics);
async function refreshConflicts(){
  const vid=$('gridVersion').value||($('manVersion').value);
  if(!vid){ $('conflicts').innerHTML='<div class="banner warn">No version selected</div>'; return; }
  const rows=await j(`${API}/conflicts?version_id=${vid}`);
  $('conflicts').innerHTML=rows.length?
    rows.slice(0,50).map(r=>`<div class="banner ${r.severity==='BLOCKING'?'err':'warn'}">[${r.rule}] ${JSON.stringify(r.detail).slice(0,120)}</div>`).join('')
    :'<div class="banner ok">✓ No conflicts detected</div>';
}
setInterval(()=>{ if(!document.getElementById('dash').classList.contains('hidden')) refreshConflicts(); }, 5000);
boot();
</script></body></html>
"""

@app.get("/", response_class=HTMLResponse)
def index(): return HTML

if __name__ == "__main__":
    import uvicorn
    os.makedirs(DATA_DIR, exist_ok=True)
    print("="*70)
    print(" Test School Timetable Planner (v1.1 — recursive data/ scan)")
    print(" Drop .xlsx files into ./data/ (sub-folders OK) then click 'Ingest'")
    print(" Open http://127.0.0.1:8000")
    print("="*70)
    uvicorn.run(app, host="127.0.0.1", port=8000)