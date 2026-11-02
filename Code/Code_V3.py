"""
================================================================================
 St. John's NEXUS School — TIME TABLE PLANNER (Single-File)
================================================================================
 RUN: python Code_V3.py
 Open: http://127.0.0.1:8000
================================================================================
"""

import os, re, json, sqlite3, datetime as dt, calendar, random
from contextlib import contextmanager
from typing import Optional
from fastapi import FastAPI, HTTPException, UploadFile, File, Body, Depends
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from passlib.context import CryptContext
from jose import jwt, JWTError
from pydantic import BaseModel

# ------------------------------------------------------------------ CONFIG
DB_PATH = "timetable.db"
DATA_DIR = "data"
ACADEMIC_YEAR = "2026-27"
YEAR_START = dt.date(2026, 10, 1)
YEAR_END   = dt.date(2027, 4, 30)

SECRET_KEY = "nexus-super-secret-key-123"
ALGORITHM = "HS256"
DEFAULT_PWD = "Password123"
pwd_ctx = CryptContext(schemes=["argon2"], deprecated="auto")

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
CREATE TABLE IF NOT EXISTS academic_year(id INTEGER PRIMARY KEY, name TEXT, start_date TEXT, end_date TEXT, status TEXT DEFAULT 'active');
CREATE TABLE IF NOT EXISTS grade(id INTEGER PRIMARY KEY, name TEXT UNIQUE, level_group TEXT, day_start TEXT, day_end TEXT);
CREATE TABLE IF NOT EXISTS section(id INTEGER PRIMARY KEY, grade_id INTEGER REFERENCES grade(id), label TEXT, strength INTEGER, UNIQUE(grade_id,label));
CREATE TABLE IF NOT EXISTS break_period(id INTEGER PRIMARY KEY, label TEXT, start_time TEXT, end_time TEXT, applies_to TEXT);
CREATE TABLE IF NOT EXISTS teacher(id INTEGER PRIMARY KEY, name TEXT, employment_status TEXT, shift_start TEXT, shift_end TEXT, is_flexible INTEGER DEFAULT 0, max_periods_per_day INTEGER DEFAULT 6, max_periods_per_week INTEGER DEFAULT 30);
CREATE TABLE IF NOT EXISTS subject(id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT, category TEXT);
CREATE TABLE IF NOT EXISTS teacher_capability(teacher_id INTEGER REFERENCES teacher(id), subject_id INTEGER REFERENCES subject(id), proficiency_rank INTEGER DEFAULT 1, PRIMARY KEY(teacher_id, subject_id));
CREATE TABLE IF NOT EXISTS syllabus_document(id INTEGER PRIMARY KEY, grade_id INTEGER, subject_id INTEGER, semester INTEGER, source_file TEXT, ingested_at TEXT);
CREATE TABLE IF NOT EXISTS syllabus_chapter(id INTEGER PRIMARY KEY, document_id INTEGER REFERENCES syllabus_document(id), chapter_no TEXT, chapter_name TEXT, target_month TEXT, page_range TEXT);
CREATE TABLE IF NOT EXISTS syllabus_topic(id INTEGER PRIMARY KEY, chapter_id INTEGER REFERENCES syllabus_chapter(id), sub_topic_no TEXT, sub_topic_name TEXT, learning_focus TEXT, target_month_start TEXT, target_month_end TEXT, estimated_periods INTEGER DEFAULT 1, granularity TEXT DEFAULT 'SUB_TOPIC');
CREATE TABLE IF NOT EXISTS timetable_version(id INTEGER PRIMARY KEY AUTOINCREMENT, academic_year_id INTEGER, version_no INTEGER, generated_by TEXT, status TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS schedule_slot(id INTEGER PRIMARY KEY AUTOINCREMENT, version_id INTEGER, section_id INTEGER, slot_date TEXT, start_time TEXT, end_time TEXT, slot_type TEXT, subject_id INTEGER, teacher_id INTEGER, syllabus_topic_id INTEGER, source TEXT DEFAULT 'AUTO', override_reason TEXT);
CREATE INDEX IF NOT EXISTS idx_slot_lookup ON schedule_slot(version_id, slot_date, section_id);
CREATE INDEX IF NOT EXISTS idx_slot_teacher ON schedule_slot(version_id, slot_date, teacher_id);
CREATE TABLE IF NOT EXISTS exam_event(id INTEGER PRIMARY KEY AUTOINCREMENT, grade_id INTEGER, assessment_code TEXT, subject_id INTEGER, revision_start TEXT, revision_end TEXT, exam_date TEXT, session TEXT, start_time TEXT, end_time TEXT);
CREATE TABLE IF NOT EXISTS generation_run(id INTEGER PRIMARY KEY AUTOINCREMENT, version_id INTEGER, status TEXT, seed INTEGER, solver_stats TEXT, conflict_report TEXT, started_at TEXT, finished_at TEXT);
CREATE TABLE IF NOT EXISTS conflict_log(id INTEGER PRIMARY KEY AUTOINCREMENT, version_id INTEGER, severity TEXT, rule_code TEXT, detail TEXT, resolved INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS audit_log(id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, slot_id INTEGER, action TEXT, before TEXT, after TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS app_user(id INTEGER PRIMARY KEY AUTOINCREMENT, login_id TEXT UNIQUE, password_hash TEXT, role TEXT, teacher_id INTEGER);
"""

def init_db():
    with db() as c:
        c.executescript(SCHEMA)
        c.execute("INSERT OR IGNORE INTO academic_year(id,name,start_date,end_date,status) VALUES(1,?,?,?,'active')", (ACADEMIC_YEAR, YEAR_START.isoformat(), YEAR_END.isoformat()))
        grades = [("EYP1","PRE_PRIMARY","08:40","12:30"),("EYP2","PRE_PRIMARY","08:40","12:30"),("EYP3","PRE_PRIMARY","08:40","12:30")] + [(str(i),"MIDDLE" if i<=8 else "SECONDARY","09:00","15:45") for i in range(1,10)] + [("10","BOARD","09:00","17:30")]
        for g in grades: c.execute("INSERT OR IGNORE INTO grade(name,level_group,day_start,day_end) VALUES(?,?,?,?)", g)
        for row in c.execute("SELECT id,name FROM grade").fetchall():
            n_sec = 4 if (row["name"] in ("EYP1","EYP2","EYP3") or row["name"].isdigit() and int(row["name"])<=7) else 2
            strength = 35 if n_sec==4 else 40
            sections_list = ["EVO", "TERRA", "VIA", "NEO"]
            for lbl in sections_list[:n_sec]: 
                c.execute("INSERT OR IGNORE INTO section(grade_id,label,strength) VALUES(?,?,?)", (row["id"], lbl, strength))
        for b in [("First Break","10:20","10:30"),("Lunch Break","12:25","12:55"),("Third Break","14:15","14:25")]: c.execute("INSERT OR IGNORE INTO break_period(label,start_time,end_time,applies_to) VALUES(?,?,?,'ALL')", b)
        subs = [("TELUGU","Telugu","LANGUAGE"),("ENGLISH","English","LANGUAGE"),("HINDI","Hindi","LANGUAGE"),("MATHS","Mathematics","CORE"),("PHYSICS","Physics","CORE"),("BIOLOGY","Biology","CORE"),("SOCIAL","Social","CORE"),("ICT","ICT","ACTIVITY"),("PET","PET","ACTIVITY"),("TAEKWONDO","Taekwondo","ACTIVITY"),("LIBRARY","Library","ACTIVITY"),("DRAWING","Drawing","ACTIVITY"),("DANCE","Dance","ACTIVITY")]
        for s in subs: c.execute("INSERT OR IGNORE INTO subject(code,name,category) VALUES(?,?,?)", s)
        
        if not c.execute("SELECT 1 FROM app_user WHERE login_id='ADMIN'").fetchone():
            c.execute("INSERT INTO app_user(login_id, password_hash, role) VALUES(?,?,?)", ("ADMIN", pwd_ctx.hash(DEFAULT_PWD), "admin"))

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
    m1 = month_num(parts[0]); m2 = month_num(parts[-1]) if len(parts) > 1 else m1
    if not m1 or not m2: return None, None
    def date_for(m, first):
        y = year_start.year
        if m < year_start.month: y += 1
        last = calendar.monthrange(y, m)[1]
        return dt.date(y, m, 1 if first else last)
    start = date_for(m1, True); end = date_for(m2, False)
    if end < start: end = date_for(m2, False)
    return start, end

# ---------------------------------------------------------- EXCEL INGESTION
SUBJECT_HINTS = {"maths": "MATHS", "mat": "MATHS", "mathematics": "MATHS", "social": "SOCIAL", "s.s": "SOCIAL", "social science": "SOCIAL", "science": "SCIENCE", "physics": "PHYSICS", "biology": "BIOLOGY", "telugu": "TELUGU", "english": "ENGLISH", "hindi": "HINDI", "languages": "LANGUAGES", "ict": "ICT", "i.c.t": "ICT", "computer": "ICT", "computers": "ICT", "pet": "PET", "p.e.t": "PET", "physical education": "PET", "sports": "PET", "taekwondo": "TAEKWONDO", "library": "LIBRARY", "drawing": "DRAWING", "art": "DRAWING", "dance": "DANCE", "music": "DANCE"}
def guess_subject(filename):
    fn = filename.lower()
    for k, v in SUBJECT_HINTS.items():
        if k in fn: return v
    return "UNKNOWN"
def guess_grade(filename):
    m = re.search(r"class[_\s]*(\d+)", filename, re.I)
    return m.group(1) if m else None

def parse_syllabus_sheet(ws):
    rows = []; header = None; header_row_idx = 0
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
    idx = {k: col(v) for k, v in {"chapter_no":"chapter no","chapter_name":"chapter name","target_month":"target month","page_range":"page","sub_topic_no":"sub-topic no","sub_topic_name":"sub-topic name","learning_focus":"learning focus"}.items()}
    for row in ws.iter_rows(min_row=header_row_idx+2, values_only=True):
        if not any(row): continue
        def get(k):
            j = idx[k]
            return str(row[j]).strip() if j is not None and j < len(row) and row[j] is not None else ""
        ch_no = get("chapter_no") or get("sub_topic_no"); ch_name = get("chapter_name")
        if not ch_no and not ch_name: continue
        rows.append({"chapter_no": ch_no, "chapter_name": ch_name, "target_month": get("target_month"), "page_range": get("page_range"), "sub_topic_no": get("sub_topic_no"), "sub_topic_name": get("sub_topic_name"), "learning_focus": get("learning_focus")})
    return rows

def ingest_syllabus_file(path):
    filename = os.path.basename(path); grade_name = guess_grade(filename); subject_code = guess_subject(filename)
    if not grade_name: return {"file": filename, "skipped": "no grade in filename"}
    wb = openpyxl.load_workbook(path, data_only=True); inserted = 0; replaced_docs = 0
    for si, sheet_name in enumerate(wb.sheetnames):
        ws = wb[sheet_name]; sem = 2 if re.search(r"sem[\s_-]*2|semester[\s_-]*2", sheet_name, re.I) else 1
        rows = parse_syllabus_sheet(ws)
        if not rows: continue
        with db() as c:
            g = c.execute("SELECT id FROM grade WHERE name=?", (grade_name,)).fetchone()
            s = c.execute("SELECT id FROM subject WHERE code=?", (subject_code,)).fetchone()
            if not g: return {"file": filename, "error": f"grade {grade_name} not found"}
            grade_id = g["id"]; subject_id = s["id"] if s else None
            if not subject_id:
                c.execute("INSERT OR IGNORE INTO subject(code,name,category) VALUES(?,?,?)", (subject_code, subject_code.title(), "CORE"))
                subject_id = c.execute("SELECT id FROM subject WHERE code=?", (subject_code,)).fetchone()["id"]
            old = c.execute("SELECT id FROM syllabus_document WHERE grade_id=? AND subject_id=? AND semester=?", (grade_id, subject_id, sem)).fetchall()
            for od in old:
                c.execute("DELETE FROM syllabus_topic WHERE chapter_id IN (SELECT id FROM syllabus_chapter WHERE document_id=?)", (od["id"],))
                c.execute("DELETE FROM syllabus_chapter WHERE document_id=?", (od["id"],))
                c.execute("DELETE FROM syllabus_document WHERE id=?", (od["id"],))
                replaced_docs += 1
            cur = c.execute("INSERT INTO syllabus_document(grade_id,subject_id,semester,source_file,ingested_at) VALUES(?,?,?,?,?)", (grade_id, subject_id, sem, filename, dt.datetime.now().isoformat()))
            doc_id = cur.lastrowid; chapter_cache = {}
            for r in rows:
                key = (r["chapter_no"], r["chapter_name"])
                if key not in chapter_cache:
                    cur = c.execute("INSERT INTO syllabus_chapter(document_id,chapter_no,chapter_name,target_month,page_range) VALUES(?,?,?,?,?)", (doc_id, r["chapter_no"], r["chapter_name"], r["target_month"], r["page_range"]))
                    chapter_cache[key] = cur.lastrowid
                current_chapter_id = chapter_cache[key]
                ms, me = parse_target_month(r["target_month"])
                c.execute("INSERT INTO syllabus_topic(chapter_id,sub_topic_no,sub_topic_name,learning_focus,target_month_start,target_month_end,estimated_periods) VALUES(?,?,?,?,?,?,1)", (current_chapter_id, r["sub_topic_no"], r["sub_topic_name"], r["learning_focus"], ms.isoformat() if ms else None, me.isoformat() if me else None))
                inserted += 1
    return {"file": filename, "topics_inserted": inserted, "documents_replaced": replaced_docs}

def ingest_teachers_file(path):
    wb = openpyxl.load_workbook(path, data_only=True); ws = wb.active; inserted = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or not row or not row[0]: continue
        name = str(row[0]).strip(); cap = str(row[1]).strip() if row[1] else ""
        status = str(row[2]).strip().upper() if row[2] else "FULL_TIME"; shift = str(row[3]).strip() if row[3] else ""
        if "flexible" in shift.lower() or "part" in status.lower():
            s_start, s_end, flex = None, None, 1
        else:
            m = re.match(r"(\d{1,2}):(\d{2})\s*(AM|PM)", shift, re.I); m2 = re.search(r"to\s+(\d{1,2}):(\d{2})\s*(AM|PM)", shift, re.I)
            def to24(m):
                if not m: return None
                h, mi, ap = int(m.group(1)), m.group(2), m.group(3).upper()
                if ap == "PM" and h != 12: h += 12
                if ap == "AM" and h == 12: h = 0
                return f"{h:02d}:{mi}"
            s_start = to24(m); s_end = to24(m2); flex = 0
        with db() as c:
            cap_lower = cap.lower()
            if "dance" in cap_lower: subj_code = "DANCE"
            elif "pet" in cap_lower or "physical" in cap_lower: subj_code = "PET"
            elif "ict" in cap_lower or "computer" in cap_lower: subj_code = "ICT"
            elif "taekwondo" in cap_lower: subj_code = "TAEKWONDO"
            elif "library" in cap_lower: subj_code = "LIBRARY"
            elif "drawing" in cap_lower or "art" in cap_lower: subj_code = "DRAWING"
            elif "telugu" in cap_lower: subj_code = "TELUGU"
            elif "english" in cap_lower: subj_code = "ENGLISH"
            elif "hindi" in cap_lower: subj_code = "HINDI"
            elif "math" in cap_lower: subj_code = "MATHS"
            elif "social" in cap_lower: subj_code = "SOCIAL"
            elif "physics" in cap_lower: subj_code = "PHYSICS"
            elif "biology" in cap_lower: subj_code = "BIOLOGY"
            else: subj_code = cap.upper().replace(" ", "_")
            subj = c.execute("SELECT id FROM subject WHERE UPPER(TRIM(code))=?", (subj_code,)).fetchone()
            if not subj:
                c.execute("INSERT INTO subject(code,name,category) VALUES(?,?,?)", (subj_code, cap, "ACTIVITY"))
                subj = c.execute("SELECT id FROM subject WHERE code=?", (subj_code,)).fetchone()
            existing = c.execute("SELECT id FROM teacher WHERE name=?", (name,)).fetchone()
            if existing:
                tid = existing["id"]
                c.execute("UPDATE teacher SET employment_status=?, shift_start=?, shift_end=?, is_flexible=? WHERE id=?", (status, s_start, s_end, flex, tid))
            else:
                cur = c.execute("INSERT INTO teacher(name,employment_status,shift_start,shift_end,is_flexible) VALUES(?,?,?,?,?)", (name, status, s_start, s_end, flex))
                tid = cur.lastrowid
            c.execute("INSERT OR IGNORE INTO teacher_capability(teacher_id,subject_id) VALUES(?,?)", (tid, subj["id"]))
            if not c.execute("SELECT 1 FROM app_user WHERE login_id=?", (name,)).fetchone():
                c.execute("INSERT INTO app_user(login_id, password_hash, role, teacher_id) VALUES(?,?,?,?)", (name, pwd_ctx.hash(DEFAULT_PWD), "teacher", tid))
            inserted += 1
    return {"file": os.path.basename(path), "teachers_inserted": inserted}

def ingest_exams_file(path):
    wb = openpyxl.load_workbook(path, data_only=True); inserted = 0
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]; grade_name = None
        m = re.search(r"class\s*(\d+)", sheet_name, re.I)
        if m: grade_name = m.group(1)
        rows = list(ws.iter_rows(values_only=True))
        for i, row in enumerate(rows):
            cells = [str(c).strip() if c is not None else "" for c in row]; line = " ".join(cells).lower()
            if not grade_name: continue
            if re.search(r"class\s*\d+\s*-\s*(fa-\d|summative|sa-\d)", line):
                hdr_idx = None
                for j in range(i, min(i+3, len(rows))):
                    h = " ".join(str(x).lower() for x in rows[j] if x)
                    if "assessment" in h and "exam date" in h: hdr_idx = j; break
                if hdr_idx is None: continue
                am = re.search(r"(fa-\d|sa-\d|summative assessment-\d)", line); assessment = am.group(1).upper() if am else "EXAM"
                for j in range(hdr_idx+1, len(rows)):
                    r = rows[j]
                    if not r or not r[0]:
                        if j > hdr_idx+1 and not any(r): break
                        continue
                    vals = [str(x).strip() if x is not None else "" for x in r]
                    if "assessment" in vals[0].lower(): break
                    if len(vals) < 5: continue
                    dm = re.search(r"(\d{2})/(\d{2})/(\d{2})", vals[2])
                    if not dm: continue
                    exam_date = dt.date(2000+int(dm.group(3)), int(dm.group(2)), int(dm.group(1))); rev_start = rev_end = None
                    rm = re.search(r"(\d{2})/(\d{2})/(\d{2})\s*to\s*(\d{2})/(\d{2})/(\d{2})", vals[1])
                    if rm:
                        rev_start = dt.date(2000+int(rm.group(3)), int(rm.group(2)), int(rm.group(1))); rev_end = dt.date(2000+int(rm.group(6)), int(rm.group(5)), int(rm.group(4)))
                    session = "FN" if "f.n" in vals[3].lower() else "AN" if "a.n" in vals[3].lower() else "FULL"
                    tm = re.search(r"(\d{1,2})[.:](\d{2})\s*(a\.?m|p\.?m)", vals[3], re.I); start_time = end_time = None
                    if tm:
                        h = int(tm.group(1)); mi = tm.group(2); ap = tm.group(3).lower()
                        if "p" in ap and h != 12: h += 12
                        start_time = f"{h:02d}:{mi}"
                        em = re.search(r"(\d{1,2})[.:](\d{2})\s*(a\.?m|p\.?m)", vals[3][tm.end():], re.I)
                        if em:
                            h2 = int(em.group(1)); mi2 = em.group(2); ap2 = em.group(3).lower()
                            if "p" in ap2 and h2 != 12: h2 += 12
                            end_time = f"{h2:02d}:{mi2}"
                    with db() as c:
                        g = c.execute("SELECT id FROM grade WHERE name=?", (grade_name,)).fetchone()
                        if not g: continue
                        subj_code = None
                        for k, v in SUBJECT_HINTS.items():
                            if k in vals[4].lower(): subj_code = v; break
                        sid = None
                        if subj_code:
                            s = c.execute("SELECT id FROM subject WHERE code=?", (subj_code,)).fetchone(); sid = s["id"] if s else None
                        c.execute("INSERT INTO exam_event(grade_id,assessment_code,subject_id,revision_start,revision_end,exam_date,session,start_time,end_time) VALUES(?,?,?,?,?,?,?,?,?)", (g["id"], assessment, sid, rev_start.isoformat() if rev_start else None, rev_end.isoformat() if rev_end else None, exam_date.isoformat(), session, start_time, end_time))
                        inserted += 1
    return {"file": os.path.basename(path), "exam_events_inserted": inserted}

def ingest_all(data_dir=DATA_DIR):
    results = []
    if not os.path.isdir(data_dir): return {"error": f"{data_dir}/ not found"}
    with db() as c:
        c.execute("DELETE FROM teacher_capability"); c.execute("DELETE FROM teacher")
        c.execute("DELETE FROM syllabus_topic"); c.execute("DELETE FROM syllabus_chapter")
        c.execute("DELETE FROM syllabus_document"); c.execute("DELETE FROM exam_event")
    for root, _dirs, files in os.walk(data_dir):
        for fn in sorted(files):
            if not fn.lower().endswith((".xlsx", ".xlsm")): continue
            path = os.path.join(root, fn); rel = os.path.relpath(path, data_dir)
            try:
                low = fn.lower()
                if "teacher" in low: r = ingest_teachers_file(path)
                elif "exam" in low or "time table" in low: r = ingest_exams_file(path)
                else: r = ingest_syllabus_file(path)
                r["source"] = rel; results.append(r)
            except Exception as e: results.append({"file": rel, "error": str(e)})
    return {"ingested": results, "total_files": len(results), "scanned_root": os.path.abspath(data_dir)}

# ---------------------------------------------------------------- SCHEDULING
def build_day_slot_template(grade_row, breaks):
    def t2m(t):
        h, m = t.split(":"); return int(h)*60 + int(m)
    def m2t(x): return f"{x//60:02d}:{x%60:02d}"
    start = t2m(grade_row["day_start"]); end = t2m(grade_row["day_end"]); bset = sorted([(t2m(b["start_time"]), t2m(b["end_time"])) for b in breaks])
    slots = []; cur = start
    for bs, be in bset:
        if bs < start or be > end: continue
        if cur < bs:
            length = bs - cur
            if length <= 0: continue
            n = max(1, round(length / 45)); per = length // n
            for k in range(n): slots.append((m2t(cur + k*per), m2t(cur + (k+1)*per)))
        cur = max(cur, be)
    if cur < end:
        length = end - cur
        if length > 0:
            n = max(1, round(length / 45)); per = length // n
            for k in range(n): slots.append((m2t(cur + k*per), m2t(cur + (k+1)*per)))
    return slots

def get_teachers_for_subject(c, subject_id):
    return c.execute("SELECT t.* FROM teacher t JOIN teacher_capability tc ON tc.teacher_id=t.id WHERE tc.subject_id=? ORDER BY tc.proficiency_rank, t.id", (subject_id,)).fetchall()

def teacher_free(c, version_id, date, start, end, teacher_id):
    return c.execute("SELECT 1 FROM schedule_slot WHERE version_id=? AND slot_date=? AND teacher_id=? AND slot_type='TEACHING' AND NOT (end_time<=? OR start_time>=?)", (version_id, date, teacher_id, start, end)).fetchone() is None

def section_free(c, version_id, date, start, end, section_id):
    return c.execute("SELECT 1 FROM schedule_slot WHERE version_id=? AND slot_date=? AND section_id=? AND NOT (end_time<=? OR start_time>=?)", (version_id, date, section_id, start, end)).fetchone() is None

def get_exam_block(c, grade_id, date):
    r = c.execute("SELECT * FROM exam_event WHERE grade_id=? AND exam_date=?", (grade_id, date)).fetchall()
    if r: return {"type":"EXAM","events":[dict(x) for x in r]}
    r = c.execute("SELECT * FROM exam_event WHERE grade_id=? AND revision_start<=? AND revision_end>=?", (grade_id, date, date)).fetchall()
    if r: return {"type":"REVISION","events":[dict(x) for x in r]}
    return None

def generate_timetable(sections=None, grade_filter=None, seed=42, version_label="AUTO"):
    random.seed(seed); today = dt.date.today(); effective_start = max(YEAR_START, today)
    if effective_start > YEAR_END: raise HTTPException(400, "No schedulable days")
    with db() as c:
        maxv = c.execute("SELECT COALESCE(MAX(version_no),0) v FROM timetable_version").fetchone()["v"]
        cur = c.execute("INSERT INTO timetable_version(academic_year_id,version_no,generated_by,status,created_at) VALUES(1,?,?,?,?)", (maxv+1, version_label, "draft", dt.datetime.now().isoformat()))
        version_id = cur.lastrowid
        run = c.execute("INSERT INTO generation_run(version_id,status,seed,started_at) VALUES(?,?,?,?)", (version_id, "running", seed, dt.datetime.now().isoformat()))
        run_id = run.lastrowid
    stats = {"sections": 0, "slots_inserted": 0, "warnings": []}
    try:
        with db() as c:
            q = "SELECT s.*, g.name gname, g.level_group, g.day_start, g.day_end FROM section s JOIN grade g ON g.id=s.grade_id"
            params = []
            if grade_filter: q += " WHERE g.name=?"; params.append(grade_filter)
            if sections: q += " AND s.id IN (%s)" % ",".join("?" * len(sections)); params.extend(sections)
            sections_rows = c.execute(q, params).fetchall()
            breaks = [dict(r) for r in c.execute("SELECT * FROM break_period").fetchall()]
            days = [effective_start + dt.timedelta(days=i) for i in range((YEAR_END - effective_start).days + 1)]
            for srow in sections_rows:
                stats["sections"] += 1
                subjects_for_grade = [dict(r) for r in c.execute("SELECT DISTINCT sd.subject_id, sub.code, sub.name FROM syllabus_document sd JOIN subject sub ON sub.id=sd.subject_id WHERE sd.grade_id=?", (srow["grade_id"],)).fetchall()]
                if not subjects_for_grade: continue
                exams = c.execute("SELECT exam_date FROM exam_event WHERE grade_id=?", (srow["grade_id"],)).fetchall()
                exam_dates = [dt.date.fromisoformat(e["exam_date"]) for e in exams if e["exam_date"]]
                first_exam = min(exam_dates) if exam_dates else dt.date(2027, 3, 1)
                queues = {}
                for sub in subjects_for_grade:
                    topics = [dict(r) for r in c.execute("SELECT st.*, sc.chapter_no, sc.chapter_name FROM syllabus_topic st JOIN syllabus_chapter sc ON sc.id=st.chapter_id JOIN syllabus_document sd ON sd.id=sc.document_id WHERE sd.grade_id=? AND sd.subject_id=? ORDER BY COALESCE(st.target_month_start,'9999'), st.id", (srow["grade_id"], sub["subject_id"])).fetchall()]
                    if not topics: continue
                    teachers = get_teachers_for_subject(c, sub["subject_id"])
                    if not teachers: continue
                    queues[sub["subject_id"]] = {"topics": topics, "teachers": teachers, "name": sub["name"], "first_exam": first_exam}
                if not queues: continue
                tmpl = build_day_slot_template(srow, breaks)
                for d in days:
                    if d.weekday() >= 6: continue
                    block = get_exam_block(c, srow["grade_id"], d)
                    if block and block["type"] == "EXAM":
                        for (st, en) in tmpl:
                            c.execute("INSERT INTO schedule_slot(version_id,section_id,slot_date,start_time,end_time,slot_type,source) VALUES(?,?,?,?,?,'EXAM','AUTO')", (version_id, srow["id"], d.isoformat(), st, en))
                            stats["slots_inserted"] += 1
                        continue
                    revision_subjects = set()
                    if block and block["type"] == "REVISION":
                        for ev in block["events"]:
                            if ev["subject_id"]: revision_subjects.add(ev["subject_id"])
                    for (st, en) in tmpl:
                        if not any(q["topics"] for q in queues.values()):
                            c.execute("INSERT INTO schedule_slot(version_id,section_id,slot_date,start_time,end_time,slot_type,source) VALUES(?,?,?,?,?,'FREE','AUTO')", (version_id, srow["id"], d.isoformat(), st, en))
                            stats["slots_inserted"] += 1
                            continue
                        candidates = []
                        wk_start = (d - dt.timedelta(days=d.weekday())).isoformat()
                        wk_end = (d + dt.timedelta(days=6 - d.weekday())).isoformat()
                        
                        for subj_id, q in queues.items():
                            if not q["topics"]: continue
                            
                            # 1. Calculate how many periods this subject already has this week
                            c_count = c.execute("SELECT COUNT(*) n FROM schedule_slot WHERE version_id=? AND section_id=? AND subject_id=? AND slot_date BETWEEN ? AND ? AND slot_type='TEACHING'", 
                                                (version_id, srow["id"], subj_id, wk_start, wk_end)).fetchone()["n"]
                            
                            subj_name = str(q.get("name") or "").lower()
                            is_act_check = any(act in subj_name for act in ["pet", "ict", "drawing", "taekwondo", "dance", "library", "yoga", "art"])
                            
                            # 2. Determine Priority
                            if c_count < 3:
                                # Priority 0: MUST hit minimum 3 periods per week
                                priority = 0
                            elif is_act_check:
                                # Priority 2: Activities max out at 3
                                priority = 2
                            else:
                                # Priority 1: Core subjects with 3+ slots. Sorted by highest syllabus remaining.
                                priority = 1
                            
                            # 3. Next topic and deadline
                            next_topic = q["topics"][0]
                            deadline = next_topic.get("target_month_end") or "9999"
                            if d > (q["first_exam"] - dt.timedelta(days=14)): 
                                deadline = "URGENT"
                            elif next_topic.get("target_month_start"):
                                target_start = dt.date.fromisoformat(next_topic["target_month_start"])
                                if d >= (target_start - dt.timedelta(days=15)): 
                                    deadline = target_start.isoformat()
                                    
                            candidates.append((priority, len(q["topics"]), deadline, subj_id, next_topic))
                            
                        # Sort by Priority, then most remaining topics (heavy syllabus), then deadline, then random
                        candidates.sort(key=lambda x: (x[0], -x[1], x[2], random.random()))
                        
                        placed = False
                        for priority, remaining_topics, deadline, subj_id, topic in candidates:
                            q = queues[subj_id]
                            
                            # --- RULE 1: NO ACTIVITY IN 1ST PERIOD ---
                            subj_name = str(q.get("name") or "").strip().lower()
                            is_activity = any(act in subj_name for act in ["pet", "ict", "drawing", "taekwondo", "dance", "library", "yoga", "art"])
                            is_ict_or_pet = "ict" in subj_name or "pet" in subj_name
                            
                            if st == "09:00" and is_activity:
                                continue
                            
                            # --- RULE 2: WEEKLY LIMITS ---
                            wk_count = c.execute("SELECT COUNT(*) n FROM schedule_slot WHERE version_id=? AND section_id=? AND subject_id=? AND slot_date BETWEEN ? AND ? AND slot_type='TEACHING'", 
                                                 (version_id, srow["id"], subj_id, wk_start, wk_end)).fetchone()["n"]
                            
                            if is_activity:
                                if wk_count >= 3: 
                                    continue
                            else:
                                if wk_count >= 10: 
                                    continue
                                
                            # --- RULE 3: DAILY LIMITS ---
                            daily_count = c.execute("SELECT COUNT(*) n FROM schedule_slot WHERE version_id=? AND section_id=? AND subject_id=? AND slot_date=? AND slot_type='TEACHING'", 
                                                    (version_id, srow["id"], subj_id, d.isoformat())).fetchone()["n"]
                            if is_activity:
                                if "ict" in subj_name:
                                    if daily_count >= 2: continue
                                else:
                                    if daily_count >= 1: continue
                            else:
                                if daily_count >= 2: continue
                            
                            # --- RULE 4: NO CONSECUTIVE ACTIVITIES (Except ICT) ---
                            if not is_ict_or_pet:
                                prev_slots = c.execute("SELECT end_time, subject_id FROM schedule_slot WHERE version_id=? AND section_id=? AND slot_date=? ORDER BY start_time", (version_id, srow["id"], d.isoformat())).fetchall()
                                if prev_slots:
                                    last_slot = prev_slots[-1]
                                    if last_slot["end_time"] == st and last_slot["subject_id"]:
                                        prev_subj = c.execute("SELECT name FROM subject WHERE id=?", (last_slot["subject_id"],)).fetchone()
                                        if prev_subj:
                                            prev_name = str(prev_subj["name"]).lower()
                                            if any(act in prev_name for act in ["drawing", "taekwondo", "dance", "library", "yoga", "art", "pet"]):
                                                continue
                            
                            chosen_teacher = None
                            for t in q["teachers"]:
                                if not t["is_flexible"]:
                                    if t["shift_start"] and t["shift_start"] > st: continue
                                    if t["shift_end"] and t["shift_end"] < en: continue
                                wk_start = (d - dt.timedelta(days=d.weekday())).isoformat()
                                wk_end = (d + dt.timedelta(days=6 - d.weekday())).isoformat()
                                wk = c.execute("SELECT COUNT(*) n FROM schedule_slot WHERE version_id=? AND teacher_id=? AND slot_date BETWEEN ? AND ? AND slot_type='TEACHING'", (version_id, t["id"], wk_start, wk_end)).fetchone()["n"]
                                if wk >= t["max_periods_per_week"]: continue
                                dy = c.execute("SELECT COUNT(*) n FROM schedule_slot WHERE version_id=? AND teacher_id=? AND slot_date=? AND slot_type='TEACHING'", (version_id, t["id"], d.isoformat())).fetchone()["n"]
                                if dy >= t["max_periods_per_day"]: continue
                                if teacher_free(c, version_id, d.isoformat(), st, en, t["id"]): 
                                    chosen_teacher = t
                                    break
                            
                            if not chosen_teacher: 
                                continue
                            if not section_free(c, version_id, d.isoformat(), st, en, srow["id"]): 
                                continue
                            
                            slot_type = "REVISION" if subj_id in revision_subjects else "TEACHING"
                            c.execute("INSERT INTO schedule_slot(version_id,section_id,slot_date,start_time,end_time,slot_type,subject_id,teacher_id,syllabus_topic_id,source) VALUES(?,?,?,?,?,?,?,?,?,'AUTO')",
                                      (version_id, srow["id"], d.isoformat(), st, en, slot_type, subj_id, chosen_teacher["id"], topic["id"]))
                            q["topics"].pop(0)
                            stats["slots_inserted"] += 1
                            placed = True
                            break
                        
                        if not placed:
                            c.execute("INSERT INTO schedule_slot(version_id,section_id,slot_date,start_time,end_time,slot_type,source) VALUES(?,?,?,?,?,'FREE','AUTO')", (version_id, srow["id"], d.isoformat(), st, en))

            stats["completed"] = True
        with db() as c:
            c.execute("UPDATE generation_run SET status='succeeded', finished_at=?, solver_stats=? WHERE id=?", (dt.datetime.now().isoformat(), json.dumps(stats), run_id))
            c.execute("UPDATE timetable_version SET status='active' WHERE id=?", (version_id,))
    except Exception as e:
        with db() as c:
            c.execute("UPDATE generation_run SET status='failed', finished_at=?, solver_stats=? WHERE id=?", (dt.datetime.now().isoformat(), json.dumps({**stats, "error": str(e)}), run_id))
        raise
    return {"version_id": version_id, "run_id": run_id, "stats": stats}

def detect_conflicts(version_id):
    conflicts = []
    with db() as c:
        rows = c.execute("SELECT slot_date,start_time,end_time,teacher_id,COUNT(*) n FROM schedule_slot WHERE version_id=? AND slot_type='TEACHING' AND teacher_id IS NOT NULL GROUP BY slot_date,start_time,end_time,teacher_id HAVING n>1", (version_id,)).fetchall()
        for r in rows: conflicts.append({"severity": "BLOCKING", "rule": "TEACHER_CLASH", "detail": dict(r)})
        today = dt.date.today().isoformat()
        rows = c.execute("SELECT st.id, st.sub_topic_name, st.target_month_end FROM syllabus_topic st WHERE st.target_month_end IS NOT NULL AND st.target_month_end >= ? AND NOT EXISTS (SELECT 1 FROM schedule_slot ss WHERE ss.version_id=? AND ss.syllabus_topic_id=st.id)", (today, version_id)).fetchall()
        for r in rows:
            target_date = dt.date.fromisoformat(r["target_month_end"])
            if (target_date - dt.date.today()).days <= 30: conflicts.append({"severity": "WARNING", "rule": "SYLLABUS_UNPLACED", "detail": dict(r)})
    return conflicts

class LoginRequest(BaseModel):
    login_id: str
    password: str

app = FastAPI(title="St. John's NEXUS School Timetable Planner", version="1.2")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
from fastapi.staticfiles import StaticFiles
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.on_event("startup")
def _startup():
    os.makedirs("static", exist_ok=True)
    init_db()

@app.post("/api/login")
def api_login(req: LoginRequest):
    with db() as c:
        user = c.execute("SELECT * FROM app_user WHERE login_id=?", (req.login_id.strip(),)).fetchone()
        if not user or not pwd_ctx.verify(req.password, user["password_hash"]):
            raise HTTPException(401, "Invalid Teacher ID or password")
        token = jwt.encode({"sub": str(user["id"]), "role": user["role"], "login_id": user["login_id"]}, SECRET_KEY, algorithm=ALGORITHM)
        return {"access_token": token, "user": {"login_id": user["login_id"], "role": user["role"]}}

@app.post("/api/ingest/all")
def api_ingest_all():
    res = ingest_all()
    with db() as c:
        # FORCE CREATE TEACHER LOGINS
        for t in c.execute("SELECT id, name FROM teacher").fetchall():
            name = t["name"]
            if not c.execute("SELECT 1 FROM app_user WHERE login_id=?", (name,)).fetchone():
                c.execute("INSERT INTO app_user(login_id, password_hash, role, teacher_id) VALUES(?,?,?,?)", 
                          (name, pwd_ctx.hash(DEFAULT_PWD), "teacher", t["id"]))
    return res

@app.post("/api/generate")
def api_generate(payload: dict = Body(default={})):
    return generate_timetable(grade_filter=payload.get("grade"), seed=int(payload.get("seed", 42)))

@app.get("/api/versions")
def api_versions():
    with db() as c: return [dict(r) for r in c.execute("SELECT * FROM timetable_version ORDER BY id DESC").fetchall()]

@app.delete("/api/versions/{vid}")
def api_delete_version(vid: int):
    with db() as c:
        c.execute("DELETE FROM schedule_slot WHERE version_id=?", (vid,))
        c.execute("DELETE FROM conflict_log WHERE version_id=?", (vid,))
        c.execute("DELETE FROM generation_run WHERE version_id=?", (vid,))
        c.execute("DELETE FROM timetable_version WHERE id=?", (vid,))
    return {"ok": True}

@app.get("/api/versions/{vid}/sections/{sid}/day/{date}")
def api_day(vid: int, sid: int, date: str):
    with db() as c:
        rows = c.execute("""SELECT ss.*, sub.name subject_name, sub.code subject_code, t.name teacher_name, st.sub_topic_name topic
            FROM schedule_slot ss LEFT JOIN subject sub ON sub.id=ss.subject_id
            LEFT JOIN teacher t ON t.id=ss.teacher_id LEFT JOIN syllabus_topic st ON st.id=ss.syllabus_topic_id
            WHERE ss.version_id=? AND ss.section_id=? AND ss.slot_date=? ORDER BY ss.start_time""", (vid, sid, date)).fetchall()
        return [dict(r) for r in rows]

@app.get("/api/versions/{vid}/sections/{sid}/week/{date}")
def api_week(vid: int, sid: int, date: str):
    d = dt.date.fromisoformat(date); start = d - dt.timedelta(days=d.weekday()); end = start + dt.timedelta(days=6)
    with db() as c:
        rows = c.execute("""SELECT ss.*, sub.name subject_name, t.name teacher_name, st.sub_topic_name topic
            FROM schedule_slot ss LEFT JOIN subject sub ON sub.id=ss.subject_id
            LEFT JOIN teacher t ON t.id=ss.teacher_id LEFT JOIN syllabus_topic st ON st.id=ss.syllabus_topic_id
            WHERE ss.version_id=? AND ss.section_id=? AND ss.slot_date BETWEEN ? AND ? ORDER BY ss.slot_date, ss.start_time""", (vid, sid, start.isoformat(), end.isoformat())).fetchall()
        return [dict(r) for r in rows]

@app.get("/api/sections")
def api_sections():
    with db() as c: return [dict(r) for r in c.execute("SELECT s.id, s.label, g.name gname FROM section s JOIN grade g ON g.id=s.grade_id ORDER BY g.id, s.label").fetchall()]

@app.get("/api/subjects")
def api_subjects():
    with db() as c: return [dict(r) for r in c.execute("SELECT * FROM subject ORDER BY id").fetchall()]

@app.get("/api/teachers/eligible")
def api_eligible(subject_code: str, date: str, start: str, end: str, version_id: int):
    with db() as c:
        s = c.execute("SELECT id FROM subject WHERE UPPER(TRIM(code))=? OR UPPER(TRIM(name))=?", (subject_code.upper().strip(), subject_code.upper().strip())).fetchone()
        if not s: return []
        teachers = get_teachers_for_subject(c, s["id"]); out = []
        for t in teachers:
            ok = True; reason = None
            if not t["is_flexible"]:
                if t["shift_start"] and t["shift_start"] > start: ok, reason = False, "outside shift"
                if t["shift_end"] and t["shift_end"] < end: ok, reason = False, "outside shift"
            if ok and not teacher_free(c, version_id, date, start, end, t["id"]): ok, reason = False, "busy at this time"
            out.append({**dict(t), "eligible": ok, "reason": reason})
        return out

@app.patch("/api/slots/{slot_id}")
def api_override(slot_id: int, payload: dict = Body(...)):
    with db() as c:
        old = c.execute("SELECT * FROM schedule_slot WHERE id=?", (slot_id,)).fetchone()
        if not old: raise HTTPException(404, "slot not found")
        before = dict(old); new_subject = payload.get("subject_id", old["subject_id"])
        new_teacher = payload.get("teacher_id", old["teacher_id"]); new_topic = payload.get("syllabus_topic_id", old["syllabus_topic_id"])
        reason = payload.get("reason", "manual override")
        today_str = dt.date.today().isoformat()
        if new_teacher and old["slot_date"] >= today_str and new_teacher != old["teacher_id"]:
            if not teacher_free(c, old["version_id"], old["slot_date"], old["start_time"], old["end_time"], new_teacher):
                raise HTTPException(409, "teacher busy at that time")
        c.execute("UPDATE schedule_slot SET subject_id=?, teacher_id=?, syllabus_topic_id=?, slot_type='TEACHING', source='MANUAL', override_reason=? WHERE id=?", (new_subject, new_teacher, new_topic, reason, slot_id))
        return {"ok": True}

@app.get("/api/conflicts")
def api_conflicts(version_id: int): return detect_conflicts(version_id)

@app.get("/api/stats")
def api_stats():
    with db() as c:
        return {
            "grades": c.execute("SELECT COUNT(*) n FROM grade").fetchone()["n"], "sections": c.execute("SELECT COUNT(*) n FROM section").fetchone()["n"],
            "teachers": c.execute("SELECT COUNT(*) n FROM teacher").fetchone()["n"], "topics": c.execute("SELECT COUNT(*) n FROM syllabus_topic").fetchone()["n"],
            "exam_events": c.execute("SELECT COUNT(*) n FROM exam_event").fetchone()["n"], "versions": c.execute("SELECT COUNT(*) n FROM timetable_version").fetchone()["n"],
            "slots": c.execute("SELECT COUNT(*) n FROM schedule_slot").fetchone()["n"],
        }
@app.get("/api/export/excel")
def export_excel(version_id: int, section_id: int, date: str):
    import datetime as dt
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    from io import BytesIO
    from fastapi.responses import StreamingResponse
    
    d = dt.date.fromisoformat(date)
    start = d - dt.timedelta(days=d.weekday())
    end = start + dt.timedelta(days=5)
    
    with db() as c:
        sec = c.execute("SELECT s.label, g.name gname FROM section s JOIN grade g ON g.id=s.grade_id WHERE s.id=?", (section_id,)).fetchone()
        if not sec: raise HTTPException(404, "Section not found")
        
        rows = c.execute("""SELECT ss.slot_date, ss.start_time, ss.end_time, ss.slot_type, 
                            sub.name subject_name, t.name teacher_name, st.sub_topic_name topic
                            FROM schedule_slot ss 
                            LEFT JOIN subject sub ON sub.id=ss.subject_id
                            LEFT JOIN teacher t ON t.id=ss.teacher_id 
                            LEFT JOIN syllabus_topic st ON st.id=ss.syllabus_topic_id
                            WHERE ss.version_id=? AND ss.section_id=? AND ss.slot_date BETWEEN ? AND ? 
                            ORDER BY ss.slot_date, ss.start_time""", (version_id, section_id, start.isoformat(), end.isoformat())).fetchall()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Timetable"
    
    title = f"Timetable - {sec['gname']}-{sec['label']} - {start.isoformat()} to {end.isoformat()}"
    ws.merge_cells('A1:G1')
    ws['A1'] = title
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = PatternFill("solid", fgColor="8B0000")
    ws['A1'].alignment = Alignment(horizontal="center")
    
    ws.append(["Time", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
    for cell in ws[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="228B22")
    
    times = sorted(list(set([r["start_time"] + "-" + r["end_time"] for r in rows])))
    days = [(start + dt.timedelta(days=i)).isoformat() for i in range(6)]
    
    for t in times:
        row_data = [t]
        for day in days:
            match = next((r for r in rows if r["start_time"] + "-" + r["end_time"] == t and r["slot_date"] == day), None)
            if match:
                val = match["subject_name"] or match["slot_type"]
                if match["teacher_name"]: val += f"\n{match['teacher_name']}"
                row_data.append(val)
            else:
                row_data.append("")
        ws.append(row_data)
        
    ws.column_dimensions['A'].width = 15
    for col in ['B', 'C', 'D', 'E', 'F', 'G']:
        ws.column_dimensions[col].width = 25
        
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    
    # Format the filename as "Class-section-timetable"
    filename = f"Class{sec['gname']}-{sec['label']}-timetable.xlsx"
    
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename={filename}"})

# --------------------------------------------------------------- FRONTEND UI
HTML = r"""
<!DOCTYPE html><html><head><meta charset="utf-8"><title>St. John's NEXUS School · Timetable Planner</title>
<style>
body { 
  margin: 0; 
  font-family: 'Segoe UI', system-ui, sans-serif; 
  color: #1e1b4b; 
  background-color: #f8fafc; 
  min-height: 100vh; 
  z-index: 1; 
  position: relative;
}

body::before {
  content: "";
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  width: 50vw; 
  max-width: 800px;
  height: 80vh;
  background: url('/static/bg.jpg') no-repeat center center;
  background-size: contain; 
  opacity: 0.08; 
  z-index: -1; 
  pointer-events: none; 
}
header { background: #8B0000; color: #fff; padding: 14px 24px; display: flex; gap: 16px; align-items: center; box-shadow: 0 4px 20px rgba(139, 0, 0, 0.4); position: sticky; top: 0; z-index: 100; }
header h1 { font-size: 18px; margin: 0; font-weight: 700; letter-spacing: 0.5px; }
.tabs { display: flex; gap: 6px; margin-left: auto; }
.tab { padding: 8px 18px; border-radius: 8px; cursor: pointer; font-size: 13px; font-weight: 600; background: rgba(255, 255, 255, 0.15); transition: all 0.2s ease; border: 1px solid rgba(255, 255, 255, 0.1); }
.tab:hover { background: rgba(255, 255, 255, 0.25); }
.tab.on { background: #228B22; color: #fff; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15); transform: translateY(-1px); }
main { padding: 20px; max-width: 1500px; margin: 0 auto; }
.card { background: #fff; border: 1px solid #e2e8f0; border-radius: 14px; padding: 18px; margin-bottom: 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; }
.kpi { background: #fff; border-radius: 10px; padding: 12px; position: relative; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; }
.kpi::before { content: ""; position: absolute; top: 0; left: 0; width: 100%; height: 4px; background: #228B22; }
.kpi .v { font-size: 24px; font-weight: 800; color: #8B0000; }
.kpi .l { font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: 600; margin-top: 4px; }
button { background: #228B22; color: #fff; border: 0; padding: 9px 16px; border-radius: 8px; cursor: pointer; font-size: 13px; font-weight: 600; transition: all 0.2s ease; }
button:hover { background: #1a6b1a; }
button.sec { background: #e2e8f0; color: #475569; }
button.sec:hover { background: #cbd5e1; }
select, input { padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 13px; background: #fff; color: #1e293b; outline: none; }
select:focus, input:focus { border-color: #228B22; box-shadow: 0 0 0 3px rgba(34, 139, 34, 0.2); }
.row { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; }
table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13px; }
th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid #e2e8f0; }
th { background: #f8fafc; font-weight: 700; color: #475569; text-transform: uppercase; font-size: 11px; }
tr:hover td { background: #f8fafc; }
.grid { display: grid; grid-template-columns: 75px repeat(6, 1fr); gap: 6px; }
.slot { background: #fff; border-radius: 8px; padding: 6px 8px; font-size: 11px; min-height: 54px; box-shadow: 0 2px 6px rgba(0,0,0,0.04); border-left: 4px solid #228B22; }
.slot:hover { transform: scale(1.02); z-index: 10; }
.slot.break { background: #fffbeb; border-color: #f59e0b; text-align: center; font-weight: 600; color: #b45309; }
.slot.exam { background: #fef2f2; border-color: #ef4444; font-weight: 700; color: #991b1b; }
.slot.revision { background: #f5f3ff; border-color: #8b5cf6; color: #5b21b6; }
.slot.manual { box-shadow: inset 0 0 0 2px #f59e0b; }
.slot .s { font-weight: 700; font-size: 12px; margin-bottom: 2px; color: #8B0000; } 
.slot .t { color: #64748b; font-size: 10px; }
.hdr { background: #8B0000; color: #fff; padding: 8px; text-align: center; border-radius: 6px; font-size: 12px; font-weight: 600; }
.hdr.time { background: #228B22; }
.hidden { display: none !important; }
#loginScreen { position: fixed; inset: 0; background: #f8fafc; display: flex; align-items: center; justify-content: center; z-index: 9999; }
.login-box { background: #fff; padding: 40px; border-radius: 14px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); width: 100%; max-width: 400px; border-top: 8px solid #8B0000; }
.login-box h2 { color: #8B0000; text-align: center; margin-top: 0; margin-bottom: 24px; }
.password-wrapper { position: relative; width: 100%; box-sizing: border-box; }
.password-wrapper input { padding-right: 40px; box-sizing: border-box; width: 100%; }
.toggle-pwd { position: absolute; right: 10px; top: 50%; transform: translateY(-50%); cursor: pointer; background: none; border: none; font-size: 16px; color: #64748b; padding: 0; }
.toggle-pwd:hover { color: #8B0000; background: none; }
</style></head><body>

<div id="loginScreen">
  <div class="login-box">
    <h2 style="display:flex; flex-direction:column; align-items:center; gap:10px;">
  <img src="/static/emblem.jpg" alt="Emblem" style="height:60px; border-radius:6px;">
  St. John's NEXUS
</h2>
    <div class="row" style="flex-direction:column; gap:16px;">
      <input id="logId" placeholder="Teacher ID (e.g. Teacher1 or ADMIN)" style="width:100%; padding:12px;">
      <div class="password-wrapper">
        <input id="logPw" type="password" placeholder="Password" style="width:100%; padding:12px;">
        <button type="button" class="toggle-pwd" onclick="togglePassword()">👁️</button>
      </div>
      <button onclick="doLogin()" style="width:100%; padding:12px; font-size:16px;">Sign In</button>
      <div id="logErr" style="color:#8B0000; font-size:13px; text-align:center;"></div>
    </div>
  </div>
</div>

<header>
  <div style="display:flex; align-items:center; gap:12px;">
  <img src="/static/emblem.jpg" alt="Emblem" style="height:40px; border-radius:4px;">
  <h1>St. John's NEXUS School</h1>
</div>
  <div class="tabs">
    <div class="tab on" data-tab="dash" id="tabDash">Dashboard</div>
    <div class="tab" data-tab="grid">Day Grid</div>
    <div class="tab" data-tab="manual" id="tabMan">Manual Override</div>
  </div>
  <button class="sec" style="margin-left:20px; background:#fff; color:#8B0000;" onclick="logout()">Logout</button>
</header>
<main>
<section id="dash">
  <div class="card" id="adminSetup">
    <div class="row" style="margin-bottom:10px">
      <b style="font-size: 16px; color: #8B0000;">✨ System Setup</b>
      <button onclick="ingestAll()">📥 Ingest /data/ files</button>
      <span id="ingestMsg" style="font-size:12px; font-weight:600;"></span>
    </div>
    <div class="row">
      <b>Generate</b>
      <select id="genGrade"><option value="">All grades</option></select>
      <input id="genSeed" type="number" value="42" style="width:80px">
      <button onclick="generate()">⚙️ Generate timetable</button>
      <span id="genMsg" style="font-size:12px; font-weight:600;"></span>
    </div>
  </div>
  <div class="card"><div class="kpis" id="kpis"></div></div>
  <div class="card"><b>Versions</b><div id="versions"></div></div>
</section>
<section id="grid" class="hidden">
  <div class="card">
    <div class="row">
      <select id="gridVersion"></select>
      <select id="gridSection"></select>
      <input type="date" id="gridDate">
      <button onclick="loadWeek()">Load week</button>
<button class="sec" onclick="downloadExcel()">📥 Download Excel</button>
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
  </div>
</section>
</main>

<script>
const $=id=>document.getElementById(id);
const API='/api';

function togglePassword() {
  const pwInput = $('logPw');
  if (pwInput.type === 'password') {
    pwInput.type = 'text';
  } else {
    pwInput.type = 'password';
  }
}

let state={sections:[],subjects:[],teachers:[],versions:[],editingSlot:null,token:null,role:null};

async function doLogin() {
  const id = $('logId').value;
  const pw = $('logPw').value;
  try {
    const res = await fetch(API+'/login', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({login_id:id, password:pw})
    });
    if(!res.ok) throw new Error('Invalid ID or password');
    const data = await res.json();
    state.token = data.access_token;
    state.role = data.user.role;
    $('loginScreen').style.display = 'none';
    
    if (state.role === 'teacher') {
      $('tabDash').classList.add('hidden');
      $('tabMan').classList.add('hidden');
      $('adminSetup').classList.add('hidden');
      document.querySelector('[data-tab="grid"]').click();
    }
    boot();
  } catch(e) {
    $('logErr').textContent = e.message;
  }
}

function logout() { state.token = null; location.reload(); }

document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('on')); t.classList.add('on');
  ['dash','grid','manual'].forEach(id=>$(id).classList.add('hidden'));
  $(t.dataset.tab).classList.remove('hidden');
});

async function j(url, opts){
  opts = opts || {};
  opts.headers = opts.headers || {};
  if(state.token) opts.headers['Authorization'] = 'Bearer ' + state.token;
  const r=await fetch(url,opts); 
  if(!r.ok) throw new Error(await r.text()); 
  return r.json();
}

async function boot(){
  state.sections=await j(API+'/sections');
  state.subjects=await j(API+'/subjects');
  const grades=[...new Set(state.sections.map(s=>s.gname))];
  $('genGrade').innerHTML='<option value="">All grades</option>'+grades.map(g=>`<option>${g}</option>`).join('');
  const secOpts=state.sections.map(s=>`<option value="${s.id}">${s.gname}-${s.label}</option>`).join('');
  const today = new Date().toISOString().slice(0,10);
  
  $('gridSection').innerHTML=secOpts; $('manSection').innerHTML=secOpts;
  $('edSubject').innerHTML=state.subjects.map(s=>`<option value="${s.id}" data-code="${s.code}">${s.name}</option>`).join('');
  
  $('gridDate').value=today; $('manDate').value=today;
  await refreshStats(); await refreshVersions();
}

async function refreshStats(){
  const s=await j(API+'/stats');
  $('kpis').innerHTML=Object.entries(s).map(([k,v])=>`<div class="kpi"><div class="v">${v}</div><div class="l">${k}</div></div>`).join('');
}

async function refreshVersions(){
  state.versions=await j(API+'/versions');
  const opts=state.versions.map(v=>{
    const dateParts = v.created_at.slice(0,10).split('-');
    const formattedDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
    return `<option value="${v.id}">v${v.version_no} · ${v.status} · ${formattedDate}</option>`;
  }).join('');
  $('gridVersion').innerHTML=opts; $('manVersion').innerHTML=opts;

  if(!state.versions.length) { $('versions').innerHTML = '<div style="padding:10px; color:#8B0000;">No versions generated yet.</div>'; return; }
  
  let html = '<table><tr><th>Version</th><th>Created</th><th>Status</th><th>Actions</th></tr>';
  html += state.versions.map(v => `
    <tr>
      <td><b>v${v.version_no}</b></td>
      <td>${(() => {
        const d = new Date(v.created_at);
        return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
      })()}</td>
      <td><span style="color:${v.status==='active'?'#228B22':'#64748b'}; font-weight:600;">${v.status.toUpperCase()}</span></td>
      <td>${state.role==='admin' ? `<button class="sec" onclick="deleteVersion(${v.id})" style="padding:4px 8px; font-size:11px; background:#fee2e2; color:#b91c1c;">Delete</button>` : ''}</td>
    </tr>
  `).join('');
  html += '</table>';
  $('versions').innerHTML = html;
}

async function deleteVersion(vid) {
    if(!confirm('Delete this version?')) return;
    try { await j(`${API}/versions/${vid}`, {method: 'DELETE'}); await refreshVersions(); await refreshStats(); } catch(e) { alert(e.message); }
}

async function ingestAll(){
  $('ingestMsg').textContent='Ingesting…';
  try{
    const r=await j(API+'/ingest/all',{method:'POST'});
    $('ingestMsg').innerHTML=`<span style="color:#228B22;">✓ ${r.total_files} files ingested</span>`;
    await refreshStats();
  }catch(e){$('ingestMsg').textContent='Error: '+e.message;}
}

async function generate(){
  $('genMsg').textContent='Generating…';
  try{
    const r=await j(API+'/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({grade:$('genGrade').value||null, seed:42})});
    $('genMsg').innerHTML=`<span style="color:#228B22;">Done: v${r.version_id} · ${r.stats.slots_inserted} slots</span>`;
    await refreshStats(); await refreshVersions();
  }catch(e){$('genMsg').textContent='Error: '+e.message;}
}
function downloadExcel(){
  const vid=$('gridVersion').value, sid=$('gridSection').value, d=$('gridDate').value;
  if(!vid||!sid||!d){ alert('Please load a week first'); return; }
  window.location.href = `${API}/export/excel?version_id=${vid}&section_id=${sid}&date=${d}`;
}
async function loadWeek(){
  const vid=$('gridVersion').value, sid=$('gridSection').value, d=$('gridDate').value;
  if(!vid||!sid||!d){$('weekGrid').innerHTML='<div style="color:#8B0000;">Generate a version first</div>';return;}
  const rows=await j(`${API}/versions/${vid}/sections/${sid}/week/${d}`);
  const start=new Date(d); start.setDate(start.getDate()-start.getDay()+1);
  const days=[]; for(let i=0;i<6;i++){const x=new Date(start);x.setDate(x.getDate()+i);days.push(x);}
  const byTime={};
  rows.forEach(r=>{const k=r.start_time+'-'+r.end_time; (byTime[k]=byTime[k]||{})[r.slot_date]=r;});
  let html='<div class="grid"><div class="hdr time">Time</div>'+days.map(d=>{
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `<div class="hdr">${dd}/${mm}/${yyyy}</div>`;
  }).join('');
  const times=Object.keys(byTime).sort();
  for(const t of times){
    html+=`<div class="hdr time">${t}</div>`;
    for(const d of days){
      const iso=d.toISOString().slice(0,10); const s=byTime[t][iso];
      if(!s){html+='<div class="slot"></div>';continue;}
      let cls='slot '+(s.slot_type||'').toLowerCase()+(s.source==='MANUAL'?' manual':'');
      const label=s.slot_type==='EXAM'?'EXAM':s.slot_type==='REVISION'?'REVISION':(s.subject_name||'Free');
      html+=`<div class="${cls}"><div class="s">${label}</div><div class="t">${s.teacher_name||''}</div></div>`;
    }
  }
  html+='</div>'; $('weekGrid').innerHTML=html;
}

async function loadSlots(){
  const vid=$('manVersion').value, sid=$('manSection').value, d=$('manDate').value;
  if(!vid){$('slotList').innerHTML='<div style="color:#8B0000;">Generate first</div>';return;}
  const rows=await j(`${API}/versions/${vid}/sections/${sid}/day/${d}`);
  if(!rows.length){$('slotList').innerHTML='<div style="color:#8B0000;">No slots for this date</div>';return;}
  $('slotList').innerHTML='<table><tr><th>Time</th><th>Type</th><th>Subject</th><th>Teacher</th><th></th></tr>'+
    rows.map(r=>{
      const canEdit = r.slot_type === 'TEACHING' || r.slot_type === 'FREE';
      const btn = canEdit ? `<button onclick='startEdit(${JSON.stringify(r)})'>Edit</button>` : '';
      return `<tr><td>${r.start_time}–${r.end_time}</td><td>${r.slot_type}</td><td>${r.subject_name||''}</td><td>${r.teacher_name||''}</td><td>${btn}</td></tr>`;
    }).join('')+'</table>';
}

async function startEdit(slot){
  state.editingSlot=slot; $('editPanel').classList.remove('hidden');
  if (slot.subject_id) $('edSubject').value = slot.subject_id;
  await loadTopics();
  if(slot.syllabus_topic_id) $('edTopic').value=slot.syllabus_topic_id;
  $('edReason').value='';
  await checkValid();
}

async function loadTopics(){
  const sid=$('edSubject').value; const sub=state.subjects.find(s=>s.id==sid);
  if(!sub) return;
  const sec=state.sections.find(s=>s.id==$('manSection').value); const grade=sec?sec.gname:'8';
  try{
    const topics=await j(`${API}/syllabus/topics?grade=${encodeURIComponent(grade)}&subject_code=${encodeURIComponent(sub.code)}`);
    $('edTopic').innerHTML=topics.map(t=>`<option value="${t.id}">${t.chapter_no} · ${t.sub_topic_no||''} ${t.sub_topic_name||''}</option>`).join('')||'<option value="">(no topics)</option>';
  }catch(e){$('edTopic').innerHTML='<option value="">(no topics)</option>';}
}

async function checkValid(){
  const slot=state.editingSlot; if(!slot) return;
  const sub=state.subjects.find(s=>s.id==$('edSubject').value);
  if(!sub) { $('edTeacher').innerHTML = '<option value="">(no subject)</option>'; return; }
  try{
    const vid=$('manVersion').value;
    const eligible=await j(`${API}/teachers/eligible?grade=${encodeURIComponent(state.sections.find(s=>s.id==$('manSection').value).gname)}&subject_code=${encodeURIComponent(sub.code)}&date=${slot.slot_date}&start=${slot.start_time}&end=${slot.end_time}&version_id=${vid}`);
    $('edTeacher').innerHTML=eligible.map(t=>`<option value="${t.id}" ${!t.eligible?'disabled':''}>${t.name}${t.eligible?' ✓':' ✗'}</option>`).join('');
  }catch(e){}
}

async function saveOverride(){
  const slot=state.editingSlot; if(!slot) return;
  const body={subject_id:parseInt($('edSubject').value), teacher_id:parseInt($('edTeacher').value)||null, syllabus_topic_id:parseInt($('edTopic').value)||null, reason:$('edReason').value||'override'};
  try{ await j(`${API}/slots/${slot.id}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)}); alert('Saved'); closeEdit(); loadSlots(); }catch(e){alert('Error: '+e.message);}
}

function closeEdit(){ 
  $('editPanel').classList.add('hidden'); 
  state.editingSlot=null; 
}

$('edSubject').addEventListener('change',loadTopics);
$('edSubject').addEventListener('change',checkValid);
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def index(): return HTML
@app.get("/api/debug_users")
def debug_users():
    with db() as c:
        users = c.execute("SELECT login_id, role FROM app_user").fetchall()
        return [dict(u) for u in users]
if __name__ == "__main__":
    import uvicorn
    os.makedirs(DATA_DIR, exist_ok=True)
    print("="*70)
    print(" St. John's NEXUS School Timetable Planner")
    print(" Open http://127.0.0.1:8000")
    print("="*70)
    uvicorn.run(app, host="127.0.0.1", port=8000)