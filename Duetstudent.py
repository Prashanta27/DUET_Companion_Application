"""
DUET Student Companion — ALL FEATURES single-file PyQt5 app

Features:
- Sidebar + stacked pages (pure PyQt5)
- Schedule (add/edit/delete) + export schedule -> PDF
- Exam manager + countdown timer
- Weighted GPA calculator (supports adding semesters)
- Attendance tracking stored in SQLite
- Notes manager with Markdown preview (uses `markdown` package if available)
- Voice assistant (TTS via pyttsx3; optional STT via SpeechRecognition)
- Storage: SQLite database at ~/.duet_companion/duet.db and files for exported PDFs
- Guidance for Android APK (BeeWare / PySide) in code comments

Run: python duet_companion_advanced_allinone.py
Requires: PyQt5
Optional: markdown, pyttsx3, SpeechRecognition
"""

import sys, os, json, sqlite3, traceback
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QWidget, QMainWindow, QHBoxLayout, QVBoxLayout, QFrame,
    QPushButton, QStackedWidget, QLabel, QTableWidget, QTableWidgetItem,
    QLineEdit, QTextEdit, QMessageBox, QFileDialog, QListWidget, QListWidgetItem,
    QHeaderView, QInputDialog, QDateTimeEdit, QDateEdit, QSpinBox, QComboBox,
    QToolButton
)
from PyQt5.QtCore import Qt, QTimer, QDateTime
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QTextDocument

# Optional libs
try:
    import markdown as md_lib
except Exception:
    md_lib = None

try:
    import pyttsx3
    tts_engine = pyttsx3.init()
except Exception:
    tts_engine = None

try:
    import speech_recognition as sr
    sr_recognizer = sr.Recognizer()
except Exception:
    sr_recognizer = None

# Paths
DATA_DIR = os.path.join(os.path.expanduser('~'), '.duet_companion')
DB_FILE = os.path.join(DATA_DIR, 'duet.db')
EXPORT_DIR = os.path.join(DATA_DIR, 'exports')
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)

# ---------------------------------------------------------
# Database utilities (SQLite)
# ---------------------------------------------------------
def get_db_conn():
    conn = sqlite3.connect(DB_FILE, detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_conn()
    cur = conn.cursor()
    # schedules table: id, day, time, course, room
    cur.execute('''
    CREATE TABLE IF NOT EXISTS schedule (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        day TEXT, time TEXT, course TEXT, room TEXT
    )
    ''')
    # exams table: id, title, datetime_iso
    cur.execute('''
    CREATE TABLE IF NOT EXISTS exams (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT,
        dt TEXT
    )
    ''')
    # notes table
    cur.execute('''
    CREATE TABLE IF NOT EXISTS notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT,
        body TEXT,
        created TEXT
    )
    ''')
    # attendance table: id, course, date_iso, present INTEGER (0/1)
    cur.execute('''
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        course TEXT,
        date TEXT,
        present INTEGER
    )
    ''')
    # gpa table: id, semester TEXT, gpa REAL, credits REAL
    cur.execute('''
    CREATE TABLE IF NOT EXISTS gpa (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        semester TEXT,
        gpa REAL,
        credits REAL
    )
    ''')
    conn.commit()
    conn.close()

init_db()

# ---------------------------------------------------------
# Main Application Window (All-in-one)
# ---------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DUET Student Companion — All Features")
        self.resize(1200, 720)
        self.setStyleSheet("background:#121214; color:#eaeaea;")

        main = QHBoxLayout()
        central = QWidget()
        central.setLayout(main)
        self.setCentralWidget(central)

        # Left Sidebar
        sidebar = QFrame()
        sidebar.setFixedWidth(220)
        sidebar.setStyleSheet("background:#1f2023;")
        side_layout = QVBoxLayout(sidebar)
        side_layout.setContentsMargins(10,10,10,10)

        title = QLabel("DUET Companion")
        title.setStyleSheet("font-weight:700; font-size:18px;")
        title.setAlignment(Qt.AlignCenter)
        side_layout.addWidget(title)

        # Buttons
        self.btn_home = QPushButton("Home")
        self.btn_schedule = QPushButton("Schedule")
        self.btn_exams = QPushButton("Exams")
        self.btn_gpa = QPushButton("GPA Calculator")
        self.btn_attendance = QPushButton("Attendance")
        self.btn_notes = QPushButton("Notes (Markdown)")
        self.btn_voice = QPushButton("Voice Assistant")
        self.btn_export = QPushButton("Export Schedule PDF")
        self.btn_settings = QPushButton("Settings")

        for b in (self.btn_home, self.btn_schedule, self.btn_exams, self.btn_gpa,
                  self.btn_attendance, self.btn_notes, self.btn_voice, self.btn_export, self.btn_settings):
            b.setFixedHeight(40)
            b.setStyleSheet(self.btn_style())
            side_layout.addWidget(b)

        side_layout.addStretch()
        main.addWidget(sidebar)

        # Stacked pages
        self.stack = QStackedWidget()
        main.addWidget(self.stack, 1)

        # ---------- HOME ----------
        self.page_home = QWidget()
        hl = QVBoxLayout(self.page_home)
        lbl = QLabel("Welcome — DUET Student Companion\nAll features loaded.")
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet("font-size:20px; padding:30px;")
        hl.addWidget(lbl)
        self.stack.addWidget(self.page_home)

        # ---------- SCHEDULE ----------
        self.page_schedule = QWidget()
        sch_layout = QVBoxLayout(self.page_schedule)
        self.tbl_schedule = QTableWidget(0,4)
        self.tbl_schedule.setHorizontalHeaderLabels(["Day","Time","Course","Room"])
        self.tbl_schedule.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        sch_layout.addWidget(self.tbl_schedule)

        row_inputs = QHBoxLayout()
        self.in_day = QLineEdit(); self.in_day.setPlaceholderText("Day")
        self.in_time = QLineEdit(); self.in_time.setPlaceholderText("Time")
        self.in_course = QLineEdit(); self.in_course.setPlaceholderText("Course")
        self.in_room = QLineEdit(); self.in_room.setPlaceholderText("Room")
        row_inputs.addWidget(self.in_day); row_inputs.addWidget(self.in_time)
        row_inputs.addWidget(self.in_course); row_inputs.addWidget(self.in_room)
        sch_layout.addLayout(row_inputs)

        btns = QHBoxLayout()
        self.btn_add = QPushButton("Add")
        self.btn_edit = QPushButton("Edit Selected")
        self.btn_delete = QPushButton("Delete Selected")
        btns.addWidget(self.btn_add); btns.addWidget(self.btn_edit); btns.addWidget(self.btn_delete)
        sch_layout.addLayout(btns)
        self.stack.addWidget(self.page_schedule)

        # ---------- EXAMS ----------
        self.page_exams = QWidget()
        ex_layout = QVBoxLayout(self.page_exams)
        self.list_exams = QListWidget()
        ex_layout.addWidget(self.list_exams)
        ex_form = QHBoxLayout()
        self.ex_title = QLineEdit(); self.ex_title.setPlaceholderText("Exam Title")
        self.ex_dt = QDateTimeEdit(QDateTime.currentDateTime()); self.ex_dt.setCalendarPopup(True)
        ex_form.addWidget(self.ex_title); ex_form.addWidget(self.ex_dt)
        ex_layout.addLayout(ex_form)
        ex_btns = QHBoxLayout()
        self.btn_add_exam = QPushButton("Add Exam")
        self.btn_del_exam = QPushButton("Delete Exam")
        self.btn_start_countdown = QPushButton("Start Countdown (show top)")
        ex_btns.addWidget(self.btn_add_exam); ex_btns.addWidget(self.btn_del_exam); ex_btns.addWidget(self.btn_start_countdown)
        ex_layout.addLayout(ex_btns)
        self.lbl_countdown = QLabel("No countdown running")
        self.lbl_countdown.setStyleSheet("font-size:18px; padding:8px;")
        ex_layout.addWidget(self.lbl_countdown)
        self.stack.addWidget(self.page_exams)

        # ---------- GPA ----------
        self.page_gpa = QWidget()
        g_layout = QVBoxLayout(self.page_gpa)
        g_help = QLabel("Add semesters as GPA:credits (weighted). App computes cumulative CGPA.")
        g_layout.addWidget(g_help)
        g_form = QHBoxLayout()
        self.in_semester = QLineEdit(); self.in_semester.setPlaceholderText("Semester name (e.g., Sem1)")
        self.in_gpa = QLineEdit(); self.in_gpa.setPlaceholderText("GPA (e.g., 3.75)")
        self.in_credits = QLineEdit(); self.in_credits.setPlaceholderText("Credits (e.g., 15)")
        g_form.addWidget(self.in_semester); g_form.addWidget(self.in_gpa); g_form.addWidget(self.in_credits)
        g_layout.addLayout(g_form)
        g_btns = QHBoxLayout()
        self.btn_add_sem = QPushButton("Add Semester"); self.btn_calc_cgpa = QPushButton("Calculate CGPA")
        g_btns.addWidget(self.btn_add_sem); g_btns.addWidget(self.btn_calc_cgpa)
        g_layout.addLayout(g_btns)
        self.tbl_gpa = QTableWidget(0,3)
        self.tbl_gpa.setHorizontalHeaderLabels(["Semester","GPA","Credits"])
        self.tbl_gpa.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        g_layout.addWidget(self.tbl_gpa)
        self.lbl_cgpa = QLabel("CGPA: -")
        self.lbl_cgpa.setStyleSheet("font-size:18px;")
        g_layout.addWidget(self.lbl_cgpa)
        self.stack.addWidget(self.page_gpa)

        # ---------- ATTENDANCE ----------
        self.page_att = QWidget()
        at_layout = QVBoxLayout(self.page_att)
        at_controls = QHBoxLayout()
        self.att_course = QLineEdit(); self.att_course.setPlaceholderText("Course name")
        self.att_date = QDateEdit(); self.att_date.setCalendarPopup(True); self.att_date.setDate(datetime.now().date())
        self.att_present = QComboBox(); self.att_present.addItems(["Present","Absent"])
        at_controls.addWidget(self.att_course); at_controls.addWidget(self.att_date); at_controls.addWidget(self.att_present)
        at_layout.addLayout(at_controls)
        at_btns = QHBoxLayout()
        self.btn_mark_att = QPushButton("Mark Attendance")
        self.btn_view_att = QPushButton("Load Attendance for Course")
        at_btns.addWidget(self.btn_mark_att); at_btns.addWidget(self.btn_view_att)
        at_layout.addLayout(at_btns)
        self.tbl_att = QTableWidget(0,3)
        self.tbl_att.setHorizontalHeaderLabels(["Course","Date","Present"])
        self.tbl_att.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        at_layout.addWidget(self.tbl_att)
        self.stack.addWidget(self.page_att)

        # ---------- NOTES (Markdown) ----------
        self.page_notes = QWidget()
        n_layout = QHBoxLayout(self.page_notes)
        left = QVBoxLayout()
        self.list_notes = QListWidget()
        left.addWidget(self.list_notes)
        n_buttons = QHBoxLayout()
        self.btn_new_note = QPushButton("New")
        self.btn_save_note = QPushButton("Save")
        self.btn_delete_note = QPushButton("Delete")
        n_buttons.addWidget(self.btn_new_note); n_buttons.addWidget(self.btn_save_note); n_buttons.addWidget(self.btn_delete_note)
        left.addLayout(n_buttons)
        n_layout.addLayout(left,1)

        right = QVBoxLayout()
        self.note_title = QLineEdit(); self.note_title.setPlaceholderText("Note title")
        self.note_body = QTextEdit(); self.note_body.setPlaceholderText("Write markdown here...")
        right.addWidget(self.note_title); right.addWidget(self.note_body)
        # Preview button
        self.btn_preview = QPushButton("Preview Markdown")
        right.addWidget(self.btn_preview)
        # Preview area
        self.preview = QTextEdit(); self.preview.setReadOnly(True)
        right.addWidget(self.preview,1)
        n_layout.addLayout(right,2)
        self.stack.addWidget(self.page_notes)

        # ---------- VOICE ASSISTANT ----------
        self.page_voice = QWidget()
        v_layout = QVBoxLayout(self.page_voice)
        self.txt_speak = QLineEdit(); self.txt_speak.setPlaceholderText("Text to speak")
        v_layout.addWidget(self.txt_speak)
        v_btns = QHBoxLayout()
        self.btn_tts = QPushButton("Speak (TTS)")
        self.btn_stt = QPushButton("Listen (STT)")
        v_btns.addWidget(self.btn_tts); v_btns.addWidget(self.btn_stt)
        v_layout.addLayout(v_btns)
        self.stt_result = QLabel("STT result will show here")
        v_layout.addWidget(self.stt_result)
        self.stack.addWidget(self.page_voice)

        # connect navigation
        self.btn_home.clicked.connect(lambda: self.stack.setCurrentWidget(self.page_home))
        self.btn_schedule.clicked.connect(lambda: self.stack.setCurrentWidget(self.page_schedule))
        self.btn_exams.clicked.connect(lambda: self.stack.setCurrentWidget(self.page_exams))
        self.btn_gpa.clicked.connect(lambda: self.stack.setCurrentWidget(self.page_gpa))
        self.btn_attendance.clicked.connect(lambda: self.stack.setCurrentWidget(self.page_att))
        self.btn_notes.clicked.connect(lambda: self.stack.setCurrentWidget(self.page_notes))
        self.btn_voice.clicked.connect(lambda: self.stack.setCurrentWidget(self.page_voice))
        self.btn_export.clicked.connect(self.export_schedule_pdf)

        # schedule actions
        self.btn_add.clicked.connect(self.add_schedule)
        self.btn_delete.clicked.connect(self.delete_schedule)
        self.btn_edit.clicked.connect(self.edit_schedule)
        self.tbl_schedule.itemSelectionChanged.connect(self.on_schedule_select)

        # exams actions
        self.btn_add_exam.clicked.connect(self.add_exam)
        self.btn_del_exam.clicked.connect(self.delete_exam)
        self.btn_start_countdown.clicked.connect(self.start_countdown)
        self.exam_timer = QTimer(); self.exam_timer.timeout.connect(self._tick_countdown)
        self.current_countdown_target = None

        # GPA actions
        self.btn_add_sem.clicked.connect(self.add_semester)
        self.btn_calc_cgpa.clicked.connect(self.calculate_cgpa)

        # Attendance actions
        self.btn_mark_att.clicked.connect(self.mark_attendance)
        self.btn_view_att.clicked.connect(self.view_attendance)

        # Notes actions
        self.btn_new_note.clicked.connect(self.new_note)
        self.btn_save_note.clicked.connect(self.save_note)
        self.btn_delete_note.clicked.connect(self.delete_note)
        self.list_notes.itemClicked.connect(self.load_note)
        self.btn_preview.clicked.connect(self.preview_markdown)

        # Voice assistant actions
        self.btn_tts.clicked.connect(self.do_tts)
        self.btn_stt.clicked.connect(self.do_stt)

        # Load DB data into pages
        self.load_schedule()
        self.load_exams()
        self.load_notes_list()
        self.load_gpa_table()
        self.load_attendance_table()

    # -------------------- Styles --------------------
    def btn_style(self):
        return """
        QPushButton { background:#2b2c30; color:#eaeaea; border-radius:6px; padding:6px; }
        QPushButton:hover { background:#404246; }
        """

    # -------------------- Schedule --------------------
    def load_schedule(self):
        self.tbl_schedule.setRowCount(0)
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute("SELECT id,day,time,course,room FROM schedule ORDER BY id")
        for r in cur.fetchall():
            row = self.tbl_schedule.rowCount()
            self.tbl_schedule.insertRow(row)
            self.tbl_schedule.setItem(row,0,QTableWidgetItem(r['day']))
            self.tbl_schedule.setItem(row,1,QTableWidgetItem(r['time']))
            self.tbl_schedule.setItem(row,2,QTableWidgetItem(r['course']))
            self.tbl_schedule.setItem(row,3,QTableWidgetItem(r['room']))
        conn.close()

    def add_schedule(self):
        d = self.in_day.text().strip(); t = self.in_time.text().strip()
        c = self.in_course.text().strip(); r = self.in_room.text().strip()
        if not (d and t and c and r):
            QMessageBox.warning(self,"Missing","Fill all fields")
            return
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("INSERT INTO schedule(day,time,course,room) VALUES(?,?,?,?)",(d,t,c,r))
        conn.commit(); conn.close()
        self.load_schedule()
        self.in_day.clear(); self.in_time.clear(); self.in_course.clear(); self.in_room.clear()

    def delete_schedule(self):
        row = self.tbl_schedule.currentRow()
        if row < 0:
            QMessageBox.warning(self,"Select","Select a row first")
            return
        # identify by matching fields (simple approach)
        day = self.tbl_schedule.item(row,0).text()
        timev = self.tbl_schedule.item(row,1).text()
        course = self.tbl_schedule.item(row,2).text()
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("DELETE FROM schedule WHERE day=? AND time=? AND course=? LIMIT 1",(day,timev,course))
        conn.commit(); conn.close()
        self.load_schedule()

    def edit_schedule(self):
        row = self.tbl_schedule.currentRow()
        if row < 0:
            QMessageBox.warning(self,"Select","Select a row first")
            return
        # simple inline edit via inputs
        d = self.in_day.text().strip(); t = self.in_time.text().strip()
        c = self.in_course.text().strip(); r = self.in_room.text().strip()
        if not (d and t and c and r):
            QMessageBox.warning(self,"Missing","Fill all fields to update")
            return
        # delete old and insert new (simple)
        old_day = self.tbl_schedule.item(row,0).text(); old_time = self.tbl_schedule.item(row,1).text()
        old_course = self.tbl_schedule.item(row,2).text()
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("DELETE FROM schedule WHERE day=? AND time=? AND course=? LIMIT 1",(old_day,old_time,old_course))
        cur.execute("INSERT INTO schedule(day,time,course,room) VALUES(?,?,?,?)",(d,t,c,r))
        conn.commit(); conn.close()
        self.load_schedule()

    def on_schedule_select(self):
        row = self.tbl_schedule.currentRow()
        if row >= 0:
            self.in_day.setText(self.tbl_schedule.item(row,0).text())
            self.in_time.setText(self.tbl_schedule.item(row,1).text())
            self.in_course.setText(self.tbl_schedule.item(row,2).text())
            self.in_room.setText(self.tbl_schedule.item(row,3).text())

    # -------------------- Export PDF --------------------
    def export_schedule_pdf(self):
        # build simple HTML table from schedule and print to PDF via QPrinter
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("SELECT day,time,course,room FROM schedule")
        rows = cur.fetchall(); conn.close()
        html = "<h2>DUET Class Schedule</h2><table border='1' cellspacing='0' cellpadding='6'>"
        html += "<tr><th>Day</th><th>Time</th><th>Course</th><th>Room</th></tr>"
        for r in rows:
            html += f"<tr><td>{r['day']}</td><td>{r['time']}</td><td>{r['course']}</td><td>{r['room']}</td></tr>"
        html += "</table>"
        # ask file path
        path, _ = QFileDialog.getSaveFileName(self, "Export Schedule to PDF", EXPORT_DIR, "PDF Files (*.pdf)")
        if not path:
            return
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(path)
        doc = QTextDocument()
        doc.setHtml(html)
        doc.print_(printer)
        QMessageBox.information(self,"Exported",f"Schedule exported to:\n{path}")

    # -------------------- Exams & Countdown --------------------
    def load_exams(self):
        self.list_exams.clear()
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("SELECT id,title,dt FROM exams ORDER BY dt")
        for r in cur.fetchall():
            item = QListWidgetItem(f"{r['title']} @ {r['dt']}")
            item.setData(Qt.UserRole, r['dt'])
            self.list_exams.addItem(item)
        conn.close()

    def add_exam(self):
        title = self.ex_title.text().strip()
        dt = self.ex_dt.dateTime().toString(Qt.ISODate)
        if not title:
            QMessageBox.warning(self,"Missing","Enter exam title")
            return
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("INSERT INTO exams(title,dt) VALUES(?,?)",(title,dt))
        conn.commit(); conn.close()
        self.load_exams()

    def delete_exam(self):
        row = self.list_exams.currentRow()
        if row < 0:
            QMessageBox.warning(self,"Select","Select an exam")
            return
        txt = self.list_exams.currentItem().text()
        # simple delete by matching text
        dt = self.list_exams.currentItem().data(Qt.UserRole)
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("DELETE FROM exams WHERE dt=? LIMIT 1",(dt,))
        conn.commit(); conn.close()
        self.load_exams()

    def start_countdown(self):
        row = self.list_exams.currentRow()
        if row < 0:
            QMessageBox.warning(self,"Select","Select an exam to countdown")
            return
        dt_iso = self.list_exams.currentItem().data(Qt.UserRole)
        try:
            target = datetime.fromisoformat(dt_iso)
        except Exception:
            QMessageBox.warning(self,"Invalid","Cannot parse exam datetime")
            return
        self.current_countdown_target = target
        self.exam_timer.start(1000)
        QMessageBox.information(self,"Countdown started","Countdown started in the app (label updates).")

    def _tick_countdown(self):
        if not self.current_countdown_target:
            self.exam_timer.stop(); return
        now = datetime.now()
        diff = self.current_countdown_target - now
        if diff.total_seconds() <= 0:
            self.lbl_countdown.setText("EXAM TIME! 🎉")
            self.exam_timer.stop()
            # beep / TTS if available
            if tts_engine:
                tts_engine.say("Exam time! Best of luck.")
                tts_engine.runAndWait()
            return
        days = diff.days
        hours, rem = divmod(diff.seconds,3600)
        mins, secs = divmod(rem,60)
        self.lbl_countdown.setText(f"Countdown: {days}d {hours}h {mins}m {secs}s")

    # -------------------- GPA --------------------
    def load_gpa_table(self):
        self.tbl_gpa.setRowCount(0)
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("SELECT semester,gpa,credits FROM gpa ORDER BY id")
        for r in cur.fetchall():
            row = self.tbl_gpa.rowCount(); self.tbl_gpa.insertRow(row)
            self.tbl_gpa.setItem(row,0,QTableWidgetItem(r['semester']))
            self.tbl_gpa.setItem(row,1,QTableWidgetItem(str(r['gpa'])))
            self.tbl_gpa.setItem(row,2,QTableWidgetItem(str(r['credits'])))
        conn.close()

    def add_semester(self):
        sem = self.in_semester.text().strip()
        try:
            gpa = float(self.in_gpa.text().strip())
            credits = float(self.in_credits.text().strip())
        except Exception:
            QMessageBox.warning(self,"Invalid","GPA and credits must be numeric")
            return
        if not sem:
            QMessageBox.warning(self,"Missing","Provide semester name")
            return
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("INSERT INTO gpa(semester,gpa,credits) VALUES(?,?,?)",(sem,gpa,credits))
        conn.commit(); conn.close()
        self.in_semester.clear(); self.in_gpa.clear(); self.in_credits.clear()
        self.load_gpa_table()

    def calculate_cgpa(self):
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("SELECT gpa,credits FROM gpa")
        total_points = 0.0; total_credits = 0.0
        for r in cur.fetchall():
            total_points += r['gpa'] * r['credits']
            total_credits += r['credits']
        conn.close()
        if total_credits == 0:
            QMessageBox.information(self,"No data","No semesters added yet")
            return
        cgpa = total_points / total_credits
        self.lbl_cgpa.setText(f"CGPA: {cgpa:.3f} (Credits: {int(total_credits)})")

    # -------------------- Attendance --------------------
    def mark_attendance(self):
        course = self.att_course.text().strip()
        date_iso = self.att_date.date().toString(Qt.ISODate)
        present = 1 if self.att_present.currentText()=="Present" else 0
        if not course:
            QMessageBox.warning(self,"Missing","Course required")
            return
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("INSERT INTO attendance(course,date,present) VALUES(?,?,?)",(course,date_iso,present))
        conn.commit(); conn.close()
        QMessageBox.information(self,"Saved","Attendance saved")
        self.load_attendance_table()

    def load_attendance_table(self):
        self.tbl_att.setRowCount(0)
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("SELECT course,date,present FROM attendance ORDER BY date DESC")
        for r in cur.fetchall():
            row = self.tbl_att.rowCount(); self.tbl_att.insertRow(row)
            self.tbl_att.setItem(row,0,QTableWidgetItem(r['course']))
            self.tbl_att.setItem(row,1,QTableWidgetItem(r['date']))
            self.tbl_att.setItem(row,2,QTableWidgetItem("Yes" if r['present'] else "No"))
        conn.close()

    def view_attendance(self):
        course = self.att_course.text().strip()
        if not course:
            QMessageBox.warning(self,"Missing","Enter course to view")
            return
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("SELECT date,present FROM attendance WHERE course=? ORDER BY date DESC",(course,))
        rows = cur.fetchall(); conn.close()
        self.tbl_att.setRowCount(0)
        for r in rows:
            row = self.tbl_att.rowCount(); self.tbl_att.insertRow(row)
            self.tbl_att.setItem(row,0,QTableWidgetItem(course))
            self.tbl_att.setItem(row,1,QTableWidgetItem(r['date']))
            self.tbl_att.setItem(row,2,QTableWidgetItem("Yes" if r['present'] else "No"))

    # -------------------- Notes (Markdown) --------------------
    def load_notes_list(self):
        self.list_notes.clear()
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("SELECT id,title,created FROM notes ORDER BY id DESC")
        for r in cur.fetchall():
            it = QListWidgetItem(f"{r['title']} — {r['created']}")
            it.setData(Qt.UserRole, r['id'])
            self.list_notes.addItem(it)
        conn.close()

    def new_note(self):
        self.note_title.clear(); self.note_body.clear(); self.preview.clear()

    def save_note(self):
        title = self.note_title.text().strip()
        body = self.note_body.toPlainText()
        if not title:
            QMessageBox.warning(self,"Missing","Provide a note title")
            return
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("INSERT INTO notes(title,body,created) VALUES(?,?,?)",(title,body,datetime.now().isoformat()))
        conn.commit(); conn.close()
        self.load_notes_list()
        QMessageBox.information(self,"Saved","Note saved")

    def delete_note(self):
        row = self.list_notes.currentRow()
        if row < 0:
            QMessageBox.warning(self,"Select","Select a note")
            return
        nid = self.list_notes.currentItem().data(Qt.UserRole)
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("DELETE FROM notes WHERE id=?",(nid,))
        conn.commit(); conn.close()
        self.load_notes_list()
        self.note_title.clear(); self.note_body.clear(); self.preview.clear()

    def load_note(self, item):
        nid = item.data(Qt.UserRole)
        conn = get_db_conn(); cur = conn.cursor()
        cur.execute("SELECT title,body FROM notes WHERE id=?",(nid,))
        r = cur.fetchone()
        conn.close()
        if r:
            self.note_title.setText(r['title'])
            self.note_body.setPlainText(r['body'])
            self.preview_markdown()

    def preview_markdown(self):
        text = self.note_body.toPlainText()
        if md_lib:
            try:
                html = md_lib.markdown(text)
            except Exception:
                html = "<pre>" + text + "</pre>"
        else:
            html = "<pre>" + text + "</pre>"
        self.preview.setHtml(html)

    # -------------------- Voice Assistant --------------------
    def do_tts(self):
        text = self.txt_speak.text().strip()
        if not text:
            QMessageBox.warning(self,"Missing","Enter text to speak")
            return
        if tts_engine:
            tts_engine.say(text); tts_engine.runAndWait()
        else:
            QMessageBox.information(self,"TTS not available","Install pyttsx3 to enable TTS")

    def do_stt(self):
        if not sr_recognizer:
            QMessageBox.information(self,"STT not available","Install SpeechRecognition and microphone dependencies")
            return
        try:
            mic = sr.Microphone()
            with mic as source:
                sr_recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = sr_recognizer.listen(source, timeout=5)
            text = sr_recognizer.recognize_google(audio)
            self.stt_result.setText(text)
        except Exception as e:
            self.stt_result.setText(f"Error: {e}")

# Run
if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())
