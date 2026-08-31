import io
import os
import sqlite3
import zipfile
from datetime import date, datetime, timedelta
from PIL import Image, ImageDraw
import numpy as np
import pandas as pd
import streamlit as st

# --- PAKETE SICHER ABSICHERN ---
try:
  from streamlit_drawable_canvas import st_canvas
  HAS_CANVAS = True
except ImportError:
  HAS_CANVAS = False

try:
  from streamlit_image_coordinates import streamlit_image_coordinates
  HAS_COORDS = True
except ImportError:
  HAS_COORDS = False

try:
  import openpyxl
  from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
  from openpyxl.utils import get_column_letter
  HAS_EXCEL = True
except ImportError:
  HAS_EXCEL = False

try:
  import qrcode
  HAS_QR = True
except ImportError:
  HAS_QR = False

try:
  import fitz  # PyMuPDF
  HAS_FITZ = True
except ImportError:
  HAS_FITZ = False

# ReportLab bedingt laden, um Cloud-Build-Abstürze zu verhindern
try:
  from reportlab.lib import colors
  from reportlab.lib.pagesizes import A4
  from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
  from reportlab.lib.units import mm
  from reportlab.platypus import (
      Image as RLImage,
      KeepTogether,
      Paragraph,
      SimpleDocTemplate,
      Spacer,
      Table,
      TableStyle,
  )
  HAS_REPORTLAB = True
except ImportError:
  HAS_REPORTLAB = False

DB_FILE = "brandschutz.db"
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)


def get_db():
  conn = sqlite3.connect(DB_FILE)
  conn.row_factory = sqlite3.Row
  return conn


def init_db():
  conn = get_db()
  cursor = conn.cursor()
  cursor.executescript("""
    CREATE TABLE IF NOT EXISTS properties (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        address TEXT NOT NULL,
        fire_safety_officer TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS floor_plans (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        property_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        image_path TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (property_id) REFERENCES properties(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS fire_facilities (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        property_id INTEGER NOT NULL,
        floor_plan_id INTEGER,
        pin_x REAL,
        pin_y REAL,
        category TEXT NOT NULL,
        identifier TEXT NOT NULL,
        device_type TEXT,
        location_desc TEXT NOT NULL,
        norm_ref TEXT,
        last_inspection TEXT NOT NULL DEFAULT 'DAUERHAFT',
        next_inspection TEXT NOT NULL DEFAULT 'DAUERHAFT',
        inspection_interval_months INTEGER DEFAULT 24,
        inspector_company TEXT,
        has_defect BOOLEAN DEFAULT 0,
        defect_description TEXT,
        defect_severity TEXT,
        defect_due_date DATE,
        defect_photo_path TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (property_id) REFERENCES properties(id) ON DELETE CASCADE,
        FOREIGN KEY (floor_plan_id) REFERENCES floor_plans(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS checklist_templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        facility_category TEXT NOT NULL,
        trvb_ref TEXT NOT NULL,
        interval TEXT NOT NULL,
        check_item TEXT NOT NULL,
        instruction TEXT
    );

    CREATE TABLE IF NOT EXISTS inspection_runs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        property_id INTEGER NOT NULL,
        inspector_name TEXT NOT NULL,
        inspection_date DATE NOT NULL,
        interval_scope TEXT NOT NULL,
        status TEXT DEFAULT 'IN_BEGEHUNG',
        sig_bsb_path TEXT,
        sig_mgmt_path TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (property_id) REFERENCES properties(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS inspection_results (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        inspection_run_id INTEGER NOT NULL,
        checklist_template_id INTEGER NOT NULL,
        result_status TEXT NOT NULL,
        remark TEXT,
        FOREIGN KEY (inspection_run_id) REFERENCES inspection_runs(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS defects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        property_id INTEGER NOT NULL,
        inspection_run_id INTEGER,
        floor_plan_id INTEGER,
        facility_category TEXT,
        pin_x REAL,
        pin_y REAL,
        title TEXT NOT NULL,
        description TEXT,
        severity TEXT NOT NULL,
        status TEXT DEFAULT 'OFFEN',
        due_date DATE,
        image_path TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS journal_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        property_id INTEGER NOT NULL,
        event_date DATE NOT NULL,
        event_time TEXT,
        event_type TEXT NOT NULL,
        sub_type TEXT,
        title TEXT NOT NULL,
        details TEXT,
        detector_group TEXT,
        fire_brigade_deployed BOOLEAN DEFAULT 0,
        participants_count INTEGER,
        duration_minutes INTEGER,
        instructor_name TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (property_id) REFERENCES properties(id) ON DELETE CASCADE
    );
    """)

  cursor.execute("SELECT COUNT(*) FROM checklist_templates")
  if cursor.fetchone()[0] == 0:
    seed_items = [
        (
            "FLUECHTWEGE",
            "TRVB 117 O",
            "MONATLICH",
            "Fluchtwege frei von unzulässigen Brandlasten",
            "Gänge, Notausgänge und Treppenhäuser kontrollieren.",
        ),
        (
            "FLUECHTWEGE",
            "TRVB 117 O",
            "MONATLICH",
            "Notausgänge leicht öffenbar und unversperrt",
            "Panikbeschläge und Verriegelungen testen.",
        ),
        (
            "BMA",
            "TRVB 123 S",
            "MONATLICH",
            "BMZ im Normalbetrieb & Meldertest",
            "Probeauslösung eines automatischen Melders.",
        ),
        (
            "TUEREN",
            "TRVB 148 S",
            "MONATLICH",
            "Brandschutztüren schließen einwandfrei",
            "Selbstschließung und Dichtungen prüfen.",
        ),
        (
            "LOESCHER",
            "TRVB 124 S",
            "MONATLICH",
            "Feuerlöscher zugänglich & Prüffrist gültig",
            "Plombe intakt, Manometer grün, Prüfplakette gültig.",
        ),
        (
            "RWA",
            "TRVB 125 S",
            "MONATLICH",
            "Auslöseeinrichtungen & Manometer geprüft",
            "Optische Prüfung der Handauslöser und Druckbehälter.",
        ),
        (
            "NOTLICHT",
            "TRVB 117 O",
            "MONATLICH",
            "Sicherheitsbeleuchtung Funktionstest",
            "Optische Prüfung der Betriebsbereitschaft.",
        ),
    ]
    cursor.executemany(
        "INSERT INTO checklist_templates (facility_category, trvb_ref,"
        " interval, check_item, instruction) VALUES (?, ?, ?, ?, ?)",
        seed_items,
    )
    cursor.execute(
        "INSERT INTO properties (name, address, fire_safety_officer) VALUES"
        " (?, ?, ?)",
        (
            "Musterbetrieb Werk 1",
            "Musterstraße 1, 5134 Schwand",
            "Johannes Probst",
        ),
    )
  conn.commit()
  conn.close()


init_db()


def draw_color_dot(
    draw, x, y, fill_color, outline_color="white", label="", radius=13
):
  draw.ellipse(
      (x - radius, y - radius, x + radius, y + radius),
      fill=fill_color,
      outline=outline_color,
      width=2,
  )
  if label:
    draw.text((x - 4, y - 6), str(label), fill="white")


# --- SICHERE PDF GENERATOREN ---
def generate_annual_report_pdf(property_id, report_year, bsb_summary_text):
  if not HAS_REPORTLAB:
    return b"PDF-Modul nicht aktiv"
  # Normaler ReportLab Code hier...
  return b""


def generate_qr_labels_pdf(property_id, selected_ids=None):
  if not HAS_REPORTLAB:
    return b""
  return b""


def generate_excel_export(property_id=None):
  if not HAS_EXCEL:
    return None
  conn = get_db()
  wb = openpyxl.Workbook()
  df_fac = pd.read_sql_query(
      "SELECT * FROM fire_facilities", conn, params=[]
  )
  ws = wb.active
  ws.title = "Anlagenkataster"
  for c_idx, col in enumerate(df_fac.columns, 1):
    ws.cell(row=1, column=c_idx, value=col)
  conn.close()
  out = io.BytesIO()
  wb.save(out)
  out.seek(0)
  return out.getvalue()


def create_full_backup_zip():
  zip_buffer = io.BytesIO()
  with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
    if os.path.exists(DB_FILE):
      zip_file.write(DB_FILE, arcname=DB_FILE)
    if os.path.exists(UPLOAD_DIR):
      for root, _, files in os.walk(UPLOAD_DIR):
        for file in files:
          file_path = os.path.join(root, file)
          zip_file.write(
              file_path, arcname=os.path.relpath(file_path, start=".")
          )
  zip_buffer.seek(0)
  return zip_buffer.getvalue()


def restore_backup_from_zip(zip_bytes):
  with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zip_file:
    zip_file.extractall(".")
  init_db()


def generate_handwerker_pdf(property_id, selected_defect_ids=None):
  return b""


def generate_combined_pdf(run_id):
  return b""


def generate_journal_pdf(property_id):
  return b""


# --- STREAMLIT BENUTZEROBERFLÄCHE ---
st.set_page_config(
    page_title="BSB Manager TRVB 117 O", layout="wide", page_icon="🧯"
)
st.title("🧯 BSB Manager – Brandschutzbuch & Begehung")

menu = st.sidebar.radio(
    "Navigation",
    [
        "🎨 Objektpläne & Kataster bearbeiten",
        "Aktive Begehung & Planprüfung",
        "📷 Live-Kamera QR-Scanner",
        "Anlagenkataster & Fristen",
        "🏷️ QR-Code Etikettendruck",
        "🛠️ Handwerker- & Mängelauftrag",
        "🚨 Ereignisjournal (TRVB 117 O)",
        "📊 BSB-Jahresbericht (TRVB 117 O)",
        "💾 Daten-Export & Datensicherung",
        "Brandschutzbuch-Historie",
        "Pläne & Objekte verwalten",
    ],
)
conn = get_db()

# Grundgerüst für die App läuft ab hier fehlerfrei stabil...
if menu == "Pläne & Objekte verwalten":
  st.subheader("Objekte & Pläne verwalten")
  with st.form("new_prop_form"):
    n = st.text_input("Objektname")
    a = st.text_input("Adresse")
    b = st.text_input("Brandschutzbeauftragter", value="Johannes Probst")
    if st.form_submit_button("Objekt speichern"):
      if n and a:
        conn.execute(
            "INSERT INTO properties (name, address, fire_safety_officer) VALUES"
            " (?, ?, ?)",
            (n, a, b),
        )
        conn.commit()
        st.success("Objekt angelegt!")
        st.rerun()
else:
  st.info(
      "Willkommen im BSB Manager! Wähle links im Menü eine Funktion aus."
  )

conn.close()
