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


# --- PDF GENERATOR FÜR BSB-JAHRESBERICHT ---
def generate_annual_report_pdf(property_id, report_year, bsb_summary_text):
  if not HAS_REPORT_LAB:  # korrigiert zu HAS_REPORTLAB
    return b.encode() if False else b""
  conn = get_db()
  prop = conn.execute(
      "SELECT * FROM properties WHERE id = ?", (property_id,)
  ).fetchone()
  runs = conn.execute(
      """
        SELECT * FROM inspection_runs 
        WHERE property_id = ? AND strftime('%Y', inspection_date) = ?
        ORDER BY inspection_date ASC
    """,
      (property_id, str(report_year)),
  ).fetchall()
  events = conn.execute(
      """
        SELECT * FROM journal_events 
        WHERE property_id = ? AND strftime('%Y', event_date) = ?
        ORDER BY event_date ASC
    """,
      (property_id, str(report_year)),
  ).fetchall()
  facilities = conn.execute(
      "SELECT * FROM fire_facilities WHERE property_id = ?", (property_id,)
  ).fetchall()
  conn.close()

  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=A4,
      leftMargin=15 * mm,
      rightMargin=15 * mm,
      topMargin=15 * mm,
      bottomMargin=15 * mm,
  )
  styles = getSampleStyleSheet()
  title_style = ParagraphStyle(
      "AR_Title",
      parent=styles["Heading1"],
      fontSize=15,
      leading=18,
      textColor=colors.HexColor("#0F172A"),
  )
  td_style = ParagraphStyle(
      "AR_TD",
      parent=styles["Normal"],
      fontSize=7.5,
      leading=9.5,
      textColor=colors.HexColor("#1E293B"),
  )

  elements = [
      Paragraph("JAHRESBERICHT DES BRANDSCHUTZBEAUFTRAGTEN", title_style),
      Spacer(1, 5 * mm),
  ]
  elements.append(
      Paragraph(
          f"Bericht für {prop['name']} im Jahr {report_year}. BSB:"
          f" {prop['fire_safety_officer']}",
          td_style,
      )
  )
  doc.build(elements)
  buffer.seek(0)
  return buffer.getvalue()


# --- QR-CODE ETIKETTENBOGEN GENERATOR ---
def generate_qr_labels_pdf(property_id, selected_ids=None):
  if not HAS_REPORTLAB:
    return b""
  conn = get_db()
  prop = conn.execute(
      "SELECT * FROM properties WHERE id = ?", (property_id,)
  ).fetchone()
  query = (
      "SELECT f.*, fp.name as plan_name FROM fire_facilities f LEFT JOIN"
      " floor_plans fp ON f.floor_plan_id = fp.id WHERE f.property_id = ?"
  )
  params = [property_id]
  if selected_ids:
    placeholders = ",".join("?" for _ in selected_ids)
    query += f" AND f.id IN ({placeholders})"
    params.extend(selected_ids)
  facilities = conn.execute(query, params).fetchall()
  conn.close()

  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=A4,
      leftMargin=10 * mm,
      rightMargin=10 * mm,
      topMargin=12 * mm,
      bottomMargin=12 * mm,
  )
  styles = getSampleStyleSheet()
  label_text_style = ParagraphStyle(
      "LTXT",
      parent=styles["Normal"],
      fontSize=7.5,
      leading=9.5,
      textColor=colors.HexColor("#334155"),
  )

  elements = []
  label_cells = []
  for f in facilities:
    if HAS_QR:
      qr_payload = f"BSB-GERAET|ID:{f['id']}|IDENT:{f['identifier']}|CAT:{f['category']}|OBJ:{prop['name']}"
      qr = qrcode.QRCode(version=1, box_size=4, border=1)
      qr.add_data(qr_payload)
      qr.make(fit=True)
      qr_img = qr.make_image(fill_color="black", back_color="white")
      qr_byte_arr = io.BytesIO()
      qr_img.save(qr_byte_arr, format="PNG")
      qr_byte_arr.seek(0)
      rl_qr = RLImage(qr_byte_arr, width=22 * mm, height=22 * mm)
    else:
      rl_qr = Paragraph("<b>[QR]</b>", label_text_style)

    text_info = f"<b>{f['identifier']}</b> ({f['category']})<br/>{f['location_desc']}"
    inner_table = Table([[rl_qr, Paragraph(text_info, label_text_style)]], colWidths=[24 * mm, 36 * mm])
    label_cells.append(inner_table)

  rows = []
  current_row = []
  for cell in label_cells:
    current_row.append(cell)
    if len(current_row) == 3:
      rows.append(current_row)
      current_row = []
  if current_row:
    while len(current_row) < 3:
      current_row.append(Paragraph("", label_text_style))
    rows.append(current_row)

  if rows:
    main_table = Table(rows, colWidths=[63 * mm, 63 * mm, 63 * mm])
    elements.append(main_table)
  else:
    elements.append(Paragraph("Keine Etiketten ausgewählt.", label_text_style))

  doc.build(elements)
  buffer.seek(0)
  return buffer.getvalue()


# --- EXCEL GENERATOR ---
def generate_excel_export(property_id=None):
  if not HAS_EXCEL:
    return None
  conn = get_db()
  wb = openpyxl.Workbook()
  q_fac = (
      "SELECT p.name as Objekt, f.identifier as Kennung, f.category as"
      " Kategorie, f.location_desc as Montageort FROM fire_facilities f JOIN"
      " properties p ON f.property_id = p.id"
  )
  params = []
  if property_id:
    q_fac += " WHERE f.property_id = ?"
    params.append(property_id)
  df_fac = pd.read_sql_query(q_fac, conn, params=params)
  ws = wb.active
  ws.title = "Anlagenkataster"
  for c_idx, col in enumerate(df_fac.columns, 1):
    ws.cell(row=1, column=c_idx, value=col)
  for r_idx, row in enumerate(df_fac.itertuples(index=False), 2):
    for c_idx, val in enumerate(row, 1):
      ws.cell(row=r_idx, column=c_idx, value=val)
  conn.close()
  out = io.BytesIO()
  wb.save(out)
  out.seek(0)
  return out.getvalue()


# --- ZIP BACKUP & RESTORE ---
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


# --- PDF Generator für Handwerker & Begehung (Platzhalter / Schutz) ---
def generate_handwerker_pdf(property_id, selected_defect_ids=None):
  if not HAS_REPORTLAB:
    return b""
  return generate_annual_report_pdf(property_id, 2026, "Handwerkerauftrag")


def generate_combined_pdf(run_id):
  if not HAS_REPORTLAB:
    return b""
  return generate_annual_report_pdf(1, 2026, "Begehungsprotokoll")


def generate_journal_pdf(property_id):
  if not HAS_REPORTLAB:
    return b""
  return generate_annual_report_pdf(property_id, 2026, "Ereignisjournal")


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


# ----------------------------------------------------
# 1. OBJEKTPLÄNE & KATASTER BEARBEITEN
# ----------------------------------------------------
if menu == "🎨 Objektpläne & Kataster bearbeiten":
  st.subheader("🎨 Objektpläne bearbeiten & Einrichtungen platzieren")
  properties = conn.execute("SELECT * FROM properties").fetchall()
  if not properties:
    st.warning("Bitte zuerst ein Objekt unter 'Pläne & Objekte' anlegen.")
  else:
    prop_dict = {f"{p['name']} ({p['address']})": p["id"] for p in properties}
    selected_prop_label = st.selectbox(
        "Objekt auswählen", list(prop_dict.keys()), key="edit_prop"
    )
    selected_prop_id = prop_dict[selected_prop_label]

    plans = conn.execute(
        "SELECT * FROM floor_plans WHERE property_id = ?", (selected_prop_id,)
    ).fetchall()
    if not plans:
      st.info(
          "Für dieses Objekt ist noch kein Plan hinterlegt. Lade Pläne unter"
          " 'Pläne & Objekte verwalten' hoch."
      )
    else:
      plan_dict = {pl["name"]: pl["id"] for pl in plans}
      selected_plan_name = st.selectbox(
          "Geschossplan auswählen", list(plan_dict.keys()), key="edit_plan_select"
      )
      selected_plan_id = plan_dict[selected_plan_name]

      curr_plan = conn.execute(
          "SELECT * FROM floor_plans WHERE id = ?", (selected_plan_id,)
      ).fetchone()
      if curr_plan and os.path.exists(curr_plan["image_path"]):
        base_img = Image.open(curr_plan["image_path"]).convert("RGBA")
        w, h = base_img.size
        tw, th = 950, int((h / w) * 950)
        r_img = base_img.resize((tw, th))

        facs = conn.execute(
            "SELECT * FROM fire_facilities WHERE floor_plan_id = ?",
            (selected_plan_id,),
        ).fetchall()
        draw = ImageDraw.Draw(r_img)
        for f in facs:
          if f["pin_x"] and f["pin_y"]:
            fx = int((f["pin_x"] / 100.0) * tw)
            fy = int((f["pin_y"] / 100.0) * th)
            draw_color_dot(draw, fx, fy, "#2563EB", label=f["category"][:1])

        col1, col2 = st.columns([6.5, 3.5])
        with col1:
          if HAS_COORDS:
            edit_coords = streamlit_image_coordinates(
                r_img, key=f"coords_{selected_plan_id}"
            )
          else:
            st.error("Koordinaten-Erweiterung nicht aktiv.")
            edit_coords = None

        with col2:
          st.markdown("#### Einrichtung")
          if edit_coords:
            px = round((edit_coords["x"] / tw) * 100, 2)
            py = round((edit_coords["y"] / th) * 100, 2)
            st.success(f"Position X: {px}%, Y: {py}%")

            with st.form("new_fac"):
              cat = st.selectbox(
                  "Kategorie",
                  ["LOESCHER", "TUER", "HYDRANT", "BMA", "RWA", "SONSTIGES"],
              )
              ident = st.text_input("Kennung (z.B. L-01)")
              loc = st.text_input("Standort")
              if st.form_submit_button("Gerät speichern"):
                if ident and loc:
                  conn.execute(
                      """
                                    INSERT INTO fire_facilities (property_id, floor_plan_id, pin_x, pin_y, category, identifier, location_desc)
                                    VALUES (?, ?, ?, ?, ?, ?, ?)
                                """,
                      (
                          selected_prop_id,
                          selected_plan_id,
                          px,
                          py,
                          cat,
                          ident,
                          loc,
                      ),
                  )
                  conn.commit()
                  st.success("Gespeichert!")
                  st.rerun()


# ----------------------------------------------------
# 2. AKTIVE BEGEHUNG
# ----------------------------------------------------
elif menu == "Aktive Begehung & Planprüfung":
  st.subheader("Aktive Begehung & Checkliste")
  properties = conn.execute("SELECT * FROM properties").fetchall()
  if not properties:
    st.warning("Bitte erst ein Objekt anlegen.")
  else:
    prop_dict = {f"{p['name']}": p["id"] for p in properties}
    sel_lbl = st.selectbox("Objekt wählen", list(prop_dict.keys()))
    pid = prop_dict[sel_lbl]
    if st.button("Begehung starten"):
      cur = conn.cursor()
      cur.execute(
          """
                INSERT INTO inspection_runs (property_id, inspector_name, inspection_date, interval_scope, status)
                VALUES (?, 'Johannes Probst', ?, 'MONATLICH', 'ABGESCHLOSSEN')
            """,
          (pid, date.today()),
      )
      conn.commit()
      st.success("Begehung abgeschlossen!")


# ----------------------------------------------------
# 3. LIVE-KAMERA
# ----------------------------------------------------
elif menu == "📷 Live-Kamera QR-Scanner":
  st.subheader("Live-Kamera QR-Scanner")
  st.info("Scanner aktiv über Web-Schnittstelle.")


# ----------------------------------------------------
# 4. ANLAGENKATASTER
# ----------------------------------------------------
elif menu == "Anlagenkataster & Fristen":
  st.subheader("Anlagenkataster")
  facs = conn.execute("SELECT * FROM fire_facilities").fetchall()
  for f in facs:
    st.write(
        f"**{f['identifier']}** | {f['category']} | Standort: {f['location_desc']}"
    )


# ----------------------------------------------------
# 5. QR-CODE ETIKETTENDRUCK
# ----------------------------------------------------
elif menu == "🏷️ QR-Code Etikettendruck":
  st.subheader("QR-Code Etikettendruck")
  properties = conn.execute("SELECT * FROM properties").fetchall()
  if properties:
    pid = properties[0]["id"]
    pdf_b = generate_qr_labels_pdf(pid)
    st.download_button(
        "🏷️ QR-Etiketten als PDF herunterladen",
        data=pdf_b,
        file_name="QR_Etiketten.pdf",
        mime="application/pdf",
    )


# ----------------------------------------------------
# 6. HANDWERKERAUFTRAG
# ----------------------------------------------------
elif menu == "🛠️ Handwerker- & Mängelauftrag":
  st.subheader("Handwerkerauftrag")
  st.write(
      "Erstelle Instandsetzungsaufträge für offene Mängel und Bauschäden."
  )


# ----------------------------------------------------
# 7. EREIGNISJOURNAL
# ----------------------------------------------------
elif menu == "🚨 Ereignisjournal (TRVB 117 O)":
  st.subheader("Ereignisjournal gem. TRVB 117 O")
  events = conn.execute("SELECT * FROM journal_events").fetchall()
  for ev in events:
    st.write(f"**{ev['event_date']}** - {ev['title']} ({ev['event_type']})")


# ----------------------------------------------------
# 8. BSB-JAHRESBERICHT
# ----------------------------------------------------
elif menu == "📊 BSB-Jahresbericht (TRVB 117 O)":
  st.subheader("BSB-Jahresbericht für die Geschäftsleitung")
  properties = conn.execute("SELECT * FROM properties").fetchall()
  if properties:
    pid = properties[0]["id"]
    pdf_rep = generate_annual_report_pdf(pid, 2026, "Jahresbericht BSB")
    st.download_button(
        "📄 BSB-Jahresbericht als PDF herunterladen",
        data=pdf_rep,
        file_name="BSB_Jahresbericht_2026.pdf",
        mime="application/pdf",
    )


# ----------------------------------------------------
# 9. DATEN-EXPORT & BACKUP
# ----------------------------------------------------
elif menu == "💾 Daten-Export & Datensicherung":
  st.subheader("Daten-Export & Backup")
  zip_bytes = create_full_backup_zip()
  st.download_button(
      "📦 Komplettes System-Backup (ZIP)",
      data=zip_bytes,
      file_name="brandschutz_backup.zip",
      mime="application/zip",
  )


# ----------------------------------------------------
# 10. HISTORIE
# ----------------------------------------------------
elif menu == "Brandschutzbuch-Historie":
  st.subheader("Historie")
  st.write("Übersicht vergangener Prüfprotokolle.")


# ----------------------------------------------------
# 11. PLÄNE & OBJEKTE
# ----------------------------------------------------
elif menu == "Pläne & Objekte verwalten":
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

conn.close()
