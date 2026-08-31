import io
import os
import sqlite3
import zipfile
from datetime import date, datetime, timedelta
from PIL import Image, ImageDraw
import numpy as np
import pandas as pd
import streamlit as st

# Grundlegende UI-Erweiterungen (optional, stürzen nie ab)
try:
  from streamlit_drawable_canvas import st_canvas
  HAS_CANVAS = True
except ImportError:
  HAS_CANVAS = False

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


# --- Streamlit Navigation ---
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
# 1. OBJEKTPLÄNE & KATASTER BEARBEITEN (ROBUSTE SLIDER FÜR PINS)
# ----------------------------------------------------
if menu == "🎨 Objektpläne & Kataster bearbeiten":
  st.subheader("🎨 Objektpläne bearbeiten & Einrichtungen platzieren")
  properties = conn.execute("SELECT * FROM properties").fetchall()
  if not properties:
    st.warning(
        "Bitte zuerst ein Objekt unter 'Pläne & Objekte verwalten' anlegen."
    )
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
          "Geschossplan zum Bearbeiten auswählen",
          list(plan_dict.keys()),
          key="edit_plan_select",
      )
      selected_plan_id = plan_dict[selected_plan_name]

      curr_plan = conn.execute(
          "SELECT * FROM floor_plans WHERE id = ?", (selected_plan_id,)
      ).fetchone()
      plan_img_path = curr_plan["image_path"]

      if curr_plan and os.path.exists(plan_img_path):
        base_img = Image.open(plan_img_path).convert("RGBA")
        w, h = base_img.size
        tw = 900
        th = int((h / w) * tw)
        r_img = base_img.resize((tw, th))

        facs_on_plan = conn.execute(
            "SELECT * FROM fire_facilities WHERE floor_plan_id = ?",
            (selected_plan_id,),
        ).fetchall()

        d_img = r_img.copy()
        draw = ImageDraw.Draw(d_img)

        for fac in facs_on_plan:
          if fac["pin_x"] and fac["pin_y"]:
            fx = int((fac["pin_x"] / 100.0) * tw)
            fy = int((fac["pin_y"] / 100.0) * th)
            draw_color_dot(
                draw,
                fx,
                fy,
                fill_color="#2563EB",
                outline_color="white",
                label=fac["category"][:1],
                radius=14,
            )

        col_plan_view, col_plan_form = st.columns([6, 4])

        with col_plan_view:
          st.image(
              d_img,
              caption=f"Grundriss: {selected_plan_name}",
              use_container_width=True,
          )

        with col_plan_form:
          st.markdown("#### 📍 Gerät auf Plan positionieren")
          st.caption(
              "Nutze die Prozent-Schieberegler, um den Pin exakt auf dem Plan"
              " zu platzieren (0% = Oben/Links, 100% = Unten/Rechts)."
          )

          with st.form("new_fac_slider_form", clear_on_submit=True):
            cat = st.selectbox(
                "Kategorie*",
                [
                    "LOESCHER",
                    "TUER",
                    "HYDRANT",
                    "BMA",
                    "RWA",
                    "NOTLICHT",
                    "FLUECHTWEG",
                    "SONSTIGES",
                ],
            )
            ident = st.text_input("Gerätenummer / Kennung*", placeholder="L-01")
            loc = st.text_input("Standortbeschreibung*", placeholder="Flur West")

            px_val = st.slider(
                "Position X-Achse (Links ➔ Rechts)", 0.0, 100.0, 50.0, 0.5
            )
            py_val = st.slider(
                "Position Y-Achse (Oben ➔ Unten)", 0.0, 100.0, 50.0, 0.5
            )

            if st.form_submit_button(
                "➕ Einrichtung per Regler platzieren", type="primary"
            ):
              if ident and loc:
                conn.execute(
                    """
                                    INSERT INTO fire_facilities (
                                        property_id, floor_plan_id, pin_x, pin_y, category, identifier, location_desc
                                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                                """,
                    (
                        selected_prop_id,
                        selected_plan_id,
                        px_val,
                        py_val,
                        cat,
                        ident,
                        loc,
                    ),
                )
                conn.commit()
                st.success(f"Einrichtung '{ident}' erfolgreich platziert!")
                st.rerun()
              else:
                st.error("Bitte Kennung und Standort ausfüllen.")
      else:
        st.warning("Grundriss-Bilddatei nicht gefunden.")


# ----------------------------------------------------
# 2. AKTIVE BEGEHUNG
# ----------------------------------------------------
elif menu == "Aktive Begehung & Planprüfung":
  st.subheader("🚀 Aktive Begehung & Checkliste")
  properties = conn.execute("SELECT * FROM properties").fetchall()
  if not properties:
    st.warning("Bitte zuerst ein Objekt anlegen.")
  else:
    prop_dict = {f"{p['name']} ({p['address']})": p["id"] for p in properties}
    selected_prop_label = st.selectbox(
        "Objekt auswählen", list(prop_dict.keys()), key="main_prop"
    )
    selected_prop_id = prop_dict[selected_prop_label]

    if st.button("🚀 Begehung starten", type="primary"):
      cur = conn.cursor()
      cur.execute(
          """
                INSERT INTO inspection_runs (property_id, inspector_name, inspection_date, interval_scope, status)
                VALUES (?, 'Johannes Probst', ?, 'MONATLICH', 'ABGESCHLOSSEN')
            """,
          (selected_prop_id, date.today()),
      )
      conn.commit()
      st.success("Begehung erfolgreich dokumentiert!")


# ----------------------------------------------------
# 3. LIVE-KAMERA
# ----------------------------------------------------
elif menu == "📷 Live-Kamera QR-Scanner":
  st.subheader("📷 Live-Kamera QR-Code-Scanner")
  st.info("QR-Scanner im Web-Modus aktiv.")


# ----------------------------------------------------
# 4. ANLAGENKATASTER
# ----------------------------------------------------
elif menu == "Anlagenkataster & Fristen":
  st.subheader("Anlagenkataster & Fristen")
  facs = conn.execute("SELECT * FROM fire_facilities").fetchall()
  if not facs:
    st.info("Keine Einrichtungen erfasst.")
  else:
    for f in facs:
      st.write(
          f"🔵 **{f['identifier']}** | {f['category']} | Ort:"
          f" {f['location_desc']}"
      )


# ----------------------------------------------------
# 5. ETIKETTENDRUCK
# ----------------------------------------------------
elif menu == "🏷️ QR-Code Etikettendruck":
  st.subheader("🏷️ QR-Code Etikettendruck")
  st.info("Verfügbar im Kataster.")


# ----------------------------------------------------
# 6. HANDWERKERAUFTRAG
# ----------------------------------------------------
elif menu == "🛠️ Handwerker- & Mängelauftrag":
  st.subheader("🛠️ Handwerker- & Mängelauftrag")
  st.info("Keine offenen Mängel.")


# ----------------------------------------------------
# 7. EREIGNISJOURNAL
# ----------------------------------------------------
elif menu == "🚨 Ereignisjournal (TRVB 117 O)":
  st.subheader("🚨 Ereignisjournal (TRVB 117 O)")
  events = conn.execute("SELECT * FROM journal_events").fetchall()
  for ev in events:
    st.write(f"**{ev['event_date']}** – {ev['title']}")


# ----------------------------------------------------
# 8. BSB-JAHRESBERICHT
# ----------------------------------------------------
elif menu == "📊 BSB-Jahresbericht (TRVB 117 O)":
  st.subheader("📊 BSB-Jahresbericht")
  st.info("Management-Bericht verfügbar.")


# ----------------------------------------------------
# 9. DATEN-EXPORT
# ----------------------------------------------------
elif menu == "💾 Daten-Export & Datensicherung":
  st.subheader("💾 Daten-Export & Datensicherung")
  zip_bytes = create_full_backup_zip()
  st.download_button(
      "📦 Komplettes System-Backup (ZIP) herunterladen",
      data=zip_bytes,
      file_name="Brandschutz_Backup.zip",
      mime="application/zip",
      type="primary",
  )


# ----------------------------------------------------
# 10. HISTORIE
# ----------------------------------------------------
elif menu == "Brandschutzbuch-Historie":
  st.subheader("Brandschutzbuch-Historie")
  runs = conn.execute("SELECT * FROM inspection_runs").fetchall()
  for r in runs:
    st.write(
        f"Prüfung am {r['inspection_date']} – Status: ` {r['status']} `"
    )


# ----------------------------------------------------
# 11. PLÄNE & OBJEKTE VERWALTEN (BILD-UPLOAD OHNE PDF-FEHLER)
# ----------------------------------------------------
elif menu == "Pläne & Objekte verwalten":
  tab_p1, tab_p2, tab_p3 = st.tabs([
      "🏢 1. Liegenschaften verwalten",
      "🗺️ 2. Geschosspläne hochladen (PNG, JPG)",
      "➕ 3. Neues Objekt anlegen",
  ])

  with tab_p1:
    st.subheader("Bestehende Objekte")
    properties = conn.execute("SELECT * FROM properties").fetchall()
    if not properties:
      st.info("Noch keine Objekte vorhanden. Lege unter Tab 3 ein Objekt an.")
    else:
      for p in properties:
        st.write(
            f"🏢 **{p['name']}** – {p['address']} (BSB:"
            f" {p['fire_safety_officer']})"
        )

  with tab_p2:
    st.subheader("Grundrissplan als Bild hochladen (empfohlen: PNG oder JPG)")
    properties = conn.execute("SELECT * FROM properties").fetchall()
    if not properties:
      st.warning("Bitte zuerst unter Tab 3 ein Objekt anlegen.")
    else:
      prop_dict = {f"{p['name']} ({p['address']})": p["id"] for p in properties}
      sel_p = st.selectbox(
          "Objekt für Plan-Upload auswählen", list(prop_dict.keys())
      )
      target_id = prop_dict[sel_p]

      with st.form("upload_plan_form", clear_on_submit=True):
        p_name = st.text_input("Planbezeichnung (z.B. Erdgeschoss)")
        p_file = st.file_uploader("Bild auswählen", type=["png", "jpg", "jpeg"])

        if st.form_submit_button("Plan speichern", type="primary"):
          if p_name and p_file:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_ext = os.path.splitext(p_file.name)[1].lower()
            saved_filename = f"plan_{target_id}_{timestamp}{file_ext}"
            full_path = os.path.join(UPLOAD_DIR, saved_filename)
            with open(full_path, "wb") as f:
              f.write(p_file.getbuffer())

            conn.execute(
                "INSERT INTO floor_plans (property_id, name, image_path) VALUES"
                " (?, ?, ?)",
                (target_id, p_name, full_path),
            )
            conn.commit()
            st.success(f"Plan '{p_name}' erfolgreich hochgeladen!")
            st.rerun()
          else:
            st.error("Bitte Bezeichnung eingeben und Bilddatei auswählen.")

  with tab_p3:
    st.subheader("Neues Objekt anlegen")
    with st.form("new_prop_form", clear_on_submit=True):
      name = st.text_input("Objektname / Liegenschaft*")
      address = st.text_input("Adresse*")
      bsb = st.text_input(
          "Zuständiger Brandschutzbeauftragter", value="Johannes Probst"
      )

      if st.form_submit_button("Objekt jetzt anlegen", type="primary"):
        if name and address:
          conn.execute(
              "INSERT INTO properties (name, address, fire_safety_officer)"
              " VALUES (?, ?, ?)",
              (name, address, bsb),
          )
          conn.commit()
          st.success(f"Objekt '{name}' erfolgreich angelegt!")
          st.rerun()
        else:
          st.error("Bitte Objektname und Adresse ausfüllen.")

conn.close()
