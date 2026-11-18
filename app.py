import io
import os
import re

import pandas as pd
import pymysql
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="BGV Master Tracker - Rupeek", layout="wide")

QUERY = """
SELECT
  ecm.case_ars_no AS `ARS Number`,
  ecff.case_flex_field1 AS `Emp id`,
  CONCAT(ecc1.first_name, ' ', IFNULL(ecc1.middle_name, ''), ' ', IFNULL(ecc1.last_name, '')) AS `Emp Name`,
  fn_case_status(ecm.case_status) AS `BGV Status`,

  (SELECT fn_check_status(check_status) FROM ec_case_checks ecc
   WHERE ecc.case_id=ecm.case_id AND check_status <>9 LIMIT 1) AS `Check_Status`,

  (
    SELECT GROUP_CONCAT(CONCAT(ecc2.check_name, ' -> ', ec.check_ops_name) ORDER BY ecc2.check_name SEPARATOR ',')
    FROM ec_case_checks ecc2
    LEFT JOIN ec_checks ec ON ecc2.check_id = ec.check_id
    WHERE ecm.case_id = ecc2.case_id
      AND ecc2.check_id IN (235,57,101,25,144,132,61,275,274,72,1,41)
      AND ecc2.check_status <> 9
  ) AS `dept`,

  (
    SELECT GROUP_CONCAT(CONCAT(ecc2.check_name, ' -> ', ecc2.check_severity) ORDER BY ecc2.check_name SEPARATOR ',')
    FROM ec_case_checks ecc2
    WHERE ecm.case_id = ecc2.case_id
      AND ecc2.check_id IN (235,57,101,25,144,132,61,275,274,72,1,41)
      AND ecc2.check_status <> 9
  ) AS `Colour code`,

  (
    SELECT GROUP_CONCAT(CONCAT(ecc2.check_name, ' -> ', ecc2.check_severity) ORDER BY ecc2.check_name SEPARATOR ',')
    FROM ec_case_checks ecc2
    WHERE ecm.case_id = ecc2.case_id
      AND ecc2.check_severity IN ('Red','Orange','Insufficient')
      AND ecc2.check_id IN (235,57,101,25,144,132,61,275,274,72,1,41)
      AND ecc2.check_status <> 9
  ) AS `Pending checks (if Insuff, Red or Orange)`,

  (
    SELECT GROUP_CONCAT(CONCAT(ecc2.check_name, ' -> ', ecc2.check_severity, ' -> ', ecc2.closure_comments) ORDER BY ecc2.check_name SEPARATOR ',')
    FROM ec_case_checks ecc2
    WHERE ecm.case_id = ecc2.case_id
      AND ecc2.check_severity IN ('Red','Orange','Insufficient')
      AND ecc2.check_id IN (235,57,101,25,144,132,61,275,274,72,1,41)
      AND ecc2.check_status <> 9
  ) AS `Comments`

FROM ec_case_master ecm
LEFT JOIN ec_case_fields ecff ON ecm.case_id = ecff.case_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id = ecm.candidate_id
LEFT JOIN ec_client_process ecp ON ecm.process_id = ecp.process_id
LEFT JOIN ec_client ec1 ON ecm.client_id = ec1.client_id

WHERE ecm.client_id = 255820
  AND ecc1.first_name NOT LIKE '%%Test%%'
  AND ecc1.first_name NOT LIKE '%%test%%'
  AND ecc1.first_name NOT LIKE '%%dummy%%'
  AND ecc1.first_name NOT LIKE '%%Dummy%%'
  AND ecm.case_status NOT IN (8,13,14)
"""

# GROUP_CONCAT is truncated at 1024 chars by default; raise it for the session.
SESSION_SETUP = "SET SESSION group_concat_max_len = 1000000"


def _cfg(key: str, default: str = "") -> str:
    """Read config from st.secrets['db'] first, then environment variables."""
    try:
        if "db" in st.secrets and key in st.secrets["db"]:
            return str(st.secrets["db"][key])
    except Exception:
        pass
    return os.getenv(f"DB_{key.upper()}", default)


@st.cache_data(ttl=300, show_spinner="Fetching data from database...")
def load_data() -> pd.DataFrame:
    conn = pymysql.connect(
        host=_cfg("host", "localhost"),
        port=int(_cfg("port", "3306")),
        user=_cfg("user"),
        password=_cfg("password"),
        database=_cfg("name"),
        charset="utf8mb4",
        connect_timeout=15,
    )
    try:
        with conn.cursor() as cur:
            cur.execute(SESSION_SETUP)
            cur.execute(QUERY)
            cols = [c[0] for c in cur.description]
            rows = cur.fetchall()
    finally:
        conn.close()
    return pd.DataFrame(rows, columns=cols)


st.title("BGV Master Tracker - Rupeek")

try:
    data = load_data()
except Exception as e:
    st.error(f"Database error: {e}")
    st.info("Set DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME in a .env file (see .env.example).")
    st.stop()


def build_excel(df: pd.DataFrame) -> bytes:
    """Tracker as .xlsx: bold header, light gray fill, no wrap, thin borders."""
    df = df.copy()
    # Emp id as real numbers where possible, so Excel doesn't flag "number stored as text".
    # Ids with letters or leading zeros (e.g. INT850, 0123) stay as text.
    df["Emp id"] = df["Emp id"].map(
        lambda v: int(v) if isinstance(v, str) and re.fullmatch(r"[1-9]\d{0,14}|0", v.strip()) else v
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Tracker")
        ws = xw.sheets["Tracker"]
        side = Side(style="thin", color="999999")
        border = Border(left=side, right=side, top=side, bottom=side)
        for row in ws.iter_rows():
            for c in row:
                c.border = border
                c.alignment = Alignment(wrap_text=False, vertical="top")
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="D9D9D9")
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(width + 2, 60)
        ws.freeze_panes = "A2"

    return buf.getvalue()


st.success(f"Tracker ready: {len(data)} cases (not displayed).")
c1, c2 = st.columns(2)
c1.download_button(
    "Download Excel",
    build_excel(data),
    "bgv_tracker.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
c2.download_button("Download CSV", data.to_csv(index=False).encode("utf-8-sig"), "bgv_tracker.csv", "text/csv")
if st.button("Refresh data"):
    load_data.clear()
    st.rerun()
