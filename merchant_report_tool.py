"""
Merchant Report & Email Tool v2.4
Professional PyQt6 application for generating merchant transaction reports
and sending them via Email or WhatsApp.
Developed for Hyper Knowledge (HK)
"""

import sys
import os
import re
import logging
import webbrowser
import urllib.parse
import subprocess
import shutil
from datetime import datetime
from pathlib import Path
import math
from typing import Optional, Dict, List, Tuple

import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QPushButton, QLabel, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QFileDialog, QDialog,
    QTextEdit, QTextBrowser, QMessageBox, QFrame, QSplitter, QStatusBar, QToolBar,
    QSizePolicy, QAbstractItemView, QStyle, QScrollArea, QGroupBox,
    QCheckBox, QSpacerItem, QStyledItemDelegate, QMenu, QLineEdit,
    QDateEdit, QTabWidget
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QTimer, QUrl, QDate
from PyQt6.QtGui import (
    QFont, QColor, QPalette, QIcon, QAction, QPixmap, QPainter,
    QLinearGradient, QBrush, QPen, QFontDatabase, QDesktopServices
)

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors as rl_colors
from reportlab.lib.units import mm, inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

from openpyxl import Workbook
from openpyxl.styles import (
    Font as XlFont, Alignment as XlAlign, Border as XlBorder,
    Side as XlSide, PatternFill as XlFill, numbers
)
from openpyxl.utils import get_column_letter

try:
    import win32com.client
    HAS_OUTLOOK = True
except ImportError:
    HAS_OUTLOOK = False


# ── Logging ──
class LogHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.records: List[str] = []
        self.callback = None
    def emit(self, record):
        msg = self.format(record)
        self.records.append(msg)
        if self.callback:
            self.callback(msg)

log_handler = LogHandler()
log_handler.setFormatter(logging.Formatter(
    '%(asctime)s | %(levelname)-8s | %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))
logger = logging.getLogger('MerchantTool')
logger.setLevel(logging.DEBUG)
logger.addHandler(log_handler)


# ── Themes ──
DARK_THEME = {
    "name":"dark","bg_primary":"#0f1419","bg_secondary":"#1a1f2e",
    "bg_tertiary":"#242b3d","bg_card":"#1e2536",
    "accent":"#3b82f6","accent_hover":"#2563eb","accent_green":"#10b981",
    "accent_orange":"#f59e0b","accent_red":"#ef4444","accent_purple":"#8b5cf6",
    "accent_wa":"#25d366","accent_wa_hover":"#1eba57",
    "text_primary":"#e8eaed","text_secondary":"#9ca3af","text_muted":"#6b7280",
    "border":"#2d3548","border_light":"#374151",
    "success":"#059669","warning":"#d97706","error":"#dc2626",
    "table_header":"#1e293b","table_row_alt":"#1a2235","table_hover":"#253352",
    "scrollbar_bg":"#1a1f2e","scrollbar_handle":"#374151",
}
LIGHT_THEME = {
    "name":"light","bg_primary":"#f0f2f5","bg_secondary":"#ffffff",
    "bg_tertiary":"#f8f9fa","bg_card":"#ffffff",
    "accent":"#2563eb","accent_hover":"#1d4ed8","accent_green":"#059669",
    "accent_orange":"#d97706","accent_red":"#dc2626","accent_purple":"#7c3aed",
    "accent_wa":"#25d366","accent_wa_hover":"#1eba57",
    "text_primary":"#111827","text_secondary":"#4b5563","text_muted":"#9ca3af",
    "border":"#e5e7eb","border_light":"#d1d5db",
    "success":"#059669","warning":"#d97706","error":"#dc2626",
    "table_header":"#e2e8f0","table_row_alt":"#f8fafc","table_hover":"#e0e7ff",
    "scrollbar_bg":"#f0f2f5","scrollbar_handle":"#c4c8d0",
}

def build_stylesheet(t):
    return f"""
    QMainWindow, QWidget {{
        background-color: {t['bg_primary']}; color: {t['text_primary']};
        font-family: 'Segoe UI', sans-serif; font-size: 13px;
    }}
    QGroupBox {{
        background-color: {t['bg_secondary']}; border: 1px solid {t['border']};
        border-radius: 8px; margin-top: 15px; padding: 15px 10px 8px 10px;
        font-weight: 600; font-size: 12px;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin; subcontrol-position: top center;
        padding: 0px 10px; color: {t['accent']}; font-weight: 700;
    }}
    QPushButton {{
        background-color: {t['bg_tertiary']}; color: {t['text_primary']};
        border: 1px solid {t['border']}; border-radius: 6px;
        padding: 5px 14px; font-weight: 600; font-size: 11px; min-height: 18px;
    }}
    QPushButton:hover {{ background-color: {t['accent']}; color: #ffffff; border-color: {t['accent']}; }}
    QPushButton:pressed {{ background-color: {t['accent_hover']}; }}
    QPushButton:disabled {{ background-color: {t['bg_tertiary']}; color: {t['text_muted']}; border-color: {t['border']}; }}
    QPushButton#btnAccent {{ background-color: {t['accent']}; color: #ffffff; border: none; }}
    QPushButton#btnAccent:hover {{ background-color: {t['accent_hover']}; }}
    QPushButton#btnGreen {{ background-color: {t['accent_green']}; color: #ffffff; border: none; }}
    QPushButton#btnGreen:hover {{ background-color: {t['success']}; }}
    QPushButton#btnPurple {{ background-color: {t['accent_purple']}; color: #ffffff; border: none; }}
    QPushButton#btnDanger {{ background-color: {t['accent_red']}; color: #ffffff; border: none; }}
    QPushButton#btnWA {{ background-color: {t['accent_wa']}; color: #ffffff; border: none; font-weight: 700; }}
    QPushButton#btnWA:hover {{ background-color: {t['accent_wa_hover']}; }}
    QPushButton#btnToolbar {{
        background-color: transparent; border: 1px solid {t['border']};
        border-radius: 6px; padding: 4px 12px; font-size: 12px;
    }}
    QPushButton#btnToolbar:hover {{ background-color: {t['bg_tertiary']}; border-color: {t['accent']}; color: {t['text_primary']}; }}
    QPushButton#btnFile {{
        background-color: {t['bg_card']}; border: 1px dashed {t['border_light']};
        border-radius: 6px; padding: 6px 14px; font-size: 11px; font-weight: 600;
    }}
    QPushButton#btnFile:hover {{ border-color: {t['accent']}; background-color: {t['bg_tertiary']}; color: {t['text_primary']}; }}
    QPushButton#btnFile:disabled {{ background-color: {t['bg_tertiary']}; color: {t['text_muted']}; border: 1px dashed {t['border_light']}; }}
    QProgressBar {{
        background-color: {t['bg_tertiary']}; border: 1px solid {t['border']};
        border-radius: 5px; text-align: center; font-size: 10px;
        min-height: 10px; max-height: 10px;
    }}
    QProgressBar::chunk {{
        background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {t['accent']}, stop:1 {t['accent_purple']});
        border-radius: 4px;
    }}
    QTableWidget {{
        background-color: {t['bg_secondary']}; alternate-background-color: {t['table_row_alt']};
        border: 1px solid {t['border']}; border-radius: 8px; gridline-color: {t['border']};
        font-size: 11px; selection-background-color: {t['table_hover']}; selection-color: {t['text_primary']};
    }}
    QTableWidget::item {{ padding: 4px 8px; border-bottom: 1px solid {t['border']}; }}
    QTableWidget::item:hover {{ background-color: {t['table_hover']}; color: {t['text_primary']}; }}
    QTableWidget::item:selected {{ background-color: {t['table_hover']}; color: {t['text_primary']}; }}
    QHeaderView::section {{
        background-color: {t['table_header']}; color: {t['text_primary']};
        font-weight: 700; font-size: 11px; padding: 7px 6px;
        border: none; border-bottom: 2px solid {t['accent']}; border-right: 1px solid {t['border']};
    }}
    QComboBox {{
        background-color: {t['bg_tertiary']}; color: {t['text_primary']};
        border: 1px solid {t['border']}; border-radius: 5px; padding: 3px 8px;
        font-size: 11px; min-width: 70px;
    }}
    QComboBox:hover {{ border-color: {t['accent']}; }}
    QComboBox::drop-down {{ border: none; width: 20px; }}
    QComboBox::down-arrow {{ image: none; border-left: 4px solid transparent; border-right: 4px solid transparent; border-top: 5px solid {t['text_secondary']}; margin-right: 6px; }}
    QComboBox QAbstractItemView {{ background-color: {t['bg_secondary']}; color: {t['text_primary']}; border: 1px solid {t['border']}; selection-background-color: {t['accent']}; }}
    QLabel {{ color: {t['text_primary']}; }}
    QLabel#labelTitle {{ font-size: 16px; font-weight: 800; color: {t['accent']}; }}
    QLabel#labelSubtitle {{ font-size: 10px; color: {t['text_secondary']}; }}
    QLabel#labelSection {{ font-size: 13px; font-weight: 700; }}
    QLabel#labelStatus {{ font-size: 10px; color: {t['text_muted']}; padding: 1px 4px; }}
    QLabel#labelBadge {{ background-color: {t['accent']}; color: #fff; border-radius: 8px; padding: 2px 8px; font-size: 10px; font-weight: 700; }}
    QLabel#labelFileOK {{ font-size: 10px; color: {t['accent_green']}; font-weight: 600; }}
    QLabel#labelFileErr {{ font-size: 10px; color: {t['accent_red']}; font-weight: 600; }}
    QStatusBar {{ background-color: {t['bg_secondary']}; color: {t['text_secondary']}; border-top: 1px solid {t['border']}; font-size: 11px; }}
    QScrollBar:vertical {{ background-color: {t['bg_secondary']}; width: 8px; border-radius: 4px; }}
    QScrollBar::handle:vertical {{ background-color: {t['scrollbar_handle']}; border-radius: 4px; min-height: 30px; }}
    QScrollBar::handle:vertical:hover {{ background-color: {t['accent']}; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar:horizontal {{ background-color: {t['bg_secondary']}; height: 8px; border-radius: 4px; }}
    QScrollBar::handle:horizontal {{ background-color: {t['scrollbar_handle']}; border-radius: 4px; min-width: 30px; }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
    QDialog {{ background-color: {t['bg_primary']}; }}
    QTextEdit {{ background-color: {t['bg_secondary']}; color: {t['text_primary']}; border: 1px solid {t['border']}; border-radius: 6px; font-family: 'Consolas', monospace; font-size: 12px; padding: 8px; }}
    QFrame#separator {{ background-color: {t['border']}; max-height: 1px; }}
    QDateEdit {{
        background-color: {t['bg_tertiary']}; color: {t['text_primary']};
        border: 1px solid {t['border']}; border-radius: 5px; padding: 4px 8px;
        font-size: 12px; font-weight: 600; min-width: 110px;
    }}
    QDateEdit:hover {{ border-color: {t['accent']}; }}
    QDateEdit::drop-down {{ border: none; width: 20px; }}
    QDateEdit::down-arrow {{ image: none; border-left: 4px solid transparent; border-right: 4px solid transparent; border-top: 5px solid {t['text_secondary']}; margin-right: 6px; }}
    QCalendarWidget {{ background-color: {t['bg_secondary']}; color: {t['text_primary']}; }}
    QTabWidget::pane {{ border: 1px solid {t['border']}; border-radius: 4px; top: -1px; }}
    QTabBar::tab {{
        background-color: {t['bg_tertiary']}; color: {t['text_secondary']};
        border: 1px solid {t['border']}; padding: 8px 16px; margin-right: 2px;
        border-top-left-radius: 4px; border-top-right-radius: 4px;
        font-weight: 600; font-size: 13px;
    }}
    QTabBar::tab:selected {{
        background-color: {t['bg_secondary']}; color: {t['accent']};
        border-bottom-color: {t['bg_secondary']};
    }}
    QTabBar::tab:hover:!selected {{ background-color: {t['table_hover']}; color: {t['text_primary']}; }}
    """


# ── Workers ──
class FileLoaderThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(object, str)
    error = pyqtSignal(str)
    def __init__(self, filepath, file_label):
        super().__init__()
        self.filepath = filepath
        self.file_label = file_label
    def run(self):
        try:
            logger.info(f"Loading: {self.file_label} -> {os.path.basename(self.filepath)}")
            self.progress.emit(10)
            sz = os.path.getsize(self.filepath)
            logger.info(f"File size: {sz/(1024*1024):.1f} MB")
            self.progress.emit(30)
            ext = os.path.splitext(self.filepath)[1].lower()
            df = None
            
            def try_read_csv():
                try: return pd.read_csv(self.filepath, encoding='utf-8')
                except Exception:
                    try: return pd.read_csv(self.filepath, encoding='windows-1256')
                    except Exception: return pd.read_csv(self.filepath, encoding='latin1', errors='replace')
            
            if ext == '.csv':
                df = try_read_csv()
            else:
                try:
                    df = pd.read_excel(self.filepath)
                except Exception as ex:
                    logger.warning(f"read_excel failed ({ex}), treating as CSV fallback...")
                    df = try_read_csv()
                    
            self.progress.emit(80)
            r, c = df.shape
            logger.info(f"Loaded {self.file_label}: {r:,} rows x {c} cols")
            logger.info(f"Columns: {list(df.columns)}")
            self.progress.emit(100)
            self.finished.emit(df, self.file_label)
        except Exception as e:
            logger.error(f"Failed to load {self.file_label}: {e}")
            self.error.emit(f"Error loading {self.file_label}:\n{e}")


class ReportGeneratorThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, list, str)
    error = pyqtSignal(str)
    match_result = pyqtSignal(str, float, float, bool)  # (merchant_id, payment_amount, net_total, matched)
    def __init__(self, merchant_name, merchant_id, report_format, transactions_df, clients_df, output_dir, to_date=None, is_b2b=False):
        self.is_b2b = is_b2b
        super().__init__()
        self.merchant_name = merchant_name
        self.merchant_id = merchant_id
        self.report_format = report_format
        self.transactions_df = transactions_df
        self.clients_df = clients_df
        self.output_dir = output_dir
        self.to_date = to_date

    def run(self):
        try:
            import traceback
            generated_files = []
            mid = int(self.merchant_id)
            logger.info(f"Generating {self.report_format} for: {self.merchant_name} (ID: {mid})")
            self.progress.emit(10)

            # ── Get payment_amount & IBAN from clients ───────────────────
            payment_amount = 0.0; iban = ''
            if self.clients_df is not None:
                cdf = self.clients_df.copy()
                cdf.columns = [str(c).strip() for c in cdf.columns]
                # Safely compare merchant_id as strings to avoid type mismatches
                cr = cdf[cdf['merchant_id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() == str(self.merchant_id).strip()]
                if not cr.empty:
                    # Clean commas if stored as string and calculate sum
                    _pa = cr['payment_amount'].astype(str).str.replace(',', '', regex=False)
                    payment_amount = float(pd.to_numeric(_pa, errors='coerce').fillna(0).sum())
                    if 'iban' in cr.columns:
                        iban = str(cr['iban'].iloc[0])
                    logger.info(f"IBAN: {iban}  payment_amount: {payment_amount:,.2f} IQD")
                else:
                    logger.warning(f"Merchant {mid} not in clients file")

            # ── Filter transactions ──────────────────────────────────────────
            all_txn = self.transactions_df[self.transactions_df['Merchant ID'] == mid].copy()
            if all_txn.empty:
                msg = f"No transactions for Merchant ID: {mid}"
                logger.warning(msg); self.error.emit(msg); return
            logger.info(f"Found {len(all_txn):,} total transactions")
            self.progress.emit(20)

            # ── Fallback to Streamlit exact column names if present ──────────
            cols_lower = {c.lower().strip(): c for c in all_txn.columns}
            amt_col = cols_lower.get('vtr dcc amou', 'Transaction Amount')
            date_col = cols_lower.get('process date', 'Transaction Date')
            
            # Clean commas if any, ensure numeric
            _amt = all_txn[amt_col].astype(str).str.replace(',', '', regex=False)
            _chr = all_txn['Merchant Commission'].astype(str).str.replace(',', '', regex=False)
            all_txn['Transaction Amount_calc'] = pd.to_numeric(_amt, errors='coerce').fillna(0)
            all_txn['Merchant Commission_calc'] = pd.to_numeric(_chr, errors='coerce').fillna(0)
            
            # net = amount + commission (no absolute values, preserves positive/negative exactly as Streamlit)
            all_txn['net_amount'] = all_txn['Transaction Amount_calc'] + all_txn['Merchant Commission_calc']
            all_txn['_date'] = pd.to_datetime(all_txn[date_col], errors='coerce').dt.normalize()
            all_txn = all_txn.sort_values('_date')

            # ── Payment Amount Matching Algorithm ────────────────────────────
            matched = False
            if self.to_date is not None and payment_amount > 0:
                to_dt = pd.to_datetime(self.to_date)
                # Filter up to to_dt (now precisely midnight-comparable because _date is normalized)
                txn_pool = all_txn[all_txn['_date'] <= to_dt].copy()
                
                # Streamlit logic: valid_dates sorted ascending (oldest to newest)
                valid_dates = txn_pool['_date'].dropna().sort_values().unique()
                
                for start in valid_dates:
                    start_ts = pd.to_datetime(start)
                    temp = txn_pool[
                        (txn_pool['_date'] >= start_ts) &
                        (txn_pool['_date'] <= to_dt)
                    ]
                    current_net = temp['net_amount'].sum()
                    
                    check_net = abs(current_net) if getattr(self, 'is_b2b', False) else current_net
                    if math.isclose(check_net, payment_amount, rel_tol=1e-5, abs_tol=0.05):
                        txn = temp.copy()
                        matched = True
                        logger.info(f"✅ Match! start={start_ts.date()}, net={current_net:,.2f} = payment={payment_amount:,.2f}")
                        break
                
                if not matched:
                    txn = txn_pool.copy()
                    logger.warning(f"❌ No match. total net={txn['net_amount'].sum():,.2f} ≠ payment={payment_amount:,.2f}")
            else:
                txn = all_txn.copy()
                if payment_amount > 0:
                    sum_net = txn['net_amount'].sum()
                    if getattr(self, 'is_b2b', False):
                        matched = math.isclose(abs(sum_net), payment_amount, rel_tol=1e-5, abs_tol=0.05)
                    else:
                        matched = math.isclose(sum_net, payment_amount, rel_tol=1e-5, abs_tol=0.05)
                logger.info(f"No to_date, using all {len(txn)} txns")

            if txn.empty:
                msg = f"No transactions in date range for Merchant ID: {mid}"
                logger.warning(msg); self.error.emit(msg); return

            net_total = txn['net_amount'].sum()
            self.match_result.emit(str(mid), payment_amount, net_total, matched)
            logger.info(f"Using {len(txn):,} txns | net={net_total:,.2f} | payment={payment_amount:,.2f} | matched={matched}")

            # Sort by Terminal ID then date/time
            txn = txn.sort_values(['Terminal ID', 'Transaction Date', 'Transaction Time'])

            # ── Compute display amounts ────────────────────────────────────
            if getattr(self, 'is_b2b', False):
                txn['Trx Amount'] = -txn['Transaction Amount_calc'].abs()
                txn['Chr Amount'] = txn['Merchant Commission_calc'].abs()
            else:
                txn['Trx Amount'] = txn['Transaction Amount_calc'].abs()
                txn['Chr Amount'] = -txn['Merchant Commission_calc'].abs()
            
            txn['Net Amount'] = txn['Trx Amount'] + txn['Chr Amount']

            # ── FIX 2: MDR rate string (matches Streamlit format) ─────────────
            def _mdr(row):
                amt = abs(row['Trx Amount'])
                if amt == 0: return '0.0%'
                rate = abs(row['Chr Amount']) / amt * 100
                return f'{rate:.1f}%'
            txn['MDR'] = txn.apply(_mdr, axis=1)

            # ── FIX 3: Type mapping (N\I column) ────────────────────────────
            ni_col = 'N\\I'
            if ni_col in txn.columns:
                txn['Txn Type'] = txn[ni_col].astype(str).str[:4].str.upper()
            else:
                txn['Txn Type'] = 'NATI'

            # ── FIX 4: Trx Type from correct column ─────────────────────────
            tc_col = cols_lower.get('transacion code', cols_lower.get('transaction code'))
            if tc_col:
                txn['Trx Type'] = txn[tc_col]
            else:
                txn['Trx Type'] = 'Sales draft'

            txn['Currency'] = 'IQD'

            # Date range from filtered transactions
            d_min = pd.to_datetime(txn[date_col]).min()
            d_max = pd.to_datetime(txn[date_col]).max()
            dr = f"From {d_min.strftime('%d-%m-%Y')} To {d_max.strftime('%d-%m-%Y')}"
            logger.info(f"Date range: {dr}")
            self.progress.emit(30)
            os.makedirs(self.output_dir, exist_ok=True)

            # ── CALCULATE TOTAL CASH BACK ───────────────────────────────────
            cash_back_total = 0.0
            cb_mask = txn['Trx Type'].astype(str).str.contains('Cash Back', na=False, case=False)
            cash_back_total += txn.loc[cb_mask, 'Transaction Amount_calc'].sum()

            vtr_col = cols_lower.get('vtr_dcc_rebate_amou')
            if vtr_col:
                cash_back_total += pd.to_numeric(txn[vtr_col], errors='coerce').fillna(0).sum()

            # Assign Cash Back rows to a dedicated Terminal ID like the source bank reports
            txn['Terminal ID'] = txn['Terminal ID'].astype(object) # Avoid dtype error if it was strictly float64
            txn.loc[cb_mask, 'Terminal ID'] = 'Portal/Mobile'
            
            # Force Txn Type to ONUS for Cash Back
            txn.loc[cb_mask, 'Txn Type'] = 'ONUS'

            # ── FIX 6: Correct summary totals ───────────────────────────────
            summary = {
                'merchant_name': self.merchant_name,
                'merchant_id': mid,
                'date_range': dr,
                'iban': iban,
                'settled_amount': payment_amount,
                'total_transactions': len(txn),
                'total_trx_amount': txn['Trx Amount'].sum(),      # positive total (e.g. 19,314,506)
                'total_chr_amount': txn['Chr Amount'].sum(),      # negative fees  (e.g. -198,522)
                'total_net_amount': txn['Net Amount'].sum(),      # positive net   (e.g. 19,115,983)
                'cash_back': float(cash_back_total),              # explicit float
                # keep old keys as aliases so email/WA functions still work
                'total_amount': txn['Trx Amount'].sum(),
                'total_commission': txn['Chr Amount'].sum(),
            }

            # ── FIX 7: Correct columns for report ───────────────────────────
            txn['Trx Date'] = pd.to_datetime(txn['Transaction Date']).dt.strftime('%d/%m/%Y')
            report_cols = ['Trx Date', 'Masked Card Number', 'Trx Type', 'Currency',
                           'Trx Amount', 'Chr Amount', 'Net Amount', 'MDR',
                           'Auth Code', 'Txn Type', 'Terminal ID']
            rdata = txn[[c for c in report_cols if c in txn.columns]].copy()
            
            # Clean Terminal ID from .0 (convert float 123.0 to string '123')
            if 'Terminal ID' in rdata.columns:
                def _clean_tid(val):
                    if pd.isna(val) or val == '': return ''
                    try:
                        f = float(val)
                        return str(int(f)) if f.is_integer() else str(val)
                    except:
                        return str(val)
                rdata['Terminal ID'] = rdata['Terminal ID'].apply(_clean_tid)

            fmt = self.report_format.lower()
            if fmt in ('pdf', 'both'):
                self.progress.emit(40)
                p = os.path.join(self.output_dir, f"{self.merchant_name}_{mid}_Report.pdf")
                self._gen_pdf(rdata, summary, p)
                generated_files.append(p)
                logger.info(f"PDF saved: {p}")
            if fmt in ('excel', 'both'):
                self.progress.emit(70)
                p = os.path.join(self.output_dir, f"{self.merchant_name}_{mid}_Report.xlsx")
                self._gen_xlsx(rdata, summary, p)
                generated_files.append(p)
                logger.info(f"Excel saved: {p}")

            self.progress.emit(100)
            logger.info(f"Done: {self.merchant_name} ({len(generated_files)} files)")
            self.finished.emit(self.merchant_name, generated_files, dr)
        except Exception as e:
            logger.error(f"Report failed for {self.merchant_name}: {e}")
            logger.error(traceback.format_exc())
            self.error.emit(f"Failed:\n{e}")

    def _gen_pdf(self, data, summary, filepath):
        logger.info("Generating PDF...")
        doc = SimpleDocTemplate(filepath, pagesize=landscape(A4),
            rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
        styles = getSampleStyleSheet()
        els = []
        aw = landscape(A4)[0] - 20*mm

        # ── Header block ────────────────────────────────────────────────────
        tot_net_w_cb = summary['total_net_amount'] + summary.get('cash_back', 0.0)
        hdr_data = [
            ['Merchant ID:', str(summary['merchant_id']),
             'Merchant Name:', summary['merchant_name']],
            ['IBAN:', summary.get('iban', 'N/A'),
             'Date Range:', summary['date_range']],
            ['Total Transactions:', f"{summary['total_transactions']:,}", '', ''],
            ['Generated:', datetime.now().strftime('%d/%m/%Y %H:%M:%S'), '', '']
        ]
        ht = Table(hdr_data, colWidths=[aw*0.12, aw*0.28, aw*0.14, aw*0.46])
        ht.setStyle(TableStyle([
            ('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),
            ('FONTNAME',(2,0),(2,-1),'Helvetica-Bold'),
            ('FONTSIZE',(0,0),(-1,-1),8),
            ('TEXTCOLOR',(0,0),(0,-1),rl_colors.HexColor('#4b5563')),
            ('TEXTCOLOR',(2,0),(2,-1),rl_colors.HexColor('#4b5563')),
            ('BACKGROUND',(0,0),(-1,-1),rl_colors.HexColor('#f0f9ff')),
            ('BOX',(0,0),(-1,-1),0.5,rl_colors.HexColor('#bfdbfe')),
            ('INNERGRID',(0,0),(-1,-1),0.25,rl_colors.HexColor('#dbeafe')),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
            ('LEFTPADDING',(0,0),(-1,-1),7),
        ]))
        els.append(ht)
        els.append(Spacer(1, 4*mm))

        # ── Column definitions (matching original PDF) ───────────────────────
        col_map = {
            'Trx Date':           ('Trx Date',   0.085),
            'Masked Card Number': ('Card Number', 0.155),
            'Trx Type':           ('Trx Type',   0.105),
            'Currency':           ('Cur',         0.045),
            'Trx Amount':         ('Trx Amount',  0.115),
            'Chr Amount':         ('Commission',  0.095),
            'Net Amount':         ('Net Amount',  0.115),
            'MDR':                ('MDR',          0.065),
            'Auth Code':          ('Auth',         0.085),
            'Txn Type':           ('Type',         0.075),
        }
        display_cols = [c for c in col_map if c in data.columns]
        hdrs       = [col_map[c][0] for c in display_cols]
        total_prop = sum(col_map[c][1] for c in display_cols)
        cw         = [aw * col_map[c][1] / total_prop for c in display_cols]
        amount_cols = {'Trx Amount', 'Chr Amount', 'Net Amount'}

        def fmt_val(col, v):
            if v is None or (isinstance(v, float) and pd.isna(v)): return ''
            if col in amount_cols:
                return f"{float(v):,.2f}"
            return str(v)

        def row_style_for_trx_type(trx_type_val):
            """Merchant Cash Back rows get a light blue tint."""
            if 'cash back' in str(trx_type_val).lower():
                return rl_colors.HexColor('#eff6ff')
            return None

        # ── Group by Terminal ID ────────────────────────────────────────────
        terminals = sorted(data['Terminal ID'].unique(), key=lambda x: str(x)) if 'Terminal ID' in data.columns else [None]
        acct_count = 0; acct_trx = 0.0; acct_chr = 0.0; acct_net = 0.0
        CHUNK = 40   # rows per page chunk

        term_label_style = ParagraphStyle(
            'TL', parent=styles['Normal'], fontSize=8, fontName='Helvetica-Bold',
            textColor=rl_colors.HexColor('#1e40af'), spaceBefore=5, spaceAfter=2)

        for term_id in terminals:
            tdf = data[data['Terminal ID'] == term_id] if 'Terminal ID' in data.columns else data
            if tdf.empty: continue

            # Terminal header (skip label for Portal/Mobile)
            if str(term_id) != 'Portal/Mobile':
                t_hdr = f"Terminal ID: {term_id}"
                els.append(Paragraph(t_hdr, term_label_style))

            # Build rows
            rows = [hdrs]
            t_trx = 0.0; t_chr = 0.0; t_net = 0.0
            trx_type_per_row = []  # track for row-level styling

            for _, r in tdf.iterrows():
                rd = [fmt_val(c, r[c]) for c in display_cols]
                rows.append(rd)
                try: t_trx += float(r.get('Trx Amount', 0) or 0)
                except: pass
                try: t_chr += float(r.get('Chr Amount', 0) or 0)
                except: pass
                try: t_net += float(r.get('Net Amount', 0) or 0)
                except: pass
                trx_type_per_row.append(str(r.get('Trx Type', '')))

            # Total By Terminal row
            sub = [''] * len(display_cols)
            sub[0] = f"{len(tdf)}"
            for col_name, total_val in [('Trx Amount', t_trx),
                                          ('Chr Amount', t_chr),
                                          ('Net Amount', t_net)]:
                if col_name in display_cols:
                    sub[display_cols.index(col_name)] = f"{total_val:,.2f}"
            rows.append(sub)

            # Render in chunks
            data_rows  = rows[1:-1]   # exclude header & total
            total_row  = rows[-1]
            for chunk_start in range(0, max(len(data_rows), 1), CHUNK):
                chunk = [rows[0]] + data_rows[chunk_start:chunk_start+CHUNK]
                is_last_chunk = (chunk_start + CHUNK >= len(data_rows))
                if is_last_chunk:
                    chunk.append(total_row)

                n = len(chunk)
                tbl = Table(chunk, colWidths=cw, repeatRows=1)
                ts_cmds = [
                    ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                    ('FONTSIZE',(0,0),(-1,0),7.5),
                    ('FONTSIZE',(0,1),(-1,-1),7),
                    ('BACKGROUND',(0,0),(-1,0),rl_colors.HexColor('#1e40af')),
                    ('TEXTCOLOR',(0,0),(-1,0),rl_colors.white),
                    ('ALIGN',(0,0),(-1,-1),'CENTER'),
                    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                    ('TOPPADDING',(0,0),(-1,-1),3),
                    ('BOTTOMPADDING',(0,0),(-1,-1),3),
                    ('ROWBACKGROUNDS',(0,1),(-1,-2 if is_last_chunk else -1),
                     [rl_colors.white, rl_colors.HexColor('#f8fafc')]),
                    ('BOX',(0,0),(-1,-1),0.5,rl_colors.HexColor('#cbd5e1')),
                    ('INNERGRID',(0,0),(-1,-1),0.25,rl_colors.HexColor('#e2e8f0')),
                ]
                if is_last_chunk:
                    # Style the Total By Terminal row
                    ts_cmds += [
                        ('BACKGROUND',(0,n-1),(-1,n-1),rl_colors.HexColor('#dbeafe')),
                        ('TEXTCOLOR',(0,n-1),(-1,n-1),rl_colors.HexColor('#1e3a8a')),
                        ('FONTNAME',(0,n-1),(-1,n-1),'Helvetica-Bold'),
                        ('FONTSIZE',(0,n-1),(-1,n-1),7.5),
                    ]
                # Right-align amount columns
                for ac in amount_cols:
                    if ac in display_cols:
                        ci = display_cols.index(ac)
                        ts_cmds.append(('ALIGN',(ci,1),(ci,-1),'RIGHT'))
                tbl.setStyle(TableStyle(ts_cmds))
                els.append(tbl)
                if not is_last_chunk:
                    els.append(PageBreak())

            els.append(Spacer(1, 3*mm))
            acct_count += len(tdf)
            acct_trx += t_trx; acct_chr += t_chr; acct_net += t_net

        # ── Total By Account ────────────────────────────────────────────────
        els.append(Spacer(1, 3*mm))
        acct_row = [[
            f"Total — Transactions : {acct_count:,} ",
            f"Trx Amount:  {acct_trx:,.2f} IQD",
            f"Total Commission:  {acct_chr:,.2f} IQD",
            f"Net Amount:  {acct_net:,.2f} IQD",
        ]]
        at = Table(acct_row, colWidths=[aw*0.3, aw*0.23, aw*0.23, aw*0.24])
        at.setStyle(TableStyle([
            ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'),
            ('FONTSIZE',(0,0),(-1,-1),9),
            ('BACKGROUND',(0,0),(-1,-1),rl_colors.HexColor('#1e3a8a')),
            ('TEXTCOLOR',(0,0),(-1,-1),rl_colors.white),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('BOX',(0,0),(-1,-1),1,rl_colors.HexColor('#1e40af')),
            ('TOPPADDING',(0,0),(-1,-1),9),
            ('BOTTOMPADDING',(0,0),(-1,-1),9),
        ]))
        els.append(at)
        doc.build(els)

    def _gen_xlsx(self, data, summary, filepath):
        logger.info("Generating Excel...")
        wb = Workbook(); ws = wb.active; ws.title = "Transaction Report"

        # ── Styles ────────────────────────────────────────────────────────
        hf  = XlFont(name='Segoe UI', bold=True, size=9, color='FFFFFF')
        hfl = XlFill(start_color='1E40AF', end_color='1E40AF', fill_type='solid')
        ha  = XlAlign(horizontal='center', vertical='center', wrap_text=True)
        df_ = XlFont(name='Segoe UI', size=9)
        da  = XlAlign(horizontal='center', vertical='center')
        ra  = XlAlign(horizontal='right',  vertical='center')
        af  = '#,##0.00'
        tb  = XlBorder(
            left=XlSide(style='thin', color='D1D5DB'),
            right=XlSide(style='thin', color='D1D5DB'),
            top=XlSide(style='thin', color='D1D5DB'),
            bottom=XlSide(style='thin', color='D1D5DB'))
        sub_fill    = XlFill(start_color='DBEAFE', end_color='DBEAFE', fill_type='solid')
        sub_font    = XlFont(name='Segoe UI', bold=True, size=9, color='1E3A8A')
        acct_fill   = XlFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
        acct_font   = XlFont(name='Segoe UI', bold=True, size=9, color='FFFFFF')
        term_fill   = XlFill(start_color='EFF6FF', end_color='EFF6FF', fill_type='solid')
        alt_fill    = XlFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
        ifl         = XlFill(start_color='F0F9FF', end_color='F0F9FF', fill_type='solid')
        ikf         = XlFont(name='Segoe UI', bold=True, size=10, color='4B5563')
        ivf         = XlFont(name='Segoe UI', size=10, color='111827')

        # ── Title ─────────────────────────────────────────────────────────
        ws.merge_cells('A1:K1')
        ws['A1'] = 'Transaction Settlement Report'
        ws['A1'].font = XlFont(name='Segoe UI', bold=True, size=14, color='1E40AF')
        ws['A1'].alignment = XlAlign(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 32

        ws.merge_cells('A2:K2')
        ws['A2'] = f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
        ws['A2'].font = XlFont(name='Segoe UI', size=9, color='6B7280')
        ws['A2'].alignment = XlAlign(horizontal='center')

        # ── Summary info block ────────────────────────────────────────────
        tot_net_w_cb = summary['total_net_amount'] + summary.get('cash_back', 0.0)
        irows = [
            ('Merchant Name:',   summary['merchant_name'],
             'Merchant ID:',     str(summary['merchant_id'])),
            ('Date Range:',      summary['date_range'],
             'Total Txns:',      f"{summary['total_transactions']:,}"),
            ('IBAN:',            summary.get('iban', 'N/A'),
             'Trx Amount:',      f"{summary['total_trx_amount']:,.2f} IQD"),
            ('Total Commission:',f"{summary['total_chr_amount']:,.2f} IQD",
             'Net Amount:',      f"{summary['total_net_amount']:,.2f} IQD"),
            ('', '', '', ''),
        ]
        for i, rd in enumerate(irows, start=4):
            for c, v, fnt in [(1,rd[0],ikf),(2,rd[1],ivf),(4,rd[2],ikf),(5,rd[3],ivf)]:
                cell = ws.cell(row=i, column=c, value=v)
                cell.font = fnt; cell.fill = ifl; cell.border = tb
                cell.alignment = XlAlign(horizontal='left', vertical='center')

        # ── Column headers for data table ─────────────────────────────────
        # FIX: Correct column mapping matching original PDF
        col_map = {
            'Trx Date':           'Trx Date',
            'Masked Card Number': 'Card Number',
            'Trx Type':           'Trx Type',
            'Currency':           'Cur',
            'Trx Amount':         'Trx Amount',
            'Chr Amount':         'Commission',
            'Net Amount':         'Net Amount',
            'MDR':                'MDR',
            'Auth Code':          'Auth',
            'Txn Type':           'Type',
        }
        display_cols = [c for c in col_map if c in data.columns]
        hdrs = [col_map[c] for c in display_cols]
        amount_cols = {'Trx Amount', 'Chr Amount', 'Net Amount'}

        # Keep track of current row
        cur_row = 9
        terminals = sorted(data['Terminal ID'].unique(), key=lambda x: str(x)) if 'Terminal ID' in data.columns else [None]

        acct_count = 0; acct_trx = 0.0; acct_chr = 0.0; acct_net = 0.0

        for term_id in terminals:
            tdf = data[data['Terminal ID'] == term_id] if 'Terminal ID' in data.columns else data
            if tdf.empty: continue

            # ── Terminal ID label row ──
            if str(term_id) != 'Portal/Mobile':
                tc = ws.cell(row=cur_row, column=1,
                             value=f"Terminal ID: {term_id}  ({len(tdf)} transactions)")
                tc.font = XlFont(name='Segoe UI', bold=True, size=9, color='1E40AF')
                tc.fill = term_fill
                tc.alignment = XlAlign(horizontal='left', vertical='center')
                tc.border = tb
                for ci2 in range(2, len(display_cols) + 1):
                    blank = ws.cell(row=cur_row, column=ci2)
                    blank.fill = term_fill; blank.border = tb
                ws.row_dimensions[cur_row].height = 20
                cur_row += 1

            # ── Column headers ────────────────────────────────────────────
            for ci, h in enumerate(hdrs, start=1):
                cell = ws.cell(row=cur_row, column=ci, value=h)
                cell.font = hf; cell.fill = hfl; cell.alignment = ha; cell.border = tb
            ws.row_dimensions[cur_row].height = 24
            hdr_row = cur_row
            cur_row += 1

            # ── Data rows ─────────────────────────────────────────────────
            t_trx = 0.0; t_chr = 0.0; t_net = 0.0
            for ri, (_, row) in enumerate(tdf.iterrows()):
                fill = alt_fill if ri % 2 == 1 else None
                for ci, cn in enumerate(display_cols, start=1):
                    v = row[cn] if cn in row.index else ''
                    if pd.isna(v): v = ''
                    cell = ws.cell(row=cur_row, column=ci, value=v)
                    cell.font = df_; cell.border = tb
                    if cn in amount_cols and isinstance(v, (int, float)):
                        cell.number_format = af
                        cell.alignment = ra
                    else:
                        cell.alignment = da
                    if fill: cell.fill = fill
                    # Accumulate totals
                    if cn == 'Trx Amount' and isinstance(v, (int, float)): t_trx += v
                    if cn == 'Chr Amount' and isinstance(v, (int, float)): t_chr += v
                    if cn == 'Net Amount' and isinstance(v, (int, float)): t_net += v
                ws.row_dimensions[cur_row].height = 18
                cur_row += 1

            # ── Total By Terminal row ─────────────────────────────────────
            for ci, cn in enumerate(display_cols, start=1):
                cell = ws.cell(row=cur_row, column=ci)
                cell.font = sub_font; cell.fill = sub_fill; cell.border = tb
                cell.alignment = ra if cn in amount_cols else da
                if ci == 1:
                    cell.value = f"Terminal Total: {len(tdf)}"
                    cell.alignment = XlAlign(horizontal='left', vertical='center')
                elif cn == 'Trx Amount':
                    cell.value = t_trx; cell.number_format = af
                elif cn == 'Chr Amount':
                    cell.value = t_chr; cell.number_format = af
                elif cn == 'Net Amount':
                    cell.value = t_net; cell.number_format = af
            ws.row_dimensions[cur_row].height = 20
            cur_row += 2   # blank row between terminals

            acct_count += len(tdf)
            acct_trx += t_trx; acct_chr += t_chr; acct_net += t_net

        # ── Total By Account row (Fix 5: no merge → no MergedCell errors) ──
        for ci in range(1, len(display_cols) + 1):
            cell = ws.cell(row=cur_row, column=ci)
            cell.font = acct_font; cell.fill = acct_fill; cell.border = tb
            cell.alignment = XlAlign(horizontal='center', vertical='center')
        ws.cell(row=cur_row, column=1).value = f"Total — Transactions :  {acct_count:,}"
        ws.cell(row=cur_row, column=1).alignment = XlAlign(horizontal='left', vertical='center')
        for cn, val in [('Trx Amount', acct_trx),
                         ('Chr Amount', acct_chr),
                         ('Net Amount', acct_net)]:
            if cn in display_cols:
                ci = display_cols.index(cn) + 1
                cell = ws.cell(row=cur_row, column=ci)
                cell.value = val; cell.number_format = af
                cell.alignment = XlAlign(horizontal='right', vertical='center')
        ws.row_dimensions[cur_row].height = 24

        # ── Column widths ─────────────────────────────────────────────────
        for ci, cn in enumerate(display_cols, start=1):
            max_len = len(hdrs[ci-1])
            ltr = get_column_letter(ci)
            ws.column_dimensions[ltr].width = min(max_len + 5, 28)
        # Fixed widths for key columns
        width_map = {
            'Trx Date': 12, 'Masked Card Number': 20, 'Trx Type': 16,
            'Currency': 8, 'Trx Amount': 16, 'Chr Amount': 16, 'Net Amount': 16,
            'MDR': 9, 'Auth Code': 11, 'Txn Type': 8, 'Terminal ID': 14,
        }
        for ci, cn in enumerate(display_cols, start=1):
            if cn in width_map:
                ws.column_dimensions[get_column_letter(ci)].width = width_map[cn]

        # Auto-filter on the first header row we created (not the last - that
        # would only filter the final terminal's table and look broken).
        # Skipping auto_filter here because each terminal has its own header;
        # a single auto_filter range can't span multiple disjoint tables.

        # FIX: Previously this was `ws.freeze_panes = f'A{hdr_row + 1}'` which
        # froze EVERY row above the LAST terminal's header — making most of
        # the sheet frozen and un-scrollable. Instead, freeze just the top
        # summary block (rows 1-8: title + merchant info) so totals stay in
        # view while the user scrolls through transactions.
        ws.freeze_panes = 'A9'
        wb.save(filepath)


# ── B2B Workers ──
class CsvLoaderThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(object, str)
    error = pyqtSignal(str)
    def __init__(self, filepath, file_label):
        super().__init__()
        self.filepath = filepath
        self.file_label = file_label
    def run(self):
        try:
            logger.info(f"Loading CSV: {self.file_label} -> {os.path.basename(self.filepath)}")
            self.progress.emit(10)
            df = pd.read_csv(self.filepath)
            self.progress.emit(80)
            r, c = df.shape
            logger.info(f"Loaded {self.file_label}: {r:,} rows x {c} cols")
            logger.info(f"Columns: {list(df.columns)}")
            self.progress.emit(100)
            self.finished.emit(df, self.file_label)
        except Exception as e:
            logger.error(f"Failed to load {self.file_label}: {e}")
            self.error.emit(f"Error loading {self.file_label}:\n{e}")


def parse_contact_info(email_str):
    result = {'emails':[],'phones':[],'is_silent':False,'is_wa_group':False,'wa_link':None,'raw':str(email_str).strip()}
    if pd.isna(email_str): result['is_silent']=True; return result
    text = str(email_str).strip()
    if not text: result['is_silent']=True; return result
    
    # Detect WhatsApp Group Links (robust regex)
    wa_links = re.findall(r'(?:https?://)?chat\.whatsapp\.com/[\w-]+', text)
    if wa_links:
        link = wa_links[0]
        if not link.startswith('http'): link = 'https://' + link
        result['wa_link'] = link
        result['is_wa_group'] = True
    elif 'group on wa' in text.lower() or 'كروب' in text.lower():
        result['is_wa_group'] = True
        
    if text.lower().startswith('silent') and '@' not in text:
        result['is_silent']=True
        phones = re.findall(r'\+?\d[\d\s\-]{8,}', text)
        if phones: result['phones']=[re.sub(r'[\s\-]','',p) for p in phones]; result['is_silent']=False
        return result
        
    result['emails'] = list(set(re.findall(r'[\w.+-]+@[\w.-]+\.\w+', text)))
    phones = re.findall(r'\+?\d[\d\s\-]{8,}', text)
    result['phones'] = [re.sub(r'[\s\-]','',p) for p in phones]
    
    if not result['emails'] and not result['phones'] and not result['wa_link']:
        result['is_silent']=True
    return result

def normalize_type(type_str):
    if pd.isna(type_str): return 'pdf'
    t = str(type_str).lower().strip()
    if 'amount' in t: return 'amount'
    if 'excel' in t and 'pdf' in t: return 'both'
    if 'excel' in t: return 'excel'
    return 'pdf'


# ── Fix 2: ComboBox that ignores scroll wheel ──────────────────────────────
class NoScrollComboBox(QComboBox):
    """Prevents accidental value change when scrolling over the combobox."""
    def wheelEvent(self, event):
        event.ignore()   # do nothing — only keyboard/click changes the value


# ── Fix 1: Table with full-row hover highlight ─────────────────────────────
class RowHoverDelegate(QStyledItemDelegate):
    """Paints the entire hovered row with the theme's hover color."""
    HOVER_LIGHT = QColor('#e0e7ff')
    HOVER_DARK  = QColor('#253352')

    def __init__(self, table: 'HoverTableWidget', is_dark: bool = True):
        super().__init__(table)
        self._table   = table
        self._is_dark = is_dark

    def set_dark(self, dark: bool):
        self._is_dark = dark

    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        if index.row() == self._table._hover_row:
            color = self.HOVER_DARK if self._is_dark else self.HOVER_LIGHT
            option.backgroundBrush = QBrush(color)


class HoverTableWidget(QTableWidget):
    """QTableWidget that highlights the entire hovered row."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._hover_row = -1
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)

    def mouseMoveEvent(self, event):
        idx = self.indexAt(event.pos())
        row = idx.row() if idx.isValid() else -1
        if row != self._hover_row:
            self._hover_row = row
            self.viewport().update()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        if self._hover_row != -1:
            self._hover_row = -1
            self.viewport().update()
        super().leaveEvent(event)


# ── Fix 6: Required columns per file index ────────────────────────────────
FILE_REQUIRED_COLUMNS = {
    0: {   # Send Info (معلومات الارسال)
        'required': ['Merchant Name', 'Merchant ID', 'Type '],
        'label': 'Send Info File',
        'hint':  'يجب أن يحتوي على: Merchant Name, Merchant ID, Type, (Email or To)'
    },
    1: {   # Clients / الوكلاء
        'required': ['iban', 'merchant_id', 'merchant_name'],
        'label': 'Clients File',
        'hint':  'يجب أن يحتوي على: iban, merchant_id, merchant_name, payment_amount'
    },
    2: {   # Transactions / الحركات المالية
        'required': ['Merchant ID', 'Transaction Amount', 'Merchant Commission',
                     'Terminal ID', 'Auth Code', 'Transaction Date'],
        'label': 'Transactions File',
        'hint':  'يجب أن يحتوي على: Merchant ID, Transaction Amount, Merchant Commission, '
                 'Terminal ID, Auth Code, Transaction Date'
    },
}

class AutoWaSendThread(QThread):
    set_clipboard_text = pyqtSignal(str)

    def __init__(self, msg, delay=4.0):
        super().__init__()
        self.msg = msg
        self.delay = delay

    def run(self):
        if sys.platform != 'win32': return
        import ctypes
        import time
        
        # Wait for WhatsApp to fully open and focus text field
        time.sleep(self.delay)
        
        user32 = ctypes.windll.user32
        VK_CONTROL = 0x11; VK_V = 0x56; KEYEVENTF_KEYUP = 0x0002

        # 1. Paste files (Ctrl+V)
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        user32.keybd_event(VK_V, 0, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        
        time.sleep(1.5) # Wait for attachment preview window
        
        # 2. Add text to clipboard
        self.set_clipboard_text.emit(self.msg)
        time.sleep(0.5)
        
        # 3. Paste text into caption (Ctrl+V)
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        user32.keybd_event(VK_V, 0, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)


# ── Dialogs ──
class LogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Application Log"); self.setMinimumSize(900,550)
        layout = QVBoxLayout(self); layout.setContentsMargins(14,14,14,14); layout.setSpacing(8)
        tb = QHBoxLayout()
        tb.addWidget(QLabel("Application Log")); tb.addStretch()
        bc = QPushButton("🗑️ Clear"); bc.setObjectName("btnDanger"); bc.setFixedWidth(80); bc.clicked.connect(self.clear_logs); tb.addWidget(bc)
        be = QPushButton("💾 Export"); be.setObjectName("btnAccent"); be.setFixedWidth(80); be.clicked.connect(self.export_logs); tb.addWidget(be)
        layout.addLayout(tb)
        self.log_text = QTextEdit(); self.log_text.setReadOnly(True); self.log_text.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        layout.addWidget(self.log_text)
        self.load_logs(); log_handler.callback = self.append_log
    def load_logs(self):
        self.log_text.clear()
        for r in log_handler.records: self._fmt(r)
    def append_log(self, msg): self._fmt(msg)
    def _fmt(self, msg):
        is_light = self.parent() and hasattr(self.parent(), 'current_theme') and self.parent().current_theme['name'] == 'light'
        if '| ERROR' in msg: c = '#ef4444'
        elif '| WARNING' in msg: c = '#f59e0b'
        elif '| INFO' in msg: c = '#000000' if is_light else '#10b981'
        else: c = None
        if c: self.log_text.append(f'<span style="color:{c}">{msg}</span>')
        else: self.log_text.append(f'<span>{msg}</span>')
    def clear_logs(self): log_handler.records.clear(); self.log_text.clear(); logger.info("Log cleared")
    def export_logs(self):
        p,_ = QFileDialog.getSaveFileName(self,"Export",f"log_{datetime.now():%Y%m%d_%H%M%S}.txt","Text (*.txt)")
        if p:
            with open(p,'w',encoding='utf-8') as f: f.write('\n'.join(log_handler.records))
            logger.info(f"Log exported: {p}")
    def closeEvent(self, e): log_handler.callback=None; super().closeEvent(e)


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About"); self.setMinimumSize(620, 620)
        layout = QVBoxLayout(self); layout.setContentsMargins(20,20,20,20); layout.setSpacing(10)
        t = QLabel("Merchant Report & Email Tool"); t.setObjectName("labelTitle"); t.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(t)
        v = QLabel("Version 2.4  |  ***** — Business Support"); v.setObjectName("labelSubtitle"); v.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(v)
        s = QFrame(); s.setObjectName("separator"); s.setFrameShape(QFrame.Shape.HLine); layout.addWidget(s)
        tx = QTextBrowser(); tx.setReadOnly(True)
        tx.setOpenExternalLinks(True)
        tx.setOpenLinks(True)
        tx.setStyleSheet("background: transparent; border: none;")
        tx.setHtml("""<style>
body { font-family: 'Segoe UI', sans-serif; font-size: 14px; line-height: 1.6; }
h3 { color: #3b82f6; margin-top: 15px; margin-bottom: 5px; font-size: 15px; }
b { color: #3b82f6; }
.s { margin-left: 15px; margin-bottom: 5px; }
hr { border: none; border-top: 1px solid rgba(150, 150, 150, 0.3); margin: 20px 0; }
.credits { text-align: center; padding: 10px; }
.tool-name { font-size: 18px; font-weight: bold; margin-bottom: 5px; color: #3b82f6; }
.tool-sub { font-size: 13px; margin-bottom: 15px; font-style: italic; }
.company { font-size: 15px; font-weight: bold; margin-bottom: 2px; }
.dept { font-size: 12px; margin-bottom: 12px; }
.purpose { font-size: 13px; margin-bottom: 15px; line-height: 1.6; }
.dev-name { font-size: 15px; font-weight: bold; margin-top: 10px; }
.dev-title { font-size: 12px; margin-bottom: 15px; }
.li-btn { display: inline-block; background-color: #0a66c2; color: #ffffff !important; font-size: 12px; font-weight: bold; padding: 8px 18px; border-radius: 6px; text-decoration: none; }
.copy { font-size: 11px; margin-top: 20px; margin-bottom: 2px; }
.hk { font-size: 13px; font-weight: bold; letter-spacing: 1px; color: #3b82f6; }
</style>
<h3>How to Use</h3>
<div class="s"><b>1.</b> Load three Excel/CSV files using the top panel in each tab (POS or B2B)</div>
<div class="s"><b>2.</b> Merchant table auto-populates from the Info file</div>
<div class="s"><b>3.</b> Select report format (PDF / Excel / Both)</div>
<div class="s"><b>4.</b> Click <b>Generate</b> — saved to "Customer Email Report" (POS) or "B2B Email Report" (B2B) folder</div>
<div class="s"><b>5.</b> Click <b>View</b> — previews the generated reports before sending</div>
<div class="s"><b>6.</b> Click <b>Send</b> — opens pre-filled email or WhatsApp</div>
<h3>B2B Features</h3>
<div class="s"><b>Kashif Format:</b> B2B reports are automatically formatted to match the Kashif bank settlement structure (negative sales, positive commission).</div>
<div class="s"><b>Auto-Matching:</b> Compares the calculated net amount with the payment file (amna.csv) automatically.</div>
<h3>Status Columns</h3>
<div class="s"><b>Report Status:</b> Tracks generation (Pending / Generating / Done / Failed)</div>
<div class="s"><b>Send Status:</b> Tracks send action (Pending / Opened / Failed)</div>
<h3>Buttons</h3>
<div class="s"><b>Generate:</b> Builds the merchant's settlement report in the chosen format</div>
<div class="s"><b>View:</b> Opens the generated file(s) in the OS default app</div>
<div class="s"><b>Send:</b> Email or WhatsApp, with report auto-attached</div>
<h3>Contact Methods</h3>
<div class="s"><b>Email:</b> 3-tier method: (1) Outlook COM, (2) VBScript automation, (3) mailto fallback</div>
<div class="s"><b>WA Send:</b> Tries WhatsApp Desktop first, falls back to WhatsApp Web</div>
<hr>
<div class="credits">
  <div class="tool-name">&#9881; Merchant Report &amp; Email Tool</div>
  <div class="tool-sub">Automation &amp; Workflow Simplification Tool</div>
  <div class="company">*****</div>
  <div class="dept">Business Support Department</div>
  <div class="purpose">
    Developed exclusively for <b>*****</b> — Business Support Team<br>
    to automate and streamline merchant settlement reporting workflows (POS &amp; B2B).
  </div>
  <hr>
  <div class="dev-name">*****</div>
  <div class="dev-title">  Team Leader</div>
  <div><a class="li-btn" href="*****">&#128279; LinkedIn Profile</a></div>
  <div class="copy">All Rights Reserved &copy; 2026</div>
  <div class="hk">Hyper Knowledge</div>
</div>""")
        layout.addWidget(tx)
        bc = QPushButton("❌ Close"); bc.setObjectName("btnAccent"); bc.setFixedWidth(100); bc.clicked.connect(self.close)
        cl = QHBoxLayout(); cl.addStretch(); cl.addWidget(bc); cl.addStretch(); layout.addLayout(cl)


# ── Main Window ──
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Merchant Report & Email Tool")
        self.setMinimumSize(1200,700)
        self.current_theme = DARK_THEME
        self.df_send_info = None
        self.df_clients = None
        self.df_transactions = None
        self.generated_reports: Dict[str, list] = {}
        self.active_threads: List[QThread] = []
        self.output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Customer Email Report")
        # B2B state
        self.b2b_df_info = None
        self.b2b_df_payments = None
        self.b2b_df_transactions = None
        self.b2b_generated_reports: Dict[str, list] = {}
        self.b2b_output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "B2B Email Report")
        logger.info("Application started")
        logger.info(f"Output dir: {self.output_dir}")
        self.setup_ui()
        self.apply_theme(DARK_THEME)

    def setup_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        main_layout = QVBoxLayout(central); main_layout.setContentsMargins(6,6,6,6); main_layout.setSpacing(6)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Tab 1: POS
        self.tab_pos = QWidget()
        self.tabs.addTab(self.tab_pos, "POS")

        # Tab 2: B2B
        self.tab_b2b = QWidget()
        self.tabs.addTab(self.tab_b2b, "B2B")
        self.setup_b2b_ui()

        ml = QVBoxLayout(self.tab_pos); ml.setContentsMargins(12,6,12,6); ml.setSpacing(6)

        # Toolbar
        tb = QHBoxLayout(); tb.setSpacing(6)
        tv = QVBoxLayout(); tv.setSpacing(0)
        lt = QLabel("Merchant Report Tool"); lt.setObjectName("labelTitle"); tv.addWidget(lt)
        ls = QLabel("Generate & Send Transaction Settlement Reports"); ls.setObjectName("labelSubtitle"); tv.addWidget(ls)
        tb.addLayout(tv); tb.addStretch()
        self.btn_log = QPushButton("📄 Log"); self.btn_log.setObjectName("btnToolbar"); self.btn_log.clicked.connect(self.show_log); tb.addWidget(self.btn_log)
        self.btn_about = QPushButton("ℹ️ About"); self.btn_about.setObjectName("btnToolbar"); self.btn_about.clicked.connect(self.show_about); tb.addWidget(self.btn_about)
        self.btn_theme = QPushButton("☀️ Light" if self.current_theme['name']=='dark' else "🌙 Dark"); self.btn_theme.setObjectName("btnToolbar"); self.btn_theme.clicked.connect(self.toggle_theme); tb.addWidget(self.btn_theme)
        ml.addLayout(tb)

        # File section — horizontal compact
        fg = QGroupBox(" Load Data Files")
        fl = QHBoxLayout(fg); fl.setSpacing(12); fl.setContentsMargins(10,14,10,8)
        self.file_buttons = []; self.file_progress = []; self.file_status = []
        fcfg = [("📋 Merchant Data File","Contacts & types"),("👥 Merchant Payment File","IBAN & amounts"),("💳 Pos Merchant Report File","Financial records")]
        for i,(bt,dt) in enumerate(fcfg):
            col = QVBoxLayout(); col.setSpacing(3)
            btn = QPushButton(f" {bt}"); btn.setObjectName("btnFile")
            btn.clicked.connect(lambda ch, idx=i: self.load_file(idx))
            col.addWidget(btn); self.file_buttons.append(btn)
            rb = QHBoxLayout(); rb.setSpacing(6)
            prog = QProgressBar(); prog.setValue(0); prog.setTextVisible(False); rb.addWidget(prog); self.file_progress.append(prog)
            st = QLabel("➖ Not loaded"); st.setObjectName("labelStatus"); st.setFixedWidth(110); rb.addWidget(st); self.file_status.append(st)
            col.addLayout(rb); fl.addLayout(col)
        ml.addWidget(fg)

        # Action bar
        ab = QHBoxLayout(); ab.setSpacing(8)
        lm = QLabel("Merchant List"); lm.setObjectName("labelSection"); ab.addWidget(lm)
        self.lbl_count = QLabel("👥 0 merchants"); self.lbl_count.setObjectName("labelBadge"); ab.addWidget(self.lbl_count)

        self.btn_gen_all = QPushButton("⚙️ Generate All Reports"); self.btn_gen_all.setObjectName("btnAccent"); self.btn_gen_all.setEnabled(False)
        self.btn_gen_all.clicked.connect(self.generate_all_reports); ab.addWidget(self.btn_gen_all)

        self.pb_gen_all = QProgressBar(); self.pb_gen_all.setValue(0); self.pb_gen_all.setTextVisible(True)
        self.pb_gen_all.setFormat("%v / %m"); self.pb_gen_all.hide(); self.pb_gen_all.setFixedWidth(120)
        self.pb_gen_all.setStyleSheet("QProgressBar { min-height: 18px; max-height: 18px; font-size: 11px; }")
        ab.addWidget(self.pb_gen_all)

        # Date picker — Global "To Date" (matches Streamlit's global_to_date)
        lbl_to = QLabel("📅 To:"); lbl_to.setObjectName("labelSection")
        ab.addWidget(lbl_to)
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setDisplayFormat("dd/MM/yyyy")
        self.date_to.setFixedWidth(130)
        ab.addWidget(self.date_to)

        ab.addStretch()

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("🔍 Search by Name or ID...")
        self.search_box.setFixedWidth(200)
        self.search_box.returnPressed.connect(self.search_table)
        ab.addWidget(self.search_box)
        
        self.btn_search = QPushButton("Search")
        self.btn_search.setObjectName("btnAccent")
        self.btn_search.setFixedWidth(60)
        self.btn_search.clicked.connect(self.search_table)
        ab.addWidget(self.btn_search)
        
        ab.addStretch()
        
        self.btn_clear = QPushButton("🗑️ Clear"); self.btn_clear.setObjectName("btnPurple"); self.btn_clear.setEnabled(False)
        self.btn_clear.clicked.connect(self.clear_all_files); ab.addWidget(self.btn_clear)
        self.btn_folder = QPushButton("📁 Open Reports Folder"); self.btn_folder.setObjectName("btnPurple")
        self.btn_folder.clicked.connect(self.open_reports_folder); ab.addWidget(self.btn_folder)
        ml.addLayout(ab)

        # Table — 9 columns
        self.table = HoverTableWidget()
        self._row_delegate = RowHoverDelegate(self.table, is_dark=True)
        self.table.setItemDelegate(self._row_delegate)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSortingEnabled(True)
        # Context menu for ID and Name
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        
        # Table — 12 columns (# | Name | ID | Pay Amt | Contact | Format | Match | Report | View | Send | Rep Status | Send Status)
        cols = ['#','Merchant Name','Merchant ID','Payment Amount','Contact','Format','Match','Report','View','Send','Report Status','Send Status']
        self.table.setColumnCount(len(cols)); self.table.setHorizontalHeaderLabels(cols)
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for c in range(2, len(cols)): h.setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0,35);  self.table.setColumnWidth(2,110); self.table.setColumnWidth(3,110); self.table.setColumnWidth(4,80)
        self.table.setColumnWidth(5,75);  self.table.setColumnWidth(6,200); self.table.setColumnWidth(7,100)
        self.table.setColumnWidth(8,80);  self.table.setColumnWidth(9,100); self.table.setColumnWidth(10,105)
        self.table.setColumnWidth(11,105)
        ml.addWidget(self.table, stretch=1)

        self.status_bar = QStatusBar(); self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready  —  Load the three data files to begin")
        self.lbl_brand = QLabel(" ***** "); self.lbl_brand.setObjectName("labelSubtitle")
        self.status_bar.addPermanentWidget(self.lbl_brand)

    def apply_theme(self, theme):
        self.current_theme = theme; self.setStyleSheet(build_stylesheet(theme))
        self.btn_theme.setText("☀️ Light" if theme['name']=='dark' else "🌙 Dark")
        # Fix 1: update hover delegate color for light/dark mode
        if hasattr(self, '_row_delegate'):
            self._row_delegate.set_dark(theme['name'] == 'dark')
        if hasattr(self, '_b2b_row_delegate'):
            self._b2b_row_delegate.set_dark(theme['name'] == 'dark')
        logger.info(f"Theme: {theme['name']}")

    def show_context_menu(self, pos):
        item = self.table.itemAt(pos)
        if not item: return
        col = item.column()
        # Allows copying from Name (1), ID (2), or Payment Amount (3) columns
        if col in (1, 2, 3):
            menu = QMenu(self)
            copy_action = menu.addAction("📋 Copy")
            action = menu.exec(self.table.viewport().mapToGlobal(pos))
            if action == copy_action:
                QApplication.clipboard().setText(item.text())

    def search_table(self):
        query = self.search_box.text().strip().lower()
        if not query: return
            
        found = False
        for row in range(self.table.rowCount()):
            name_item = self.table.item(row, 1)
            id_item = self.table.item(row, 2)
            if (name_item and query in name_item.text().lower()) or \
               (id_item and query in id_item.text().lower()):
                self.table.selectRow(row)
                self.table.scrollToItem(name_item)
                
                # Highlight in purple
                self.table.setStyleSheet("""
                    QTableWidget::item:selected {
                        background-color: #9333ea;
                        color: white;
                    }
                """)
                found = True
                break
        
        if not found:
            QMessageBox.information(self, "Search", "No matching results found.")

    def toggle_theme(self):
        self.apply_theme(LIGHT_THEME if self.current_theme['name']=='dark' else DARK_THEME)

    def show_log(self):
        d = LogDialog(self); d.setStyleSheet(self.styleSheet()); d.exec()

    def show_about(self):
        d = AboutDialog(self); d.setStyleSheet(self.styleSheet()); d.exec()

    def load_file(self, index):
        fp,_ = QFileDialog.getOpenFileName(self,"Select File","","Data Files (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)")
        if not fp: return
        labels = ["📋 Send Info File","👥 Clients File","💳 Transactions File"]
        self.file_progress[index].setValue(0); self.file_status[index].setText("⏳ Loading...")
        self.file_buttons[index].setEnabled(False)
        logger.info(f"Selected {labels[index]}: {os.path.basename(fp)}")
        t = FileLoaderThread(fp, labels[index])
        t.progress.connect(lambda v, idx=index: self.file_progress[idx].setValue(v))
        t.finished.connect(lambda df, label, idx=index: self.on_file_loaded(df, label, idx))
        t.error.connect(lambda msg, idx=index: self.on_file_error(msg, idx))
        t.start(); self.active_threads.append(t)

    def on_file_loaded(self, df, label, index):
        # ── Fix 6: Validate required columns ────────────────────────────────
        spec = FILE_REQUIRED_COLUMNS.get(index, {})
        required = spec.get('required', [])
        # Strip spaces from columns before checking
        actual_cols = [c.strip().lower() for c in df.columns]
        missing = [c for c in required if c.strip().lower() not in actual_cols]
        
        # Check for contact column (Email or To)
        if index == 0:
            if 'email' not in actual_cols and 'to' not in actual_cols:
                missing.append('Email or To')
                
        if missing:
            self.file_buttons[index].setEnabled(True)
            self.file_progress[index].setValue(0)
            self.file_status[index].setText("⚠️ Wrong file!")
            self.file_status[index].setObjectName("labelFileErr")
            self.file_status[index].style().unpolish(self.file_status[index])
            self.file_status[index].style().polish(self.file_status[index])
            QMessageBox.critical(
                self, "Wrong File",
                f"❌ هذا الملف ليس {spec.get('label', 'المطلوب')}.\n\n"
                f"{spec.get('hint', '')}\n\n"
                f"الأعمدة المفقودة:\n  • " + "\n  • ".join(missing)
            )
            return

        self.file_buttons[index].setEnabled(True)
        r, c = df.shape
        if index == 0:
            self.df_send_info = df
            self.file_status[index].setText(f"✅ {r} merchants")
            self.populate_table()
        elif index == 1:
            # Normalize column names (strip spaces)
            df.columns = [str(col).strip() for col in df.columns]
            if 'merchant_id' in df.columns:
                df['merchant_id'] = df['merchant_id'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            self.df_clients = df
            self.file_status[index].setText(f"✅ {r} records")
            self.populate_table()
        elif index == 2:
            self.df_transactions = df
            self.file_status[index].setText(f"✅ {r:,} txns")
        self.file_status[index].setObjectName("labelFileOK")
        self.file_status[index].style().unpolish(self.file_status[index])
        self.file_status[index].style().polish(self.file_status[index])
        self.btn_clear.setEnabled(True)
        self.update_status()

    def clear_all_files(self):
        r = QMessageBox.question(self,"Clear Files","Are you sure you want to clear all loaded files?",
            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            self.df_send_info = None; self.df_clients = None; self.df_transactions = None
            self.generated_reports.clear(); self.table.setRowCount(0)
            self.lbl_count.setText("👥 0 merchants")
            for i in range(3):
                self.file_status[i].setText("➖ Not loaded"); self.file_status[i].setObjectName("labelStatus")
                self.file_status[i].style().unpolish(self.file_status[i]); self.file_status[i].style().polish(self.file_status[i])
                self.file_progress[i].setValue(0)
            self.btn_clear.setEnabled(False)
            self.update_status()

    def on_file_error(self, msg, index):
        self.file_buttons[index].setEnabled(True); self.file_progress[index].setValue(0)
        self.file_status[index].setText("⚠️ Error!"); self.file_status[index].setObjectName("labelFileErr")
        self.file_status[index].style().unpolish(self.file_status[index])
        self.file_status[index].style().polish(self.file_status[index])
        QMessageBox.critical(self, "File Load Error", msg)

    def update_status(self):
        loaded = sum([self.df_send_info is not None, self.df_clients is not None, self.df_transactions is not None])
        self.status_bar.showMessage(f"Files: {loaded}/3  |  Reports: {len(self.generated_reports)}")
        al = all([self.df_send_info is not None, self.df_clients is not None, self.df_transactions is not None])
        self.btn_gen_all.setEnabled(al)
        self.update_gen_btns()

    def populate_table(self):
        if self.df_send_info is None: return
        self.table.setSortingEnabled(False); self.table.setRowCount(0)
        df = self.df_send_info.copy()
        
        # More robust lookup for the Type/Format column regardless of exact spaces
        type_col = None
        for col in df.columns:
            # Common names for the report format column
            if str(col).strip().lower() in ('type', 'format', 'report_type', 'report type', 'report format', 'file type', 'filetype'):
                type_col = col
                break
                
        if type_col is not None:
            df['_tn'] = df[type_col].apply(normalize_type)
        else:
            df['_tn'] = 'pdf' # Fallback
            
        # Previously here was a line dropping 'amount' — removed to allow it!

        # Dynamic Contact Column Detection (Sub-string search for flexibility)
        to_col = None
        cc_col = None
        for col in df.columns:
            c_clean = str(col).strip().lower()
            if c_clean == 'to': to_col = col
            if c_clean == 'cc': cc_col = col
        
        # If 'To' not found explicitly, fallback to old keyword search
        contact_col = to_col
        if not contact_col:
            contact_keywords = ['email', 'contact', 'whatsapp', 'phone', 'wa', 'تواصل', 'ايميل', 'link', 'رابط']
            for col in df.columns:
                c_lower = str(col).strip().lower()
                if any(k in c_lower for k in contact_keywords):
                    contact_col = col
                    break
                
        if self.df_clients is not None:
            # Create a set of normalized merchant IDs from the payment file
            valid_mids = set(self.df_clients['merchant_id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().unique())
            
            def _is_valid(mid_val):
                if pd.isna(mid_val): return False
                # Normalize the ID from the info file for comparison
                m = str(mid_val).replace('.0', '').strip()
                return m in valid_mids
                
            df = df[df['Merchant ID'].apply(_is_valid)].reset_index(drop=True)

        self.table.setRowCount(len(df))
        mc = 0
        for ri, (_, row) in enumerate(df.iterrows()):
            mc += 1
            mid = str(int(row['Merchant ID'])) if pd.notna(row['Merchant ID']) else ''
            mn = str(row['Merchant Name']).strip()
            mt = row['_tn']
            
            # Use dynamic contact column if found
            contact_val = row[contact_col] if contact_col else (row.get('Email', '') or row.get('Contact', ''))
            ct = parse_contact_info(contact_val)
            
            # Also parse CC if available
            cc_ct = None
            if cc_col:
                cc_ct = parse_contact_info(row[cc_col])
            
            self.table.setRowHeight(ri, 38)
            # #
            it = QTableWidgetItem(str(mc)); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.table.setItem(ri,0,it)
            # Name
            it = QTableWidgetItem(mn); it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            f = QFont("Segoe UI",10); f.setBold(True); it.setFont(f); self.table.setItem(ri,1,it)
            # ID
            it = QTableWidgetItem(mid); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); it.setFont(QFont("Consolas",10)); self.table.setItem(ri,2,it)
            # Pay Amount (col 3)
            pa = 0.0
            if self.df_clients is not None and 'merchant_id' in self.df_clients.columns and 'payment_amount' in self.df_clients.columns:
                _cr = self.df_clients[self.df_clients['merchant_id'] == mid]
                if not _cr.empty:
                    _pa_str = _cr['payment_amount'].astype(str).str.replace(',', '', regex=False)
                    pa = float(pd.to_numeric(_pa_str, errors='coerce').fillna(0).sum())
            it = QTableWidgetItem(f"{pa:,.2f}"); it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); it.setForeground(QColor('#3b82f6')); self.table.setItem(ri,3,it)

            # Contact (col 4)
            if ct['emails']: 
                ctx = f"Email ({len(ct['emails'])})"
                if cc_ct and cc_ct['emails']:
                    ctx += f" + Cc({len(cc_ct['emails'])})"
            elif ct['phones']: ctx = "WhatsApp"
            elif ct['is_wa_group']: ctx = "WA Group"
            elif ct['is_silent']: ctx = "Silent"
            else: ctx = "None"
            it = QTableWidgetItem(ctx); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.table.setItem(ri,4,it)
            # Format — Fix 2: use NoScrollComboBox to prevent wheel changes
            cb = NoScrollComboBox(); cb.addItems(["PDF","Excel","Both","Amount"])
            if mt=='excel': cb.setCurrentIndex(1)
            elif mt=='both': cb.setCurrentIndex(2)
            elif mt=='amount': cb.setCurrentIndex(3)
            self.table.setCellWidget(ri,5,cb)
            # Match (col 6) — shows payment vs net after generation
            it = QTableWidgetItem("⏳"); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.table.setItem(ri,6,it)
            # Generate  (col 7)
            bg = QPushButton("⚙️ Generate"); bg.setObjectName("btnGreen"); bg.setEnabled(False)
            bg.clicked.connect(lambda ch, r=ri, n=mn, i=mid: self.generate_report(r,n,i))
            self.table.setCellWidget(ri,7,bg)
            # View  (col 8)
            bv = QPushButton("👁 View"); bv.setObjectName("btnPurple"); bv.setEnabled(False)
            bv.clicked.connect(lambda ch, n=mn, i=mid: self.view_files(i, n))
            self.table.setCellWidget(ri,8,bv)
            # Send  (col 9)
            if ct['emails']:
                bs = QPushButton("📧 Email"); bs.setObjectName("btnAccent")
                bs.clicked.connect(lambda ch, n=mn, i=mid, c=ct, cc=cc_ct, r=ri: self.send_email(n,i,c,r,cc_contact=cc))
            elif (ct['phones'] or ct['wa_link']) and not ct['is_silent']:
                bs = QPushButton("💬 WA Send"); bs.setObjectName("btnWA")
                bs.clicked.connect(lambda ch, n=mn, i=mid, c=ct, r=ri: self.send_whatsapp(n,i,c,r))
            else:
                bs = QPushButton("—"); bs.setEnabled(False)
            self.table.setCellWidget(ri,9,bs)
            # Report Status  (col 10)
            it = QTableWidgetItem("⏳ Pending"); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.table.setItem(ri,10,it)
            # Send Status  (col 11)
            it = QTableWidgetItem("⏳ Pending"); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.table.setItem(ri,11,it)
        self.table.setSortingEnabled(True)
        self.lbl_count.setText(f"👥 {mc} merchants")
        self.update_gen_btns()
        logger.info(f"Table: {mc} merchants (excl. amount)")

    def update_gen_btns(self):
        al = all([self.df_send_info is not None, self.df_clients is not None, self.df_transactions is not None])
        for r in range(self.table.rowCount()):
            b = self.table.cellWidget(r,7)  # col 7 = Generate button
            if b: b.setEnabled(al)

    def generate_report(self, row_idx, merchant_name, merchant_id):
        if self.df_transactions is None:
            QMessageBox.warning(self,"Missing Data","Load Transactions file first."); return
        cb = self.table.cellWidget(row_idx,5)
        rf = cb.currentText().lower()
        it = self.table.item(row_idx,10)  # col 10 = Report Status
        if it: it.setText("⚙️ Generating...")
        # Get to_date from the date picker
        to_date = self.date_to.date().toPyDate()
        logger.info(f"Report: {merchant_name} ({merchant_id}) - {rf} - to_date={to_date}")
        t = ReportGeneratorThread(merchant_name, merchant_id, rf, self.df_transactions, self.df_clients, self.output_dir, to_date)
        
        def _update_prog(v, r=row_idx):
            self.table.setSortingEnabled(False)
            it = self.table.item(r, 10)
            if it: it.setText(f"⚙️ Gen {v}%")
            self.table.setSortingEnabled(True)
            
        t.progress.connect(_update_prog)
        t.finished.connect(lambda n, f, d, r=row_idx, mi=merchant_id: self.on_report_done(n,f,d,r,mi))
        t.error.connect(lambda m, r=row_idx: self.on_report_err(m,r))
        t.match_result.connect(lambda mid, pa, nt, m, r=row_idx: self.on_match_result(r, pa, nt, m))
        t.start(); self.active_threads.append(t)

    def _increment_gen_pb(self):
        if hasattr(self, 'gen_all_total') and self.pb_gen_all.isVisible():
            self.gen_all_completed += 1; self.pb_gen_all.setValue(self.gen_all_completed)
            if self.gen_all_completed >= self.gen_all_total:
                QTimer.singleShot(2000, self.pb_gen_all.hide)

    def on_report_done(self, name, files, date_range, row_idx, merchant_id=None):
        # FIX email-attachment bug: use the merchant_id passed from generate_report's
        # lambda as the storage key. Previously we used the cell text at row_idx,
        # but if the user sorted the table between Generate and the thread finishing,
        # row_idx would point to a different merchant and we'd store under the wrong key.
        if merchant_id is not None:
            key = str(merchant_id)
        else:
            mi = self.table.item(row_idx, 2)
            key = mi.text() if mi else None
        if key:
            self.generated_reports[key] = files
            self.generated_date_ranges = getattr(self, 'generated_date_ranges', {})
            self.generated_date_ranges[key] = date_range
            logger.info(f"Stored {len(files)} file(s) under key '{key}'")
        # Enable buttons for this row
        bg = self.table.cellWidget(row_idx, 7)
        if bg: bg.setEnabled(True)
        bv = self.table.cellWidget(row_idx, 8)
        if bv: bv.setEnabled(True)
        bs = self.table.cellWidget(row_idx, 9)
        if bs: bs.setEnabled(True)
        # Update Report Status  (col 10)
        self.table.setSortingEnabled(False)
        it = self.table.item(row_idx, 10)
        if it:
            has_pdf = any(f.lower().endswith('.pdf')  for f in files)
            has_xl  = any(f.lower().endswith('.xlsx') for f in files)
            if has_pdf and has_xl: label = "✅ PDF + Excel"
            elif has_pdf:          label = "✅ PDF"
            elif has_xl:           label = "✅ Excel"
            else:                  label = "✅ Done"
            it.setText(label)
        self.table.setSortingEnabled(True)
        self.update_status()
        self.status_bar.showMessage(f"Report: {name} ({len(files)} file(s))", 5000)
        self._increment_gen_pb()

    def on_report_err(self, msg, row_idx):
        self.table.setSortingEnabled(False)
        bg = self.table.cellWidget(row_idx, 7)
        if bg: bg.setEnabled(True)
        it = self.table.item(row_idx, 10)  # col 10 = Report Status
        if it: it.setText("❌ Failed")
        self.table.setSortingEnabled(True)
        QMessageBox.warning(self,"Report Error", msg)
        self._increment_gen_pb()

    def on_match_result(self, row_idx, payment_amount, net_total, matched):
        """Update the Match column (col 6) with payment vs net comparison."""
        self.table.setSortingEnabled(False)
        it = self.table.item(row_idx, 6)
        if it:
            if matched:
                it.setText(f"✅ {net_total:,.2f} = {payment_amount:,.2f}")
                it.setForeground(QColor('#10b981'))
            else:
                it.setText(f"❌ {net_total:,.2f} ≠ {payment_amount:,.2f}")
                it.setForeground(QColor('#ef4444'))
        self.table.setSortingEnabled(True)

    def generate_all_reports(self):
        if not all([self.df_send_info is not None, self.df_clients is not None, self.df_transactions is not None]):
            QMessageBox.warning(self,"Missing Data","Load all files first."); return
        cnt = self.table.rowCount()
        r = QMessageBox.question(self,"Generate All",f"Generate for all {cnt} merchants?",
            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            logger.info(f"Batch: {cnt} merchants")
            self.gen_all_total = cnt; self.gen_all_completed = 0
            self.pb_gen_all.setMaximum(cnt); self.pb_gen_all.setValue(0); self.pb_gen_all.show()
            for r in range(cnt):
                ni = self.table.item(r,1); ii = self.table.item(r,2)
                if ni and ii:
                    QTimer.singleShot(r*500, lambda r=r,n=ni.text(),i=ii.text(): self.generate_report(r,n,i))

    def _get_email_data(self, merchant_id, is_b2b=False):
        settled = 0
        # Select correct dataframes based on tab
        clients_df = self.b2b_df_payments if is_b2b else self.df_clients
        trans_df = self.b2b_df_transactions if is_b2b else self.df_transactions
        
        # Robust string-based matching for Merchant ID
        mid_str = str(merchant_id).strip()
        
        if clients_df is not None:
            # Match using the same logic as Table population
            mask = clients_df['merchant_id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() == mid_str
            cr = clients_df[mask]
            if not cr.empty:
                _pa = cr['payment_amount'].astype(str).str.replace(',', '', regex=False)
                settled = float(pd.to_numeric(_pa, errors='coerce').fillna(0).sum())

        dr = "N/A"
        if trans_df is not None:
            mask_t = trans_df['Merchant ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() == mid_str
            txn = trans_df[mask_t]
            if not txn.empty:
                date_col = 'Transaction Date' if 'Transaction Date' in txn.columns else (
                           'Process Date' if 'Process Date' in txn.columns else None)
                if date_col:
                    try:
                        dmin = pd.to_datetime(txn[date_col]).min()
                        dmax = pd.to_datetime(txn[date_col]).max()
                        dr = f"From {dmin.strftime('%d-%m-%Y')} To {dmax.strftime('%d-%m-%Y')}"
                    except Exception as e:
                        logger.warning(f"Date range parse failed: {e}")
        return settled, dr

    def view_files(self, merchant_id, merchant_name):
        """Open generated report file(s) in the OS's default application."""
        key = str(merchant_id)
        files = self.generated_reports.get(key, [])
        if not files:
            QMessageBox.information(
                self, "No Files",
                f"No report generated yet for {merchant_name}.\n"
                f"Click 'Generate' first."
            )
            return
        opened = 0
        for fp in files:
            if not os.path.exists(fp):
                logger.warning(f"File missing on disk: {fp}")
                continue
            try:
                if sys.platform == 'win32':
                    os.startfile(fp)
                elif sys.platform == 'darwin':
                    subprocess.run(['open', fp])
                else:
                    subprocess.run(['xdg-open', fp])
                logger.info(f"Opened: {fp}")
                opened += 1
            except Exception as e:
                logger.error(f"Open failed for {fp}: {e}")
                QMessageBox.warning(self, "Open Failed",
                    f"Couldn't open {os.path.basename(fp)}:\n{e}")
        if opened:
            self.status_bar.showMessage(f"Opened {opened} file(s) for {merchant_name}", 4000)

    def _outlook_compose_via_vbs(self, to, subject, body, files, cc="", is_html=False):
        """
        Open an Outlook Classic compose window with real attachments using a
        generated VBScript that drives Outlook COM. Works on any Windows
        machine where Outlook Classic is installed — does NOT depend on
        pywin32/win32com. Returns True if wscript.exe was launched.

        We write the script as UTF-16 LE with a BOM so VBScript's default
        parser handles non-ASCII characters (merchant names, Arabic) correctly.
        """
        if sys.platform != 'win32':
            return False
        import tempfile

        def esc(s):
            """Escape a VBS string literal: double-up embedded quotes."""
            return str(s).replace('"', '""')

        body_prop = "HTMLBody" if is_html else "Body"

        if is_html:
            # For HTML body, keep as a single string — do NOT split on newlines
            # because HTML ignores \n and needs <br> tags instead.
            body_vbs = '"' + esc(body).replace('\n', '') + '"'
        else:
            # Convert Python '\n' newlines into VBS `vbCrLf` concatenations so
            # the email body keeps its line breaks.
            body_parts = esc(body).split('\n')
            body_vbs = ' & vbCrLf & '.join(f'"{p}"' for p in body_parts) if body_parts else '""'

        lines = [
            'On Error Resume Next',
            'Set oOutlook = CreateObject("Outlook.Application")',
            'If Err.Number <> 0 Then WScript.Quit 1',
            'Set oMail = oOutlook.CreateItem(0)',
            f'oMail.To = "{esc(to)}"',
        ]
        if cc:
            lines.append(f'oMail.CC = "{esc(cc)}"')
        
        lines.extend([
            f'oMail.Subject = "{esc(subject)}"',
            f'oMail.{body_prop} = {body_vbs}',
        ])
        
        for fp in files:
            lines.append(f'oMail.Attachments.Add "{esc(fp)}"')
        lines.append('oMail.Display')

        script = '\r\n'.join(lines)

        # Write as UTF-16 LE with BOM — required for VBScript to handle
        # non-ASCII strings (Arabic merchant names, accented characters).
        fd, path = tempfile.mkstemp(suffix='.vbs', prefix='mtool_email_')
        os.close(fd)
        try:
            with open(path, 'wb') as f:
                f.write(b'\xff\xfe')               # UTF-16 LE BOM
                f.write(script.encode('utf-16-le'))
            subprocess.Popen(['wscript.exe', path], shell=False)
            logger.info(f"VBS Outlook opener launched ({len(files)} attachment(s))")
            # Clean up the temp script after 60 seconds
            QTimer.singleShot(60000, lambda p=path: os.remove(p) if os.path.exists(p) else None)
            return True
        except Exception as e:
            logger.error(f"VBS launch failed: {e}")
            try: os.remove(path)
            except: pass
            return False

    def send_email(self, name, mid, contact, row_idx, cc_contact=None):
        logger.info(f"Email: {name} ({mid})")
        settled, dr_fallback = self._get_email_data(mid)
        # Use exact Date Range computed during report generation
        dr = getattr(self, 'generated_date_ranges', {}).get(str(mid), dr_fallback)
        to = ';'.join(contact['emails'])
        cc_str = ';'.join(cc_contact['emails']) if cc_contact and cc_contact['emails'] else ""
        
        subj = f"Transaction Settlement Report - {name}"
        # Format the date range for the email body to match user's request exactly
        dr_body = dr.replace('From ', 'From: ').replace(' To ', '\nTo: ')
        
        body = (f"Dear Merchant,\n\n"
                f"Please find attached your transaction settlement report.\n\n"
                f"Merchant Details\n"
                f"Merchant Name: {name}\n"
                f"Merchant ID: {mid}\n\n"
                f"Settlement Period\n"
                f"{dr_body}\n\n"
                f"Settled Amount\n"
                f"IQD {settled:,.2f}\n\n"
                f"Best regards,")
        # HTML Body for Outlook to prevent extra spacing and improve organization
        html_body = (
            f"<div style='font-family: Calibri, sans-serif; font-size: 11pt; color: #000;'>"
            f"Dear Merchant,<br><br>"
            f"Please find attached your transaction settlement report.<br><br>"
            f"<b>Merchant Details</b><br>"
            f"Merchant Name: {name}<br>"
            f"Merchant ID: {mid}<br><br>"
            f"<b>Settlement Period</b><br>"
            f"{dr_body.replace('\n', '<br>')}<br><br>"
            f"<b>Settled Amount</b><br>"
            f"IQD {settled:,.2f}<br><br>"
            f"Best regards,"
            f"</div>"
        )
        rfiles = self.generated_reports.get(str(mid), [])
        valid_files = [os.path.abspath(f) for f in rfiles if os.path.exists(f)]
        logger.info(f"Email {name}: {len(valid_files)}/{len(rfiles)} file(s) on disk")

        sent = False
        attached = 0

        # ── Method 1: Direct Outlook COM via pywin32 (fastest) ─────────────
        if HAS_OUTLOOK:
            try:
                ol = win32com.client.Dispatch('Outlook.Application')
                mail = ol.CreateItem(0)
                mail.To = to
                if cc_str:
                    mail.CC = cc_str
                mail.Subject = subj
                mail.HTMLBody = html_body
                for fp in valid_files:
                    try:
                        mail.Attachments.Add(Source=fp)
                        attached += 1
                        logger.info(f"  Attached (COM): {os.path.basename(fp)}")
                    except Exception as ae:
                        logger.error(f"  COM attach failed for {fp}: {ae}")
                mail.Display()
                sent = True
                logger.info(f"Outlook opened via COM: {attached}/{len(valid_files)} attached")
            except Exception as e:
                logger.warning(f"COM method failed, falling back to VBS: {e}")

        # ── Method 2: VBS → Outlook Classic (guaranteed on Windows+Outlook) ─
        if not sent and sys.platform == 'win32':
            if self._outlook_compose_via_vbs(to, subj, html_body, valid_files, cc=cc_str, is_html=True):
                attached = len(valid_files)
                sent = True
                logger.info(f"Outlook opened via VBS: {attached} attached")

        # ── Method 3: Last-resort mailto: (no attachment support) ──────────
        if not sent:
            mu = (f"mailto:{urllib.parse.quote(to)}"
                  f"?subject={urllib.parse.quote(subj)}"
                  f"&body={urllib.parse.quote(body)}")
            if cc_str:
                mu += f"&cc={urllib.parse.quote(cc_str)}"
            try:
                QDesktopServices.openUrl(QUrl(mu))
                sent = True
                logger.info(f"mailto: opened for {name} (no attachment support)")
            except Exception as e:
                logger.error(f"Email failed: {e}")
                QMessageBox.critical(self, "Email Error", str(e))

        # ── Status feedback ────────────────────────────────────────────────
        if sent and attached > 0:
            self.status_bar.showMessage(
                f"✉ Email opened for {name} — {attached} file(s) attached", 6000)
        elif sent:
            self.status_bar.showMessage(f"✉ Email opened for {name}", 5000)

        self.table.setSortingEnabled(False)
        it = self.table.item(row_idx, 11)   # col 11 = Send Status
        if it:
            if sent and attached > 0:
                it.setText(f"✅ Sent ({attached})")
            elif sent:
                it.setText("✅ Opened")
            else:
                it.setText("❌ Failed")
        self.table.setSortingEnabled(True)

    def send_whatsapp(self, name, mid, contact, row_idx):
        logger.info(f"WhatsApp: {name} ({mid})")
        wa_link = contact.get('wa_link')
        pc = ""
        if not contact['phones'] and not wa_link:
            QMessageBox.warning(self,"No Contact","No phone number or group link found."); return
            
        if contact['phones']:
            phone = contact['phones'][0]
            if not phone.startswith('+'): phone = '+' + phone
            pc = re.sub(r'[^\d]','', phone)
            
        settled, dr_fallback = self._get_email_data(mid)
        # Use exact Date Range computed during report generation
        dr = getattr(self, 'generated_date_ranges', {}).get(str(mid), dr_fallback)
        # Format the date range for the WA body to match user's request exactly
        dr_body = dr.replace('From ', 'From: ').replace(' To ', '\nTo: ')
        
        msg = (f"Dear Merchant,\n\n"
               f"Please find attached your transaction settlement report.\n\n"
               f"Merchant Details\n"
               f"Merchant Name: {name}\n"
               f"Merchant ID: {mid}\n\n"
               f"Settlement Period\n"
               f"{dr_body}\n\n"
               f"Settled Amount\n"
               f"IQD {settled:,.2f}\n\n"
               f"Best regards,")
        em = urllib.parse.quote(msg)
        has_attachment = False
        rfiles = self.generated_reports.get(str(mid), [])
        if rfiles:
            from PyQt6.QtCore import QMimeData, QUrl
            mime = QMimeData()
            urls = [QUrl.fromLocalFile(os.path.abspath(f)) for f in rfiles if os.path.exists(f)]
            if urls:
                mime.setUrls(urls)
                QApplication.clipboard().setMimeData(mime)
                has_attachment = True

        if contact['wa_link']:
            # Convert https link to whatsapp:// protocol for groups
            invite_code = contact['wa_link'].split('/')[-1]
            wa_desktop = f"whatsapp://chat?code={invite_code}"
            wa_web = contact['wa_link']
            has_attachment = True # Force clipboard copy for groups
        elif has_attachment:
            wa_desktop = f"whatsapp://send?phone={pc}"
            wa_web = f"https://web.whatsapp.com/send?phone={pc}"
        else:
            wa_desktop = f"whatsapp://send?phone={pc}&text={em}"
            wa_web = f"https://web.whatsapp.com/send?phone={pc}&text={em}"

        sent = False
        is_web = False
        # Try Desktop app first
        try:
            if sys.platform == 'win32':
                os.startfile(wa_desktop)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', wa_desktop])
            else:
                subprocess.Popen(['xdg-open', wa_desktop])
            log_target = phone if contact['phones'] else "Group Link"
            logger.info(f"WA Desktop opened: {name} ({log_target})")
            sent = True
        except Exception as e:
            logger.warning(f"WA Desktop failed: {e}")
            
        # Fallback to Web
        if not sent:
            try:
                webbrowser.open(wa_web)
                log_target = phone if contact['phones'] else "Group Link"
                logger.info(f"WA Web opened: {name} ({log_target})")
                sent = True
                is_web = True
            except Exception as e:
                logger.error(f"WA Web failed: {e}")
                QMessageBox.critical(self,"WhatsApp Error",str(e))
        
        # If it's a group link or attachment, trigger the auto-paste helper
        if sent and (has_attachment or contact['wa_link']):
            # 7 seconds for desktop, 12 for web to be safe
            delay = 12.0 if is_web else 7.0
            self.auto_wa = AutoWaSendThread(msg=msg, delay=delay)
            self.auto_wa.set_clipboard_text.connect(lambda t: QApplication.clipboard().setText(t))
            self.auto_wa.start()
        
        self.table.setSortingEnabled(False)
        it = self.table.item(row_idx, 11)  # col 11 = Send Status
        if it: it.setText("✅ Opened" if sent else "❌ Failed")
        self.table.setSortingEnabled(True)

    def open_reports_folder(self):
        os.makedirs(self.output_dir, exist_ok=True)
        try:
            if sys.platform=='win32': os.startfile(self.output_dir)
            elif sys.platform=='darwin': subprocess.run(['open',self.output_dir])
            else: subprocess.run(['xdg-open',self.output_dir])
            logger.info(f"Opened: {self.output_dir}")
        except Exception as e: logger.error(f"Open folder failed: {e}")

    # ══════════════════════════════════════════════════════════════════════
    # ██  B2B  TAB  ██
    # ══════════════════════════════════════════════════════════════════════
    def setup_b2b_ui(self):
        ml = QVBoxLayout(self.tab_b2b); ml.setContentsMargins(12,6,12,6); ml.setSpacing(6)
        # Toolbar
        tb = QHBoxLayout(); tb.setSpacing(6)
        tv = QVBoxLayout(); tv.setSpacing(0)
        lt = QLabel("B2B Settlement Verification"); lt.setObjectName("labelTitle"); tv.addWidget(lt)
        ls = QLabel("Generate, Verify & Send B2B Transaction Settlement Reports"); ls.setObjectName("labelSubtitle"); tv.addWidget(ls)
        tb.addLayout(tv); tb.addStretch()
        self.btn_b2b_log = QPushButton("📄 Log"); self.btn_b2b_log.setObjectName("btnToolbar"); self.btn_b2b_log.clicked.connect(self.show_log); tb.addWidget(self.btn_b2b_log)
        self.btn_b2b_about = QPushButton("ℹ️ About"); self.btn_b2b_about.setObjectName("btnToolbar"); self.btn_b2b_about.clicked.connect(self.show_about); tb.addWidget(self.btn_b2b_about)
        self.btn_b2b_theme = QPushButton("☀️ Light" if self.current_theme['name']=='dark' else "🌙 Dark"); self.btn_b2b_theme.setObjectName("btnToolbar"); self.btn_b2b_theme.clicked.connect(self.toggle_theme); tb.addWidget(self.btn_b2b_theme)
        ml.addLayout(tb)
        # File section
        fg = QGroupBox(" Load B2B Data Files")
        fl = QHBoxLayout(fg); fl.setSpacing(12); fl.setContentsMargins(10,14,10,8)
        self.b2b_file_buttons = []; self.b2b_file_progress = []; self.b2b_file_status = []
        fcfg = [("📋 Merchant Info File","Contacts & emails (.xlsx)"),("💰 Payment File","Payment amounts (.csv)"),("💳 HyperPay Transactions File","Transactions (.xlsx)")]
        for i,(bt,dt) in enumerate(fcfg):
            col = QVBoxLayout(); col.setSpacing(3)
            btn = QPushButton(f" {bt}"); btn.setObjectName("btnFile")
            btn.clicked.connect(lambda ch, idx=i: self.load_b2b_file(idx))
            col.addWidget(btn); self.b2b_file_buttons.append(btn)
            rb = QHBoxLayout(); rb.setSpacing(6)
            prog = QProgressBar(); prog.setValue(0); prog.setTextVisible(False); rb.addWidget(prog); self.b2b_file_progress.append(prog)
            st = QLabel("➖ Not loaded"); st.setObjectName("labelStatus"); st.setFixedWidth(110); rb.addWidget(st); self.b2b_file_status.append(st)
            col.addLayout(rb); fl.addLayout(col)
        ml.addWidget(fg)
        
        # Action bar
        ab = QHBoxLayout(); ab.setSpacing(8)
        btn_gen_all = QPushButton("⚙️ Generate All Reports"); btn_gen_all.setObjectName("btnAccent")
        btn_gen_all.clicked.connect(self.generate_all_b2b_reports); ab.addWidget(btn_gen_all)
        
        lbl_to = QLabel("📅 To:"); lbl_to.setObjectName("labelSection")
        ab.addWidget(lbl_to)
        self.b2b_date_to = QDateEdit()
        self.b2b_date_to.setCalendarPopup(True)
        self.b2b_date_to.setDate(QDate.currentDate())
        self.b2b_date_to.setDisplayFormat("dd/MM/yyyy")
        self.b2b_date_to.setFixedWidth(130)
        ab.addWidget(self.b2b_date_to)
        
        ab.addStretch()
        lm = QLabel("Merchant List"); lm.setObjectName("labelSection"); ab.addWidget(lm)
        self.b2b_lbl_count = QLabel("👥 0 merchants"); self.b2b_lbl_count.setObjectName("labelBadge"); ab.addWidget(self.b2b_lbl_count)
        ab.addStretch()
        
        self.b2b_search_box = QLineEdit(); self.b2b_search_box.setPlaceholderText("🔍 Search by Name or ID..."); self.b2b_search_box.setFixedWidth(200)
        self.b2b_search_box.returnPressed.connect(self.search_b2b_table); ab.addWidget(self.b2b_search_box)
        b2b_bs = QPushButton("Search"); b2b_bs.setObjectName("btnAccent"); b2b_bs.setFixedWidth(60); b2b_bs.clicked.connect(self.search_b2b_table); ab.addWidget(b2b_bs)
        
        ab.addStretch()
        self.b2b_btn_clear = QPushButton("🗑️ Clear"); self.b2b_btn_clear.setObjectName("btnPurple"); self.b2b_btn_clear.setEnabled(False)
        self.b2b_btn_clear.clicked.connect(self.clear_b2b_files); ab.addWidget(self.b2b_btn_clear)
        self.b2b_btn_folder = QPushButton("📁 Open B2B Reports"); self.b2b_btn_folder.setObjectName("btnPurple")
        self.b2b_btn_folder.clicked.connect(self.open_b2b_reports_folder); ab.addWidget(self.b2b_btn_folder)
        ml.addLayout(ab)
        
        # Table — 11 columns
        self.b2b_table = HoverTableWidget()
        self._b2b_row_delegate = RowHoverDelegate(self.b2b_table, is_dark=True)
        self.b2b_table.setItemDelegate(self._b2b_row_delegate)
        self.b2b_table.setAlternatingRowColors(True)
        self.b2b_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.b2b_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.b2b_table.verticalHeader().setVisible(False)
        self.b2b_table.setShowGrid(False)
        self.b2b_table.horizontalHeader().setStretchLastSection(True)
        self.b2b_table.setSortingEnabled(True)
        self.b2b_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.b2b_table.customContextMenuRequested.connect(self.show_b2b_context_menu)
        
        # Table — 12 columns
        cols = ['#','Merchant Name','Merchant ID','Payment Amount','Contact','Format','Match','Report','View','Send','Report Status','Send Status']
        self.b2b_table.setColumnCount(len(cols)); self.b2b_table.setHorizontalHeaderLabels(cols)
        h = self.b2b_table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for c in range(2, len(cols)): h.setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
        self.b2b_table.setColumnWidth(0,35);  self.b2b_table.setColumnWidth(2,110); self.b2b_table.setColumnWidth(3,110); self.b2b_table.setColumnWidth(4,80)
        self.b2b_table.setColumnWidth(5,75);  self.b2b_table.setColumnWidth(6,200); self.b2b_table.setColumnWidth(7,100)
        self.b2b_table.setColumnWidth(8,80);  self.b2b_table.setColumnWidth(9,100); self.b2b_table.setColumnWidth(10,105)
        self.b2b_table.setColumnWidth(11,105)
        ml.addWidget(self.b2b_table, stretch=1)

    def open_b2b_reports_folder(self):
        os.makedirs(self.b2b_output_dir, exist_ok=True)
        try:
            if sys.platform=='win32': os.startfile(self.b2b_output_dir)
            elif sys.platform=='darwin': subprocess.run(['open',self.b2b_output_dir])
            else: subprocess.run(['xdg-open',self.b2b_output_dir])
            logger.info(f"Opened: {self.b2b_output_dir}")
        except Exception as e: logger.error(f"Open folder failed: {e}")

    def show_b2b_context_menu(self, pos):
        item = self.b2b_table.itemAt(pos)
        if not item: return
        if item.column() in (1, 2, 3):
            menu = QMenu(self)
            copy_action = menu.addAction("📋 Copy")
            action = menu.exec(self.b2b_table.viewport().mapToGlobal(pos))
            if action == copy_action:
                QApplication.clipboard().setText(item.text())

    def search_b2b_table(self):
        query = self.b2b_search_box.text().strip().lower()
        if not query: return
        for row in range(self.b2b_table.rowCount()):
            ni = self.b2b_table.item(row, 1); ii = self.b2b_table.item(row, 2)
            if (ni and query in ni.text().lower()) or (ii and query in ii.text().lower()):
                self.b2b_table.selectRow(row); self.b2b_table.scrollToItem(ni); return
        QMessageBox.information(self, "Search", "No matching results found.")

    def load_b2b_file(self, index):
        # Allow all types for all B2B files
        all_filters = "Data Files (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)"
        labels = ["📋 Info File", "💰 Payment File", "💳 Transactions File"]
        fp, _ = QFileDialog.getOpenFileName(self, f"Select {labels[index]}", "", all_filters)
        if not fp: return
        self.b2b_file_progress[index].setValue(0); self.b2b_file_status[index].setText("⏳ Loading...")
        self.b2b_file_buttons[index].setEnabled(False)
        logger.info(f"B2B: Selected {labels[index]}: {os.path.basename(fp)}")
        if index == 0:
            t = FileLoaderThread(fp, labels[0])
            t.progress.connect(lambda v: self.b2b_file_progress[0].setValue(v))
            t.finished.connect(lambda df, lbl: self._on_b2b_info_loaded(df))
            t.error.connect(lambda msg: self._on_b2b_file_error(msg, 0))
            t.start(); self.active_threads.append(t)
        elif index == 1:
            t = CsvLoaderThread(fp, labels[1])
            t.progress.connect(lambda v: self.b2b_file_progress[1].setValue(v))
            t.finished.connect(lambda df, lbl: self._on_b2b_csv_loaded(df))
            t.error.connect(lambda msg: self._on_b2b_file_error(msg, 1))
            t.start(); self.active_threads.append(t)
        elif index == 2:
            t = FileLoaderThread(fp, labels[2])
            t.progress.connect(lambda v: self.b2b_file_progress[2].setValue(v))
            t.finished.connect(lambda df, lbl: self._on_b2b_transactions_loaded(df))
            t.error.connect(lambda msg: self._on_b2b_file_error(msg, 2))
            t.start(); self.active_threads.append(t)

    def _on_b2b_info_loaded(self, df):
        self.b2b_file_buttons[0].setEnabled(True)
        # Check for required columns, allowing To or Email
        actual = [c.strip().lower() for c in df.columns]
        req = ['merchant name', 'merchant id']
        missing = [c for c in req if c not in actual]
        
        if 'email' not in actual and 'to' not in actual:
            missing.append('Email or To')
            
        if missing:
            self.b2b_file_progress[0].setValue(0); self.b2b_file_status[0].setText("⚠️ Wrong file!")
            self.b2b_file_status[0].setObjectName("labelFileErr")
            self.b2b_file_status[0].style().unpolish(self.b2b_file_status[0]); self.b2b_file_status[0].style().polish(self.b2b_file_status[0])
            QMessageBox.critical(self, "Wrong File", f"Missing columns:\n• " + "\n• ".join(missing)); return
            
        self.b2b_df_info = df
        self.b2b_file_status[0].setText(f"✅ {len(df)} merchants")
        self.b2b_file_status[0].setObjectName("labelFileOK")
        self.b2b_file_status[0].style().unpolish(self.b2b_file_status[0]); self.b2b_file_status[0].style().polish(self.b2b_file_status[0])
        self.b2b_btn_clear.setEnabled(True)
        self._try_populate_b2b()

    def _on_b2b_csv_loaded(self, df):
        self.b2b_file_buttons[1].setEnabled(True)
        df.columns = [str(c).strip() for c in df.columns]
        if 'merchant_id' not in df.columns or 'payment_amount' not in df.columns:
            self.b2b_file_progress[1].setValue(0); self.b2b_file_status[1].setText("⚠️ Wrong file!")
            self.b2b_file_status[1].setObjectName("labelFileErr")
            self.b2b_file_status[1].style().unpolish(self.b2b_file_status[1]); self.b2b_file_status[1].style().polish(self.b2b_file_status[1])
            QMessageBox.critical(self, "Wrong File", "Missing: merchant_id or payment_amount"); return
            
        # Safely cast merchant_id to string so it matches `mid` correctly
        df['merchant_id'] = df['merchant_id'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        self.b2b_df_payments = df
        self.b2b_file_status[1].setText(f"✅ {len(df):,} records")
        self.b2b_file_status[1].setObjectName("labelFileOK")
        self.b2b_file_status[1].style().unpolish(self.b2b_file_status[1]); self.b2b_file_status[1].style().polish(self.b2b_file_status[1])
        self.b2b_btn_clear.setEnabled(True)
        self._try_populate_b2b()

    def _on_b2b_transactions_loaded(self, df):
        self.b2b_file_buttons[2].setEnabled(True)
        self.b2b_df_transactions = df
        self.b2b_file_status[2].setText(f"✅ {len(df):,} records")
        self.b2b_file_status[2].setObjectName("labelFileOK")
        self.b2b_file_status[2].style().unpolish(self.b2b_file_status[2]); self.b2b_file_status[2].style().polish(self.b2b_file_status[2])
        self.b2b_btn_clear.setEnabled(True)
        self._try_populate_b2b()

    def _on_b2b_file_error(self, msg, index):
        self.b2b_file_buttons[index].setEnabled(True); self.b2b_file_progress[index].setValue(0)
        self.b2b_file_status[index].setText("⚠️ Error!"); self.b2b_file_status[index].setObjectName("labelFileErr")
        self.b2b_file_status[index].style().unpolish(self.b2b_file_status[index]); self.b2b_file_status[index].style().polish(self.b2b_file_status[index])
        QMessageBox.critical(self, "File Error", msg)

    def _try_populate_b2b(self):
        if self.b2b_df_info is not None:
            self._populate_b2b_table()

    def _populate_b2b_table(self):
        self.b2b_table.setSortingEnabled(False); self.b2b_table.setRowCount(0)
        df_info = self.b2b_df_info.copy()
        
        # Filter: Only show merchants present in both Info File and Payments File
        if self.b2b_df_payments is not None:
            valid_mids = set(self.b2b_df_payments['merchant_id'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().unique())
            
            def _is_valid_b2b(mid_val):
                if pd.isna(mid_val): return False
                m = str(mid_val).replace('.0', '').strip()
                return m in valid_mids
                
            df_info = df_info[df_info['Merchant ID'].apply(_is_valid_b2b)].reset_index(drop=True)
        
        mc = 0
        self.b2b_table.setRowCount(len(df_info))
        
        # Detect To and Cc columns for B2B
        b2b_to_col = None
        b2b_cc_col = None
        for col in df_info.columns:
            c_clean = str(col).strip().lower()
            if c_clean == 'to': b2b_to_col = col
            if c_clean == 'cc': b2b_cc_col = col
            
        for ri, (_, row) in enumerate(df_info.iterrows()):
            mc += 1
            mid = str(int(row['Merchant ID'])) if pd.notna(row['Merchant ID']) else ''
            mn = str(row['Merchant Name']).strip()
            
            # Use detected To column or fallback to 'Email'
            to_val = row[b2b_to_col] if b2b_to_col else row.get('Email', '')
            ct = parse_contact_info(to_val)
            
            # Also parse CC if available
            cc_ct = None
            if b2b_cc_col:
                cc_ct = parse_contact_info(row[b2b_cc_col])
                
            self.b2b_table.setRowHeight(ri, 38)
            # Col 0: #
            it = QTableWidgetItem(str(mc)); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.b2b_table.setItem(ri,0,it)
            # Col 1: Name
            it = QTableWidgetItem(mn); it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable)
            f = QFont("Segoe UI",10); f.setBold(True); it.setFont(f); self.b2b_table.setItem(ri,1,it)
            # Col 2: ID
            it = QTableWidgetItem(mid); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); it.setFont(QFont("Consolas",10)); self.b2b_table.setItem(ri,2,it)
            # Col 3: Pay Amount
            pa = 0.0
            if self.b2b_df_payments is not None and 'merchant_id' in self.b2b_df_payments.columns and 'payment_amount' in self.b2b_df_payments.columns:
                _cr = self.b2b_df_payments[self.b2b_df_payments['merchant_id'] == mid]
                if not _cr.empty:
                    _pa_str = _cr['payment_amount'].astype(str).str.replace(',', '', regex=False)
                    pa = float(pd.to_numeric(_pa_str, errors='coerce').fillna(0).sum())
            it = QTableWidgetItem(f"{pa:,.2f}"); it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); it.setForeground(QColor('#3b82f6')); self.b2b_table.setItem(ri,3,it)

            # Col 4: Contact
            if ct['emails']: 
                ctx = f"Email ({len(ct['emails'])})"
                if cc_ct and cc_ct['emails']:
                    ctx += f" + Cc({len(cc_ct['emails'])})"
            elif ct['phones']: ctx = "WhatsApp"
            elif ct['is_wa_group']: ctx = "WA Group"
            elif ct['is_silent']: ctx = "Silent"
            else: ctx = "None"
            it = QTableWidgetItem(ctx); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.b2b_table.setItem(ri,4,it)
            
            # Col 5: Format combo
            cb = NoScrollComboBox(); cb.addItems(["Both", "PDF", "Excel", "Amount"]); cb.setCurrentText("Both")
            
            # Find the format column dynamically like in POS
            b2b_type_col = None
            for col in df_info.columns:
                if str(col).strip().lower() in ('type', 'format', 'report_type', 'report type', 'report format', 'file type', 'filetype'):
                    b2b_type_col = col
                    break
            
            if b2b_type_col is not None:
                fmt = normalize_type(row.get(b2b_type_col, ''))
                if fmt == 'pdf': cb.setCurrentText("PDF")
                elif fmt == 'excel': cb.setCurrentText("Excel")
                elif fmt == 'both': cb.setCurrentText("Both")
                elif fmt == 'amount': cb.setCurrentText("Amount")
            
            self.b2b_table.setCellWidget(ri, 5, cb)
            
            # Col 6: Match
            it = QTableWidgetItem("—"); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.b2b_table.setItem(ri,6,it)
            
            # Col 7: Generate button
            bg = QPushButton("⚙️ Generate"); bg.setObjectName("btnAccent")
            bg.clicked.connect(lambda ch, r=ri, n=mn, i=mid: self.generate_b2b_report(r, n, i))
            self.b2b_table.setCellWidget(ri,7,bg)
            
            # Col 8: View button
            bv = QPushButton("👁️ View"); bv.setEnabled(False)
            bv.clicked.connect(lambda ch, m=mid: self.view_b2b_files(m))
            self.b2b_table.setCellWidget(ri,8,bv)
            
            # Col 9: Send button (Automatic detection)
            wa_can_send = (ct['phones'] or ct['wa_link']) and not ct['is_silent']
            if ct['emails']:
                bs = QPushButton("📧 Email"); bs.setObjectName("btnAccent"); bs.setEnabled(False)
                bs.clicked.connect(lambda ch, n=mn, i=mid, c=ct, cc=cc_ct, r=ri: self.send_b2b_email(n,i,c,r,cc_contact=cc))
            elif wa_can_send:
                bs = QPushButton("💬 WA Send"); bs.setObjectName("btnWA"); bs.setEnabled(False)
                bs.clicked.connect(lambda ch, n=mn, i=mid, c=ct, r=ri: self.send_b2b_whatsapp(n,i,c,r))
            else:
                bs = QPushButton("—"); bs.setEnabled(False)
            self.b2b_table.setCellWidget(ri,9,bs)
            
            # Col 10: Report Status
            it = QTableWidgetItem("⏳ Pending"); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.b2b_table.setItem(ri,10,it)
            
            # Col 11: Send Status
            it = QTableWidgetItem("—"); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.b2b_table.setItem(ri,11,it)
            
        self.b2b_table.setSortingEnabled(True)
        self.b2b_lbl_count.setText(f"👥 {mc} merchants")
        logger.info(f"B2B Table: {mc} merchants populated")

    def generate_all_b2b_reports(self):
        if self.b2b_df_transactions is None:
            QMessageBox.warning(self, "Warning", "Please load transactions file first."); return
        count = 0
        for row in range(self.b2b_table.rowCount()):
            mid_item = self.b2b_table.item(row, 2)
            name_item = self.b2b_table.item(row, 1)
            if not mid_item or not name_item: continue
            self.generate_b2b_report(row, name_item.text(), mid_item.text())
            count += 1
        if count == 0: QMessageBox.warning(self, "No Data", "No merchants to process.")

    def generate_b2b_report(self, row_idx, merchant_name, merchant_id):
        if self.b2b_df_transactions is None:
            QMessageBox.warning(self, "Warning", "Please load B2B transactions file first."); return
        cb = self.b2b_table.cellWidget(row_idx, 5)
        fmt = cb.currentText() if cb else "Both"
        to_date = self.b2b_date_to.date().toString("yyyy-MM-dd")
        logger.info(f"B2B Trigger gen: {merchant_name} ({merchant_id}) [{fmt}]")
        self.b2b_table.setSortingEnabled(False)
        it = self.b2b_table.item(row_idx, 10)
        if it: it.setText("🔄 Generating...")
        self.b2b_table.setSortingEnabled(True)
        
        bg = self.b2b_table.cellWidget(row_idx, 7)
        if bg: bg.setEnabled(False)
        
        os.makedirs(self.b2b_output_dir, exist_ok=True)
        t = ReportGeneratorThread(merchant_name, merchant_id, fmt, self.b2b_df_transactions, self.b2b_df_payments, self.b2b_output_dir, to_date, is_b2b=True)
        t.finished.connect(lambda n, f, rt, r=row_idx, m=merchant_id: self.on_b2b_report_done(r, m, f, rt))
        t.match_result.connect(lambda m, p, n, ok: self.on_b2b_match_result(row_idx, m, p, n, ok))
        t.error.connect(lambda e: self.on_b2b_report_err(row_idx, e))
        t.start(); self.active_threads.append(t)

    def on_b2b_match_result(self, row_idx, mid, pay_amt, net_amt, matched):
        self.b2b_table.setSortingEnabled(False)
        it = self.b2b_table.item(row_idx, 6)
        if it:
            if matched:
                it.setText(f"✅ Net={abs(net_amt):,.2f} = Pay={pay_amt:,.2f}")
                it.setForeground(QColor('#10b981'))
            elif pay_amt == 0 and net_amt != 0:
                it.setText(f"⚠️ Not in Payment File (Net={abs(net_amt):,.2f})")
                it.setForeground(QColor('#f59e0b'))
            else:
                it.setText(f"❌ Net={abs(net_amt):,.2f} ≠ Pay={pay_amt:,.2f}")
                it.setForeground(QColor('#ef4444'))
        self.b2b_table.setSortingEnabled(True)

    def on_b2b_report_done(self, row_idx, mid, files, rtype):
        self.b2b_generated_reports[str(mid)] = files
        self.generated_date_ranges = getattr(self, 'generated_date_ranges', {})
        self.generated_date_ranges[str(mid)] = rtype
        self.b2b_table.setSortingEnabled(False)
        bg = self.b2b_table.cellWidget(row_idx, 7)
        if bg: bg.setEnabled(True)
        bv = self.b2b_table.cellWidget(row_idx, 8)
        if bv: bv.setEnabled(True)
        bs = self.b2b_table.cellWidget(row_idx, 9)
        if bs: bs.setEnabled(True)
        it = self.b2b_table.item(row_idx, 10)
        if it: it.setText("✅ Done")
        it_s = self.b2b_table.item(row_idx, 11)
        if it_s: it_s.setText("⏳ Pending")
        self.b2b_table.setSortingEnabled(True)

    def on_b2b_report_err(self, row_idx, msg):
        self.b2b_table.setSortingEnabled(False)
        bg = self.b2b_table.cellWidget(row_idx, 7)
        if bg: bg.setEnabled(True)
        it = self.b2b_table.item(row_idx, 10)
        if it: it.setText("❌ Failed")
        self.b2b_table.setSortingEnabled(True)
        QMessageBox.warning(self, "B2B Report Error", msg)

    def view_b2b_files(self, mid):
        files = self.b2b_generated_reports.get(str(mid), [])
        if not files: QMessageBox.information(self, "No Files", "No reports generated."); return
        for f in files:
            if os.path.exists(f):
                try:
                    if sys.platform=='win32': os.startfile(f)
                    elif sys.platform=='darwin': subprocess.run(['open',f])
                    else: subprocess.run(['xdg-open',f])
                except Exception as e: logger.error(f"Open file failed: {e}")

    def send_b2b_email(self, name, mid, contact, row_idx, cc_contact=None):
        logger.info(f"B2B Email: {name} ({mid})")
        # Get settled amount and date range
        settled, dr_fallback = self._get_email_data(mid, is_b2b=True)
        # Use generated date range if available
        dr = getattr(self, 'generated_date_ranges', {}).get(str(mid), dr_fallback)
        
        to = ';'.join(contact['emails'])
        cc_str = ';'.join(cc_contact['emails']) if cc_contact and cc_contact['emails'] else ""
        
        subj = f"Transaction Settlement Report - {name}"
        files = self.b2b_generated_reports.get(str(mid), [])
        # Format the date range for the email body to match user's request exactly
        dr_body = dr.replace('From ', 'From: ').replace(' To ', '\nTo: ')
        
        body = (f"Dear Merchant,\n\n"
                f"Please find attached your transaction settlement report.\n\n"
                f"Merchant Details\n"
                f"Merchant Name: {name}\n"
                f"Merchant ID: {mid}\n\n"
                f"Settlement Period\n"
                f"{dr_body}\n\n"
                f"Settled Amount\n"
                f"IQD {settled:,.2f}\n\n"
                f"Best regards,")
        # HTML Body for Outlook to prevent extra spacing and improve organization
        html_body = (
            f"<div style='font-family: Calibri, sans-serif; font-size: 11pt; color: #000;'>"
            f"Dear Merchant,<br><br>"
            f"Please find attached your transaction settlement report.<br><br>"
            f"<b>Merchant Details</b><br>"
            f"Merchant Name: {name}<br>"
            f"Merchant ID: {mid}<br><br>"
            f"<b>Settlement Period</b><br>"
            f"{dr_body.replace('\n', '<br>')}<br><br>"
            f"<b>Settled Amount</b><br>"
            f"IQD {settled:,.2f}<br><br>"
            f"Best regards,"
            f"</div>"
        )
        sent = False; attached = 0
        if HAS_OUTLOOK:
            try:
                ol = win32com.client.Dispatch('Outlook.Application')
                mail = ol.CreateItem(0); mail.To = to
                if cc_str:
                    mail.CC = cc_str
                mail.Subject = subj; mail.HTMLBody = html_body
                for fp in files:
                    if os.path.exists(fp):
                        try: mail.Attachments.Add(Source=os.path.abspath(fp)); attached += 1
                        except: pass
                mail.Display(); sent = True
            except: pass
        if not sent and sys.platform == 'win32':
            if self._outlook_compose_via_vbs(to, subj, html_body, files, cc=cc_str, is_html=True):
                attached = len(files); sent = True
        if not sent:
            mu = f"mailto:{urllib.parse.quote(to)}?subject={urllib.parse.quote(subj)}&body={urllib.parse.quote(body)}"
            if cc_str:
                mu += f"&cc={urllib.parse.quote(cc_str)}"
            try: QDesktopServices.openUrl(QUrl(mu)); sent = True
            except: pass
        self.b2b_table.setSortingEnabled(False)
        it = self.b2b_table.item(row_idx, 11)
        if it:
            if sent and attached > 0: it.setText(f"✅ Sent ({attached})")
            elif sent: it.setText("✅ Opened")
            else: it.setText("❌ Failed")
        self.b2b_table.setSortingEnabled(True)

    def send_b2b_whatsapp(self, name, mid, contact, row_idx):
        logger.info(f"B2B WhatsApp: {name} ({mid})")
        # B2B WhatsApp Handling
        wa_link = contact.get('wa_link')
        pc = ""
        # Check if we have at least one way to send
        if not contact['phones'] and not wa_link:
            return # Just exit silently or log it
        
        if contact['phones']:
            phone = contact['phones'][0]
            if not phone.startswith('+'): phone = '+' + phone
            pc = re.sub(r'[^\d]','', phone)
            
        settled, dr_fallback = self._get_email_data(mid, is_b2b=True)
        dr = getattr(self, 'generated_date_ranges', {}).get(str(mid), dr_fallback)
        # Format the date range for the WA body to match user's request exactly
        dr_body = dr.replace('From ', 'From: ').replace(' To ', '\nTo: ')
        
        body = (f"Dear Merchant,\n\n"
                f"Please find attached your transaction settlement report.\n\n"
                f"Merchant Details\n"
                f"Merchant Name: {name}\n"
                f"Merchant ID: {mid}\n\n"
                f"Settlement Period\n"
                f"{dr_body}\n\n"
                f"Settled Amount\n"
                f"IQD {settled:,.2f}\n\n"
                f"Best regards,")
        em = urllib.parse.quote(body)
        files = self.b2b_generated_reports.get(mid, [])
        has_attachment = False
        if files and os.path.exists(files[0]):
            from PyQt6.QtCore import QMimeData, QUrl as QUrl2
            mime = QMimeData()
            mime.setUrls([QUrl2.fromLocalFile(os.path.abspath(f)) for f in files if os.path.exists(f)])
            QApplication.clipboard().setMimeData(mime)
            has_attachment = True
        if wa_link:
            invite_code = wa_link.split('/')[-1]
            wa_desktop = f"whatsapp://chat?code={invite_code}"
            wa_web = wa_link
            has_attachment = True # Force clipboard copy for groups
        else:
            wa_desktop = f"whatsapp://send?phone={pc}" if has_attachment else f"whatsapp://send?phone={pc}&text={em}"
            wa_web = f"https://web.whatsapp.com/send?phone={pc}" if has_attachment else f"https://web.whatsapp.com/send?phone={pc}&text={em}"
            
        sent = False; is_web = False
        try:
            if sys.platform == 'win32': os.startfile(wa_desktop)
            else: subprocess.Popen(['xdg-open', wa_desktop])
            sent = True
        except: pass
        if not sent:
            try: webbrowser.open(wa_web); sent = True; is_web = True
            except: pass
        if sent and (has_attachment or wa_link):
            delay = 12.0 if is_web else 7.0
            self.b2b_auto_wa = AutoWaSendThread(msg=body, delay=delay)
            self.b2b_auto_wa.set_clipboard_text.connect(lambda t: QApplication.clipboard().setText(t))
            self.b2b_auto_wa.start()
        self.b2b_table.setSortingEnabled(False)
        it = self.b2b_table.item(row_idx, 11)
        if it: it.setText("✅ Opened" if sent else "❌ Failed")
        self.b2b_table.setSortingEnabled(True)

    def clear_b2b_files(self):
        r = QMessageBox.question(self,"Clear B2B Files","Clear all loaded B2B files?",
            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            self.b2b_df_info = None; self.b2b_df_payments = None; self.b2b_df_transactions = None; self.b2b_generated_reports = {}
            self.b2b_table.setRowCount(0); self.b2b_lbl_count.setText("👥 0 merchants")
            for i in range(3):
                self.b2b_file_status[i].setText("➖ Not loaded"); self.b2b_file_status[i].setObjectName("labelStatus")
                self.b2b_file_status[i].style().unpolish(self.b2b_file_status[i]); self.b2b_file_status[i].style().polish(self.b2b_file_status[i])
                self.b2b_file_progress[i].setValue(0)
            self.b2b_btn_clear.setEnabled(False)

    def closeEvent(self, event):
        logger.info("Application closed")
        for t in self.active_threads:
            if t.isRunning(): t.quit(); t.wait(2000)
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setFont(QFont("Segoe UI", 10))
    w = MainWindow()
    w.showMaximized()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
