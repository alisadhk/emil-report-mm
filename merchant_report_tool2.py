"""
Merchant Report & Email Tool v1.1
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
from typing import Optional, Dict, List, Tuple

import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QPushButton, QLabel, QProgressBar, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QFileDialog, QDialog,
    QTextEdit, QMessageBox, QFrame, QSplitter, QStatusBar, QToolBar,
    QSizePolicy, QAbstractItemView, QStyle, QScrollArea, QGroupBox,
    QCheckBox, QSpacerItem
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QTimer, QUrl
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
            df = pd.read_excel(self.filepath, engine='openpyxl')
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
    finished = pyqtSignal(str, list)
    error = pyqtSignal(str)
    def __init__(self, merchant_name, merchant_id, report_format, transactions_df, clients_df, output_dir):
        super().__init__()
        self.merchant_name = merchant_name
        self.merchant_id = merchant_id
        self.report_format = report_format
        self.transactions_df = transactions_df
        self.clients_df = clients_df
        self.output_dir = output_dir

    def run(self):
        try:
            generated_files = []
            mid = int(self.merchant_id)
            logger.info(f"Generating {self.report_format} for: {self.merchant_name} (ID: {mid})")
            self.progress.emit(10)
            txn = self.transactions_df[self.transactions_df['Merchant ID'] == mid].copy()
            if txn.empty:
                msg = f"No transactions for Merchant ID: {mid}"
                logger.warning(msg)
                self.error.emit(msg)
                return
            logger.info(f"Found {len(txn):,} transactions")
            self.progress.emit(20)
            txn = txn.sort_values(['Transaction Date', 'Transaction Time'])
            settled = 0
            if self.clients_df is not None:
                cr = self.clients_df[self.clients_df['merchant_id'] == mid]
                if not cr.empty:
                    settled = cr['  payment_amount  '].sum()
                    logger.info(f"Settled: {settled:,.2f} IQD")
                else:
                    logger.warning(f"Merchant {mid} not in clients file")
            d_min = txn['Process Date'].min()
            d_max = txn['Process Date'].max()
            dr = f"{d_min.strftime('%d/%m/%Y')} to {d_max.strftime('%d/%m/%Y')}"
            logger.info(f"Date range: {dr}")
            self.progress.emit(30)
            os.makedirs(self.output_dir, exist_ok=True)
            summary = {
                'merchant_name': self.merchant_name, 'merchant_id': mid,
                'date_range': dr, 'settled_amount': settled,
                'total_transactions': len(txn),
                'total_amount': txn['Transaction Amount'].sum(),
                'total_commission': txn['Merchant Commission'].sum(),
            }
            rcols = ['Transaction Date','Transaction Time','Masked Card Number','result',
                     'Transaction Amount','Merchant Commission','PAYMENT_MODE','Terminal ID',
                     'Auth Code','RRN','N\\I']
            ecols = [c for c in rcols if c in txn.columns]
            rdata = txn[ecols].copy()
            if 'Transaction Date' in rdata.columns:
                rdata['Transaction Date'] = pd.to_datetime(rdata['Transaction Date']).dt.strftime('%d/%m/%Y')
            fmt = self.report_format.lower()
            if fmt in ('pdf','both'):
                self.progress.emit(40)
                p = os.path.join(self.output_dir, f"{self.merchant_name}_{mid}_Report.pdf")
                self._gen_pdf(rdata, summary, p)
                generated_files.append(p)
                logger.info(f"PDF saved: {p}")
            if fmt in ('excel','both'):
                self.progress.emit(70)
                p = os.path.join(self.output_dir, f"{self.merchant_name}_{mid}_Report.xlsx")
                self._gen_xlsx(rdata, summary, p)
                generated_files.append(p)
                logger.info(f"Excel saved: {p}")
            self.progress.emit(100)
            logger.info(f"Done: {self.merchant_name} ({len(generated_files)} files)")
            self.finished.emit(self.merchant_name, generated_files)
        except Exception as e:
            logger.error(f"Report failed for {self.merchant_name}: {e}")
            self.error.emit(f"Failed:\n{e}")

    def _gen_pdf(self, data, summary, filepath):
        logger.info("Generating PDF...")
        doc = SimpleDocTemplate(filepath, pagesize=landscape(A4),
            rightMargin=15*mm, leftMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
        styles = getSampleStyleSheet()
        els = []
        ts = ParagraphStyle('T', parent=styles['Heading1'], fontSize=18,
            textColor=rl_colors.HexColor('#1e40af'), spaceAfter=6, alignment=TA_CENTER, fontName='Helvetica-Bold')
        ss = ParagraphStyle('S', parent=styles['Normal'], fontSize=11,
            textColor=rl_colors.HexColor('#4b5563'), spaceAfter=4, alignment=TA_CENTER)
        els.append(Paragraph("Transaction Settlement Report", ts))
        els.append(Paragraph(f"Generated on {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ss))
        els.append(Spacer(1, 8*mm))
        idata = [
            ['Merchant Name:', summary['merchant_name'], 'Merchant ID:', str(summary['merchant_id'])],
            ['Date Range:', summary['date_range'], 'Total Transactions:', f"{summary['total_transactions']:,}"],
            ['Settled Amount:', f"{summary['settled_amount']:,.2f} IQD", 'Total Commission:', f"{summary['total_commission']:,.2f} IQD"],
        ]
        it = Table(idata, colWidths=[100,200,120,200])
        it.setStyle(TableStyle([
            ('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),('FONTNAME',(2,0),(2,-1),'Helvetica-Bold'),
            ('FONTSIZE',(0,0),(-1,-1),9),
            ('TEXTCOLOR',(0,0),(0,-1),rl_colors.HexColor('#4b5563')),('TEXTCOLOR',(2,0),(2,-1),rl_colors.HexColor('#4b5563')),
            ('TEXTCOLOR',(1,0),(1,-1),rl_colors.HexColor('#111827')),('TEXTCOLOR',(3,0),(3,-1),rl_colors.HexColor('#111827')),
            ('BACKGROUND',(0,0),(-1,-1),rl_colors.HexColor('#f8fafc')),
            ('BOX',(0,0),(-1,-1),0.5,rl_colors.HexColor('#e2e8f0')),
            ('INNERGRID',(0,0),(-1,-1),0.25,rl_colors.HexColor('#e2e8f0')),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),6),
            ('BOTTOMPADDING',(0,0),(-1,-1),6),('LEFTPADDING',(0,0),(-1,-1),8),
        ]))
        els.append(it)
        els.append(Spacer(1, 8*mm))
        cm = {'Transaction Date':'Date','Transaction Time':'Time','Masked Card Number':'Card Number',
              'result':'Card Type','Transaction Amount':'Amount','Merchant Commission':'Commission',
              'PAYMENT_MODE':'Pay Mode','Terminal ID':'Terminal','Auth Code':'Auth Code','RRN':'RRN','N\\I':'N/I'}
        hdrs = [cm.get(c,c) for c in data.columns]
        td = [hdrs]
        for _, row in data.iterrows():
            rd = []
            for col in data.columns:
                v = row[col]
                if pd.isna(v): rd.append('')
                elif isinstance(v, float):
                    if col in ('Transaction Amount','Merchant Commission'): rd.append(f"{v:,.2f}")
                    else: rd.append(str(int(v)) if v==int(v) else f"{v:.2f}")
                else: rd.append(str(v))
            td.append(rd)
        nc = len(hdrs)
        aw = landscape(A4)[0] - 30*mm
        cw = [aw/nc]*nc
        cs = 40
        for s in range(0, len(td)-1, cs):
            ch = [td[0]] + td[s+1:s+1+cs]
            tbl = Table(ch, colWidths=cw, repeatRows=1)
            tbl.setStyle(TableStyle([
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),8),('FONTSIZE',(0,1),(-1,-1),7),
                ('BACKGROUND',(0,0),(-1,0),rl_colors.HexColor('#1e40af')),('TEXTCOLOR',(0,0),(-1,0),rl_colors.white),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
                ('ROWBACKGROUNDS',(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor('#f1f5f9')]),
                ('BOX',(0,0),(-1,-1),0.5,rl_colors.HexColor('#cbd5e1')),
                ('INNERGRID',(0,0),(-1,-1),0.25,rl_colors.HexColor('#e2e8f0')),
            ]))
            els.append(tbl)
            if s+cs < len(td)-1: els.append(PageBreak())
        els.append(Spacer(1, 6*mm))
        sd = [
            ['Total Transactions',f"{summary['total_transactions']:,}",'Total Amount',f"{summary['total_amount']:,.2f} IQD"],
            ['Total Commission',f"{summary['total_commission']:,.2f} IQD",'Settled Amount',f"{summary['settled_amount']:,.2f} IQD"],
        ]
        st = Table(sd, colWidths=[120,180,120,180])
        st.setStyle(TableStyle([
            ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),
            ('BACKGROUND',(0,0),(-1,-1),rl_colors.HexColor('#eef2ff')),
            ('TEXTCOLOR',(0,0),(0,-1),rl_colors.HexColor('#3730a3')),('TEXTCOLOR',(2,0),(2,-1),rl_colors.HexColor('#3730a3')),
            ('BOX',(0,0),(-1,-1),1,rl_colors.HexColor('#818cf8')),
            ('INNERGRID',(0,0),(-1,-1),0.5,rl_colors.HexColor('#c7d2fe')),
            ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),('LEFTPADDING',(0,0),(-1,-1),10),
        ]))
        els.append(st)
        doc.build(els)

    def _gen_xlsx(self, data, summary, filepath):
        logger.info("Generating Excel...")
        wb = Workbook(); ws = wb.active; ws.title = "Transaction Report"
        hf = XlFont(name='Segoe UI',bold=True,size=9,color='FFFFFF')
        hfl = XlFill(start_color='1E40AF',end_color='1E40AF',fill_type='solid')
        ha = XlAlign(horizontal='center',vertical='center',wrap_text=True)
        df_ = XlFont(name='Segoe UI',size=9)
        da = XlAlign(horizontal='center',vertical='center')
        af = '#,##0.00'
        tb = XlBorder(left=XlSide(style='thin',color='D1D5DB'),right=XlSide(style='thin',color='D1D5DB'),
                      top=XlSide(style='thin',color='D1D5DB'),bottom=XlSide(style='thin',color='D1D5DB'))
        ws.merge_cells('A1:K1')
        ws['A1'] = 'Transaction Settlement Report'
        ws['A1'].font = XlFont(name='Segoe UI',bold=True,size=14,color='1E40AF')
        ws['A1'].alignment = XlAlign(horizontal='center',vertical='center')
        ws.row_dimensions[1].height = 35
        ws.merge_cells('A2:K2')
        ws['A2'] = f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
        ws['A2'].font = XlFont(name='Segoe UI',size=9,color='6B7280')
        ws['A2'].alignment = XlAlign(horizontal='center')
        ikf = XlFont(name='Segoe UI',bold=True,size=10,color='4B5563')
        ivf = XlFont(name='Segoe UI',size=10,color='111827')
        ifl = XlFill(start_color='F8FAFC',end_color='F8FAFC',fill_type='solid')
        irows = [
            ('Merchant Name:',summary['merchant_name'],'Merchant ID:',str(summary['merchant_id'])),
            ('Date Range:',summary['date_range'],'Total Transactions:',f"{summary['total_transactions']:,}"),
            ('Settled Amount:',f"{summary['settled_amount']:,.2f} IQD",'Total Commission:',f"{summary['total_commission']:,.2f} IQD"),
        ]
        for i,rd in enumerate(irows, start=4):
            ws.cell(row=i,column=1,value=rd[0]).font=ikf; ws.cell(row=i,column=2,value=rd[1]).font=ivf
            ws.cell(row=i,column=4,value=rd[2]).font=ikf; ws.cell(row=i,column=5,value=rd[3]).font=ivf
            for c in range(1,7): ws.cell(row=i,column=c).fill=ifl; ws.cell(row=i,column=c).border=tb
        cm = {'Transaction Date':'Date','Transaction Time':'Time','Masked Card Number':'Card Number',
              'result':'Card Type','Transaction Amount':'Amount (IQD)','Merchant Commission':'Commission (IQD)',
              'PAYMENT_MODE':'Payment Mode','Terminal ID':'Terminal ID','Auth Code':'Auth Code','RRN':'RRN','N\\I':'N/I'}
        sr = 8
        hdrs = [cm.get(c,c) for c in data.columns]
        for ci,h in enumerate(hdrs,start=1):
            cell=ws.cell(row=sr,column=ci,value=h); cell.font=hf; cell.fill=hfl; cell.alignment=ha; cell.border=tb
        ws.row_dimensions[sr].height = 28
        afl = XlFill(start_color='F1F5F9',end_color='F1F5F9',fill_type='solid')
        acols = {'Transaction Amount','Merchant Commission'}
        for ri,(_, row) in enumerate(data.iterrows(), start=sr+1):
            for ci,cn in enumerate(data.columns, start=1):
                v = row[cn]
                if pd.isna(v): v=''
                cell=ws.cell(row=ri,column=ci,value=v); cell.font=df_; cell.alignment=da; cell.border=tb
                if cn in acols and isinstance(v,(int,float)): cell.number_format=af
                if ri%2==0: cell.fill=afl
        for ci in range(1,len(hdrs)+1):
            ml = len(str(hdrs[ci-1]))
            for row in ws.iter_rows(min_row=sr,max_row=min(sr+50,ws.max_row),min_col=ci,max_col=ci):
                for cell in row:
                    if cell.value: ml=max(ml,len(str(cell.value)))
            ws.column_dimensions[get_column_letter(ci)].width = min(ml+4,30)
        ws.freeze_panes = f'A{sr+1}'
        lc = get_column_letter(len(hdrs))
        ws.auto_filter.ref = f'A{sr}:{lc}{ws.max_row}'
        wb.save(filepath)


# ── Utilities ──
def parse_contact_info(email_str):
    result = {'emails':[],'phones':[],'is_silent':False,'is_wa_group':False,'raw':str(email_str).strip()}
    if pd.isna(email_str): result['is_silent']=True; return result
    text = str(email_str).strip()
    if not text: result['is_silent']=True; return result
    if 'group on wa' in text.lower(): result['is_wa_group']=True; return result
    if text.lower().startswith('silent') and '@' not in text:
        result['is_silent']=True
        phones = re.findall(r'\+?\d[\d\s\-]{8,}', text)
        if phones: result['phones']=[re.sub(r'[\s\-]','',p) for p in phones]; result['is_silent']=False
        return result
    result['emails'] = list(set(re.findall(r'[\w.+-]+@[\w.-]+\.\w+', text)))
    phones = re.findall(r'\+?\d[\d\s\-]{8,}', text)
    result['phones'] = [re.sub(r'[\s\-]','',p) for p in phones]
    if not result['emails'] and not result['phones']: result['is_silent']=True
    return result

def normalize_type(type_str):
    if pd.isna(type_str): return 'pdf'
    t = str(type_str).lower().strip()
    if 'amount' in t: return 'amount'
    if 'excel' in t and 'pdf' in t: return 'both'
    if 'excel' in t: return 'excel'
    return 'pdf'

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
        self.setWindowTitle("About"); self.setMinimumSize(580,480)
        layout = QVBoxLayout(self); layout.setContentsMargins(20,20,20,20); layout.setSpacing(10)
        t = QLabel("Merchant Report & Email Tool"); t.setObjectName("labelTitle"); t.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(t)
        v = QLabel("Version 1.1  |  *****"); v.setObjectName("labelSubtitle"); v.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(v)
        s = QFrame(); s.setObjectName("separator"); s.setFrameShape(QFrame.Shape.HLine); layout.addWidget(s)
        tx = QTextEdit(); tx.setReadOnly(True)
        tx.setHtml("""<style>body{font-family:'Segoe UI';font-size:13px;line-height:1.6;}h3{color:#3b82f6;margin-top:12px;}b{color:#60a5fa;}.s{margin-left:10px;margin-bottom:3px;}</style>
        <h3>How to Use</h3>
        <div class="s"><b>1.</b> Load three Excel files using the top panel</div>
        <div class="s"><b>2.</b> Merchant table auto-populates from Send Info file</div>
        <div class="s"><b>3.</b> Select report format (PDF / Excel / Both)</div>
        <div class="s"><b>4.</b> Click <b>Generate</b> — saved to "Customer Email Report" folder</div>
        <div class="s"><b>5.</b> Click <b>Send</b> — opens pre-filled email or WhatsApp</div>
        <div class="s"><b>6.</b> Review & send manually</div>
        <h3>Status Columns</h3>
        <div class="s"><b>Report Status:</b> Tracks generation (Pending / Generating / Done / Failed)</div>
        <div class="s"><b>Send Status:</b> Tracks send action (Pending / Opened / Failed)</div>
        <h3>Contact Methods</h3>
        <div class="s"><b>Email:</b> Opens Outlook or default mail client with attachments</div>
        <div class="s"><b>WA Send:</b> Tries WhatsApp Desktop first, falls back to WhatsApp Web (text only, no attachment)</div>
        <h3>Extra</h3>
        <div class="s"><b>Log:</b> Detailed logs for troubleshooting</div>
        <div class="s"><b>Theme:</b> Dark/Light mode toggle</div>
        <div class="s"><b>Generate All:</b> Batch-generate all reports</div>""")
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
        logger.info("Application started")
        logger.info(f"Output dir: {self.output_dir}")
        self.setup_ui()
        self.apply_theme(DARK_THEME)

    def setup_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        ml = QVBoxLayout(central); ml.setContentsMargins(12,6,12,6); ml.setSpacing(6)

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
        fcfg = [("📋 Send Info File","Contacts & types"),("👥 Clients File","IBAN & amounts"),("💳 Transactions File","Financial records")]
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

        ab.addStretch()
        
        self.btn_clear = QPushButton("🗑️ Clear"); self.btn_clear.setObjectName("btnPurple"); self.btn_clear.setEnabled(False)
        self.btn_clear.clicked.connect(self.clear_all_files); ab.addWidget(self.btn_clear)
        self.btn_folder = QPushButton("📁 Open Reports Folder"); self.btn_folder.setObjectName("btnPurple")
        self.btn_folder.clicked.connect(self.open_reports_folder); ab.addWidget(self.btn_folder)
        ml.addLayout(ab)

        # Table — 9 columns
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSortingEnabled(True)
        cols = ['#','Merchant Name','Merchant ID','Contact','Format','Report','Send','Report Status','Send Status']
        self.table.setColumnCount(len(cols)); self.table.setHorizontalHeaderLabels(cols)
        h = self.table.horizontalHeader()
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        h.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for c in range(2,9): h.setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0,35); self.table.setColumnWidth(2,110); self.table.setColumnWidth(3,100)
        self.table.setColumnWidth(4,85); self.table.setColumnWidth(5,110); self.table.setColumnWidth(6,110)
        self.table.setColumnWidth(7,115); self.table.setColumnWidth(8,115)
        ml.addWidget(self.table, stretch=1)

        self.status_bar = QStatusBar(); self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready  —  Load the three data files to begin")
        self.lbl_brand = QLabel(" ***** "); self.lbl_brand.setObjectName("labelSubtitle")
        self.status_bar.addPermanentWidget(self.lbl_brand)

    def apply_theme(self, theme):
        self.current_theme = theme; self.setStyleSheet(build_stylesheet(theme))
        self.btn_theme.setText("☀️ Light" if theme['name']=='dark' else "🌙 Dark")
        logger.info(f"Theme: {theme['name']}")

    def toggle_theme(self):
        self.apply_theme(LIGHT_THEME if self.current_theme['name']=='dark' else DARK_THEME)

    def show_log(self):
        d = LogDialog(self); d.setStyleSheet(self.styleSheet()); d.exec()

    def show_about(self):
        d = AboutDialog(self); d.setStyleSheet(self.styleSheet()); d.exec()

    def load_file(self, index):
        fp,_ = QFileDialog.getOpenFileName(self,"Select File","","Excel Files (*.xlsx *.xls);;All Files (*)")
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
        self.file_buttons[index].setEnabled(True)
        r, c = df.shape
        if index==0: self.df_send_info=df; self.file_status[index].setText(f"✅ {r} merchants"); self.populate_table()
        elif index==1: self.df_clients=df; self.file_status[index].setText(f"✅ {r} records")
        elif index==2: self.df_transactions=df; self.file_status[index].setText(f"✅ {r:,} txns")
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
        df['_tn'] = df['Type '].apply(normalize_type)
        df = df[df['_tn']!='amount'].reset_index(drop=True)
        self.table.setRowCount(len(df))
        mc = 0
        for ri, (_, row) in enumerate(df.iterrows()):
            mc += 1
            mid = str(int(row['Merchant ID'])) if pd.notna(row['Merchant ID']) else ''
            mn = str(row['Merchant Name']).strip()
            mt = row['_tn']
            ct = parse_contact_info(row['Email'])
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
            # Contact
            if ct['emails']: ctx = f"Email ({len(ct['emails'])})"
            elif ct['phones']: ctx = "WhatsApp"
            elif ct['is_wa_group']: ctx = "WA Group"
            elif ct['is_silent']: ctx = "Silent"
            else: ctx = "None"
            it = QTableWidgetItem(ctx); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.table.setItem(ri,3,it)
            # Format
            cb = QComboBox(); cb.addItems(["PDF","Excel","Both"])
            if mt=='excel': cb.setCurrentIndex(1)
            elif mt=='both': cb.setCurrentIndex(2)
            self.table.setCellWidget(ri,4,cb)
            # Generate
            bg = QPushButton("⚙️ Generate"); bg.setObjectName("btnGreen"); bg.setEnabled(False)
            bg.clicked.connect(lambda ch, r=ri, n=mn, i=mid: self.generate_report(r,n,i))
            self.table.setCellWidget(ri,5,bg)
            # Send
            if ct['emails']:
                bs = QPushButton("📧 Email"); bs.setObjectName("btnAccent")
                bs.clicked.connect(lambda ch, n=mn, i=mid, c=ct, r=ri: self.send_email(n,i,c,r))
            elif ct['phones'] and not ct['is_silent']:
                bs = QPushButton("💬 WA Send"); bs.setObjectName("btnWA")
                bs.clicked.connect(lambda ch, n=mn, i=mid, c=ct, r=ri: self.send_whatsapp(n,i,c,r))
            else:
                bs = QPushButton("—"); bs.setEnabled(False)
            self.table.setCellWidget(ri,6,bs)
            # Report Status
            it = QTableWidgetItem("⏳ Pending"); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.table.setItem(ri,7,it)
            # Send Status
            it = QTableWidgetItem("⏳ Pending"); it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            it.setFlags(it.flags() & ~Qt.ItemFlag.ItemIsEditable); self.table.setItem(ri,8,it)
        self.table.setSortingEnabled(True)
        self.lbl_count.setText(f"👥 {mc} merchants")
        self.update_gen_btns()
        logger.info(f"Table: {mc} merchants (excl. amount)")

    def update_gen_btns(self):
        al = all([self.df_send_info is not None, self.df_clients is not None, self.df_transactions is not None])
        for r in range(self.table.rowCount()):
            b = self.table.cellWidget(r,5)
            if b: b.setEnabled(al)

    def generate_report(self, row_idx, merchant_name, merchant_id):
        if self.df_transactions is None:
            QMessageBox.warning(self,"Missing Data","Load Transactions file first."); return
        cb = self.table.cellWidget(row_idx,4)
        rf = cb.currentText().lower()
        it = self.table.item(row_idx,7)
        if it: it.setText("⚙️ Generating...")
        logger.info(f"Report: {merchant_name} ({merchant_id}) - {rf}")
        t = ReportGeneratorThread(merchant_name, merchant_id, rf, self.df_transactions, self.df_clients, self.output_dir)
        t.progress.connect(lambda v, r=row_idx: self.table.item(r,7).setText(f"⚙️ Gen {v}%") if self.table.item(r,7) else None)
        t.finished.connect(lambda n, f, r=row_idx: self.on_report_done(n,f,r))
        t.error.connect(lambda m, r=row_idx: self.on_report_err(m,r))
        t.start(); self.active_threads.append(t)

    def _increment_gen_pb(self):
        if hasattr(self, 'gen_all_total') and self.pb_gen_all.isVisible():
            self.gen_all_completed += 1; self.pb_gen_all.setValue(self.gen_all_completed)
            if self.gen_all_completed >= self.gen_all_total:
                QTimer.singleShot(2000, self.pb_gen_all.hide)

    def on_report_done(self, name, files, row_idx):
        mi = self.table.item(row_idx,2)
        if mi: self.generated_reports[mi.text()] = files
        it = self.table.item(row_idx,7)
        if it: it.setText("✅ Done")
        self.update_status()
        self.status_bar.showMessage(f"Report: {name} ({len(files)} file(s))", 5000)
        self._increment_gen_pb()

    def on_report_err(self, msg, row_idx):
        it = self.table.item(row_idx,7)
        if it: it.setText("❌ Failed")
        QMessageBox.warning(self,"Report Error", msg)
        self._increment_gen_pb()

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

    def _get_email_data(self, merchant_id):
        mid = int(merchant_id)
        settled = 0
        if self.df_clients is not None:
            cr = self.df_clients[self.df_clients['merchant_id']==mid]
            if not cr.empty: settled = cr['  payment_amount  '].sum()
        dr = "N/A"
        if self.df_transactions is not None:
            txn = self.df_transactions[self.df_transactions['Merchant ID']==mid]
            if not txn.empty:
                dr = f"{txn['Process Date'].min().strftime('%d/%m/%Y')} to {txn['Process Date'].max().strftime('%d/%m/%Y')}"
        return settled, dr

    def send_email(self, name, mid, contact, row_idx):
        logger.info(f"Email: {name} ({mid})")
        settled, dr = self._get_email_data(mid)
        to = ';'.join(contact['emails'])
        subj = f"Transaction Settlement Report - {name}"
        body = f"Dear,\n\nPlease find attached your transaction settlement report.\n\nMerchant ID: {mid}\nDate Range: {dr}\nSettled Amount: {settled:,.2f} IQD\n\nBest regards"
        rfiles = self.generated_reports.get(mid, [])
        sent = False
        if HAS_OUTLOOK:
            try:
                ol = win32com.client.Dispatch('Outlook.Application')
                mail = ol.CreateItem(0); mail.To=to; mail.Subject=subj; mail.Body=body
                for fp in rfiles:
                    if os.path.exists(fp): mail.Attachments.Add(os.path.abspath(fp))
                mail.Display()
                logger.info(f"Outlook opened: {name} ({len(rfiles)} attachments)")
                sent = True
            except Exception as e: logger.warning(f"Outlook failed: {e}")
        if not sent:
            mu = f"mailto:{urllib.parse.quote(to)}?subject={urllib.parse.quote(subj)}&body={urllib.parse.quote(body)}"
            try:
                QDesktopServices.openUrl(QUrl(mu))
                logger.info(f"mailto: opened: {name}"); sent = True
                if rfiles:
                    logger.warning("mailto: no attachment support")
                    QMessageBox.information(self,"Attach Files",f"Email opened.\nAttach report(s) from:\n{self.output_dir}")
            except Exception as e:
                logger.error(f"Email failed: {e}"); QMessageBox.critical(self,"Email Error",str(e))
        it = self.table.item(row_idx,8)
        if it: it.setText("✅ Opened" if sent else "❌ Failed")

    def send_whatsapp(self, name, mid, contact, row_idx):
        logger.info(f"WhatsApp: {name} ({mid})")
        if not contact['phones']:
            QMessageBox.warning(self,"No Phone","No phone number found."); return
        settled, dr = self._get_email_data(mid)
        phone = contact['phones'][0]
        if not phone.startswith('+'): phone = '+' + phone
        pc = re.sub(r'[^\d]','', phone)
        msg = f"Dear,\n\nPlease find your transaction settlement report.\n\nMerchant ID: {mid}\nDate Range: {dr}\nSettled Amount: {settled:,.2f} IQD\n\nBest regards"
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

        if has_attachment:
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
            logger.info(f"WA Desktop opened: {name} ({phone})")
            sent = True
        except Exception as e:
            logger.warning(f"WA Desktop failed: {e}")
            
        # Fallback to Web
        if not sent:
            try:
                webbrowser.open(wa_web)
                logger.info(f"WA Web opened: {name} ({phone})")
                sent = True
                is_web = True
            except Exception as e:
                logger.error(f"WA Web failed: {e}")
                QMessageBox.critical(self,"WhatsApp Error",str(e))
                
        if sent and has_attachment:
            # Launch automation sequence
            delay = 12.0 if is_web else 4.0
            self.auto_wa_thread = AutoWaSendThread(msg=msg, delay=delay)
            self.auto_wa_thread.set_clipboard_text.connect(lambda t: QApplication.clipboard().setText(t))
            self.auto_wa_thread.start()
        
        it = self.table.item(row_idx,8)
        if it: it.setText("✅ Opened" if sent else "❌ Failed")

    def open_reports_folder(self):
        os.makedirs(self.output_dir, exist_ok=True)
        try:
            if sys.platform=='win32': os.startfile(self.output_dir)
            elif sys.platform=='darwin': subprocess.run(['open',self.output_dir])
            else: subprocess.run(['xdg-open',self.output_dir])
            logger.info(f"Opened: {self.output_dir}")
        except Exception as e: logger.error(f"Open folder failed: {e}")

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
