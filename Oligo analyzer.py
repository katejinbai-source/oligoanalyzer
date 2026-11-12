from __future__ import print_function
from base64 import b64encode
from urllib import request, parse
from urllib.error import HTTPError, URLError
import requests
import json
import re
import traceback
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5 import QtGui, QtCore, QtWidgets
from screeninfo import get_monitors
import sys
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import pandas as pd
import logging
import io
import time
import os

# client_id = 'zhijieyang'
# client_secret = 'f2700e1f-20c4-4336-9647-bbbc9c953c52'
# idt_username = 'zhijie.yang@atilabiosystems.com'
# idt_password = 'Atila123'
#
# idturl = 'https://www.idtdna.com/Restapi/v1/OligoAnalyzer/Analyze'
# # idturl = 'https://www.idtdna.com/restapi/v1/OligoAnalyzer/Hairpin'
#
# authorization_string = b64encode(bytes(client_id + ":" + client_secret, "utf-8")).decode()
#
# request_headers = {"Content-Type": "application/x-www-form-urlencoded",
#                    "Authorization": "Basic " + authorization_string}
#
# data_dict = {"grant_type": "password",
#              "scope": "test",
#              "username": idt_username,
#              "password": idt_password}
#
# request_data = parse.urlencode(data_dict).encode()
#
# post_request = request.Request("https://www.idtdna.com/Identityserver/connect/token",
#                                data=request_data,
#                                headers=request_headers,
#                                method="POST")
#
# # Transmit the HTTP request and get HTTP response
# response = request.urlopen(post_request)
#
# # Process the HTTP response for the desired data
# body = response.read().decode()
#
# # Error and return the response from the endpoint if there was a problem
# if (response.status != 200):
#     raise RuntimeError("Request failed with error code:" + response.status + "\nBody:\n" + body)
#
# body_dict = json.loads(body)
# # print(body_dict)
# token = body_dict["access_token"]

# sequence = "cccccgatcgatcgtagctcgaggggg"
# na_conc = 50
# mg_conc = 2
# dntps_conc = 0.8
# oligo_conc = 0.1
# nucleotide_type = "DNA"
# folding_temp = 30
# max_retry = 5
# delay = 1
#
# sequence = sequence.upper()
#
# payload = json.dumps({
#     "Sequence": sequence,
#     "NaConc": na_conc,
#     "MgConc": mg_conc,
#     "dNTPsConc": dntps_conc,
#     "OligoConc": oligo_conc,
#     "NucleotideType": nucleotide_type,
#     "FoldingTemp": folding_temp
# })
#
# headers = {
#     'Content-Type': 'application/json',
#     'Authorization': f'Bearer {token}',
# }
#
# response = requests.request("POST",
#                             idturl,
#                             headers = headers,
#                             data = payload)
# # Raise an exception for non-2xx responses
# response.raise_for_status()
#
# if response is not None:
#     oligo_info = json.loads(response.text)[0]
#
# for key in oligo_info.keys():
#     print(key, oligo_info[key])

IDT_TOKEN_URL = "https://www.idtdna.com/Identityserver/connect/token"
IDT_ANALYZE_URL = "https://www.idtdna.com/Restapi/v1/OligoAnalyzer/Analyze"
REQUEST_DELAY = 0.25  # pause between sequences so IDT does not throttle the requests

# A plain sequence: IUPAC bases (any case), * (phosphorothioate), + (LNA) and /modification/ codes
_SEQ_PATTERN = re.compile(r"(?:/[^/]+/|[ACGTURYKMSWBDHVN*+])+", re.IGNORECASE)


def clean_sequence(seq):
    """Remove whitespace and uppercase the bases; /modification/ codes are kept as typed."""
    seq = re.sub(r"\s+", "", str(seq))
    parts = re.split(r"(/[^/]*/)", seq)
    return "".join(p if len(p) > 1 and p.startswith("/") and p.endswith("/") else p.upper()
                   for p in parts)


def looks_like_sequence(text):
    return _SEQ_PATTERN.fullmatch(re.sub(r"\s+", "", text)) is not None


def parse_sequence_text(text):
    """Turn pasted or loaded text into a list of (name, sequence) pairs.

    Each line can be:
      Name,Sequence  or  Name<Tab>Sequence (cells copied from Excel); extra columns are ignored
      >Name with the sequence on the following line(s) (FASTA)
      a name on one line and its sequence on the next line
      a sequence on its own (name is None; the caller assigns one)
    Returns (entries, problems); problems describes lines that were skipped.
    """
    entries, problems = [], []
    name, parts, fasta = None, [], False

    def close():
        nonlocal name, parts, fasta
        if parts:
            entries.append((name, "".join(parts)))
        elif name is not None:
            problems.append("No sequence found for '%s', skipped." % name)
        name, parts, fasta = None, [], False

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue

        if line.startswith(">"):
            close()
            name, fasta = (line[1:].strip() or None), True
            continue

        delim = "\t" if "\t" in line else ("," if "," in line else None)
        if delim:
            close()
            fields = [f.strip() for f in line.split(delim)]
            nm = fields[0]
            sq = fields[1] if len(fields) > 1 else ""
            if nm.lower() == "name" and sq.lower().startswith("seq"):
                continue  # header row
            if not sq:
                problems.append("No sequence after the name in line '%s', skipped." % line)
                continue
            entries.append((nm or None, sq))
            continue

        if looks_like_sequence(line):
            if fasta:
                parts.append(line)
            elif name is not None:
                parts.append(line)
                close()
            else:
                entries.append((None, line))
        else:
            close()
            name = line  # name for the sequence on the next line

    close()
    return entries, problems


class WrapTextDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        text = index.data(Qt.DisplayRole)
        if not text:
            return super().paint(painter, option, index)

        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setPlainText(text)
        doc.setTextWidth(option.rect.width())
        painter.save()
        painter.translate(option.rect.topLeft())
        ctx = QAbstractTextDocumentLayout.PaintContext()
        doc.documentLayout().draw(painter, ctx)
        painter.restore()

    def sizeHint(self, option, index):
        text = index.data(Qt.DisplayRole)
        if not text:
            return super().sizeHint(option, index)

        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setPlainText(text)
        doc.setTextWidth(option.rect.width())
        return QSize(option.rect.width(), int(doc.size().height()))


class AutoResizingTable(QTableWidget):
    def __init__(self):
        super().__init__()
        self.setWordWrap(True)
        self.setItemDelegate(WrapTextDelegate(self))

        self.itemChanged.connect(self.auto_resize_rows)
        self.horizontalHeader().sectionResized.connect(self.auto_resize_rows)

    def auto_resize_rows(self, *args):
        self.resizeRowsToContents()

class MainWin(QMainWindow):

    def __init__(self):
        super().__init__()
        self.token = None
        self.initUI()

    def screensize(self):
        return [get_monitors()[0].width, get_monitors()[0].height]

    def initUI(self):

        self.main_frame = QWidget()
        self.statusBar().showMessage('Ready')
        self.statusBar()

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        # LoadAct = QAction("&Load", self)
        # fileMenu.addAction(LoadAct)
        # LoadAct.setStatusTip('Load the QIAcuity exported zip files or RFU files')
        # LoadAct.triggered.connect(self.LoadFile)
        exitAct = QAction('&Exit', self)
        exitAct.setShortcut('Ctrl+Q')
        exitAct.setStatusTip('Exit application')
        exitAct.triggered.connect(self.close)
        fileMenu.addAction(exitAct)

        VMenu = menubar.addMenu('&Help')
        VAct = QAction("&About", self)
        VMenu.addAction(VAct)
        VAct.triggered.connect(self.AboutEvent)

        screen_w = self.screensize()[0]
        screen_h = self.screensize()[1]

        self.max_sw = screen_w / 19.2

        font = QtGui.QFont("Arial", 10)
        font.setBold(True)
        font.setPointSize(10)

        tabbox = QVBoxLayout()
        tabbox.setAlignment(QtCore.Qt.AlignTop)

        self.na_l = QLabel("Na Conc", self)
        self.na_l.setMinimumWidth(int(self.max_sw))
        self.na_l.setMaximumWidth(int(self.max_sw))

        self.mg_l = QLabel("Mg Conc", self)
        self.mg_l.setMinimumWidth(int(self.max_sw))
        self.mg_l.setMaximumWidth(int(self.max_sw))

        self.oligo_l = QLabel("Oligo Conc", self)
        self.oligo_l.setMinimumWidth(int(self.max_sw))
        self.oligo_l.setMaximumWidth(int(self.max_sw))

        self.dntp_l = QLabel("dNTPs Conc", self)
        self.dntp_l.setMinimumWidth(int(self.max_sw))
        self.dntp_l.setMaximumWidth(int(self.max_sw))

        self.na = QLineEdit("50", self)
        self.na.setMinimumWidth(int(self.max_sw))
        self.na.setMaximumWidth(int(self.max_sw))

        self.mg = QLineEdit("2", self)
        self.mg.setMinimumWidth(int(self.max_sw))
        self.mg.setMaximumWidth(int(self.max_sw))

        self.oligo = QLineEdit("0.2", self)
        self.oligo.setMinimumWidth(int(self.max_sw))
        self.oligo.setMaximumWidth(int(self.max_sw))

        self.dntp = QLineEdit("0.8", self)
        self.dntp.setMinimumWidth(int(self.max_sw))
        self.dntp.setMaximumWidth(int(self.max_sw))

        self.uM = QLabel("uM")
        self.mM1 = QLabel("mM")
        self.mM2 = QLabel("mM")
        self.mM3 = QLabel("mM")

        hbox1 = QHBoxLayout()
        hbox1.setAlignment(QtCore.Qt.AlignLeft)
        hbox2 = QHBoxLayout()
        hbox2.setAlignment(QtCore.Qt.AlignLeft)
        hbox3 = QHBoxLayout()
        hbox3.setAlignment(QtCore.Qt.AlignLeft)
        hbox4 = QHBoxLayout()
        hbox4.setAlignment(QtCore.Qt.AlignLeft)

        hbox1.addWidget(self.oligo_l)
        hbox1.addWidget(self.oligo)
        hbox1.addWidget(self.uM)

        hbox2.addWidget(self.na_l)
        hbox2.addWidget(self.na)
        hbox2.addWidget(self.mM1)

        hbox3.addWidget(self.mg_l)
        hbox3.addWidget(self.mg)
        hbox3.addWidget(self.mM2)

        hbox4.addWidget(self.dntp_l)
        hbox4.addWidget(self.dntp)
        hbox4.addWidget(self.mM3)

        hbox5 = QHBoxLayout()
        load_seq = QPushButton("Load sequences from file", self)
        load_seq.setFont(font)
        load_seq.clicked.connect(self.load)

        add_seq = QPushButton("Add sequences from input", self)
        add_seq.setFont(font)
        add_seq.clicked.connect(self.add)

        analyze_seq = QPushButton("Analyze sequences", self)
        analyze_seq.setFont(font)
        analyze_seq.clicked.connect(self.analyze)

        clear_table = QPushButton("Clear table", self)
        clear_table.setFont(font)
        clear_table.clicked.connect(self.clear_table)

        hbox5.setAlignment(QtCore.Qt.AlignLeft)
        hbox5.addWidget(load_seq)
        hbox5.addWidget(add_seq)
        hbox5.addWidget(analyze_seq)
        hbox5.addWidget(clear_table)

        line1 = QFrame()
        line1.setFrameShape(QFrame.HLine)
        line1.setFrameShadow(QFrame.Sunken)  # Optional: gives a nice beveled look

        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)  # Optional: gives a nice beveled look

        input_l = QLabel("Input sequences (one per line: Name,Sequence; a sequence alone also works)")
        input_l.setFont(font)
        self.input = QTextEdit(self)
        self.input.setMaximumHeight(int(screen_h * 0.1))

        error_l = QLabel("Error log")
        error_l.setFont(font)
        self.error = QTextEdit(self)
        self.error.setMaximumHeight(int(screen_h * 0.1))

        hbox6 = QHBoxLayout()
        hbox6.addWidget(input_l, 7)
        hbox6.addWidget(error_l, 3)

        hbox7 = QHBoxLayout()
        hbox7.addWidget(self.input, 7)
        hbox7.addWidget(self.error, 3)

        # self.textEdit1 = QTextEdit(tabbox)
        # self.textEdit1.setMaximumWidth(int(self.max_sw * 4))

        self.result_table = AutoResizingTable()
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.result_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.result_table.customContextMenuRequested.connect(self.tablemenu)
        self.result_table.setRowCount(0)
        self.result_table.setColumnCount(7)
        self.result_table.horizontalHeader().setFont(font)
        self.result_table.horizontalHeader().setFixedHeight(30)
        # self.result_table.verticalHeader().setFixedWidth(20)
        self.result_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        colname = ["Name", "Sequence", "Reverse Complement",
                   "Length", "GC%", "Tm", "MW"]

        for n in range(7):
            total_width = screen_w - 40
            if n in [1,2]:
                width = int(total_width * 500 / 1650)
            elif n == 0:
                width = int(total_width * 250 / 1650)
            else:
                width = int(total_width * 100 / 1650)
            header_content = QTableWidgetItem(colname[n])
            self.result_table.setHorizontalHeaderItem(n, header_content)
            self.result_table.setColumnWidth(n, width)
            # self.result_table.horizontalHeader().setSectionResizeMode(n, QHeaderView.ResizeToContents)

        tabbox.addLayout(hbox1)
        tabbox.addLayout(hbox2)
        tabbox.addLayout(hbox3)
        tabbox.addLayout(hbox4)
        tabbox.addWidget(line1)
        tabbox.addLayout(hbox5)
        tabbox.addWidget(line2)
        tabbox.addLayout(hbox6)
        tabbox.addLayout(hbox7)
        tabbox.addWidget(self.result_table)

        self.main_frame.setLayout(tabbox)
        self.setCentralWidget(self.main_frame)

        self.setGeometry(0, 0, screen_w, screen_h)
        self.setWindowTitle('Oligo Analyzer')
        self.showMaximized()


    def closeEvent(self, event):
        # print("closeEvent triggered")  # helpful debug
        reply = QMessageBox.question(
            self,
            "Exit App",
            "Are you sure you want to quit?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # print("User confirmed exit.")
            event.accept()
        else:
            # print("User canceled exit.")
            event.ignore()

    def AboutEvent(self):
        """This function is for mainwin help menu bar-about content"""

        msg = QMessageBox()
        msg.setWindowTitle("About")
        # For version control, please update this part when there are some changes to this software
        msg.setText("Version 1.01")
        # msg.setWindowIcon(QtGui.QIcon(self.icon_file))
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    # ---------- helpers ----------

    def log_error(self, message, with_traceback=False):
        """Write a message (optionally with the current exception's traceback) to the error log box."""
        text = message
        if with_traceback:
            text += "\n" + traceback.format_exc()
        self.error.moveCursor(QTextCursor.End)
        self.error.insertPlainText(text.rstrip() + "\n")
        self.error.insertPlainText("----------------------------------------------------\n")
        self.error.moveCursor(QTextCursor.End)
        QtGui.QGuiApplication.processEvents()

    def _set_cell(self, row, col, value, left=True):
        item = QTableWidgetItem("" if value is None else str(value))
        if left:
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.result_table.setItem(row, col, item)

    def _append_rows(self, entries):
        """Add (name, sequence) pairs to the table; unnamed sequences get Seq_<row number>."""
        for name, seq in entries:
            row = self.result_table.rowCount()
            self.result_table.insertRow(row)
            self._set_cell(row, 0, name or "Seq_%d" % (row + 1))
            self._set_cell(row, 1, clean_sequence(seq))
        QtGui.QGuiApplication.processEvents()
        return len(entries)

    # ---------- loading / adding sequences ----------

    def load(self):
        seq_file = QFileDialog.getOpenFileName(
            self, "Load Sequences", "",
            "Sequence files (*.xlsx *.xlsm *.xls *.csv *.tsv *.txt *.fasta *.fa);;All files (*)")[0]

        if seq_file == "":
            self.log_error("No sequence is loaded.")
            return

        try:
            if seq_file.lower().endswith((".xlsx", ".xlsm", ".xls")):
                # Turn each spreadsheet row into a tab-separated line: Name<Tab>Sequence
                sheet = pd.read_excel(seq_file, header=None)
                lines = []
                for _, cells in sheet.iterrows():
                    values = [str(c).strip() for c in cells if pd.notna(c) and str(c).strip()]
                    lines.append("\t".join(values))
                text = "\n".join(lines)
            else:
                with open(seq_file, encoding="utf-8-sig", errors="replace") as f:
                    text = f.read()
        except Exception:
            self.log_error("File format is not recognized: %s" % seq_file, with_traceback=True)
            return

        entries, problems = parse_sequence_text(text)
        for p in problems:
            self.log_error(p)
        added = self._append_rows(entries)
        if added == 0:
            self.log_error("No sequences found in %s." % seq_file)
        self.statusBar().showMessage("Loaded %d sequence(s) from file." % added)

    def add(self):
        text = self.input.toPlainText()
        if not text.strip():
            self.log_error("The input box is empty. Paste one sequence per line as Name,Sequence.")
            return 0

        entries, problems = parse_sequence_text(text)
        for p in problems:
            self.log_error(p)
        added = self._append_rows(entries)
        if added:
            # Clear the box so the same sequences are not added twice
            self.input.clear()
        self.statusBar().showMessage("Added %d sequence(s) to the table." % added)
        return added

    # ---------- IDT analysis ----------

    def get_idt_token(self):
        """Log in to the IDT SciTools API and return an access token."""
        client_id = 'zhijieyang'
        client_secret = 'f2700e1f-20c4-4336-9647-bbbc9c953c52'
        idt_username = 'zhijie.yang@atilabiosystems.com'
        idt_password = 'Atila123'

        authorization_string = b64encode(bytes(client_id + ":" + client_secret, "utf-8")).decode()

        request_headers = {"Content-Type": "application/x-www-form-urlencoded",
                           "Authorization": "Basic " + authorization_string}

        data_dict = {"grant_type": "password",
                     "scope": "test",
                     "username": idt_username,
                     "password": idt_password}

        request_data = parse.urlencode(data_dict).encode()

        post_request = request.Request(IDT_TOKEN_URL,
                                       data=request_data,
                                       headers=request_headers,
                                       method="POST")

        # urlopen raises HTTPError for any non-2xx reply
        response = request.urlopen(post_request, timeout=30)
        body_dict = json.loads(response.read().decode())
        if "access_token" not in body_dict:
            raise RuntimeError("IDT login reply has no access token: %s" % body_dict)
        return body_dict["access_token"]

    def analyze_one(self, sequence, conditions, max_retry=5):
        """Send one sequence to IDT OligoAnalyzer and return its result dict."""
        payload = json.dumps(dict(conditions, Sequence=sequence))
        delay = 0.5
        refreshed = False
        response, last_error = None, None

        for attempt in range(max_retry):
            headers = {'Content-Type': 'application/json',
                       'Authorization': 'Bearer %s' % self.token}
            try:
                response = requests.post(IDT_ANALYZE_URL, headers=headers, data=payload, timeout=30)
            except requests.RequestException as e:
                response, last_error = None, e
            else:
                if response.status_code == 401 and not refreshed:
                    # Token expired during a long batch: log in again once
                    self.token = self.get_idt_token()
                    refreshed = True
                    continue
                if response.status_code != 429 and response.status_code < 500:
                    break
            # Network error, rate limit (429) or server error (5xx): wait and retry
            time.sleep(delay)
            delay *= 2

        if response is None:
            raise RuntimeError("no response from IDT (%s)" % last_error)
        if response.status_code != 200:
            raise RuntimeError("IDT returned HTTP %d: %s" % (response.status_code, response.text[:300]))

        try:
            info = response.json()
        except ValueError:
            raise RuntimeError("IDT reply was not JSON: %s" % response.text[:300])
        if isinstance(info, list):
            info = info[0] if info else {}

        missing = [k for k in ("Length", "GCContent", "MeltTemp", "MolecularWeight") if k not in info]
        if missing:
            raise RuntimeError("unexpected reply from IDT (no %s): %s"
                               % (", ".join(missing), response.text[:300]))
        return info

    def analyze(self):
        # Sequences still sitting in the input box are added to the table first
        if self.input.toPlainText().strip():
            self.add()

        n_rows = self.result_table.rowCount()
        if n_rows == 0:
            self.log_error("Nothing to analyze: the table is empty. Paste sequences in the input box "
                           "(one per line, Name,Sequence) or load a file, then click Analyze.")
            self.statusBar().showMessage("Nothing to analyze")
            return

        try:
            conditions = {"NaConc": float(self.na.text()),
                          "MgConc": float(self.mg.text()),
                          "dNTPsConc": float(self.dntp.text()),
                          "OligoConc": float(self.oligo.text()),
                          "NucleotideType": "DNA"}
        except ValueError:
            self.log_error("Oligo, Na, Mg and dNTPs concentrations must all be numbers.")
            return

        self.statusBar().showMessage("Logging in to IDT...")
        QtGui.QGuiApplication.processEvents()
        try:
            self.token = self.get_idt_token()
        except HTTPError as e:
            detail = e.read().decode("utf-8", errors="replace")[:500]
            self.log_error("IDT login failed (HTTP %s). The IDT account or API key in this script may "
                           "have changed or expired.\n%s" % (e.code, detail))
            self.statusBar().showMessage("IDT login failed - see the error log")
            return
        except URLError:
            self.log_error("Could not connect to IDT. Check the internet connection or VPN/proxy.",
                           with_traceback=True)
            self.statusBar().showMessage("Could not connect to IDT - see the error log")
            return
        except Exception:
            self.log_error("IDT login failed.", with_traceback=True)
            self.statusBar().showMessage("IDT login failed - see the error log")
            return

        done = failed = 0
        for row in range(n_rows):
            self.statusBar().showMessage("Analyzing %d of %d..." % (row + 1, n_rows))
            QtGui.QGuiApplication.processEvents()

            name_item = self.result_table.item(row, 0)
            seq_item = self.result_table.item(row, 1)
            name = name_item.text() if name_item and name_item.text() else "Row %d" % (row + 1)
            sequence = clean_sequence(seq_item.text()) if seq_item else ""

            # Clear old results so a failed re-analysis does not leave stale numbers
            for col in range(2, 7):
                self._set_cell(row, col, "")

            if not sequence:
                self.log_error("%s: no sequence, skipped." % name)
                failed += 1
                continue

            try:
                oligo_info = self.analyze_one(sequence, conditions)
                self._set_cell(row, 2, str(oligo_info.get("Complement", "")).replace(" ", ""))
                self._set_cell(row, 3, oligo_info["Length"], left=False)
                self._set_cell(row, 4, oligo_info["GCContent"], left=False)
                self._set_cell(row, 5, oligo_info["MeltTemp"], left=False)
                self._set_cell(row, 6, oligo_info["MolecularWeight"], left=False)
                done += 1
            except Exception as e:
                # One bad sequence no longer stops the rest of the batch
                self.log_error("%s: %s" % (name, e))
                failed += 1

            QtGui.QGuiApplication.processEvents()
            time.sleep(REQUEST_DELAY)

        self.statusBar().showMessage("Finished: %d analyzed, %d failed%s"
                                     % (done, failed, " - see the error log" if failed else ""))

    def clear_table(self):
        self.result_table.setRowCount(0)

    def tablemenu(self, event):
        tmenu = QMenu(self)
        copy_cells = tmenu.addAction("Copy the selected cells")
        delete_row = tmenu.addAction("Delete rows")
        clear_table = tmenu.addAction("Clear table")
        action = tmenu.exec_(self.mapToGlobal(QPoint(QCursor.pos().x(), QCursor.pos().y() - 30)))
        # objs = self.result_table.selectedIndexes()
        # for obj in objs:
        #     print (obj.row(),obj.column())
        if action == copy_cells:
            objs = self.result_table.selectedIndexes()
            rows = sorted(set([obj.row() for obj in objs]))
            cols = sorted(set([obj.column() for obj in objs]))
            temp_text = ""
            for row in rows:
                for col in cols:
                    item = self.result_table.item(row, col)
                    temp_text += item.text() if item else ""
                    if col != cols[-1]:
                        temp_text += "\t"
                if row != rows[-1]:
                    temp_text += "\n"
            QApplication.clipboard().setText(temp_text)
        elif action == delete_row:

            objs = self.result_table.selectedIndexes()
            rows = list(set([obj.row() for obj in objs]))
            rows.sort(reverse=True)
            for row in rows:
                if row >= 0:
                    self.result_table.removeRow(row)

        elif action == clear_table:
            self.result_table.setRowCount(0)


if __name__ == "__main__":
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    os.environ["QT_SCALE_FACTOR_ROUNDING_POLICY"] = "PassThrough"

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    ex = MainWin()
    sys.exit(app.exec_())