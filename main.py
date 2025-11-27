#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Crypto Exam Generator (v1.8b)

- Export DOCX ze šablony (QWizard) + zachování formátování otázek:
  * tučné/kurzíva/podtržení, barva textu, zarovnání odstavce, odrážky/číslování (jednoduše jako „• “ / „1. “ prefix).
  * Nahrazují se pouze hodnoty v ostrých závorkách <>. Zbytek dokumentu zůstává beze změny.
- Zachováno: autosave (1.7c), WYSIWYG editor bez náhledu (1.7d), import DOCX, DnD, filtr, multiselect.

Poznámka k listům: Word numbering je složitější (numbering.xml). Zde používáme vizuální prefixy („• “, „1. “), aby
se vzhledově zachovala sémantika bez úprav numbering.xml (minimální změna).

Autor: Python PySide6 aplikace - vývoj a úprava
"""
from __future__ import annotations

import json
import sys
import uuid as _uuid
import re
import html as _html
import zipfile
from xml.etree import ElementTree as ET
from dataclasses import dataclass, asdict, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from html.parser import HTMLParser

from PySide6.QtCore import Qt, QSize, QSaveFile, QByteArray, QTimer, QDateTime
from PySide6.QtGui import (
    QAction,
    QActionGroup,
    QKeySequence,
    QTextCharFormat,
    QTextCursor,
    QTextListFormat,
    QTextBlockFormat,
    QColor,
    QPalette,
    QFont,
)
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTreeWidget,
    QTreeWidgetItem,
    QSplitter,
    QToolBar,
    QTextEdit,
    QFileDialog,
    QMessageBox,
    QLineEdit,
    QPushButton,
    QFormLayout,
    QSpinBox,
    QDoubleSpinBox,
    QComboBox,
    QColorDialog,
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QStyle,
    QScrollArea,
    QWizard,
    QWizardPage,
    QDateTimeEdit,
)

APP_NAME = "Crypto Exam Generator"
APP_VERSION = "1.8b"


# --------------------------- Datové typy ---------------------------

@dataclass
class Question:
    id: str
    type: str  # "classic" | "bonus"
    text_html: str
    title: str = ""
    points: int = 1            # jen pro classic
    bonus_correct: float = 0.0 # jen pro bonus
    bonus_wrong: float = 0.0   # jen pro bonus (může být záporné)
    created_at: str = ""       # ISO

    @staticmethod
    def new_default(qtype: str = "classic") -> "Question":
        now = datetime.now().isoformat(timespec="seconds")
        if qtype == "bonus":
            return Question(
                id=str(_uuid.uuid4()),
                type="bonus",
                text_html="<p><br></p>",
                title="BONUS otázka",
                points=0,
                bonus_correct=1.0,
                bonus_wrong=0.0,
                created_at=now,
            )
        return Question(
            id=str(_uuid.uuid4()),
            type="classic",
            text_html="<p><br></p>",
            title="Otázka",
            points=1,
            bonus_correct=0.0,
            bonus_wrong=0.0,
            created_at=now,
        )


@dataclass
class Subgroup:
    id: str
    name: str
    subgroups: List["Subgroup"] = field(default_factory=list)
    questions: List[Question] = field(default_factory=list)


@dataclass
class Group:
    id: str
    name: str
    subgroups: List[Subgroup]


@dataclass
class RootData:
    groups: List[Group]


# --------------------------- Utility: Dark theme ---------------------------

def apply_dark_theme(app: QApplication) -> None:
    QApplication.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(37, 37, 38))
    palette.setColor(QPalette.WindowText, Qt.white)
    palette.setColor(QPalette.Base, QColor(30, 30, 30))
    palette.setColor(QPalette.AlternateBase, QColor(45, 45, 45))
    palette.setColor(QPalette.ToolTipBase, Qt.white)
    palette.setColor(QPalette.ToolTipText, Qt.white)
    palette.setColor(QPalette.Text, Qt.white)
    palette.setColor(QPalette.Button, QColor(45, 45, 48))
    palette.setColor(QPalette.ButtonText, Qt.white)
    palette.setColor(QPalette.BrightText, Qt.red)
    palette.setColor(QPalette.Highlight, QColor(10, 132, 255))
    palette.setColor(QPalette.HighlightedText, Qt.black)
    app.setPalette(palette)


# --------------------------- DnD Tree ---------------------------

class DnDTree(QTreeWidget):
    """QTreeWidget s podporou drag&drop, po přesunu synchronizuje datový model."""

    def __init__(self, owner: "MainWindow") -> None:
        super().__init__()
        self.owner = owner
        self.setHeaderLabels(["Název", "Typ / body"])
        self.setColumnWidth(0, 300)
        self.setContextMenuPolicy(Qt.CustomContextMenu)

        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setDragDropMode(QAbstractItemView.InternalMove)

    def dropEvent(self, event) -> None:
        ids_before = self.owner._selected_question_ids()
        super().dropEvent(event)
        self.owner._sync_model_from_tree()
        self.owner._refresh_tree()
        self.owner._reselect_questions(ids_before)
        self.owner.save_data()
        self.owner.statusBar().showMessage("Přesun dokončen (uloženo).", 3000)


# ---------------------- Dialog pro výběr cíle ----------------------

class MoveTargetDialog(QDialog):
    """Dialog pro výběr cílové skupiny/podskupiny pomocí stromu."""

    def __init__(self, owner: "MainWindow") -> None:
        super().__init__(owner)
        self.setWindowTitle("Vyberte cílovou skupinu/podskupinu")
        self.resize(520, 560)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        self.info = QLabel("Vyberte podskupinu (nebo skupinu – vytvoří se Default).")
        layout.addWidget(self.info)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Název", "Typ"])
        layout.addWidget(self.tree, 1)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        for g in owner.root.groups:
            g_item = QTreeWidgetItem([g.name, "Skupina"])
            g_item.setData(0, Qt.UserRole, {"kind": "group", "id": g.id})
            g_item.setIcon(0, owner.style().standardIcon(QStyle.SP_DirIcon))
            f = g_item.font(0)
            f.setBold(True)
            g_item.setFont(0, f)
            self.tree.addTopLevelItem(g_item)
            self._add_subs(owner, g_item, g.id, g.subgroups)

        self.tree.expandAll()

    def _add_subs(self, owner: "MainWindow", parent_item: QTreeWidgetItem, gid: str, subs: List[Subgroup]) -> None:
        for sg in subs:
            it = QTreeWidgetItem([sg.name, "Podskupina"])
            it.setData(0, Qt.UserRole, {"kind": "subgroup", "id": sg.id, "parent_group_id": gid})
            it.setIcon(0, owner.style().standardIcon(QStyle.SP_DirOpenIcon))
            parent_item.addChild(it)
            if sg.subgroups:
                self._add_subs(owner, it, gid, sg.subgroups)

    def selected_target(self) -> tuple[str, Optional[str]]:
        items = self.tree.selectedItems()
        if not items:
            return "", None
        meta = items[0].data(0, Qt.UserRole) or {}
        if meta.get("kind") == "subgroup":
            return meta.get("parent_group_id"), meta.get("id")
        elif meta.get("kind") == "group":
            return meta.get("id"), None
        return "", None


# --------------------------- Export Wizard ---------------------------

def cz_day_of_week(dt: datetime) -> str:
    days = ["pondělí","úterý","středa","čtvrtek","pátek","sobota","neděle"]
    return days[dt.weekday()]

def round_dt_to_10m(dt: QDateTime) -> QDateTime:
    m = dt.time().minute()
    rounded = (m // 10) * 10
    return QDateTime(dt.date(), dt.time().addSecs((rounded - m) * 60))


# ---- HTML -> jednoduché mezireprezentace pro DOCX ----

from typing import Optional

class HTMLToDocxParser(HTMLParser):
    """
    Převádí podmnožinu HTML z QTextEdit na seznam odstavců:
    paragraphs = [ { 'align': 'left|center|right|justify',
                     'runs': [ {'text': str, 'b':bool, 'i':bool, 'u':bool, 'color': 'RRGGBB'|None } ],
                     'prefix': '' | '• ' | '1. ' | 'a. ' } , ... ]
    """
    def __init__(self) -> None:
        super().__init__()
        self.paragraphs: List[dict] = []
        self._stack: List[dict] = []
        self._current_runs: List[dict] = []
        self._align: str = "left"
        self._in_ul = False
        self._in_ol = False
        self._ol_counter = 0
        self._format = {'b': False, 'i': False, 'u': False, 'color': None}

    def _start_paragraph(self, prefix: str = "") -> None:
        if self._current_runs:
            self._end_paragraph()
        self._current_runs = []
        self._align = "left"
        self._stack.append({'tag': 'p', 'prefix': prefix})

    def _end_paragraph(self) -> None:
        if not self._stack or self._stack[-1].get('tag') != 'p':
            return
        prefix = self._stack[-1].get('prefix', '')
        merged: List[dict] = []
        for r in self._current_runs:
            if r.get('text') == "":
                continue
            if merged and all(merged[-1].get(k) == r.get(k) for k in ('b','i','u','color')):
                merged[-1]['text'] += r['text']
            else:
                merged.append(r.copy())
        self.paragraphs.append({'align': self._align, 'runs': merged, 'prefix': prefix})
        self._current_runs = []
        self._stack.pop()

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        tag = tag.lower()
        if tag in ('p','div'):
            self._start_paragraph()
            style = (attrs.get('style') or '').lower()
            align = None
            m = re.search(r'text-align\s*:\s*(left|center|right|justify)', style)
            if m:
                align = m.group(1)
            if attrs.get('align'):
                align = attrs['align'].lower()
            if align in ('left','center','right','justify'):
                self._align = align
        elif tag == 'br':
            self._append_text("\n")
        elif tag in ('b','strong'):
            self._format['b'] = True
        elif tag in ('i','em'):
            self._format['i'] = True
        elif tag == 'u':
            self._format['u'] = True
        elif tag == 'span':
            style = (attrs.get('style') or '').lower()
            col = None
            m = re.search(r'color\s*:\s*#?([0-9a-f]{6})', style)
            if m:
                col = m.group(1)
            else:
                m2 = re.search(r'color\s*:\s*rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', style)
                if m2:
                    r,g,b = [max(0, min(255, int(x))) for x in m2.groups()]
                    col = f"{r:02X}{g:02X}{b:02X}"
            if col:
                self._stack.append({'tag':'span','prev_color': self._format['color']})
                self._format['color'] = col
            else:
                self._stack.append({'tag':'span','prev_color': self._format['color']})
        elif tag == 'ul':
            self._in_ul = True
        elif tag == 'ol':
            self._in_ol = True
            self._ol_counter = 0
        elif tag == 'li':
            if self._in_ul:
                self._start_paragraph(prefix="• ")
            elif self._in_ol:
                self._ol_counter += 1
                self._start_paragraph(prefix=f"{self._ol_counter}. ")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ('p','div'):
            self._end_paragraph()
        elif tag in ('b','strong'):
            self._format['b'] = False
        elif tag in ('i','em'):
            self._format['i'] = False
        elif tag == 'u':
            self._format['u'] = False
        elif tag == 'span':
            if self._stack and self._stack[-1].get('tag') == 'span':
                self._format['color'] = self._stack[-1].get('prev_color')
                self._stack.pop()
        elif tag == 'ul':
            self._in_ul = False
        elif tag == 'ol':
            self._in_ol = False

    def handle_data(self, data):
        if not data:
            return
        if not self._stack or self._stack[-1].get('tag') != 'p':
            self._start_paragraph()
        self._append_text(data)

    def _append_text(self, text: str):
        self._current_runs.append({
            'text': text,
            'b': self._format['b'],
            'i': self._format['i'],
            'u': self._format['u'],
            'color': self._format['color'],
        })

def parse_html_to_paragraphs(html: str) -> List[dict]:
    parser = HTMLToDocxParser()
    parser.feed(html or "<p></p>")
    if not parser.paragraphs:
        parser._start_paragraph()
        parser._append_text("")
        parser._end_paragraph()
    return parser.paragraphs


# ---- Tvorba WordprocessingML prvků (w:p, w:r) ----

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {'w': W_NS}

def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"

def make_w_run(text: str, b: bool=False, i: bool=False, u: bool=False, color: Optional[str]=None) -> ET.Element:
    r = ET.Element(w_tag("r"))
    rpr = ET.SubElement(r, w_tag("rPr"))
    if b:
        ET.SubElement(rpr, w_tag("b"))
    if i:
        ET.SubElement(rpr, w_tag("i"))
    if u:
        uel = ET.SubElement(rpr, w_tag("u"))
        uel.set(w_tag("val"), "single")
    if color:
        cel = ET.SubElement(rpr, w_tag("color"))
        cel.set(w_tag("val"), color.upper())
    t = ET.SubElement(r, w_tag("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text or ""
    return r

def make_w_paragraph(align: str, runs: List[dict], prefix: str="") -> ET.Element:
    p = ET.Element(w_tag("p"))
    ppr = ET.SubElement(p, w_tag("pPr"))
    jc = ET.SubElement(ppr, w_tag("jc"))
    jc.set(w_tag("val"), {"left":"left","center":"center","right":"right","justify":"both"}.get(align, "left"))
    if prefix:
        p.append(make_w_run(prefix))
    for r in runs:
        p.append(make_w_run(r.get('text',''), r.get('b'), r.get('i'), r.get('u'), r.get('color')))
    return p


# --------------------------- Hlavní okno (část exportu) ---------------------------

class ExportWizard(QWizard):
    def __init__(self, owner: "MainWindow") -> None:
        super().__init__(owner)
        self.setWindowTitle("Export DOCX – průvodce")
        self.setWizardStyle(QWizard.ModernStyle)
        self.owner = owner

        self.template_path: Optional[Path] = None
        self.output_path: Optional[Path] = None
        self.placeholders_q: List[str] = []
        self.placeholders_b: List[str] = []
        self.has_datumcas = False
        self.has_pozn = False
        self.has_minmax = (False, False)

        self._build_pages()

    def _build_pages(self):
        self.page1 = QWizardPage(); self.page1.setTitle("Krok 1/3 – Šablona a parametry")
        l = QVBoxLayout(self.page1)

        form = QFormLayout()
        self.le_template = QLineEdit(); self.btn_template = QPushButton("Vybrat…")
        h1 = QHBoxLayout(); h1.addWidget(self.le_template, 1); h1.addWidget(self.btn_template)
        w1 = QWidget(); w1.setLayout(h1); form.addRow("Šablona DOCX:", w1)

        self.le_output = QLineEdit(); self.btn_output = QPushButton("Uložit jako…")
        h2 = QHBoxLayout(); h2.addWidget(self.le_output, 1); h2.addWidget(self.btn_output)
        w2 = QWidget(); w2.setLayout(h2); form.addRow("Výstupní soubor:", w2)

        self.le_prefix = QLineEdit("MůjText"); form.addRow("Prefix pro <PoznamkaVerze>:", self.le_prefix)
        self.dt_edit = QDateTimeEdit(QDateTime.currentDateTime()); self.dt_edit.setDisplayFormat("dd.MM.yyyy HH:mm"); self.dt_edit.setCalendarPopup(True)
        form.addRow("Datum a čas (<DatumČas>):", self.dt_edit)

        l.addLayout(form)
        self.lbl_info = QLabel("(Po výběru šablony se vyhodnotí placeholdery.)"); l.addWidget(self.lbl_info)

        self.btn_template.clicked.connect(self._choose_template)
        self.btn_output.clicked.connect(self._choose_output)

        self.addPage(self.page1)

        # Page 2
        self.page2 = QWizardPage(); self.page2.setTitle("Krok 2/3 – Výběr otázek")
        l2 = QVBoxLayout(self.page2)
        self.area = QScrollArea(); self.area.setWidgetResizable(True)
        self.inner = QWidget(); self.form2 = QFormLayout(self.inner)
        self.area.setWidget(self.inner); l2.addWidget(self.area)
        self.lbl_hint = QLabel("Vyberte otázky pro <OtázkaX> (klasické) a <BONUSY> (bonusové).")
        l2.addWidget(self.lbl_hint)
        self.addPage(self.page2)

        # Page 3
        self.page3 = QWizardPage(); self.page3.setTitle("Krok 3/3 – Souhrn a export")
        l3 = QFormLayout(self.page3)
        self.lbl_counts = QLabel("-"); self.lbl_minmax = QLabel("-")
        l3.addRow("Souhrn:", self.lbl_counts); l3.addRow("Body:", self.lbl_minmax)
        self.addPage(self.page3)

        self.currentIdChanged.connect(self._on_page_changed)

    def _choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Vybrat šablonu DOCX", str(self.owner.project_root), "Word dokument (*.docx)")
        if not path:
            return
        self.le_template.setText(path); self.template_path = Path(path); self._scan_placeholders()

    def _choose_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Uložit výstupní DOCX", str(self.owner.project_root / "vystup.docx"), "Word dokument (*.docx)")
        if path:
            self.le_output.setText(path); self.output_path = Path(path)

    def _scan_placeholders(self):
        try:
            with zipfile.ZipFile(self.template_path, "r") as z:
                texts = []
                for name in z.namelist():
                    if name.startswith("word/") and name.endswith(".xml"):
                        texts.append(z.read(name).decode("utf-8", errors="ignore"))
                xml = "\n".join(texts)
        except Exception as e:
            QMessageBox.warning(self, "Šablona", f"Nelze číst šablonu:\n{e}"); return

        ph = re.findall(r"&lt;([^&<>]+)&gt;", xml) + re.findall(r"<([A-Za-z0-9ÁČĎÉĚÍŇÓŘŠŤÚŮÝŽáčďéěíňóřšťúůýž]+[0-9]*)>", xml)
        seen, uniq = set(), []
        for p in ph:
            if p not in seen:
                uniq.append(p); seen.add(p)

        self.placeholders_q = sorted([p for p in uniq if re.match(r"^Otázka\d+$", p)], key=lambda x: int(re.findall(r"\d+", x)[0]))
        self.placeholders_b = sorted([p for p in uniq if re.match(r"^BONUS\d+$", p)], key=lambda x: int(re.findall(r"\d+", x)[0]))
        self.has_datumcas = "DatumČas" in uniq
        self.has_pozn = "PoznamkaVerze" in uniq
        self.has_minmax = ("MinBody" in uniq, "MaxBody" in uniq)

        self._rebuild_page2()
        self.lbl_info.setText(f"Nalezeno {len(self.placeholders_q)}× Otázka, {len(self.placeholders_b)}× BONUS; "
                              f"{'DatumČas ' if self.has_datumcas else ''}{'PoznamkaVerze ' if self.has_pozn else ''}"
                              f"{'MinBody ' if self.has_minmax[0] else ''}{'MaxBody ' if self.has_minmax[1] else ''}")

    def _rebuild_page2(self):
        while self.form2.rowCount():
            self.form2.removeRow(0)

        self.cmb_q: Dict[str, QComboBox] = {}
        self.cmb_b: Dict[str, QComboBox] = {}

        classics = self.owner._all_questions_by_type("classic")
        bonuses  = self.owner._all_questions_by_type("bonus")

        for name in self.placeholders_q:
            cb = QComboBox(); cb.addItem("-- nevybráno --", "")
            for q in classics:
                cb.addItem(q.title or "Otázka", q.id)
            self.form2.addRow(name + ":", cb); self.cmb_q[name] = cb
            cb.currentIndexChanged.connect(self._update_summary)

        for name in self.placeholders_b:
            cb = QComboBox(); cb.addItem("-- nevybráno --", "")
            for q in bonuses:
                cb.addItem(f"{q.title or 'BONUS'} (+{q.bonus_correct:.2f}/ {q.bonus_wrong:.2f})", q.id)
            self.form2.addRow(name + ":", cb); self.cmb_b[name] = cb
            cb.currentIndexChanged.connect(self._update_summary)

    def _on_page_changed(self, idx: int):
        if idx == 2:
            self._update_summary()

    def _update_summary(self):
        classics_sel, bonuses_sel = [], []
        for cb in getattr(self, 'cmb_q', {}).values():
            qid = cb.currentData()
            if qid:
                q = self.owner._find_question_by_id(qid)
                if q: classics_sel.append(q)
        for cb in getattr(self, 'cmb_b', {}).values():
            qid = cb.currentData()
            if qid:
                q = self.owner._find_question_by_id(qid)
                if q: bonuses_sel.append(q)
        min_body = sum(float(q.bonus_wrong) for q in bonuses_sel)
        max_body = sum(int(q.points) for q in classics_sel) + sum(float(q.bonus_correct) for q in bonuses_sel)
        self.lbl_counts.setText(f"Klasických: {len(classics_sel)} / {len(self.placeholders_q)}, BONUS: {len(bonuses_sel)} / {len(self.placeholders_b)}")
        self.lbl_minmax.setText(f"MinBody = {min_body:.2f} | MaxBody = {max_body:.2f}")

    def accept(self) -> None:
        if not self.template_path or not self.output_path:
            QMessageBox.warning(self, "Export", "Vyberte šablonu i výstupní soubor."); return

        repl_plain: Dict[str, str] = {}

        if self.has_datumcas:
            dt = round_dt_to_10m(self.dt_edit.dateTime())
            repl_plain["DatumČas"] = f"{cz_day_of_week(dt.toPython())} {dt.toString('dd.MM.yyyy HH:mm')}"
        if self.has_pozn:
            prefix = (self.le_prefix.text().strip() or "MůjText")
            today = datetime.now().strftime("%Y-%m-%d"); rnd = str(_uuid.uuid4())[:8]
            repl_plain["PoznamkaVerze"] = f"{prefix}_{today}_{rnd}"

        classics_sel: Dict[str, Question] = {}
        bonuses_sel: Dict[str, Question] = {}

        for name, cb in self.cmb_q.items():
            qid = cb.currentData()
            if not qid: QMessageBox.warning(self, "Export", f"Vyberte otázku pro {name}."); return
            q = self.owner._find_question_by_id(qid)
            if not q or q.type != "classic": QMessageBox.warning(self, "Export", f"{name} musí být klasická."); return
            classics_sel[name] = q

        for name, cb in self.cmb_b.items():
            qid = cb.currentData()
            if not qid: QMessageBox.warning(self, "Export", f"Vyberte otázku pro {name}."); return
            q = self.owner._find_question_by_id(qid)
            if not q or q.type != "bonus": QMessageBox.warning(self, "Export", f"{name} musí být BONUS."); return
            bonuses_sel[name] = q

        if self.has_minmax[0]:
            repl_plain["MinBody"] = f"{sum(float(q.bonus_wrong) for q in bonuses_sel.values()):.2f}"
        if self.has_minmax[1]:
            maxb = sum(int(q.points) for q in classics_sel.values()) + sum(float(q.bonus_correct) for q in bonuses_sel.values())
            repl_plain["MaxBody"] = f"{maxb:.2f}"

        rich_map: Dict[str, str] = {}
        for k, q in classics_sel.items():
            rich_map[k] = q.text_html or "<p></p>"
        for k, q in bonuses_sel.items():
            rich_map[k] = q.text_html or "<p></p>"

        try:
            self.owner._generate_docx_from_template(self.template_path, self.output_path, repl_plain, rich_map)
        except Exception as e:
            QMessageBox.critical(self, "Export", f"Chyba při exportu:\n{e}"); return

        QMessageBox.information(self, "Export", f"Hotovo:\n{self.output_path}")
        super().accept()


# --------------------------- Hlavní okno (UI + logika) ---------------------------

class MainWindow(QMainWindow):
    """Hlavní okno aplikace."""

    def _selected_question_ids(self) -> List[str]:
        ids: List[str] = []
        for it in self.tree.selectedItems():
            meta = it.data(0, Qt.UserRole) or {}
            if meta.get("kind") == "question":
                ids.append(meta.get("id"))
        return ids

    def _reselect_questions(self, ids: List[str]) -> None:
        if not ids:
            return
        wanted = set(ids)
        self.tree.clearSelection()

        def walk(item: QTreeWidgetItem):
            meta = item.data(0, Qt.UserRole) or {}
            if meta.get("kind") == "question" and meta.get("id") in wanted:
                item.setSelected(True)
            for i in range(item.childCount()):
                walk(item.child(i))

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))

        items = self.tree.selectedItems()
        if items:
            self.tree.setCurrentItem(items[0])
            self._on_tree_selection_changed()

    def __init__(self, data_path: Optional[Path] = None) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1200, 880)

        self.project_root = Path.cwd()
        default_data_dir = self.project_root / "data"
        default_data_dir.mkdir(parents=True, exist_ok=True)
        self.data_path = data_path or (default_data_dir / "questions.json")

        self.root: RootData = RootData(groups=[])
        self._current_question_id: Optional[str] = None
        self._current_node_kind: Optional[str] = None

        self._autosave_timer = QTimer(self)
        self._autosave_timer.setSingleShot(True)
        self._autosave_timer.setInterval(1200)
        self._autosave_timer.timeout.connect(self._autosave_current_question)

        self._build_ui()
        self._connect_signals()
        self._build_menus()
        self.load_data()
        self._refresh_tree()

    def _build_ui(self) -> None:
        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        filter_bar = QWidget()
        filter_layout = QHBoxLayout(filter_bar)
        filter_layout.setContentsMargins(6, 6, 6, 0)
        filter_layout.setSpacing(6)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Filtr: název / obsah otázky…")
        self.btn_move_selected = QPushButton("Přesunout vybrané…")
        self.btn_delete_selected = QPushButton("Smazat vybrané")
        filter_layout.addWidget(self.filter_edit, 1)
        filter_layout.addWidget(self.btn_move_selected)
        filter_layout.addWidget(self.btn_delete_selected)
        left_layout.addWidget(filter_bar)

        self.tree = DnDTree(self)
        left_layout.addWidget(self.tree, 1)

        self.detail_stack = QWidget()
        self.detail_layout = QVBoxLayout(self.detail_stack)
        self.detail_layout.setContentsMargins(6, 6, 6, 6)
        self.detail_layout.setSpacing(8)

        self.editor_toolbar = QToolBar("Formát")
        self.editor_toolbar.setIconSize(QSize(18, 18))
        self.action_bold = QAction("Tučné", self); self.action_bold.setCheckable(True); self.action_bold.setShortcut(QKeySequence.Bold)
        self.action_italic = QAction("Kurzíva", self); self.action_italic.setCheckable(True); self.action_italic.setShortcut(QKeySequence.Italic)
        self.action_underline = QAction("Podtržení", self); self.action_underline.setCheckable(True); self.action_underline.setShortcut(QKeySequence.Underline)
        self.action_color = QAction("Barva", self)
        self.action_bullets = QAction("Odrážky", self); self.action_bullets.setCheckable(True)

        self.action_align_left = QAction("Vlevo", self)
        self.action_align_center = QAction("Na střed", self)
        self.action_align_right = QAction("Vpravo", self)
        self.action_align_justify = QAction("Do bloku", self)
        self.align_group = QActionGroup(self)
        for a in (self.action_align_left, self.action_align_center, self.action_align_right, self.action_align_justify):
            a.setCheckable(True); self.align_group.addAction(a)

        self.editor_toolbar.addAction(self.action_bold)
        self.editor_toolbar.addAction(self.action_italic)
        self.editor_toolbar.addAction(self.action_underline)
        self.editor_toolbar.addSeparator()
        self.editor_toolbar.addAction(self.action_color)
        self.editor_toolbar.addSeparator()
        self.editor_toolbar.addAction(self.action_bullets)
        self.editor_toolbar.addSeparator()
        self.editor_toolbar.addAction(self.action_align_left)
        self.editor_toolbar.addAction(self.action_align_center)
        self.editor_toolbar.addAction(self.action_align_right)
        self.editor_toolbar.addAction(self.action_align_justify)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)

        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("Krátký název otázky…")

        self.combo_type = QComboBox(); self.combo_type.addItems(["Klasická", "BONUS"])
        self.spin_points = QSpinBox(); self.spin_points.setRange(-999, 999); self.spin_points.setValue(1)
        self.spin_bonus_correct = QDoubleSpinBox(); self.spin_bonus_correct.setDecimals(2); self.spin_bonus_correct.setSingleStep(0.01); self.spin_bonus_correct.setRange(-999.99, 999.99); self.spin_bonus_correct.setValue(1.00)
        self.spin_bonus_wrong = QDoubleSpinBox(); self.spin_bonus_wrong.setDecimals(2); self.spin_bonus_wrong.setSingleStep(0.01); self.spin_bonus_wrong.setRange(-999.99, 999.99); self.spin_bonus_wrong.setValue(0.00)

        form.addRow("Název otázky:", self.title_edit)
        form.addRow("Typ otázky:", self.combo_type)
        form.addRow("Body (klasická):", self.spin_points)
        form.addRow("Body za správně (BONUS):", self.spin_bonus_correct)
        form.addRow("Body za špatně (BONUS):", self.spin_bonus_wrong)

        self.text_edit = QTextEdit()
        self.text_edit.setAcceptRichText(True)
        self.text_edit.setPlaceholderText("Sem napište znění otázky…\nPodporováno: tučné, kurzíva, podtržení, barva, odrážky, zarovnání.")
        self.text_edit.setMinimumHeight(360)

        self.btn_save_question = QPushButton("Uložit změny otázky"); self.btn_save_question.setDefault(True)

        self.rename_panel = QWidget()
        rename_layout = QFormLayout(self.rename_panel)
        self.rename_line = QLineEdit()
        self.btn_rename = QPushButton("Uložit název")
        rename_layout.addRow("Název:", self.rename_line)
        rename_layout.addRow(self.btn_rename)

        self.detail_layout.addWidget(self.editor_toolbar)
        self.detail_layout.addLayout(form)
        self.detail_layout.addWidget(self.text_edit, 1)
        self.detail_layout.addWidget(self.btn_save_question)
        self.detail_layout.addWidget(self.rename_panel)
        self._set_editor_enabled(False)

        splitter.addWidget(left_panel)
        splitter.addWidget(self.detail_stack)
        splitter.setStretchFactor(1, 1)
        self.setCentralWidget(splitter)

        tb = self.addToolBar("Hlavní")
        tb.setIconSize(QSize(18, 18))
        self.act_add_group = QAction("Přidat skupinu", self)
        self.act_add_subgroup = QAction("Přidat podskupinu", self)
        self.act_add_question = QAction("Přidat otázku", self)
        self.act_delete = QAction("Smazat", self)
        self.act_save_all = QAction("Uložit vše", self)
        self.act_choose_data = QAction("Zvolit soubor s daty…", self)

        self.act_add_group.setShortcut("Ctrl+G")
        self.act_add_subgroup.setShortcut("Ctrl+Shift+G")
        self.act_add_question.setShortcut(QKeySequence.New)
        self.act_delete.setShortcut(QKeySequence.Delete)
        self.act_save_all.setShortcut(QKeySequence.Save)

        tb.addAction(self.act_add_group)
        tb.addAction(self.act_add_subgroup)
        tb.addAction(self.act_add_question)
        tb.addSeparator()
        tb.addAction(self.act_delete)
        tb.addSeparator()
        tb.addAction(self.act_save_all)
        tb.addSeparator()
        tb.addAction(self.act_choose_data)

        self.statusBar().showMessage(f"Datový soubor: {self.data_path}")

    def _set_editor_enabled(self, enabled: bool) -> None:
        self.editor_toolbar.setEnabled(enabled)
        self.title_edit.setEnabled(enabled)
        self.combo_type.setEnabled(enabled)
        self.spin_points.setEnabled(enabled)
        self.spin_bonus_correct.setEnabled(enabled and self.combo_type.currentIndex() == 1)
        self.spin_bonus_wrong.setEnabled(enabled and self.combo_type.currentIndex() == 1)
        self.text_edit.setEnabled(enabled)
        self.btn_save_question.setEnabled(enabled)

    def _connect_signals(self) -> None:
        self.tree.itemSelectionChanged.connect(self._on_tree_selection_changed)
        self.btn_save_question.clicked.connect(self._on_save_question_clicked)
        self.btn_rename.clicked.connect(self._on_rename_clicked)
        self.combo_type.currentIndexChanged.connect(self._on_type_changed_ui)

        self.action_bold.triggered.connect(lambda: self._toggle_format("bold"))
        self.action_italic.triggered.connect(lambda: self._toggle_format("italic"))
        self.action_underline.triggered.connect(lambda: self._toggle_format("underline"))
        self.action_color.triggered.connect(self._choose_color)
        self.action_bullets.triggered.connect(self._toggle_bullets)
        self.action_align_left.triggered.connect(lambda: self._apply_alignment(Qt.AlignLeft))
        self.action_align_center.triggered.connect(lambda: self._apply_alignment(Qt.AlignHCenter))
        self.action_align_right.triggered.connect(lambda: self._apply_alignment(Qt.AlignRight))
        self.action_align_justify.triggered.connect(lambda: self._apply_alignment(Qt.AlignJustify))

        self.text_edit.cursorPositionChanged.connect(self._sync_toolbar_to_cursor)

        self.text_edit.textChanged.connect(self._autosave_schedule)
        self.title_edit.textChanged.connect(self._autosave_schedule)
        self.combo_type.currentIndexChanged.connect(self._autosave_schedule)
        self.spin_points.valueChanged.connect(self._autosave_schedule)
        self.spin_bonus_correct.valueChanged.connect(self._autosave_schedule)
        self.spin_bonus_wrong.valueChanged.connect(self._autosave_schedule)

        self.act_add_group.triggered.connect(self._add_group)
        self.act_add_subgroup.triggered.connect(self._add_subgroup)
        self.act_add_question.triggered.connect(self._add_question)
        self.act_delete.triggered.connect(self._delete_selected)
        self.act_save_all.triggered.connect(self.save_data)
        self.act_choose_data.triggered.connect(self._choose_data_file)

        self.filter_edit.textChanged.connect(self._apply_filter)
        self.btn_move_selected.clicked.connect(self._bulk_move_selected)
        self.btn_delete_selected.clicked.connect(self._bulk_delete_selected)

    def _build_menus(self) -> None:
        bar = self.menuBar()
        self.file_menu = bar.addMenu("Soubor")
        edit_menu = bar.addMenu("Úpravy")

        self.act_import_docx = QAction("Import z DOCX…", self)
        self.act_move_question = QAction("Přesunout otázku…", self)
        self.act_move_selected = QAction("Přesunout vybrané…", self)
        self.act_delete_selected = QAction("Smazat vybrané", self)
        self.act_export_docx = QAction("Export do DOCX (šablona)…", self)

        self.file_menu.addAction(self.act_import_docx)
        self.file_menu.addAction(self.act_export_docx)
        edit_menu.addAction(self.act_move_question)
        edit_menu.addAction(self.act_move_selected)
        edit_menu.addAction(self.act_delete_selected)

        self.act_import_docx.setShortcut("Ctrl+I")
        self.act_import_docx.triggered.connect(self._import_from_docx)
        self.act_move_question.triggered.connect(self._move_question)
        self.act_move_selected.triggered.connect(self._bulk_move_selected)
        self.act_delete_selected.triggered.connect(self._bulk_delete_selected)
        self.act_export_docx.triggered.connect(self._export_docx_wizard)

        tb_import = self.addToolBar("Import/Export")
        tb_import.setIconSize(QSize(18, 18))
        tb_import.addAction(self.act_import_docx)
        tb_import.addAction(self.act_export_docx)

    # -------------------- Práce s daty (JSON) --------------------

    def default_root_obj(self) -> RootData:
        return RootData(groups=[])

    def load_data(self) -> None:
        if self.data_path.exists():
            try:
                with self.data_path.open("r", encoding="utf-8") as f:
                    raw = json.load(f)
                groups: List[Group] = []
                for g in raw.get("groups", []):
                    groups.append(self._parse_group(g))
                self.root = RootData(groups=groups)
            except Exception as e:
                QMessageBox.warning(self, "Načtení selhalo", f"Soubor {self.data_path} nelze načíst: {e}\nVytvořen prázdný projekt.")
                self.root = self.default_root_obj()
        else:
            self.root = self.default_root_obj()

    def save_data(self) -> None:
        self._apply_editor_to_current_question(silent=True)
        self.data_path.parent.mkdir(parents=True, exist_ok=True)
        data = {"groups": [self._serialize_group(g) for g in self.root.groups]}
        try:
            from PySide6.QtCore import QSaveFile, QByteArray
            sf = QSaveFile(str(self.data_path))
            sf.open(QSaveFile.WriteOnly)
            payload = json.dumps(data, ensure_ascii=False, indent=2)
            sf.write(QByteArray(payload.encode("utf-8")))
            sf.commit()
            self.statusBar().showMessage(f"Uloženo: {self.data_path}", 1500)
        except Exception as e:
            QMessageBox.critical(self, "Uložení selhalo", f"Chyba při ukládání do {self.data_path}:\n{e}")

    def _parse_group(self, g: dict) -> Group:
        subgroups = [self._parse_subgroup(sg) for sg in g.get("subgroups", [])]
        return Group(id=g["id"], name=g["name"], subgroups=subgroups)

    def _parse_subgroup(self, sg: dict) -> Subgroup:
        subgroups_raw = sg.get("subgroups", [])
        subgroups = [self._parse_subgroup(s) for s in subgroups_raw]
        questions = [self._parse_question(q) for q in sg.get("questions", [])]
        return Subgroup(id=sg["id"], name=sg["name"], subgroups=subgroups, questions=questions)

    def _parse_question(self, q: dict) -> Question:
        title = q.get("title") or self._derive_title_from_html(q.get("text_html") or "<p></p>", prefix=("BONUS: " if q.get("type") == "bonus" else ""))
        bc_default = 1.0 if q.get("type") == "bonus" else 0.0
        bw_default = 0.0
        try:
            bc = round(float(q.get("bonus_correct", bc_default)), 2)
        except Exception:
            bc = round(float(bc_default), 2)
        try:
            bw = round(float(q.get("bonus_wrong", bw_default)), 2)
        except Exception:
            bw = round(float(bw_default), 2)
        return Question(
            id=q.get("id", ""),
            type=q.get("type", "classic"),
            text_html=q.get("text_html", "<p><br></p>"),
            title=title,
            points=int(q.get("points", 1)),
            bonus_correct=bc,
            bonus_wrong=bw,
            created_at=q.get("created_at", ""),
        )

    def _serialize_group(self, g: Group) -> dict:
        return {"id": g.id, "name": g.name, "subgroups": [self._serialize_subgroup(sg) for sg in g.subgroups]}

    def _serialize_subgroup(self, sg: Subgroup) -> dict:
        return {"id": sg.id, "name": sg.name, "subgroups": [self._serialize_subgroup(s) for s in sg.subgroups], "questions": [asdict(q) for q in sg.questions]}

    # -------------------- Tree helpery --------------------

    def _bonus_points_label(self, q: Question) -> str:
        return f"+{q.bonus_correct:.2f}/ {q.bonus_wrong:.2f}"

    def _apply_question_item_visuals(self, item: QTreeWidgetItem, q_type: str) -> None:
        item.setIcon(0, self.style().standardIcon(QStyle.SP_FileIcon))
        if q_type == "bonus":
            item.setForeground(0, self.palette().brush(QPalette.Highlight))
            f = item.font(0); f.setBold(True); item.setFont(0, f)
        else:
            item.setForeground(0, self.palette().brush(QPalette.Text))
            f = item.font(0); f.setBold(False); item.setFont(0, f)

    def _refresh_tree(self) -> None:
        self.tree.clear()
        for g in self.root.groups:
            g_item = QTreeWidgetItem([g.name, "Skupina"])
            g_item.setData(0, Qt.UserRole, {"kind": "group", "id": g.id})
            g_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
            f = g_item.font(0); f.setBold(True); g_item.setFont(0, f)
            self.tree.addTopLevelItem(g_item)
            self._add_subgroups_to_item(g_item, g.id, g.subgroups)
        self.tree.expandAll(); self.tree.resizeColumnToContents(0)

    def _add_subgroups_to_item(self, parent_item: QTreeWidgetItem, group_id: str, subgroups: List[Subgroup]) -> None:
        for sg in subgroups:
            parent_meta = parent_item.data(0, Qt.UserRole) or {}
            parent_sub_id = parent_meta.get("id") if parent_meta.get("kind") == "subgroup" else None
            sg_item = QTreeWidgetItem([sg.name, "Podskupina"])
            sg_item.setData(0, Qt.UserRole, {"kind": "subgroup", "id": sg.id, "parent_group_id": group_id, "parent_subgroup_id": parent_sub_id})
            sg_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirOpenIcon))
            parent_item.addChild(sg_item)
            for q in sg.questions:
                label = "Klasická" if q.type == "classic" else "BONUS"
                pts = q.points if q.type == "classic" else self._bonus_points_label(q)
                q_item = QTreeWidgetItem([q.title or "Otázka", f"{label} | {pts}"])
                q_item.setData(0, Qt.UserRole, {"kind": "question", "id": q.id, "parent_group_id": group_id, "parent_subgroup_id": sg.id})
                self._apply_question_item_visuals(q_item, q.type)
                sg_item.addChild(q_item)
            if sg.subgroups:
                self._add_subgroups_to_item(sg_item, group_id, sg.subgroups)

    def _selected_node(self) -> Tuple[Optional[str], Optional[dict]]:
        items = self.tree.selectedItems()
        if not items:
            return None, None
        item = items[0]
        meta = item.data(0, Qt.UserRole)
        if not meta:
            return None, None
        return meta.get("kind"), meta

    def _sync_model_from_tree(self) -> None:
        group_map = {g.id: g for g in self.root.groups}
        subgroup_map: dict[str, Subgroup] = {}
        question_map: dict[str, Question] = {}

        def scan_subgroups(lst: List[Subgroup]):
            for sg in lst:
                subgroup_map[sg.id] = sg
                for q in sg.questions:
                    question_map[q.id] = q
                scan_subgroups(sg.subgroups)

        for g in self.root.groups:
            scan_subgroups(g.subgroups)

        new_groups: List[Group] = []

        def build_from_item(item: QTreeWidgetItem, container) -> None:
            for i in range(item.childCount()):
                ch = item.child(i)
                meta = ch.data(0, Qt.UserRole) or {}
                kind = meta.get("kind")
                if kind == "subgroup":
                    old = subgroup_map.get(meta["id"])
                    if not old:
                        continue
                    new_sg = Subgroup(id=old.id, name=ch.text(0), subgroups=[], questions=[])
                    container.subgroups.append(new_sg)
                    build_from_item(ch, new_sg)
                elif kind == "question":
                    q = question_map.get(meta["id"])
                    if not q:
                        continue
                    if isinstance(container, Group):
                        if not container.subgroups:
                            container.subgroups.append(Subgroup(id=str(_uuid.uuid4()), name="Default", subgroups=[], questions=[]))
                        container.subgroups[0].questions.append(q)
                    else:
                        container.questions.append(q)

        for i in range(self.tree.topLevelItemCount()):
            gi = self.tree.topLevelItem(i)
            meta = gi.data(0, Qt.UserRole) or {}
            if meta.get("kind") != "group":
                continue
            old_g = group_map.get(meta["id"])
            if not old_g:
                old_g = Group(id=meta["id"], name=gi.text(0), subgroups=[])
            new_g = Group(id=old_g.id, name=gi.text(0), subgroups=[])
            build_from_item(gi, new_g)
            new_groups.append(new_g)

        self.root.groups = new_groups

    # -------------------- Akce: přidání/mazání/přejmenování --------------------

    def _add_group(self) -> None:
        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(self, "Nová skupina", "Název skupiny:")
        if not ok or not name.strip():
            return
        g = Group(id=str(_uuid.uuid4()), name=name.strip(), subgroups=[])
        self.root.groups.append(g)
        self._refresh_tree()
        self.save_data()

    def _add_subgroup(self) -> None:
        kind, meta = self._selected_node()
        if kind not in ("group", "subgroup"):
            QMessageBox.information(self, "Výběr", "Vyberte skupinu (nebo podskupinu) pro přidání podskupiny.")
            return
        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(self, "Nová podskupina", "Název podskupiny:")
        if not ok or not name.strip():
            return

        if kind == "group":
            g = self._find_group(meta["id"])
            if not g:
                return
            g.subgroups.append(Subgroup(id=str(_uuid.uuid4()), name=name.strip(), subgroups=[], questions=[]))
        else:
            parent_sg = self._find_subgroup(meta["parent_group_id"], meta["id"])
            if not parent_sg:
                return
            parent_sg.subgroups.append(Subgroup(id=str(_uuid.uuid4()), name=name.strip(), subgroups=[], questions=[]))

        self._refresh_tree()
        self.save_data()

    def _add_question(self) -> None:
        kind, meta = self._selected_node()
        if kind not in ("group", "subgroup"):
            QMessageBox.information(self, "Výběr", "Vyberte skupinu nebo podskupinu, do které chcete přidat otázku.")
            return

        target_sg: Optional[Subgroup] = None
        if kind == "group":
            g = self._find_group(meta["id"])
            if not g:
                return
            if not g.subgroups:
                sg = Subgroup(id=str(_uuid.uuid4()), name="Default", subgroups=[], questions=[])
                g.subgroups.append(sg)
                target_sg = sg
            else:
                target_sg = g.subgroups[0]
        else:
            target_sg = self._find_subgroup(meta["parent_group_id"], meta["id"])
        if not target_sg:
            return

        q = Question.new_default("classic")
        target_sg.questions.append(q)
        self._refresh_tree()
        self._select_question(q.id)
        self.save_data()

    def _delete_selected(self) -> None:
        kind, meta = self._selected_node()
        if not kind:
            return
        if kind == "question":
            qid = meta["id"]
            gid = meta["parent_group_id"]
            sgid = meta["parent_subgroup_id"]
            if QMessageBox.question(self, "Smazat otázku", "Opravdu smazat vybranou otázku?") == QMessageBox.Yes:
                sg = self._find_subgroup(gid, sgid)
                if sg:
                    sg.questions = [q for q in sg.questions if q.id != qid]
                    self._refresh_tree()
                    self._clear_editor()
                    self.save_data()
        elif kind == "subgroup":
            gid = meta["parent_group_id"]
            sgid = meta["id"]
            if QMessageBox.question(self, "Smazat podskupinu", "Smazat podskupinu včetně podřízených podskupin a otázek?") == QMessageBox.Yes:
                parent_sgid = meta.get("parent_subgroup_id")
                if parent_sgid:
                    parent = self._find_subgroup(gid, parent_sgid)
                    if parent:
                        parent.subgroups = [s for s in parent.subgroups if s.id != sgid]
                else:
                    g = self._find_group(gid)
                    if g:
                        g.subgroups = [s for s in g.subgroups if s.id != sgid]
                self._refresh_tree()
                self._clear_editor()
                self.save_data()
        elif kind == "group":
            gid = meta["id"]
            if QMessageBox.question(self, "Smazat skupinu", "Smazat celou skupinu včetně podskupin a otázek?") == QMessageBox.Yes:
                self.root.groups = [g for g in self.root.groups if g.id != gid]
                self._refresh_tree()
                self._clear_editor()
                self.save_data()

    def _bulk_delete_selected(self) -> None:
        items = [it for it in self.tree.selectedItems() if (it.data(0, Qt.UserRole) or {}).get("kind") == "question"]
        if not items:
            QMessageBox.information(self, "Smazat vybrané", "Vyberte ve stromu alespoň jednu otázku.")
            return
        if QMessageBox.question(self, "Smazat vybrané", f"Opravdu smazat {len(items)} vybraných otázek?") != QMessageBox.Yes:
            return
        deleted = 0
        for it in items:
            meta = it.data(0, Qt.UserRole) or {}
            gid = meta.get("parent_group_id")
            sgid = meta.get("parent_subgroup_id")
            qid = meta.get("id")
            sg = self._find_subgroup(gid, sgid)
            if sg:
                sg.questions = [q for q in sg.questions if q.id != qid]
                deleted += 1
        self._refresh_tree()
        self._clear_editor()
        self.save_data()
        self.statusBar().showMessage(f"Smazáno {deleted} otázek.", 4000)

    def _on_rename_clicked(self) -> None:
        kind, meta = self._selected_node()
        if kind not in ("group", "subgroup"):
            return
        new_name = self.rename_line.text().strip()
        if not new_name:
            return
        if kind == "group":
            g = self._find_group(meta["id"])
            if g:
                g.name = new_name
        else:
            sg = self._find_subgroup(meta["parent_group_id"], meta["id"])
            if sg:
                sg.name = new_name
        self._refresh_tree()
        self.save_data()

    def _on_tree_selection_changed(self) -> None:
        kind, meta = self._selected_node()
        self._current_node_kind = kind
        if kind == "question":
            q = self._find_question(meta["parent_group_id"], meta["parent_subgroup_id"], meta["id"])
            if q:
                self._load_question_to_editor(q)
                self._set_editor_enabled(True)
                self.rename_panel.hide()
        elif kind in ("group", "subgroup"):
            name = ""
            if kind == "group":
                g = self._find_group(meta["id"]); name = g.name if g else ""
            else:
                sg = self._find_subgroup(meta["parent_group_id"], meta["id"]); name = sg.name if sg else ""
            self.rename_line.setText(name)
            self._set_editor_enabled(False)
            self.rename_panel.show()
        else:
            self._clear_editor(); self.rename_panel.hide()

    def _clear_editor(self) -> None:
        self._current_question_id = None
        self.text_edit.clear()
        self.spin_points.setValue(1)
        self.spin_bonus_correct.setValue(1.00)
        self.spin_bonus_wrong.setValue(0.00)
        self.combo_type.setCurrentIndex(0)
        self.title_edit.clear()
        self._set_editor_enabled(False)

    def _load_question_to_editor(self, q: Question) -> None:
        self._current_question_id = q.id
        self.combo_type.setCurrentIndex(0 if q.type == "classic" else 1)
        self.spin_points.setValue(int(q.points))
        self.spin_bonus_correct.setValue(float(q.bonus_correct))
        self.spin_bonus_wrong.setValue(float(q.bonus_wrong))
        self.text_edit.setHtml(q.text_html or "<p><br></p>")
        self.title_edit.setText(q.title or self._derive_title_from_html(q.text_html))
        self._set_editor_enabled(True)
        self.rename_panel.hide()

    def _apply_editor_to_current_question(self, silent: bool = False) -> None:
        if not self._current_question_id:
            return
        def apply_in(sgs: List[Subgroup]) -> bool:
            for sg in sgs:
                for i, q in enumerate(sg.questions):
                    if q.id == self._current_question_id:
                        q.type = "classic" if self.combo_type.currentIndex() == 0 else "bonus"
                        q.text_html = self.text_edit.toHtml()
                        q.title = (self.title_edit.text().strip() or self._derive_title_from_html(q.text_html, prefix=("BONUS: " if q.type == "bonus" else "")))
                        if q.type == "classic":
                            q.points = int(self.spin_points.value()); q.bonus_correct = 0.0; q.bonus_wrong = 0.0
                        else:
                            q.points = 0; q.bonus_correct = round(float(self.spin_bonus_correct.value()), 2); q.bonus_wrong = round(float(self.spin_bonus_wrong.value()), 2)
                        sg.questions[i] = q
                        label = "Klasická" if q.type == "classic" else "BONUS"
                        pts = q.points if q.type == "classic" else self._bonus_points_label(q)
                        self._update_selected_question_item_title(q.title)
                        self._update_selected_question_item_subtitle(f"{label} | {pts}")
                        items = self.tree.selectedItems()
                        if items:
                            self._apply_question_item_visuals(items[0], q.type)
                        if not silent:
                            self.statusBar().showMessage("Změny otázky uloženy (lokálně).", 1200)
                        return True
                if apply_in(sg.subgroups):
                    return True
            return False
        for g in self.root.groups:
            if apply_in(g.subgroups):
                break

    def _autosave_schedule(self) -> None:
        if not self._current_question_id:
            return
        self._autosave_timer.stop(); self._autosave_timer.start()

    def _autosave_current_question(self) -> None:
        if not self._current_question_id:
            return
        self._apply_editor_to_current_question(silent=True)
        self.save_data()

    def _on_save_question_clicked(self) -> None:
        self._apply_editor_to_current_question(silent=True)
        self.save_data()
        self.statusBar().showMessage("Otázka uložena.", 1500)

    def _update_selected_question_item_title(self, text: str) -> None:
        items = self.tree.selectedItems()
        if items: items[0].setText(0, text or "Otázka")

    def _update_selected_question_item_subtitle(self, text: str) -> None:
        items = self.tree.selectedItems()
        if items: items[0].setText(1, text)

    def _find_group(self, gid: str) -> Optional[Group]:
        for g in self.root.groups:
            if g.id == gid:
                return g
        return None

    def _find_subgroup(self, gid: str, sgid: Optional[str]) -> Optional[Subgroup]:
        g = self._find_group(gid)
        if not g: return None
        def rec(lst: List[Subgroup]) -> Optional[Subgroup]:
            for sg in lst:
                if sg.id == sgid:
                    return sg
                found = rec(sg.subgroups)
                if found: return found
            return None
        return rec(g.subgroups)

    def _find_question(self, gid: str, sgid: Optional[str], qid: str) -> Optional[Question]:
        sg = self._find_subgroup(gid, sgid)
        if not sg: return None
        for q in sg.questions:
            if q.id == qid: return q
        return None

    def _find_question_by_id(self, qid: str) -> Optional[Question]:
        def walk(lst: List[Subgroup]) -> Optional[Question]:
            for sg in lst:
                for q in sg.questions:
                    if q.id == qid:
                        return q
                r = walk(sg.subgroups)
                if r: return r
            return None
        for g in self.root.groups:
            r = walk(g.subgroups)
            if r: return r
        return None

    def _select_question(self, qid: str) -> None:
        def _walk(item: QTreeWidgetItem) -> Optional[QTreeWidgetItem]:
            meta = item.data(0, Qt.UserRole)
            if meta and meta.get("kind") == "question" and meta.get("id") == qid:
                return item
            for i in range(item.childCount()):
                found = _walk(item.child(i))
                if found: return found
            return None
        for i in range(self.tree.topLevelItemCount()):
            found = _walk(self.tree.topLevelItem(i))
            if found:
                self.tree.setCurrentItem(found); self.tree.scrollToItem(found)
                break

    def _merge_format_on_selection(self, fmt: QTextCharFormat) -> None:
        cursor = self.text_edit.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.text_edit.mergeCurrentCharFormat(fmt)

    def _toggle_format(self, which: str) -> None:
        fmt = QTextCharFormat()
        if which == "bold":
            new_weight = QFont.Weight.Bold if self.action_bold.isChecked() else QFont.Weight.Normal
            fmt.setFontWeight(int(new_weight))
        elif which == "italic":
            fmt.setFontItalic(self.action_italic.isChecked())
        elif which == "underline":
            fmt.setFontUnderline(self.action_underline.isChecked())
        self._merge_format_on_selection(fmt); self._autosave_schedule()

    def _choose_color(self) -> None:
        color = QColorDialog.getColor(parent=self, title="Vyberte barvu textu")
        if color.isValid():
            fmt = QTextCharFormat(); fmt.setForeground(color)
            self._merge_format_on_selection(fmt); self._autosave_schedule()

    def _toggle_bullets(self) -> None:
        cursor = self.text_edit.textCursor()
        block = cursor.block()
        in_list = block.textList() is not None
        if in_list:
            lst = block.textList(); fmt = lst.format(); fmt.setStyle(QTextListFormat.ListStyleUndefined); cursor.createList(fmt)
        else:
            fmt = QTextListFormat(); fmt.setStyle(QTextListFormat.ListDisc); cursor.createList(fmt)
        self._autosave_schedule()

    def _apply_alignment(self, align_flag: Qt.AlignmentFlag) -> None:
        cursor = self.text_edit.textCursor()
        bf = QTextBlockFormat(); bf.setAlignment(align_flag)
        if cursor.hasSelection():
            start = cursor.selectionStart(); end = cursor.selectionEnd()
            cursor.beginEditBlock()
            tmp = QTextCursor(self.text_edit.document()); tmp.setPosition(start)
            while tmp.position() <= end:
                block_cursor = QTextCursor(tmp.block())
                block_format = block_cursor.blockFormat()
                block_format.setAlignment(align_flag)
                block_cursor.setBlockFormat(block_format)
                if tmp.block().position() + tmp.block().length() > end:
                    break
                tmp.movePosition(QTextCursor.NextBlock)
            cursor.endEditBlock()
        else:
            cursor.mergeBlockFormat(bf)
        self.text_edit.setTextCursor(cursor)
        self._sync_toolbar_to_cursor()
        self._autosave_schedule()

    def _sync_toolbar_to_cursor(self) -> None:
        fmt = self.text_edit.currentCharFormat()
        try:
            self.action_bold.setChecked(int(fmt.fontWeight()) == int(QFont.Weight.Bold))
        except Exception:
            self.action_bold.setChecked(False)
        self.action_italic.setChecked(fmt.fontItalic())
        self.action_underline.setChecked(fmt.fontUnderline())
        in_list = self.text_edit.textCursor().block().textList() is not None
        self.action_bullets.setChecked(in_list)
        align = self.text_edit.textCursor().blockFormat().alignment()
        self.action_align_left.setChecked(bool(align & Qt.AlignLeft))
        self.action_align_center.setChecked(bool(align & Qt.AlignHCenter))
        self.action_align_right.setChecked(bool(align & Qt.AlignRight))
        self.action_align_justify.setChecked(bool(align & Qt.AlignJustify))

    def _on_type_changed_ui(self) -> None:
        is_classic = self.combo_type.currentIndex() == 0
        self.spin_points.setEnabled(is_classic)
        self.spin_bonus_correct.setEnabled(not is_classic)
        self.spin_bonus_wrong.setEnabled(not is_classic)
        self._autosave_schedule()

    def _choose_data_file(self) -> None:
        new_path, _ = QFileDialog.getSaveFileName(self, "Zvolit/uložit JSON s otázkami", str(self.data_path), "JSON (*.json)")
        if new_path:
            self.data_path = Path(new_path)
            self.statusBar().showMessage(f"Datový soubor změněn na: {self.data_path}", 4000)
            self.load_data(); self._refresh_tree()

    def _apply_filter(self, text: str) -> None:
        pat = (text or '').strip().lower()
        def question_matches(qid: str) -> bool:
            q = None
            for g in self.root.groups:
                stack = list(g.subgroups)
                while stack:
                    sg = stack.pop()
                    for qq in sg.questions:
                        if qq.id == qid:
                            q = qq; break
                    if q: break
                    stack.extend(sg.subgroups)
                if q: break
            if not q: return False
            plain = re.sub(r'<[^>]+>', ' ', q.text_html)
            plain = _html.unescape(plain).lower()
            title = (q.title or '').lower()
            return (pat in title) or (pat in plain)
        def apply_item(item) -> bool:
            meta = item.data(0, Qt.UserRole) or {}
            kind = meta.get('kind')
            any_child = False
            for i in range(item.childCount()):
                if apply_item(item.child(i)):
                    any_child = True
            self_match = False
            if not pat:
                self_match = True
            elif kind in ('group', 'subgroup'):
                self_match = pat in item.text(0).lower()
            elif kind == 'question':
                self_match = question_matches(meta.get('id'))
            show = self_match or any_child
            item.setHidden(not show)
            return show
        for i in range(self.tree.topLevelItemCount()):
            apply_item(self.tree.topLevelItem(i))

    # -------------------- Import z DOCX --------------------

    def _read_numbering_map(self, z: zipfile.ZipFile) -> tuple[dict[int, int], dict[tuple[int, int], str]]:
        num_to_abstract: dict[int, int] = {}
        fmt_map: dict[tuple[int, int], str] = {}
        try:
            with z.open("word/numbering.xml") as f:
                xml = f.read()
        except KeyError:
            return num_to_abstract, fmt_map
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(xml)
        for num in root.findall(".//w:num", ns):
            nid = num.find("w:abstractNumId", ns)
            val = nid.get("{%s}val" % ns["w"]) if nid is not None else None
            num_id_el = num.find("w:numId", ns)
            num_id_val = num_id_el.get("{%s}val" % ns["w"]) if num_id_el is not None else None
            if val is not None and num_id_val is not None:
                try:
                    num_to_abstract[int(num_id_val)] = int(val)
                except Exception:
                    pass
        for absn in root.findall(".//w:abstractNum", ns):
            aid_attr = absn.get("{%s}abstractNumId" % ns["w"])
            if aid_attr is None:
                continue
            try:
                abs_id = int(aid_attr)
            except Exception:
                continue
            for lvl in absn.findall("w:lvl", ns):
                ilvl_attr = lvl.get("{%s}ilvl" % ns["w"])
                if ilvl_attr is None:
                    continue
                try:
                    ilvl = int(ilvl_attr)
                except Exception:
                    ilvl = 0
                numfmt = lvl.find("w:numFmt", ns)
                fmt = numfmt.get("{%s}val" % ns["w"]) if numfmt is not None else "decimal"
                fmt_map[(abs_id, ilvl)] = fmt
        return num_to_abstract, fmt_map

    def _extract_paragraphs_from_docx(self, path: Path) -> List[dict]:
        with zipfile.ZipFile(path, "r") as z:
            with z.open("word/document.xml") as f:
                xml_bytes = f.read()
            num_to_abs, fmt_map = self._read_numbering_map(z)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(xml_bytes)
        out: List[dict] = []
        for p in root.findall(".//w:p", ns):
            ppr = p.find("w:pPr", ns)
            is_num = False
            num_fmt = None
            ilvl = None
            if ppr is not None:
                numpr = ppr.find("w:numPr", ns)
                if numpr is not None:
                    is_num = True
                    num_id_el = numpr.find("w:numId", ns)
                    ilvl_el = numpr.find("w:ilvl", ns)
                    if ilvl_el is not None and ilvl_el.get("{%s}val" % ns["w"]) is not None:
                        try:
                            ilvl = int(ilvl_el.get("{%s}val" % ns["w"]))
                        except Exception:
                            ilvl = 0
                    if num_id_el is not None and num_id_el.get("{%s}val" % ns["w"]) is not None:
                        try:
                            num_id = int(num_id_el.get("{%s}val" % ns["w"]))
                            abs_id = num_to_abs.get(num_id)
                            if abs_id is not None:
                                num_fmt = fmt_map.get((abs_id, ilvl or 0), None)
                        except Exception:
                            pass
            texts = [t.text or "" for t in p.findall(".//w:t", ns)]
            txt = "".join(texts).strip()
            out.append({"text": txt, "is_numbered": is_num, "num_fmt": num_fmt, "ilvl": ilvl})
        return out

    def _parse_questions_from_paragraphs(self, paragraphs: List[dict]) -> List[Question]:
        if paragraphs and isinstance(paragraphs[0], tuple):
            paragraphs = [{'text': t[0], 'is_numbered': bool(t[1]), 'num_fmt': None, 'ilvl': None} for t in paragraphs]
        out: List[Question] = []
        i = 0; n = len(paragraphs)
        rx_scale = re.compile(r'^\s*[A-F]\s*->\s*<[^>]+>\s*bod', re.IGNORECASE)
        rx_bonus_start = re.compile(r'^\s*Otázka\s+\d+', re.IGNORECASE)
        rx_classic_numtxt = re.compile(r'^\s*\d+[\.)]\s')
        rx_question_verb = re.compile(r'\b(Popište|Uveďte|Zašifrujte|Vysvětlete|Porovnejte|Jaký|Jak|Stručně|Lze|Kolik|Která|Co je)\b', re.IGNORECASE)
        def is_noise(text_line: str) -> bool:
            t0 = (text_line or "").strip().lower()
            if not t0: return True
            keys = ['datum:', 'jméno', 'podpis', 'na uvedené otázky', 'maximum bodů', 'klasifikační', 'stupnice', 'souhlasíte', 'cookies']
            if any(k in t0 for k in keys): return True
            if rx_scale.search(t0): return True
            return False
        def is_question_like(t: str) -> bool:
            return ('?' in (t or "")) or bool(rx_question_verb.search(t or ""))
        def html_escape(s: str) -> str:
            return _html.escape(s or "")
        def wrap_list(items: List[tuple[str, int, str]]) -> str:
            if not items: return ""
            fmt = items[0][2] or "decimal"
            if fmt == "bullet":
                tag_open, tag_close = "<ul>", "</ul>"
            elif fmt == "lowerLetter":
                tag_open, tag_close = '<ol type="a">', "</ol>"
            elif fmt == "upperLetter":
                tag_open, tag_close = '<ol type="A">', "</ol>"
            else:
                tag_open, tag_close = "<ol>", "</ol>"
            lis = "".join(f"<li>{html_escape(t)}</li>" for (t, _lvl, _f) in items if t.strip())
            return f"{tag_open}{lis}{tag_close}"
        while i < n:
            para = paragraphs[i]
            txt = para.get("text", "")
            is_num = bool(para.get("is_numbered")); ilvl = para.get("ilvl")
            if is_noise(txt):
                i += 1; continue
            if rx_bonus_start.match(txt) or ("bonus" in (txt or "").lower()):
                block_html = f"<p>{html_escape(txt)}</p>"
                j = i + 1
                while j < n:
                    nt = paragraphs[j].get("text", "")
                    n_isnum = bool(paragraphs[j].get("is_numbered"))
                    if not nt.strip() or is_noise(nt) or n_isnum or rx_bonus_start.match(nt) or rx_classic_numtxt.match(nt) or is_question_like(nt):
                        break
                    if len(nt.strip()) <= 120:
                        block_html += f"<p>{html_escape(nt.strip())}</p>"; j += 1
                    else:
                        break
                q = Question.new_default("bonus")
                q.text_html = block_html; q.title = self._derive_title_from_html(block_html, prefix="BONUS: ")
                q.bonus_correct, q.bonus_wrong = 1.0, 0.0
                out.append(q); i = j; continue
            is_top_numbered = is_num and (ilvl is None or ilvl == 0)
            if is_top_numbered or rx_classic_numtxt.match(txt) or is_question_like(txt):
                block_html = f"<p>{html_escape(txt)}</p>"
                list_buffer: List[tuple[str, int, str]] = []
                j = i + 1
                while j < n:
                    next_txt = paragraphs[j].get("text", "")
                    next_isnum = bool(paragraphs[j].get("is_numbered"))
                    next_ilvl = paragraphs[j].get("ilvl")
                    if not next_txt.strip() or is_noise(next_txt):
                        j += 1; continue
                    if (next_isnum and (next_ilvl is None or next_ilvl == 0)) or rx_bonus_start.match(next_txt) or rx_classic_numtxt.match(next_txt) or is_question_like(next_txt):
                        break
                    if next_isnum:
                        list_buffer.append((next_txt.strip(), next_ilvl or 0, "decimal")); j += 1; continue
                    if len(next_txt.strip()) <= 120:
                        block_html += f"<p>{html_escape(next_txt.strip())}</p>"; j += 1; continue
                    break
                if list_buffer:
                    block_html += wrap_list(list_buffer)
                q = Question.new_default("classic")
                q.text_html = block_html; q.title = self._derive_title_from_html(block_html); q.points = 1
                out.append(q); i = j; continue
            i += 1
        return out

    def _ensure_unassigned_group(self) -> tuple[str, str]:
        name = "Neroztříděné"
        g = next((g for g in self.root.groups if g.name == name), None)
        if not g:
            g = Group(id=str(_uuid.uuid4()), name=name, subgroups=[]); self.root.groups.append(g)
        if not g.subgroups:
            g.subgroups.append(Subgroup(id=str(_uuid.uuid4()), name="Default", subgroups=[], questions=[]))
        return g.id, g.subgroups[0].id

    def _move_question(self) -> None:
        kind, meta = self._selected_node()
        if kind != "question":
            QMessageBox.information(self, "Přesun", "Vyberte nejprve otázku ve stromu."); return
        dlg = MoveTargetDialog(self)
        if dlg.exec() != QDialog.Accepted: return
        g_id, sg_id = dlg.selected_target()
        if not g_id: return
        g = self._find_group(g_id); 
        if not g: return
        target_sg = self._find_subgroup(g_id, sg_id) if sg_id else None
        if not target_sg:
            if not g.subgroups:
                g.subgroups.append(Subgroup(id=str(_uuid.uuid4()), name="Default", subgroups=[], questions=[]))
            target_sg = g.subgroups[0]
        src_gid = meta["parent_group_id"]; src_sgid = meta["parent_subgroup_id"]; qid = meta["id"]
        src_sg = self._find_subgroup(src_gid, src_sgid); q = self._find_question(src_gid, src_sgid, qid)
        if not (src_sg and q): return
        src_sg.questions = [qq for qq in src_sg.questions if qq.id != qid]
        target_sg.questions.append(q)
        self._refresh_tree(); self.save_data()
        g_name = g.name if g else ""; sg_name = target_sg.name if target_sg else "Default"
        self.statusBar().showMessage(f"Otázka přesunuta do {g_name} / {sg_name}.", 4000)

    def _bulk_move_selected(self) -> None:
        items = [it for it in self.tree.selectedItems() if (it.data(0, Qt.UserRole) or {}).get('kind') == 'question']
        if not items:
            QMessageBox.information(self, "Přesun", "Vyberte ve stromu alespoň jednu otázku."); return
        dlg = MoveTargetDialog(self)
        if dlg.exec() != QDialog.Accepted: return
        g_id, sg_id = dlg.selected_target()
        if not g_id: return
        g = self._find_group(g_id); 
        if not g: return
        target_sg = self._find_subgroup(g_id, sg_id) if sg_id else None
        if not target_sg:
            if not g.subgroups:
                g.subgroups.append(Subgroup(id=str(_uuid.uuid4()), name="Default", subgroups=[], questions=[]))
            target_sg = g.subgroups[0]
        moved = 0
        for it in items:
            meta = it.data(0, Qt.UserRole) or {}
            qid = meta.get('id')
            sg = self._find_subgroup(meta.get('parent_group_id'), meta.get('parent_subgroup_id'))
            q = self._find_question(meta.get('parent_group_id'), meta.get('parent_subgroup_id'), qid)
            if sg and q:
                sg.questions = [qq for qq in sg.questions if qq.id != qid]
                target_sg.questions.append(q); moved += 1
        self._refresh_tree(); self.save_data()
        g_name = g.name if g else ""; sg_name = target_sg.name if target_sg else "Default"
        self.statusBar().showMessage(f"Přesunuto {moved} otázek do {g_name} / {sg_name}.", 4000)

    # -------------------- Export DOCX ze šablony --------------------

    def _all_questions_by_type(self, qtype: str) -> List[Question]:
        out: List[Question] = []
        def walk_subs(lst: List[Subgroup]):
            for sg in lst:
                for q in sg.questions:
                    if q.type == qtype: out.append(q)
                walk_subs(sg.subgroups)
        for g in self.root.groups:
            walk_subs(g.subgroups)
        return out

    def _export_docx_wizard(self) -> None:
        wiz = ExportWizard(self); wiz.le_output.setText(str(self.project_root / "test_vystup.docx")); wiz.exec()

    def _generate_docx_from_template(self, template_path: Path, output_path: Path,
                                     simple_repl: Dict[str, str], rich_repl_html: Dict[str, str]) -> None:
        def replace_in_tree(tree: ET.ElementTree) -> None:
            root = tree.getroot()
            for ph_name, html in rich_repl_html.items():
                token1 = f"<{ph_name}>"; token2 = f"&lt;{ph_name}&gt;"
                paragraphs = list(root.findall(".//w:p", NSMAP))
                for p in paragraphs:
                    texts = [t.text or "" for t in p.findall(".//w:t", NSMAP)]
                    full = "".join(texts).strip()
                    if full == token1 or full == token2:
                        paras = parse_html_to_paragraphs(html)
                        new_elements = [ make_w_paragraph(d['align'], d['runs'], d.get('prefix','')) for d in paras ]
                        parent = None
                        def find_parent(r, child):
                            for elem in r.iter():
                                for e in list(elem):
                                    if e is child: return elem
                            return None
                        parent = find_parent(root, p)
                        if parent is not None:
                            idx = list(parent).index(p); parent.remove(p)
                            for off, newp in enumerate(new_elements):
                                parent.insert(idx + off, newp)
            for t in root.findall(".//w:t", NSMAP):
                txt = t.text or ""
                for k, v in simple_repl.items():
                    token1 = f"<{k}>"; token2 = f"&lt;{k}&gt;"
                    if token1 in txt or token2 in txt:
                        txt = txt.replace(token1, v).replace(token2, v)
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = txt
        with zipfile.ZipFile(template_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        tree = ET.ElementTree(ET.fromstring(data.decode("utf-8")))
                        replace_in_tree(tree)
                        data = ET.tostring(tree.getroot(), encoding="utf-8", xml_declaration=False)
                    except Exception:
                        pass
                zout.writestr(item, data)

    # -------------------- Pomocné --------------------

    def _derive_title_from_html(self, html: str, prefix: str = "") -> str:
        import re as _re, html as _h
        txt = _re.sub(r'<[^>]+>', ' ', html or '')
        txt = _h.unescape(txt).strip()
        if not txt: return (prefix + "Otázka").strip()
        parts = _re.split(r'[.!?]\s', txt)
        head = parts[0] if parts and parts[0] else txt
        head = head.strip()
        if len(head) > 80: head = head[:77].rstrip() + '…'
        return (prefix + head).strip()


# --------------------------- main ---------------------------

def main() -> int:
    app = QApplication(sys.argv)
    apply_dark_theme(app)
    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
