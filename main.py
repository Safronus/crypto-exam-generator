#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Crypto Exam Generator (v1.8c)

Novinky v 1.8c
- Přidána aplikace ikona (icon/icon.png). Pokud existuje, nastaví se na QApplication i hlavní okno.
- Oprava chyby: chybějící metoda MainWindow._import_from_docx (menu „Import z DOCX…“).
- Import DOCX: jemné vylepšení – přenáší informaci o typu číslování (decimal / lowerLetter / upperLetter / bullet).

Pozn.: Word numbering (numbering.xml) je mapováno pouze na výsledné vizuální <ol>/<ul> v HTML, bez úprav numbering.xml
(vizualně věrné, minimální zásah).
"""
from __future__ import annotations

import hashlib
import secrets

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

import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
    QIcon, QBrush
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
    # Nové importy pro v4.0 UI
    QGroupBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QTreeWidgetItemIterator,
    QHeaderView, QMenu, QTabWidget
)

APP_NAME = "Crypto Exam Generator"
APP_VERSION = "6.3.1"

# ---------------------------------------------------------------------------
# Globální pomocné funkce
# ---------------------------------------------------------------------------

def parse_html_to_paragraphs(html: str) -> List[dict]:
    if not html:
        return []
    
    parser = HTMLToDocxParser()
    try:
        parser.feed(html)
        parser._end_paragraph() # Flush
    except Exception as e:
        print(f"Chyba při parsování HTML: {e}")
        return [{'align': 'left', 'runs': [{'text': html, 'b': False, 'i': False, 'u': False, 'color': None}], 'prefix': ''}]
    
    # Filtrování prázdných odstavců na začátku a konci (Trim)
    res = parser.paragraphs
    
    # Remove empty start
    while res and not any(r['text'].strip() for r in res[0]['runs']):
        res.pop(0)
    # Remove empty end
    while res and not any(r['text'].strip() for r in res[-1]['runs']):
        res.pop()
        
    if not res:
         # Pokud po promazání nic nezbylo, vrátíme aspoň jeden prázdný (nebo nic, podle logiky)
         # Pro insert_rich chceme raději nic, než prázdný řádek navíc
         return []
         
    return res

# --------------------------- Datové typy ---------------------------

# Přidat do importů (pokud tam není QTableWidget atd., ale v 5.9.2 byly):
# from PySide6.QtWidgets import ..., QTableWidget, QTableWidgetItem, QAbstractItemView

@dataclass
class FunnyAnswer:
    text: str
    author: str
    date: str

@dataclass
class Question:
    id: str
    type: str  # "classic" nebo "bonus"
    text_html: str
    title: str
    points: int = 1
    bonus_correct: float = 0.0
    bonus_wrong: float = 0.0
    created_at: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    # Nová pole
    correct_answer: str = ""
    funny_answers: List[FunnyAnswer] = field(default_factory=list)

    @staticmethod
    def new_default(q_type: str = "classic") -> "Question":
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if q_type == "bonus":
            return Question(
                id=str(_uuid.uuid4()),
                type="bonus",
                text_html="<p>Znění bonusové otázky...</p>",
                title="BONUS otázka",
                points=0,
                bonus_correct=1.0,
                bonus_wrong=0.0,
                created_at=now,
                correct_answer="",
                funny_answers=[]
            )
        return Question(
            id=str(_uuid.uuid4()),
            type="classic",
            text_html="<p>Znění otázky...</p>",
            title="Otázka",
            points=1,
            bonus_correct=0.0,
            bonus_wrong=0.0,
            created_at=now,
            correct_answer="",
            funny_answers=[]
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
        
        # Nastavení chování hlavičky pro správné roztažení
        header = self.header()
        # 0. sloupec (Název) se roztáhne do zbytku
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        # 1. sloupec (Typ/body) se přizpůsobí obsahu
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        
        # Důležité: StretchLastSection musí být False, jinak přebije naše nastavení
        header.setStretchLastSection(False)

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

class HTMLToDocxParser(HTMLParser):
    """
    Převádí HTML na seznam odstavců pro DOCX.
    Podporuje vnořené styly a ignoruje balast.
    """
    def __init__(self) -> None:
        super().__init__()
        self.paragraphs: List[dict] = []
        
        # Zásobník otevřených elementů a jejich stylů
        # Každá položka: {'tag': str, 'attrs': dict, 'styles': dict}
        self._stack: List[dict] = [] 
        
        self._list_stack: List[dict] = [] # Pro seznamy
        self._current_runs: List[dict] = []
        self._current_align: str = "left"
        
        # Globální stav ignorování
        self._in_body = False
        self._has_body_tag = False
        self._ignore_content = False

    # -- STAVOVÉ METODY PRO PARAGRAFY --

    def _start_paragraph(self, prefix: str = "") -> None:
        if self._current_runs:
            self._end_paragraph()
        
        self._current_runs = []
        self._current_align = "left"
        
        # Hledáme zarovnání v aktuálním kontextu (poslední blokový element na zásobníku)
        # Procházíme od konce zásobníku, první p/div/li určí zarovnání
        for item in reversed(self._stack):
            if item['tag'] in ('p', 'div', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
                if item['styles'].get('align'):
                    self._current_align = item['styles']['align']
                break # Našli jsme blok, končíme hledání (zarovnání se nedědí z nadřazeného bloku v docx tak snadno)
                
        self._current_prefix = prefix


    def _end_paragraph(self) -> None:
        # Sloučíme runy a uložíme
        if not self._current_runs and not hasattr(self, '_current_prefix'):
            return

        merged = []
        for r in self._current_runs:
            if r['text'] == "": continue
            # Merge
            if merged and all(merged[-1][k] == r[k] for k in ('b','i','u','color')):
                merged[-1]['text'] += r['text']
            else:
                merged.append(r)
        
        prefix = getattr(self, '_current_prefix', '')
        
        # Uložíme (i prázdný, pokud má prefix - např. prázdná odrážka)
        if merged or prefix:
            self.paragraphs.append({
                'align': self._current_align,
                'runs': merged,
                'prefix': prefix
            })
        
        self._current_runs = []
        if hasattr(self, '_current_prefix'): del self._current_prefix

    def _append_text(self, text: str):
        # Zjistíme aktuální styl ze zásobníku
        b = any(item['styles'].get('b') for item in self._stack)
        i = any(item['styles'].get('i') for item in self._stack)
        u = any(item['styles'].get('u') for item in self._stack)
        
        # Barva: poslední platná na zásobníku
        color = None
        for item in reversed(self._stack):
            if item['styles'].get('color'):
                color = item['styles']['color']
                break
        
        self._current_runs.append({
            'text': text,
            'b': b, 'i': i, 'u': u, 'color': color
        })

    # -- PARSOVACÍ LOGIKA --

    def feed(self, data: str) -> None:
        if "<body" in data.lower():
            self._has_body_tag = True
            self._in_body = False
        else:
            self._has_body_tag = False
            self._in_body = True
        super().feed(data)

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        attrs_d = dict(attrs)
        
        if tag == 'body':
            self._in_body = True
            return
        if tag in ('head', 'style', 'script', 'meta', 'title', 'html', '!doctype'):
            self._ignore_content = True
            return
        if not self._in_body or self._ignore_content:
            return

        # -- ANALÝZA STYLŮ (CSS) --
        styles = {}
        
        # 1. Z atributu style="..."
        if 'style' in attrs_d:
            s = attrs_d['style'].lower()
            
            # Color
            m = re.search(r'color\s*:\s*#?([0-9a-f]{6})', s)
            if m: styles['color'] = m.group(1)
            else:
                m2 = re.search(r'color\s*:\s*rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', s)
                if m2:
                    r,g,b = [max(0, min(255, int(x))) for x in m2.groups()]
                    styles['color'] = f"{r:02X}{g:02X}{b:02X}"
            
            # Align
            m_align = re.search(r'text-align\s*:\s*(left|center|right|justify)', s)
            if m_align: styles['align'] = m_align.group(1)
            
            # Bold (font-weight)
            # Qt používá font-weight:600, 700, 800... nebo bold
            if 'font-weight' in s:
                if 'bold' in s:
                    styles['b'] = True
                else:
                    m_fw = re.search(r'font-weight\s*:\s*(\d+)', s)
                    if m_fw and int(m_fw.group(1)) >= 600:
                        styles['b'] = True
            
            # Italic (font-style)
            if 'font-style' in s and 'italic' in s:
                styles['i'] = True
            
            # Underline (text-decoration)
            if 'text-decoration' in s and 'underline' in s:
                styles['u'] = True
        
        # 2. Z atributů tagu (staré HTML)
        if attrs_d.get('align'): styles['align'] = attrs_d['align'].lower()
        if tag in ('b', 'strong'): styles['b'] = True
        if tag in ('i', 'em'): styles['i'] = True
        if tag == 'u': styles['u'] = True
        
        # Přidáme na zásobník
        self._stack.append({'tag': tag, 'attrs': attrs_d, 'styles': styles})

        # -- ZPRACOVÁNÍ TAGŮ --
        
        if tag in ('p', 'div'):
            self._start_paragraph()
        
        elif tag == 'br':
            self._append_text("\n")
            
        elif tag in ('ul', 'ol'):
            t = attrs_d.get('type', '1')
            self._list_stack.append({'tag': tag, 'type': t, 'count': 0})
            
        elif tag == 'li':
            prefix = self._get_list_prefix()
            self._start_paragraph(prefix=prefix)


    def handle_endtag(self, tag):
        tag = tag.lower()
        
        if tag == 'body':
            self._in_body = False; return
        if tag in ('head', 'style', 'script', 'meta', 'title', 'html'):
            self._ignore_content = False; return
        if not self._in_body or self._ignore_content:
            return

        # Strukturální ukončení
        if tag in ('p', 'div', 'li'):
            self._end_paragraph()
        
        if tag in ('ul', 'ol'):
            if self._list_stack: self._list_stack.pop()

        # Odstranění ze zásobníku (hledáme od konce odpovídající tag)
        # Protože HTML může být validně neuzavřené (např. <p>...<p>), 
        # musíme být opatrní. Ale v QT rich textu je to obvykle well-formed.
        for i in range(len(self._stack) - 1, -1, -1):
            if self._stack[i]['tag'] == tag:
                # Odebereme tento a všechny nad ním (pokud byly neuzavřené)
                del self._stack[i:]
                break

    def handle_data(self, data):
        if not self._in_body or self._ignore_content: return
        if not data: return
        
        # Normalizace pevných mezer, pokud je třeba (volitelné)
        # data = data.replace('\xa0', ' ') 
        
        in_block = False
        for item in reversed(self._stack):
            if item['tag'] in ('p', 'div', 'li'):
                in_block = True
                break
        
        if not in_block:
            if not data.strip(): return
            self._start_paragraph()
            
        self._append_text(data)


    # Helpers (stejné jako minule)
    def _get_list_prefix(self) -> str:
        if not self._list_stack: return ""
        L = self._list_stack[-1]
        if L['tag'] == 'ul': return "•\t"
        elif L['tag'] == 'ol':
            L['count'] += 1
            n = L['count']
            t = L.get('type', '1')
            val = f"{n}."
            if t == 'a': val = f"{self._to_alpha(n, False)}."
            elif t == 'A': val = f"{self._to_alpha(n, True)}."
            elif t == 'i': val = f"{self._to_roman(n, False)}."
            elif t == 'I': val = f"{self._to_roman(n, True)}."
            return f"{val}\t"
        return ""
        
    def _to_alpha(self, n: int, upper: bool) -> str:
        s = ""
        n -= 1
        while n >= 0:
            s = chr(97 + n % 26) + s
            n = n // 26 - 1
        return s.upper() if upper else s
        
    def _to_roman(self, n: int, upper: bool) -> str:
        val = [(50, 'L'), (40, 'XL'), (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
        res = ""
        for v, r in val:
            while n >= v: res += r; n -= v
        return res if upper else res.lower()


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


# --------------------------- Export Wizard ---------------------------

class ExportWizard(QWizard):
    def __init__(self, owner: "MainWindow") -> None:
        super().__init__(owner)
        self.setWindowTitle("Export DOCX – Průvodce")
        self.setWizardStyle(QWizard.ModernStyle)
        self.owner = owner
        self.resize(1600, 1200)

        # Lokalizace
        self.setButtonText(QWizard.BackButton, "< Zpět")
        self.setButtonText(QWizard.NextButton, "Další >")
        self.setButtonText(QWizard.CommitButton, "Dokončit")
        self.setButtonText(QWizard.FinishButton, "Dokončit")
        self.setButtonText(QWizard.CancelButton, "Zrušit")

        # IDs
        self.ID_PAGE1 = 0
        self.ID_PAGE2 = 1
        self.ID_PAGE3 = 2

        # Data
        self.template_path: Optional[Path] = None
        self.output_path: Optional[Path] = None
        self.output_changed_manually = False
        self.placeholders_q = []
        self.placeholders_b = []
        self.selection_map = {}

        # Cesty (Default)
        self.templates_dir = self.owner.project_root / "data" / "Šablony"
        self.output_dir = self.owner.project_root / "data" / "Vygenerované testy"
        self.templates_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.default_template = self.templates_dir / "template_AK3KR_předtermín.docx"
        # Inicializace default_output (aby nepadal dialog)
        self.default_output = self.output_dir / "export.docx" 

        # --- INIT STRÁNEK ---
        self.page1 = QWizardPage()
        self.page2 = QWizardPage()
        self.page3 = QWizardPage()

        self._build_page1_content()
        self._build_page2_content()
        self._build_page3_content()

        self.setPage(self.ID_PAGE1, self.page1)
        self.setPage(self.ID_PAGE2, self.page2)
        self.setPage(self.ID_PAGE3, self.page3)
        
        self.setStartId(self.ID_PAGE1)

        # Auto-load Šablona
        if self.default_template.exists():
            self.le_template.setText(str(self.default_template))
            self.template_path = self.default_template
            QTimer.singleShot(100, self._scan_placeholders)
        
        # Auto-generate Output Name (Hned teď!)
        # Toto nastaví self.output_path a text v QLineEdit
        self._update_default_output()


    # --- Build Content Methods ---

    def _build_page1_content(self):
        self.page1.setTitle("Krok 1: Výběr šablony a globální nastavení")
        l1 = QVBoxLayout(self.page1)
        
        # GroupBox: Parametry
        gb_params = QGroupBox("Parametry testu")
        form_params = QFormLayout()
        
        self.le_prefix = QLineEdit("MůjTest")
        self.le_prefix.textChanged.connect(self._update_default_output)
        form_params.addRow("Prefix verze:", self.le_prefix)
        
        self.dt_edit = QDateTimeEdit(QDateTime.currentDateTime())
        self.dt_edit.setDisplayFormat("dd.MM.yyyy HH:mm")
        self.dt_edit.setCalendarPopup(True)
        self.dt_edit.dateTimeChanged.connect(self._update_default_output)
        form_params.addRow("Datum testu:", self.dt_edit)
        
        gb_params.setLayout(form_params)
        l1.addWidget(gb_params)

        # GroupBox: Soubory
        gb_files = QGroupBox("Soubory")
        form_files = QFormLayout()
        
        self.le_template = QLineEdit()
        self.le_template.textChanged.connect(self._on_templ_change)
        btn_t = QPushButton("Vybrat šablonu...")
        btn_t.clicked.connect(self._choose_template)
        h_t = QHBoxLayout(); h_t.addWidget(self.le_template); h_t.addWidget(btn_t)
        form_files.addRow("Šablona:", h_t)
        
        self.le_output = QLineEdit()
        self.le_output.textChanged.connect(self._on_output_text_changed)
        btn_o = QPushButton("Cíl exportu...")
        btn_o.clicked.connect(self._choose_output)
        h_o = QHBoxLayout(); h_o.addWidget(self.le_output); h_o.addWidget(btn_o)
        form_files.addRow("Výstup:", h_o)
        
        gb_files.setLayout(form_files)
        l1.addWidget(gb_files)
        
        self.lbl_scan_info = QLabel("Info: Čekám na načtení šablony...")
        self.lbl_scan_info.setStyleSheet("color: gray; font-style: italic;")
        l1.addWidget(self.lbl_scan_info)

    def _build_page2_content(self):
        self.page2.setTitle("Krok 2: Přiřazení otázek do šablony")
        self.page2.initializePage = self._init_page2
        
        main_layout = QVBoxLayout(self.page2)
        
        # 1. Info Panel
        self.info_box_p2 = QGroupBox("Kontext exportu")
        self.info_box_p2.setStyleSheet("QGroupBox { font-weight: bold; border: 1px solid #555; margin-top: 6px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; }")
        l_info = QFormLayout(self.info_box_p2)
        self.lbl_templ_p2 = QLabel("-")
        self.lbl_out_p2 = QLabel("-")
        l_info.addRow("Vstupní šablona:", self.lbl_templ_p2)
        l_info.addRow("Výstupní soubor:", self.lbl_out_p2)
        main_layout.addWidget(self.info_box_p2)

        # 2. Hlavní obsah (Dva sloupce: Strom | Sloty)
        columns_layout = QHBoxLayout()
        
        # Levý panel: Strom
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("<b>Dostupné otázky:</b>"))
        self.tree_source = QTreeWidget()
        self.tree_source.setHeaderLabels(["Struktura otázek"])
        self.tree_source.setSelectionMode(QAbstractItemView.ExtendedSelection)
        
        # Nastavení signálů
        self.tree_source.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_source.customContextMenuRequested.connect(self._on_tree_source_context_menu)
        
        # OPRAVA: Vrácení signálu pro výběr (náhled a single selection logic)
        # Pokud metoda _on_tree_selection v původním kódu existuje (což asi ano), musíme ji zapojit.
        if hasattr(self, "_on_tree_selection"):
            self.tree_source.itemSelectionChanged.connect(self._on_tree_selection)
        
        left_layout.addWidget(self.tree_source)
        
        self.btn_assign_multi = QPushButton(">> Přiřadit vybrané na volné pozice >>")
        self.btn_assign_multi.setToolTip("Doplní vybrané otázky (zleva) na první volná místa v šabloně (vpravo).")
        self.btn_assign_multi.clicked.connect(self._assign_selected_multi)
        left_layout.addWidget(self.btn_assign_multi)
        
        columns_layout.addLayout(left_layout, 4)
        
        # Pravý panel: Sloty
        right_layout = QVBoxLayout()
        
        right_header = QHBoxLayout()
        right_header.addWidget(QLabel("<b>Sloty v šabloně:</b>"))
        right_header.addStretch()
        
        self.btn_clear_all = QPushButton("Vyprázdnit vše")
        self.btn_clear_all.setToolTip("Zruší přiřazení všech otázek.")
        self.btn_clear_all.clicked.connect(self._clear_all_assignments)
        right_header.addWidget(self.btn_clear_all)
        
        right_layout.addLayout(right_header)
        
        self.scroll_slots = QScrollArea()
        self.scroll_slots.setWidgetResizable(True)
        self.widget_slots = QWidget()
        self.layout_slots = QVBoxLayout(self.widget_slots)
        self.layout_slots.setSpacing(6)
        self.layout_slots.addStretch()
        self.scroll_slots.setWidget(self.widget_slots)
        right_layout.addWidget(self.scroll_slots)
        columns_layout.addLayout(right_layout, 6)
        
        main_layout.addLayout(columns_layout, 3)

        # 3. Náhled
        preview_box = QGroupBox("Náhled vybrané otázky")
        preview_layout = QVBoxLayout(preview_box)
        preview_layout.setContentsMargins(5,5,5,5)
        
        self.text_preview_q = QTextEdit()
        self.text_preview_q.setReadOnly(True)
        self.text_preview_q.setMaximumHeight(120)
        self.text_preview_q.setStyleSheet("""
            QTextEdit { 
                background-color: #2e2e2e; 
                color: #ffffff; 
                font-size: 14px; 
                border: 1px solid #555;
                padding: 5px;
            }
        """)
        preview_layout.addWidget(self.text_preview_q)
        
        main_layout.addWidget(preview_box, 1)


    def _clear_all_assignments(self) -> None:
        """Vymaže všechna přiřazení otázek ve slotech."""
        if not self.selection_map:
            return

        if QMessageBox.question(self, "Vyprázdnit", "Opravdu zrušit všechna přiřazení?") != QMessageBox.Yes:
            return

        self.selection_map.clear()
        self._init_page2() # Obnoví UI
        self.owner.statusBar().showMessage("Všechna přiřazení byla zrušena.", 3000)


    def _build_page3_content(self):
        self.page3.setTitle("Krok 3: Kontrola a Export")
        self.page3.initializePage = self._init_page3
        
        main_layout = QVBoxLayout(self.page3)
        
        # Info Panel (PŮVODNÍ KÓD OBNOVEN)
        self.info_box_p3 = QGroupBox("Kontext exportu")
        self.info_box_p3.setStyleSheet("QGroupBox { font-weight: bold; border: 1px solid #555; margin-top: 6px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; }")
        l_info = QFormLayout(self.info_box_p3)
        self.lbl_templ_p3 = QLabel("-")
        self.lbl_out_p3 = QLabel("-")
        l_info.addRow("Vstupní šablona:", self.lbl_templ_p3)
        l_info.addRow("Výstupní soubor:", self.lbl_out_p3)
        main_layout.addWidget(self.info_box_p3)
        
        # Náhled
        lbl_prev = QLabel("<b>Náhled obsahu testu:</b>")
        main_layout.addWidget(lbl_prev)
        
        self.preview_edit = QTextEdit()
        self.preview_edit.setReadOnly(True)
        # Dark Theme CSS pro QTextEdit
        self.preview_edit.setStyleSheet("QTextEdit { background-color: #252526; color: #e0e0e0; border: 1px solid #3e3e42; }")
        main_layout.addWidget(self.preview_edit)

        # NOVÉ: Label pro kontrolní hash (PŘIDÁNO NA KONEC)
        self.lbl_hash_preview = QLabel("Hash: -")
        self.lbl_hash_preview.setWordWrap(True)
        self.lbl_hash_preview.setStyleSheet("color: #555; font-family: Monospace; font-size: 10px; margin-top: 5px;")
        self.lbl_hash_preview.setTextInteractionFlags(Qt.TextSelectableByMouse)
        main_layout.addWidget(self.lbl_hash_preview)
    # --- Helpers & Logic ---

    def _update_default_output(self):
        if self.output_changed_manually:
            return
        prefix = self.le_prefix.text().strip() or "Test"
        dt = self.dt_edit.dateTime()
        filename = f"{prefix}_{dt.toString('yyyy-MM-dd_HHmm')}.docx"
        new_path = self.output_dir / filename
        
        self.le_output.blockSignals(True)
        self.le_output.setText(str(new_path))
        self.le_output.blockSignals(False)
        self.output_path = new_path

    def _on_output_text_changed(self, text):
        self.output_changed_manually = True
        self.output_path = Path(text)

    def _choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Vybrat šablonu", str(self.templates_dir), "*.docx")
        if path:
            self.le_template.setText(path)

    def _choose_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Cíl exportu", str(self.default_output), "*.docx")
        if path:
            self.le_output.setText(path)

    def _on_templ_change(self, text):
        path = Path(text)
        if path.exists() and path.suffix == '.docx':
            self.template_path = path
            self._scan_placeholders()
        else:
            self.template_path = None
            self.lbl_scan_info.setText("Šablona neexistuje nebo není platná.")

    def _scan_placeholders(self):
        if not self.template_path or not self.template_path.exists():
            return

        try:
            doc = docx.Document(self.template_path)
            full_text = ""
            # Načteme text z celého dokumentu pro regex
            for p in doc.paragraphs: full_text += p.text + "\n"
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        for p in c.paragraphs: full_text += p.text + "\n"
            
            import re
            # Regex pro <Placeholder> i {Placeholder}
            placeholders = re.findall(r'[<{]([A-Za-z0-9ÁČĎÉĚÍŇÓŘŠŤÚŮÝŽáčďéěíňóřšťúůýž]+)[>}]', full_text)
            placeholders = sorted(list(set(placeholders)))
            
            self.placeholders_q = [p for p in placeholders if re.match(r'^Otázka\d+$', p)]
            self.placeholders_q.sort(key=lambda x: int(re.findall(r'\d+', x)[0]))
            
            self.placeholders_b = [p for p in placeholders if re.match(r'^BONUS\d+$', p)]
            self.placeholders_b.sort(key=lambda x: int(re.findall(r'\d+', x)[0]))
            
            self.has_datumcas = any(x in placeholders for x in ['DatumČas', 'DatumCas', 'DATUMCAS'])
            self.has_pozn = any(x in placeholders for x in ['PoznamkaVerze', 'POZNAMKAVERZE'])
            self.has_minmax = (any('MinBody' in x for x in placeholders), any('MaxBody' in x for x in placeholders))
            
            msg = f"Nalezeno: {len(self.placeholders_q)}x Otázka, {len(self.placeholders_b)}x BONUS."
            if self.has_minmax[0]: msg += " (S body)."
            self.lbl_scan_info.setText(msg)
            
        except Exception as e:
            self.lbl_scan_info.setText(f"Chyba čtení šablony: {e}")

    # --- Page Initializers ---
    
    def _on_tree_selection(self):
        sel = self.tree_source.selectedItems()
        if not sel:
            self.text_preview_q.clear()
            return
            
        item = sel[0]
        qid = item.data(0, Qt.UserRole)
        
        if not qid:
            self.text_preview_q.setText("--- (Vyberte konkrétní otázku pro náhled) ---")
            return
            
        q = self.owner._find_question_by_id(qid)
        if q:
            import re
            import html
            
            html_content = q.text_html or ""
            
            # 1. Odstranění <style>...</style> a <head>...</head> i s obsahem
            # Flag re.DOTALL zajistí, že . matchuje i newlines
            clean = re.sub(r'<style.*?>.*?</style>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
            clean = re.sub(r'<head.*?>.*?</head>', '', clean, flags=re.DOTALL | re.IGNORECASE)
            
            # 2. Odstranění všech ostatních tagů <...>
            clean = re.sub(r'<[^>]+>', ' ', clean)
            
            # 3. Decode entities (&nbsp;, &lt;...)
            clean = html.unescape(clean)
            
            # 4. Squeeze whitespace (více mezer/tabulátorů na jednu mezeru, ale zachovat newlines pokud chceme, 
            # nebo vše na jeden řádek. Pro náhled je asi lepší zachovat základní odstavce, 
            # ale QTextEdit plain text to zvládne)
            
            # Zkusíme odstranit vícenásobné prázdné řádky
            lines = [line.strip() for line in clean.splitlines() if line.strip()]
            final_text = "\n".join(lines)
            
            self.text_preview_q.setText(final_text)
        else:
            self.text_preview_q.clear()

    def _init_page2(self):
        try:
            # Info update
            t_name = self.template_path.name if self.template_path else "Nevybráno"
            o_name = self.output_path.name if self.output_path else "Nevybráno"
            self.lbl_templ_p2.setText(t_name)
            self.lbl_out_p2.setText(o_name)

            # Rescan
            if not self.placeholders_q and not self.placeholders_b:
                self._scan_placeholders()

            # 1. Clear Tree
            self.tree_source.blockSignals(True)
            self.tree_source.clear()
            
            # 2. Clear Slots
            while self.layout_slots.count():
                item = self.layout_slots.takeAt(0)
                if item.widget(): item.widget().deleteLater()
            self.layout_slots.addStretch()
            
            # 4. Populate Tree
            def add_subgroup_recursive(parent_item, subgroup_list, parent_gid):
                for sg in subgroup_list:
                    sg_item = QTreeWidgetItem([sg.name])
                    sg_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
                    sg_item.setData(0, Qt.UserRole, {
                        "kind": "subgroup", 
                        "id": sg.id, 
                        "parent_group_id": parent_gid
                    })
                    parent_item.addChild(sg_item)
                    
                    for q in sg.questions:
                        info = f"({q.points} b)" if q.type == 'classic' else f"(Bonus: {q.bonus_correct})"
                        label = f"{q.title} {info}"
                        
                        q_item = QTreeWidgetItem([label])
                        q_item.setData(0, Qt.UserRole, {
                            "kind": "question",
                            "id": q.id,
                            "parent_group_id": parent_gid,
                            "parent_subgroup_id": sg.id
                        })
                        q_item.setIcon(0, self.style().standardIcon(QStyle.SP_FileIcon))
                        
                        sg_item.addChild(q_item)
                    
                    if sg.subgroups:
                        add_subgroup_recursive(sg_item, sg.subgroups, parent_gid)

            groups = self.owner.root.groups
            for g in groups:
                g_item = QTreeWidgetItem([g.name])
                g_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
                f = g_item.font(0); f.setBold(True); g_item.setFont(0, f)
                g_item.setData(0, Qt.UserRole, {
                    "kind": "group", 
                    "id": g.id
                })
                self.tree_source.addTopLevelItem(g_item)
                add_subgroup_recursive(g_item, g.subgroups, g.id)
                
            self.tree_source.expandAll()
            self.tree_source.blockSignals(False)

            # 5. Create Slots
            def create_slot_row(ph, is_bonus):
                row_w = QWidget()
                row_l = QHBoxLayout(row_w)
                row_l.setContentsMargins(0,0,0,0)
                
                lbl_name = QLabel(f"{ph}:")
                lbl_name.setFixedWidth(120)
                if is_bonus: lbl_name.setStyleSheet("color: #ffcc00;")
                else: lbl_name.setStyleSheet("color: #4da6ff;")
                
                btn_assign = QPushButton("Vybrat...")
                qid = self.selection_map.get(ph)
                if qid:
                    q = self.owner._find_question_by_id(qid)
                    if q:
                        btn_assign.setText(q.title)
                        btn_assign.setToolTip(f"{q.text_html[:100]}...")
                    else:
                        btn_assign.setText("???")
                else:
                    btn_assign.setText("--- Volné ---")
                    
                btn_assign.clicked.connect(lambda checked, p=ph: self._on_slot_assign_clicked(p))
                
                btn_clear = QPushButton("X")
                btn_clear.setFixedWidth(30)
                btn_clear.clicked.connect(lambda checked, p=ph: self._on_slot_clear_clicked(p))
                
                row_l.addWidget(lbl_name)
                row_l.addWidget(btn_assign, 1)
                row_l.addWidget(btn_clear)
                row_w.setProperty("placeholder", ph)
                self.layout_slots.insertWidget(self.layout_slots.count()-1, row_w)

            if self.placeholders_q:
                lbl = QLabel("--- KLASICKÉ OTÁZKY ---")
                lbl.setStyleSheet("font-weight:bold; color:#4da6ff; margin-top:5px;")
                self.layout_slots.insertWidget(self.layout_slots.count()-1, lbl)
                for ph in self.placeholders_q:
                    create_slot_row(ph, False)

            if self.placeholders_b:
                lbl = QLabel("--- BONUSOVÉ OTÁZKY ---")
                lbl.setStyleSheet("font-weight:bold; color:#ffcc00; margin-top:10px;")
                self.layout_slots.insertWidget(self.layout_slots.count()-1, lbl)
                for ph in self.placeholders_b:
                    create_slot_row(ph, True)
            
            # DŮLEŽITÉ: Aktualizace vizuálů stromu na konci
            self._refresh_tree_visuals()
                    
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Chyba", f"Chyba při inicializaci stránky 2:\n{e}")


    def _on_slot_assign_clicked(self, ph: str) -> None:
        # Jednoduchý výběr: Otevře dialog se seznamem dostupných otázek
        # Pro jednoduchost: jen zrušíme výběr, aby se uvolnilo místo? Ne, má to přiřadit.
        # V rámci "Minimal intervention" bych neměl přidávat komplexní dialogy, pokud tam nebyly.
        # Ale tlačítko tam je.
        # Pokud tam metoda byla, je to ok. Pokud ne, přidám basic logiku.
        # Zkusíme předpokládat, že tam nebyla a přidáme ji.
        
        # Dialog pro výběr konkrétní otázky do slotu
        # (Může využít existující výběr ve stromu nebo nový dialog)
        
        # Pro teď: Pokud je vybrána otázka ve stromu, přiřadí ji.
        items = self.tree_source.selectedItems()
        if not items:
            QMessageBox.information(self, "Výběr", "Vyberte otázku ve stromu vlevo a pak klikněte sem.")
            return
            
        meta = items[0].data(0, Qt.UserRole) or {}
        if meta.get("kind") != "question":
            return
            
        q = self.owner._find_question(meta["parent_group_id"], meta["parent_subgroup_id"], meta["id"])
        if q:
            self._assign_question_to_slot(ph, q)

    def _on_slot_clear_clicked(self, ph: str) -> None:
        if ph in self.selection_map:
            del self.selection_map[ph]
            self._init_page2() # Refresh

    def _on_tree_source_context_menu(self, pos) -> None:
        """Kontextové menu nad stromem zdrojových otázek (Krok 2)."""
        # Zjistíme, kolik je vybraných položek
        items = self.tree_source.selectedItems()
        if not items:
            return
            
        menu = QMenu(self)
        has_action = False

        # PŘÍPAD 1: Multi-select (více než 1 položka)
        if len(items) > 1:
            # Nabídneme hromadné přiřazení
            act_multi = QAction(f"Přiřadit vybrané ({len(items)}) na volné pozice", self)
            act_multi.triggered.connect(self._assign_selected_multi)
            menu.addAction(act_multi)
            has_action = True
            
        # PŘÍPAD 2: Single-select (1 položka)
        elif len(items) == 1:
            item = items[0]
            meta = item.data(0, Qt.UserRole) or {}
            kind = meta.get("kind")
            
            # Skupina/Podskupina -> Náhodný výběr
            if kind in ("group", "subgroup"):
                act_random = QAction("Naplnit volné pozice náhodně z této větve", self)
                act_random.triggered.connect(lambda: self._assign_random_from_context(meta))
                menu.addAction(act_random)
                has_action = True
                
            # Otázka -> Single assign (pokud chceme i v kontext menu, nebo necháme jen dblclick/tlačítko)
            # Požadavek byl "může to fungovat ještě pro multi-select", o single v menu explicitně nepadlo slovo,
            # ale v minulé verzi jsem to přidal. Nechám to tam.
            if kind == "question":
                act_assign = QAction("Přiřadit na první volné místo", self)
                act_assign.triggered.connect(lambda: self._assign_single_question_from_context(meta))
                menu.addAction(act_assign)
                has_action = True

        if has_action:
            menu.exec(self.tree_source.mapToGlobal(pos))

    def _refresh_tree_visuals(self) -> None:
        """Aktualizuje vizuální stav položek ve stromu (zvýrazní vybrané)."""
        iterator = QTreeWidgetItemIterator(self.tree_source)
        used_ids = set(self.selection_map.values())
        
        # Barvy pro dark theme
        color_used = QColor("#666666")
        color_normal = QColor("#e0e0e0")
        
        while iterator.value():
            item = iterator.value()
            meta = item.data(0, Qt.UserRole) or {}
            
            if meta.get("kind") == "question":
                qid = meta.get("id")
                txt = item.text(0)
                
                # Odstraníme případný starý suffix
                clean_txt = txt.replace(" [VYBRÁNO]", "")
                
                if qid in used_ids:
                    # Je vybrána
                    item.setText(0, clean_txt + " [VYBRÁNO]")
                    item.setForeground(0, QBrush(color_used))
                    f = item.font(0); f.setItalic(True); item.setFont(0, f)
                else:
                    # Není vybrána
                    item.setText(0, clean_txt)
                    item.setForeground(0, QBrush(color_normal))
                    f = item.font(0); f.setItalic(False); item.setFont(0, f)
            
            iterator += 1

    def _assign_single_question_from_context(self, meta: dict) -> None:
        """Přiřadí jednu otázku z kontextového menu."""
        q = self.owner._find_question(meta["parent_group_id"], meta["parent_subgroup_id"], meta["id"])
        if not q:
            return
            
        # Využijeme logiku z multi-selectu (je to jako select 1 item)
        # Ale pro jednoduchost přímo:
        if q.id in self.selection_map.values():
            self.owner.statusBar().showMessage("Otázka již je přiřazena.", 2000)
            return
            
        target_ph = None
        if q.type == "classic":
            for ph in self.placeholders_q:
                if ph not in self.selection_map:
                    target_ph = ph; break
        elif q.type == "bonus":
            for ph in self.placeholders_b:
                if ph not in self.selection_map:
                    target_ph = ph; break
                    
        if target_ph:
            self._assign_question_to_slot(target_ph, q)
            self.owner.statusBar().showMessage("Otázka přiřazena.", 2000)
        else:
            QMessageBox.information(self, "Plno", "Není volné místo pro tento typ otázky.")


    def _assign_random_from_context(self, meta: dict) -> None:
        """Vybere náhodné otázky z dané větve a doplní je na volná místa."""
        # 1. Zjistit volné sloty
        free_slots = []
        # Musíme iterovat přes layout slotů, ale nemáme přímý list.
        # Můžeme projít placeholders_q a placeholders_b a zkontrolovat selection_map.
        
        # Spojíme seznamy placeholderů (nejprve klasické, pak bonusové, nebo jak chceme plnit)
        # Obvykle chceme plnit klasické klasickými a bonusové bonusovými? 
        # Zadání specifikuje "otázky", ale v aplikaci je rozdělení na typy.
        # Pro jednoduchost a konzervativnost: 
        # Pokud je slot pro klasickou otázku (v placeholders_q), hledáme klasické otázky.
        # Pokud je slot pro bonus (v placeholders_b), hledáme bonusové.
        
        # Sběr všech otázek ve větvi
        all_questions_in_branch = []
        
        def collect_recursive(gid, sgid):
            if sgid:
                sg = self.owner._find_subgroup(gid, sgid)
                if sg:
                    all_questions_in_branch.extend(sg.questions)
                    for sub in sg.subgroups:
                        collect_recursive(gid, sub.id)
            else:
                g = self.owner._find_group(gid)
                if g:
                    for sub in g.subgroups:
                        collect_recursive(gid, sub.id)

        gid = meta.get("id") if meta.get("kind") == "group" else meta.get("parent_group_id")
        sgid = None if meta.get("kind") == "group" else meta.get("id")
        
        collect_recursive(gid, sgid)
        
        if not all_questions_in_branch:
            QMessageBox.information(self, "Info", "V této větvi nejsou žádné otázky.")
            return

        # Rozdělení dostupných otázek podle typu
        available_classic = [q for q in all_questions_in_branch if q.type == "classic"]
        available_bonus = [q for q in all_questions_in_branch if q.type == "bonus"]
        
        # Promíchat pro náhodnost
        import random
        random.shuffle(available_classic)
        random.shuffle(available_bonus)
        
        # 2. Plnění slotů
        # (Iterujeme přes placeholdery a pokud je volný, vezmeme otázku)
        
        assigned_count = 0
        
        # Klasické sloty
        for ph in self.placeholders_q:
            if ph not in self.selection_map: # Volný slot
                # Najít otázku, která ještě NENÍ použita v mapě
                for q in available_classic:
                    if q.id not in self.selection_map.values():
                        self._assign_question_to_slot(ph, q)
                        available_classic.remove(q) # Odebrat, aby se neopakovala
                        assigned_count += 1
                        break
        
        # Bonusové sloty
        for ph in self.placeholders_b:
            if ph not in self.selection_map:
                for q in available_bonus:
                    if q.id not in self.selection_map.values():
                        self._assign_question_to_slot(ph, q)
                        available_bonus.remove(q)
                        assigned_count += 1
                        break
                        
        if assigned_count > 0:
            self.owner.statusBar().showMessage(f"Doplněno {assigned_count} otázek.", 3000)
        else:
            QMessageBox.information(self, "Info", "Nebylo možné doplnit žádné další otázky (buď nejsou volná místa, nebo došly unikátní otázky).")

    def _assign_selected_multi(self) -> None:
        """Přiřadí vybrané otázky ve stromu na první volná místa."""
        items = self.tree_source.selectedItems()
        selected_questions = []
        
        for it in items:
            meta = it.data(0, Qt.UserRole) or {}
            if meta.get("kind") == "question":
                q = self.owner._find_question(meta["parent_group_id"], meta["parent_subgroup_id"], meta["id"])
                if q:
                    selected_questions.append(q)
        
        if not selected_questions:
            QMessageBox.information(self, "Info", "Vyberte ve stromu alespoň jednu otázku.")
            return

        assigned_count = 0
        
        for q in selected_questions:
            if q.id in self.selection_map.values():
                continue
                
            target_ph = None
            if q.type == "classic":
                for ph in self.placeholders_q:
                    if ph not in self.selection_map:
                        target_ph = ph; break
            elif q.type == "bonus":
                for ph in self.placeholders_b:
                    if ph not in self.selection_map:
                        target_ph = ph; break
            
            if target_ph:
                self.selection_map[target_ph] = q.id # Přímý zápis
                assigned_count += 1
        
        if assigned_count > 0:
            self._init_page2() # Hromadný refresh na konci (aktualizuje sloty i strom)
            self.owner.statusBar().showMessage(f"Přiřazeno {assigned_count} otázek.", 3000)
        else:
            self.owner.statusBar().showMessage("Nebylo co přiřadit (vše plné nebo vybrané už použité).", 3000)

    def _assign_question_to_slot(self, ph: str, q: Question) -> None:
        self.selection_map[ph] = q.id
        self._init_page2() # Refresh UI
        
        # Aktualizace UI (tlačítka slotu)
        # Musíme najít widget odpovídající tomuto placeholderu v layoutu
        # Protože nemáme přímou referenci ph -> widget, projdeme layout.
        # (Nebo si můžeme držet mapu ph -> button při vytváření, ale "Maintenance mode" velí neměnit init příliš).
        
        count = self.layout_slots.count()
        for i in range(count):
            item = self.layout_slots.itemAt(i)
            w = item.widget()
            if w and hasattr(w, "property") and w.property("placeholder") == ph:
                # Našli jsme widget (SlotRowWidget nebo podobný)
                # Předpokládám, že má metodu set_question nebo update_ui
                # Pokud neznám vnitřní strukturu SlotRowWidget, musím ji odhadnout z _init_page2 loopu.
                # Ale jelikož nemám kód SlotRowWidget (pokud existuje), udělám to přes refresh celé stránky nebo chytřeji.
                
                # Z logiky _init_page2 (kterou jsem neviděl celou, ale bývá to loop):
                # Obvykle se sloty generují znovu.
                # Pro jednoduchost a robustnost: Zavoláme refresh slotů.
                # Ale to by bylo pomalé.
                
                # Zkusíme najít tlačítko/label v tom widgetu.
                # Předpoklad: w je nějaký container.
                
                # Nejčistší v rámci "Black Box":
                # Znovu vygenerovat sloty je jistota.
                pass
        
        # Protože nemám detailní přístup k widgetům slotů, zavolám obnovu UI slotů.
        # Toto je sice méně efektivní, ale bezpečné.
        self._refresh_slots_ui()

    def _refresh_slots_ui(self) -> None:
        """Znovu vykreslí pravý panel se sloty (volá _init_page2 logiku pro sloty)."""
        # Protože _init_page2 dělá clear a populate, můžeme ji zavolat, 
        # ALE musíme dát pozor, aby nesmazala selection_map (což _init_page2 obvykle nedělá, ta ji čte).
        # Pokud _init_page2 RESCANUJE placeholdery, mohlo by to vadit.
        # V _init_page2 je: if not self.placeholders_q ... scan. Takže ok.
        
        # Zavoláme část _init_page2, která kreslí sloty.
        # Nebo jednoduše celou _init_page2, pokud je idempotentní.
        # Z kódu výše: _init_page2 maže tree a sloty a plní je.
        # To je trochu heavy (refresh tree zruší výběr).
        # Takže raději jen sloty.
        
        # Implementace refresh slotů (zkopírováno/vytaženo z _init_page2 logiky):
        
        # 1. Clear Slots
        while self.layout_slots.count():
            item = self.layout_slots.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.layout_slots.addStretch()
        
        # 2. Re-populate
        # (Tuto logiku nemám k dispozici ve snippetu, musím ji "získat" nebo napsat znovu podle logiky aplikace)
        # Pokud nemám kód pro plnění slotů (loop over placeholders), nemohu to napsat.
        
        # ŘEŠENÍ: Zavolám self._init_page2() s tím, že se smířím se zrušením výběru ve stromu.
        # Uživatel právě klikl na tlačítko nebo menu, takže akce skončila.
        # Obnova stránky je akceptovatelná.
        self._init_page2()


    def _add_slot_widget(self, placeholder_name, allowed_type):
        w = QWidget()
        l = QHBoxLayout(w)
        l.setContentsMargins(0,2,0,2)
        
        lbl_ph = QLabel(f"{placeholder_name}:")
        lbl_ph.setFixedWidth(80)
        
        btn_sel = QPushButton("Vybrat...")
        btn_sel.setFixedWidth(80)
        
        current_qid = self.selection_map.get(placeholder_name)
        current_title = "(nevybráno)"
        if current_qid:
            q = self.owner._find_question_by_id(current_qid)
            if q: current_title = q.title

        lbl_val = QLabel(current_title)
        lbl_val.setStyleSheet("color: gray; font-style: italic;" if not current_qid else "color: white; font-weight: bold;")
        
        btn_clr = QPushButton("X")
        btn_clr.setFixedWidth(30)
        btn_clr.setEnabled(bool(current_qid))
        
        l.addWidget(lbl_ph)
        l.addWidget(lbl_val, 1)
        l.addWidget(btn_sel)
        l.addWidget(btn_clr)
        
        def select_current():
            sel = self.tree_source.selectedItems()
            if not sel:
                QMessageBox.information(self, "Info", "Nejprve označte otázku v levém seznamu.")
                return
            item = sel[0]
            qid = item.data(0, Qt.UserRole)
            if not qid: return
            
            q = self.owner._find_question_by_id(qid)
            if not q: return
            
            if q.type != allowed_type:
                QMessageBox.warning(self, "Typ nesedí", f"Do slotu {placeholder_name} nelze vložit otázku typu {q.type}.")
                return
            
            old_qid = self.selection_map.get(placeholder_name)
            if old_qid: self._show_tree_item(old_qid)
            
            self.selection_map[placeholder_name] = qid
            lbl_val.setText(q.title)
            lbl_val.setStyleSheet("color: white; font-weight: bold;")
            btn_clr.setEnabled(True)
            item.setHidden(True)

        def clear_current():
            old_qid = self.selection_map.get(placeholder_name)
            if old_qid:
                del self.selection_map[placeholder_name]
                self._show_tree_item(old_qid)
            lbl_val.setText("(nevybráno)")
            lbl_val.setStyleSheet("color: gray; font-style: italic;")
            btn_clr.setEnabled(False)

        btn_sel.clicked.connect(select_current)
        btn_clr.clicked.connect(clear_current)
        self.layout_slots.insertWidget(self.layout_slots.count()-1, w)

    def _show_tree_item(self, qid):
        it = QTreeWidgetItemIterator(self.tree_source)
        while it.value():
            if it.value().data(0, Qt.UserRole) == qid:
                it.value().setHidden(False)
                break
            it += 1

    def _show_context_menu(self, position):
        item = self.tree_source.itemAt(position)
        if not item: return
        qid = item.data(0, Qt.UserRole)
        if not qid: return
        
        q = self.owner._find_question_by_id(qid)
        if not q: return

        from PySide6.QtWidgets import QMenu
        from PySide6.QtGui import QAction
        
        menu = QMenu()
        menu_assign = menu.addMenu("Přiřadit k...")
        
        free_slots = []
        if q.type == 'classic':
            for ph in self.placeholders_q:
                if ph not in self.selection_map: free_slots.append(ph)
        else:
            for ph in self.placeholders_b:
                if ph not in self.selection_map: free_slots.append(ph)
        
        if not free_slots:
            a = menu_assign.addAction("(Žádné volné sloty)")
            a.setEnabled(False)
        else:
            for slot in free_slots:
                action = QAction(slot, self.tree_source)
                action.triggered.connect(lambda checked=False, s=slot, q_id=qid: self._assign_from_context(s, q_id))
                menu_assign.addAction(action)

        menu.exec(self.tree_source.viewport().mapToGlobal(position))

    def _assign_from_context(self, slot_name, qid):
        self.selection_map[slot_name] = qid
        
        # Refresh slot UI
        for i in range(self.layout_slots.count()):
            w = self.layout_slots.itemAt(i).widget()
            if w and isinstance(w, QWidget):
                children = w.findChildren(QLabel)
                if children and children[0].text() == f"{slot_name}:":
                    # Found slot widget
                    lbl_val = w.layout().itemAt(1).widget()
                    btn_clr = w.layout().itemAt(3).widget()
                    q = self.owner._find_question_by_id(qid)
                    lbl_val.setText(q.title)
                    lbl_val.setStyleSheet("color: white; font-weight: bold;")
                    btn_clr.setEnabled(True)
                    break
        
        # Hide tree item
        it = QTreeWidgetItemIterator(self.tree_source)
        while it.value():
            if it.value().data(0, Qt.UserRole) == qid:
                it.value().setHidden(True)
                break
            it += 1

    def _init_page3(self):
        try:
            # 1. Generování hashe (NOVÉ)
            ts = str(datetime.now().timestamp())
            salt = secrets.token_hex(16)
            data_to_hash = f"{ts}{salt}"
            self._cached_hash = hashlib.sha3_256(data_to_hash.encode("utf-8")).hexdigest()
            
            # Zobrazení hashe v labelu (pokud existuje z _build_page3_content)
            if hasattr(self, "lbl_hash_preview"):
                self.lbl_hash_preview.setText(f"SHA3-512 Hash:\n{self._cached_hash}")

            t_name = self.template_path.name if self.template_path else "Nevybráno"
            o_name = self.output_path.name if self.output_path else "Nevybráno"
            self.lbl_templ_p3.setText(t_name)
            self.lbl_out_p3.setText(o_name)

            total_bonus_points = 0.0
            min_loss = 0.0
            
            # Barvy
            bg_color = "#252526"; text_color = "#e0e0e0"; border_color = "#555555"
            sec_q_bg = "#2d3845"; sec_b_bg = "#453d2d"; sec_s_bg = "#2d452d"
            
            # Sestavení verze (ZMĚNA: Bez UUID na konci)
            prefix = self.le_prefix.text().strip()
            today = datetime.now().strftime("%Y-%m-%d")
            verze_preview = f"{prefix} {today}" 
            
            html = f"""
            <html>
            <body style='font-family: Arial, sans-serif; background-color: {bg_color}; color: {text_color};'>
            <h2 style='color: #61dafb; border-bottom: 2px solid #61dafb;'>Souhrn testu</h2>
            <table width='100%' style='margin-bottom: 20px; color: {text_color};'>
                <tr>
                    <td><b>Verze:</b> {verze_preview}</td>
                    <td align='right'><b>Datum:</b> {self.dt_edit.dateTime().toString("dd.MM.yyyy HH:mm")}</td>
                </tr>
                <tr>
                    <td colspan='2' style='font-size: 10px; color: #888;'><b>Hash:</b> {self._cached_hash[:32]}...</td>
                </tr>
            </table>
            """

           # Klasické
            html += f"<h3 style='background-color: {sec_q_bg}; padding: 5px; border-left: 4px solid #4da6ff;'>1. Klasické otázky</h3>"
            html += f"<table width='100%' border='0' cellspacing='0' cellpadding='5' style='color: {text_color};'>"
            for ph in self.placeholders_q:
                qid = self.selection_map.get(ph)
                if qid:
                    q = self.owner._find_question_by_id(qid)
                    if q:
                        title_clean = re.sub(r'<[^>]+>', '', q.title)
                        html += f"<tr><td width='100' style='color:#888;'>{ph}:</td><td><b>{title_clean}</b></td><td align='right'>({q.points} b)</td></tr>"
                else:
                    html += f"<tr><td width='100' style='color:#ff5555;'>{ph}:</td><td colspan='2' style='color:#ff5555;'>--- NEVYPLNĚNO ---</td></tr>"
            html += "</table>"

            # Bonusy
            html += f"<h3 style='background-color: {sec_b_bg}; padding: 5px; border-left: 4px solid #ffcc00;'>2. Bonusové otázky</h3>"
            html += f"<table width='100%' border='0' cellspacing='0' cellpadding='5' style='color: {text_color};'>"
            for ph in self.placeholders_b:
                qid = self.selection_map.get(ph)
                if qid:
                    q = self.owner._find_question_by_id(qid)
                    if q:
                        total_bonus_points += float(q.bonus_correct)
                        min_loss += float(q.bonus_wrong)
                        title_clean = re.sub(r'<[^>]+>', '', q.title)
                        html += f"<tr><td width='100' style='color:#888;'>{ph}:</td><td><b>{title_clean}</b></td><td align='right' style='color:#81c784;'>+{q.bonus_correct} / <span style='color:#e57373;'>{q.bonus_wrong}</span></td></tr>"
                else:
                    html += f"<tr><td width='100' style='color:#ff5555;'>{ph}:</td><td colspan='2' style='color:#ff5555;'>--- NEVYPLNĚNO ---</td></tr>"
            html += "</table>"

            # Výpočet MaxBody
            max_body_val = 10.0 + total_bonus_points

            # Klasifikace
            html += f"<h3 style='background-color: {sec_s_bg}; padding: 5px; border-left: 4px solid #66bb6a;'>3. Klasifikace</h3>"
            html += f"""
            <p><b>Max. bodů:</b> {max_body_val:.2f} (10 + {total_bonus_points:.2f}) &nbsp;&nbsp;|&nbsp;&nbsp; <b>Min. bodů (penalizace):</b> {min_loss:.2f}</p>
            <table width='60%' border='1' cellspacing='0' cellpadding='5' style='border-collapse: collapse; border: 1px solid {border_color}; color: {text_color};'>
                <tr style='background-color: #333;'><th>Známka</th><th>Interval</th></tr>
                <tr><td align='center' style='color:#81c784'><b>A</b></td><td>&lt; 9.2 ; <b>{max_body_val:.2f}</b> &gt;</td></tr>
                <tr><td align='center' style='color:#a5d6a7'><b>B</b></td><td>&lt; 8.4 ; 9.2 )</td></tr>
                <tr><td align='center' style='color:#c8e6c9'><b>C</b></td><td>&lt; 7.6 ; 8.4 )</td></tr>
                <tr><td align='center' style='color:#fff59d'><b>D</b></td><td>&lt; 6.8 ; 7.6 )</td></tr>
                <tr><td align='center' style='color:#ffcc80'><b>E</b></td><td>&lt; 6.0 ; 6.8 )</td></tr>
                <tr><td align='center' style='color:#ef5350'><b>F</b></td><td>&lt; <b>{min_loss:.2f}</b> ; 6.0 )</td></tr>
            </table>
            """
            html += "</body></html>"
            self.preview_edit.setHtml(html)
            
        except Exception as e:
            print(f"CRITICAL ERROR in _init_page3: {e}")
            import traceback
            traceback.print_exc()
            self.preview_edit.setText(f"Chyba při generování náhledu: {e}")

    def accept(self) -> None:
        if not self.template_path or not self.output_path:
            return

        repl_plain: Dict[str, str] = {}
        
        # Datum
        dt = round_dt_to_10m(self.dt_edit.dateTime())
        dt_str = f"{cz_day_of_week(dt.toPython())} {dt.toString('dd.MM.yyyy HH:mm')}"
        repl_plain["DatumČas"] = dt_str
        repl_plain["DatumCas"] = dt_str
        repl_plain["DATUMCAS"] = dt_str
        
        # Verze
        prefix = self.le_prefix.text().strip()
        today = datetime.now().strftime("%Y-%m-%d")
        verze_str = f"{prefix} {today}"
        repl_plain["PoznamkaVerze"] = verze_str
        repl_plain["POZNAMKAVERZE"] = verze_str
        
        # Kontrolní Hash
        k_hash = getattr(self, "_cached_hash", "")
        if not k_hash:
            ts = str(datetime.now().timestamp())
            salt = secrets.token_hex(16)
            data_to_hash = f"{ts}{salt}"
            k_hash = hashlib.sha3_512(data_to_hash.encode("utf-8")).hexdigest()

        repl_plain["KontrolniHash"] = k_hash
        repl_plain["KONTROLNIHASH"] = k_hash
        
        # Body (MaxBody = 10 + Bonusy)
        total_bonus = 0.0
        min_loss = 0.0
        
        for qid in self.selection_map.values():
            q = self.owner._find_question_by_id(qid)
            if not q: continue
            
            if q.type == 'bonus':
                total_bonus += float(q.bonus_correct)
                min_loss += float(q.bonus_wrong)
        
        max_body = 10.0 + total_bonus
        
        repl_plain["MaxBody"] = f"{max_body:.2f}"
        repl_plain["MAXBODY"] = f"{max_body:.2f}"
        repl_plain["MinBody"] = f"{min_loss:.2f}"
        repl_plain["MINBODY"] = f"{min_loss:.2f}"

        # Rich text map
        rich_map: Dict[str, str] = {}
        for ph, qid in self.selection_map.items():
            q = self.owner._find_question_by_id(qid)
            if q:
                rich_map[ph] = q.text_html

        try:
            self.owner._generate_docx_from_template(self.template_path, self.output_path, repl_plain, rich_map)
        except Exception as e:
            QMessageBox.critical(self, "Export", f"Chyba při exportu:\n{e}")
            return

        # NOVÉ: Registrace exportu do historie
        self.owner.register_export(self.output_path.name, k_hash)

        QMessageBox.information(self, "Export", f"Export dokončen.\nSoubor uložen:\n{self.output_path}")
        super().accept()

# --------------------------- Hlavní okno (UI + logika) ---------------------------

class FunnyAnswerDialog(QDialog):
    """Dialog pro přidání nové vtipné odpovědi."""
    def __init__(self, parent=None, project_root: Optional[Path] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Přidat vtipnou odpověď")
        self.resize(600, 400)
        
        self.project_root = project_root
        
        layout = QVBoxLayout(self)
        form = QFormLayout()
        
        # Výběr zdrojové písemky
        self.combo_source = QComboBox()
        self.combo_source.addItem("(Ruční zadání / Bez vazby na soubor)", None)
        self.combo_source.currentIndexChanged.connect(self._on_source_changed)
        
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Znění vtipné odpovědi...")
        
        self.author_edit = QLineEdit()
        self.author_edit.setPlaceholderText("Např. Student, Anonym...")
        
        self.date_edit = QDateTimeEdit(QDateTime.currentDateTime())
        self.date_edit.setDisplayFormat("dd.MM.yyyy HH:mm") # Přidán i čas pro kontrolu
        self.date_edit.setCalendarPopup(True)

        form.addRow("Zdroj písemky:", self.combo_source)
        form.addRow("Odpověď:", self.text_edit)
        form.addRow("Autor:", self.author_edit)
        form.addRow("Datum:", self.date_edit)
        
        layout.addLayout(form)
        
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)
        
        # Načtení souborů
        self._load_files()

    def _load_files(self) -> None:
        if not self.project_root:
            return
            
        base_dir = self.project_root / "data" / "Vygenerované testy"
        if not base_dir.exists():
            return
            
        # Rekurzivní vyhledání všech .docx
        found_files = sorted(base_dir.rglob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
        
        for p in found_files:
            # Zobrazíme relativní cestu vůči složce vygenerovaných testů pro přehlednost
            try:
                rel_path = p.relative_to(base_dir)
                display_text = str(rel_path)
            except ValueError:
                display_text = p.name
            
            self.combo_source.addItem(display_text, str(p))

    def _on_source_changed(self, index: int) -> None:
        data = self.combo_source.itemData(index)
        if not data:
            return
            
        path = Path(data)
        filename = path.name
        
        # Očekávaný formát: Prefix_YYYY-MM-DD_HHMM_...
        parts = filename.split('_')
        
        if len(parts) >= 3:
            date_str = parts[1]  # YYYY-MM-DD
            time_str = parts[2]  # HHMM
            
            # Validace a parsování
            try:
                # Datum
                qdate = QDateTime.fromString(date_str, "yyyy-MM-dd").date()
                
                # Čas (HHMM)
                qtime = QDateTime.fromString(time_str, "HHmm").time()
                
                if qdate.isValid() and qtime.isValid():
                    dt = QDateTime(qdate, qtime)
                    self.date_edit.setDateTime(dt)
            except Exception:
                # Pokud parsování selže, neděláme nic (necháme aktuální)
                pass
            
    def set_data(self, text: str, date_str: str, author: str) -> None:
        """Naplní formulář daty pro editaci."""
        self.text_edit.setText(text)
        self.author_edit.setText(author)
        
        # Pokusíme se parsovat datum (očekáváme dd.MM.yyyy nebo dd.MM.yyyy HH:mm)
        # Zkusíme s časem
        dt = QDateTime.fromString(date_str, "dd.MM.yyyy HH:mm")
        if not dt.isValid():
            # Zkusíme bez času
            dt = QDateTime.fromString(date_str, "dd.MM.yyyy")
        
        if dt.isValid():
            self.date_edit.setDateTime(dt)


    def get_data(self) -> tuple[str, str, str]:
        # Vracíme datum ve formátu stringu, jak je zvykem v aplikaci (bez času, nebo s časem dle preference?)
        # Původní kód používal dd.MM.yyyy, ale formát v tabulce FunnyAnswer je jen string.
        # Pokud chceme zachovat čas získaný z názvu souboru, formátujeme ho.
        return (
            self.text_edit.toPlainText().strip(),
            self.date_edit.dateTime().toString("dd.MM.yyyy HH:mm"), 
            self.author_edit.text().strip()
        )

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
        self.resize(1800, 900)

        self.project_root = Path.cwd()
        default_data_dir = self.project_root / "data"
        default_data_dir.mkdir(parents=True, exist_ok=True)
        self.data_path = data_path or (default_data_dir / "questions.json")

        # Aplikace ikona (pokud existuje)
        icon_file = self.project_root / "icon" / "icon.png"
        if icon_file.exists():
            app_icon = QIcon(str(icon_file))
            self.setWindowIcon(app_icon)
            QApplication.instance().setWindowIcon(app_icon)

        self.root: RootData = RootData(groups=[])
        self._current_question_id: Optional[str] = None
        self._current_node_kind: Optional[str] = None

        self._autosave_timer = QTimer(self)
        self._autosave_timer.setSingleShot(True)
        self._autosave_timer.setInterval(1200)
        self._autosave_timer.timeout.connect(self._autosave_current_question)

        self._build_ui()
        self._connect_signals()
        
        # Nové: Kontextové menu pro strom
        self.tree.customContextMenuRequested.connect(self._on_context_menu)
        
        self._build_menus()
        self.load_data()
        self._refresh_tree()
        
        # ZMĚNA: Strom 60%, Editor 40% (cca 840px : 560px)
        # Nyní, když je self.splitter správně nastaven v _build_ui, můžeme přímo nastavit velikosti.
        self.splitter.setSizes([940, 860])

    def _on_context_menu(self, pos) -> None:
        item = self.tree.itemAt(pos)
        if not item:
            return
            
        meta = item.data(0, Qt.UserRole) or {}
        kind = meta.get("kind")
        
        if kind == "question":
            menu = QMenu(self)
            action_dup = QAction("Duplikovat otázku", self)
            action_dup.triggered.connect(self._duplicate_question)
            menu.addAction(action_dup)
            menu.exec(self.tree.mapToGlobal(pos))

    def _duplicate_question(self) -> None:
        kind, meta = self._selected_node()
        if kind != "question":
            return

        gid = meta.get("parent_group_id")
        sgid = meta.get("parent_subgroup_id")
        qid = meta.get("id")

        q_orig = self._find_question(gid, sgid, qid)
        if not q_orig:
            return

        # Vytvoření kopie
        data = asdict(q_orig)
        data["id"] = str(_uuid.uuid4())
        data["title"] = (q_orig.title or "Otázka") + " (kopie)"
        
        new_q = Question(**data)

        # Vložení do správné podskupiny
        target_sg = self._find_subgroup(gid, sgid)
        if target_sg:
            target_sg.questions.append(new_q)
            self._refresh_tree()
            self._select_question(new_q.id)
            self.save_data()
            self.statusBar().showMessage("Otázka byla duplikována.", 3000)

    def _build_ui(self) -> None:
        self.splitter = QSplitter()
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(8)

        # LEVÝ PANEL (nyní obsahuje záložky)
        left_panel_container = QWidget()
        left_container_layout = QVBoxLayout(left_panel_container)
        left_container_layout.setContentsMargins(0, 0, 0, 0)

        self.left_tabs = QTabWidget()
        
        # --- ZÁLOŽKA 1: OTÁZKY (Původní obsah) ---
        self.tab_questions = QWidget()
        questions_layout = QVBoxLayout(self.tab_questions)
        questions_layout.setContentsMargins(4, 4, 4, 4)
        questions_layout.setSpacing(6)

        filter_bar = QWidget()
        filter_layout = QHBoxLayout(filter_bar)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(6)
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Filtr: název / obsah otázky…")
        self.btn_move_selected = QPushButton("Přesunout vybrané…")
        self.btn_delete_selected = QPushButton("Smazat vybrané")
        filter_layout.addWidget(self.filter_edit, 1)
        filter_layout.addWidget(self.btn_move_selected)
        filter_layout.addWidget(self.btn_delete_selected)
        questions_layout.addWidget(filter_bar)

        self.tree = DnDTree(self)
        questions_layout.addWidget(self.tree, 1)
        
        self.left_tabs.addTab(self.tab_questions, "Otázky")

        # --- ZÁLOŽKA 2: HISTORIE (Upraveno) ---
        self.tab_history = QWidget()
        history_layout = QVBoxLayout(self.tab_history)
        history_layout.setContentsMargins(4, 4, 4, 4)
        
        self.table_history = QTableWidget(0, 2)
        self.table_history.setHorizontalHeaderLabels(["Soubor", "Hash (SHA3-256)"])
        self.table_history.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents) 
        self.table_history.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_history.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_history.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_history.setSortingEnabled(True)
        
        # NOVÉ: Kontextové menu pro historii
        self.table_history.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_history.customContextMenuRequested.connect(self._on_history_context_menu)

        history_layout.addWidget(self.table_history)
        
        btn_refresh_hist = QPushButton("Obnovit historii")
        btn_refresh_hist.clicked.connect(self._refresh_history_table)
        history_layout.addWidget(btn_refresh_hist)

        self.left_tabs.addTab(self.tab_history, "Historie")
        
        left_container_layout.addWidget(self.left_tabs)


        # PRAVÝ PANEL (Detail / Editor)
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

        self.form_layout = QFormLayout()
        self.form_layout.setLabelAlignment(Qt.AlignLeft)

        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("Krátký název otázky…")

        self.combo_type = QComboBox(); self.combo_type.addItems(["Klasická", "BONUS"])
        self.spin_points = QSpinBox(); self.spin_points.setRange(-999, 999); self.spin_points.setValue(1)
        self.spin_bonus_correct = QDoubleSpinBox(); self.spin_bonus_correct.setDecimals(2); self.spin_bonus_correct.setSingleStep(0.01); self.spin_bonus_correct.setRange(-999.99, 999.99); self.spin_bonus_correct.setValue(1.00)
        self.spin_bonus_wrong = QDoubleSpinBox(); self.spin_bonus_wrong.setDecimals(2); self.spin_bonus_wrong.setSingleStep(0.01); self.spin_bonus_wrong.setRange(-999.99, 999.99); self.spin_bonus_wrong.setValue(0.00)

        self.form_layout.addRow("Název otázky:", self.title_edit)
        self.form_layout.addRow("Typ otázky:", self.combo_type)
        self.form_layout.addRow("Body (klasická):", self.spin_points)
        self.form_layout.addRow("Body za správně (BONUS):", self.spin_bonus_correct)
        self.form_layout.addRow("Body za špatně (BONUS):", self.spin_bonus_wrong)

        # --- NOVÉ: Správná odpověď ---
        self.edit_correct_answer = QTextEdit()
        self.edit_correct_answer.setPlaceholderText("Volitelný text správné odpovědi...")
        self.edit_correct_answer.setFixedHeight(60)
        self.form_layout.addRow("Správná odpověď:", self.edit_correct_answer)

        # --- NOVÉ: Vtipné odpovědi (Tabulka + Tlačítka) ---
        self.funny_container = QWidget()
        fc_layout = QVBoxLayout(self.funny_container)
        fc_layout.setContentsMargins(0,0,0,0)
        
        self.table_funny = QTableWidget(0, 3)
        self.table_funny.setHorizontalHeaderLabels(["Odpověď", "Datum", "Jméno"])
        self.table_funny.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table_funny.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_funny.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table_funny.setFixedHeight(120)
        self.table_funny.setSelectionBehavior(QAbstractItemView.SelectRows)
        
        btns_layout = QHBoxLayout()
        self.btn_add_funny = QPushButton("Přidat vtipnou odpoveď")
        self.btn_rem_funny = QPushButton("Odebrat")
        btns_layout.addWidget(self.btn_add_funny)
        btns_layout.addWidget(self.btn_rem_funny)
        btns_layout.addStretch()

        fc_layout.addLayout(btns_layout)
        fc_layout.addWidget(self.table_funny)

        self.form_layout.addRow("Vtipné odpovědi:", self.funny_container)

        self.text_edit = QTextEdit()
        self.text_edit.setAcceptRichText(True)
        self.text_edit.setPlaceholderText("Sem napište znění otázky…\nPodporováno: tučné, kurzíva, podtržení, barva, odrážky, zarovnání.")
        self.text_edit.setMinimumHeight(200)

        self.btn_save_question = QPushButton("Uložit změny otázky"); self.btn_save_question.setDefault(True)

        self.rename_panel = QWidget()
        rename_layout = QFormLayout(self.rename_panel)
        self.rename_line = QLineEdit()
        self.btn_rename = QPushButton("Uložit název")
        rename_layout.addRow("Název:", self.rename_line)
        rename_layout.addRow(self.btn_rename)

        self.detail_layout.addWidget(self.editor_toolbar)
        self.detail_layout.addLayout(self.form_layout)
        self.detail_layout.addWidget(self.text_edit, 1)
        self.detail_layout.addWidget(self.btn_save_question)
        self.detail_layout.addWidget(self.rename_panel)
        self._set_editor_enabled(False)

        self.splitter.addWidget(left_panel_container)
        self.splitter.addWidget(self.detail_stack)
        self.splitter.setStretchFactor(1, 1)
        self.setCentralWidget(self.splitter)

        tb = self.addToolBar("Hlavní")
        tb.setIconSize(QSize(18, 18))
        self.act_add_group = QAction("Přidat skupinu", self)
        self.act_add_subgroup = QAction("Přidat podskupinu", self)
        self.act_add_question = QAction("Přidat otázku", self)
        self.act_delete = QAction("Smazat", self)

        self.act_add_group.setShortcut("Ctrl+G")
        self.act_add_subgroup.setShortcut("Ctrl+Shift+G")
        self.act_add_question.setShortcut(QKeySequence.New)
        self.act_delete.setShortcut(QKeySequence.Delete)

        tb.addAction(self.act_add_group)
        tb.addAction(self.act_add_subgroup)
        tb.addAction(self.act_add_question)


        tb.addSeparator()
        tb.addAction(self.act_delete)

        self.statusBar().showMessage(f"Datový soubor: {self.data_path}")
        
        self._refresh_history_table()

    def _refresh_history_table(self) -> None:
        """Načte historii exportů z history.json a naplní tabulku."""
        history_file = self.project_root / "data" / "history.json"
        history = []
        if history_file.exists():
            try:
                with open(history_file, "r", encoding="utf-8") as f:
                    history = json.load(f)
            except Exception as e:
                print(f"Chyba při čtení historie: {e}")
        
        self.table_history.setRowCount(0)
        self.table_history.setSortingEnabled(False)
        
        for entry in history:
            row = self.table_history.rowCount()
            self.table_history.insertRow(row)
            
            fn = entry.get("filename", "")
            h = entry.get("hash", "")
            
            self.table_history.setItem(row, 0, QTableWidgetItem(fn))
            
            h_item = QTableWidgetItem(h)
            h_item.setTextAlignment(Qt.AlignCenter)
            self.table_history.setItem(row, 1, h_item)
            
        self.table_history.setSortingEnabled(True)

    def _on_history_context_menu(self, pos) -> None:
        """Zobrazí kontextové menu pro tabulku historie."""
        items = self.table_history.selectedItems()
        if not items:
            return
            
        menu = QMenu(self)
        act_del = QAction("Smazat záznam(y)", self)
        act_del.triggered.connect(self._delete_history_items)
        menu.addAction(act_del)
        menu.exec(self.table_history.mapToGlobal(pos))

    def _delete_history_items(self) -> None:
        """Smaže vybrané záznamy z historie."""
        # Získáme unikátní řádky
        rows = sorted(set(index.row() for index in self.table_history.selectedIndexes()), reverse=True)
        if not rows:
            return

        if QMessageBox.question(self, "Smazat", f"Opravdu smazat {len(rows)} záznamů z historie?") != QMessageBox.Yes:
            return

        # Musíme smazat data z JSONu.
        # Protože tabulka může být seřazená jinak než JSON, musíme identifikovat záznamy podle obsahu (filename + hash).
        # Nebo jednodušeji: Načteme JSON, odstraníme ty, co odpovídají vybraným řádkům.
        
        to_remove = [] # List of (filename, hash)
        for r in rows:
            fn = self.table_history.item(r, 0).text()
            h = self.table_history.item(r, 1).text()
            to_remove.append((fn, h))

        history_file = self.project_root / "data" / "history.json"
        history = []
        if history_file.exists():
            try:
                with open(history_file, "r", encoding="utf-8") as f:
                    history = json.load(f)
            except Exception:
                pass

        # Filtrace (ponecháme ty, co nejsou v to_remove)
        new_history = []
        for entry in history:
            match = False
            for r_fn, r_h in to_remove:
                if entry.get("filename") == r_fn and entry.get("hash") == r_h:
                    match = True
                    break
            if not match:
                new_history.append(entry)
        
        # Uložení
        try:
            with open(history_file, "w", encoding="utf-8") as f:
                json.dump(new_history, f, indent=2, ensure_ascii=False)
        except Exception as e:
            QMessageBox.warning(self, "Chyba", f"Nelze uložit historii:\n{e}")

        self._refresh_history_table()


    def register_export(self, filename: str, k_hash: str) -> None:
        """Zaznamená nový export a obnoví tabulku."""
        history_file = self.project_root / "data" / "history.json"
        history = []
        if history_file.exists():
            try:
                with open(history_file, "r", encoding="utf-8") as f:
                    history = json.load(f)
            except Exception:
                pass # Ignorujeme chyby čtení, vytvoříme nový seznam
        
        # Přidání záznamu
        history.append({
            "filename": filename, 
            "hash": k_hash, 
            "date": datetime.now().isoformat()
        })
        
        # Uložení
        try:
            with open(history_file, "w", encoding="utf-8") as f:
                json.dump(history, f, indent=2, ensure_ascii=False)
        except Exception as e:
            QMessageBox.warning(self, "Chyba historie", f"Nepodařilo se uložit historii exportu:\n{e}")
            
        self._refresh_history_table()


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
        # self.tree.itemChanged.connect(self._on_tree_item_changed) # REMOVED previously
        
        self.btn_save_question.clicked.connect(self._on_save_question_clicked)
        self.btn_rename.clicked.connect(self._on_rename_clicked)
        
        # Tree actions context menu
        self.act_add_group.triggered.connect(self._add_group)
        self.act_add_subgroup.triggered.connect(self._add_subgroup)
        self.act_add_question.triggered.connect(self._add_question)
        
        # Delete shortcut
        self.act_delete.triggered.connect(self._bulk_delete_selected) 
        self.btn_delete_selected.clicked.connect(self._bulk_delete_selected)
        
        # Autosave triggers
        self.title_edit.textChanged.connect(self._autosave_schedule)
        self.combo_type.currentIndexChanged.connect(self._on_type_changed_ui)
        self.spin_points.valueChanged.connect(self._autosave_schedule)
        self.spin_bonus_correct.valueChanged.connect(self._autosave_schedule)
        self.spin_bonus_wrong.valueChanged.connect(self._autosave_schedule)
        self.text_edit.textChanged.connect(self._autosave_schedule)
        
        # Autosave triggers (New fields)
        self.edit_correct_answer.textChanged.connect(self._autosave_schedule)
        self.table_funny.itemChanged.connect(self._autosave_schedule)
        
        # NOVÉ: Kontextové menu pro vtipné odpovědi (Editace)
        self.table_funny.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_funny.customContextMenuRequested.connect(self._on_funny_context_menu)

        # Formátování
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
        
        # Filter
        self.filter_edit.textChanged.connect(self._apply_filter)
        
        # Drag Drop Move (btn)
        self.btn_move_selected.clicked.connect(self._move_selected_dialog)

        # Tlačítka pro vtipné odpovědi
        self.btn_add_funny.clicked.connect(self._add_funny_row)
        self.btn_rem_funny.clicked.connect(self._remove_funny_row)

    def _add_funny_row(self) -> None:
        # Předáváme self.project_root pro vyhledání souborů
        dlg = FunnyAnswerDialog(self, project_root=self.project_root)
        
        if dlg.exec() == QDialog.Accepted:
            text, date_str, author = dlg.get_data()
            
            row = self.table_funny.rowCount()
            self.table_funny.insertRow(row)
            
            self.table_funny.setItem(row, 0, QTableWidgetItem(text))
            self.table_funny.setItem(row, 1, QTableWidgetItem(date_str))
            self.table_funny.setItem(row, 2, QTableWidgetItem(author))
            
            self._autosave_schedule()


    def _remove_funny_row(self) -> None:
        rows = sorted(set(index.row() for index in self.table_funny.selectedIndexes()), reverse=True)
        for r in rows:
            self.table_funny.removeRow(r)
        if rows:
            self._autosave_schedule()

    def _on_funny_context_menu(self, pos) -> None:
        """Kontextové menu pro tabulku vtipných odpovědí."""
        item = self.table_funny.itemAt(pos)
        if not item:
            return
        
        menu = QMenu(self)
        act_edit = QAction("Upravit odpověď", self)
        act_edit.triggered.connect(lambda: self._edit_funny_row(item.row()))
        menu.addAction(act_edit)
        
        act_del = QAction("Smazat odpověď", self)
        act_del.triggered.connect(self._remove_funny_row) # Použije selected items
        menu.addAction(act_del)
        
        menu.exec(self.table_funny.mapToGlobal(pos))

    def _edit_funny_row(self, row: int) -> None:
        """Otevře dialog pro editaci vtipné odpovědi na daném řádku."""
        # Načtení dat z tabulky
        text_item = self.table_funny.item(row, 0)
        date_item = self.table_funny.item(row, 1)
        author_item = self.table_funny.item(row, 2)
        
        if not text_item or not date_item or not author_item:
            return
            
        old_text = text_item.text()
        old_date = date_item.text()
        old_author = author_item.text()
        
        # Otevření dialogu
        dlg = FunnyAnswerDialog(self, project_root=self.project_root)
        dlg.setWindowTitle("Upravit vtipnou odpověď")
        dlg.set_data(old_text, old_date, old_author)
        
        if dlg.exec() == QDialog.Accepted:
            new_text, new_date, new_author = dlg.get_data()
            
            # Uložení zpět do tabulky
            self.table_funny.setItem(row, 0, QTableWidgetItem(new_text))
            self.table_funny.setItem(row, 1, QTableWidgetItem(new_date))
            self.table_funny.setItem(row, 2, QTableWidgetItem(new_author))
            
            self._autosave_schedule()


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
        
    def _move_selected_dialog(self) -> None:
        """Otevře dialog pro přesun vybraných otázek do jiné skupiny/podskupiny."""
        ids = self._selected_question_ids()
        if not ids:
            QMessageBox.information(self, "Přesun", "Vyberte otázky k přesunu.")
            return

        # MoveTargetDialog musí být definován v souboru (byl vidět v původním výpisu)
        dlg = MoveTargetDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return

        gid, sgid = dlg.selected_target()
        if not gid:
            return

        # Nalezení cílové podskupiny
        target_sg: Optional[Subgroup] = None
        if sgid:
            target_sg = self._find_subgroup(gid, sgid)
        else:
            # Cíl je skupina -> zkusíme najít první podskupinu nebo vytvoříme Default
            g = self._find_group(gid)
            if g:
                if g.subgroups:
                    target_sg = g.subgroups[0]
                else:
                    # Vytvoření defaultní podskupiny, pokud skupina žádnou nemá
                    new_sg = Subgroup(id=str(_uuid.uuid4()), name="Default", subgroups=[], questions=[])
                    g.subgroups.append(new_sg)
                    target_sg = new_sg

        if not target_sg:
            QMessageBox.warning(self, "Chyba", "Cílová skupina/podskupina nebyla nalezena.")
            return

        # PROVEDENÍ PŘESUNU
        # 1. Najdeme a vyjmeme otázky z původních umístění
        moved_questions: List[Question] = []

        def remove_from_list(sgs: List[Subgroup]):
            for sg in sgs:
                # Ponecháme jen ty, které NEJSOU v seznamu k přesunu
                # Ty co JSOU, si uložíme
                to_keep = []
                for q in sg.questions:
                    if q.id in ids:
                        moved_questions.append(q)
                    else:
                        to_keep.append(q)
                sg.questions = to_keep
                
                # Rekurze
                remove_from_list(sg.subgroups)

        for g in self.root.groups:
            remove_from_list(g.subgroups)

        # 2. Vložíme je do cíle
        # (Otázky se přidají na konec cílové podskupiny)
        target_sg.questions.extend(moved_questions)

        # 3. Uložit a obnovit
        self._refresh_tree()
        self._reselect_questions(ids) # Zkusíme znovu označit přesunuté
        self.save_data()
        self.statusBar().showMessage(f"Přesunuto {len(moved_questions)} otázek.", 3000)


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
            
        # NOVÉ: Deserializace vtipných odpovědí
        f_answers_raw = q.get("funny_answers", [])
        f_answers = []
        for item in f_answers_raw:
            if isinstance(item, dict):
                f_answers.append(FunnyAnswer(
                    text=item.get("text", ""),
                    date=item.get("date", ""),
                    author=item.get("author", "")
                ))
                
        return Question(
            id=q.get("id", ""),
            type=q.get("type", "classic"),
            text_html=q.get("text_html", "<p><br></p>"),
            title=title,
            points=int(q.get("points", 1)),
            bonus_correct=bc,
            bonus_wrong=bw,
            created_at=q.get("created_at", ""),
            correct_answer=q.get("correct_answer", ""),
            funny_answers=f_answers
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
        
        # Skupiny neřadíme, bereme jak jsou
        for g in self.root.groups:
            g_item = QTreeWidgetItem([g.name, ""]) # Prázdný text ve sloupci 1
            g_item.setData(0, Qt.UserRole, {"kind": "group", "id": g.id})
            g_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
            f = g_item.font(0); f.setBold(True); g_item.setFont(0, f)
            self.tree.addTopLevelItem(g_item)
            
            # Podskupiny seřadíme abecedně podle jména (case-insensitive)
            sorted_subgroups = sorted(g.subgroups, key=lambda s: s.name.lower())
            self._add_subgroups_to_item(g_item, g.id, sorted_subgroups)

        self.tree.expandAll()
        # Vynutíme přepočet šířky sloupce 1 podle obsahu
        self.tree.resizeColumnToContents(1)

    def _add_subgroups_to_item(self, parent_item: QTreeWidgetItem, group_id: str, subgroups: List[Subgroup]) -> None:
        # Pozn.: Vstupní 'subgroups' už může být seřazený z _refresh_tree, ale pro rekurzi (vnořené podskupiny)
        # a pro otázky to musíme řešit zde.
        
        # Pokud bychom spoléhali na to, že 'subgroups' na vstupu je seřazené, je to OK pro první úroveň volání.
        # Pro rekurzi si to raději pojistíme nebo seřadíme při volání.
        # Zde iterujeme přes seznam, který nám byl předán.
        
        for sg in subgroups:
            parent_meta = parent_item.data(0, Qt.UserRole) or {}
            parent_sub_id = parent_meta.get("id") if parent_meta.get("kind") == "subgroup" else None
            
            sg_item = QTreeWidgetItem([sg.name, ""])
            sg_item.setData(0, Qt.UserRole, {
                "kind": "subgroup", 
                "id": sg.id, 
                "parent_group_id": group_id, 
                "parent_subgroup_id": parent_sub_id
            })
            sg_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirOpenIcon))
            parent_item.addChild(sg_item)
            
            # 1. Přidat Otázky (Seřazené abecedně podle titulku)
            sorted_questions = sorted(sg.questions, key=lambda q: (q.title or "").lower())
            
            for q in sorted_questions:
                label = "Klasická" if q.type == "classic" else "BONUS"
                pts = q.points if q.type == "classic" else self._bonus_points_label(q)
                
                q_item = QTreeWidgetItem([q.title or "Otázka", f"{label} | {pts}"])
                q_item.setData(0, Qt.UserRole, {
                    "kind": "question", 
                    "id": q.id, 
                    "parent_group_id": group_id, 
                    "parent_subgroup_id": sg.id
                })
                self._apply_question_item_visuals(q_item, q.type)
                sg_item.addChild(q_item)
            
            # 2. Rekurze pro vnořené podskupiny (Seřazené)
            if sg.subgroups:
                sorted_nested_subgroups = sorted(sg.subgroups, key=lambda s: s.name.lower())
                self._add_subgroups_to_item(sg_item, group_id, sorted_nested_subgroups)

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
        """Deprecated: Redirects to bulk delete."""
        self._bulk_delete_selected()

    def _bulk_delete_selected(self) -> None:
        """Hromadné mazání vybraných položek (otázky, podskupiny, skupiny)."""
        items = self.tree.selectedItems()
        if not items:
            QMessageBox.information(self, "Smazat", "Vyberte položky ke smazání.")
            return

        count = len(items)
        msg = f"Opravdu smazat {count} vybraných položek?\n(Včetně obsahu skupin/podskupin)"
        if QMessageBox.question(self, "Smazat vybrané", msg) != QMessageBox.Yes:
            return

        # Sběr IDček k smazání
        to_delete_q_ids = set()      # otázky
        to_delete_sg_ids = set()     # podskupiny
        to_delete_g_ids = set()      # skupiny

        for it in items:
            meta = it.data(0, Qt.UserRole) or {}
            kind = meta.get("kind")
            if kind == "question":
                to_delete_q_ids.add(meta.get("id"))
            elif kind == "subgroup":
                to_delete_sg_ids.add(meta.get("id"))
            elif kind == "group":
                to_delete_g_ids.add(meta.get("id"))

        # 1. Filtrace skupin (nejvyšší úroveň)
        # Pokud mažeme skupinu, zmizí vše pod ní, takže nemusíme řešit její podskupiny/otázky
        self.root.groups = [g for g in self.root.groups if g.id not in to_delete_g_ids]

        # 2. Procházení zbytku a mazání podskupin a otázek
        for g in self.root.groups:
            # Filtrace podskupin v této skupině
            g.subgroups = [sg for sg in g.subgroups if sg.id not in to_delete_sg_ids]
            
            # Rekurzivní čištění uvnitř podskupin (pro otázky a vnořené podskupiny)
            self._clean_subgroups_recursive(g.subgroups, to_delete_sg_ids, to_delete_q_ids)

        self._refresh_tree()
        self._clear_editor()
        self.save_data()
        self.statusBar().showMessage(f"Smazáno {count} položek.", 4000)

    def _clean_subgroups_recursive(self, subgroups: List[Subgroup], delete_sg_ids: set, delete_q_ids: set) -> None:
        """Pomocná metoda pro rekurzivní čištění."""
        for sg in subgroups:
            # Smazání otázek v aktuální podskupině
            if delete_q_ids:
                sg.questions = [q for q in sg.questions if q.id not in delete_q_ids]
            
            # Filtrace vnořených podskupin (pokud existují)
            if sg.subgroups:
                sg.subgroups = [s for s in sg.subgroups if s.id not in delete_sg_ids]
                # Rekurze do hloubky
                self._clean_subgroups_recursive(sg.subgroups, delete_sg_ids, delete_q_ids)

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
                self._set_question_editor_visible(True)
                self.rename_panel.hide()
                self._set_editor_enabled(True)
        elif kind in ("group", "subgroup"):
            name = ""
            if kind == "group":
                g = self._find_group(meta["id"]); name = g.name if g else ""
            else:
                sg = self._find_subgroup(meta["parent_group_id"], meta["id"]); name = sg.name if sg else ""
            
            self.rename_line.setText(name)
            self._set_question_editor_visible(False)
            self.rename_panel.show()
            self._set_editor_enabled(False) 
        else:
            # Pokud metoda _clear_editor existuje, zavoláme ji. 
            # Pokud ne, implementujte ji viz výše.
            self._clear_editor()
            self._set_question_editor_visible(False)
            self.rename_panel.hide()

    def _clear_editor(self) -> None:
        self._current_question_id = None
        self.text_edit.clear()
        self.spin_points.setValue(1)
        self.spin_bonus_correct.setValue(1.00)
        self.spin_bonus_wrong.setValue(0.00)
        self.combo_type.setCurrentIndex(0)
        self.title_edit.clear()
        self.edit_correct_answer.clear() # NOVÉ: vymazat i toto
        self.table_funny.setRowCount(0) # NOVÉ: vymazat i toto
        self._set_editor_enabled(False)


    def _set_question_editor_visible(self, visible: bool) -> None:
        """Zobrazí nebo skryje kompletní editor otázky (toolbar, formulář, text)."""
        self.editor_toolbar.setVisible(visible)
        self.text_edit.setVisible(visible)
        self.btn_save_question.setVisible(visible)
        
        # Skrytí/Zobrazení prvků formuláře
        widgets = [
            self.title_edit, 
            self.combo_type, 
            self.spin_points, 
            self.spin_bonus_correct, 
            self.spin_bonus_wrong,
            # NOVÉ:
            self.edit_correct_answer,
            self.funny_container
        ]
        
        for w in widgets:
            w.setVisible(visible)
            lbl = self.form_layout.labelForField(w)
            if lbl:
                lbl.setVisible(visible)
        
        if visible:
            self._on_type_changed_ui()


    def _load_question_to_editor(self, q: Question) -> None:
        self._current_question_id = q.id
        self.combo_type.setCurrentIndex(0 if q.type == "classic" else 1)
        self.spin_points.setValue(int(q.points))
        self.spin_bonus_correct.setValue(float(q.bonus_correct))
        self.spin_bonus_wrong.setValue(float(q.bonus_wrong))
        self.text_edit.setHtml(q.text_html or "<p><br></p>")
        self.title_edit.setText(q.title or self._derive_title_from_html(q.text_html))
        
        # NOVÉ: Načtení správné odpovědi
        self.edit_correct_answer.setPlainText(q.correct_answer or "")
        
        # NOVÉ: Načtení vtipných odpovědí
        self.table_funny.setRowCount(0)
        # Pojistka pro případ starého JSONu kde funny_answers může být None nebo chybět v __init__ (pokud by se nepoužil default)
        f_answers = getattr(q, "funny_answers", []) or []
        
        for fa in f_answers:
            # fa může být dict (z JSONu) nebo objekt FunnyAnswer
            text = fa.text if isinstance(fa, FunnyAnswer) else fa.get("text", "")
            date = fa.date if isinstance(fa, FunnyAnswer) else fa.get("date", "")
            author = fa.author if isinstance(fa, FunnyAnswer) else fa.get("author", "")
            
            row = self.table_funny.rowCount()
            self.table_funny.insertRow(row)
            self.table_funny.setItem(row, 0, QTableWidgetItem(text))
            self.table_funny.setItem(row, 1, QTableWidgetItem(date))
            self.table_funny.setItem(row, 2, QTableWidgetItem(author))

        self._set_editor_enabled(True)
        
        # Synchronizace viditelnosti polí podle načteného typu
        self._on_type_changed_ui()

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
                        
                        # Uložení bodů
                        if q.type == "classic":
                            q.points = int(self.spin_points.value()); q.bonus_correct = 0.0; q.bonus_wrong = 0.0
                        else:
                            q.points = 0; q.bonus_correct = round(float(self.spin_bonus_correct.value()), 2); q.bonus_wrong = round(float(self.spin_bonus_wrong.value()), 2)
                        
                        # NOVÉ: Uložení správné odpovědi
                        q.correct_answer = self.edit_correct_answer.toPlainText()
                        
                        # NOVÉ: Uložení vtipných odpovědí z tabulky
                        new_funny = []
                        for r in range(self.table_funny.rowCount()):
                            t = self.table_funny.item(r, 0).text()
                            d = self.table_funny.item(r, 1).text()
                            a = self.table_funny.item(r, 2).text()
                            new_funny.append(FunnyAnswer(text=t, date=d, author=a))
                        q.funny_answers = new_funny

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

    # -------------------- Vyhledávače --------------------

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

    # -------------------- Formátování Rich text --------------------

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

        # Skrytí/zobrazení polí a jejich popisků
        self.spin_points.setVisible(is_classic)
        lbl_points = self.form_layout.labelForField(self.spin_points)
        if lbl_points: 
            lbl_points.setVisible(is_classic)

        self.spin_bonus_correct.setVisible(not is_classic)
        lbl_bonus_correct = self.form_layout.labelForField(self.spin_bonus_correct)
        if lbl_bonus_correct: 
            lbl_bonus_correct.setVisible(not is_classic)

        self.spin_bonus_wrong.setVisible(not is_classic)
        lbl_bonus_wrong = self.form_layout.labelForField(self.spin_bonus_wrong)
        if lbl_bonus_wrong: 
            lbl_bonus_wrong.setVisible(not is_classic)

        self._autosave_schedule()


    # -------------------- Výběr datového souboru --------------------

    def _choose_data_file(self) -> None:
        new_path, _ = QFileDialog.getSaveFileName(self, "Zvolit/uložit JSON s otázkami", str(self.data_path), "JSON (*.json)")
        if new_path:
            self.data_path = Path(new_path)
            self.statusBar().showMessage(f"Datový soubor změněn na: {self.data_path}", 4000)
            self.load_data(); self._refresh_tree()

    # -------------------- Filtr --------------------

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
        if 'docx' not in sys.modules:
            QMessageBox.critical(self, "Chyba", "Knihovna python-docx není nainstalována.")
            return []

        doc = docx.Document(path)
        out = []

        # Cache pro numbering definitions
        numbering_cache = {} # numId -> (abstractId, {ilvl: fmt})

        # Přednačtení definic z XML, abychom to nemuseli lovit per paragraph
        try:
            numbering_part = doc.part.numbering_part
            if numbering_part:
                # Mapování abstractNumId -> {ilvl: fmt}
                abstract_formats = {}
                for abstract_id, abstract in numbering_part.numbering_definitions._abstract_nums.items():
                    levels = {}
                    for lvl in abstract.levels:
                        levels[lvl.ilvl] = lvl.num_fmt
                    abstract_formats[abstract_id] = levels
                
                # Mapování numId -> abstractNumId
                for num_id, num in numbering_part.numbering_definitions._nums.items():
                    if num.abstractNumId in abstract_formats:
                        numbering_cache[num_id] = (num.abstractNumId, abstract_formats[num.abstractNumId])
        except Exception:
            pass # Pokud selže přístup k internals, pojedeme bez formátů

        def get_numbering(p: Paragraph):
            p_element = p._p
            pPr = p_element.pPr
            if pPr is None: return None, None, None
            numPr = pPr.numPr
            if numPr is None: return None, None, None
            
            val_ilvl = numPr.ilvl.val if numPr.ilvl is not None else 0
            val_numId = numPr.numId.val if numPr.numId is not None else None
            
            fmt = "decimal" # default fallback
            
            if val_numId in numbering_cache:
                _, levels = numbering_cache[val_numId]
                if val_ilvl in levels:
                    fmt = levels[val_ilvl]
            
            return val_numId, val_ilvl, fmt

        for p in doc.paragraphs:
            txt = p.text.strip()
            numId, ilvl, fmt = get_numbering(p)
            is_numbered = (numId is not None)
            
            out.append({
                "text": txt,
                "is_numbered": is_numbered,
                "ilvl": ilvl,
                "num_fmt": fmt,
                "num_id": numId # Ukládáme si i ID seznamu pro detekci změny kontextu
            })
        
        return out

    def _parse_questions_from_paragraphs(self, paragraphs: List[dict]) -> List[Question]:
        out: List[Question] = []
        i = 0
        n = len(paragraphs)

        rx_bonus = re.compile(r'^\s*Otázka\s+\d+.*BONUS', re.IGNORECASE)
        rx_question_start_text = re.compile(r'^[A-ZŽŠČŘĎŤŇ].*[\?\.]$') # Začíná velkým, končí ? nebo .
        rx_not_question_start = re.compile(r'^(Slovník|Tabulka|Obrázek|Příklad|Body|Poznámka)', re.IGNORECASE)

        def html_escape(s: str) -> str:
            return _html.escape(s or "")

        def wrap_list(items: List[tuple[str, int, str]]) -> str:
            if not items: return ""
            fmt = items[0][2] or "decimal"
            is_bullet = (fmt == "bullet")
            tag_open = "<ul>" if is_bullet else "<ol>"
            tag_close = "</ul>" if is_bullet else "</ol>"
            lis = "".join(f"<li>{html_escape(t)}</li>" for (t, _, _) in items if t.strip())
            return f"{tag_open}{lis}{tag_close}"

        # Sledujeme poslední numId hlavní otázky, abychom poznali změnu seznamu
        last_question_num_id = None

        while i < n:
            p = paragraphs[i]
            txt = p["text"]
            if not txt: 
                i += 1; continue

            # 1. BONUS
            if rx_bonus.search(txt):
                block_html = f"<p><b>{html_escape(txt)}</b></p>"
                j = i + 1
                while j < n:
                    next_p = paragraphs[j]
                    next_txt = next_p["text"]
                    if rx_bonus.search(next_txt): break
                    # Pokud narazíme na HLAVNÍ otázku (číslovaná, level 0, decimal, a vypadá jako otázka)
                    if next_p["is_numbered"] and next_p["ilvl"] == 0 and next_p["num_fmt"] != "bullet" and not rx_not_question_start.match(next_txt):
                         break
                    if next_txt: block_html += f"<p>{html_escape(next_txt)}</p>"
                    j += 1
                q = Question.new_default("bonus")
                q.title = self._derive_title_from_html(block_html, prefix="BONUS: ")
                q.text_html = block_html
                q.bonus_correct, q.bonus_wrong = 1.0, 0.0
                out.append(q); i = j; continue

            # 2. KLASICKÁ
            # Podmínky pro novou otázku:
            # a) Je číslovaná
            # b) Je na levelu 0
            # c) Není to bullet
            # d) Není to explicitně vyloučený text (Slovník...)
            is_num = p["is_numbered"]
            ilvl = p["ilvl"]
            fmt = p["num_fmt"]
            nid = p["num_id"]

            is_potential_question = (is_num and (ilvl == 0 or ilvl is None) and fmt != "bullet")
            
            # Heuristika pro "Slovník" problém:
            # Pokud se změnilo numId (oproti minulé otázce) a text nevypadá jako otázka (začíná na Slovník),
            # tak to pravděpodobně NENÍ nová otázka, ale součást minulé (pokud nějaká byla).
            if is_potential_question and nid != last_question_num_id and last_question_num_id is not None:
                if rx_not_question_start.match(txt):
                    is_potential_question = False

            if is_potential_question:
                last_question_num_id = nid # Uložíme si ID seznamu této otázky
                
                block_html = f"<p>{html_escape(txt)}</p>"
                j = i + 1
                list_buffer = []
                
                while j < n:
                    next_p = paragraphs[j]
                    next_txt = next_p["text"]
                    
                    # Check start of new element (Bonus or New Question)
                    if rx_bonus.search(next_txt): break
                    
                    next_is_num = next_p["is_numbered"]
                    next_ilvl = next_p["ilvl"]
                    next_fmt = next_p["num_fmt"]
                    next_nid = next_p["num_id"]

                    # Je to začátek další otázky?
                    if next_is_num and (next_ilvl == 0 or next_ilvl is None) and next_fmt != "bullet":
                        # Výjimka: Pokud je to "Slovník..." (tedy změna numId, ale textově to není otázka),
                        # tak to NENÍ nová otázka, ale pokračování této.
                        is_really_new = True
                        if next_nid != last_question_num_id:
                             if rx_not_question_start.match(next_txt):
                                 is_really_new = False
                        
                        if is_really_new:
                            break
                    
                    if not next_txt: j += 1; continue

                    # Je to list item?
                    # - Buď ilvl > 0
                    # - Nebo fmt == bullet
                    # - Nebo ilvl == 0, ale je to ten "Slovník" případ (is_really_new=False výše propadne sem)
                    is_list_item = False
                    if next_is_num:
                        if next_ilvl > 0 or next_fmt == "bullet":
                            is_list_item = True
                        elif next_nid != last_question_num_id and rx_not_question_start.match(next_txt):
                             # To je ten případ "Slovník" na levelu 0
                             is_list_item = True

                    if is_list_item:
                        list_buffer.append((next_txt, next_ilvl, next_fmt))
                    else:
                        if list_buffer:
                            block_html += wrap_list(list_buffer)
                            list_buffer = []
                        block_html += f"<p>{html_escape(next_txt)}</p>"
                    j += 1
                
                if list_buffer:
                    block_html += wrap_list(list_buffer)
                
                q = Question.new_default("classic")
                q.title = self._derive_title_from_html(block_html)
                q.text_html = block_html
                q.points = 1
                out.append(q); i = j; continue

            # Text, který není součástí žádné otázky (úvod atd.)
            i += 1

        return out


    def _ensure_unassigned_group(self) -> tuple[str, Optional[str]]:
        """Zajistí existenci skupiny 'Neroztříděné'. Vrací (group_id, None)."""
        name = "Neroztříděné"
        g = next((g for g in self.root.groups if g.name == name), None)
        if not g:
            g = Group(id=str(_uuid.uuid4()), name=name, subgroups=[])
            self.root.groups.append(g)
        # Nevytváříme "Default" podskupinu automaticky, pokud není potřeba.
        # V importu si vytvoříme "Klasické" a "Bonusové" specificky.
        return g.id, None


    def _import_from_docx(self) -> None:
        # Výchozí složka pro import
        import_dir = self.project_root / "data" / "Staré písemky"
        import_dir.mkdir(parents=True, exist_ok=True)

        paths, _ = QFileDialog.getOpenFileNames(self, "Import z DOCX", str(import_dir), "Word dokument (*.docx)")
        if not paths:
            return

        # 1. Získání cílových podskupin v "Neroztříděné"
        #    (Použijeme _ensure_unassigned_group pro získání/vytvoření hlavní skupiny,
        #    ale pak si ručně najdeme/vytvoříme specifické podskupiny.)
        g_id, _ = self._ensure_unassigned_group()
        unassigned_group = self._find_group(g_id)
        if not unassigned_group:
            # Fallback, nemělo by nastat
            return

        def get_or_create_subgroup(g: Group, name: str) -> Subgroup:
            sg = next((s for s in g.subgroups if s.name == name), None)
            if not sg:
                sg = Subgroup(id=str(_uuid.uuid4()), name=name, subgroups=[], questions=[])
                g.subgroups.append(sg)
            return sg

        target_classic = get_or_create_subgroup(unassigned_group, "Klasické")
        target_bonus = get_or_create_subgroup(unassigned_group, "Bonusové")

        # 2. Vytvoření indexu existujících otázek pro kontrolu duplicit
        #    Jako klíč použijeme ostripovaný HTML obsah.
        existing_hashes = set()

        def index_questions(node):
            # Rekurzivně projít strom
            if isinstance(node, RootData):
                for gr in node.groups: index_questions(gr)
            elif isinstance(node, Group):
                for sgr in node.subgroups: index_questions(sgr)
            elif isinstance(node, Subgroup):
                for q in node.questions:
                    if q.text_html:
                        existing_hashes.add(q.text_html.strip())
                for sub in node.subgroups:
                    index_questions(sub)

        index_questions(self.root)

        total_imported = 0
        total_duplicates = 0

        for p in paths:
            try:
                paras = self._extract_paragraphs_from_docx(Path(p))
                qs = self._parse_questions_from_paragraphs(paras)
                
                if not qs:
                    # Info, ale nepovažujeme za chybu přerušující ostatní soubory
                    continue

                file_imported_count = 0
                
                for q in qs:
                    content_hash = (q.text_html or "").strip()
                    
                    # Kontrola duplicit
                    if content_hash in existing_hashes:
                        total_duplicates += 1
                        continue
                    
                    # Pokud není duplicitní, přidáme do DB a aktualizujeme hashset (proti duplicitám v rámci jednoho importu)
                    existing_hashes.add(content_hash)
                    
                    if q.type == "classic":
                        target_classic.questions.append(q)
                    else:
                        target_bonus.questions.append(q)
                    
                    file_imported_count += 1

                total_imported += file_imported_count

            except Exception as e:
                QMessageBox.warning(self, "Import – chyba", f"Soubor: {p}\n{e}")

        self._refresh_tree()
        self.save_data()

        msg = f"Import dokončen.\n\nÚspěšně importováno: {total_imported}\nDuplicitních (přeskočeno): {total_duplicates}"
        QMessageBox.information(self, "Výsledek importu", msg)


    # -------------------- Přesun otázky --------------------

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

    # -------------------- Export DOCX --------------------

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

    def _export_docx_wizard(self):
        # PŮVODNÍ (ŠPATNĚ):
        # wiz = ExportWizard(self)
        # wiz.le_output.setText(str(self.project_root / "test_vystup.docx")) <--- TOTO SMAZAT!
        # wiz.exec()

        # NOVÉ (SPRÁVNĚ):
        wiz = ExportWizard(self)
        wiz.exec()


    def _generate_docx_from_template(self, template_path: Path, output_path: Path,
                                     simple_repl: Dict[str, str], rich_repl_html: Dict[str, str]) -> None:
        
        try:
            doc = docx.Document(template_path)
        except Exception as e:
            QMessageBox.critical(self, "Export chyba", f"Nelze otevřít šablonu pomocí python-docx:\n{e}")
            return

        def insert_rich_question_block(paragraph, html_content):
            paras_data = parse_html_to_paragraphs(html_content)
            if not paras_data: 
                paragraph.clear()
                return
            
            p_insert = paragraph._p
            
            for i, p_data in enumerate(paras_data):
                if i == 0:
                    new_p = paragraph
                    new_p.clear()
                else:
                    new_p = doc.add_paragraph()
                    p_insert.addnext(new_p._p)
                    p_insert = new_p._p

                # -- ZAROVNÁNÍ --
                align = p_data.get('align', 'left')
                if align == 'center': new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif align == 'right': new_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif align == 'justify': new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                else: new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                # -- MEZEROVÁNÍ (Spacing) --
                # Nastavíme minimální mezery, aby to vypadalo jako v editoru
                new_p.paragraph_format.space_before = Pt(0)
                new_p.paragraph_format.space_after = Pt(0)
                # Případně line_spacing_rule (single)
                # new_p.paragraph_format.line_spacing = 1.0 

                # Prefix
                if p_data.get('prefix'):
                    new_p.paragraph_format.left_indent = Pt(48)
                    new_p.paragraph_format.first_line_indent = Pt(-24)
                    new_p.add_run(p_data['prefix'])
                
                # Runs
                for r_data in p_data['runs']:
                    text_content = r_data['text']
                    
                    # Řešení <br> -> \n
                    # Pokud text obsahuje \n, musíme ho rozdělit a vložit breaky
                    parts = text_content.split('\n')
                    for idx, part in enumerate(parts):
                        if part:
                            run = new_p.add_run(part)
                            if r_data.get('b'): run.bold = True
                            if r_data.get('i'): run.italic = True
                            if r_data.get('u'): run.underline = True
                            if r_data.get('color'):
                                try:
                                    rgb = r_data['color']
                                    run.font.color.rgb = RGBColor(int(rgb[:2], 16), int(rgb[2:4], 16), int(rgb[4:], 16))
                                except: pass
                        
                        # Pokud to není poslední část, znamená to, že následoval \n -> Soft Break
                        if idx < len(parts) - 1:
                            run = new_p.add_run()
                            run.add_break() # Shift+Enter ve Wordu


        # -- Helper: Zpracování jednoho odstavce (Inline i Block) --
        def process_paragraph(p):
            full_text = p.text
            # Ignorujeme prázdné (pokud neobsahují obrázky, ale my řešíme text)
            if not full_text.strip(): return

            # 1. BLOCK CHECK (Je odstavec POUZE placeholderem?)
            # Ořízneme whitespace pro porovnání
            txt_clean = full_text.strip()
            matched_rich = None
            for ph, html in rich_repl_html.items():
                if txt_clean == f"<{ph}>" or txt_clean == f"{{{ph}}}":
                    matched_rich = (ph, html)
                    break
            
            if matched_rich:
                insert_rich_question_block(p, matched_rich[1])
                return

            # 2. INLINE CHECK (Obsahuje odstavec placeholder uvnitř textu?)
            # Kombinace Simple (Plain) a Rich (Formatted) nahrazování
            
            # Zjistíme, zda odstavec obsahuje nějaký klíč
            keys_found = []
            
            # Simple keys
            for k in simple_repl.keys():
                if f"<{k}>" in full_text or f"{{{k}}}" in full_text:
                    keys_found.append(k)
            # Rich keys (Bonusy)
            for k in rich_repl_html.keys():
                if f"<{k}>" in full_text or f"{{{k}}}" in full_text:
                    keys_found.append(k)
            
            if not keys_found:
                return

            # Pokud jsme něco našli, musíme odstavec "přeskládat".
            # Text rozdělíme na segmenty.
            # Segment je buď string (zachovat) nebo dict (nahradit).
            
            segments = [full_text]
            
            # Aplikujeme rozdělení pro všechny nalezené klíče
            all_repl_data = {}
            # Simple data
            for k, v in simple_repl.items(): 
                all_repl_data[k] = {'type': 'simple', 'val': v}
            # Rich data
            for k, html in rich_repl_html.items(): 
                all_repl_data[k] = {'type': 'rich', 'val': html}
            
            for k in keys_found:
                info = all_repl_data[k]
                # Hledáme <Klic> i {Klic}
                tokens = [f"<{k}>", f"{{{k}}}"]
                
                for token in tokens:
                    new_segments = []
                    for seg in segments:
                        if isinstance(seg, str):
                            # Split string by token
                            parts = seg.split(token)
                            for i, part in enumerate(parts):
                                if part: new_segments.append(part) # Zachováme text
                                if i < len(parts) - 1:
                                    # Vložíme placeholder pro nahrazení
                                    new_segments.append(info)
                        else:
                            new_segments.append(seg) # Už zpracovaný segment
                    segments = new_segments

            # -- APLIKACE ZMĚN --
            # Uchováme styl prvního runu pro "dědění" stylu (font, size)
            base_font_name = None
            base_font_size = None
            base_bold = None
            
            if p.runs:
                r0 = p.runs[0]
                base_font_name = r0.font.name
                base_font_size = r0.font.size
                base_bold = r0.bold # Použijeme jen pro simple text

            p.clear() # Smažeme starý obsah
            
            for seg in segments:
                if isinstance(seg, str):
                    run = p.add_run(seg)
                    if base_font_name: run.font.name = base_font_name
                    if base_font_size: run.font.size = base_font_size
                    # run.bold = base_bold # Nechceme vynucovat bold na okolní text, pokud nebyl
                    
                elif isinstance(seg, dict):
                    val = seg['val']
                    if seg['type'] == 'simple':
                        # Simple text (Datum)
                        run = p.add_run(val)
                        if base_font_name: run.font.name = base_font_name
                        if base_font_size: run.font.size = base_font_size
                        if base_bold is not None: run.bold = base_bold # Datum zdědí tučnost
                        
                    elif seg['type'] == 'rich':
                        # Inline Rich Text (Bonus)
                        # Rozparsujeme HTML na runy
                        paras = parse_html_to_paragraphs(val)
                        
                        # Vložíme runy.
                        for p_idx, p_data in enumerate(paras):
                            # Pokud to není první "odstavec" z HTML, vložíme zalomení řádku (Soft Break)
                            if p_idx > 0:
                                run_break = p.add_run()
                                run_break.add_break()
                            
                            for r_data in p_data['runs']:
                                text_content = r_data['text']
                                # I uvnitř runu může být \n (z <br>)
                                parts = text_content.split('\n')
                                for idx_part, part in enumerate(parts):
                                    if part:
                                        run = p.add_run(part)
                                        if r_data.get('b'): run.bold = True
                                        if r_data.get('i'): run.italic = True
                                        if r_data.get('u'): run.underline = True
                                        if r_data.get('color'):
                                            try:
                                                rgb = r_data['color']
                                                run.font.color.rgb = RGBColor(int(rgb[:2], 16), int(rgb[2:4], 16), int(rgb[4:], 16))
                                            except: pass
                                        
                                        # Base style
                                        if base_font_name: run.font.name = base_font_name
                                        if base_font_size: run.font.size = base_font_size
                                    
                                    if idx_part < len(parts) - 1:
                                        p.add_run().add_break()



        # 1. Body
        for p in doc.paragraphs:
            process_paragraph(p)
            
        # 2. Tables in Body
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        process_paragraph(p)
        
        # 3. Headers / Footers
        for section in doc.sections:
            for h in [section.header, section.first_page_header]:
                if h:
                    for p in h.paragraphs: process_paragraph(p)
                    for t in h.tables:
                        for r in t.rows:
                            for c in r.cells:
                                for p in c.paragraphs: process_paragraph(p)
            for f in [section.footer, section.first_page_footer]:
                if f:
                    for p in f.paragraphs: process_paragraph(p)
                    for t in f.tables:
                        for r in t.rows:
                            for c in r.cells:
                                for p in c.paragraphs: process_paragraph(p)

        try:
            doc.save(output_path)
        except Exception as e:
            QMessageBox.critical(self, "Chyba uložení", f"Nelze uložit DOCX:\n{e}")



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

    # Nastavení ikony na úrovni aplikace (pokud existuje)
    project_root = Path.cwd()
    icon_file = project_root / "icon" / "icon.png"
    if icon_file.exists():
        app.setWindowIcon(QIcon(str(icon_file)))

    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
