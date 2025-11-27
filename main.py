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
    QIcon
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
    QTreeWidgetItemIterator
)

APP_NAME = "Crypto Exam Generator"
APP_VERSION = "1.8c"

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
        self.resize(1100, 800)

        self.template_path: Optional[Path] = None
        self.output_path: Optional[Path] = None
        
        # Placeholdery načtené ze šablony
        self.placeholders_q: List[str] = [] 
        self.placeholders_b: List[str] = [] 
        self.has_datumcas = False
        self.has_pozn = False
        self.has_minmax = (False, False)

        # Výběr uživatele {placeholder: question_id}
        self.selection_map: Dict[str, str] = {} 

        # Default paths
        self.templates_dir = self.owner.project_root / "data" / "Šablony"
        self.output_dir = self.owner.project_root / "data" / "Vygenerované testy"
        self.templates_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.default_template = self.templates_dir / "template_AK3KR_předtermín.docx"
        self.default_output = self.output_dir / f"test_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

        # --- BUILD PAGES ---
        self._build_page1()
        self._build_page2()
        self._build_page3()
        
        # Auto-load
        if self.default_template.exists():
            self.le_template.setText(str(self.default_template))
            self.template_path = self.default_template
            QTimer.singleShot(100, self._scan_placeholders)
        
        self.le_output.setText(str(self.default_output))
        self.output_path = self.default_output

    def _build_page1(self):
        self.page1 = QWizardPage()
        self.page1.setTitle("Krok 1: Výběr šablony a globální nastavení")
        l1 = QVBoxLayout(self.page1)
        
        # GroupBox: Soubory
        gb_files = QGroupBox("Soubory")
        form_files = QFormLayout()
        self.le_template = QLineEdit()
        btn_t = QPushButton("Vybrat šablonu..."); btn_t.clicked.connect(self._choose_template)
        h_t = QHBoxLayout(); h_t.addWidget(self.le_template); h_t.addWidget(btn_t)
        form_files.addRow("Šablona:", h_t)
        
        self.le_output = QLineEdit()
        btn_o = QPushButton("Cíl exportu..."); btn_o.clicked.connect(self._choose_output)
        h_o = QHBoxLayout(); h_o.addWidget(self.le_output); h_o.addWidget(btn_o)
        form_files.addRow("Výstup:", h_o)
        gb_files.setLayout(form_files)
        l1.addWidget(gb_files)
        
        # GroupBox: Parametry
        gb_params = QGroupBox("Parametry testu")
        form_params = QFormLayout()
        self.le_prefix = QLineEdit("MůjTest")
        form_params.addRow("Prefix verze:", self.le_prefix)
        
        self.dt_edit = QDateTimeEdit(QDateTime.currentDateTime())
        self.dt_edit.setDisplayFormat("dd.MM.yyyy HH:mm")
        self.dt_edit.setCalendarPopup(True)
        form_params.addRow("Datum testu:", self.dt_edit)
        gb_params.setLayout(form_params)
        l1.addWidget(gb_params)
        
        self.lbl_scan_info = QLabel("Info: Čekám na načtení šablony...")
        self.lbl_scan_info.setStyleSheet("color: gray; font-style: italic;")
        l1.addWidget(self.lbl_scan_info)
        
        self.le_template.textChanged.connect(self._on_templ_change)
        self.le_output.textChanged.connect(lambda t: setattr(self, 'output_path', Path(t)))
        self.addPage(self.page1)

    def _build_page2(self):
        self.page2 = QWizardPage()
        self.page2.setTitle("Krok 2: Přiřazení otázek do šablony")
        l2 = QHBoxLayout(self.page2)
        
        # Levý panel: Dostupné otázky (Tree)
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("<b>Dostupné otázky v databázi:</b>"))
        self.tree_source = QTreeWidget()
        self.tree_source.setHeaderLabels(["Název otázky", "Typ", "Body"])
        self.tree_source.setColumnWidth(0, 280)
        left_layout.addWidget(self.tree_source)
        l2.addLayout(left_layout, 4) # Ratio 4
        
        # Pravý panel: Sloty
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("<b>Sloty v šabloně:</b>"))
        
        self.scroll_slots = QScrollArea()
        self.scroll_slots.setWidgetResizable(True)
        self.widget_slots = QWidget()
        self.layout_slots = QVBoxLayout(self.widget_slots)
        self.layout_slots.setSpacing(6)
        self.layout_slots.addStretch()
        self.scroll_slots.setWidget(self.widget_slots)
        
        right_layout.addWidget(self.scroll_slots)
        l2.addLayout(right_layout, 6) # Ratio 6
        
        self.addPage(self.page2)
        self.page2.initializePage = self._init_page2

    def _build_page3(self):
        self.page3 = QWizardPage()
        self.page3.setTitle("Krok 3: Kontrola a Export")
        l3 = QVBoxLayout(self.page3)
        
        self.lbl_summary = QLabel()
        self.lbl_summary.setWordWrap(True)
        l3.addWidget(self.lbl_summary)
        
        self.tbl_preview = QTableWidget()
        self.tbl_preview.setColumnCount(2)
        self.tbl_preview.setHorizontalHeaderLabels(["Položka", "Hodnota"])
        self.tbl_preview.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        l3.addWidget(self.tbl_preview)
        
        self.addPage(self.page3)
        self.page3.initializePage = self._init_page3

    # --- Logic Krok 1 ---
    def _choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Šablona", str(self.templates_dir), "*.docx")
        if path: self.le_template.setText(path)

    def _choose_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Výstup", str(self.default_output), "*.docx")
        if path: self.le_output.setText(path)

    def _on_templ_change(self, text):
        path = Path(text)
        if path.exists() and path.suffix == '.docx':
            self.template_path = path
            self._scan_placeholders()
        else:
            self.template_path = None
            self.lbl_scan_info.setText("Šablona neexistuje.")

    def _scan_placeholders(self):
        try:
            doc = docx.Document(self.template_path)
            full_text = ""
            for p in doc.paragraphs: full_text += p.text + "\n"
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        for p in c.paragraphs: full_text += p.text + "\n"
            
            import re
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
            self.lbl_scan_info.setText(msg)
            
        except Exception as e:
            self.lbl_scan_info.setText(f"Chyba čtení šablony: {e}")

    # --- Logic Krok 2 ---
    def _init_page2(self):
        self.tree_source.clear()
        while self.layout_slots.count():
            item = self.layout_slots.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.layout_slots.addStretch()
        
        # Load Tree
        groups = self.owner.root.groups
        for g in groups:
            g_item = QTreeWidgetItem([g.name, "", ""])
            g_item.setExpanded(True)
            self.tree_source.addTopLevelItem(g_item)
            for sg in g.subgroups:
                sg_item = QTreeWidgetItem([sg.name, "", ""])
                sg_item.setExpanded(True)
                g_item.addChild(sg_item)
                for q in sg.questions:
                    label = q.title
                    # Pokud už je vybrána, označíme ji
                    if q.id in self.selection_map.values():
                        label += " (VYBRÁNO)"
                    
                    q_item = QTreeWidgetItem([label, q.type, str(q.points) if q.type=='classic' else f"B:{q.bonus_correct}"])
                    q_item.setData(0, Qt.UserRole, q.id)
                    
                    # Pokud je vybrána, skryjeme ji (aby nešla vybrat znovu)
                    if q.id in self.selection_map.values():
                        q_item.setHidden(True)
                    
                    sg_item.addChild(q_item)

        # Load Slots
        if self.placeholders_q:
            lbl = QLabel("--- KLASICKÉ OTÁZKY ---"); lbl.setStyleSheet("font-weight:bold; color:#4da6ff; margin-top:5px;")
            self.layout_slots.insertWidget(self.layout_slots.count()-1, lbl)
            for ph in self.placeholders_q:
                self._add_slot_widget(ph, 'classic')

        if self.placeholders_b:
            lbl = QLabel("--- BONUSOVÉ OTÁZKY ---"); lbl.setStyleSheet("font-weight:bold; color:#ffcc00; margin-top:10px;")
            self.layout_slots.insertWidget(self.layout_slots.count()-1, lbl)
            for ph in self.placeholders_b:
                self._add_slot_widget(ph, 'bonus')
    
    def _add_slot_widget(self, placeholder_name, allowed_type):
        w = QWidget()
        l = QHBoxLayout(w)
        l.setContentsMargins(0,2,0,2)
        
        lbl_ph = QLabel(f"{placeholder_name}:")
        lbl_ph.setFixedWidth(80)
        
        btn_sel = QPushButton("Vybrat...")
        btn_sel.setFixedWidth(80)
        
        # Zjistíme, zda už je něco vybráno (z paměti)
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
        l.addWidget(lbl_val, 1) # Stretch
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
            
            # Pokud už tam něco bylo, vrátíme to do stromu
            old_qid = self.selection_map.get(placeholder_name)
            if old_qid:
                self._show_tree_item(old_qid)
            
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
        # Najít item v tree a odkrýt ho
        it = QTreeWidgetItemIterator(self.tree_source)
        while it.value():
            item = it.value()
            if item.data(0, Qt.UserRole) == qid:
                item.setHidden(False)
                break
            it += 1

    # --- Logic Krok 3 ---
    def _init_page3(self):
        self.tbl_preview.setRowCount(0)
        total_points = 0.0
        min_loss = 0.0
        
        # Tabulka
        row = 0
        for ph in self.placeholders_q + self.placeholders_b:
            qid = self.selection_map.get(ph)
            self.tbl_preview.insertRow(row)
            self.tbl_preview.setItem(row, 0, QTableWidgetItem(ph))
            
            if qid:
                q = self.owner._find_question_by_id(qid)
                if q.type == 'classic':
                    total_points += float(q.points)
                else:
                    total_points += float(q.bonus_correct)
                    min_loss += float(q.bonus_wrong)
                self.tbl_preview.setItem(row, 1, QTableWidgetItem(q.title))
            else:
                 self.tbl_preview.setItem(row, 1, QTableWidgetItem("---"))
            row += 1
            
        # Stupnice
        scale_text = f"""
        MaxBody: {total_points:.2f}
        MinBody: {min_loss:.2f}
        
        A: <9.2; {total_points:.2f}>
        B: <8.4; 9.2)
        C: <7.6; 8.4)
        D: <6.8; 7.6)
        E: <6.0; 6.8)
        F: <{min_loss:.2f}; 6.0)
        """
        self.lbl_summary.setText(scale_text)

    def accept(self) -> None:
        if not self.template_path or not self.output_path: return

        # Build Replacements
        repl_plain: Dict[str, str] = {}
        
        # Datum
        dt = round_dt_to_10m(self.dt_edit.dateTime())
        repl_plain["DatumČas"] = f"{cz_day_of_week(dt.toPython())} {dt.toString('dd.MM.yyyy HH:mm')}"
        repl_plain["DatumCas"] = repl_plain["DatumČas"] # alias
        repl_plain["DATUMCAS"] = repl_plain["DatumČas"]
        
        # Pozn
        prefix = self.le_prefix.text()
        today = datetime.now().strftime("%Y-%m-%d")
        repl_plain["PoznamkaVerze"] = f"{prefix} {today}_{str(_uuid.uuid4())[:8]}"
        repl_plain["POZNAMKAVERZE"] = repl_plain["PoznamkaVerze"]
        
        # Body
        # Spočítáme znovu
        total_points = 0.0
        min_loss = 0.0
        for qid in self.selection_map.values():
            q = self.owner._find_question_by_id(qid)
            if q.type == 'classic': total_points += float(q.points)
            else:
                total_points += float(q.bonus_correct)
                min_loss += float(q.bonus_wrong)
        
        repl_plain["MaxBody"] = f"{total_points:.2f}"
        repl_plain["MAXBODY"] = f"{total_points:.2f}"
        repl_plain["MinBody"] = f"{min_loss:.2f}"
        repl_plain["MINBODY"] = f"{min_loss:.2f}"

        # Rich Map
        rich_map: Dict[str, str] = {}
        for ph, qid in self.selection_map.items():
            q = self.owner._find_question_by_id(qid)
            if q:
                rich_map[ph] = q.text_html

        try:
            self.owner._generate_docx_from_template(self.template_path, self.output_path, repl_plain, rich_map)
        except Exception as e:
            QMessageBox.critical(self, "Export", f"Chyba:\n{e}"); return

        QMessageBox.information(self, "Export", f"Uloženo:\n{self.output_path}")
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
        self.spin_points.setEnabled(is_classic)
        self.spin_bonus_correct.setEnabled(not is_classic)
        self.spin_bonus_wrong.setEnabled(not is_classic)
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
                    next_fmt = paragraphs[j].get("num_fmt") or "decimal"
                    if not next_txt.strip() or is_noise(next_txt):
                        j += 1; continue
                    if (next_isnum and (next_ilvl is None or next_ilvl == 0)) or rx_bonus_start.match(next_txt) or rx_classic_numtxt.match(next_txt) or is_question_like(next_txt):
                        break
                    if next_isnum:
                        list_buffer.append((next_txt.strip(), next_ilvl or 0, next_fmt)); j += 1; continue
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

    def _import_from_docx(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "Import z DOCX", str(self.project_root), "Word dokument (*.docx)")
        if not paths:
            return
        g_id, sg_id = self._ensure_unassigned_group()
        target_sg = self._find_subgroup(g_id, sg_id)
        total = 0
        for p in paths:
            try:
                paras = self._extract_paragraphs_from_docx(Path(p))
                qs = self._parse_questions_from_paragraphs(paras)
                if not qs:
                    QMessageBox.information(self, "Import", f"{p}\nNebyla nalezena žádná otázka.")
                    continue
                if target_sg is not None:
                    target_sg.questions.extend(qs)
                total += len(qs)
            except Exception as e:
                QMessageBox.warning(self, "Import – chyba", f"Soubor: {p}\n{e}")
        self._refresh_tree()
        self.save_data()
        if total:
            self.statusBar().showMessage(f"Import hotov: {total} otázek do 'Neroztříděné'.", 6000)

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

    def _export_docx_wizard(self) -> None:
        wiz = ExportWizard(self); wiz.le_output.setText(str(self.project_root / "test_vystup.docx")); wiz.exec()

    def _generate_docx_from_template(self, template_path: Path, output_path: Path,
                                     simple_repl: Dict[str, str], rich_repl_html: Dict[str, str]) -> None:
        
        print(f"DEBUG: Start exportu pomocí python-docx (v3.6 Inline). Template: {template_path}")
        
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
            print(f"DEBUG: Uloženo do {output_path}")
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
