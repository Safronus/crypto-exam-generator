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

import subprocess

import json
import sys
import uuid as _uuid
import re
import os
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

from PySide6.QtCore import Qt, QSize, QSaveFile, QByteArray, QTimer, QDateTime, QPoint
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
    QPixmap, QPainter, QIcon, QBrush
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
    QHeaderView, QCheckBox, QGridLayout,
    QTreeWidgetItemIterator, QButtonGroup,
    QHeaderView, QMenu, QTabWidget, QRadioButton,
    QTreeWidget, QTreeWidgetItem, QSizePolicy
)

APP_NAME = "Crypto Exam Generator"
APP_VERSION = "6.7.2"

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
    # Nové pole – uložený zdrojový dokument (cesta k souboru, nebo prázdný string)
    source_doc: str = ""

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
    palette.setColor(QPalette.ToolTipBase, Qt.black)
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
    Podporuje vnořené styly, seznamy a odsazení (indentation).
    """
    def __init__(self) -> None:
        super().__init__()
        self.paragraphs: List[dict] = []
        self._stack: List[dict] = [] 
        self._list_stack: List[dict] = [] 
        self._current_runs: List[dict] = []
        self._current_align: str = "left"
        self._current_indent: int = 0
        self._in_body = False
        self._has_body_tag = False
        self._ignore_content = False

    def _start_paragraph(self, prefix: str = "", indent: int = 0) -> None:
        if self._current_runs:
            self._end_paragraph()
        self._current_runs = []
        self._current_align = "left"
        self._current_indent = indent
        for item in reversed(self._stack):
            if item['tag'] in ('p', 'div', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
                if item['styles'].get('align'):
                    self._current_align = item['styles']['align']
                break
        self._current_prefix = prefix

    def _end_paragraph(self) -> None:
        if not self._current_runs and not hasattr(self, '_current_prefix'): return
        merged = []
        for r in self._current_runs:
            if r['text'] == "": continue
            if merged and all(merged[-1][k] == r[k] for k in ('b','i','u','color')):
                merged[-1]['text'] += r['text']
            else:
                merged.append(r)
        prefix = getattr(self, '_current_prefix', '')
        if merged or prefix:
            self.paragraphs.append({
                'align': self._current_align,
                'runs': merged,
                'prefix': prefix,
                'indent': self._current_indent
            })
        self._current_runs = []
        if hasattr(self, '_current_prefix'): del self._current_prefix

    def _append_text(self, text: str):
        b = any(item['styles'].get('b') for item in self._stack)
        i = any(item['styles'].get('i') for item in self._stack)
        u = any(item['styles'].get('u') for item in self._stack)
        color = None
        for item in reversed(self._stack):
            if item['styles'].get('color'):
                color = item['styles']['color']
                break
        self._current_runs.append({'text': text, 'b': b, 'i': i, 'u': u, 'color': color})

    def feed(self, data: str) -> None:
        if "<body" in data.lower():
            self._has_body_tag = True; self._in_body = False
        else:
            self._has_body_tag = False; self._in_body = True
        super().feed(data)

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        attrs_d = dict(attrs)
        if tag == 'body': self._in_body = True; return
        if tag in ('head', 'style', 'script', 'meta', 'title', 'html', '!doctype'): self._ignore_content = True; return
        if not self._in_body or self._ignore_content: return

        styles = self._parse_style(attrs_d.get('style', ''))
        if attrs_d.get('align'): styles['align'] = attrs_d['align'].lower()
        if tag in ('b', 'strong'): styles['b'] = True
        if tag in ('i', 'em'): styles['i'] = True
        if tag == 'u': styles['u'] = True
        
        # Explicitní margin - ZVÝŠENÁ CITLIVOST (20px = 1 level)
        margin_left = styles.get('margin-left', '0px')
        explicit_indent_val = 0
        if 'px' in margin_left:
            try:
                px = float(margin_left.replace('px', '').strip())
                explicit_indent_val = int(px / 20) # Changed from 30 to 20
            except: pass

        parent_style = self._stack[-1]['styles'] if self._stack else {}
        merged_styles = parent_style.copy()
        merged_styles.update(styles)
        self._stack.append({'tag': tag, 'attrs': attrs_d, 'styles': merged_styles})

        list_nesting = len(self._list_stack)
        base_indent = max(0, list_nesting - 1) if tag == 'li' else max(0, list_nesting)
        final_indent = base_indent + explicit_indent_val

        if tag in ('p', 'div'):
            self._start_paragraph(indent=final_indent)
        elif tag == 'br':
            self._append_text("\n")
        elif tag in ('ul', 'ol'):
            t = attrs_d.get('type', '1')
            self._list_stack.append({'tag': tag, 'type': t, 'count': 0})
        elif tag == 'li':
            prefix = self._get_list_prefix()
            self._start_paragraph(prefix=prefix, indent=final_indent)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == 'body': self._in_body = False; return
        if tag in ('head', 'style', 'script', 'meta', 'title', 'html'): self._ignore_content = False; return
        if not self._in_body or self._ignore_content: return
        if tag in ('p', 'div', 'li'): self._end_paragraph()
        if tag in ('ul', 'ol'):
            if self._list_stack: self._list_stack.pop()
        for i in range(len(self._stack) - 1, -1, -1):
            if self._stack[i]['tag'] == tag:
                del self._stack[i:]
                break

    def handle_data(self, data):
        if not self._in_body or self._ignore_content: return
        if not data: return
        in_block = False
        current_indent = 0
        list_nesting = len(self._list_stack)
        current_indent = max(0, list_nesting)
        for item in reversed(self._stack):
            if item['tag'] in ('p', 'div', 'li'):
                in_block = True
                break
        if not in_block:
            if not data.strip(): return
            self._start_paragraph(indent=current_indent)
        self._append_text(data)

    def _parse_style(self, style_str: str) -> dict:
        res = {}
        if not style_str: return res
        for part in style_str.split(';'):
            if ':' in part:
                k,v = part.split(':', 1)
                k=k.strip().lower(); v=v.strip().lower()
                if k=='color':
                    m = re.search(r'#?([0-9a-f]{6})', v)
                    if m: res['color']=m.group(1)
                    else:
                         m2 = re.search(r'rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', v)
                         if m2: r,g,b=[int(x) for x in m2.groups()]; res['color']=f"{r:02X}{g:02X}{b:02X}"
                elif k=='text-align': res['align']=v
                elif k=='margin-left': res['margin-left']=v
                elif k=='font-weight':
                    if 'bold' in v or (re.search(r'\d+',v) and int(re.search(r'\d+',v).group(0))>=600): res['b']=True
                elif k=='font-style' and 'italic' in v: res['i']=True
                elif k=='text-decoration' and 'underline' in v: res['u']=True
        return res

    def _get_list_prefix(self) -> str:
        if not self._list_stack: return ""
        L = self._list_stack[-1]
        if L['tag'] == 'ul': return "•\t"
        elif L['tag'] == 'ol':
            L['count'] += 1; n = L['count']; t = L.get('type', '1')
            val = f"{n}."
            if t == 'a': val = f"{self._to_alpha(n, False)}."
            elif t == 'A': val = f"{self._to_alpha(n, True)}."
            elif t == 'i': val = f"{self._to_roman(n, False)}."
            elif t == 'I': val = f"{self._to_roman(n, True)}."
            return f"{val}\t"
        return ""
    def _to_alpha(self, n, upper):
        s=""; n-=1
        while n>=0: s=chr(97+n%26)+s; n=n//26-1
        return s.upper() if upper else s
    def _to_roman(self, n, upper):
        val=[(50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]; res=""
        for v,r in val:
            while n>=v: res+=r; n-=v
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
        # Načtení uložených cest
        self.settings_file = self.owner.project_root / "data" / "export_settings.json"
        self.stored_settings = self._load_settings()

        # Cesty (Default)
        default_templates_dir = self.owner.project_root / "data" / "Šablony"
        default_output_dir = self.owner.project_root / "data" / "Vygenerované testy"
        default_print_dir = self.owner.project_root / "data" / "Tisk"
        
        # Vytvoření složek, pokud neexistují
        default_templates_dir.mkdir(parents=True, exist_ok=True)
        default_output_dir.mkdir(parents=True, exist_ok=True)
        default_print_dir.mkdir(parents=True, exist_ok=True)
        
        # Aplikace uložených nebo defaultních cest
        self.templates_dir = Path(self.stored_settings.get("templates_dir", default_templates_dir))
        self.output_dir = Path(self.stored_settings.get("output_dir", default_output_dir))
        self.print_dir = Path(self.stored_settings.get("print_dir", default_print_dir))
        
        # Šablona
        last_template = self.stored_settings.get("last_template")
        if last_template and Path(last_template).exists():
            self.default_template = Path(last_template)
        else:
            self.default_template = self.templates_dir / "template.docx"

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
        else:
            self.le_template.setText(str(self.default_template))
            self._update_path_indicators()
        
        # Auto-generate Output Name
        self._update_default_output()
        
        # Update indikátorů
        self._update_path_indicators()
        
    def _load_settings(self) -> dict:
        if self.settings_file.exists():
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return {}

    def _save_settings(self):
        data = {
            "templates_dir": str(self.templates_dir),
            "output_dir": str(self.output_dir),
            "print_dir": str(self.print_dir),
            "last_template": self.le_template.text()
        }
        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Chyba ukládání nastavení exportu: {e}")

    def _update_path_indicators(self):
        # Šablona
        t_path = Path(self.le_template.text())
        if t_path.exists() and t_path.is_file():
            self.lbl_status_template.setText("✅ OK")
            self.lbl_status_template.setStyleSheet("color: #81c784; font-weight: bold;")
        else:
            self.lbl_status_template.setText("❌ Chybí")
            self.lbl_status_template.setStyleSheet("color: #ef5350; font-weight: bold;")
            
        # Výstupní složka
        if self.output_dir.exists() and self.output_dir.is_dir():
            self.lbl_status_out_dir.setText("✅ OK")
            self.lbl_status_out_dir.setStyleSheet("color: #81c784; font-weight: bold;")
        else:
            self.lbl_status_out_dir.setText("❌ Chybí")
            self.lbl_status_out_dir.setStyleSheet("color: #ef5350; font-weight: bold;")

        # Tisková složka
        if self.print_dir.exists() and self.print_dir.is_dir():
            self.lbl_status_print_dir.setText("✅ OK")
            self.lbl_status_print_dir.setStyleSheet("color: #81c784; font-weight: bold;")
        else:
            self.lbl_status_print_dir.setText("❌ Chybí")
            self.lbl_status_print_dir.setStyleSheet("color: #ef5350; font-weight: bold;")


    # --- Build Content Methods ---

    def _build_page1_content(self):
        self.page1.setTitle("Krok 1: Výběr šablony a nastavení cest")
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

        # GroupBox: Soubory a složky
        gb_files = QGroupBox("Cesty k souborům")
        grid_files = QGridLayout()
        grid_files.setColumnStretch(1, 1) # Input stretch
        
        # 1. Šablona
        grid_files.addWidget(QLabel("Šablona:"), 0, 0)
        self.le_template = QLineEdit()
        self.le_template.textChanged.connect(self._on_templ_change)
        grid_files.addWidget(self.le_template, 0, 1)
        
        btn_t = QPushButton("Vybrat...")
        btn_t.clicked.connect(self._choose_template)
        grid_files.addWidget(btn_t, 0, 2)
        
        self.lbl_status_template = QLabel("?")
        self.lbl_status_template.setFixedWidth(80)
        grid_files.addWidget(self.lbl_status_template, 0, 3)

        # 2. Výstup DOCX (Složka)
        grid_files.addWidget(QLabel("Složka pro testy:"), 1, 0)
        self.le_out_dir = QLineEdit(str(self.output_dir))
        self.le_out_dir.setReadOnly(True) # Editace jen přes dialog pro bezpečí
        grid_files.addWidget(self.le_out_dir, 1, 1)
        
        btn_od = QPushButton("Změnit...")
        btn_od.clicked.connect(self._choose_output_dir)
        grid_files.addWidget(btn_od, 1, 2)
        
        self.lbl_status_out_dir = QLabel("?")
        grid_files.addWidget(self.lbl_status_out_dir, 1, 3)
        
        # 3. Výstup PDF Tisk (Složka)
        grid_files.addWidget(QLabel("Složka pro tisk:"), 2, 0)
        self.le_print_dir = QLineEdit(str(self.print_dir))
        self.le_print_dir.setReadOnly(True)
        grid_files.addWidget(self.le_print_dir, 2, 1)
        
        btn_pd = QPushButton("Změnit...")
        btn_pd.clicked.connect(self._choose_print_dir)
        grid_files.addWidget(btn_pd, 2, 2)
        
        self.lbl_status_print_dir = QLabel("?")
        grid_files.addWidget(self.lbl_status_print_dir, 2, 3)
        
        # 4. Konkrétní soubor (náhled názvu)
        grid_files.addWidget(QLabel("Název souboru:"), 3, 0)
        self.le_output = QLineEdit()
        self.le_output.textChanged.connect(self._on_output_text_changed)
        grid_files.addWidget(self.le_output, 3, 1, 1, 2) # Span 2 sloupce
        
        gb_files.setLayout(grid_files)
        l1.addWidget(gb_files)
        
        self.lbl_scan_info = QLabel("Info: Čekám na načtení šablony...")
        self.lbl_scan_info.setStyleSheet("color: gray; font-style: italic; margin-top: 10px;")
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

        # 2. Volba režimu
        self.mode_box = QGroupBox("Režim exportu")
        self.mode_box.setStyleSheet("font-weight: bold; margin-top: 10px;")
        l_mode = QHBoxLayout(self.mode_box)
        
        self.rb_mode_single = QRadioButton("Jednotlivý export (Standardní)")
        self.rb_mode_single.setToolTip("Vytvoří jeden soubor s ručně vybranými otázkami.")
        self.rb_mode_single.setChecked(True)
        
        self.rb_mode_multi = QRadioButton("Hromadný export (Generátor variant)")
        self.rb_mode_multi.setToolTip("Vytvoří více kopií testu. Otázky 1-10 budou vybrány náhodně pro každou kopii.")
        
        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.rb_mode_single, 0)
        self.mode_group.addButton(self.rb_mode_multi, 1)
        self.mode_group.buttonToggled.connect(self._on_mode_toggled)
        
        l_mode.addWidget(self.rb_mode_single)
        l_mode.addWidget(self.rb_mode_multi)
        l_mode.addStretch()
        main_layout.addWidget(self.mode_box)

        # 3. Nastavení pro hromadný export (Skryté by default)
        self.widget_multi_options = QWidget()
        self.widget_multi_options.setVisible(False)
        self.widget_multi_options.setStyleSheet("background-color: #2d2d30; border-radius: 4px; padding: 4px;")
        l_multi = QHBoxLayout(self.widget_multi_options)
        l_multi.setContentsMargins(5, 5, 5, 5)
        
        l_multi.addWidget(QLabel("Počet kopií:"))
        self.spin_multi_count = QSpinBox()
        self.spin_multi_count.setRange(2, 50)
        self.spin_multi_count.setValue(2)
        l_multi.addWidget(self.spin_multi_count)
        
        l_multi.addWidget(QLabel("Zdroj otázek (pro <Otázka1-10>):"))
        self.combo_multi_source = QComboBox()
        self.combo_multi_source.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        l_multi.addWidget(self.combo_multi_source, 1)
        
        main_layout.addWidget(self.widget_multi_options)

        # 4. Hlavní obsah (Dva sloupce: Strom | Sloty)
        columns_layout = QHBoxLayout()
        
        # Levý panel: Strom (NOVÉ: Obaleno ve widgetu pro skrývání)
        self.widget_left_panel = QWidget()
        left_layout = QVBoxLayout(self.widget_left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        left_layout.addWidget(QLabel("<b>Dostupné otázky:</b>"))
        self.tree_source = QTreeWidget()
        self.tree_source.setHeaderLabels(["Struktura otázek"])
        self.tree_source.setSelectionMode(QAbstractItemView.ExtendedSelection)
        
        # Nastavení signálů
        self.tree_source.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_source.customContextMenuRequested.connect(self._on_tree_source_context_menu)
        
        if hasattr(self, "_on_tree_selection"):
            self.tree_source.itemSelectionChanged.connect(self._on_tree_selection)
        
        left_layout.addWidget(self.tree_source)
        
        self.btn_assign_multi = QPushButton(">> Přiřadit vybrané na volné pozice >>")
        self.btn_assign_multi.setToolTip("Doplní vybrané otázky (zleva) na první volná místa v šabloně (vpravo).")
        self.btn_assign_multi.clicked.connect(self._assign_selected_multi)
        left_layout.addWidget(self.btn_assign_multi)
        
        columns_layout.addWidget(self.widget_left_panel, 4)
        
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

        # 5. Náhled
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

    def _on_mode_toggled(self, btn, checked):
        """Reaguje na změnu režimu exportu (Single vs Multi)."""
        if not checked: return
        
        is_multi = (self.mode_group.checkedId() == 1)
        self.widget_multi_options.setVisible(is_multi)
        
        # NOVÉ: Skrytí levého panelu s otázkami při multi režimu
        self.widget_left_panel.setVisible(not is_multi)
        
        self._update_slots_visuals(is_multi)

    def _update_slots_visuals(self, is_multi: bool):
        """Aktualizuje vzhled slotů podle režimu."""
        # Projdeme všechny widgety ve slot layoutu
        for i in range(self.layout_slots.count()):
            item = self.layout_slots.itemAt(i)
            w = item.widget()
            if not w: continue
            
            ph = w.property("placeholder")
            if not ph: continue # Není to slot (např. nadpis)
            
            # Získáme reference na tlačítka uvnitř slotu
            row_layout = w.layout()
            if not row_layout or row_layout.count() < 3: continue
            
            btn_assign = row_layout.itemAt(1).widget()
            btn_clear = row_layout.itemAt(2).widget()
            
            # Logika pro MULTI: Zamknout VŠECHNY klasické otázky (OtázkaX)
            # Pokud je placeholder v seznamu klasických otázek
            if is_multi and ph in self.placeholders_q:
                btn_assign.setText("🎲 NÁHODNĚ Z VARIANT")
                btn_assign.setStyleSheet("color: #ffcc00; font-weight: bold; border: 1px dashed #ffcc00;")
                btn_assign.setEnabled(False)
                btn_clear.setEnabled(False)
            else:
                # Obnovíme standardní stav (nebo pro Bonusy v multi režimu - ty zůstávají manuální/prázdné?)
                # Požadavek byl jen na "počet otázek", předpokládám že Bonusy se negenerují náhodně (nebo ano?)
                # Zadání: "počet klasických otázek se i pro tento režim bude odvíjet od šablony"
                # Takže jen self.placeholders_q
                
                btn_assign.setEnabled(True)
                btn_clear.setEnabled(True)
                btn_assign.setStyleSheet("")
                
                qid = self.selection_map.get(ph)
                if qid:
                    q = self.owner._find_question_by_id(qid)
                    if q:
                        btn_assign.setText(q.title)
                    else:
                        btn_assign.setText("???")
                else:
                    btn_assign.setText("Vybrat...")

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
        
        # Info Panel
        self.info_box_p3 = QGroupBox("Kontext exportu")
        self.info_box_p3.setStyleSheet("QGroupBox { font-weight: bold; border: 1px solid #555; margin-top: 6px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; }")
        l_info = QFormLayout(self.info_box_p3)
        self.lbl_templ_p3 = QLabel("-")
        self.lbl_out_p3 = QLabel("-")
        l_info.addRow("Vstupní šablona:", self.lbl_templ_p3)
        l_info.addRow("Výstupní soubor:", self.lbl_out_p3)
        main_layout.addWidget(self.info_box_p3)
        
        # NOVÉ: Exportní volby
        self.options_box = QGroupBox("Exportní volby")
        self.options_box.setStyleSheet("font-weight: bold; margin-top: 6px;")
        l_opts = QVBoxLayout(self.options_box)
        
        self.chk_export_pdf = QCheckBox("Exportovat do PDF pro tisk (složka /data/Tisk/)")
        self.chk_export_pdf.setChecked(True)
        self.chk_export_pdf.setToolTip("DOCX se automaticky převede na PDF. V hromadném režimu budou všechny varianty spojeny do jednoho PDF.")
        l_opts.addWidget(self.chk_export_pdf)
        
        main_layout.addWidget(self.options_box)
        
        # Náhled
        lbl_prev = QLabel("<b>Náhled obsahu testu:</b>")
        main_layout.addWidget(lbl_prev)
        
        self.preview_edit = QTextEdit()
        self.preview_edit.setReadOnly(True)
        # Dark Theme CSS pro QTextEdit
        self.preview_edit.setStyleSheet("QTextEdit { background-color: #252526; color: #e0e0e0; border: 1px solid #3e3e42; }")
        main_layout.addWidget(self.preview_edit)

        # Hash label (OPRAVA FONTU)
        self.lbl_hash_preview = QLabel("Hash: -")
        self.lbl_hash_preview.setWordWrap(True)
        # Změna z "Monospace" na "Consolas, Monaco, monospace"
        self.lbl_hash_preview.setStyleSheet("color: #555; font-family: Consolas, Monaco, monospace; font-size: 10px; margin-top: 5px;")
        self.lbl_hash_preview.setTextInteractionFlags(Qt.TextSelectableByMouse)
        main_layout.addWidget(self.lbl_hash_preview)

    # --- Helpers & Logic ---

    def _update_default_output(self):
        if self.output_changed_manually and self.sender() == self.le_output:
            # Pokud uživatel edituje ručně celou cestu, necháme ho
            self.output_path = Path(self.le_output.text())
            return

        prefix = self.le_prefix.text().strip()
        # Nahrazení nepovolených znaků
        prefix = re.sub(r'[\\/*?:"<>|]', "_", prefix)
        
        dt = self.dt_edit.dateTime()
        # 2. část: YYYY-MM-DD_HHMM
        date_time_str = dt.toString("yyyy-MM-dd_HHmm")
        
        # 3. část: Timestamp (aktuální čas v sekundách)
        timestamp = str(int(datetime.now().timestamp()))
        
        # Sestavení názvu: PREFIX_DATETIME_TIMESTAMP.docx
        filename = f"{prefix}_{date_time_str}_{timestamp}.docx"
        
        self.output_path = self.output_dir / filename
        
        # Blokujeme signál, abychom necyklili přes _on_output_text_changed
        self.le_output.blockSignals(True)
        self.le_output.setText(str(self.output_path))
        self.le_output.blockSignals(False)
        
        self.page1.completeChanged.emit()

    def _on_output_text_changed(self, text):
        self.output_changed_manually = True
        self.output_path = Path(text)

    def _choose_template(self):
        start_dir = str(self.templates_dir) if self.templates_dir.exists() else str(self.owner.project_root)
        path, _ = QFileDialog.getOpenFileName(self, "Vybrat šablonu", start_dir, "*.docx")
        if path:
            self.le_template.setText(path)
            self.templates_dir = Path(path).parent
            # Uložit nastavení ihned
            self._save_settings()

    def _choose_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Cíl exportu", str(self.default_output), "*.docx")
        if path:
            self.le_output.setText(path)
            
    def _choose_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Vybrat složku pro testy", str(self.output_dir))
        if d:
            self.output_dir = Path(d)
            self.le_out_dir.setText(d)
            # Uložit nastavení ihned
            self._save_settings()
            
            self._update_path_indicators()
            self._update_default_output() # Přegenerovat cestu k souboru

    def _choose_print_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Vybrat složku pro tisk", str(self.print_dir))
        if d:
            self.print_dir = Path(d)
            self.le_print_dir.setText(d)
            # Uložit nastavení ihned
            self._save_settings()
            
            self._update_path_indicators()

    def _on_templ_change(self, text: str):
        self.template_path = Path(text)
        self._update_path_indicators()
        
        # Uložit nastavení ihned
        self._save_settings()
        
        if self.template_path.exists() and self.template_path.suffix.lower() == '.docx':
             self._scan_placeholders()
        else:
             self.lbl_scan_info.setText("Šablona nenalezena nebo není .docx")
        
        self.page1.completeChanged.emit()


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
        # Náhled funguje jen pokud je vybrána přesně jedna položka
        if not sel or len(sel) != 1:
            self.text_preview_q.clear()
            return

        item = sel[0]
        data = item.data(0, Qt.UserRole)

        # Očekáváme buď přímo ID otázky, nebo dict s metadaty (kind/id/...)
        if not data:
            self.text_preview_q.setText("--- (Vyberte konkrétní otázku pro náhled) ---")
            return

        if isinstance(data, dict):
            # Náhled má smysl jen pro položku typu 'question'
            if data.get("kind") != "question":
                self.text_preview_q.setText("--- (Vyberte konkrétní otázku pro náhled) ---")
                return
            qid = data.get("id")
        else:
            qid = data

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

            # 4. Squeeze whitespace – odstranění prázdných řádků
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

            # 1. Clear Tree & Combo
            self.tree_source.blockSignals(True)
            self.tree_source.clear()
            self.combo_multi_source.clear()
            
            # 2. Clear Slots
            while self.layout_slots.count():
                item = self.layout_slots.takeAt(0)
                if item.widget(): item.widget().deleteLater()
            self.layout_slots.addStretch()
            
            # 4. Populate Tree and Combo
            def add_subgroup_recursive(parent_item, subgroup_list, parent_gid, level=1):
                for sg in subgroup_list:
                    # Tree Item
                    sg_item = QTreeWidgetItem([sg.name])
                    sg_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
                    sg_item.setData(0, Qt.UserRole, {
                        "kind": "subgroup", 
                        "id": sg.id, 
                        "parent_group_id": parent_gid
                    })
                    parent_item.addChild(sg_item)
                    
                    # Combo Item
                    indent = "  " * level
                    self.combo_multi_source.addItem(f"{indent}📂 {sg.name}", {"type": "subgroup", "id": sg.id})
                    
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
                        add_subgroup_recursive(sg_item, sg.subgroups, parent_gid, level + 1)

            groups = self.owner.root.groups
            for g in groups:
                # Tree
                g_item = QTreeWidgetItem([g.name])
                g_item.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
                f = g_item.font(0); f.setBold(True); g_item.setFont(0, f)
                g_item.setData(0, Qt.UserRole, {
                    "kind": "group", 
                    "id": g.id
                })
                self.tree_source.addTopLevelItem(g_item)
                
                # Combo
                self.combo_multi_source.addItem(f"📁 {g.name}", {"type": "group", "id": g.id})
                
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
            
            # Aplikovat vizuální stav podle aktuálního režimu
            is_multi = (self.mode_group.checkedId() == 1)
            self._update_slots_visuals(is_multi)
            
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
            # 1. Generování hashe
            ts = str(datetime.now().timestamp())
            salt = secrets.token_hex(16)
            data_to_hash = f"{ts}{salt}"
            self._cached_hash = hashlib.sha3_256(data_to_hash.encode("utf-8")).hexdigest()
            
            if hasattr(self, "lbl_hash_preview"):
                self.lbl_hash_preview.setText(f"SHA3-256 Hash:\n{self._cached_hash}")

            t_name = self.template_path.name if self.template_path else "Nevybráno"
            o_name = self.output_path.name if self.output_path else "Nevybráno"
            self.lbl_templ_p3.setText(t_name)
            self.lbl_out_p3.setText(o_name)

            total_bonus_points = 0.0
            min_loss = 0.0
            
            bg_color = "#252526"; text_color = "#e0e0e0"; border_color = "#555555"
            sec_q_bg = "#2d3845"; sec_b_bg = "#453d2d"; sec_s_bg = "#2d452d"
            
            prefix = self.le_prefix.text().strip()
            today = datetime.now().strftime("%Y-%m-%d")
            verze_preview = f"{prefix} {today}" 
            
            is_multi = (self.mode_group.checkedId() == 1)
            
            multi_info = ""
            if is_multi:
                multi_count = self.spin_multi_count.value()
                multi_info = f"<tr><td colspan='2' style='color: #ffcc00; font-weight: bold;'>⚡ Hromadný export: {multi_count} verzí (stejný hash)</td></tr>"

            html = f"""
            <html>
            <body style='font-family: Arial, sans-serif; background-color: {bg_color}; color: {text_color};'>
            <h2 style='color: #61dafb; border-bottom: 2px solid #61dafb;'>Souhrn testu</h2>
            <table width='100%' style='margin-bottom: 20px; color: {text_color};'>
                <tr>
                    <td><b>Verze:</b> {verze_preview}</td>
                    <td align='right'><b>Datum:</b> {self.dt_edit.dateTime().toString("dd.MM.yyyy HH:mm")}</td>
                </tr>
                {multi_info}
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
                    # Pokud je multi a placeholder je v seznamu klasických otázek -> Náhodný
                    if is_multi and ph in self.placeholders_q:
                        html += f"<tr><td width='100' style='color:#888;'>{ph}:</td><td colspan='2' style='color:#ffcc00;'>[Náhodný výběr pro každou verzi]</td></tr>"
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
        self._save_settings()
        
        if not self.template_path or not self.output_path:
            return

        # Kontrolní Hash
        k_hash = getattr(self, "_cached_hash", "")
        if not k_hash:
            ts = str(datetime.now().timestamp())
            salt = secrets.token_hex(16)
            data_to_hash = f"{ts}{salt}"
            k_hash = hashlib.sha3_256(data_to_hash.encode("utf-8")).hexdigest()

        is_multi = (self.mode_group.checkedId() == 1)
        count = self.spin_multi_count.value() if is_multi else 1
        do_pdf_export = self.chk_export_pdf.isChecked()
        
        # Příprava poolu otázek
        question_pool = []
        if is_multi:
            data = self.combo_multi_source.currentData()
            if data:
                def collect_questions(group_id, is_subgroup):
                    qs = []
                    nodes_to_visit = list(self.owner.root.groups)
                    target_node = None
                    while nodes_to_visit:
                        curr = nodes_to_visit.pop(0)
                        if curr.id == group_id:
                            target_node = curr
                            break
                        if hasattr(curr, "subgroups") and curr.subgroups:
                            nodes_to_visit.extend(curr.subgroups)
                    if target_node:
                        def extract_q(node):
                            valid_qs = []
                            if hasattr(node, "questions"):
                                valid_qs.extend([q.id for q in node.questions if q.type == 'classic'])
                            if hasattr(node, "subgroups") and node.subgroups:
                                for sub in node.subgroups:
                                    valid_qs.extend(extract_q(sub))
                            return valid_qs
                        qs = extract_q(target_node)
                    return qs
                is_sub = (data["type"] == "subgroup")
                question_pool = collect_questions(data["id"], is_sub)
        
        base_output_path = self.output_path
        success_count = 0
        generated_docx_files = []
        generated_pdf_files = []

        print_folder = self.print_dir
        if do_pdf_export:
            print_folder.mkdir(parents=True, exist_ok=True)

        # Loop generování DOCX
        for i in range(count):
            current_selection = self.selection_map.copy()
            
            if is_multi and question_pool:
                import random
                # ZMĚNA: Cílem jsou VŠECHNY klasické placeholdery ze šablony
                targets = self.placeholders_q
                needed = len(targets)
                
                if len(question_pool) >= needed:
                    picked = random.sample(question_pool, needed)
                    for idx, ph in enumerate(targets):
                        current_selection[ph] = picked[idx]
                else:
                    if len(question_pool) > 0:
                        for ph in targets:
                            current_selection[ph] = random.choice(question_pool)
            
            repl_plain: Dict[str, str] = {}
            dt = round_dt_to_10m(self.dt_edit.dateTime())
            dt_str = f"{cz_day_of_week(dt.toPython())} {dt.toString('dd.MM.yyyy HH:mm')}"
            repl_plain["DatumČas"] = dt_str
            repl_plain["DatumCas"] = dt_str
            repl_plain["DATUMCAS"] = dt_str
            
            prefix = self.le_prefix.text().strip()
            today = datetime.now().strftime("%Y-%m-%d")
            verze_str = f"{prefix} {today}"
            repl_plain["PoznamkaVerze"] = verze_str
            repl_plain["POZNAMKAVERZE"] = verze_str
            repl_plain["KontrolniHash"] = k_hash
            repl_plain["KONTROLNIHASH"] = k_hash
            
            total_bonus = 0.0
            min_loss = 0.0
            for qid in current_selection.values():
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

            rich_map: Dict[str, str] = {}
            for ph, qid in current_selection.items():
                q = self.owner._find_question_by_id(qid)
                if q:
                    rich_map[ph] = q.text_html

            if is_multi:
                p = Path(base_output_path)
                new_name = f"{p.stem}_v{i+1}{p.suffix}"
                target_path = p.parent / new_name
            else:
                target_path = base_output_path

            try:
                self.owner._generate_docx_from_template(self.template_path, target_path, repl_plain, rich_map)
                success_count += 1
                generated_docx_files.append(target_path)
            except Exception as e:
                QMessageBox.critical(self, "Export", f"Chyba při exportu verze {i+1}:\n{e}")
                if not is_multi: return

        # Historie
        if is_multi:
            record_name = f"Balík {count} verzí: {base_output_path.name}"
            self.owner.register_export(record_name, k_hash)
        else:
            self.owner.register_export(base_output_path.name, k_hash)

        # PDF Export
        pdf_success_msg = ""
        if do_pdf_export and generated_docx_files:
            try:
                for docx_file in generated_docx_files:
                    pdf_file = self.owner._convert_docx_to_pdf(docx_file)
                    if pdf_file and pdf_file.exists():
                        generated_pdf_files.append(pdf_file)
                
                if generated_pdf_files:
                    if is_multi and len(generated_pdf_files) > 1:
                        merged_name = f"{base_output_path.stem}_merged.pdf"
                        final_pdf = print_folder / merged_name
                        
                        if self.owner._merge_pdfs(generated_pdf_files, final_pdf, cleanup=True):
                            pdf_success_msg = f"\n\nPDF pro tisk (sloučené) uloženo do:\n{final_pdf}"
                        else:
                            pdf_success_msg = f"\n\nPOZOR: Slučování selhalo. Jednotlivá PDF jsou v:\n{print_folder}"
                            import shutil
                            for tmp_pdf in generated_pdf_files:
                                if tmp_pdf.exists():
                                    try:
                                        dest = print_folder / tmp_pdf.name
                                        shutil.move(str(tmp_pdf), str(dest))
                                    except: pass
                    else:
                        import shutil
                        final_pdf = print_folder / generated_pdf_files[0].name
                        shutil.move(str(generated_pdf_files[0]), str(final_pdf))
                        pdf_success_msg = f"\n\nPDF pro tisk uloženo do:\n{final_pdf}"
                            
            except Exception as e:
                QMessageBox.warning(self, "PDF Export", f"Kritická chyba při exportu PDF:\n{e}")
                import traceback
                traceback.print_exc()
        
        if is_multi:
            msg = f"Hromadný export dokončen.\nVygenerováno {success_count} souborů DOCX.{pdf_success_msg}"
            QMessageBox.information(self, "Export", msg)
        else:
            msg = f"Export dokončen.\nSoubor uložen:\n{base_output_path}{pdf_success_msg}"
            QMessageBox.information(self, "Export", msg)
            
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
            
    def set_data(self, text: str, date_str: str, author: str, source_doc: Optional[str] = None) -> None:
        """Naplní formulář daty pro editaci."""
        self.text_edit.setText(text)
        self.author_edit.setText(author)

        # Pokusíme se parsovat datum (očekáváme dd.MM.yyyy nebo dd.MM.yyyy HH:mm)
        dt = QDateTime.fromString(date_str, "dd.MM.yyyy HH:mm")
        if not dt.isValid():
            dt = QDateTime.fromString(date_str, "dd.MM.yyyy")

        if dt.isValid():
            self.date_edit.setDateTime(dt)

        # Nastavení zdrojového dokumentu v comboboxu
        if source_doc:
            idx = self.combo_source.findData(source_doc)
            if idx >= 0:
                self.combo_source.setCurrentIndex(idx)
            else:
                self.combo_source.setCurrentIndex(0)
        else:
            self.combo_source.setCurrentIndex(0)
            
    def get_data(self) -> tuple[str, str, str, str]:
        # Vracíme text, datum, autora a (případně) cestu ke zdrojovému dokumentu
        text = self.text_edit.toPlainText().strip()
        date_str = self.date_edit.dateTime().toString("dd.MM.yyyy HH:mm")
        author = self.author_edit.text().strip()
        data = self.combo_source.currentData()
        source_doc = str(data) if data is not None else ""
        return text, date_str, author, source_doc

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
        
        
        self._build_menus()
        self.load_data()
        self._refresh_tree()
        self._refresh_funny_answers_tab()

        # ZMĚNA: Strom 60%, Editor 40% (cca 840px : 560px)
        # Nyní, když je self.splitter správně nastaven v _build_ui, můžeme přímo nastavit velikosti.
        self.splitter.setSizes([940, 860])

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

        # LEVÝ PANEL
        left_panel_container = QWidget()
        left_container_layout = QVBoxLayout(left_panel_container)
        left_container_layout.setContentsMargins(0, 0, 0, 0)
        self.left_tabs = QTabWidget()
        
        # ZÁLOŽKA 1: OTÁZKY
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
        # NOVÉ: Kontextové menu
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._on_tree_context_menu)
        questions_layout.addWidget(self.tree, 1)
        self.left_tabs.addTab(self.tab_questions, "Otázky")

        # ZÁLOŽKA 2: HISTORIE
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
        self.table_history.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_history.customContextMenuRequested.connect(self._on_history_context_menu)
        history_layout.addWidget(self.table_history)
        btn_refresh_hist = QPushButton("Obnovit historii")
        btn_refresh_hist.clicked.connect(self._refresh_history_table)
        history_layout.addWidget(btn_refresh_hist)
        self.left_tabs.addTab(self.tab_history, "Historie")
        
        self._init_funny_answers_tab()
        left_container_layout.addWidget(self.left_tabs)

        # PRAVÝ PANEL (Detail / Editor)
        self.detail_stack = QWidget()
        self.detail_layout = QVBoxLayout(self.detail_stack)
        self.detail_layout.setContentsMargins(6, 6, 6, 6)
        self.detail_layout.setSpacing(8)

        # Toolbar
        self.editor_toolbar = QToolBar("Formát")
        self.editor_toolbar.setIconSize(QSize(18, 18))
        self.action_bold = QAction("Tučné", self); self.action_bold.setCheckable(True); self.action_bold.setShortcut(QKeySequence.Bold)
        self.action_italic = QAction("Kurzíva", self); self.action_italic.setCheckable(True); self.action_italic.setShortcut(QKeySequence.Italic)
        self.action_underline = QAction("Podtržení", self); self.action_underline.setCheckable(True); self.action_underline.setShortcut(QKeySequence.Underline)
        self.action_color = QAction("Barva", self)
        self.action_bullets = QAction("Odrážky", self); self.action_bullets.setCheckable(True)
        self.action_indent_dec = QAction("< Odsadit", self); self.action_indent_dec.setToolTip("Zmenšit odsazení")
        self.action_indent_inc = QAction("> Odsadit", self); self.action_indent_inc.setToolTip("Zvětšit odsazení")
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
        self.editor_toolbar.addAction(self.action_indent_dec)
        self.editor_toolbar.addAction(self.action_indent_inc)
        self.editor_toolbar.addSeparator()
        self.editor_toolbar.addAction(self.action_align_left)
        self.editor_toolbar.addAction(self.action_align_center)
        self.editor_toolbar.addAction(self.action_align_right)
        self.editor_toolbar.addAction(self.action_align_justify)

        # Horní formulář (Název, Typ, Body)
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

        # Správná odpověď (přesunuto dolů, definice zůstává zde)
        self.edit_correct_answer = QTextEdit()
        self.edit_correct_answer.setPlaceholderText("Volitelný text správné odpovědi...")
        self.edit_correct_answer.setFixedHeight(60)
        
        # Vtipné odpovědi (přesunuto dolů, definice zůstává zde)
        self.funny_container = QWidget()
        fc_layout = QVBoxLayout(self.funny_container)
        fc_layout.setContentsMargins(0,0,0,0)
        self.table_funny = QTableWidget(0, 4)
        self.table_funny.setHorizontalHeaderLabels(["Odpověď", "Datum", "Jméno", "Zdroj"])
        self.table_funny.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table_funny.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_funny.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table_funny.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
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

        # Obsah otázky (Text Edit)
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

        # SKLÁDÁNÍ LAYOUTU
        self.detail_layout.addWidget(self.editor_toolbar)
        self.detail_layout.addLayout(self.form_layout)
        
        # 1. Obsah otázky (uložíme label do self pro skrývání)
        self.lbl_content = QLabel("<b>Obsah otázky:</b>")
        self.detail_layout.addWidget(self.lbl_content)
        self.detail_layout.addWidget(self.text_edit, 1) 
        
        # 2. Správná odpověď
        self.lbl_correct = QLabel("<b>Správná odpověď:</b>")
        self.detail_layout.addWidget(self.lbl_correct)
        self.detail_layout.addWidget(self.edit_correct_answer)
        
        # 3. Vtipné odpovědi
        self.lbl_funny = QLabel("<b>Vtipné odpovědi:</b>")
        self.detail_layout.addWidget(self.lbl_funny)
        self.detail_layout.addWidget(self.funny_container)
        
        # Tlačítko uložit
        self.detail_layout.addWidget(self.btn_save_question)
        self.detail_layout.addWidget(self.rename_panel)

        self._set_editor_enabled(False)
        self.splitter.addWidget(left_panel_container)
        self.splitter.addWidget(self.detail_stack)
        self.splitter.setStretchFactor(1, 1)
        self.setCentralWidget(self.splitter)

        # Toolbar aplikace
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
        
    def _on_tree_context_menu(self, pos: QPoint) -> None:
        """Kontextové menu stromu otázek (v6.7.2)."""
        item = self.tree.itemAt(pos)
        if not item:
            return
            
        self.tree.setCurrentItem(item)

        # Robustní získání metadat (podpora tuple i dict)
        raw_data = item.data(0, Qt.UserRole)
        kind = "unknown"
        
        if isinstance(raw_data, tuple) and len(raw_data) >= 1:
            kind = raw_data[0]
        elif isinstance(raw_data, dict):
            kind = raw_data.get("kind", "unknown")

        menu = QMenu(self.tree)
        has_action = False

        # 1. Přidat podskupinu (Group/Subgroup)
        if kind in ("group", "subgroup"):
            act = menu.addAction("Přidat podskupinu")
            act.triggered.connect(self._add_subgroup)
            has_action = True

        # 2. Přidat otázku (Subgroup)
        if kind == "subgroup":
            act = menu.addAction("Přidat otázku")
            act.triggered.connect(self._add_question)
            has_action = True
            
        # 3. Duplikovat otázku (Question)
        if kind == "question":
            act = menu.addAction("Duplikovat otázku")
            act.triggered.connect(self._duplicate_question)
            has_action = True

        if has_action:
            menu.addSeparator()

        # 4. Smazat (Vše) -> Použijeme existující metodu _delete_selected
        act_del = menu.addAction("Smazat")
        act_del.triggered.connect(self._delete_selected)

        menu.exec(self.tree.mapToGlobal(pos))

    def _change_indent(self, steps: int) -> None:
        """Změní odsazení aktuálního bloku nebo listu."""
        cursor = self.text_edit.textCursor()
        cursor.beginEditBlock()
        
        current_list = cursor.currentList()
        if current_list:
            # Pokud jsme v listu, měníme level (odsazení formátu listu)
            fmt = current_list.format()
            current_indent = fmt.indent()
            new_indent = max(1, current_indent + steps)
            fmt.setIndent(new_indent)
            current_list.setFormat(fmt)
        else:
            # Pokud je to běžný text, měníme margin bloku
            block_fmt = cursor.blockFormat()
            current_margin = block_fmt.leftMargin()
            # Krok odsazení např. 20px
            new_margin = max(0, current_margin + (steps * 20))
            block_fmt.setLeftMargin(new_margin)
            cursor.setBlockFormat(block_fmt)
            
        cursor.endEditBlock()
        self._autosave_schedule()
        self.text_edit.setFocus()


    def _on_format_bullets(self, checked: bool) -> None:
        """Přepne aktuální výběr na odrážky s lepším odsazením."""
        cursor = self.text_edit.textCursor()
        cursor.beginEditBlock()
        
        if checked:
            # Vytvoříme formát seznamu
            list_fmt = QTextListFormat()
            list_fmt.setStyle(QTextListFormat.ListDisc)
            list_fmt.setIndent(1) # Level 1
            
            # Nastavení odsazení (v pixelech/bodech) pro vizuální úpravu
            # NumberSuffix = mezera za odrážkou
            list_fmt.setNumberPrefix("")
            list_fmt.setNumberSuffix(" ") 
            
            cursor.createList(list_fmt)
            
            # Vynucení odsazení bloku (pro celý list item)
            block_fmt = cursor.blockFormat()
            block_fmt.setLeftMargin(15)  # Odsazení celého bloku zleva
            block_fmt.setTextIndent(-10) # Předsazení odrážky (aby byla vlevo od textu)
            cursor.setBlockFormat(block_fmt)
        else:
            # Zrušit list
            # Standardní blok bez listu
            block_fmt = QTextBlockFormat()
            block_fmt.setObjectIndex(-1) # Not a list
            cursor.setBlockFormat(block_fmt)
            
            # Reset odsazení
            cursor.setBlockFormat(QTextBlockFormat())

        cursor.endEditBlock()
        # Synchronizace stavu tlačítka (pokud by cursor change změnil stav zpět)
        self.text_edit.setFocus()


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

    def _init_funny_answers_tab(self):
        """Inicializuje záložku 'Hall of Shame' s detailním náhledem."""
        self.tab_funny = QWidget()
        layout = QVBoxLayout(self.tab_funny)
        layout.setContentsMargins(4, 4, 4, 4)
        
        # Filtr a tlačítka
        top_bar = QHBoxLayout()
        self.le_funny_filter = QLineEdit()
        self.le_funny_filter.setPlaceholderText("Hledat v odpovědích...")
        self.le_funny_filter.textChanged.connect(self._filter_funny_answers)
        
        btn_refresh = QPushButton("Obnovit")
        btn_refresh.clicked.connect(self._refresh_funny_answers_tab)
        
        top_bar.addWidget(self.le_funny_filter)
        top_bar.addWidget(btn_refresh)
        layout.addLayout(top_bar)
        
        # Splitter pro Strom a Detail
        splitter = QSplitter(Qt.Vertical)
        
        # Strom odpovědí (TreeWidget)
        self.tree_funny = QTreeWidget()
        self.tree_funny.setHeaderLabels(["Odpověď / Otázka", "Datum", "Jméno", "Zdroj"])
        self.tree_funny.setColumnWidth(0, 400)
        self.tree_funny.itemSelectionChanged.connect(self._on_funny_tree_select) # Signál výběru
        
        splitter.addWidget(self.tree_funny)
        
        # Detail odpovědi
        detail_container = QWidget()
        detail_layout = QVBoxLayout(detail_container)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.addWidget(QLabel("<b>Celé znění odpovědi:</b>"))
        
        self.text_funny_detail = QTextEdit()
        self.text_funny_detail.setReadOnly(True)
        self.text_funny_detail.setPlaceholderText("Vyberte odpověď pro zobrazení detailu...")
        detail_layout.addWidget(self.text_funny_detail)
        
        splitter.addWidget(detail_container)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        
        layout.addWidget(splitter)
        
        self.left_tabs.addTab(self.tab_funny, "Hall of Shame")

    def _on_funny_tree_select(self):
        """Zobrazí detail vybrané vtipné odpovědi ze stromu."""
        selected = self.tree_funny.selectedItems()
        if not selected:
            self.text_funny_detail.clear()
            return
            
        item = selected[0]
        # Pokud je to top-level item (otázka), nic nezobrazujeme (nebo název otázky)
        if item.parent() is None:
            self.text_funny_detail.clear()
            return
            
        # Získáme plný text z UserRole
        full_text = item.data(0, Qt.UserRole)
        if full_text:
            self.text_funny_detail.setText(full_text)
        else:
            # Fallback na text položky, pokud data chybí
            self.text_funny_detail.setText(item.text(0))

    def _filter_funny_answers(self, text: str):
        """Filtruje položky ve stromu vtipných odpovědí."""
        search_text = text.lower()
        
        # Projdeme všechny top-level položky (otázky)
        root = self.tree_funny.invisibleRootItem()
        for i in range(root.childCount()):
            q_item = root.child(i)
            
            # Projdeme odpovědi (děti)
            has_visible_child = False
            for j in range(q_item.childCount()):
                child = q_item.child(j)
                child_text = child.text(0).lower()
                
                if not search_text or search_text in child_text:
                    child.setHidden(False)
                    has_visible_child = True
                else:
                    child.setHidden(True)
            
            # Pokud otázka nemá viditelné odpovědi a sama neodpovídá filtru, skryjeme ji
            # (Zde pro jednoduchost filtrujeme jen podle obsahu odpovědí)
            q_item.setHidden(not has_visible_child)
            
            # Pokud filtr není prázdný a našli jsme shodu, rozbalíme
            if search_text and has_visible_child:
                q_item.setExpanded(True)
            elif not search_text:
                q_item.setExpanded(False)

    def _refresh_funny_answers_tab(self) -> None:
        """Znovu vygeneruje strom 'Seznam vtipných odpovědí' ze struktury otázek."""
        if not hasattr(self, "tree_funny"):
            return

        self.tree_funny.clear()
        self.text_funny_detail.clear()

        root = getattr(self, "root", None)
        if root is None or not root.groups:
            return

        question_brush = QBrush(QColor("white"))

        pix = QPixmap(16, 16)
        pix.fill(Qt.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QColor("teal"))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(1, 1, 14, 14)
        painter.setPen(QColor("white"))
        font = painter.font()
        font.setBold(True)
        font.setPointSize(10)
        painter.setFont(font)
        painter.drawText(pix.rect(), Qt.AlignCenter, "?")
        painter.end()
        question_icon = QIcon(pix)

        all_questions = []

        def collect_questions(subgroups: List[Subgroup]) -> None:
            for sg in subgroups:
                for q in sg.questions:
                    f_list = getattr(q, "funny_answers", []) or []
                    if f_list:
                        all_questions.append(q)
                collect_questions(sg.subgroups)

        for g in root.groups:
            collect_questions(g.subgroups)

        # Sort questions alphabetically
        all_questions.sort(key=lambda q: (q.title or "").lower())

        # Helper pro parsování data
        def parse_date_safe(date_str):
            if not date_str: return datetime.min
            # Formáty: dd.mm.yyyy, yyyy-mm-dd
            for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M"):
                try:
                    return datetime.strptime(str(date_str).strip(), fmt)
                except ValueError:
                    pass
            return datetime.min

        for q in all_questions:
            if self.tree_funny.topLevelItemCount() > 0:
                spacer = QTreeWidgetItem()
                spacer.setFlags(Qt.NoItemFlags)
                spacer.setSizeHint(0, QSize(0, 15))
                self.tree_funny.addTopLevelItem(spacer)

            q_title = q.title or "(bez názvu)"
            q_item = QTreeWidgetItem()
            q_item.setText(0, q_title)
            q_item.setIcon(0, question_icon)
            
            for col in range(4):
                q_item.setForeground(col, question_brush)

            self.tree_funny.addTopLevelItem(q_item)

            f_list = getattr(q, "funny_answers", [])
            
            # Sort answers by parsed date (Descending = Newest first)
            f_list_sorted = sorted(f_list, key=lambda x: parse_date_safe(x.date if isinstance(x, FunnyAnswer) else x.get('date')), reverse=True)

            for fa in f_list_sorted:
                if isinstance(fa, FunnyAnswer):
                    text = fa.text
                    author = fa.author
                    date = fa.date
                    source_doc = getattr(fa, "source_doc", "")
                else:
                    text = fa.get("text", "")
                    author = fa.get("author", "")
                    date = fa.get("date", "")
                    source_doc = fa.get("source_doc", "")

                snippet = (text or "").replace("\n", " ")
                if len(snippet) > 120:
                    snippet = snippet[:117] + "..."

                child = QTreeWidgetItem(q_item)
                child.setText(0, snippet)
                child.setText(1, date or "")
                child.setText(2, author or "")
                child.setText(3, os.path.basename(source_doc) if source_doc else "")
                
                child.setData(0, Qt.UserRole, text) 

            q_item.setExpanded(True)

        for i in range(4):
            self.tree_funny.resizeColumnToContents(i)

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
            text, date_str, author, source_doc = dlg.get_data()

            row = self.table_funny.rowCount()
            self.table_funny.insertRow(row)

            self.table_funny.setItem(row, 0, QTableWidgetItem(text))
            self.table_funny.setItem(row, 1, QTableWidgetItem(date_str))
            self.table_funny.setItem(row, 2, QTableWidgetItem(author))

            # Ve sloupci "Zdroj" zobrazíme jen název souboru,
            # ale do UserRole uložíme plnou cestu
            display_source = os.path.basename(source_doc) if source_doc else ""
            source_item = QTableWidgetItem(display_source)
            source_item.setData(Qt.UserRole, source_doc)
            self.table_funny.setItem(row, 3, source_item)

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
        source_item = self.table_funny.item(row, 3) if self.table_funny.columnCount() > 3 else None

        if not text_item or not date_item or not author_item:
            return

        old_text = text_item.text()
        old_date = date_item.text()
        old_author = author_item.text()

        # Plná cesta je v UserRole, pokud není, použijeme text
        if source_item is not None:
            data = source_item.data(Qt.UserRole)
            if isinstance(data, str) and data:
                old_source = data
            else:
                old_source = source_item.text()
        else:
            old_source = ""

        # Otevření dialogu
        dlg = FunnyAnswerDialog(self, project_root=self.project_root)
        dlg.setWindowTitle("Upravit vtipnou odpověď")
        dlg.set_data(old_text, old_date, old_author, old_source)

        if dlg.exec() == QDialog.Accepted:
            new_text, new_date, new_author, new_source = dlg.get_data()

            # Uložení zpět do tabulky
            self.table_funny.setItem(row, 0, QTableWidgetItem(new_text))
            self.table_funny.setItem(row, 1, QTableWidgetItem(new_date))
            self.table_funny.setItem(row, 2, QTableWidgetItem(new_author))

            display_source = os.path.basename(new_source) if new_source else ""
            new_source_item = QTableWidgetItem(display_source)
            new_source_item.setData(Qt.UserRole, new_source)
            self.table_funny.setItem(row, 3, new_source_item)

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
        title = q.get("title") or self._derive_title_from_html(
            q.get("text_html") or "<p></p>",
            prefix=("BONUS: " if q.get("type") == "bonus" else ""),
        )
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

        # Deserializace vtipných odpovědí (včetně zdrojového dokumentu, pokud je uložen)
        f_answers_raw = q.get("funny_answers", [])
        f_answers: List[FunnyAnswer] = []
        for item in f_answers_raw:
            if isinstance(item, dict):
                f_answers.append(
                    FunnyAnswer(
                        text=item.get("text", ""),
                        author=item.get("author", ""),
                        date=item.get("date", ""),
                        source_doc=item.get("source_doc", ""),
                    )
                )

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
            funny_answers=f_answers,
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
            # NOVÉ: Widgety a jejich labely
            self.edit_correct_answer,
            self.funny_container,
            # NOVÉ: Samostatné labely sekcí
            self.lbl_content,
            self.lbl_correct,
            self.lbl_funny
        ]
        
        for w in widgets:
            if hasattr(self, w.objectName()) or w in widgets: # Check existence
                w.setVisible(visible)
            
        # Skrytí labelů ve form layoutu
        for i in range(self.form_layout.rowCount()):
            item = self.form_layout.itemAt(i, QFormLayout.LabelRole)
            if item and item.widget():
                item.widget().setVisible(visible)
            item = self.form_layout.itemAt(i, QFormLayout.FieldRole)
            if item and item.widget():
                item.widget().setVisible(visible)
        
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

        # Načtení správné odpovědi
        self.edit_correct_answer.setPlainText(q.correct_answer or "")

        # Načtení vtipných odpovědí
        self.table_funny.setRowCount(0)
        # Pojistka pro případ starého JSONu kde funny_answers může být None
        f_answers = getattr(q, "funny_answers", []) or []

        for fa in f_answers:
            # fa může být dict (z JSONu) nebo objekt FunnyAnswer
            if isinstance(fa, FunnyAnswer):
                text = fa.text
                date = fa.date
                author = fa.author
                source_doc = fa.source_doc
            else:
                text = fa.get("text", "")
                date = fa.get("date", "")
                author = fa.get("author", "")
                source_doc = fa.get("source_doc", "")

            row = self.table_funny.rowCount()
            self.table_funny.insertRow(row)
            self.table_funny.setItem(row, 0, QTableWidgetItem(text))
            self.table_funny.setItem(row, 1, QTableWidgetItem(date))
            self.table_funny.setItem(row, 2, QTableWidgetItem(author))

            display_source = os.path.basename(source_doc) if source_doc else ""
            source_item = QTableWidgetItem(display_source)
            source_item.setData(Qt.UserRole, source_doc)
            self.table_funny.setItem(row, 3, source_item)

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
                        q.title = (
                            self.title_edit.text().strip()
                            or self._derive_title_from_html(
                                q.text_html,
                                prefix=("BONUS: " if q.type == "bonus" else ""),
                            )
                        )

                        # Uložení bodů
                        if q.type == "classic":
                            q.points = int(self.spin_points.value())
                            q.bonus_correct = 0.0
                            q.bonus_wrong = 0.0
                        else:
                            q.points = 0
                            q.bonus_correct = round(float(self.spin_bonus_correct.value()), 2)
                            q.bonus_wrong = round(float(self.spin_bonus_wrong.value()), 2)

                        # Uložení správné odpovědi
                        q.correct_answer = self.edit_correct_answer.toPlainText()

                        # Uložení vtipných odpovědí z tabulky (včetně zdrojového dokumentu)
                        new_funny: List[FunnyAnswer] = []
                        for r in range(self.table_funny.rowCount()):
                            t_item = self.table_funny.item(r, 0)
                            if not t_item:
                                continue
                            d_item = self.table_funny.item(r, 1)
                            a_item = self.table_funny.item(r, 2)
                            s_item = self.table_funny.item(r, 3) if self.table_funny.columnCount() > 3 else None

                            text = t_item.text()
                            date = d_item.text() if d_item else ""
                            author = a_item.text() if a_item else ""

                            if s_item is not None:
                                data = s_item.data(Qt.UserRole)
                                if isinstance(data, str) and data:
                                    source_doc = data
                                else:
                                    source_doc = s_item.text()
                            else:
                                source_doc = ""

                            new_funny.append(
                                FunnyAnswer(
                                    text=text,
                                    author=author,
                                    date=date,
                                    source_doc=source_doc,
                                )
                            )
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

                        # 🔁 po každé změně otázky obnovíme přehled vtipných odpovědí
                        self._refresh_funny_answers_tab()
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
        """Přepne aktuální výběr na odrážky s lepším odsazením."""
        cursor = self.text_edit.textCursor()
        cursor.beginEditBlock()
        
        # Zjistíme, zda už jsme v listu (podle prvního bloku ve výběru)
        current_list = cursor.currentList()
        
        if current_list:
            # Zrušit list (nastavit styl na Undefined, což ho odstraní)
            block_fmt = QTextBlockFormat()
            block_fmt.setObjectIndex(-1) # Zruší vazbu na list
            cursor.setBlockFormat(block_fmt)
        else:
            # Vytvořit nový list s lepším formátováním
            list_fmt = QTextListFormat()
            list_fmt.setStyle(QTextListFormat.ListDisc)
            list_fmt.setIndent(1) # Level 1
            
            # Odsazení čísla/odrážky
            # Ve Wordu/Docx to odpovídá Hanging Indent
            
            cursor.createList(list_fmt)
            
            # Aplikovat odsazení bloku pro vizuální shodu
            bf = cursor.blockFormat()
            # Nastavíme levý margin (celé odsuneme) a text indent (první řádek vrátíme zpět pro odrážku)
            # Hodnoty jsou v px/pt (záleží na DPI, ale 20/-15 je rozumný start)
            bf.setLeftMargin(20)
            bf.setTextIndent(-15) 
            cursor.setBlockFormat(bf)
            
        cursor.endEditBlock()
        self._autosave_schedule()
        self.text_edit.setFocus()


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

    def _convert_docx_to_pdf(self, docx_path: Path) -> Optional[Path]:
        """
        Konvertuje DOCX na PDF pomocí LibreOffice (hledá spustitelný soubor i v /Applications).
        Vrací cestu k PDF souboru nebo None pokud selhalo.
        """
        import shutil
        
        # 1. Hledání spustitelného souboru LibreOffice
        lo_candidates = [
            "libreoffice",                                            # Standard Linux/PATH
            "soffice",                                                # Generic bin
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",   # macOS standard path
            "/usr/bin/libreoffice",
            "/usr/local/bin/libreoffice",
            r"C:\Program Files\LibreOffice\program\soffice.exe",      # Windows x64
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe" # Windows x86
        ]
        
        lo_executable = None
        for cand in lo_candidates:
            # shutil.which hledá v PATH, Path(cand).exists() hledá konkrétní soubor
            if shutil.which(cand) or Path(cand).exists():
                lo_executable = cand
                break
        
        if not lo_executable:
             QMessageBox.warning(
                self, 
                "LibreOffice nenalezen",
                "Nemohu najít nainstalovaný LibreOffice.\n"
                "Pokud jej máte nainstalovaný, ujistěte se, že je ve standardní složce "
                "(/Applications/LibreOffice.app na macOS)."
            )
             return None

        # 2. Samotná konverze
        try:
            pdf_path = docx_path.with_suffix('.pdf')
            
            cmd = [
                lo_executable,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(pdf_path.parent),
                str(docx_path)
            ]
            
            # Spuštění procesu
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60
            )
            
            if result.returncode != 0:
                print(f"LibreOffice chyba (Code {result.returncode}):\nSTDERR: {result.stderr}\nSTDOUT: {result.stdout}")
            
            # Ověření, že PDF byl vytvořen
            if pdf_path.exists():
                return pdf_path
            else:
                # Někdy se stane, že returncode je 0, ale soubor nikde (např. sandbox issues)
                return None
                
        except subprocess.TimeoutExpired:
            QMessageBox.warning(self, "Chyba konverze", "Konverze trvala příliš dlouho (timeout).")
            return None
        except Exception as e:
            QMessageBox.warning(self, "Chyba konverze", f"Neočekávaná chyba při konverzi PDF:\n{e}")
            return None

    def _merge_pdfs(self, pdf_paths: List[Path], output_path: Path, cleanup: bool = True) -> bool:
        """
        Spojí více PDF souborů do jednoho.
        Zkouší: PyPDF2 -> macOS join.py -> Ghostscript.
        Pokud vše selže, vyhodí chybovou hlášku s instrukcemi.
        """
        success = False
        missing_tools = []
        
        # 1. Zkusíme PyPDF2 (Preferované)
        try:
            from PyPDF2 import PdfMerger
            merger = PdfMerger()
            for pdf_path in pdf_paths:
                if pdf_path.exists():
                    merger.append(str(pdf_path))
            merger.write(str(output_path))
            merger.close()
            success = True
        except ImportError:
            missing_tools.append("PyPDF2 (pip install PyPDF2)")
            # Pokračujeme na fallback
        except Exception as e:
            print(f"Chyba PyPDF2: {e}")
            # Pokračujeme na fallback

        if not success:
            # 2. Fallback: macOS Native Script nebo Ghostscript
            success = self._merge_pdfs_subprocess(pdf_paths, output_path)
            if not success:
                missing_tools.append("Ghostscript (brew install ghostscript)")

        # Pokud selhalo úplně
        if not success:
            msg = (
                "Nepodařilo se sloučit PDF soubory.\n"
                "Chybí potřebné nástroje.\n\n"
                "Řešení (vyberte jedno):\n"
                "1. Nainstalujte python knihovnu:  pip install PyPDF2\n"
                "2. Nainstalujte Ghostscript:      brew install ghostscript\n\n"
                "Jednotlivé PDF soubory byly ponechány ve složce."
            )
            QMessageBox.warning(self, "Chyba slučování PDF", msg)
            # Vracíme False a NEPROVÁDÍME cleanup, aby uživateli zůstaly aspoň jednotlivé soubory
            return False

        # Vyčištění (pouze při úspěchu)
        if success and cleanup:
            for pdf_path in pdf_paths:
                try:
                    pdf_path.unlink()
                except Exception as e:
                    print(f"Nemohu smazat dočasný PDF {pdf_path}: {e}")
        
        return True


    def _merge_pdfs_subprocess(self, pdf_paths: List[Path], output_path: Path) -> bool:
        """Fallback: slučování PDF pomocí externích nástrojů (macOS join.py nebo GS)."""
        import sys
        
        # A. macOS Built-in Script (Automator)
        # Tento skript je standardně přítomen na macOS
        macos_join_script = "/System/Library/Automator/Combine PDF Pages.action/Contents/Resources/join.py"
        if sys.platform == "darwin" and Path(macos_join_script).exists():
            try:
                cmd = [
                    "python3",  # Použijeme systémový python nebo ten co běží
                    macos_join_script,
                    "-o", str(output_path)
                ]
                cmd.extend([str(p) for p in pdf_paths if p.exists()])
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                if result.returncode == 0:
                    return True
                else:
                    print(f"macOS join.py failed: {result.stderr}")
            except Exception as e:
                print(f"macOS join.py exception: {e}")

        # B. Ghostscript
        try:
            cmd = ['gs', '-q', '-dNOPAUSE', '-dBATCH', '-dSAFER', '-sDEVICE=pdfwrite', f'-sOutputFile={output_path}']
            cmd.extend([str(p) for p in pdf_paths if p.exists()])
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            return result.returncode == 0
            
        except FileNotFoundError:
            # GS není nainstalován
            print("Ghostscript (gs) nebyl nalezen.")
            return False
        except Exception as e:
            print(f"Ghostscript error: {e}")
            return False

    def _generate_docx_from_template(self, template_path: Path, output_path: Path,
                                     simple_repl: Dict[str, str], rich_repl_html: Dict[str, str]) -> None:
        try:
            doc = docx.Document(template_path)
        except Exception as e:
            QMessageBox.critical(self, "Export chyba", f"Nelze otevřít šablonu pomocí python-docx:\n{e}")
            return

        # Helper function updated to accept style info
        def insert_rich_question_block(paragraph, html_content, base_style=None, base_font=None):
            paras_data = parse_html_to_paragraphs(html_content)
            if not paras_data: 
                paragraph.clear()
                return
            
            p_insert = paragraph._p
            
            for i, p_data in enumerate(paras_data):
                if i == 0:
                    new_p = paragraph
                    new_p.clear()
                    # Restore base style for the first paragraph (reused)
                    if base_style: new_p.style = base_style
                else:
                    # Create new paragraph
                    new_p = doc.add_paragraph()
                    # Apply base style from template
                    if base_style: new_p.style = base_style
                    
                    p_insert.addnext(new_p._p)
                    p_insert = new_p._p

                # Apply formatting from HTML
                align = p_data.get('align', 'left')
                if align == 'center': new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif align == 'right': new_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif align == 'justify': new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                else: new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                new_p.paragraph_format.space_before = Pt(0)
                new_p.paragraph_format.space_after = Pt(0)

                indent_lvl = p_data.get('indent', 0)
                
                if p_data.get('prefix'):
                    # Seznam
                    base_indent_pt = 36
                    level_indent = indent_lvl * 36
                    new_p.paragraph_format.left_indent = Pt(base_indent_pt + level_indent)
                    new_p.paragraph_format.first_line_indent = Pt(-18)
                    
                    # Prefix run (bullet/number) - inherits font too
                    r_prefix = new_p.add_run(p_data['prefix'])
                    if base_font:
                        if base_font.get('name'): r_prefix.font.name = base_font['name']
                        if base_font.get('size'): r_prefix.font.size = base_font['size']
                else:
                    # Běžný odstavec
                    if indent_lvl > 0:
                        new_p.paragraph_format.left_indent = Pt(indent_lvl * 36)

                for r_data in p_data['runs']:
                    text_content = r_data['text']
                    parts = text_content.split('\n')
                    for idx, part in enumerate(parts):
                        if part:
                            run = new_p.add_run(part)
                            # 1. Apply base font from template (if exists)
                            if base_font:
                                if base_font.get('name'): run.font.name = base_font['name']
                                if base_font.get('size'): run.font.size = base_font['size']
                                # We don't enforce bold/italic from template if HTML has its own, 
                                # but we could use it as fallback. Here we prioritize HTML format.

                            # 2. Apply HTML format overrides
                            if r_data.get('b'): run.bold = True
                            if r_data.get('i'): run.italic = True
                            if r_data.get('u'): run.underline = True
                            if r_data.get('color'):
                                try:
                                    rgb = r_data['color']
                                    run.font.color.rgb = RGBColor(int(rgb[:2], 16), int(rgb[2:4], 16), int(rgb[4:], 16))
                                except: pass
                        
                        if idx < len(parts) - 1:
                            run = new_p.add_run()
                            run.add_break()

        def process_paragraph(p):
            full_text = p.text
            if not full_text.strip(): return

            # Extract Base Style & Font BEFORE modification
            base_style = p.style
            base_font = {}
            if p.runs:
                # Use the first run as reference for font properties
                r0 = p.runs[0]
                base_font = {
                    'name': r0.font.name,
                    'size': r0.font.size,
                    'bold': r0.bold,
                    'italic': r0.italic,
                    'underline': r0.underline,
                    'color': r0.font.color.rgb if r0.font.color else None
                }

            # 1. Rich Check (Block Replacement)
            txt_clean = full_text.strip()
            matched_rich = None
            for ph, html in rich_repl_html.items():
                if txt_clean == f"<{ph}>" or txt_clean == f"{{{ph}}}":
                    matched_rich = (ph, html)
                    break
            
            if matched_rich:
                # Pass style info to insertion function
                insert_rich_question_block(p, matched_rich[1], base_style=base_style, base_font=base_font)
                return

            # 2. Inline Check
            keys_found = []
            for k in simple_repl.keys():
                if f"<{k}>" in full_text or f"{{{k}}}" in full_text: keys_found.append(k)
            for k in rich_repl_html.keys():
                if f"<{k}>" in full_text or f"{{{k}}}" in full_text: keys_found.append(k)
            
            if not keys_found: return

            has_rich_key = any(k in rich_repl_html for k in keys_found)

            # Safe Replacement (Simple)
            if not has_rich_key:
                replacements_done = 0
                for run in p.runs:
                    t = run.text
                    original_t = t
                    for k in keys_found:
                        if k in simple_repl:
                            val = str(simple_repl[k])
                            t = t.replace(f"<{k}>", val).replace(f"{{{k}}}", val)
                    if t != original_t:
                        run.text = t
                        replacements_done += 1
                
                final_text = p.text
                still_has_keys = any((f"<{k}>" in final_text or f"{{{k}}}" in final_text) for k in keys_found)
                if replacements_done > 0 and not still_has_keys:
                    return 

            # Fallback Reconstruction (Rich/Complex)
            segments = [full_text]
            all_repl_data = {}
            for k, v in simple_repl.items(): all_repl_data[k] = {'type': 'simple', 'val': v}
            for k, html in rich_repl_html.items(): all_repl_data[k] = {'type': 'rich', 'val': html}
            
            for k in keys_found:
                info = all_repl_data[k]
                tokens = [f"<{k}>", f"{{{k}}}"]
                for token in tokens:
                    new_segments = []
                    for seg in segments:
                        if isinstance(seg, str):
                            parts = seg.split(token)
                            for i, part in enumerate(parts):
                                if part: new_segments.append(part)
                                if i < len(parts) - 1: new_segments.append(info)
                        else:
                            new_segments.append(seg)
                    segments = new_segments

            # For inline reconstruction, we rely on base_font captured earlier
            p.clear()
            
            for seg in segments:
                if isinstance(seg, str):
                    run = p.add_run(seg)
                    # Apply base font
                    if base_font.get('name'): run.font.name = base_font['name']
                    if base_font.get('size'): run.font.size = base_font['size']
                    run.bold = base_font.get('bold')
                    run.italic = base_font.get('italic')
                    
                elif isinstance(seg, dict):
                    val = seg['val']
                    if seg['type'] == 'simple':
                        run = p.add_run(val)
                        # Inherit base font
                        if base_font.get('name'): run.font.name = base_font['name']
                        if base_font.get('size'): run.font.size = base_font['size']
                        run.bold = base_font.get('bold')
                        run.italic = base_font.get('italic')
                        run.underline = base_font.get('underline')
                        if base_font.get('color'): run.font.color.rgb = base_font['color']
                        
                    elif seg['type'] == 'rich':
                        # Inline rich text
                        paras = parse_html_to_paragraphs(val)
                        for p_idx, p_data in enumerate(paras):
                            if p_idx > 0: p.add_run().add_break()
                            for r_data in p_data['runs']:
                                text_content = r_data['text']
                                parts = text_content.split('\n')
                                for idx_part, part in enumerate(parts):
                                    if part:
                                        run = p.add_run(part)
                                        # Apply formatting
                                        if r_data.get('b'): run.bold = True
                                        if r_data.get('i'): run.italic = True
                                        if r_data.get('u'): run.underline = True
                                        if r_data.get('color'):
                                            try:
                                                rgb = r_data['color']
                                                run.font.color.rgb = RGBColor(int(rgb[:2], 16), int(rgb[2:4], 16), int(rgb[4:], 16))
                                            except: pass
                                        
                                        # Fallback to base font if not overridden
                                        if not r_data.get('b') and base_font.get('name'): run.font.name = base_font['name']
                                        if not r_data.get('b') and base_font.get('size'): run.font.size = base_font['size']
                                        
                                    if idx_part < len(parts) - 1: p.add_run().add_break()

        for p in doc.paragraphs: process_paragraph(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: process_paragraph(p)
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
