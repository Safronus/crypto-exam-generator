#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Crypto Exam Generator (v1.8e)

Změny v 1.8e
- Export DOCX: 1:1 substituce všech placeholderů i když jsou **rozsekané do více w:t**, ve **všech word/*.xml**.
- Zachování číslování šablony: **nepřepisujeme celý <w:p>**, ale pouze jeho **runs**.
- `<OtázkaX>/<BONUSX>`:
  * INLINE (token je uvnitř textu odstavce) → HTML otázky se vloží jako **runs** do téhož odstavce. Pokud má otázka více odstavců nebo položky listu, vkládají se **<w:br/>** mezi části → **číslování řádku zůstává**.
  * BLOCK (odstavec obsahuje **jen** `<Token>`) → první část otázky nahradí runs v témže `<w:p>` (zachová pPr včetně číslování). Další části otázky se vloží jako **nové odstavce za původní** s čistým pPr (bez numPr).
- Wizard: sken placeholderů přes jednotlivé **odstavce** (w:p), čímž najde i `<BONUS3>` rozsekaný do víc runů.
- Výchozí cesty: šablona `.../data/Šablony/template_AK3KR.docx`, výstup do `.../data/Vygenerované testy/`.
"""
from __future__ import annotations

import sys, uuid as _uuid, html as _html, re, json, zipfile
from dataclasses import dataclass, asdict, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Dict, Tuple

from html.parser import HTMLParser

from PySide6.QtCore import Qt, QSize, QSaveFile, QByteArray, QTimer, QDateTime
from PySide6.QtGui import QAction, QActionGroup, QKeySequence, QTextCharFormat, QTextCursor, QTextListFormat, QTextBlockFormat, QColor, QPalette, QFont, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem, QSplitter, QToolBar,
    QTextEdit, QFileDialog, QMessageBox, QLineEdit, QPushButton, QFormLayout, QSpinBox, QDoubleSpinBox, QComboBox,
    QColorDialog, QAbstractItemView, QDialog, QDialogButtonBox, QLabel, QStyle, QScrollArea, QWizard, QWizardPage,
    QDateTimeEdit
)

APP_NAME = "Crypto Exam Generator"
APP_VERSION = "1.8e"


# ------------------ Datové modely ------------------

@dataclass
class Question:
    id: str
    type: str  # classic|bonus
    text_html: str
    title: str = ""
    points: int = 1
    bonus_correct: float = 0.0
    bonus_wrong: float = 0.0
    created_at: str = ""

    @staticmethod
    def new_default(qtype: str="classic") -> "Question":
        now = datetime.now().isoformat(timespec="seconds")
        if qtype == "bonus":
            return Question(str(_uuid.uuid4()), "bonus", "<p><br></p>", "BONUS otázka", 0, 1.00, 0.00, now)
        return Question(str(_uuid.uuid4()), "classic", "<p><br></p>", "Otázka", 1, 0.0, 0.0, now)


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
    subgroups: List[Subgroup] = field(default_factory=list)


@dataclass
class RootData:
    groups: List[Group] = field(default_factory=list)


# ------------------ Dark theme ------------------

def apply_dark_theme(app: QApplication) -> None:
    QApplication.setStyle("Fusion")
    pal = QPalette()
    pal.setColor(QPalette.Window, QColor(37,37,38))
    pal.setColor(QPalette.WindowText, Qt.white)
    pal.setColor(QPalette.Base, QColor(30,30,30))
    pal.setColor(QPalette.Text, Qt.white)
    pal.setColor(QPalette.Button, QColor(45,45,48))
    pal.setColor(QPalette.ButtonText, Qt.white)
    pal.setColor(QPalette.Highlight, QColor(10,132,255))
    pal.setColor(QPalette.HighlightedText, Qt.black)
    app.setPalette(pal)


# ------------------ Tree s DnD ------------------

class DnDTree(QTreeWidget):
    def __init__(self, owner: "MainWindow") -> None:
        super().__init__()
        self.owner = owner
        self.setHeaderLabels(["Název", "Typ / body"])
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)

    def dropEvent(self, event) -> None:
        ids_before = self.owner._selected_qids()
        super().dropEvent(event)
        self.owner._sync_model_from_tree() if hasattr(self.owner, "_sync_model_from_tree") else None
        self.owner._refresh_tree()
        self.owner._reselect(ids_before)
        self.owner.save_data()
        self.owner.statusBar().showMessage("Přesun dokončen (uloženo).", 3000)


# ------------------ Dialog pro volbu cíle ------------------

class MoveTargetDialog(QDialog):
    def __init__(self, owner: "MainWindow") -> None:
        super().__init__(owner)
        self.setWindowTitle("Vyberte cílovou skupinu/podskupinu")
        self.resize(520, 560)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Vyberte podskupinu (nebo skupinu – vytvoří se Default)."))
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Název", "Typ"])
        lay.addWidget(self.tree, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept); bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        for g in owner.root.groups:
            g_it = QTreeWidgetItem([g.name, "Skupina"])
            g_it.setData(0, Qt.UserRole, {"kind":"group", "id": g.id})
            g_it.setIcon(0, owner.style().standardIcon(QStyle.SP_DirIcon))
            f = g_it.font(0); f.setBold(True); g_it.setFont(0, f)
            self.tree.addTopLevelItem(g_it)
            self._add_subs(owner, g_it, g.id, g.subgroups)
        self.tree.expandAll()

    def _add_subs(self, owner: "MainWindow", parent_item: QTreeWidgetItem, gid: str, subs: List[Subgroup]) -> None:
        for sg in subs:
            it = QTreeWidgetItem([sg.name, "Podskupina"])
            it.setData(0, Qt.UserRole, {"kind":"subgroup","id": sg.id, "parent_group_id": gid})
            it.setIcon(0, owner.style().standardIcon(QStyle.SP_DirOpenIcon))
            parent_item.addChild(it)
            if sg.subgroups:
                self._add_subs(owner, it, gid, sg.subgroups)

    def selected_target(self) -> tuple[str, Optional[str]]:
        items = self.tree.selectedItems()
        if not items: return "", None
        meta = items[0].data(0, Qt.UserRole) or {}
        if meta.get("kind") == "subgroup":
            return meta.get("parent_group_id"), meta.get("id")
        if meta.get("kind") == "group":
            return meta.get("id"), None
        return "", None


# ------------------ Export: HTML parser → meziformát ------------------

class HTMLToBlocks(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.blocks: List[dict] = []
        self.cur_runs: List[dict] = []
        self._in_p = False

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        if t in ("p","div","li"):
            if self._in_p: self._end_p()
            self._in_p = True; self.cur_runs = []
        elif t == "br":
            self.cur_runs.append({'text': "\n", 'b': False, 'i': False, 'u': False, 'color': None})
        elif t in ("b","strong"):
            self.cur_runs.append({'text': "__B_ON__", 'b': True, 'i': None, 'u': None, 'color': None})
        elif t in ("i","em"):
            self.cur_runs.append({'text': "__I_ON__", 'b': None, 'i': True, 'u': None, 'color': None})
        elif t == "u":
            self.cur_runs.append({'text': "__U_ON__", 'b': None, 'i': None, 'u': True, 'color': None})
        elif t == "span":
            style = dict(attrs).get("style","").lower()
            m = re.search(r'color\s*:\s*#?([0-9a-f]{6})', style)
            if m:
                self.cur_runs.append({'text': f"__C_#{m.group(1)}__", 'b': None, 'i': None, 'u': None, 'color': f"#{m.group(1)}"})

    def handle_endtag(self, tag):
        t = tag.lower()
        if t in ("p","div","li"):
            self._end_p()
        elif t in ("b","strong"):
            self.cur_runs.append({'text': "__B_OFF__", 'b': False, 'i': None, 'u': None, 'color': None})
        elif t in ("i","em"):
            self.cur_runs.append({'text': "__I_OFF__", 'b': None, 'i': False, 'u': None, 'color': None})
        elif t == "u":
            self.cur_runs.append({'text': "__U_OFF__", 'b': None, 'i': None, 'u': False, 'color': None})

    def handle_data(self, data):
        if not self._in_p:
            self._in_p = True; self.cur_runs = []
        self.cur_runs.append({'text': data, 'b': None, 'i': None, 'u': None, 'color': None})

    def _end_p(self):
        if not self._in_p: return
        runs = []
        bold = False; italic = False; underline = False; color = None
        for r in self.cur_runs:
            txt = r.get('text') or ""
            if txt == "__B_ON__": bold = True; continue
            if txt == "__B_OFF__": bold = False; continue
            if txt == "__I_ON__": italic = True; continue
            if txt == "__I_OFF__": italic = False; continue
            if txt == "__U_ON__": underline = True; continue
            if txt == "__U_OFF__": underline = False; continue
            m = re.match(r"__C_#([0-9A-Fa-f]{6})__", txt)
            if m: color = f"#{m.group(1)}"; continue
            if not txt: continue
            runs.append({'text': txt, 'b': bold, 'i': italic, 'u': underline, 'color': color})
        self.blocks.append({'align':'left', 'runs': runs})
        self.cur_runs = []; self._in_p = False

def parse_html_blocks(html: str) -> list:
    p = HTMLToBlocks(); p.feed(html or "<p></p>")
    if not p.blocks:
        p._in_p = True; p._end_p()
    return p.blocks


# ------------------ DOCX util ------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
def w(name: str) -> str: return f"{{{W_NS}}}{name}"

def make_run(text: str, b: bool=False, i: bool=False, u: bool=False, color: str|None=None) -> ET.Element:
    r = ET.Element(w("r"))
    rpr = ET.SubElement(r, w("rPr"))
    if b: ET.SubElement(rpr, w("b"))
    if i: ET.SubElement(rpr, w("i"))
    if u:
        uel = ET.SubElement(rpr, w("u"))
        uel.set(w("val"), "single")
    if color:
        cel = ET.SubElement(rpr, w("color"))
        cel.set(w("val"), color.lstrip("#").upper())
    t = ET.SubElement(r, w("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r

def append_br(p: ET.Element) -> None:
    r = ET.SubElement(p, w("r")); ET.SubElement(r, w("br"))

def clear_runs(p: ET.Element) -> None:
    for r in list(p.findall("w:r", NS)): p.remove(r)

def paragraph_text(p: ET.Element) -> str:
    return "".join((t.text or "") for t in p.findall(".//w:t", NS))

def replace_inline_in_p(p: ET.Element, replacements_simple: dict, replacements_rich: dict) -> None:
    text = paragraph_text(p)
    if not text.strip(): return
    def normalize_tokens(s: str) -> str:
        return re.sub(r"<\s*([A-Za-z0-9ÁČĎÉĚÍŇÓŘŠŤÚŮÝŽáčďéěíňóřšťúůýž]+[0-9]*)\s*>", r"<\1>", s)
    text = normalize_tokens(text)
    had_change = False
    for key, val in replacements_simple.items():
        tk = f"<{key}>"
        if tk in text:
            text = text.replace(tk, val); had_change = True
    # rich inline
    for key, blocks in replacements_rich.items():
        tk = f"<{key}>"
        if tk in text:
            prefix, suffix = text.split(tk, 1)
            clear_runs(p)
            if prefix: p.append(make_run(prefix))
            for b_idx, blk in enumerate(blocks):
                for r in blk['runs']:
                    p.append(make_run(r['text'], r['b'], r['i'], r['u'], r['color']))
                if b_idx < len(blocks)-1: append_br(p)
            if suffix: p.append(make_run(suffix))
            text = prefix + "".join(r['text'] for b in blocks for r in b['runs']) + suffix
            had_change = True
    if had_change and not list(p.findall("w:r", NS)):
        clear_runs(p); p.append(make_run(text))

def replace_block_in_p(tree: ET.ElementTree, p: ET.Element, blocks: list) -> None:
    clear_runs(p)
    for r in blocks[0]['runs']:
        p.append(make_run(r['text'], r['b'], r['i'], r['u'], r['color']))
    # další bloky → nové odstavce (bez numPr)
    parent = None
    for elem in tree.getroot().iter():
        for ch in list(elem):
            if ch is p:
                parent = elem; break
        if parent is not None: break
    if parent is None: return
    idx = list(parent).index(p)
    for k in range(1, len(blocks)):
        np = ET.Element(w("p"))
        for r in blocks[k]['runs']:
            np.append(make_run(r['text'], r['b'], r['i'], r['u'], r['color']))
        parent.insert(idx + k, np)

# ------------------ Hlavní okno ------------------

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME); self.resize(1200, 880)
        self.project_root = Path.cwd()
        self.data_path = self.project_root / "data" / "questions.json"
        self.data_path.parent.mkdir(parents=True, exist_ok=True)
        icon_file = self.project_root / "icon" / "icon.png"
        if icon_file.exists():
            app_icon = QIcon(str(icon_file)); self.setWindowIcon(app_icon); QApplication.instance().setWindowIcon(app_icon)
        self.root = RootData([]); self._current_qid: str|None = None
        self._build_ui(); self._connect(); self._build_menus(); self.load_data(); self._refresh_tree()

    # UI
    def _build_ui(self):
        splitter = QSplitter()
        left = QWidget(); lv = QVBoxLayout(left); lv.setContentsMargins(0,0,0,0); lv.setSpacing(6)
        bar = QWidget(); hb = QHBoxLayout(bar); hb.setContentsMargins(6,6,6,0); hb.setSpacing(6)
        self.filter_edit = QLineEdit(); self.filter_edit.setPlaceholderText("Filtr: název / obsah otázky…")
        self.btn_move_selected = QPushButton("Přesunout vybrané…"); self.btn_delete_selected = QPushButton("Smazat vybrané")
        hb.addWidget(self.filter_edit,1); hb.addWidget(self.btn_move_selected); hb.addWidget(self.btn_delete_selected); lv.addWidget(bar)
        self.tree = DnDTree(self); lv.addWidget(self.tree, 1)
        right = QWidget(); rv = QVBoxLayout(right); rv.setContentsMargins(6,6,6,6); rv.setSpacing(8)
        self.editor_tb = QToolBar("Formát"); self.editor_tb.setIconSize(QSize(18,18))
        self.act_b = QAction("Tučné", self); self.act_b.setCheckable(True); self.act_b.setShortcut(QKeySequence.Bold)
        self.act_i = QAction("Kurzíva", self); self.act_i.setCheckable(True); self.act_i.setShortcut(QKeySequence.Italic)
        self.act_u = QAction("Podtržení", self); self.act_u.setCheckable(True); self.act_u.setShortcut(QKeySequence.Underline)
        self.act_color = QAction("Barva", self)
        self.act_bul = QAction("Odrážky", self); self.act_bul.setCheckable(True)
        self.editor_tb.addAction(self.act_b); self.editor_tb.addAction(self.act_i); self.editor_tb.addAction(self.act_u)
        self.editor_tb.addSeparator(); self.editor_tb.addAction(self.act_color); self.editor_tb.addSeparator(); self.editor_tb.addAction(self.act_bul)
        rv.addWidget(self.editor_tb)
        form = QFormLayout()
        self.title_edit = QLineEdit(); self.combo_type = QComboBox(); self.combo_type.addItems(["Klasická","BONUS"])
        self.spin_pts = QSpinBox(); self.spin_pts.setRange(-999,999); self.spin_pts.setValue(1)
        self.spin_bc = QDoubleSpinBox(); self.spin_bc.setDecimals(2); self.spin_bc.setRange(-999.99,999.99); self.spin_bc.setValue(1.00)
        self.spin_bw = QDoubleSpinBox(); self.spin_bw.setDecimals(2); self.spin_bw.setRange(-999.99,999.99); self.spin_bw.setValue(0.00)
        form.addRow("Název:", self.title_edit); form.addRow("Typ:", self.combo_type); form.addRow("Body (klasická):", self.spin_pts); form.addRow("Správně (BONUS):", self.spin_bc); form.addRow("Špatně (BONUS):", self.spin_bw)
        rv.addLayout(form)
        self.text_edit = QTextEdit(); self.text_edit.setAcceptRichText(True); self.text_edit.setMinimumHeight(360); rv.addWidget(self.text_edit, 1)
        self.btn_save_q = QPushButton("Uložit změny otázky"); self.btn_save_q.setDefault(True); rv.addWidget(self.btn_save_q)
        splitter.addWidget(left); splitter.addWidget(right); splitter.setStretchFactor(1,1); self.setCentralWidget(splitter)
        tb = self.addToolBar("Hlavní"); tb.setIconSize(QSize(18,18))
        self.act_add_g = QAction("Přidat skupinu", self); self.act_add_sg = QAction("Přidat podskupinu", self); self.act_add_q = QAction("Přidat otázku", self)
        self.act_del = QAction("Smazat", self); self.act_save_all = QAction("Uložit vše", self); self.act_choose_data = QAction("Zvolit soubor s daty…", self)
        tb.addAction(self.act_add_g); tb.addAction(self.act_add_sg); tb.addAction(self.act_add_q); tb.addSeparator(); tb.addAction(self.act_del); tb.addSeparator(); tb.addAction(self.act_save_all); tb.addSeparator(); tb.addAction(self.act_choose_data)
        self.statusBar().showMessage(f"Datový soubor: {self.data_path}")

    def _connect(self):
        self.tree.itemSelectionChanged.connect(self._on_tree_sel)
        self.btn_save_q.clicked.connect(lambda: self._save_current_q(False))
        self.act_add_g.triggered.connect(self._add_group)
        self.act_add_sg.triggered.connect(self._add_subgroup)
        self.act_add_q.triggered.connect(self._add_question)
        self.act_del.triggered.connect(self._delete_selected)
        self.act_save_all.triggered.connect(self.save_data)
        self.act_choose_data.triggered.connect(self._choose_data_file)
        self.filter_edit.textChanged.connect(self._apply_filter)

    def _build_menus(self):
        bar = self.menuBar(); m_file = bar.addMenu("Soubor"); m_edit = bar.addMenu("Úpravy")
        self.act_import_docx = QAction("Import z DOCX…", self); self.act_export_docx = QAction("Export do DOCX (šablona)…", self)
        self.act_move_selected = QAction("Přesunout vybrané…", self); self.act_delete_selected = QAction("Smazat vybrané", self)
        m_file.addAction(self.act_import_docx); m_file.addAction(self.act_export_docx); m_edit.addAction(self.act_move_selected); m_edit.addAction(self.act_delete_selected)
        self.act_import_docx.setShortcut("Ctrl+I"); self.act_import_docx.triggered.connect(self._import_from_docx)
        self.act_export_docx.triggered.connect(self._export_docx_wizard)
        self.act_move_selected.triggered.connect(self._bulk_move_selected); self.act_delete_selected.triggered.connect(self._bulk_delete_selected)

    # Data
    def load_data(self):
        if self.data_path.exists():
            try:
                raw = json.loads(self.data_path.read_text(encoding="utf-8"))
                groups = []
                for g in raw.get("groups", []):
                    groups.append(self._parse_group(g))
                self.root = RootData(groups)
            except Exception:
                self.root = RootData([])
        else:
            self.root = RootData([])

    def save_data(self):
        self._save_current_q(silent=True)
        self.data_path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps({"groups": [self._ser_group(g) for g in self.root.groups]}, ensure_ascii=False, indent=2)
        sf = QSaveFile(str(self.data_path)); sf.open(QSaveFile.WriteOnly); sf.write(QByteArray(payload.encode("utf-8"))); sf.commit()
        self.statusBar().showMessage("Uloženo.", 1200)

    def _parse_group(self, g: dict) -> Group:
        return Group(g["id"], g["name"], [self._parse_subgroup(s) for s in g.get("subgroups",[])])

    def _parse_subgroup(self, sg: dict) -> Subgroup:
        return Subgroup(sg["id"], sg["name"], [self._parse_subgroup(s) for s in sg.get("subgroups",[])], [self._parse_question(q) for q in sg.get("questions",[])])

    def _parse_question(self, q: dict) -> Question:
        def safef(x, d):
            try: return round(float(x),2)
            except: return d
        title = q.get("title") or self._derive_title(q.get("text_html",""))
        return Question(q.get("id",""), q.get("type","classic"), q.get("text_html","<p><br></p>"), title, int(q.get("points",1)), safef(q.get("bonus_correct",1.0 if q.get('type')=='bonus' else 0.0),1.0), safef(q.get("bonus_wrong",0.0),0.0), q.get("created_at",""))

    def _ser_group(self, g: Group) -> dict:
        return {"id": g.id, "name": g.name, "subgroups": [self._ser_subgroup(s) for s in g.subgroups]}

    def _ser_subgroup(self, sg: Subgroup) -> dict:
        return {"id": sg.id, "name": sg.name, "subgroups": [self._ser_subgroup(s) for s in sg.subgroups], "questions": [asdict(q) for q in sg.questions]}

    # Strom
    def _refresh_tree(self):
        self.tree.clear()
        for g in self.root.groups:
            gi = QTreeWidgetItem([g.name, "Skupina"])
            gi.setData(0, Qt.UserRole, {"kind":"group","id": g.id})
            gi.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
            self.tree.addTopLevelItem(gi)
            self._add_subs_to_item(gi, g.id, g.subgroups)
        self.tree.expandAll()

    def _add_subs_to_item(self, parent: QTreeWidgetItem, gid: str, subs: List[Subgroup]):
        for sg in subs:
            it = QTreeWidgetItem([sg.name, "Podskupina"])
            it.setData(0, Qt.UserRole, {"kind":"subgroup","id": sg.id, "parent_group_id": gid})
            it.setIcon(0, self.style().standardIcon(QStyle.SP_DirOpenIcon))
            parent.addChild(it)
            for q in sg.questions:
                label = "Klasická" if q.type=="classic" else "BONUS"
                pts = str(q.points) if q.type=="classic" else f"+{q.bonus_correct:.2f}/ {q.bonus_wrong:.2f}"
                qi = QTreeWidgetItem([q.title or "Otázka", f"{label} | {pts}"])
                qi.setData(0, Qt.UserRole, {"kind":"question","id": q.id, "parent_group_id": gid, "parent_subgroup_id": sg.id})
                parent.child(parent.indexOfChild(it)).addChild(qi) if False else it.addChild(qi)
            if sg.subgroups:
                self._add_subs_to_item(it, gid, sg.subgroups)

    def _selected_qids(self) -> List[str]:
        out = []
        for it in self.tree.selectedItems():
            meta = it.data(0, Qt.UserRole) or {}
            if meta.get("kind")=="question": out.append(meta.get("id"))
        return out

    def _reselect(self, ids: List[str]) -> None:
        if not ids: return
        want = set(ids)
        def walk(item: QTreeWidgetItem):
            meta = item.data(0, Qt.UserRole) or {}
            if meta.get("kind")=="question" and meta.get("id") in want: item.setSelected(True)
            for i in range(item.childCount()): walk(item.child(i))
        for i in range(self.tree.topLevelItemCount()): walk(self.tree.topLevelItem(i))

    def _on_tree_sel(self):
        its = self.tree.selectedItems()
        if not its: return
        meta = its[0].data(0, Qt.UserRole) or {}
        if meta.get("kind")=="question":
            q = self._find_q(meta.get("parent_group_id"), meta.get("parent_subgroup_id"), meta.get("id"))
            if q: self._load_q(q)

    # CRUD a editor
    def _add_group(self):
        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(self, "Nová skupina", "Název:")
        if not ok or not name.strip(): return
        self.root.groups.append(Group(str(_uuid.uuid4()), name.strip(), []))
        self._refresh_tree(); self.save_data()

    def _add_subgroup(self):
        from PySide6.QtWidgets import QInputDialog
        its = self.tree.selectedItems()
        if not its: return
        meta = its[0].data(0, Qt.UserRole) or {}
        if meta.get("kind") not in ("group","subgroup"):
            QMessageBox.information(self, "Podskupina", "Vyberte skupinu/podskupinu."); return
        name, ok = QInputDialog.getText(self, "Nová podskupina", "Název:")
        if not ok or not name.strip(): return
        if meta.get("kind")=="group":
            g = self._find_g(meta["id"]); 
            if g: g.subgroups.append(Subgroup(str(_uuid.uuid4()), name.strip(), [], []))
        else:
            sg = self._find_sg(meta["parent_group_id"], meta["id"])
            if sg: sg.subgroups.append(Subgroup(str(_uuid.uuid4()), name.strip(), [], []))
        self._refresh_tree(); self.save_data()

    def _add_question(self):
        its = self.tree.selectedItems()
        if not its: 
            QMessageBox.information(self, "Otázka", "Vyberte skupinu/podskupinu."); return
        meta = its[0].data(0, Qt.UserRole) or {}
        if meta.get("kind") not in ("group","subgroup"):
            QMessageBox.information(self, "Otázka", "Vyberte skupinu/podskupinu."); return
        if meta.get("kind")=="group":
            g = self._find_g(meta["id"]); 
            if not g: return
            if not g.subgroups:
                g.subgroups.append(Subgroup(str(_uuid.uuid4()), "Default", [], []))
            sg = g.subgroups[0]
        else:
            sg = self._find_sg(meta["parent_group_id"], meta["id"])
        if not sg: return
        q = Question.new_default("classic"); sg.questions.append(q)
        self._refresh_tree(); self.save_data()

    def _delete_selected(self):
        its = self.tree.selectedItems()
        if not its: return
        meta = its[0].data(0, Qt.UserRole) or {}
        if meta.get("kind")!="question": return
        gid, sgid, qid = meta.get("parent_group_id"), meta.get("parent_subgroup_id"), meta.get("id")
        sg = self._find_sg(gid, sgid); 
        if not sg: return
        sg.questions = [qq for qq in sg.questions if qq.id != qid]
        self._refresh_tree(); self.save_data()

    def _save_current_q(self, silent: bool=True):
        if not self._current_qid: return
        def apply_in(subs: List[Subgroup]) -> bool:
            for sg in subs:
                for i, q in enumerate(sg.questions):
                    if q.id == self._current_qid:
                        q.title = self.title_edit.text().strip() or self._derive_title(self.text_edit.toHtml())
                        q.type = "classic" if self.combo_type.currentIndex()==0 else "bonus"
                        if q.type=="classic":
                            q.points = int(self.spin_pts.value()); q.bonus_correct = 0.0; q.bonus_wrong = 0.0
                        else:
                            q.points = 0; q.bonus_correct = round(float(self.spin_bc.value()),2); q.bonus_wrong = round(float(self.spin_bw.value()),2)
                        q.text_html = self.text_edit.toHtml()
                        sg.questions[i] = q; return True
                if apply_in(sg.subgroups): return True
            return False
        for g in self.root.groups:
            if apply_in(g.subgroups): break
        if not silent: self.statusBar().showMessage("Otázka uložena.", 1200)

    def _find_g(self, gid: str) -> Optional[Group]:
        return next((g for g in self.root.groups if g.id==gid), None)

    def _find_sg(self, gid: str, sgid: str) -> Optional[Subgroup]:
        g = self._find_g(gid)
        if not g: return None
        def walk(lst: List[Subgroup]):
            for sg in lst:
                if sg.id == sgid: return sg
                r = walk(sg.subgroups)
                if r: return r
            return None
        return walk(g.subgroups)

    def _find_q(self, gid: str, sgid: str, qid: str) -> Optional[Question]:
        sg = self._find_sg(gid, sgid)
        if not sg: return None
        return next((q for q in sg.questions if q.id==qid), None)

    def _all_by_type(self, qtype: str) -> List[Question]:
        out = []
        def walk(lst: List[Subgroup]):
            for sg in lst:
                for q in sg.questions:
                    if q.type == qtype: out.append(q)
                walk(sg.subgroups)
        for g in self.root.groups: walk(g.subgroups)
        return out

    def _load_q(self, q: Question):
        self._current_qid = q.id
        self.title_edit.setText(q.title or self._derive_title(q.text_html))
        self.combo_type.setCurrentIndex(0 if q.type=="classic" else 1)
        self.spin_pts.setValue(int(q.points))
        self.spin_bc.setValue(float(q.bonus_correct))
        self.spin_bw.setValue(float(q.bonus_wrong))
        self.text_edit.setHtml(q.text_html or "<p><br></p>")

    def _derive_title(self, html: str) -> str:
        txt = re.sub(r'<[^>]+>', ' ', html or '')
        txt = _html.unescape(txt).strip()
        if not txt: return "Otázka"
        head = re.split(r'[.!?]\s', txt)[0] or txt
        return head[:80] + ('…' if len(head)>80 else '')

    # Import DOCX (zjednodušený parser, otázky 1. … a BONUS "Otázka N" / text 'BONUS')
    def _import_from_docx(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Import z DOCX", str(self.project_root), "Word dokument (*.docx)")
        if not paths: return
        gid, sgid = self._ensure_unassigned()
        target = self._find_sg(gid, sgid)
        total = 0
        for p in paths:
            try:
                with zipfile.ZipFile(p, "r") as z:
                    xml = z.read("word/document.xml")
                root = ET.fromstring(xml)
                paras = ["".join((t.text or "") for t in p.findall(".//w:t", NS)).strip() for p in root.findall(".//w:p", NS)]
                rx_bonus = re.compile(r'^\s*Otázka\s+\d+|bonus', re.IGNORECASE)
                rx_classic = re.compile(r'^\s*\d+[\.)]\s')
                for t in paras:
                    if not t: continue
                    if rx_bonus.search(t):
                        q = Question.new_default("bonus"); q.text_html = f"<p>{_html.escape(t)}</p>"; q.title = self._derive_title(q.text_html)
                        if target: target.questions.append(q); total += 1
                    elif rx_classic.match(t):
                        q = Question.new_default("classic"); q.text_html = f"<p>{_html.escape(t)}</p>"; q.title = self._derive_title(q.text_html)
                        if target: target.questions.append(q); total += 1
            except Exception as e:
                QMessageBox.warning(self, "Import", f"{p}\n{e}")
        self._refresh_tree(); self.save_data()
        self.statusBar().showMessage(f"Import hotov: {total} otázek do 'Neroztříděné'.", 6000)

    def _ensure_unassigned(self) -> tuple[str,str]:
        g = next((g for g in self.root.groups if g.name=="Neroztříděné"), None)
        if not g:
            g = Group(str(_uuid.uuid4()), "Neroztříděné", []); self.root.groups.append(g)
        if not g.subgroups:
            g.subgroups.append(Subgroup(str(_uuid.uuid4()), "Default", [], []))
        return g.id, g.subgroups[0].id

    # Přesuny
    def _bulk_move_selected(self):
        from PySide6.QtWidgets import QInputDialog
        items = [it for it in self.tree.selectedItems() if (it.data(0, Qt.UserRole) or {}).get("kind")=="question"]
        if not items: 
            QMessageBox.information(self, "Přesun", "Vyberte alespoň jednu otázku."); return
        dlg = MoveTargetDialog(self)
        if dlg.exec() != QDialog.Accepted: return
        gid, sgid = dlg.selected_target()
        g = self._find_g(gid); target = self._find_sg(gid, sgid) if g else None
        if not target:
            if g and not g.subgroups: g.subgroups.append(Subgroup(str(_uuid.uuid4()), "Default", [], []))
            target = g.subgroups[0] if g else None
        moved = 0
        for it in items:
            meta = it.data(0, Qt.UserRole) or {}
            src = self._find_sg(meta.get("parent_group_id"), meta.get("parent_subgroup_id"))
            q = self._find_q(meta.get("parent_group_id"), meta.get("parent_subgroup_id"), meta.get("id"))
            if src and q:
                src.questions = [qq for qq in src.questions if qq.id != q.id]
                target.questions.append(q); moved += 1
        self._refresh_tree(); self.save_data()
        self.statusBar().showMessage(f"Přesunuto {moved} otázek.", 4000)

    # Export (default cesty + robustní náhrada)
    def _export_docx_wizard(self):
        default_template = self.project_root / "data" / "Šablony" / "template_AK3KR.docx"
        default_out_dir = self.project_root / "data" / "Vygenerované testy"
        default_out_dir.mkdir(parents=True, exist_ok=True)
        default_out = default_out_dir / f"Test_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"

        templ_path, _ = QFileDialog.getOpenFileName(self, "Zvolte šablonu DOCX", str(default_template if default_template.exists() else self.project_root), "Word dokument (*.docx)")
        if not templ_path:
            templ_path = str(default_template)
        out_path, _ = QFileDialog.getSaveFileName(self, "Uložit výstup DOCX", str(default_out), "Word dokument (*.docx)")
        if not out_path:
            out_path = str(default_out)

        try:
            ph = self._scan_placeholders(Path(templ_path))
        except Exception as e:
            QMessageBox.critical(self, "Šablona", f"Nelze číst šablonu:\n{e}"); return

        classics = self._all_by_type("classic")
        bonuses  = self._all_by_type("bonus")
        repl_simple: Dict[str,str] = {}
        repl_rich: Dict[str,str] = {}

        if "PoznamkaVerze" in ph["any"]:
            repl_simple["PoznamkaVerze"] = f"MůjText_{datetime.now().strftime('%Y-%m-%d')}_{str(_uuid.uuid4())[:8]}"
        if "DatumČas" in ph["any"]:
            repl_simple["DatumČas"] = datetime.now().strftime("%A %d.%m.%Y %H:%M")

        for name in sorted([t for t in ph["q"]], key=lambda s: int(re.findall(r"\d+", s)[0])):
            idx = int(re.findall(r"\d+", name)[0]) - 1
            if idx < len(classics): repl_rich[name] = classics[idx].text_html
        for name in sorted([t for t in ph["b"]], key=lambda s: int(re.findall(r"\d+", s)[0])):
            idx = int(re.findall(r"\d+", name)[0]) - 1
            if idx < len(bonuses): repl_rich[name] = bonuses[idx].text_html

        if "MinBody" in ph["any"]:
            repl_simple["MinBody"] = f"{sum(float(q.bonus_wrong) for q in bonuses):.2f}"
        if "MaxBody" in ph["any"]:
            repl_simple["MaxBody"] = f"{sum(int(q.points) for q in classics) + sum(float(q.bonus_correct) for q in bonuses):.2f}"

        try:
            self._generate_from_template(Path(templ_path), Path(out_path), repl_simple, repl_rich)
        except Exception as e:
            QMessageBox.critical(self, "Export", f"Chyba při exportu:\n{e}"); return
        QMessageBox.information(self, "Export", f"Vygenerováno:\n{out_path}")

    def _scan_placeholders(self, template: Path) -> Dict[str, set]:
        with zipfile.ZipFile(template, "r") as z:
            texts = []
            for name in z.namelist():
                if name.startswith("word/") and name.endswith(".xml"):
                    try:
                        tree = ET.ElementTree(ET.fromstring(z.read(name).decode("utf-8", errors="ignore")))
                    except Exception:
                        continue
                    for p in tree.getroot().findall(".//w:p", NS):
                        s = "".join((t.text or "") for t in p.findall(".//w:t", NS))
                        s = re.sub(r"<\s*([A-Za-z0-9ÁČĎÉĚÍŇÓŘŠŤÚŮÝŽáčďéěíňóřšťúůýž]+[0-9]*)\s*>", r"<\1>", s)
                        if s.strip(): texts.append(s)
        all_tokens = set()
        for s in texts:
            for m in re.findall(r"<([A-Za-z0-9ÁČĎÉĚÍŇÓŘŠŤÚŮÝŽáčďéěíňóřšťúůýž]+[0-9]*)>", s):
                all_tokens.add(m)
        qs = {t for t in all_tokens if re.match(r"^Otázka\d+$", t)}
        bs = {t for t in all_tokens if re.match(r"^BONUS\d+$", t)}
        return {"q": qs, "b": bs, "any": all_tokens}

    def _generate_from_template(self, template: Path, out: Path, simple: Dict[str,str], rich_html: Dict[str,str]) -> None:
        rich_blocks = {k: parse_html_blocks(v) for k,v in rich_html.items()}
        def normalize(s: str) -> str:
            return re.sub(r"<\s*([A-Za-z0-9ÁČĎÉĚÍŇÓŘŠŤÚŮÝŽáčďéěíňóřšťúůýž]+[0-9]*)\s*>", r"<\1>", s or "")
        with zipfile.ZipFile(template, "r") as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        tree = ET.ElementTree(ET.fromstring(data.decode("utf-8")))
                        root = tree.getroot()
                        for p in list(root.findall(".//w:p", NS)):
                            s = normalize(paragraph_text(p).strip())
                            if not s: continue
                            m = re.fullmatch(r"<([A-Za-z0-9ÁČĎÉĚÍŇÓŘŠŤÚŮÝŽáčďéěíňóřšťúůýž]+[0-9]*)>", s)
                            if m:
                                name = m.group(1)
                                if name in rich_blocks:
                                    replace_block_in_p(tree, p, rich_blocks[name]); continue
                                if name in simple:
                                    clear_runs(p); p.append(make_run(simple[name])); continue
                            replace_inline_in_p(p, simple, rich_blocks)
                        data = ET.tostring(root, encoding="utf-8", xml_declaration=False)
                    except Exception as e:
                        pass
                zout.writestr(item, data)

# main
def main() -> int:
    app = QApplication(sys.argv); apply_dark_theme(app)
    icon_file = Path.cwd() / "icon" / "icon.png"
    if icon_file.exists(): app.setWindowIcon(QIcon(str(icon_file)))
    w = MainWindow(); w.show(); return app.exec()

if __name__ == "__main__":
    sys.exit(main())
