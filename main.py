#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Crypto Exam Generator (v1.2)
- Hierarchie podskupin libovolné hloubky (jako složky)
- Drag & drop přesouvání a řazení podskupin i otázek v levém stromu
- Import z DOCX, editor formátování, JSON úložiště

Platforma: macOS (podporováno i jinde), výchozí dark theme (Fusion).
Autor: (doplní uživatel)
Licence: MIT nebo dle potřeby
"""

from __future__ import annotations

import json
import sys
import uuid
import re
import html
import zipfile
from xml.etree import ElementTree as ET
from dataclasses import dataclass, asdict, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from PySide6.QtCore import Qt, QSize, QSaveFile, QByteArray
from PySide6.QtGui import QAction, QKeySequence, QTextCharFormat, QTextCursor, QTextListFormat, QColor, QPalette
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
    QComboBox,
    QColorDialog,
    QAbstractItemView,
)


APP_NAME = "Crypto Exam Generator"
APP_VERSION = "1.3"  # minor: nested subgroups + drag&drop

# --------------------------- Datové typy ---------------------------

@dataclass
class Question:
    id: str
    type: str  # "classic" | "bonus"
    text_html: str
    points: int = 1          # pouze pro classic
    bonus_correct: int = 1   # pouze pro bonus
    bonus_wrong: int = 0     # pouze pro bonus (může být záporné)
    created_at: str = ""     # ISO

    @staticmethod
    def new_default(qtype: str = "classic") -> "Question":
        now = datetime.now().isoformat(timespec="seconds")
        if qtype == "bonus":
            return Question(
                id=str(uuid.uuid4()),
                type="bonus",
                text_html="<p><br></p>",
                points=0,
                bonus_correct=1,
                bonus_wrong=0,
                created_at=now,
            )
        return Question(
            id=str(uuid.uuid4()),
            type="classic",
            text_html="<p><br></p>",
            points=1,
            bonus_correct=0,
            bonus_wrong=0,
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
    """Nastaví Fusion dark theme s jemnými barvami, vhodné pro macOS/Retina."""
    app.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
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
    # v1.3: multi-select zapnuto (ExtendedSelection)
    """QTreeWidget s podporou drag&drop, po přesunu synchronizuje datový model."""

    def __init__(self, owner: "MainWindow") -> None:
        super().__init__()
        self.owner = owner
        self.setHeaderLabels(["Název", "Typ / body"])
        self.setColumnWidth(0, 280)
        self.setContextMenuPolicy(Qt.CustomContextMenu)

        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setDragDropMode(QAbstractItemView.InternalMove)

    def dropEvent(self, event) -> None:
        super().dropEvent(event)
        # Po vizuálním přesunu stromu přegenerujeme datový model z aktuální struktury
        self.owner._sync_model_from_tree()
        self.owner.save_data()
        self.owner.statusBar().showMessage("Přesun dokončen (uloženo).", 3000)


# --------------------------- Hlavní okno ---------------------------

class MainWindow(QMainWindow):
    """Hlavní okno aplikace."""

    def __init__(self, data_path: Optional[Path] = None) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1200, 800)

        # cesta k datům (JSON)
        self.project_root = Path.cwd()
        default_data_dir = self.project_root / "data"
        default_data_dir.mkdir(parents=True, exist_ok=True)
        self.data_path = data_path or (default_data_dir / "questions.json")

        self.root: RootData = RootData(groups=[])
        self._current_question_id: Optional[str] = None
        self._current_node_kind: Optional[str] = None  # "group"|"subgroup"|"question"

        self._build_ui()
        self._connect_signals()
        self._build_menus()
        self.load_data()
        self._refresh_tree()

    # -------------------- UI konstrukce --------------------

    def _build_ui(self) -> None:
        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)

        # Levý panel: filtrace + strom
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        # Filtr a hromadné akce
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

        # Strom
        self.tree = DnDTree(self)
        left_layout.addWidget(self.tree, 1)

        # Pravý panel: editor/props
        self.detail_stack = QWidget()
        self.detail_layout = QVBoxLayout(self.detail_stack)
        self.detail_layout.setContentsMargins(6, 6, 6, 6)
        self.detail_layout.setSpacing(8)

        # --- Toolbar pro rich text
        self.editor_toolbar = QToolBar("Formát")
        self.editor_toolbar.setIconSize(QSize(18, 18))
        self.action_bold = QAction("Tučné", self)
        self.action_bold.setCheckable(True)
        self.action_bold.setShortcut(QKeySequence.Bold)
        self.action_italic = QAction("Kurzíva", self)
        self.action_italic.setCheckable(True)
        self.action_italic.setShortcut(QKeySequence.Italic)
        self.action_underline = QAction("Podtržení", self)
        self.action_underline.setCheckable(True)
        self.action_underline.setShortcut(QKeySequence.Underline)
        self.action_color = QAction("Barva", self)
        self.action_bullets = QAction("Odrážky", self)
        self.action_bullets.setCheckable(True)

        self.editor_toolbar.addAction(self.action_bold)
        self.editor_toolbar.addAction(self.action_italic)
        self.editor_toolbar.addAction(self.action_underline)
        self.editor_toolbar.addSeparator()
        self.editor_toolbar.addAction(self.action_color)
        self.editor_toolbar.addSeparator()
        self.editor_toolbar.addAction(self.action_bullets)

        # --- Panel s typem a body
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)

        self.combo_type = QComboBox()
        self.combo_type.addItems(["Klasická", "BONUS"])  # mapujeme ručně

        self.spin_points = QSpinBox()
        self.spin_points.setRange(-999, 999)
        self.spin_points.setValue(1)

        self.spin_bonus_correct = QSpinBox()
        self.spin_bonus_correct.setRange(-999, 999)
        self.spin_bonus_correct.setValue(1)

        self.spin_bonus_wrong = QSpinBox()
        self.spin_bonus_wrong.setRange(-999, 999)
        self.spin_bonus_wrong.setValue(0)

        form.addRow("Typ otázky:", self.combo_type)
        form.addRow("Body (klasická):", self.spin_points)
        form.addRow("Body za správně (BONUS):", self.spin_bonus_correct)
        form.addRow("Body za špatně (BONUS):", self.spin_bonus_wrong)

        # --- Rich text editor
        self.text_edit = QTextEdit()
        self.text_edit.setAcceptRichText(True)
        self.text_edit.setPlaceholderText("Sem napište znění otázky…\nPodporováno: tučné, kurzíva, podtržení, barva, odrážky.")
        self.text_edit.setMinimumHeight(200)

        # --- Tlačítka akce
        self.btn_save_question = QPushButton("Uložit změny otázky")
        self.btn_save_question.setDefault(True)

        # --- Panel pro přejmenování skupiny/podskupiny
        self.rename_panel = QWidget()
        rename_layout = QFormLayout(self.rename_panel)
        self.rename_line = QLineEdit()
        self.btn_rename = QPushButton("Uložit název")
        rename_layout.addRow("Název:", self.rename_line)
        rename_layout.addRow(self.btn_rename)

        # Skládání pravého panelu
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

        # Hlavní toolbar: přidání/mazání/uložení
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
        self.combo_type.setEnabled(enabled)
        self.spin_points.setEnabled(enabled)
        self.spin_bonus_correct.setEnabled(enabled)
        self.spin_bonus_wrong.setEnabled(enabled)
        self.text_edit.setEnabled(enabled)
        self.btn_save_question.setEnabled(enabled)

    def _connect_signals(self) -> None:
        self.tree.itemSelectionChanged.connect(self._on_tree_selection_changed)
        self.btn_save_question.clicked.connect(self._on_save_question_clicked)
        self.btn_rename.clicked.connect(self._on_rename_clicked)
        self.combo_type.currentIndexChanged.connect(self._on_type_changed_ui)

        # Rich text actions
        self.action_bold.triggered.connect(lambda: self._toggle_format("bold"))
        self.action_italic.triggered.connect(lambda: self._toggle_format("italic"))
        self.action_underline.triggered.connect(lambda: self._toggle_format("underline"))
        self.action_color.triggered.connect(self._choose_color)
        self.action_bullets.triggered.connect(self._toggle_bullets)
        self.text_edit.cursorPositionChanged.connect(self._sync_toolbar_to_cursor)

        # Toolbar actions
        self.act_add_group.triggered.connect(self._add_group)
        self.act_add_subgroup.triggered.connect(self._add_subgroup)
        self.act_add_question.triggered.connect(self._add_question)
        self.act_delete.triggered.connect(self._delete_selected)
        self.act_save_all.triggered.connect(self.save_data)
        self.act_choose_data.triggered.connect(self._choose_data_file)
        # Filtr + hromadné akce
        self.filter_edit.textChanged.connect(self._apply_filter)
        self.btn_move_selected.clicked.connect(self._bulk_move_selected)
        self.btn_delete_selected.clicked.connect(self._bulk_delete_selected)


    # -------------------- Menu (import/přesun) --------------------

    def _build_menus(self) -> None:
        bar = self.menuBar()
        file_menu = bar.addMenu("Soubor")
        edit_menu = bar.addMenu("Úpravy")

        self.act_import_docx = QAction("Import z DOCX…", self)
        self.act_move_question = QAction("Přesunout otázku…", self)
        self.act_move_selected = QAction("Přesunout vybrané…", self)
        self.act_delete_selected = QAction("Smazat vybrané", self)

        # Zkratka pro rychlý import
        self.act_import_docx.setShortcut("Ctrl+I")

        file_menu.addAction(self.act_import_docx)
        edit_menu.addAction(self.act_move_question)
        edit_menu.addAction(self.act_move_selected)
        edit_menu.addAction(self.act_delete_selected)

        self.act_import_docx.triggered.connect(self._import_from_docx)
        self.act_move_question.triggered.connect(self._move_question)

        # Viditelné tlačítko do toolbaru (pro snazší nalezení)
        tb_import = self.addToolBar("Import")
        tb_import.setIconSize(QSize(18, 18))
        tb_import.addAction(self.act_import_docx)

        self.act_move_selected.triggered.connect(self._bulk_move_selected)
        self.act_delete_selected.triggered.connect(self._bulk_delete_selected)


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
        # promítnout rozpracované změny otázky
        self._apply_editor_to_current_question(silent=True)

        self.data_path.parent.mkdir(parents=True, exist_ok=True)
        data = {"groups": [self._serialize_group(g) for g in self.root.groups]}

        try:
            sf = QSaveFile(str(self.data_path))
            sf.open(QSaveFile.WriteOnly)
            payload = json.dumps(data, ensure_ascii=False, indent=2)
            sf.write(QByteArray(payload.encode("utf-8")))
            sf.commit()
            self.statusBar().showMessage(f"Uloženo: {self.data_path}", 3000)
        except Exception as e:
            QMessageBox.critical(self, "Uložení selhalo", f"Chyba při ukládání do {self.data_path}:\n{e}")

    # ---- JSON parse/serialize helpers (rekurzivně) ----

    def _parse_group(self, g: dict) -> Group:
        subgroups = [self._parse_subgroup(sg) for sg in g.get("subgroups", [])]
        return Group(id=g["id"], name=g["name"], subgroups=subgroups)

    def _parse_subgroup(self, sg: dict) -> Subgroup:
        # kompatibilita se starou strukturou (bez "subgroups")
        subgroups_raw = sg.get("subgroups", [])
        subgroups = [self._parse_subgroup(s) for s in subgroups_raw]
        questions = []
        for q in sg.get("questions", []):
            questions.append(Question(**q))
        return Subgroup(id=sg["id"], name=sg["name"], subgroups=subgroups, questions=questions)

    def _serialize_group(self, g: Group) -> dict:
        return {
            "id": g.id,
            "name": g.name,
            "subgroups": [self._serialize_subgroup(sg) for sg in g.subgroups],
        }

    def _serialize_subgroup(self, sg: Subgroup) -> dict:
        return {
            "id": sg.id,
            "name": sg.name,
            "subgroups": [self._serialize_subgroup(s) for s in sg.subgroups],
            "questions": [asdict(q) for q in sg.questions],
        }

    # -------------------- Filtr --------------------
    def _apply_filter(self, text: str) -> None:
        pat = (text or '').strip().lower()
        def question_matches(qid: str) -> bool:
            q = None
            # najdi otázku podle id
            for g in self.root.groups:
                stack = list(g.subgroups)
                while stack:
                    sg = stack.pop()
                    for qq in sg.questions:
                        if qq.id == qid:
                            q = qq
                            break
                    if q: break
                    stack.extend(sg.subgroups)
                if q: break
            if not q:
                return False
            import re, html as _html
            plain = re.sub(r'<[^>]+>', ' ', q.text_html)
            plain = _html.unescape(plain).lower()
            return pat in plain
        def apply_item(item) -> bool:
            meta = item.data(0, Qt.UserRole) or {}
            kind = meta.get('kind')
            # rekurzivně na děti
            any_child = False
            for i in range(item.childCount()):
                if apply_item(item.child(i)):
                    any_child = True
            # self match
            self_match = False
            if not pat:
                self_match = True
            elif kind == 'group' or kind == 'subgroup':
                name = item.text(0).lower()
                self_match = pat in name
            elif kind == 'question':
                self_match = question_matches(meta.get('id'))
            show = self_match or any_child
            item.setHidden(not show)
            return show
        for i in range(self.tree.topLevelItemCount()):
            apply_item(self.tree.topLevelItem(i))

# -------------------- Tree helpery --------------------

    def _refresh_tree(self) -> None:
        self.tree.clear()
        for g in self.root.groups:
            g_item = QTreeWidgetItem([g.name, "Skupina"])
            g_item.setData(0, Qt.UserRole, {"kind": "group", "id": g.id})
            self.tree.addTopLevelItem(g_item)
            self._add_subgroups_to_item(g_item, g.id, g.subgroups)
        self.tree.expandAll()
        self.tree.resizeColumnToContents(0)

    def _add_subgroups_to_item(self, parent_item: QTreeWidgetItem, group_id: str, subgroups: List[Subgroup]) -> None:
        for sg in subgroups:
            sg_item = QTreeWidgetItem([sg.name, "Podskupina"])
            parent_meta = parent_item.data(0, Qt.UserRole) or {}
            parent_sub_id = parent_meta.get("id") if parent_meta.get("kind") == "subgroup" else None
            sg_item.setData(0, Qt.UserRole, {"kind": "subgroup", "id": sg.id, "parent_group_id": group_id, "parent_subgroup_id": parent_sub_id})
            parent_item.addChild(sg_item)
            # questions
            for q in sg.questions:
                label = "Klasická" if q.type == "classic" else "BONUS"
                pts = q.points if q.type == "classic" else f"+{q.bonus_correct}/ {q.bonus_wrong}"
                q_item = QTreeWidgetItem(["Otázka", f"{label} | {pts}"])
                q_item.setData(0, Qt.UserRole, {"kind": "question", "id": q.id, "parent_group_id": group_id, "parent_subgroup_id": sg.id})
                sg_item.addChild(q_item)
            # nested subgroups
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
        """Zreplikuje datový model podle aktuální struktury stromu (po DnD)."""
        # Připravíme mapy id -> objekt
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
            # container je buď Group nebo Subgroup (append do jeho lists)
            child_count = item.childCount()
            for i in range(child_count):
                ch = item.child(i)
                meta = ch.data(0, Qt.UserRole) or {}
                kind = meta.get("kind")
                if kind == "subgroup":
                    old = subgroup_map.get(meta["id"])
                    if not old:
                        continue
                    new_sg = Subgroup(id=old.id, name=ch.text(0), subgroups=[], questions=[])
                    # přidej
                    container.subgroups.append(new_sg)
                    # dive
                    build_from_item(ch, new_sg)
                elif kind == "question":
                    q = question_map.get(meta["id"])
                    if not q:
                        continue
                    if isinstance(container, Group):
                        # otázka bez podskupiny -> vytvoř Default
                        if not container.subgroups:
                            container.subgroups.append(Subgroup(id=str(uuid.uuid4()), name="Default", subgroups=[], questions=[]))
                        container.subgroups[0].questions.append(q)
                    else:
                        container.questions.append(q)
                elif kind == "group":
                    pass  # ignoruj

        # top-level (groups)
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
        g = Group(id=str(uuid.uuid4()), name=name.strip(), subgroups=[])
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

        # Urči nadřazený kontejner
        if kind == "group":
            g = self._find_group(meta["id"])
            if not g:
                return
            g.subgroups.append(Subgroup(id=str(uuid.uuid4()), name=name.strip(), subgroups=[], questions=[]))
        else:
            parent_sg = self._find_subgroup(meta["parent_group_id"], meta["id"])
            if not parent_sg:
                return
            parent_sg.subgroups.append(Subgroup(id=str(uuid.uuid4()), name=name.strip(), subgroups=[], questions=[]))

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
                sg = Subgroup(id=str(uuid.uuid4()), name="Default", subgroups=[], questions=[])
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

    # -------------------- Výběr v tree → načtení detailu --------------------

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
                g = self._find_group(meta["id"])
                name = g.name if g else ""
            else:
                sg = self._find_subgroup(meta["parent_group_id"], meta["id"])
                name = sg.name if sg else ""
            self.rename_line.setText(name)
            self._set_editor_enabled(False)
            self.rename_panel.show()
        else:
            self._clear_editor()
            self.rename_panel.hide()

    def _clear_editor(self) -> None:
        self._current_question_id = None
        self.text_edit.clear()
        self.spin_points.setValue(1)
        self.spin_bonus_correct.setValue(1)
        self.spin_bonus_wrong.setValue(0)
        self.combo_type.setCurrentIndex(0)
        self._set_editor_enabled(False)

    def _load_question_to_editor(self, q: Question) -> None:
        self._current_question_id = q.id
        self.combo_type.setCurrentIndex(0 if q.type == "classic" else 1)
        self.spin_points.setValue(int(q.points))
        self.spin_bonus_correct.setValue(int(q.bonus_correct))
        self.spin_bonus_wrong.setValue(int(q.bonus_wrong))
        self.text_edit.setHtml(q.text_html or "<p><br></p>")
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
                        if q.type == "classic":
                            q.points = int(self.spin_points.value())
                            q.bonus_correct = 0
                            q.bonus_wrong = 0
                        else:
                            q.points = 0
                            q.bonus_correct = int(self.spin_bonus_correct.value())
                            q.bonus_wrong = int(self.spin_bonus_wrong.value())
                        sg.questions[i] = q
                        label = "Klasická" if q.type == "classic" else "BONUS"
                        pts = q.points if q.type == "classic" else f"+{q.bonus_correct}/ {q.bonus_wrong}"
                        self._update_selected_question_item_subtitle(f"{label} | {pts}")
                        if not silent:
                            self.statusBar().showMessage("Změny otázky byly uloženy (lokálně).", 3000)
                        return True
                if apply_in(sg.subgroups):
                    return True
            return False
        for g in self.root.groups:
            if apply_in(g.subgroups):
                break

    def _on_save_question_clicked(self) -> None:
        self._apply_editor_to_current_question(silent=False)

    def _update_selected_question_item_subtitle(self, text: str) -> None:
        items = self.tree.selectedItems()
        if items:
            items[0].setText(1, text)

    # -------------------- Vyhledávače v datech (rekurzivně) --------------------

    def _find_group(self, gid: str) -> Optional[Group]:
        for g in self.root.groups:
            if g.id == gid:
                return g
        return None

    def _find_subgroup(self, gid: str, sgid: str) -> Optional[Subgroup]:
        g = self._find_group(gid)
        if not g:
            return None
        def rec(lst: List[Subgroup]) -> Optional[Subgroup]:
            for sg in lst:
                if sg.id == sgid:
                    return sg
                found = rec(sg.subgroups)
                if found:
                    return found
            return None
        return rec(g.subgroups)

    def _find_question(self, gid: str, sgid: str, qid: str) -> Optional[Question]:
        sg = self._find_subgroup(gid, sgid)
        if not sg:
            return None
        for q in sg.questions:
            if q.id == qid:
                return q
        return None

    def _select_question(self, qid: str) -> None:
        def _walk(item: QTreeWidgetItem) -> Optional[QTreeWidgetItem]:
            meta = item.data(0, Qt.UserRole)
            if meta and meta.get("kind") == "question" and meta.get("id") == qid:
                return item
            for i in range(item.childCount()):
                found = _walk(item.child(i))
                if found:
                    return found
            return None

        for i in range(self.tree.topLevelItemCount()):
            found = _walk(self.tree.topLevelItem(i))
            if found:
                self.tree.setCurrentItem(found)
                self.tree.scrollToItem(found)
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
            fmt.setFontWeight(Qt.Bold if not self.action_bold.isChecked() else Qt.Normal)
        elif which == "italic":
            fmt.setFontItalic(self.action_italic.isChecked())
        elif which == "underline":
            fmt.setFontUnderline(self.action_underline.isChecked())
        self._merge_format_on_selection(fmt)

    def _choose_color(self) -> None:
        color = QColorDialog.getColor(parent=self, title="Vyberte barvu textu")
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setForeground(color)
            self._merge_format_on_selection(fmt)

    def _toggle_bullets(self) -> None:
        cursor = self.text_edit.textCursor()
        block = cursor.block()
        in_list = block.textList() is not None
        if in_list:
            lst = block.textList()
            fmt = lst.format()
            fmt.setStyle(QTextListFormat.ListStyleUndefined)
            cursor.createList(fmt)
        else:
            fmt = QTextListFormat()
            fmt.setStyle(QTextListFormat.ListDisc)
            cursor.createList(fmt)

    def _sync_toolbar_to_cursor(self) -> None:
        fmt = self.text_edit.currentCharFormat()
        self.action_bold.setChecked(fmt.fontWeight() == Qt.Bold)
        self.action_italic.setChecked(fmt.fontItalic())
        self.action_underline.setChecked(fmt.fontUnderline())
        in_list = self.text_edit.textCursor().block().textList() is not None
        self.action_bullets.setChecked(in_list)

    def _on_type_changed_ui(self) -> None:
        is_classic = self.combo_type.currentIndex() == 0
        self.spin_points.setEnabled(is_classic)
        self.spin_bonus_correct.setEnabled(not is_classic)
        self.spin_bonus_wrong.setEnabled(not is_classic)

    # -------------------- Výběr datového souboru --------------------

    def _choose_data_file(self) -> None:
        new_path, _ = QFileDialog.getSaveFileName(self, "Zvolit/uložit JSON s otázkami", str(self.data_path), "JSON (*.json)")
        if new_path:
            self.data_path = Path(new_path)
            self.statusBar().showMessage(f"Datový soubor změněn na: {self.data_path}", 4000)
            self.load_data()
            self._refresh_tree()

    # -------------------- Import z DOCX --------------------

    def _import_from_docx(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "Vyberte .docx soubory", str(self.project_root), "Word dokumenty (*.docx)")
        if not paths:
            return
        imported = 0
        group_id, subgroup_id = self._ensure_unassigned_group()
        for p in paths:
            try:
                paras = self._extract_paragraphs_from_docx(Path(p))
                questions = self._parse_questions_from_paragraphs(paras)
                if questions:
                    sg = self._find_subgroup(group_id, subgroup_id)
                    if not sg:
                        continue
                    for q in questions:
                        sg.questions.append(q)
                    imported += len(questions)
            except Exception as e:
                QMessageBox.warning(self, "Import selhal", f"{p}\n{e}")
        if imported:
            self._refresh_tree()
            self.save_data()
            self.statusBar().showMessage(f"Importováno {imported} otázek do skupiny 'Neroztříděné'.", 5000)
        else:
            QMessageBox.information(self, "Import", "Nebyla nalezena žádná otázka.")

    

    def _extract_paragraphs_from_docx(self, path: Path) -> list[tuple[str, bool]]:
        """Vrátí seznam (text, is_numbered) pro jednotlivé odstavce v DOCX.
        is_numbered je True, pokud odstavec obsahuje w:numPr (číslovaný seznam ve Wordu).
        """
        with zipfile.ZipFile(path, 'r') as z:
            with z.open('word/document.xml') as f:
                xml_bytes = f.read()
        root = ET.fromstring(xml_bytes)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        out: list[tuple[str, bool]] = []
        for p in root.findall('.//w:p', ns):
            ppr = p.find('w:pPr', ns)
            is_num = False
            if ppr is not None and ppr.find('w:numPr', ns) is not None:
                is_num = True
            texts = [t.text or '' for t in p.findall('.//w:t', ns)]
            txt = ''.join(texts).strip()
            out.append((txt, is_num))
        return out


    def _ensure_unassigned_group(self) -> tuple[str, str]:
        name = "Neroztříděné"
        g = next((g for g in self.root.groups if g.name == name), None)
        if not g:
            g = Group(id=str(uuid.uuid4()), name=name, subgroups=[])
            self.root.groups.append(g)
        if not g.subgroups:
            g.subgroups.append(Subgroup(id=str(uuid.uuid4()), name="Default", subgroups=[], questions=[]))
        return g.id, g.subgroups[0].id

    def _parse_questions_from_paragraphs(self, paragraphs: list[tuple[str, bool]]) -> list[Question]:
        lines = text.splitlines()
        blocks = []
        buf = []
        def flush():
            s = "\n".join(buf).strip()
            if s:
                blocks.append(s)
            buf.clear()
        for ln in lines:
            if re.match(r"\s*Otázka\s+\d+", ln, flags=re.IGNORECASE):
                flush()
                buf.append(ln)
            elif not ln.strip():
                flush()
            else:
                buf.append(ln)
        flush()

        def is_noise(block: str) -> bool:
            hay = block.lower()
            noise_keys = [
                "datum:", "jméno", "podpis", "klasifikační", "stupnice", "maximum bodů",
                "na uvedené otázky", "souhlasíte", "cookies"
            ]
            return any(k in hay for k in noise_keys)

        out: list[Question] = []
        for b in blocks:
            if is_noise(b):
                continue
            b_stripped = b.strip()
            is_bonus = bool(re.search(r"\bOtázka\s+\d+", b_stripped, flags=re.IGNORECASE)) or ("bonus" in b_stripped.lower())
            is_classic = bool(re.match(r"\s*\d+[\.)]\s", b_stripped)) or (" bod" in b_stripped.lower())
            qtype = "bonus" if is_bonus else "classic"
            if qtype == "classic" and not is_classic:
                if not ("?" in b_stripped or re.search(r"\b(Popište|Uveďte|Zašifrujte|Vysvětlete|Porovnejte|Jaký|Jak|Stručně)\b", b_stripped, re.IGNORECASE)):
                    continue
            html_text = "<p>" + html.escape(b_stripped).replace("\n", "<br>") + "</p>"
            if qtype == "bonus":
                q = Question.new_default("bonus")
                q.text_html = html_text
                q.bonus_correct = 1
                q.bonus_wrong = 0
            else:
                q = Question.new_default("classic")
                q.text_html = html_text
                q.points = 1
            out.append(q)
        return out

    # -------------------- Přesun otázky (menu – zachováno) --------------------

    def _move_question(self) -> None:
        kind, meta = self._selected_node()
        if kind != "question":
            QMessageBox.information(self, "Přesun", "Vyberte nejprve otázku ve stromu.")
            return
        from PySide6.QtWidgets import QInputDialog
        group_names = [g.name for g in self.root.groups]
        if not group_names:
            QMessageBox.information(self, "Přesun", "Neexistují žádné skupiny.")
            return
        g_name, ok = QInputDialog.getItem(self, "Přesun otázky", "Cílová skupina:", group_names, 0, False)
        if not ok or not g_name:
            return
        g = next((g for g in self.root.groups if g.name == g_name), None)
        if not g:
            return
        # výběr cílové podskupiny (z celé hierarchie)
        def flatten_subgroups(lst: List[Subgroup], acc: List[Subgroup]):
            for s in lst:
                acc.append(s)
                flatten_subgroups(s.subgroups, acc)
        all_sg: List[Subgroup] = []
        flatten_subgroups(g.subgroups, all_sg)
        if not all_sg:
            g.subgroups.append(Subgroup(id=str(uuid.uuid4()), name="Default", subgroups=[], questions=[]))
            all_sg = [g.subgroups[0]]
        sg_names = [sg.name for sg in all_sg]
        sg_name, ok = QInputDialog.getItem(self, "Přesun otázky", "Cílová podskupina:", sg_names, 0, False)
        if not ok or not sg_name:
            return
        target_sg = next((s for s in all_sg if s.name == sg_name), None)
        if not target_sg:
            return

    def _bulk_move_selected(self) -> None:
        items = [it for it in self.tree.selectedItems() if (it.data(0, Qt.UserRole) or {}).get('kind') == 'question']
        if not items:
            QMessageBox.information(self, 'Přesun', 'Vyberte ve stromu alespoň jednu otázku.')
            return
        # výběr cíle (skupina -> podskupina)
        from PySide6.QtWidgets import QInputDialog
        group_names = [g.name for g in self.root.groups]
        g_name, ok = QInputDialog.getItem(self, 'Přesun vybraných', 'Cílová skupina:', group_names, 0, False)
        if not ok or not g_name:
            return
        g = next((g for g in self.root.groups if g.name == g_name), None)
        if not g:
            return
        def flatten(lst, acc):
            for s in lst:
                acc.append(s)
                flatten(s.subgroups, acc)
        all_sg = []
        flatten(g.subgroups, all_sg)
        if not all_sg:
            g.subgroups.append(Subgroup(id=str(uuid.uuid4()), name='Default', subgroups=[], questions=[]))
            all_sg = [g.subgroups[0]]
        sg_names = [sg.name for sg in all_sg]
        sg_name, ok = QInputDialog.getItem(self, 'Přesun vybraných', 'Cílová podskupina:', sg_names, 0, False)
        if not ok or not sg_name:
            return
        target_sg = next((s for s in all_sg if s.name == sg_name), None)
        if not target_sg:
            return
        # přesun
        moved = 0
        for it in items:
            meta = it.data(0, Qt.UserRole)
            qid = meta.get('id')
            sg = self._find_subgroup(meta.get('parent_group_id'), meta.get('parent_subgroup_id'))
            q = self._find_question(meta.get('parent_group_id'), meta.get('parent_subgroup_id'), qid)
            if sg and q:
                sg.questions = [qq for qq in sg.questions if qq.id != qid]
                target_sg.questions.append(q)
                moved += 1
        self._refresh_tree()
        self.save_data()
        self.statusBar().showMessage(f'Přesunuto {moved} otázek do {g_name} / {sg_name}.', 4000)

    def _bulk_delete_selected(self) -> None:
        items = [it for it in self.tree.selectedItems() if (it.data(0, Qt.UserRole) or {}).get('kind') == 'question']
        if not items:
            QMessageBox.information(self, 'Mazání', 'Vyberte ve stromu alespoň jednu otázku.')
            return
        if QMessageBox.question(self, 'Smazat vybrané', f'Opravdu smazat {len(items)} otázek?') != QMessageBox.Yes:
            return
        deleted = 0
        for it in items:
            meta = it.data(0, Qt.UserRole) or {}
            qid = meta.get('id')
            sg = self._find_subgroup(meta.get('parent_group_id'), meta.get('parent_subgroup_id'))
            if sg:
                before = len(sg.questions)
                sg.questions = [q for q in sg.questions if q.id != qid]
                if len(sg.questions) < before:
                    deleted += 1
        self._refresh_tree()
        self.save_data()
        self.statusBar().showMessage(f'Smazáno {deleted} otázek.', 4000)
        # Najdi zdroj
        src_gid = meta["parent_group_id"]
        src_sgid = meta["parent_subgroup_id"]
        qid = meta["id"]
        src_sg = self._find_subgroup(src_gid, src_sgid)
        q = self._find_question(src_gid, src_sgid, qid)
        if not (src_sg and q):
            return
        # Odeber ze zdroje + přidej do cíle
        src_sg.questions = [qq for qq in src_sg.questions if qq.id != qid]
        target_sg.questions.append(q)
        self._refresh_tree()
        self.save_data()
        self.statusBar().showMessage(f"Otázka přesunuta do {g_name} / {sg_name}.", 4000)


# --------------------------- main ---------------------------

def main() -> int:
    app = QApplication(sys.argv)
    apply_dark_theme(app)
    w = MainWindow()
    w.show()
    ret = app.exec()
    return ret


if __name__ == "__main__":
    sys.exit(main())
