#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PubMed Search — interface graphique pour pubmed_search.py

L'interface ne contient pas la logique de recherche : elle charge le script
pubmed_search.py choisi par l'utilisateur (fichier externe, modifiable à part)
et l'exécute avec le Python intégré à l'application.

Mode interne : `pubmed_gui --run-script <script.py> [arguments…]`
exécute le script avec l'interpréteur embarqué (utilisé par l'interface).
"""

import json
import os
import re
import sys
import threading

APP_NAME = "PubMed Search"
APP_ID = "pubmed-search-gui"
APP_VERSION = "1.2.0"


# --------------------------------------------------------------------------- #
# Mode « exécuteur » : lance pubmed_search.py avec le Python embarqué
# --------------------------------------------------------------------------- #
def _run_script_mode(script, args):
    import runpy
    import _thread
    import warnings
    warnings.filterwarnings("ignore", message=".*ARC4.*")      # bruit de pypdf/cryptography
    warnings.filterwarnings("ignore", category=DeprecationWarning)
    import runtime_deps  # noqa: F401  (embarque requests, openpyxl, pypdf et la stdlib utile)

    if os.name == "nt":
        # Exécutable Windows « fenêtré » : Python ne crée pas sys.stdout/stdin même quand
        # l'interface lui passe des tubes ; on les rattache aux poignées héritées.
        try:
            import ctypes
            import msvcrt
            k32 = ctypes.windll.kernel32
            for name, num, mode in (("stdin", -10, "r"), ("stdout", -11, "w"), ("stderr", -12, "w")):
                if getattr(sys, name) is None:
                    h = k32.GetStdHandle(num)
                    if h and h != -1:
                        fd = msvcrt.open_osfhandle(h, os.O_RDONLY if mode == "r" else os.O_WRONLY)
                        setattr(sys, name, open(fd, mode, encoding="utf-8", errors="replace",
                                                buffering=1, closefd=False))
        except Exception:  # noqa
            pass

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
        except Exception:  # noqa
            pass

    def watch_stdin():
        # L'interface écrit « STOP » sur l'entrée standard : on lève KeyboardInterrupt
        # dans le script, qui écrit alors ses résultats partiels (Excel, error.txt…).
        try:
            for line in sys.stdin:
                if line.strip() == "STOP":
                    _thread.interrupt_main()
                    return
        except Exception:  # noqa
            return

    if sys.stdin is not None:
        threading.Thread(target=watch_stdin, daemon=True).start()

    script = os.path.abspath(script)
    sys.argv = [script] + list(args)
    sys.path.insert(0, os.path.dirname(script))
    try:
        runpy.run_path(script, run_name="__main__")
    except SystemExit as e:
        code = e.code
        if isinstance(code, str):
            print(code, file=sys.stderr)
            code = 1
        sys.exit(code or 0)
    except KeyboardInterrupt:
        sys.exit(130)
    sys.exit(0)


if len(sys.argv) >= 3 and sys.argv[1] == "--run-script":
    _run_script_mode(sys.argv[2], sys.argv[3:])


# --------------------------------------------------------------------------- #
# Interface Qt
# --------------------------------------------------------------------------- #
from PySide6.QtCore import (Qt, QDateTime, QProcess, QProcessEnvironment,  # noqa: E402
                            QStandardPaths, QTimer, QUrl, Signal, QCoreApplication,
                            QLockFile, QFileSystemWatcher)
from PySide6.QtGui import (QColor, QDesktopServices, QFont, QFontDatabase,  # noqa: E402
                           QIcon, QPainter, QPainterPath, QSyntaxHighlighter,
                           QTextCharFormat, QTextCursor)
from PySide6.QtWidgets import (QApplication, QButtonGroup, QCheckBox, QDateTimeEdit,  # noqa: E402
                               QFileDialog, QFrame, QGridLayout, QHBoxLayout,
                               QInputDialog, QLabel, QLineEdit, QMainWindow,
                               QMessageBox, QPlainTextEdit, QPushButton,
                               QRadioButton, QScrollArea, QSizePolicy, QSpinBox,
                               QSplitter, QTabWidget, QToolButton, QVBoxLayout,
                               QWidget, QTableWidget, QTableWidgetItem, QHeaderView,
                               QAbstractItemView)

FROZEN = getattr(sys, "frozen", False)
BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
RES = os.path.join(BASE_DIR, "resources")

# Palette « coloration H&E » : hématoxyline (violet-bleu) et éosine (rose)
C = {
    "slide": "#F5F4F8",      # fond : lame de verre
    "panel": "#FFFFFF",
    "ink": "#1C1D33",
    "muted": "#62637A",
    "rule": "#DCDBE6",
    "hema": "#3A3D8F",       # hématoxyline : actions, sélection
    "hema_soft": "#E6E6F3",
    "eosin": "#C8467A",      # éosine : PDF obtenus
    "eosin_soft": "#F7E3EB",
    "log_bg": "#17182B",
    "log_fg": "#D9D8EA",
    "log_dim": "#8C8DA8",
    "ok": "#2E7D5B",
    "warn": "#B26A00",
    "err": "#B3261E",
}


def config_dir():
    d = QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation)
    if not d:
        d = os.path.join(os.path.expanduser("~"), ".config", APP_ID)
    os.makedirs(d, exist_ok=True)
    return d


def bundled_script():
    p = os.path.join(RES, "pubmed_search.py")
    return p if os.path.exists(p) else ""


def script_version(path):
    try:
        with open(path, encoding="utf-8") as f:
            m = re.search(r'^VERSION\s*=\s*["\']([^"\']+)["\']', f.read(), re.M)
        return m.group(1) if m else "?"
    except OSError:
        return ""


EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}$")


# --------------------------------------------------------------------------- #
# Coloration syntaxique de la requête PubMed
# --------------------------------------------------------------------------- #
class QueryHighlighter(QSyntaxHighlighter):
    def __init__(self, doc):
        super().__init__(doc)
        self.f_tag = QTextCharFormat()
        self.f_tag.setForeground(QColor(C["eosin"]))
        self.f_op = QTextCharFormat()
        self.f_op.setForeground(QColor(C["hema"]))
        self.f_op.setFontWeight(QFont.Weight.DemiBold)
        self.f_quote = QTextCharFormat()
        self.f_quote.setForeground(QColor(C["ok"]))
        self.f_paren = QTextCharFormat()
        self.f_paren.setForeground(QColor(C["muted"]))

    def highlightBlock(self, text):
        for m in re.finditer(r"[()]", text):
            self.setFormat(m.start(), 1, self.f_paren)
        for m in re.finditer(r'"[^"]*"', text):
            self.setFormat(m.start(), m.end() - m.start(), self.f_quote)
        for m in re.finditer(r"\b(AND|OR|NOT)\b", text):
            self.setFormat(m.start(), m.end() - m.start(), self.f_op)
        for m in re.finditer(r"\[[^\]]+\]", text):
            self.setFormat(m.start(), m.end() - m.start(), self.f_tag)


# --------------------------------------------------------------------------- #
# Petits composants
# --------------------------------------------------------------------------- #
class ProgressStrip(QWidget):
    """Barre segmentée : PDF obtenus (éosine) / traités sans PDF (hématoxyline) / restants."""

    def __init__(self):
        super().__init__()
        self.total = self.ok = self.miss = 0
        self.setFixedHeight(10)

    def set_values(self, total, ok, miss):
        self.total, self.ok, self.miss = total, ok, miss
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        r = self.rect()
        path = QPainterPath()
        path.addRoundedRect(r, 5, 5)
        p.setClipPath(path)
        p.fillRect(r, QColor("#2A2B45"))
        if self.total:
            w_ok = r.width() * self.ok / self.total
            w_miss = r.width() * self.miss / self.total
            p.fillRect(0, 0, int(w_ok), r.height(), QColor(C["eosin"]))
            p.fillRect(int(w_ok), 0, int(w_miss) + 1, r.height(), QColor("#6E71C4"))
        p.end()


class Stat(QWidget):
    def __init__(self, label, color):
        super().__init__()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self.value = QLabel("–")
        self.value.setObjectName("statValue")
        self.value.setStyleSheet(f"color:{color};")
        cap = QLabel(label)
        cap.setObjectName("statLabel")
        lay.addWidget(self.value)
        lay.addWidget(cap)

    def set(self, v):
        self.value.setText(str(v))


def field_label(text, buddy=None):
    lab = QLabel(text)
    lab.setObjectName("fieldLabel")
    if buddy is not None:
        lab.setBuddy(buddy)
    return lab


def hint(text):
    lab = QLabel(text)
    lab.setObjectName("hint")
    lab.setWordWrap(True)
    return lab


def error_label():
    lab = QLabel("")
    lab.setObjectName("fieldError")
    lab.setWordWrap(True)
    lab.hide()
    return lab


def section_title(text):
    lab = QLabel(text)
    lab.setObjectName("sectionTitle")
    return lab


class SecretEdit(QWidget):
    """Champ de clé API masqué, avec bouton Afficher."""

    def __init__(self, placeholder=""):
        super().__init__()
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        self.edit = QLineEdit()
        self.edit.setPlaceholderText(placeholder)
        self.edit.setEchoMode(QLineEdit.Password)
        self.btn = QToolButton()
        self.btn.setText("Afficher")
        self.btn.setCheckable(True)
        self.btn.setObjectName("ghost")
        self.btn.toggled.connect(self._toggle)
        lay.addWidget(self.edit, 1)
        lay.addWidget(self.btn)

    def _toggle(self, on):
        self.edit.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password)
        self.btn.setText("Masquer" if on else "Afficher")

    def text(self):
        return self.edit.text().strip()

    def setText(self, t):
        self.edit.setText(t or "")


class Disclosure(QWidget):
    """Volet dépliable (« menu déroulant ») : en-tête cliquable + contenu."""

    def __init__(self, title):
        super().__init__()
        self.title = title
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self.head = QToolButton()
        self.head.setObjectName("disclosure")
        self.head.setCheckable(True)
        self.head.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.head.setArrowType(Qt.RightArrow)
        self.head.setText(title)
        self.head.toggled.connect(self._toggle)
        self.body = QWidget()
        self.body.setObjectName("disclosureBody")
        self.body_lay = QVBoxLayout(self.body)
        self.body_lay.setContentsMargins(22, 4, 0, 8)
        self.body_lay.setSpacing(6)
        self.body.hide()
        lay.addWidget(self.head, 0, Qt.AlignLeft)
        lay.addWidget(self.body)

    def _toggle(self, on):
        self.body.setVisible(on)
        self.head.setArrowType(Qt.DownArrow if on else Qt.RightArrow)

    def set_count(self, n):
        self.head.setText(self.title + (f"  ({n} coché{'s' if n > 1 else ''})" if n else ""))


def user_filters_path():
    return os.path.join(config_dir(), "filtres.json")


def load_filters():
    """Filtres de l'onglet Recherche : copie de l'utilisateur si elle existe, sinon ceux fournis."""
    for p in (user_filters_path(), os.path.join(RES, "filtres.json")):
        try:
            with open(p, encoding="utf-8") as f:
                data = json.load(f)
            groups = [g for g in data.get("groupes", []) if g.get("titre") and g.get("elements")]
            if groups:
                return groups, p
        except (OSError, ValueError):
            continue
    return [], ""


def compose_query(base, groups):
    """(requête) AND (cases du groupe 1 en OR) AND (groupe 2) …"""
    parts = []
    base = " ".join((base or "").split())
    if base:
        parts.append(base)
    for queries in groups:
        if queries:
            parts.append(" OR ".join(f"({q})" if " OR " in q else q for q in queries))
    if len(parts) == 1:
        return parts[0]
    return " AND ".join(f"({p})" for p in parts)


def scroll_wrap(widget):
    sa = QScrollArea()
    sa.setWidgetResizable(True)
    sa.setFrameShape(QFrame.NoFrame)
    sa.setWidget(widget)
    return sa


# --------------------------------------------------------------------------- #
# Fenêtre principale
# --------------------------------------------------------------------------- #
class MainWindow(QMainWindow):
    log_signal = Signal(str, str)
    TAB_SEARCH, TAB_QUEUE, TAB_IMPORT, TAB_SETTINGS = range(4)

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(QIcon(os.path.join(RES, "icon.png")))
        self.resize(1240, 800)
        self.setMinimumSize(980, 640)
        self.settings_path = os.path.join(config_dir(), "settings.json")
        self.proc = None
        self.run_kind = None
        self.results_dir = ""
        self.excel_path = ""
        self.script_flags = set()
        self.link_cursor = 0
        self._buf = ""
        self._counts = {"total": 0, "ok": 0, "miss": 0}

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._build_header())

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)
        split.setHandleWidth(1)
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.queue = []
        self.queue_active = False
        self.queue_stop = False
        self.current_job = None
        self.queue_timer = QTimer(self)
        self.queue_timer.setInterval(15000)
        self.queue_timer.timeout.connect(self._queue_tick)
        self.tabs.addTab(scroll_wrap(self._build_search_tab()), "Recherche")
        self.tabs.addTab(scroll_wrap(self._build_queue_tab()), "File d'attente")
        self.tabs.addTab(scroll_wrap(self._build_import_tab()), "Import manuel")
        self.tabs.addTab(scroll_wrap(self._build_settings_tab()), "Réglages")
        split.addWidget(self.tabs)
        split.addWidget(self._build_log_panel())
        split.setStretchFactor(0, 5)
        split.setStretchFactor(1, 6)
        split.setSizes([540, 700])
        root.addWidget(split, 1)

        self._load_settings()
        self._start_inbox()
        self._refresh_script()
        self._update_preview()
        self._sync_sort_state()

    # ----------------------------------------------------------------- header
    def _build_header(self):
        bar = QFrame()
        bar.setObjectName("header")
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(24, 14, 24, 14)
        mark = QLabel()
        mark.setPixmap(QIcon(os.path.join(RES, "icon.png")).pixmap(30, 30))
        title = QLabel(APP_NAME)
        title.setObjectName("appTitle")
        self.script_chip = QLabel("Aucun script chargé")
        self.script_chip.setObjectName("scriptChip")
        change = QPushButton("Changer de script")
        change.setObjectName("ghost")
        change.clicked.connect(self._choose_script)
        lay.addWidget(mark)
        lay.addSpacing(8)
        lay.addWidget(title)
        lay.addStretch(1)
        lay.addWidget(self.script_chip)
        lay.addWidget(change)
        return bar

    # ----------------------------------------------------------- recherche
    def _build_search_tab(self):
        w = QWidget()
        w.setObjectName("formPage")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(28, 24, 28, 28)
        lay.setSpacing(6)

        self.query = QPlainTextEdit()
        self.query.setObjectName("query")
        self.query.setPlaceholderText('(NETosis[ti] OR "extracellular trap*"[ti]) AND review[pt]')
        self.query.setTabChangesFocus(True)
        self.query.setFixedHeight(104)
        QueryHighlighter(self.query.document())
        self.query.textChanged.connect(lambda: self._clear_error(self.err_query))
        lay.addWidget(field_label("Requête PubMed", self.query))
        lay.addWidget(self.query)
        self.err_query = error_label()
        lay.addWidget(self.err_query)
        lay.addWidget(hint("Même syntaxe que sur pubmed.ncbi.nlm.nih.gov : champs entre crochets "
                           "([ti], [tiab], [pt]…) et opérateurs AND, OR, NOT en majuscules."))
        lay.addSpacing(14)

        # Filtres (volets dépliables) : ajoutés à la requête entre parenthèses
        self.filter_groups, self.filters_file = load_filters()
        self.filter_boxes = []      # (index du groupe, label, requête, case)
        self.filter_sections = []
        for gi, g in enumerate(self.filter_groups):
            sec = Disclosure(g["titre"])
            grid = QGridLayout()
            grid.setHorizontalSpacing(18)
            grid.setVerticalSpacing(4)
            for i, el in enumerate(g["elements"]):
                cb = QCheckBox(el.get("label", "?"))
                cb.setToolTip(el.get("query", ""))
                cb.toggled.connect(self._filters_changed)
                grid.addWidget(cb, i // 2, i % 2)
                self.filter_boxes.append((gi, el.get("label", ""), el.get("query", ""), cb))
            sec.body_lay.addLayout(grid)
            self.filter_sections.append(sec)
            lay.addWidget(sec)
        if self.filter_groups:
            lay.addWidget(hint("Cases d'un même volet combinées par OR ; volets et requête combinés par AND. "
                               "Survolez une case pour voir les termes ajoutés."))
        self.final_label = field_label("Requête envoyée à PubMed")
        self.final_query = QPlainTextEdit()
        self.final_query.setObjectName("finalQuery")
        self.final_query.setReadOnly(True)
        self.final_query.setFixedHeight(92)
        QueryHighlighter(self.final_query.document())
        self.final_info = hint("")
        for wdg in (self.final_label, self.final_query, self.final_info):
            wdg.hide()
            lay.addWidget(wdg)
        self.query.textChanged.connect(self._filters_changed)
        lay.addSpacing(18)

        # Tri
        lay.addWidget(field_label("Tri des résultats"))
        seg = QHBoxLayout()
        seg.setSpacing(0)
        self.sort_group = QButtonGroup(self)
        for i, (key, text) in enumerate((("best", "Pertinence"), ("recent", "Plus récents"),
                                         ("both", "Les deux"))):
            b = QPushButton(text)
            b.setCheckable(True)
            b.setObjectName("segment")
            b.setProperty("pos", "first" if i == 0 else ("last" if i == 2 else "mid"))
            b.setProperty("key", key)
            self.sort_group.addButton(b, i)
            seg.addWidget(b)
        self.sort_group.button(0).setChecked(True)
        self.sort_group.idToggled.connect(lambda *_: self._sync_sort_state())
        seg.addStretch(1)
        lay.addLayout(seg)
        self.sort_hint = hint("")
        lay.addWidget(self.sort_hint)
        lay.addSpacing(14)

        grid = QGridLayout()
        grid.setHorizontalSpacing(24)
        grid.setVerticalSpacing(6)
        self.max_results = QSpinBox()
        self.max_results.setRange(1, 10000)
        self.max_results.setValue(100)
        self.max_results.setSingleStep(50)
        self.passes = QSpinBox()
        self.passes.setRange(0, 10)
        self.passes.setValue(0)
        grid.addWidget(field_label("Articles maximum par tri", self.max_results), 0, 0)
        grid.addWidget(field_label("Passes supplémentaires", self.passes), 0, 1)
        grid.addWidget(self.max_results, 1, 0)
        grid.addWidget(self.passes, 1, 1)
        grid.addWidget(hint("Limite NCBI : 10 000."), 2, 0)
        grid.addWidget(hint("Relance sur les gratuits non téléchargés. 0 = aucune."), 2, 1)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        lay.addLayout(grid)
        lay.addSpacing(10)
        self.getpdf = QCheckBox("Télécharger les PDF des articles gratuits")
        self.getpdf.setChecked(True)
        self.getpdf.toggled.connect(self.passes.setEnabled)
        lay.addWidget(self.getpdf)
        lay.addSpacing(18)

        # Sortie
        lay.addWidget(section_title("Où enregistrer"))
        self.outdir = QLineEdit()
        self.outdir.setPlaceholderText("Dossier qui recevra le dossier de résultats")
        self.outdir.textChanged.connect(self._update_preview)
        self.outdir.textChanged.connect(lambda: self._clear_error(self.err_outdir))
        b_choose = QPushButton("Choisir…")
        b_choose.setObjectName("secondary")
        b_choose.clicked.connect(self._choose_outdir)
        b_new = QPushButton("Nouveau dossier…")
        b_new.setObjectName("secondary")
        b_new.clicked.connect(self._new_outdir)
        row = QHBoxLayout()
        row.setSpacing(6)
        row.addWidget(self.outdir, 1)
        row.addWidget(b_choose)
        row.addWidget(b_new)
        lay.addWidget(field_label("Dossier de destination", self.outdir))
        lay.addLayout(row)
        self.err_outdir = error_label()
        lay.addWidget(self.err_outdir)
        lay.addSpacing(10)

        self.excel = QLineEdit()
        self.excel.setPlaceholderText("Diabetes_review")
        self.excel.textChanged.connect(self._update_preview)
        self.excel.textChanged.connect(lambda: self._clear_error(self.err_excel))
        lay.addWidget(field_label("Nom du fichier Excel", self.excel))
        lay.addWidget(self.excel)
        self.err_excel = error_label()
        lay.addWidget(self.err_excel)
        self.preview = QLabel("")
        self.preview.setObjectName("preview")
        self.preview.setWordWrap(True)
        self.preview.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(self.preview)
        lay.addSpacing(16)

        # Options avancées
        self.adv_toggle = QToolButton()
        self.adv_toggle.setObjectName("disclosure")
        self.adv_toggle.setText("Options avancées")
        self.adv_toggle.setCheckable(True)
        self.adv_toggle.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.adv_toggle.setArrowType(Qt.RightArrow)
        self.adv = QWidget()
        al = QVBoxLayout(self.adv)
        al.setContentsMargins(0, 6, 0, 0)
        al.setSpacing(6)
        self.schedule = QCheckBox("Programmer le lancement")
        self.schedule_at = QDateTimeEdit(QDateTime.currentDateTime().addSecs(3600))
        self.schedule_at.setDisplayFormat("dd/MM/yyyy  HH:mm")
        self.schedule_at.setCalendarPopup(True)
        self.schedule_at.setEnabled(False)
        self.schedule.toggled.connect(self.schedule_at.setEnabled)
        srow = QHBoxLayout()
        srow.addWidget(self.schedule)
        srow.addWidget(self.schedule_at)
        srow.addStretch(1)
        al.addLayout(srow)
        al.addWidget(hint("L'application et l'ordinateur doivent rester ouverts jusqu'à l'heure prévue."))
        self.adv.hide()
        self.adv_toggle.toggled.connect(self._toggle_adv)
        lay.addWidget(self.adv_toggle, 0, Qt.AlignLeft)
        lay.addWidget(self.adv)
        lay.addSpacing(22)

        brow = QHBoxLayout()
        brow.setSpacing(10)
        self.add_queue_btn = QPushButton("Ajouter à la file")
        self.add_queue_btn.setObjectName("secondary")
        self.add_queue_btn.setMinimumHeight(44)
        self.add_queue_btn.clicked.connect(self._add_to_queue)
        self.run_btn = QPushButton("Lancer la recherche")
        self.run_btn.setObjectName("primary")
        self.run_btn.setMinimumHeight(44)
        self.run_btn.clicked.connect(self._start_search)
        brow.addWidget(self.add_queue_btn, 1)
        brow.addWidget(self.run_btn, 2)
        lay.addLayout(brow)
        self.queue_note = hint("")
        lay.addWidget(self.queue_note)
        lay.addStretch(1)
        return w

    # ----------------------------------------------- liens pubmedsearch://
    def _start_inbox(self):
        self.inbox_watch = QFileSystemWatcher([inbox_dir()], self)
        self.inbox_watch.directoryChanged.connect(lambda *_: QTimer.singleShot(150, self._read_inbox))
        self._read_inbox()

    def _read_inbox(self):
        d = inbox_dir()
        for name in sorted(os.listdir(d)):
            if not name.endswith(".json"):
                continue
            p = os.path.join(d, name)
            try:
                with open(p, encoding="utf-8") as f:
                    url = json.load(f).get("url", "")
            except (OSError, ValueError):
                url = ""
            try:
                os.remove(p)
            except OSError:
                pass
            if url:
                self.handle_url(url)

    def _bring_to_front(self):
        if self.isMinimized():
            self.showNormal()
        self.show()
        self.raise_()
        self.activateWindow()

    def handle_url(self, url):
        """Requête envoyée par l'extension Firefox (ou tout lien pubmedsearch://)."""
        self._bring_to_front()
        if url.lower().startswith(URL_SCHEME + "://show"):
            return
        req = parse_request(url)
        if not req["query"]:
            self.phase.setText("Lien reçu sans requête PubMed : rien à faire.")
            return
        self.tabs.setCurrentIndex(self.TAB_SEARCH)
        self.query.setPlainText(req["query"])
        for _, _, _, cb in self.filter_boxes:
            cb.setChecked(False)
        if req["sort"]:
            self.sort_group.button({"best": 0, "recent": 1, "both": 2}[req["sort"]]).setChecked(True)
        if req["max"]:
            self.max_results.setValue(req["max"])
        if req["action"] in ("queue", "run") and (not self._stem() or self._stem().startswith("PubMed_")):
            self.excel.setText("PubMed_" + QDateTime.currentDateTime().toString("yyyyMMdd_HHmmss"))
        origin = "Firefox" if req["source"] == "firefox" else "un lien"
        if req["action"] == "queue":
            n0 = len(self.queue)
            self._add_to_queue()
            if len(self.queue) > n0:
                self.phase.setText(f"Requête reçue de {origin} et ajoutée à la file d'attente.")
            return
        if req["action"] == "run":
            if self.proc is None:
                self.phase.setText(f"Requête reçue de {origin} : lancement de la recherche.")
                self._start_search()
            else:
                self._add_to_queue()
                self.phase.setText(f"Requête reçue de {origin} : une recherche est en cours, "
                                   "elle a été ajoutée à la file d'attente.")
            return
        self.phase.setText(f"Requête reçue de {origin}. Vérifiez le formulaire puis lancez la recherche.")
        if not self._stem():
            self.excel.setFocus()

    # --------------------------------------------------------------- filtres
    def _selected_filters(self):
        """[[requêtes cochées du groupe 0], [groupe 1], …] et libellés cochés."""
        groups = [[] for _ in self.filter_groups]
        labels = []
        for gi, label, q, cb in self.filter_boxes:
            if cb.isChecked() and q:
                groups[gi].append(q)
                labels.append(label)
        return groups, labels

    def _final_query(self):
        groups, _ = self._selected_filters()
        return compose_query(self.query.toPlainText(), groups)

    def _filters_changed(self):
        groups, labels = self._selected_filters()
        for gi, sec in enumerate(self.filter_sections):
            sec.set_count(len(groups[gi]))
        show = any(groups)
        for wdg in (self.final_label, self.final_query, self.final_info):
            wdg.setVisible(show)
        if show:
            fq = self._final_query()
            self.final_query.setPlainText(fq)
            n_terms = fq.count(" OR ") + fq.count(" AND ") + 1
            self.final_info.setText(f"{len(fq)} caractères, environ {n_terms} termes. "
                                    "Filtres : " + ", ".join(labels) + ".")

    def _toggle_adv(self, on):
        self.adv.setVisible(on)
        self.adv_toggle.setArrowType(Qt.DownArrow if on else Qt.RightArrow)

    def _sort_key(self):
        b = self.sort_group.checkedButton()
        return b.property("key") if b else "best"

    def _sync_sort_state(self):
        texts = {
            "best": "Ordre « Best match » de PubMed.",
            "recent": "Articles triés par date de publication, du plus récent au plus ancien.",
            "both": "Deux recherches fusionnées : le journal indique les articles propres à chaque tri "
                    "et ceux qu'ils ont en commun.",
        }
        self.sort_hint.setText(texts[self._sort_key()])
        for b in self.sort_group.buttons():
            b.style().unpolish(b)
            b.style().polish(b)

    def _stem(self):
        name = os.path.basename(self.excel.text().strip())
        if name.lower().endswith(".xlsx"):
            name = name[:-5]
        return name

    def _update_preview(self):
        stem = self._stem() or "Nom"
        base = self.outdir.text().strip() or "…"
        safe = re.sub(r"[^A-Za-z0-9_\-]", "", stem) or "pubmed"
        when = QDateTime.currentDateTime().toString("dd.MM.yyyy_HH'h'mm")
        self.preview.setText(
            f"Les résultats seront créés dans\n{os.path.join(base, f'{safe}_{when}')}\n"
            f"avec les PDF dans le sous-dossier PDF_{safe}")

    def _choose_outdir(self):
        start = self.outdir.text().strip() or os.path.expanduser("~")
        d = QFileDialog.getExistingDirectory(self, "Dossier de destination", start)
        if d:
            self.outdir.setText(d)

    def _new_outdir(self):
        parent = QFileDialog.getExistingDirectory(
            self, "Où créer le nouveau dossier ?", self.outdir.text().strip() or os.path.expanduser("~"))
        if not parent:
            return
        name, ok = QInputDialog.getText(self, "Nouveau dossier", "Nom du dossier :",
                                        text=self._stem() or "Recherches_PubMed")
        name = (name or "").strip()
        if not ok or not name:
            return
        path = os.path.join(parent, name)
        try:
            os.makedirs(path, exist_ok=True)
        except OSError as e:
            QMessageBox.warning(self, "Dossier non créé", f"Impossible de créer {path} :\n{e}")
            return
        self.outdir.setText(path)


    # ------------------------------------------------------------ file d'attente
    STATUS = {"attente": "En attente", "cours": "En cours", "ok": "Terminée",
              "arret": "Arrêtée", "erreur": "Erreur"}
    SORTS = {"best": "Pertinence", "recent": "Plus récents", "both": "Les deux"}

    def _build_queue_tab(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(28, 24, 28, 28)
        lay.setSpacing(6)
        lay.addWidget(hint("Préparez plusieurs recherches dans l'onglet Recherche avec « Ajouter à la file », "
                           "puis lancez-les ici l'une après l'autre, tout de suite ou à une heure choisie. "
                           "Les identifiants utilisés sont ceux des Réglages au moment du lancement."))
        lay.addSpacing(12)
        self.qtable = QTableWidget(0, 5)
        self.qtable.setObjectName("queue")
        self.qtable.setHorizontalHeaderLabels(["État", "Fichier Excel", "Requête", "Tri / max", "Résultat"])
        self.qtable.verticalHeader().hide()
        self.qtable.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.qtable.setSelectionMode(QAbstractItemView.SingleSelection)
        self.qtable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.qtable.setWordWrap(False)
        self.qtable.setShowGrid(False)
        self.qtable.setAlternatingRowColors(True)
        hh = self.qtable.horizontalHeader()
        for c, wdt in ((0, 84), (1, 150), (3, 112), (4, 104)):
            hh.setSectionResizeMode(c, QHeaderView.Interactive)
            self.qtable.setColumnWidth(c, wdt)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        hh.setMinimumSectionSize(60)
        self.qtable.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.qtable.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        self.qtable.setTextElideMode(Qt.ElideRight)
        hh.setHighlightSections(False)
        self.qtable.setMinimumHeight(260)
        self.qtable.itemSelectionChanged.connect(self._queue_buttons)
        self.qtable.cellDoubleClicked.connect(self._queue_open_result)
        lay.addWidget(self.qtable, 1)
        self.queue_empty = hint("La file est vide. Remplissez l'onglet Recherche puis cliquez sur "
                                "« Ajouter à la file ».")
        lay.addWidget(self.queue_empty)

        row = QGridLayout()
        row.setHorizontalSpacing(6)
        row.setVerticalSpacing(6)
        self.q_up = QPushButton("Monter")
        self.q_down = QPushButton("Descendre")
        self.q_edit = QPushButton("Reprendre dans Recherche")
        self.q_del = QPushButton("Retirer")
        self.q_clear = QPushButton("Retirer les terminées")
        for i, b in enumerate((self.q_up, self.q_down, self.q_del)):
            b.setObjectName("secondary")
            row.addWidget(b, 0, i)
        for i, b in enumerate((self.q_edit, self.q_clear)):
            b.setObjectName("secondary")
            row.addWidget(b, 1, i * 2, 1, 2 if i == 0 else 1)
        row.setColumnStretch(3, 1)
        self.q_up.clicked.connect(lambda: self._queue_move(-1))
        self.q_down.clicked.connect(lambda: self._queue_move(1))
        self.q_edit.clicked.connect(self._queue_edit)
        self.q_del.clicked.connect(self._queue_remove)
        self.q_clear.clicked.connect(self._queue_clear_done)
        lay.addLayout(row)
        lay.addSpacing(18)

        lay.addWidget(section_title("Lancement"))
        srow = QHBoxLayout()
        self.q_schedule = QCheckBox("Programmer le lancement de la file")
        self.q_at = QDateTimeEdit(QDateTime.currentDateTime().addSecs(3600))
        self.q_at.setDisplayFormat("dd/MM/yyyy  HH:mm")
        self.q_at.setCalendarPopup(True)
        self.q_at.setEnabled(False)
        self.q_schedule.toggled.connect(self.q_at.setEnabled)
        self.q_schedule.toggled.connect(lambda on: self.q_run.setText(
            "Programmer la file" if on else "Lancer la file"))
        srow.addWidget(self.q_schedule)
        srow.addWidget(self.q_at)
        srow.addStretch(1)
        lay.addLayout(srow)
        lay.addWidget(hint("L'application et l'ordinateur doivent rester allumés jusqu'à l'heure prévue."))
        lay.addSpacing(12)
        brow = QHBoxLayout()
        brow.setSpacing(10)
        self.q_cancel = QPushButton("Annuler la programmation")
        self.q_cancel.setObjectName("secondary")
        self.q_cancel.setMinimumHeight(44)
        self.q_cancel.hide()
        self.q_cancel.clicked.connect(self._queue_cancel_schedule)
        self.q_run = QPushButton("Lancer la file")
        self.q_run.setObjectName("primary")
        self.q_run.setMinimumHeight(44)
        self.q_run.clicked.connect(self._queue_start_clicked)
        brow.addWidget(self.q_cancel, 1)
        brow.addWidget(self.q_run, 2)
        lay.addLayout(brow)
        self.q_status = hint("")
        lay.addWidget(self.q_status)
        return w

    def _queue_refresh(self):
        self.qtable.setRowCount(len(self.queue))
        colors = {"cours": C["hema"], "ok": C["ok"], "erreur": C["err"], "arret": C["warn"]}
        for r, job in enumerate(self.queue):
            filt = (" + " + ", ".join(job["filters"])) if job.get("filters") else ""
            cells = [self.STATUS.get(job["status"], job["status"]), job["stem"] + ".xlsx",
                     (job.get("base_query") or "(filtres seuls)") + filt,
                     f'{self.SORTS.get(job["sort"], job["sort"])} / {job["max"]}'
                     + ("" if job["getpdf"] else " / sans PDF"),
                     job.get("summary", "")]
            for c, text in enumerate(cells):
                it = QTableWidgetItem(text)
                if c == 2:
                    it.setToolTip(job["query"])
                if c == 0:
                    f = it.font()
                    f.setBold(True)
                    it.setFont(f)
                    if job["status"] in colors:
                        it.setForeground(QColor(colors[job["status"]]))
                self.qtable.setItem(r, c, it)
        n_wait = sum(1 for j in self.queue if j["status"] == "attente")
        self.tabs.setTabText(self.TAB_QUEUE, f"File d'attente ({n_wait})" if n_wait else "File d'attente")
        self.queue_empty.setVisible(not self.queue)
        self.qtable.setVisible(bool(self.queue))
        self._queue_buttons()

    def _queue_sel(self):
        rows = self.qtable.selectionModel().selectedRows() if self.qtable.selectionModel() else []
        return rows[0].row() if rows else -1

    def _queue_buttons(self):
        r = self._queue_sel()
        movable = r >= 0 and self.queue[r]["status"] != "cours"
        self.q_up.setEnabled(movable and r > 0)
        self.q_down.setEnabled(movable and r < len(self.queue) - 1)
        self.q_del.setEnabled(movable)
        self.q_edit.setEnabled(r >= 0)
        self.q_clear.setEnabled(any(j["status"] in ("ok", "arret", "erreur") for j in self.queue))
        n_wait = sum(1 for j in self.queue if j["status"] == "attente")
        self.q_run.setEnabled(n_wait > 0 and not self.queue_active and self.proc is None)

    def _add_to_queue(self):
        job = self._collect_job()
        if not job:
            return
        self.queue.append(job)
        self._queue_refresh()
        self._save_settings()
        n = sum(1 for j in self.queue if j["status"] == "attente")
        self.queue_note.setText(f"« {job['stem']} » ajoutée à la file ({n} en attente). "
                                "Modifiez le formulaire pour préparer la suivante.")

    def _queue_move(self, d):
        r = self._queue_sel()
        t = r + d
        if r < 0 or not 0 <= t < len(self.queue):
            return
        self.queue[r], self.queue[t] = self.queue[t], self.queue[r]
        self._queue_refresh()
        self.qtable.selectRow(t)
        self._save_settings()

    def _queue_remove(self):
        r = self._queue_sel()
        if r >= 0 and self.queue[r]["status"] != "cours":
            del self.queue[r]
            self._queue_refresh()
            self._save_settings()

    def _queue_clear_done(self):
        self.queue = [j for j in self.queue if j["status"] in ("attente", "cours")]
        self._queue_refresh()
        self._save_settings()

    def _queue_edit(self):
        """Recharge une recherche de la file dans le formulaire Recherche."""
        r = self._queue_sel()
        if r < 0:
            return
        job = self.queue[r]
        self.query.setPlainText(job.get("base_query", ""))
        for _, label, _, cb in self.filter_boxes:
            cb.setChecked(label in job.get("filters", []))
        self.excel.setText(job["stem"])
        self.outdir.setText(job["outdir"])
        self.max_results.setValue(job["max"])
        self.passes.setValue(job["passes"])
        self.getpdf.setChecked(job["getpdf"])
        self.sort_group.button({"best": 0, "recent": 1, "both": 2}.get(job["sort"], 0)).setChecked(True)
        self.tabs.setCurrentIndex(self.TAB_SEARCH)

    def _queue_open_result(self, row, _col):
        res = self.queue[row].get("results", "")
        if res and os.path.isdir(res):
            self._open_path(res)

    def _queue_start_clicked(self):
        if not self._validate_common() or not self._check_email():
            return
        if self.q_schedule.isChecked():
            when = self.q_at.dateTime()
            if when <= QDateTime.currentDateTime():
                self.q_status.setText("Choisissez une heure dans le futur, ou décochez la programmation.")
                return
            self.queue_due = when
            self.queue_timer.start()
            self.q_cancel.show()
            self.q_run.setEnabled(False)
            self._queue_tick()
        else:
            self._queue_begin()

    def _queue_cancel_schedule(self):
        self.queue_timer.stop()
        self.queue_due = None
        self.q_cancel.hide()
        self.q_status.setText("Programmation annulée.")
        self._queue_buttons()

    def _queue_tick(self):
        due = getattr(self, "queue_due", None)
        if not due:
            return
        secs = QDateTime.currentDateTime().secsTo(due)
        if secs <= 0:
            self.queue_timer.stop()
            self.queue_due = None
            self.q_cancel.hide()
            self._queue_begin()
            return
        h, m = divmod(secs // 60, 60)
        self.q_status.setText(f"Lancement prévu le {due.toString('dd/MM/yyyy à HH:mm')} "
                              f"(dans {h} h {m:02d} min). Laissez l'application ouverte.")

    def _queue_begin(self):
        if self.proc is not None:
            self.q_status.setText("Une recherche est déjà en cours : la file démarrera à sa fin.")
            QTimer.singleShot(10000, self._queue_begin)
            return
        self.queue_active = True
        self.queue_stop = False
        self._queue_next()

    def _queue_next(self):
        script = self.script_path.text().strip()
        email = self.email.text().strip()
        nxt = next((j for j in self.queue if j["status"] == "attente"), None)
        if self.queue_stop or nxt is None or not os.path.exists(script):
            done = sum(1 for j in self.queue if j["status"] == "ok")
            self.queue_active = False
            self.current_job = None
            self.q_status.setText(("File interrompue. " if self.queue_stop else "File terminée. ")
                                  + f"{done} recherche(s) terminée(s).")
            self._queue_refresh()
            self._save_settings()
            return
        try:
            os.makedirs(nxt["outdir"], exist_ok=True)
        except OSError as e:
            nxt["status"], nxt["summary"] = "erreur", f"dossier impossible : {e}"
            self._queue_refresh()
            QTimer.singleShot(0, self._queue_next)
            return
        nxt["status"] = "cours"
        self.current_job = nxt
        pos = self.queue.index(nxt) + 1
        self.q_status.setText(f"Recherche {pos} sur {len(self.queue)} en cours : {nxt['stem']}.")
        self._queue_refresh()
        self._launch(script, self._job_args(nxt, email), "search")

    def _queue_job_finished(self, code):
        job = self.current_job
        if not job:
            return
        job["status"] = "ok" if code == 0 else ("arret" if code == 130 else "erreur")
        job["results"] = self.results_dir
        ok_txt = self.s_ok.value.text()
        rate = self.s_rate.value.text()
        job["summary"] = (f"{ok_txt} PDF" + (f", {rate}" if rate not in ("–", "") else "")) \
            if ok_txt not in ("–", "") else ("erreur" if code not in (0, 130) else "")
        if code == 130:
            self.queue_stop = True
        self._queue_refresh()
        self._save_settings()
        QTimer.singleShot(1500, self._queue_next)

    # ------------------------------------------------------------- import
    def _build_import_tab(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(28, 24, 28, 28)
        lay.setSpacing(6)
        lay.addWidget(hint("Pour les articles gratuits que le script n'a pas pu télécharger : "
                           "ouvrez leurs liens, enregistrez les PDF dans un dossier, puis importez-les. "
                           "Ils sont reconnus par leur DOI, leur titre ou un PMID dans le nom du fichier, "
                           "renommés et ajoutés à l'Excel."))
        lay.addSpacing(16)

        self.imp_results = QLineEdit()
        self.imp_results.setPlaceholderText("Dossier créé par une recherche (contient le fichier Excel)")
        self.imp_results.textChanged.connect(self._on_import_results_changed)
        b = QPushButton("Choisir…")
        b.setObjectName("secondary")
        b.clicked.connect(lambda: self._pick_dir(self.imp_results, "Dossier de résultats"))
        row = QHBoxLayout()
        row.setSpacing(6)
        row.addWidget(self.imp_results, 1)
        row.addWidget(b)
        lay.addWidget(section_title("Étape 1 : ouvrir les liens"))
        lay.addWidget(field_label("Dossier de résultats", self.imp_results))
        lay.addLayout(row)
        self.err_imp_results = error_label()
        lay.addWidget(self.err_imp_results)
        self.links_info = hint("")
        lay.addWidget(self.links_info)
        lrow = QHBoxLayout()
        self.open_links_btn = QPushButton("Ouvrir les 10 liens suivants")
        self.open_links_btn.setObjectName("secondary")
        self.open_links_btn.setEnabled(False)
        self.open_links_btn.clicked.connect(self._open_next_links)
        lrow.addWidget(self.open_links_btn)
        lrow.addStretch(1)
        lay.addLayout(lrow)
        lay.addSpacing(22)

        lay.addWidget(section_title("Étape 2 : importer les PDF enregistrés"))
        self.imp_pdfs = QLineEdit()
        dl = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation) or os.path.expanduser("~")
        self.imp_pdfs.setPlaceholderText(os.path.join(dl, "pdf_import"))
        b2 = QPushButton("Choisir…")
        b2.setObjectName("secondary")
        b2.clicked.connect(lambda: self._pick_dir(self.imp_pdfs, "Dossier des PDF téléchargés"))
        row2 = QHBoxLayout()
        row2.setSpacing(6)
        row2.addWidget(self.imp_pdfs, 1)
        row2.addWidget(b2)
        lay.addWidget(field_label("Dossier des PDF enregistrés", self.imp_pdfs))
        lay.addLayout(row2)
        self.err_imp_pdfs = error_label()
        lay.addWidget(self.err_imp_pdfs)
        lay.addWidget(hint("Utilisez un dossier réservé à ces PDF : tous les PDF qu'il contient sont examinés."))
        lay.addSpacing(8)
        self.imp_move = QCheckBox("Déplacer les PDF au lieu de les copier")
        lay.addWidget(self.imp_move)
        lay.addSpacing(22)
        self.imp_btn = QPushButton("Importer les PDF")
        self.imp_btn.setObjectName("primary")
        self.imp_btn.setMinimumHeight(44)
        self.imp_btn.clicked.connect(self._start_import)
        lay.addWidget(self.imp_btn)
        lay.addStretch(1)
        return w

    def _pick_dir(self, edit, title):
        d = QFileDialog.getExistingDirectory(self, title, edit.text().strip() or os.path.expanduser("~"))
        if d:
            edit.setText(d)

    def _manual_links(self):
        p = os.path.join(self.imp_results.text().strip(), "error.txt")
        try:
            with open(p, encoding="utf-8") as f:
                return re.findall(r"À ouvrir manuellement : (\S+)", f.read())
        except OSError:
            return []

    def _on_import_results_changed(self):
        self._clear_error(self.err_imp_results)
        self.link_cursor = 0
        links = self._manual_links()
        if links:
            self.links_info.setText(f"{len(links)} articles gratuits à récupérer à la main dans error.txt.")
        elif self.imp_results.text().strip():
            self.links_info.setText("Aucun lien « À ouvrir manuellement » trouvé dans error.txt de ce dossier.")
        else:
            self.links_info.setText("")
        self.open_links_btn.setEnabled(bool(links))
        self.open_links_btn.setText("Ouvrir les 10 premiers liens" if links else "Ouvrir les 10 liens suivants")

    def _open_next_links(self):
        links = self._manual_links()
        batch = links[self.link_cursor:self.link_cursor + 10]
        for u in batch:
            QDesktopServices.openUrl(QUrl(u))
        self.link_cursor += len(batch)
        left = len(links) - self.link_cursor
        self.links_info.setText(f"Liens ouverts : {self.link_cursor} sur {len(links)}."
                                + ("" if left else " Tous les liens ont été ouverts."))
        self.open_links_btn.setEnabled(left > 0)
        self.open_links_btn.setText("Ouvrir les 10 liens suivants")

    # ------------------------------------------------------------ réglages
    def _build_settings_tab(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(28, 24, 28, 28)
        lay.setSpacing(6)

        lay.addWidget(section_title("Script de recherche"))
        self.script_path = QLineEdit()
        self.script_path.setReadOnly(True)
        b = QPushButton("Choisir…")
        b.setObjectName("secondary")
        b.clicked.connect(self._choose_script)
        b2 = QPushButton("Recharger")
        b2.setObjectName("secondary")
        b2.clicked.connect(self._refresh_script)
        b3 = QPushButton("Script fourni")
        b3.setObjectName("secondary")
        b3.setToolTip("Revenir à la version de pubmed_search.py livrée avec l'application")
        b3.clicked.connect(self._use_bundled_script)
        row = QHBoxLayout()
        row.setSpacing(6)
        row.addWidget(self.script_path, 1)
        row.addWidget(b)
        row.addWidget(b2)
        row.addWidget(b3)
        lay.addWidget(field_label("Fichier pubmed_search.py", self.script_path))
        lay.addLayout(row)
        self.script_status = hint("")
        lay.addWidget(self.script_status)
        lay.addSpacing(22)

        lay.addWidget(section_title("Identifiants"))
        self.email = QLineEdit()
        self.email.setPlaceholderText("nom@exemple.org")
        self.email.textChanged.connect(lambda: self._clear_error(self.err_email))
        lay.addWidget(field_label("Email (transmis à NCBI et Unpaywall)", self.email))
        lay.addWidget(self.email)
        self.err_email = error_label()
        lay.addWidget(self.err_email)
        lay.addSpacing(8)
        self.k_ncbi = SecretEdit("Facultative : 10 requêtes/s au lieu de 3")
        self.k_core = SecretEdit("Facultative : copies des dépôts d'universités")
        self.k_els = SecretEdit("Facultative")
        self.k_oa = SecretEdit("Facultative")
        for lab, wdg in (("Clé API NCBI (PubMed)", self.k_ncbi), ("Clé API CORE", self.k_core),
                         ("Clé API Elsevier", self.k_els), ("Clé API OpenAlex", self.k_oa)):
            lay.addWidget(field_label(lab, wdg.edit))
            lay.addWidget(wdg)
            lay.addSpacing(4)
        self.remember_keys = QCheckBox("Mémoriser les clés API sur cet ordinateur")
        lay.addWidget(self.remember_keys)
        lay.addWidget(hint(f"Les clés sont alors enregistrées en clair dans {self.settings_path}. "
                           "Elles sont transmises au script par variables d'environnement et "
                           "n'apparaissent ni dans le journal ni dans error.txt."))
        lay.addSpacing(22)

        lay.addWidget(section_title("Filtres de recherche"))
        lay.addWidget(hint("Régions, niveaux de développement et types d'articles proposés dans l'onglet "
                           "Recherche sont définis dans un fichier modifiable (libellés et requêtes PubMed). "
                           "Les modifications s'appliquent au prochain démarrage."))
        frow = QHBoxLayout()
        fb = QPushButton("Modifier les filtres…")
        fb.setObjectName("secondary")
        fb.clicked.connect(self._edit_filters)
        frow.addWidget(fb)
        frow.addStretch(1)
        lay.addLayout(frow)
        lay.addSpacing(22)

        lay.addWidget(section_title("Interpréteur Python"))
        self.py_builtin = QRadioButton("Python intégré à l'application")
        self.py_custom = QRadioButton("Autre interpréteur Python")
        self.py_builtin.setChecked(True)
        grp = QButtonGroup(self)
        grp.addButton(self.py_builtin)
        grp.addButton(self.py_custom)
        lay.addWidget(self.py_builtin)
        lay.addWidget(self.py_custom)
        self.py_path = QLineEdit()
        self.py_path.setPlaceholderText("/usr/bin/python3  ou  C:\\Python312\\python.exe")
        self.py_path.setEnabled(False)
        self.py_custom.toggled.connect(self.py_path.setEnabled)
        self.py_custom.toggled.connect(lambda *_: self._refresh_script())
        lay.addWidget(self.py_path)
        lay.addWidget(hint("Le Python intégré contient requests, openpyxl et pypdf. Choisissez un autre "
                           "interpréteur seulement si une nouvelle version du script demande d'autres bibliothèques."))
        lay.addStretch(1)
        return w

    def _edit_filters(self):
        dst = user_filters_path()
        if not os.path.exists(dst):
            try:
                import shutil
                shutil.copyfile(os.path.join(RES, "filtres.json"), dst)
            except OSError as e:
                QMessageBox.warning(self, "Filtres", f"Copie impossible : {e}")
                return
        QDesktopServices.openUrl(QUrl.fromLocalFile(dst))
        QMessageBox.information(self, "Filtres",
                                f"Le fichier des filtres s'ouvre dans votre éditeur :\n{dst}\n\n"
                                "Enregistrez-le puis redémarrez l'application pour voir les changements.")

    # ---------------------------------------------------------------- log
    def _build_log_panel(self):
        w = QFrame()
        w.setObjectName("logPanel")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)
        top = QHBoxLayout()
        t = QLabel("Journal")
        t.setObjectName("logTitle")
        self.state = QLabel("Prêt")
        self.state.setObjectName("stateChip")
        top.addWidget(t)
        top.addStretch(1)
        top.addWidget(self.state)
        lay.addLayout(top)

        stats = QHBoxLayout()
        stats.setSpacing(28)
        self.s_ok = Stat("PDF obtenus", C["eosin"])
        self.s_miss = Stat("sans PDF", "#9EA1E6")
        self.s_left = Stat("restants", C["log_fg"])
        self.s_rate = Stat("réussite", C["log_fg"])
        for s in (self.s_ok, self.s_miss, self.s_left, self.s_rate):
            stats.addWidget(s)
        stats.addStretch(1)
        lay.addLayout(stats)
        self.strip = ProgressStrip()
        lay.addWidget(self.strip)
        self.phase = QLabel("Remplissez la recherche puis lancez-la.")
        self.phase.setObjectName("phase")
        lay.addWidget(self.phase)

        self.log = QPlainTextEdit()
        self.log.setObjectName("log")
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(20000)
        self.log.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.log.setPlaceholderText("La sortie du script s'affichera ici : recherche PubMed, "
                                    "téléchargement article par article, puis statistiques finales.")
        lay.addWidget(self.log, 1)

        bot = QHBoxLayout()
        bot.setSpacing(8)
        self.stop_btn = QPushButton("Arrêter")
        self.stop_btn.setObjectName("danger")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop)
        self.open_dir_btn = QPushButton("Ouvrir le dossier")
        self.open_xlsx_btn = QPushButton("Ouvrir l'Excel")
        self.open_err_btn = QPushButton("Ouvrir error.txt")
        for b in (self.open_dir_btn, self.open_xlsx_btn, self.open_err_btn):
            b.setObjectName("onDark")
            b.setEnabled(False)
        self.open_dir_btn.clicked.connect(lambda: self._open_path(self.results_dir))
        self.open_xlsx_btn.clicked.connect(lambda: self._open_path(self.excel_path))
        self.open_err_btn.clicked.connect(lambda: self._open_path(os.path.join(self.results_dir, "error.txt")))
        copy = QPushButton("Copier")
        copy.setObjectName("onDark")
        copy.setToolTip("Copier tout le journal")
        copy.clicked.connect(lambda: QApplication.clipboard().setText(self.log.toPlainText()))
        bot.addWidget(self.stop_btn)
        bot.addStretch(1)
        bot.addWidget(self.open_dir_btn)
        bot.addWidget(self.open_xlsx_btn)
        bot.addWidget(self.open_err_btn)
        bot.addWidget(copy)
        lay.addLayout(bot)
        return w

    # ----------------------------------------------------------- script
    def _choose_script(self):
        start = os.path.dirname(self.script_path.text()) or os.path.expanduser("~")
        p, _ = QFileDialog.getOpenFileName(self, "Choisir pubmed_search.py", start, "Script Python (*.py)")
        if p:
            self.script_path.setText(p)
            self._refresh_script()

    def _use_bundled_script(self):
        p = bundled_script()
        if p:
            self.script_path.setText(p)
            self._refresh_script()

    def _python_cmd(self, script, args):
        if self.py_custom.isChecked() and self.py_path.text().strip():
            return self.py_path.text().strip(), ["-u", script] + args
        if FROZEN:
            return sys.executable, ["--run-script", script] + args
        return sys.executable, [os.path.abspath(__file__), "--run-script", script] + args

    def _refresh_script(self):
        path = self.script_path.text().strip()
        if not path or not os.path.exists(path):
            self.script_chip.setText("Aucun script chargé")
            self.script_chip.setProperty("state", "bad")
            self.script_status.setText("Choisissez le fichier pubmed_search.py à utiliser.")
            self._repolish(self.script_chip)
            self.script_flags = set()
            return
        ver = script_version(path)
        self.script_chip.setText(f"pubmed_search.py  v{ver}")
        self.script_chip.setToolTip(path)
        self.script_chip.setProperty("state", "ok")
        self._repolish(self.script_chip)
        self.script_status.setText(f"Version {ver}. Lecture des options du script…")
        # Lit les options réellement proposées par le script (--help) : l'interface
        # désactive ce que cette version du script ne connaît pas.
        prog, args = self._python_cmd(path, ["--help"])
        probe = QProcess(self)
        probe.setProcessChannelMode(QProcess.MergedChannels)
        probe.setProcessEnvironment(self._env())

        def done(*_):
            out = bytes(probe.readAll()).decode("utf-8", "replace")
            flags = set(re.findall(r"(--[a-z][a-z0-9\-]+)", out))
            self.script_flags = flags
            if "--output" not in flags:
                self.script_status.setText(f"Version {ver}. Le script n'a pas répondu à --help : "
                                           f"vérifiez qu'il s'agit bien de pubmed_search.py.\n{out[-300:]}")
            else:
                missing = [f for f in ("--core-key", "--elsevier-key", "--openalex-key", "--import-pdf",
                                       "--time-set") if f not in flags]
                txt = f"Version {ver}, prête."
                if missing:
                    txt += " Options absentes de cette version : " + ", ".join(missing) + "."
                self.script_status.setText(txt)
            self.k_core.setEnabled("--core-key" in flags)
            self.k_els.setEnabled("--elsevier-key" in flags)
            self.k_oa.setEnabled("--openalex-key" in flags)
            self.imp_btn.setEnabled("--import-pdf" in flags)
            self.schedule.setEnabled("--time-set" in flags)
            probe.deleteLater()

        probe.finished.connect(done)
        probe.errorOccurred.connect(lambda *_: self.script_status.setText(
            "Impossible de lancer l'interpréteur Python choisi."))
        probe.start(prog, args)

    @staticmethod
    def _repolish(w):
        w.style().unpolish(w)
        w.style().polish(w)

    # ------------------------------------------------------------- exécution
    def _env(self):
        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONUNBUFFERED", "1")
        env.insert("PYTHONIOENCODING", "utf-8")
        for var in ("NCBI_API_KEY", "CORE_API_KEY", "ELSEVIER_API_KEY", "OPENALEX_API_KEY"):
            env.remove(var)
        pairs = (("NCBI_API_KEY", self.k_ncbi), ("CORE_API_KEY", self.k_core),
                 ("ELSEVIER_API_KEY", self.k_els), ("OPENALEX_API_KEY", self.k_oa))
        for var, wdg in pairs:
            if wdg.isEnabled() and wdg.text():
                env.insert(var, wdg.text())
        return env

    def _set_error(self, lab, text, focus=None):
        lab.setText(text)
        lab.show()
        if focus is not None:
            focus.setFocus()

    def _clear_error(self, lab):
        lab.hide()

    def _validate_common(self):
        script = self.script_path.text().strip()
        if not script or not os.path.exists(script):
            self.tabs.setCurrentIndex(self.TAB_SETTINGS)
            QMessageBox.warning(self, "Script introuvable",
                                "Choisissez le fichier pubmed_search.py dans Réglages.")
            return None
        if self.proc is not None:
            return None
        return script

    def _check_email(self):
        email = self.email.text().strip()
        if not EMAIL_RE.match(email):
            self.tabs.setCurrentIndex(self.TAB_SETTINGS)
            self._set_error(self.err_email, "Saisissez un email valide : NCBI et Unpaywall l'exigent.", self.email)
            return None
        return email

    def _collect_job(self):
        """Valide le formulaire Recherche et renvoie la recherche à exécuter (dict) ou None."""
        ok = True
        q = self._final_query()
        if not q:
            self._set_error(self.err_query, "Saisissez une requête PubMed ou cochez au moins un filtre.",
                            self.query)
            ok = False
        outdir = self.outdir.text().strip()
        if not outdir:
            self._set_error(self.err_outdir, "Choisissez le dossier de destination.", self.outdir)
            ok = False
        elif not os.path.isdir(outdir):
            r = QMessageBox.question(self, "Dossier inexistant",
                                     f"Le dossier n'existe pas :\n{outdir}\n\nLe créer ?")
            if r != QMessageBox.Yes:
                return
            try:
                os.makedirs(outdir, exist_ok=True)
            except OSError as e:
                self._set_error(self.err_outdir, f"Création impossible : {e}", self.outdir)
                ok = False
        stem = self._stem()
        if not stem:
            self._set_error(self.err_excel, "Donnez un nom au fichier Excel, par exemple Diabetes_review.",
                            self.excel)
            ok = False
        elif re.search(r"[^A-Za-z0-9_\-]", stem):
            self._set_error(self.err_excel, "Utilisez seulement des lettres sans accent, des chiffres, "
                                            "- et _ (le nom sert aussi pour le dossier).", self.excel)
            ok = False
        if not ok:
            return None
        _, labels = self._selected_filters()
        return {"stem": stem, "query": q, "base_query": " ".join(self.query.toPlainText().split()),
                "filters": labels, "outdir": outdir, "max": self.max_results.value(),
                "sort": self._sort_key(), "getpdf": self.getpdf.isChecked(),
                "passes": self.passes.value(), "status": "attente", "summary": ""}

    @staticmethod
    def _job_args(job, email):
        args = [job["query"], "--output", job["stem"] + ".xlsx", "--email", email,
                "--max-results", str(job["max"]), "--outdir", job["outdir"]]
        if job["sort"] in ("best", "both"):
            args.append("--best-match")
        if job["sort"] in ("recent", "both"):
            args.append("--most-recent")
        if job["getpdf"]:
            args.append("--getfreepaper")
            if job["passes"]:
                args += ["--pass-number", str(job["passes"])]
        return args

    def _start_search(self):
        script = self._validate_common()
        if not script:
            return
        email = self._check_email()
        if not email:
            return
        job = self._collect_job()
        if not job:
            return
        args = self._job_args(job, email)
        if self.schedule.isChecked() and self.schedule.isEnabled():
            args += ["--time-set", self.schedule_at.dateTime().toString("dd-MM-yyyy_HH'h'mm")]
        self.current_job = None
        self._save_settings()
        self._launch(script, args, "search")

    def _start_import(self):
        script = self._validate_common()
        if not script:
            return
        res = self.imp_results.text().strip()
        pdfs = self.imp_pdfs.text().strip()
        ok = True
        if not os.path.isdir(res) or not any(f.lower().endswith(".xlsx") for f in os.listdir(res)):
            self._set_error(self.err_imp_results, "Choisissez un dossier créé par une recherche "
                                                  "(il doit contenir le fichier Excel).", self.imp_results)
            ok = False
        if not os.path.isdir(pdfs):
            self._set_error(self.err_imp_pdfs, "Choisissez le dossier où vous avez enregistré les PDF.",
                            self.imp_pdfs)
            ok = False
        if not ok:
            return
        args = ["--import-pdf", pdfs, "--results", res]
        if self.imp_move.isChecked():
            args.append("--move")
        self.results_dir = res
        self._save_settings()
        self._launch(script, args, "import")

    def _launch(self, script, args, kind):
        self.log.clear()
        self._counts = {"total": 0, "ok": 0, "miss": 0}
        self.strip.set_values(0, 0, 0)
        for s in (self.s_ok, self.s_miss, self.s_left, self.s_rate):
            s.set("–")
        if kind == "search":
            self.results_dir = ""
            self.excel_path = ""
        for b in (self.open_dir_btn, self.open_xlsx_btn, self.open_err_btn):
            b.setEnabled(False)
        prog, full = self._python_cmd(script, args)
        shown = " ".join(f'"{a}"' if " " in a else a for a in ["pubmed_search.py"] + args)
        self._append(f"$ {shown}\n", "dim")
        self.proc = QProcess(self)
        self.proc.setProcessChannelMode(QProcess.MergedChannels)
        self.proc.setProcessEnvironment(self._env())
        self.proc.setWorkingDirectory(os.path.dirname(script))
        self.proc.readyReadStandardOutput.connect(self._read_output)
        self.proc.finished.connect(self._finished)
        self.proc.errorOccurred.connect(self._proc_error)
        self.run_kind = kind
        self._set_state("En cours", "run")
        self.phase.setText("Recherche dans PubMed…" if kind == "search" else "Import des PDF…")
        self.run_btn.setEnabled(False)
        self.imp_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.proc.start(prog, full)
        self._queue_buttons()

    def _stop(self):
        if not self.proc:
            return
        self.phase.setText("Arrêt demandé : le script enregistre les résultats partiels…")
        self.proc.write(b"STOP\n")
        if self.py_custom.isChecked() and os.name == "posix":
            import signal
            try:
                os.kill(self.proc.processId(), signal.SIGINT)
            except OSError:
                pass
        QTimer.singleShot(20000, lambda: self.proc and self.proc.kill())

    def _proc_error(self, err):
        if err == QProcess.FailedToStart:
            self._append("Le script n'a pas pu démarrer : vérifiez l'interpréteur Python dans Réglages.\n", "err")

    def _read_output(self):
        data = bytes(self.proc.readAllStandardOutput()).decode("utf-8", "replace")
        self._buf += data
        *lines, self._buf = self._buf.split("\n")
        for line in lines:
            self._handle_line(line.rstrip("\r"))

    def _handle_line(self, line):
        kind = None
        if "⚠" in line:
            kind = "warn"
        elif "❌" in line or line.startswith("Traceback"):
            kind = "err"
        m = re.match(r"\s*\[(\d+)/(\d+)\] PMID \d+ .*→ (.*)$", line)
        if m:
            i, n, res = int(m.group(1)), int(m.group(2)), m.group(3)
            c = self._counts
            c["total"] = max(c["total"], n)
            if res.startswith("OK"):
                c["ok"] += 1
                kind = "ok"
            elif "non téléchargé" in res:
                c["miss"] += 1
            self.strip.set_values(c["total"], c["ok"], c["miss"])
            self.s_ok.set(c["ok"])
            self.s_miss.set(c["miss"])
            self.s_left.set(max(c["total"] - i, 0))
            self.phase.setText(f"Téléchargement : article {i} sur {n}")
        m = re.match(r"\s*⬇ Passe (\d+) : (\d+) article", line)
        if m:
            self._counts = {"total": int(m.group(2)), "ok": 0, "miss": 0}
            self.strip.set_values(int(m.group(2)), 0, 0)
            self.phase.setText(f"Passe {m.group(1)} : {m.group(2)} articles à traiter")
        m = re.match(r"Dossier : (.+)$", line)
        if m:
            self.results_dir = m.group(1).strip()
        m = re.match(r"\s*Excel : (.+)$", line)
        if m:
            p = m.group(1).strip()
            self.excel_path = p if os.path.isabs(p) else os.path.join(self.proc.workingDirectory(), p)
        m = re.match(r"\s*PDF téléchargés\s*:\s*(\d+)\s*\(([\d.]+ %|n/a)", line)
        if m:
            self.s_ok.set(m.group(1))
            self.s_rate.set(m.group(2).replace(" ", "\u202f"))
        m = re.match(r"\s*Restant à télécharger\s*:\s*(\d+)", line)
        if m:
            self.s_left.set(m.group(1))
        if "Recherche dans PubMed" in line or "ESearch" in line:
            self.phase.setText("Recherche dans PubMed…")
        elif "Récupération des métadonnées" in line:
            self.phase.setText("Lecture des notices PubMed…")
        elif "Enrichissement" in line:
            self.phase.setText("Recherche des liens de texte intégral…")
        elif "Exécution programmée" in line:
            self.phase.setText(line.strip(" ⏳"))
        self._append(line + "\n", kind)

    def _append(self, text, kind=None):
        colors = {"ok": "#E79AB8", "warn": "#E6B45C", "err": "#F08A82", "dim": C["log_dim"]}
        cur = self.log.textCursor()
        cur.movePosition(QTextCursor.End)
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(colors.get(kind, C["log_fg"])))
        cur.insertText(text, fmt)
        bar = self.log.verticalScrollBar()
        if bar.value() >= bar.maximum() - 40:
            self.log.setTextCursor(cur)
            self.log.ensureCursorVisible()

    def _finished(self, code, status):
        if self._buf:
            self._handle_line(self._buf)
            self._buf = ""
        if code == 0:
            self._set_state("Terminé", "ok")
            self.phase.setText("Recherche terminée." if self.run_kind == "search" else "Import terminé.")
        elif code == 130:
            self._set_state("Arrêté", "warn")
            self.phase.setText("Arrêtée : les résultats partiels ont été enregistrés.")
        else:
            self._set_state("Erreur", "err")
            self.phase.setText(f"Le script s'est arrêté avec une erreur (code {code}). Voir le journal.")
        if self.results_dir and os.path.isdir(self.results_dir):
            self.open_dir_btn.setEnabled(True)
            self.open_err_btn.setEnabled(os.path.exists(os.path.join(self.results_dir, "error.txt")))
            if not self.excel_path:
                xs = [f for f in os.listdir(self.results_dir) if f.lower().endswith(".xlsx")]
                self.excel_path = os.path.join(self.results_dir, xs[0]) if xs else ""
            self.open_xlsx_btn.setEnabled(bool(self.excel_path) and os.path.exists(self.excel_path))
            if self.run_kind == "search":
                self.imp_results.setText(self.results_dir)
        self.proc.deleteLater()
        self.proc = None
        if self.queue_active and self.current_job is not None:
            self._queue_job_finished(code)
        self._queue_buttons()
        self.run_btn.setEnabled(True)
        self.imp_btn.setEnabled("--import-pdf" in self.script_flags or not self.script_flags)
        self.stop_btn.setEnabled(False)

    def _set_state(self, text, state):
        self.state.setText(text)
        self.state.setProperty("state", state)
        self._repolish(self.state)

    def _open_path(self, p):
        if p and os.path.exists(p):
            QDesktopServices.openUrl(QUrl.fromLocalFile(p))

    # ------------------------------------------------------------ réglages
    def _load_settings(self):
        s = {}
        try:
            with open(self.settings_path, encoding="utf-8") as f:
                s = json.load(f)
        except (OSError, ValueError):
            pass
        self.script_path.setText(s.get("script") if s.get("script") and os.path.exists(s["script"])
                                 else bundled_script())
        self.email.setText(s.get("email", ""))
        self.query.setPlainText(s.get("query", ""))
        self.max_results.setValue(int(s.get("max_results", 100)))
        self.passes.setValue(int(s.get("passes", 0)))
        self.getpdf.setChecked(bool(s.get("getpdf", True)))
        idx = {"best": 0, "recent": 1, "both": 2}.get(s.get("sort", "best"), 0)
        self.sort_group.button(idx).setChecked(True)
        self.outdir.setText(s.get("outdir") or os.path.join(
            QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation) or os.path.expanduser("~"),
            "Recherches PubMed"))
        self.excel.setText(s.get("excel", ""))
        self.imp_pdfs.setText(s.get("import_pdfs", ""))
        self.remember_keys.setChecked(bool(s.get("remember_keys", False)))
        keys = s.get("keys", {}) if self.remember_keys.isChecked() else {}
        self.k_ncbi.setText(keys.get("ncbi", ""))
        self.k_core.setText(keys.get("core", ""))
        self.k_els.setText(keys.get("elsevier", ""))
        self.k_oa.setText(keys.get("openalex", ""))
        if s.get("python_custom"):
            self.py_custom.setChecked(True)
        self.py_path.setText(s.get("python_path", ""))
        self.queue = [j for j in s.get("queue", []) if isinstance(j, dict) and j.get("query")]
        for j in self.queue:
            if j.get("status") == "cours":
                j["status"] = "attente"
        self._queue_refresh()
        if s.get("geometry"):
            try:
                self.restoreGeometry(bytes.fromhex(s["geometry"]))
            except Exception:  # noqa
                pass

    def _save_settings(self):
        s = {
            "script": self.script_path.text().strip(),
            "email": self.email.text().strip(),
            "query": self.query.toPlainText(),
            "max_results": self.max_results.value(),
            "passes": self.passes.value(),
            "getpdf": self.getpdf.isChecked(),
            "sort": self._sort_key(),
            "outdir": self.outdir.text().strip(),
            "excel": self.excel.text().strip(),
            "import_pdfs": self.imp_pdfs.text().strip(),
            "remember_keys": self.remember_keys.isChecked(),
            "python_custom": self.py_custom.isChecked(),
            "python_path": self.py_path.text().strip(),
            "geometry": bytes(self.saveGeometry()).hex(),
            "queue": [dict(j, status=("attente" if j["status"] == "cours" else j["status"]))
                      for j in self.queue],
        }
        if self.remember_keys.isChecked():
            s["keys"] = {"ncbi": self.k_ncbi.text(), "core": self.k_core.text(),
                         "elsevier": self.k_els.text(), "openalex": self.k_oa.text()}
        try:
            tmp = self.settings_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(s, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.settings_path)
            if os.name == "posix":
                os.chmod(self.settings_path, 0o600)
        except OSError:
            pass

    def closeEvent(self, e):
        if self.proc is not None:
            r = QMessageBox.question(self, "Recherche en cours",
                                     "Une recherche est en cours. L'arrêter et quitter ?\n"
                                     "Les résultats partiels seront enregistrés.")
            if r != QMessageBox.Yes:
                e.ignore()
                return
            self.proc.write(b"STOP\n")
            self.proc.waitForFinished(15000)
            if self.proc is not None and self.proc.state() != QProcess.NotRunning:
                self.proc.kill()
        self._save_settings()
        e.accept()


# --------------------------------------------------------------------------- #
# Style
# --------------------------------------------------------------------------- #
def stylesheet(ui, mono):
    return f"""
* {{ font-family: "{ui}"; font-size: 14px; color: {C['ink']}; }}
QMainWindow, #formPage, QScrollArea, QScrollArea > QWidget > QWidget {{ background: {C['slide']}; }}
QTabWidget::pane {{ border: none; background: {C['slide']}; }}
QTabBar {{ background: {C['slide']}; }}
QTabBar::tab {{ background: transparent; color: {C['muted']}; padding: 12px 4px 10px 4px;
               margin: 0 10px 0 28px; border: none; border-bottom: 3px solid transparent; font-size: 14px; }}
QTabBar::tab:selected {{ color: {C['ink']}; border-bottom: 3px solid {C['hema']}; font-weight: 600; }}
QTabBar::tab:hover:!selected {{ color: {C['ink']}; }}
#header {{ background: {C['panel']}; border-bottom: 1px solid {C['rule']}; }}
#appTitle {{ font-size: 19px; font-weight: 700; letter-spacing: 0.2px; }}
#scriptChip {{ padding: 5px 10px; border-radius: 6px; font-family: "{mono}"; font-size: 12px; }}
#scriptChip[state="ok"] {{ background: {C['hema_soft']}; color: {C['hema']}; }}
#scriptChip[state="bad"] {{ background: #FBE5E3; color: {C['err']}; }}
#sectionTitle {{ font-size: 16px; font-weight: 700; padding-bottom: 4px; }}
#fieldLabel {{ font-weight: 600; font-size: 13px; color: {C['ink']}; padding-top: 6px; }}
#hint {{ color: {C['muted']}; font-size: 12.5px; }}
#fieldError {{ color: {C['err']}; font-size: 12.5px; font-weight: 600; }}
#preview {{ color: {C['muted']}; font-size: 12.5px; font-family: "{mono}"; background: {C['panel']};
            border: 1px dashed {C['rule']}; border-radius: 6px; padding: 8px 10px; margin-top: 6px; }}
QLineEdit, QSpinBox, QDateTimeEdit, QPlainTextEdit#query {{
    background: {C['panel']}; border: 1px solid {C['rule']}; border-radius: 6px; padding: 7px 9px;
    selection-background-color: {C['hema_soft']}; selection-color: {C['ink']}; }}
QLineEdit:focus, QSpinBox:focus, QDateTimeEdit:focus, QPlainTextEdit#query:focus {{
    border: 2px solid {C['hema']}; padding: 6px 8px; }}
QLineEdit:disabled, QSpinBox:disabled, QDateTimeEdit:disabled {{ color: #A3A4B5; background: #EFEEF4; }}
QLineEdit[readOnly="true"] {{ background: #EFEEF4; color: {C['muted']}; }}
QPlainTextEdit#query {{ font-family: "{mono}"; font-size: 15px; padding: 10px 12px; }}
QPlainTextEdit#query:focus {{ padding: 9px 11px; }}
QLineEdit, QSpinBox, QDateTimeEdit {{ min-height: 22px; }}
QSpinBox {{ padding-right: 30px; }}
QSpinBox::up-button, QSpinBox::down-button {{ subcontrol-origin: border; width: 26px;
    border: none; border-left: 1px solid {C['rule']}; background: transparent; }}
QSpinBox::up-button {{ subcontrol-position: top right; border-top-right-radius: 6px; }}
QSpinBox::down-button {{ subcontrol-position: bottom right; border-bottom-right-radius: 6px; }}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background: {C['hema_soft']}; }}
QSpinBox::up-arrow {{ image: url({os.path.join(RES, 'up.svg').replace(os.sep, '/')}); width: 10px; height: 6px; }}
QSpinBox::down-arrow {{ image: url({os.path.join(RES, 'down.svg').replace(os.sep, '/')}); width: 10px; height: 6px; }}
QPushButton {{ border-radius: 6px; padding: 8px 14px; font-weight: 600; }}
QPushButton#primary {{ background: {C['hema']}; color: white; font-size: 15px; border: none; }}
QPushButton#primary:hover {{ background: #2F327A; }}
QPushButton#primary:pressed {{ background: #262866; }}
QPushButton#primary:disabled {{ background: #A9AAD0; }}
QPushButton#primary:focus {{ outline: none; border: 2px solid {C['eosin']}; }}
QPushButton#secondary {{ background: {C['panel']}; border: 1px solid {C['rule']}; color: {C['ink']}; }}
QPushButton#secondary:hover {{ border-color: {C['hema']}; color: {C['hema']}; }}
QPushButton#secondary:focus {{ border: 2px solid {C['hema']}; }}
QPushButton#secondary:disabled {{ color: #A3A4B5; }}
QPushButton#ghost, QToolButton#ghost {{ background: transparent; border: none; color: {C['hema']};
    padding: 6px 8px; font-weight: 600; }}
QPushButton#ghost:hover, QToolButton#ghost:hover {{ text-decoration: underline; }}
QPushButton#segment {{ background: {C['panel']}; border: 1px solid {C['rule']}; color: {C['ink']};
    padding: 8px 18px; border-radius: 0; font-weight: 600; }}
QPushButton#segment[pos="first"] {{ border-top-left-radius: 6px; border-bottom-left-radius: 6px; }}
QPushButton#segment[pos="last"] {{ border-top-right-radius: 6px; border-bottom-right-radius: 6px; }}
QPushButton#segment[pos="mid"], QPushButton#segment[pos="last"] {{ border-left: none; }}
QPushButton#segment:checked {{ background: {C['hema']}; color: white; border-color: {C['hema']}; }}
QPushButton#segment:focus {{ border: 2px solid {C['eosin']}; }}
QToolButton#disclosure {{ border: none; background: transparent; color: {C['hema']}; font-weight: 600;
    padding: 5px 0; }}
QToolButton#disclosure:focus {{ text-decoration: underline; }}
QPlainTextEdit#finalQuery {{ background: {C['hema_soft']}; border: 1px solid #CFCFE6; border-radius: 6px;
    font-family: "{mono}"; font-size: 12.5px; padding: 6px 8px; }}
QTableWidget#queue {{ background: {C['panel']}; border: 1px solid {C['rule']}; border-radius: 6px;
    alternate-background-color: #FAFAFC; selection-background-color: {C['hema_soft']};
    selection-color: {C['ink']}; font-size: 13px; }}
QTableWidget#queue::item {{ padding: 6px 8px; border: none; }}
QHeaderView::section {{ background: {C['slide']}; border: none; border-bottom: 1px solid {C['rule']};
    padding: 6px 8px; font-weight: 600; font-size: 12.5px; color: {C['muted']}; }}
QCheckBox, QRadioButton {{ spacing: 8px; }}
QCheckBox::indicator, QRadioButton::indicator {{ width: 16px; height: 16px; }}
QCheckBox::indicator {{ border: 1px solid #A7A8BC; border-radius: 4px; background: {C['panel']}; }}
QCheckBox::indicator:checked {{ background: {C['hema']}; border-color: {C['hema']};
    image: url({os.path.join(RES, 'check.svg').replace(os.sep, '/')}); }}
QCheckBox:focus, QRadioButton:focus {{ color: {C['hema']}; }}
#logPanel {{ background: {C['log_bg']}; }}
#logPanel QLabel {{ color: {C['log_fg']}; }}
#logTitle {{ font-size: 16px; font-weight: 700; color: white; }}
#stateChip {{ padding: 4px 10px; border-radius: 10px; font-size: 12px; font-weight: 600;
    background: #2A2B45; color: {C['log_fg']}; }}
#stateChip[state="run"] {{ background: #34377F; color: white; }}
#stateChip[state="ok"] {{ background: #1F4D3B; color: #BDEBD3; }}
#stateChip[state="warn"] {{ background: #5A4314; color: #F3D9A4; }}
#stateChip[state="err"] {{ background: #5C1F1B; color: #F6C0BB; }}
#statValue {{ font-size: 26px; font-weight: 700; }}
#logPanel #statLabel {{ font-size: 12px; color: {C['log_dim']}; }}
#logPanel #phase {{ color: {C['log_dim']}; font-size: 12.5px; }}
QPlainTextEdit#log {{ background: #111222; border: 1px solid #2A2B45; border-radius: 6px;
    font-family: "{mono}"; font-size: 12.5px; color: {C['log_fg']}; padding: 8px;
    selection-background-color: #34377F; }}
QPushButton#onDark {{ background: transparent; border: 1px solid #3A3B5C; color: {C['log_fg']}; }}
QPushButton#onDark:hover {{ border-color: #8C8FE0; color: white; }}
QPushButton#onDark:disabled {{ color: #4E4F6E; border-color: #2A2B45; }}
QPushButton#danger {{ background: transparent; border: 1px solid #7A3A3A; color: #F2B8B5; }}
QPushButton#danger:hover {{ background: #4A1E1E; }}
QPushButton#danger:disabled {{ color: #4E4F6E; border-color: #2A2B45; }}
QSplitter::handle {{ background: {C['rule']}; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: #C9C8D8; border-radius: 4px; min-height: 30px; }}
QPlainTextEdit#log QScrollBar::handle:vertical {{ background: #3A3B5C; }}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
QToolTip {{ background: {C['ink']}; color: white; border: none; padding: 6px 8px; }}
"""


def load_fonts():
    fams = {}
    fdir = os.path.join(RES, "fonts")
    if os.path.isdir(fdir):
        for f in sorted(os.listdir(fdir)):
            if f.endswith(".ttf"):
                fid = QFontDatabase.addApplicationFont(os.path.join(fdir, f))
                for fam in QFontDatabase.applicationFontFamilies(fid):
                    fams[f.split("-")[0]] = fam
    ui = fams.get("AtkinsonHyperlegibleNext", "Sans Serif")
    mono = fams.get("IBMPlexMono", "Monospace")
    return ui, mono


URL_SCHEME = "pubmedsearch"


def inbox_dir():
    d = os.path.join(config_dir(), "inbox")
    os.makedirs(d, exist_ok=True)
    return d


def urls_from_argv(argv):
    return [a for a in argv[1:] if a.lower().startswith(URL_SCHEME + ":")]


def post_to_inbox(url):
    """Dépose une demande (lien pubmedsearch://) pour l'instance déjà ouverte."""
    import time
    name = os.path.join(inbox_dir(), f"{time.time_ns()}.json")
    with open(name + ".tmp", "w", encoding="utf-8") as f:
        json.dump({"url": url}, f)
    os.replace(name + ".tmp", name)


def parse_request(url):
    """pubmedsearch://search?q=…&sort=best|recent|both&action=fill|queue|run&max=…"""
    from urllib.parse import urlsplit, parse_qs
    parts = urlsplit(url)
    qs = {k: v[-1] for k, v in parse_qs(parts.query, keep_blank_values=True).items()}
    req = {"query": qs.get("q", "").strip(),
           "sort": qs.get("sort", "") if qs.get("sort") in ("best", "recent", "both") else "",
           "action": qs.get("action", "fill") if qs.get("action") in ("fill", "queue", "run") else "fill",
           "source": qs.get("source", "")}
    try:
        req["max"] = max(1, min(10000, int(qs.get("max", ""))))
    except ValueError:
        req["max"] = None
    return req


def main():
    QCoreApplication.setApplicationName(APP_ID)
    QCoreApplication.setOrganizationName("")
    urls = urls_from_argv(sys.argv)
    # Instance unique : si l'application est déjà ouverte, on lui transmet la demande et on quitte.
    lock = QLockFile(os.path.join(config_dir(), "instance.lock"))
    lock.setStaleLockTime(0)
    if not lock.tryLock(200):
        if urls:
            for u in urls:
                post_to_inbox(u)
            sys.exit(0)
        # lancement simple alors qu'une fenêtre existe : on la fait passer au premier plan
        post_to_inbox(URL_SCHEME + "://show")
        sys.exit(0)
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except Exception:  # noqa
            pass
    app = QApplication(sys.argv)
    app.setApplicationName(APP_ID)
    app.setApplicationDisplayName(APP_NAME)
    app.setOrganizationName("")
    app.setStyle("Fusion")
    ui, mono = load_fonts()
    f = QFont(ui)
    f.setPointSizeF(10.5)
    app.setFont(f)
    app.setStyleSheet(stylesheet(ui, mono))
    win = MainWindow()
    win.show()
    for u in urls:
        QTimer.singleShot(300, lambda u=u: win.handle_url(u))
    code = app.exec()
    lock.unlock()
    sys.exit(code)


if __name__ == "__main__":
    main()
