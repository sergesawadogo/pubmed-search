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
APP_VERSION = "1.3.0"


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
                            QLockFile, QFileSystemWatcher, QEvent, QObject)
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
                               QAbstractItemView, QColorDialog, QComboBox, QDialog,
                               QDialogButtonBox, QTreeWidget, QTreeWidgetItem)

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
    "log_pdf": "#5FD38D",    # journal : PDF obtenu (vert)
    "log_nopdf": "#E36BD8",  # journal : sans PDF (magenta)
    "log_json": "#5CC8E8",   # journal : JSON obtenu (cyan)
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
    """Barre segmentée : PDF obtenus (vert) / traités sans PDF (magenta) / restants."""

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
            p.fillRect(0, 0, int(w_ok), r.height(), QColor(C["log_pdf"]))
            p.fillRect(int(w_ok), 0, int(w_miss) + 1, r.height(), QColor(C["log_nopdf"]))
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


def bundled_filters_path():
    return os.path.join(RES, "filtres.json")


def validate_filters(data):
    """Vérifie la structure d'un fichier de filtres ; renvoie la liste des groupes ou lève ValueError."""
    if not isinstance(data, dict) or not isinstance(data.get("groupes"), list):
        raise ValueError("le fichier doit contenir une liste \"groupes\"")
    groups = []
    for i, g in enumerate(data["groupes"], 1):
        if not isinstance(g, dict) or not g.get("titre") or not isinstance(g.get("elements"), list):
            raise ValueError(f"groupe n° {i} : il faut un \"titre\" et une liste \"elements\"")
        for j, el in enumerate(g["elements"], 1):
            if not isinstance(el, dict) or not el.get("label"):
                raise ValueError(f"« {g['titre']} », élément n° {j} : \"label\" manquant")
            if not el.get("tout") and not el.get("query"):
                raise ValueError(f"« {g['titre']} » → « {el['label']} » : \"query\" manquante")
            q = el.get("query", "")
            if q.count("(") != q.count(")"):
                raise ValueError(f"« {el['label']} » : parenthèses non équilibrées")
            if q.count('"') % 2:
                raise ValueError(f"« {el['label']} » : guillemets non appariés")
        if g["elements"]:
            groups.append(g)
    return groups


def load_filters():
    """Filtres de l'onglet Recherche : copie de l'utilisateur si elle existe (complétée par les
    groupes ajoutés dans une version plus récente de l'application), sinon ceux fournis."""
    bundled = []
    try:
        with open(bundled_filters_path(), encoding="utf-8") as f:
            bundled = validate_filters(json.load(f))
    except (OSError, ValueError):
        pass
    p = user_filters_path()
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        groups = validate_filters(data)
        known = {g["titre"] for g in groups} | set(data.get("groupes_retires", []))
        added = [g for g in bundled if g["titre"] not in known]
        if added:   # nouveaux groupes fournis (ex. Burkina Faso en 1.3.0) : ajoutés à la copie
            data["groupes"] = data["groupes"] + added
            try:
                with open(p + ".tmp", "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                os.replace(p + ".tmp", p)
            except OSError:
                pass
            groups += added
        if groups:
            return groups, p
    except (OSError, ValueError):
        pass
    return bundled, bundled_filters_path() if bundled else ""


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


# sources optionnelles du script (--sources) : clé, libellé, info-bulle
SOURCES = (
    ("elsevier", "Elsevier", "API Elsevier (ScienceDirect, Cell, JBC) : PDF puis JSON. Clé requise."),
    ("springer-oa", "Springer Nature Open Access", "Texte intégral JATS → JSON (BMC, SpringerOpen, "
                                                   "Nature Communications…). Clé requise."),
    ("springer-meta", "Springer Nature Meta", "Lien PDF officiel des articles Springer libres ; "
                                              "résumé en dernier recours pour le JSON. Clé requise."),
    ("unpaywall", "Unpaywall", "Versions libres chez l'éditeur, en dépôt ou en preprint (DOI)."),
    ("core", "CORE", "Copies des dépôts d'universités. Clé requise ; lent (~10 requêtes/min)."),
    ("openalex", "OpenAlex", "Liens PDF libres connus d'OpenAlex. Clé requise."),
)


# fonds proposés pour la zone Journal : clé -> (libellé, couleur)
LOG_THEMES = {
    "nuit": ("Nuit violette (défaut)", "#111222"),
    "noir": ("Noir", "#000000"),
    "ardoise": ("Ardoise", "#1E2329"),
    "bleu": ("Bleu nuit", "#0B1E3A"),
    "vert": ("Vert sombre", "#0E2A1F"),
    "clair": ("Clair (papier)", "#FBFAF5"),
    "blanc": ("Blanc", "#FFFFFF"),
}


class FilterEditor(QDialog):
    """Éditeur des filtres intégré : volets et cases dans une liste, libellé et requête à droite.
    Vérifie la syntaxe (parenthèses, guillemets) avant d'enregistrer la copie de l'utilisateur."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Modifier les filtres")
        self.resize(980, 640)
        self._cur = None
        src = user_filters_path() if os.path.exists(user_filters_path()) else bundled_filters_path()
        try:
            with open(src, encoding="utf-8") as f:
                self.data = json.load(f)
        except (OSError, ValueError):
            with open(bundled_filters_path(), encoding="utf-8") as f:
                self.data = json.load(f)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(10)
        lay.addWidget(hint("Chaque volet regroupe des cases ; les cases cochées d'un volet sont combinées "
                           "par OR, les volets entre eux et avec la requête par AND. Sélectionnez un volet "
                           "ou une case pour modifier son libellé et sa requête PubMed."))
        body = QHBoxLayout()
        body.setSpacing(14)
        left = QVBoxLayout()
        self.tree = QTreeWidget()
        self.tree.setObjectName("filterTree")
        self.tree.setHeaderHidden(True)
        self.tree.setMinimumWidth(320)
        self.tree.currentItemChanged.connect(self._select)
        left.addWidget(self.tree, 1)
        g1 = QGridLayout()
        g1.setSpacing(6)
        btns = (("Nouveau volet", self._add_group), ("Nouvelle case", self._add_element),
                ("Monter", lambda: self._move(-1)), ("Descendre", lambda: self._move(1)),
                ("Dupliquer", self._duplicate), ("Supprimer", self._delete))
        for i, (t, fn) in enumerate(btns):
            b = QPushButton(t)
            b.setObjectName("secondary")
            b.clicked.connect(fn)
            g1.addWidget(b, i // 2, i % 2)
        left.addLayout(g1)
        body.addLayout(left, 2)

        right = QVBoxLayout()
        right.setSpacing(6)
        self.kind_lab = section_title("")
        right.addWidget(self.kind_lab)
        self.name = QLineEdit()
        self.name_lab = field_label("Libellé", self.name)
        right.addWidget(self.name_lab)
        right.addWidget(self.name)
        self.is_all = QCheckBox("Case ALL : coche toutes les cases de ce volet (pas de requête propre)")
        self.is_all.toggled.connect(self._all_toggled)
        right.addWidget(self.is_all)
        self.q_lab = field_label("Requête PubMed ajoutée quand la case est cochée")
        right.addWidget(self.q_lab)
        self.q = QPlainTextEdit()
        self.q.setObjectName("query")
        QueryHighlighter(self.q.document())
        right.addWidget(self.q, 1)
        self.q_info = hint("")
        right.addWidget(self.q_info)
        self.q.textChanged.connect(self._check_query)
        body.addLayout(right, 3)
        lay.addLayout(body, 1)

        self.err = error_label()
        lay.addWidget(self.err)
        row = QHBoxLayout()
        reset = QPushButton("Rétablir les filtres fournis")
        reset.setObjectName("secondary")
        reset.clicked.connect(self._reset)
        ext = QPushButton("Ouvrir le fichier JSON")
        ext.setObjectName("secondary")
        ext.setToolTip("Édition avancée dans un éditeur de texte externe")
        ext.clicked.connect(self._external)
        row.addWidget(reset)
        row.addWidget(ext)
        row.addStretch(1)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Save).setText("Enregistrer")
        bb.button(QDialogButtonBox.Save).setObjectName("primary")
        bb.button(QDialogButtonBox.Cancel).setText("Annuler")
        bb.button(QDialogButtonBox.Cancel).setObjectName("secondary")
        bb.accepted.connect(self._save)
        bb.rejected.connect(self.reject)
        row.addWidget(bb)
        lay.addLayout(row)
        self._populate()

    # -- arbre ---------------------------------------------------------------------- #
    def _populate(self, select_path=None):
        self._cur = None
        self.tree.clear()
        for g in self.data.get("groupes", []):
            top = self._group_item(g)
            for el in g.get("elements", []):
                top.addChild(self._element_item(el))
            top.setExpanded(True)
        first = self.tree.topLevelItem(0)
        if first is not None:
            self.tree.setCurrentItem(first)
        else:
            self._select(None, None)

    def _group_item(self, g, index=None):
        it = QTreeWidgetItem([g.get("titre", "")])
        f = it.font(0)
        f.setBold(True)
        it.setFont(0, f)
        it.setData(0, Qt.UserRole, {"titre": g.get("titre", ""), "_group": True,
                                    **{k: v for k, v in g.items() if k not in ("titre", "elements")}})
        if index is None:
            self.tree.addTopLevelItem(it)
        else:
            self.tree.insertTopLevelItem(index, it)
        return it

    @staticmethod
    def _element_item(el):
        it = QTreeWidgetItem([el.get("label", "")])
        it.setData(0, Qt.UserRole, dict(el))
        if el.get("tout"):
            f = it.font(0)
            f.setItalic(True)
            it.setFont(0, f)
        return it

    def _store(self):
        """Recopie le formulaire dans l'élément sélectionné."""
        it = self._cur
        if it is None:
            return
        d = dict(it.data(0, Qt.UserRole) or {})
        name = self.name.text().strip()
        if d.get("_group"):
            d["titre"] = name
        else:
            d["label"] = name
            if self.is_all.isChecked():
                d["tout"] = True
                d.pop("query", None)
            else:
                d.pop("tout", None)
                d["query"] = " ".join(self.q.toPlainText().split())
        it.setData(0, Qt.UserRole, d)
        it.setText(0, name)

    def _select(self, cur, _prev=None):
        self._store()
        self._cur = cur
        d = cur.data(0, Qt.UserRole) if cur is not None else None
        on = d is not None
        for w in (self.name, self.name_lab):
            w.setEnabled(on)
        if not on:
            self.kind_lab.setText("Aucun élément")
            self.name.clear()
            for w in (self.is_all, self.q_lab, self.q, self.q_info):
                w.hide()
            return
        group = bool(d.get("_group"))
        self.kind_lab.setText("Volet" if group else "Case")
        self.name_lab.setText("Titre du volet" if group else "Libellé de la case")
        self.name.setText(d.get("titre" if group else "label", ""))
        self.is_all.blockSignals(True)
        self.is_all.setChecked(bool(d.get("tout")))
        self.is_all.blockSignals(False)
        self.is_all.setVisible(not group)
        self.q.blockSignals(True)
        self.q.setPlainText("" if group else d.get("query", ""))
        self.q.blockSignals(False)
        show_q = not group and not d.get("tout")
        for w in (self.q_lab, self.q, self.q_info):
            w.setVisible(show_q)
        self._check_query()

    def _all_toggled(self, on):
        for w in (self.q_lab, self.q, self.q_info):
            w.setVisible(not on)

    def _check_query(self):
        q = self.q.toPlainText()
        probs = []
        if q.count("(") != q.count(")"):
            probs.append("parenthèses non équilibrées")
        if q.count('"') % 2:
            probs.append("guillemets non appariés")
        if re.search(r"\b(and|or|not)\b", q) and not re.search(r"\b(AND|OR|NOT)\b", q):
            probs.append("opérateurs à écrire en majuscules (AND, OR, NOT)")
        n = q.count(" OR ") + q.count(" AND ") + 1 if q.strip() else 0
        self.q_info.setText(("⚠ " + " ; ".join(probs) + ". ") if probs else
                            f"{len(q)} caractères, environ {n} termes.")

    def _target_group(self):
        it = self.tree.currentItem()
        if it is None:
            return None
        return it if it.parent() is None else it.parent()

    def _add_group(self):
        self._store()
        it = self._group_item({"titre": "Nouveau volet"})
        it.addChild(self._element_item({"label": "Nouvelle case", "query": ""}))
        it.setExpanded(True)
        self.tree.setCurrentItem(it)
        self.name.setFocus()
        self.name.selectAll()

    def _add_element(self):
        self._store()
        grp = self._target_group()
        if grp is None:
            return self._add_group()
        cur = self.tree.currentItem()
        idx = grp.indexOfChild(cur) + 1 if cur is not grp else grp.childCount()
        it = self._element_item({"label": "Nouvelle case", "query": ""})
        grp.insertChild(idx, it)
        grp.setExpanded(True)
        self.tree.setCurrentItem(it)
        self.name.setFocus()
        self.name.selectAll()

    def _duplicate(self):
        self._store()
        it = self.tree.currentItem()
        if it is None or it.parent() is None:
            return
        d = dict(it.data(0, Qt.UserRole))
        d["label"] = d.get("label", "") + " (copie)"
        new = self._element_item(d)
        it.parent().insertChild(it.parent().indexOfChild(it) + 1, new)
        self.tree.setCurrentItem(new)

    def _move(self, step):
        self._store()
        it = self.tree.currentItem()
        if it is None:
            return
        par = it.parent()
        if par is None:
            i = self.tree.indexOfTopLevelItem(it)
            j = i + step
            if not 0 <= j < self.tree.topLevelItemCount():
                return
            self._cur = None
            self.tree.takeTopLevelItem(i)
            self.tree.insertTopLevelItem(j, it)
            it.setExpanded(True)
        else:
            i = par.indexOfChild(it)
            j = i + step
            if not 0 <= j < par.childCount():
                return
            self._cur = None
            par.takeChild(i)
            par.insertChild(j, it)
        self.tree.setCurrentItem(it)

    def _delete(self):
        it = self.tree.currentItem()
        if it is None:
            return
        d = it.data(0, Qt.UserRole) or {}
        what = f"le volet « {d.get('titre')} » et toutes ses cases" if it.parent() is None \
            else f"la case « {d.get('label')} »"
        if QMessageBox.question(self, "Supprimer", f"Supprimer {what} ?") != QMessageBox.Yes:
            return
        self._cur = None
        if it.parent() is None:
            self.tree.takeTopLevelItem(self.tree.indexOfTopLevelItem(it))
        else:
            it.parent().removeChild(it)

    def _collect(self):
        self._store()
        groups = []
        for i in range(self.tree.topLevelItemCount()):
            top = self.tree.topLevelItem(i)
            g = {k: v for k, v in (top.data(0, Qt.UserRole) or {}).items() if k != "_group"}
            g["elements"] = [dict(top.child(j).data(0, Qt.UserRole)) for j in range(top.childCount())]
            groups.append(g)
        data = {k: v for k, v in self.data.items() if k != "groupes"}
        data["groupes"] = groups
        return data

    # -- actions -------------------------------------------------------------------- #
    def _reset(self):
        r = QMessageBox.question(self, "Filtres", "Revenir aux filtres fournis avec l'application ? "
                                                  "Vos volets et cases modifiés seront remplacés.")
        if r == QMessageBox.Yes:
            with open(bundled_filters_path(), encoding="utf-8") as f:
                self.data = json.load(f)
            self._populate()

    def _external(self):
        if self._save(close=False):
            QDesktopServices.openUrl(QUrl.fromLocalFile(user_filters_path()))
            QMessageBox.information(self, "Filtres", "Le fichier s'ouvre dans votre éditeur de texte. "
                                                     "Après l'avoir enregistré, rouvrez « Modifier les "
                                                     "filtres… » ou redémarrez l'application.")

    def _save(self, close=True):
        data = self._collect()
        try:
            groups = validate_filters(data)
        except ValueError as e:
            self.err.setText(f"Filtres non valides : {e}")
            self.err.show()
            return False
        self.err.hide()
        try:
            with open(bundled_filters_path(), encoding="utf-8") as f:
                bundled_titles = {g["titre"] for g in json.load(f).get("groupes", [])}
        except (OSError, ValueError):
            bundled_titles = set()
        # volets fournis volontairement supprimés : ne pas les rajouter au prochain démarrage
        data["groupes_retires"] = sorted(bundled_titles - {g["titre"] for g in groups})
        try:
            p = user_filters_path()
            with open(p + ".tmp", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(p + ".tmp", p)
        except OSError as e:
            self.err.setText(f"Enregistrement impossible : {e}")
            self.err.show()
            return False
        self.data = data
        if close:
            self.accept()
        return True


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
        self._json_ok = 0
        self._log_lines = []
        self.log_light = False

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
        self.filter_boxes = []      # (index du groupe, label, requête, case)
        self.filter_all = {}        # index du groupe -> case ALL
        self.filter_sections = []
        self.filters_host = QWidget()
        self.filters_lay = QVBoxLayout(self.filters_host)
        self.filters_lay.setContentsMargins(0, 0, 0, 0)
        self.filters_lay.setSpacing(6)
        lay.addWidget(self.filters_host)
        self._build_filters()
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
        self.getjson = QCheckBox("Télécharger les JSON (articles sans PDF)")
        self.getjson.setToolTip("Passe supplémentaire après les PDF : texte intégral structuré (JSON) "
                                "depuis NCBI BioC, Europe PMC, Springer Nature et Elsevier, pour les "
                                "articles restés sans PDF. Le PDF reste toujours prioritaire.")
        prow = QHBoxLayout()
        prow.setSpacing(24)
        prow.addWidget(self.getpdf)
        prow.addWidget(self.getjson)
        prow.addStretch(1)
        lay.addLayout(prow)
        lay.addSpacing(8)

        # Sources optionnelles (menu déroulant à cases)
        self.sources_sec = Disclosure("Sources utilisées")
        sgrid = QGridLayout()
        sgrid.setHorizontalSpacing(18)
        sgrid.setVerticalSpacing(4)
        self.source_boxes = {}
        for i, (key, label, tip) in enumerate(SOURCES):
            cb = QCheckBox(label)
            cb.setChecked(True)
            cb.setToolTip(tip)
            cb.toggled.connect(self._sources_changed)
            sgrid.addWidget(cb, i // 2, i % 2)
            self.source_boxes[key] = cb
        self.sources_sec.body_lay.addLayout(sgrid)
        self.sources_sec.body_lay.addWidget(hint(
            "PMC S3, Europe PMC, HAL et les sites des éditeurs restent toujours utilisés. Une source "
            "cochée mais sans clé dans Réglages (Elsevier, Springer Nature, CORE, OpenAlex) n'est pas "
            "interrogée."))
        lay.addWidget(self.sources_sec)
        self._sources_changed()
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
        self._set_filter_labels([])
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
    def _build_filters(self, keep_labels=None):
        """(Re)construit les volets de filtres depuis filtres.json, sans redémarrer."""
        while self.filters_lay.count():
            it = self.filters_lay.takeAt(0)
            if it.widget():
                it.widget().deleteLater()
        self.filter_groups, self.filters_file = load_filters()
        self.filter_boxes, self.filter_sections, self.filter_all = [], [], {}
        keep = set(keep_labels or [])
        for gi, g in enumerate(self.filter_groups):
            sec = Disclosure(g["titre"])
            grid = QGridLayout()
            grid.setHorizontalSpacing(18)
            grid.setVerticalSpacing(4)
            # libellés longs : une case par ligne (sinon tronqués dans le panneau)
            ncol = 1 if any(len(el.get("label", "")) > 26 for el in g["elements"] if not el.get("tout")) else 2
            i = 0
            for el in g["elements"]:
                cb = QCheckBox(el.get("label", "?"))
                if el.get("tout"):
                    # case ALL : seule sur sa ligne, coche toutes les cases du volet
                    cb.setToolTip("Toutes les recherches de ce volet, combinées par OR")
                    f = cb.font()
                    f.setBold(True)
                    cb.setFont(f)
                    if i % ncol:
                        i += ncol - i % ncol
                    grid.addWidget(cb, i // ncol, 0, 1, 2)
                    i += ncol
                    self.filter_all[gi] = cb
                    cb.toggled.connect(lambda on, g=gi: self._toggle_all(g, on))
                    continue
                cb.setToolTip(el.get("query", ""))
                cb.toggled.connect(lambda _on, g=gi: self._member_toggled(g))
                grid.addWidget(cb, i // ncol, i % ncol)
                i += 1
                self.filter_boxes.append((gi, el.get("label", ""), el.get("query", ""), cb))
            sec.body_lay.addLayout(grid)
            self.filter_sections.append(sec)
            self.filters_lay.addWidget(sec)
        if self.filter_groups:
            self.filters_lay.addWidget(hint(
                "Cases d'un même volet combinées par OR ; volets et requête combinés par AND. "
                "Survolez une case pour voir les termes ajoutés."))
        for _, label, _, cb in self.filter_boxes:
            if label in keep:
                cb.setChecked(True)
        if hasattr(self, "final_query"):
            self._filters_changed()

    def _toggle_all(self, gi, on):
        boxes = [cb for g, _, _, cb in self.filter_boxes if g == gi]
        if not on and not all(cb.isChecked() for cb in boxes):
            return   # décochée parce qu'une case du volet a été décochée
        for cb in boxes:
            cb.blockSignals(True)
            cb.setChecked(on)
            cb.blockSignals(False)
        self._filters_changed()

    def _member_toggled(self, gi):
        allcb = self.filter_all.get(gi)
        if allcb is not None:
            full = all(cb.isChecked() for g, _, _, cb in self.filter_boxes if g == gi)
            if allcb.isChecked() != full:
                allcb.blockSignals(True)
                allcb.setChecked(full)
                allcb.blockSignals(False)
        self._filters_changed()

    def _set_filter_labels(self, labels):
        labels = set(labels or [])
        for _, label, _, cb in self.filter_boxes:
            cb.blockSignals(True)
            cb.setChecked(label in labels)
            cb.blockSignals(False)
        for gi in self.filter_all:
            self._member_toggled(gi)
        self._filters_changed()

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

    def _selected_sources(self):
        return [k for k, cb in self.source_boxes.items() if cb.isChecked()]

    def _sources_changed(self, *_):
        sel = self._selected_sources()
        n_off = len(self.source_boxes) - len(sel)
        self.sources_sec.head.setText("Sources utilisées" + (
            "  (toutes)" if not n_off else f"  ({len(sel)} sur {len(self.source_boxes)})"))

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
                     + ("" if job["getpdf"] else " / sans PDF") + (" / JSON" if job.get("getjson") else ""),
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
        self._set_filter_labels(job.get("filters", []))
        self.excel.setText(job["stem"])
        self.outdir.setText(job["outdir"])
        self.max_results.setValue(job["max"])
        self.passes.setValue(job["passes"])
        self.getpdf.setChecked(job["getpdf"])
        self.getjson.setChecked(bool(job.get("getjson")))
        src = job.get("sources")
        for k, cb in self.source_boxes.items():
            cb.setChecked(src is None or k in src)
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
        self.k_els_tok = SecretEdit("Facultatif : fourni par la bibliothèque d'une institution abonnée")
        self.k_oa = SecretEdit("Facultative")
        self.k_sn_oa = SecretEdit("Facultative : texte intégral JSON (BMC, SpringerOpen…)")
        self.k_sn_meta = SecretEdit("Facultative : liens PDF Springer, résumés")
        for lab, wdg in (("Clé API NCBI (PubMed)", self.k_ncbi), ("Clé API CORE", self.k_core),
                         ("Clé API Elsevier", self.k_els),
                         ("Jeton institutionnel Elsevier (insttoken)", self.k_els_tok),
                         ("Clé API OpenAlex", self.k_oa),
                         ("Clé Springer Nature Open Access API", self.k_sn_oa),
                         ("Clé Springer Nature Meta API", self.k_sn_meta)):
            lay.addWidget(field_label(lab, wdg.edit))
            lay.addWidget(wdg)
            lay.addSpacing(4)
        lay.addWidget(hint("Elsevier ne donne le texte intégral par API qu'aux requêtes venant du réseau "
                           "d'une institution abonnée, ou munies de son jeton institutionnel. "
                           "Clés Springer Nature : dev.springernature.com (forfait gratuit, quota "
                           "journalier). Si une seule clé couvre les deux API, saisissez-la dans les deux champs."))
        lay.addSpacing(4)
        self.remember_keys = QCheckBox("Mémoriser les clés API sur cet ordinateur")
        lay.addWidget(self.remember_keys)
        lay.addWidget(hint(f"Les clés sont alors enregistrées en clair dans {self.settings_path}. "
                           "Elles sont transmises au script par variables d'environnement et "
                           "n'apparaissent ni dans le journal ni dans error.txt."))
        lay.addSpacing(22)

        lay.addWidget(section_title("Filtres de recherche"))
        lay.addWidget(hint("Régions, Burkina Faso, niveaux de développement et types d'articles proposés dans "
                           "l'onglet Recherche sont modifiables (libellés et requêtes PubMed). Les changements "
                           "s'appliquent dès l'enregistrement."))
        frow = QHBoxLayout()
        fb = QPushButton("Modifier les filtres…")
        fb.setObjectName("secondary")
        fb.clicked.connect(self._edit_filters)
        frow.addWidget(fb)
        frow.addStretch(1)
        lay.addLayout(frow)
        lay.addSpacing(22)

        lay.addWidget(section_title("Apparence du journal"))
        crow = QHBoxLayout()
        crow.setSpacing(6)
        self.log_theme = QComboBox()
        for key, (label, _bg) in LOG_THEMES.items():
            self.log_theme.addItem(label, key)
        self.log_theme.addItem("Couleur personnalisée…", "custom")
        self.log_theme.activated.connect(self._log_theme_chosen)
        self.log_swatch = QLabel()
        self.log_swatch.setFixedSize(28, 28)
        crow.addWidget(self.log_theme, 1)
        crow.addWidget(self.log_swatch)
        crow.addStretch(1)
        lay.addWidget(field_label("Fond de la zone Journal", self.log_theme))
        lay.addLayout(crow)
        lay.addWidget(hint("PDF obtenus en vert, articles sans PDF en magenta, JSON obtenus en cyan. Les "
                           "couleurs du texte s'adaptent à un fond clair ou foncé."))
        self.log_bg = LOG_THEMES["nuit"][1]
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
        _, labels = self._selected_filters()
        dlg = FilterEditor(self)
        if dlg.exec() == QDialog.Accepted:
            self._build_filters(keep_labels=labels)
            self.phase.setText("Filtres mis à jour.")

    # ------------------------------------------------------- couleur du journal
    def _log_theme_chosen(self, idx):
        key = self.log_theme.itemData(idx)
        if key == "custom":
            c = QColorDialog.getColor(QColor(self.log_bg), self, "Fond de la zone Journal")
            if not c.isValid():
                return
            self._apply_log_bg(c.name())
        else:
            self._apply_log_bg(LOG_THEMES[key][1])
        self._save_settings()

    def _apply_log_bg(self, color):
        self.log_bg = QColor(color).name() if QColor(color).isValid() else LOG_THEMES["nuit"][1]
        light = QColor(self.log_bg).lightnessF() > 0.6
        self.log_light = light
        fg = "#1C1D33" if light else C["log_fg"]
        self.log.setStyleSheet(f"QPlainTextEdit#log {{ background: {self.log_bg}; color: {fg}; "
                               f"border: 1px solid {'#C9C8D8' if light else '#2A2B45'}; }}")
        self.log_swatch.setStyleSheet(f"background:{self.log_bg}; border:1px solid {C['rule']}; "
                                      "border-radius:6px;")
        idx = next((i for i in range(self.log_theme.count())
                    if self.log_theme.itemData(i) in LOG_THEMES
                    and LOG_THEMES[self.log_theme.itemData(i)][1].lower() == self.log_bg.lower()),
                   self.log_theme.count() - 1)
        self.log_theme.setCurrentIndex(idx)
        self._recolor_log()

    def _log_palette(self):
        if getattr(self, "log_light", False):   # fond clair : teintes plus foncées, lisibles
            return {"pdf": "#1B7F45", "nopdf": "#A0219A", "json": "#0B6E8F", "warn": "#9A5B00",
                    "err": "#B3261E", "dim": "#6B6C80", None: "#1C1D33"}
        return {"pdf": C["log_pdf"], "nopdf": C["log_nopdf"], "json": C["log_json"], "warn": "#E6B45C",
                "err": "#F08A82", "dim": C["log_dim"], None: C["log_fg"]}

    def _recolor_log(self):
        """Réapplique les couleurs au journal existant après un changement de fond."""
        lines = getattr(self, "_log_lines", [])
        if not lines:
            return
        self.log.clear()
        for text, kind in lines[-20000:]:
            self._append(text, kind, record=False)

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
        self.s_ok = Stat("PDF obtenus", C["log_pdf"])
        self.s_miss = Stat("sans PDF", C["log_nopdf"])
        self.s_left = Stat("restants", C["log_fg"])
        self.s_rate = Stat("réussite", C["log_fg"])
        self.s_json = Stat("JSON obtenus", C["log_json"])
        self.s_json.hide()
        for s in (self.s_ok, self.s_miss, self.s_left, self.s_rate, self.s_json):
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
                                       "--time-set", "--getjson", "--sources", "--springer-oa-key",
                                       "--springer-meta-key") if f not in flags]
                txt = f"Version {ver}, prête."
                if missing:
                    txt += " Options absentes de cette version : " + ", ".join(missing) + "."
                self.script_status.setText(txt)
            self.k_sn_oa.setEnabled("--springer-oa-key" in flags)
            self.k_sn_meta.setEnabled("--springer-meta-key" in flags)
            self.getjson.setEnabled("--getjson" in flags)
            if "--getjson" not in flags:
                self.getjson.setChecked(False)
            self.sources_sec.setEnabled("--sources" in flags)
            self.k_els_tok.setEnabled("--elsevier-insttoken" in flags)
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
        pairs = (("NCBI_API_KEY", self.k_ncbi), ("CORE_API_KEY", self.k_core),
                 ("ELSEVIER_API_KEY", self.k_els), ("OPENALEX_API_KEY", self.k_oa),
                 ("SPRINGER_OA_API_KEY", self.k_sn_oa), ("SPRINGER_META_API_KEY", self.k_sn_meta),
                 ("ELSEVIER_INSTTOKEN", self.k_els_tok))
        for var, _ in pairs:
            env.remove(var)
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
                "getjson": self.getjson.isChecked() and self.getjson.isEnabled(),
                "sources": self._selected_sources() if self.sources_sec.isEnabled() else None,
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
        if job.get("getjson"):
            args.append("--getjson")
        src = job.get("sources")
        if src is not None and set(src) != {k for k, _, _ in SOURCES}:
            args += ["--sources", ",".join(src) if src else "none"]
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
        self._log_lines = []
        self._counts = {"total": 0, "ok": 0, "miss": 0}
        self._json_phase = False
        self._json_ok = 0
        self.strip.set_values(0, 0, 0)
        for s in (self.s_ok, self.s_miss, self.s_left, self.s_rate, self.s_json):
            s.set("–")
        self.s_json.setVisible("--getjson" in args)
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
        m = re.match(r"\s*\[(\d+)/(\d+)\] PMID \d+ .*⇒ (.*)$", line)
        if m:   # passe JSON
            i, n, res = int(m.group(1)), int(m.group(2)), m.group(3)
            if res.startswith("JSON"):
                self._json_ok += 1
                self.s_json.set(self._json_ok)
                kind = "json"
            self.phase.setText(f"Passe JSON : article {i} sur {n}")
        m = re.match(r"\s*\[(\d+)/(\d+)\] PMID \d+ .*→ (.*)$", line)
        if m:
            i, n, res = int(m.group(1)), int(m.group(2)), m.group(3)
            c = self._counts
            c["total"] = max(c["total"], n)
            if res.startswith("OK"):
                c["ok"] += 1
                kind = "pdf"
            elif "non téléchargé" in res:
                c["miss"] += 1
                kind = "nopdf"
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
        m = re.match(r"\s*JSON obtenus\s*:\s*(\d+)", line)
        if m:
            self.s_json.set(m.group(1))
            self.s_json.show()
        if "🧾 Passe JSON" in line:
            self.phase.setText("Passe JSON : texte intégral des articles sans PDF…")
        if "Recherche dans PubMed" in line or "ESearch" in line:
            self.phase.setText("Recherche dans PubMed…")
        elif "Récupération des métadonnées" in line:
            self.phase.setText("Lecture des notices PubMed…")
        elif "Enrichissement" in line:
            self.phase.setText("Recherche des liens de texte intégral…")
        elif "Exécution programmée" in line:
            self.phase.setText(line.strip(" ⏳"))
        self._append(line + "\n", kind)

    def _append(self, text, kind=None, record=True):
        colors = self._log_palette()
        if record:
            if not hasattr(self, "_log_lines"):
                self._log_lines = []
            self._log_lines.append((text, kind))
            if len(self._log_lines) > 25000:
                del self._log_lines[:5000]
        cur = self.log.textCursor()
        cur.movePosition(QTextCursor.End)
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(colors.get(kind, colors[None])))
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
        self.getjson.setChecked(bool(s.get("getjson", False)))
        src = s.get("sources")
        for k, cb in self.source_boxes.items():
            cb.setChecked(not isinstance(src, list) or k in src)
        self._apply_log_bg(s.get("log_bg") or LOG_THEMES["nuit"][1])
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
        self.k_sn_oa.setText(keys.get("springer_oa", ""))
        self.k_sn_meta.setText(keys.get("springer_meta", ""))
        self.k_els_tok.setText(keys.get("elsevier_insttoken", ""))
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
            "getjson": self.getjson.isChecked(),
            "sources": self._selected_sources(),
            "log_bg": getattr(self, "log_bg", LOG_THEMES["nuit"][1]),
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
                         "elsevier": self.k_els.text(), "openalex": self.k_oa.text(),
                         "springer_oa": self.k_sn_oa.text(), "springer_meta": self.k_sn_meta.text(),
                         "elsevier_insttoken": self.k_els_tok.text()}
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
QComboBox {{ background: {C['panel']}; border: 1px solid {C['rule']}; border-radius: 6px; padding: 6px 9px;
    min-height: 22px; }}
QComboBox:focus {{ border: 2px solid {C['hema']}; padding: 5px 8px; }}
QTreeWidget#filterTree {{ background: {C['panel']}; border: 1px solid {C['rule']}; border-radius: 6px;
    padding: 4px; font-size: 13.5px; selection-background-color: {C['hema_soft']}; selection-color: {C['ink']}; }}
QTreeWidget#filterTree::item {{ padding: 3px 2px; }}
QDialog {{ background: {C['slide']}; }}
QPlainTextEdit#filterJson {{ background: {C['panel']}; border: 1px solid {C['rule']}; border-radius: 6px;
    font-family: "{mono}"; font-size: 12.5px; padding: 8px; }}
QCheckBox:disabled, QToolButton#disclosure:disabled {{ color: #A3A4B5; }}
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


class MacUrlFilter(QObject):
    """macOS : les liens pubmedsearch:// arrivent par un événement « ouvrir l'URL » de
    l'application (Apple Event), pas par la ligne de commande."""

    def __init__(self):
        super().__init__()
        self.win = None
        self.pending = []

    def eventFilter(self, obj, event):
        if event.type() == QEvent.FileOpen:
            url = event.url().toString()
            if url.lower().startswith(URL_SCHEME + ":"):
                if self.win is not None:
                    self.win.handle_url(url)
                else:
                    self.pending.append(url)
                return True
        return False


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
    url_filter = MacUrlFilter()
    app.installEventFilter(url_filter)
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
    url_filter.win = win
    urls += url_filter.pending
    for u in urls:
        QTimer.singleShot(300, lambda u=u: win.handle_url(u))
    code = app.exec()
    lock.unlock()
    sys.exit(code)


if __name__ == "__main__":
    main()
