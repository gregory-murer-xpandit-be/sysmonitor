"""
SysMonitor v0.1 - Dashboard système Windows 11
Dépendances : pip install PyQt6 pyqtgraph psutil pywin32
Exécuter en tant qu'administrateur pour accès complet aux connexions réseau.
"""

import sys
import datetime
import csv
import os
import winreg
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed

import psutil
import win32api
import win32con

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QComboBox, QLineEdit, QGroupBox,
    QFrame, QFileDialog, QMessageBox, QProgressBar
)
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt6.QtGui import QColor

import pyqtgraph as pg

# ─── Thème sombre ──────────────────────────────────────────────────────────────
BG_DARK   = "#0f1117"
BG_PANEL  = "#1a1d27"
BG_CARD   = "#20232f"
BG_NODATE = "#1e2030"
ACCENT    = "#4f8ef7"
ACCENT2   = "#2dd4bf"
ACCENT3   = "#f59e0b"
RED       = "#ef4444"
GREEN     = "#22c55e"
ORANGE    = "#f59e0b"
TEXT_PRI  = "#e2e8f0"
TEXT_SEC  = "#94a3b8"
TEXT_MUTED= "#475569"
BORDER    = "#2d3148"

STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {BG_DARK};
    color: {TEXT_PRI};
    font-family: 'Segoe UI';
    font-size: 13px;
}}
QTabWidget::pane {{
    border: 1px solid {BORDER};
    background: {BG_PANEL};
    border-radius: 8px;
}}
QTabBar::tab {{
    background: {BG_DARK};
    color: {TEXT_SEC};
    padding: 10px 24px;
    border: 1px solid {BORDER};
    border-bottom: none;
    border-radius: 6px 6px 0 0;
    margin-right: 2px;
    font-size: 13px;
}}
QTabBar::tab:selected {{
    background: {BG_PANEL};
    color: {ACCENT};
    border-bottom: 2px solid {ACCENT};
    font-weight: 600;
}}
QTableWidget {{
    background: {BG_CARD};
    border: 1px solid {BORDER};
    border-radius: 6px;
    gridline-color: {BORDER};
    color: {TEXT_PRI};
    selection-background-color: {ACCENT};
}}
QTableWidget::item {{ padding: 6px 10px; }}
QHeaderView::section {{
    background: {BG_PANEL};
    color: {ACCENT};
    padding: 8px 10px;
    border: none;
    border-right: 1px solid {BORDER};
    font-weight: 600;
}}
QPushButton {{
    background: {ACCENT};
    color: white;
    border: none;
    border-radius: 6px;
    padding: 8px 18px;
    font-size: 13px;
    font-weight: 600;
}}
QPushButton:hover {{ background: #6ba3f9; }}
QPushButton:pressed {{ background: #3a7bf5; }}
QComboBox {{
    background: {BG_CARD};
    color: {TEXT_PRI};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 6px 12px;
    min-width: 160px;
}}
QComboBox::drop-down {{ border: none; }}
QComboBox QAbstractItemView {{
    background: {BG_CARD};
    color: {TEXT_PRI};
    selection-background-color: {ACCENT};
}}
QLineEdit {{
    background: {BG_CARD};
    color: {TEXT_PRI};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 6px 12px;
}}
QLineEdit:focus {{ border-color: {ACCENT}; }}
QScrollBar:vertical {{
    background: {BG_DARK};
    width: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {BORDER};
    border-radius: 4px;
    min-height: 20px;
}}
QGroupBox {{
    border: 1px solid {BORDER};
    border-radius: 8px;
    margin-top: 12px;
    padding: 12px;
    color: {TEXT_SEC};
    font-size: 12px;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 12px;
    color: {ACCENT};
    font-weight: 600;
}}
QProgressBar {{
    background: {BG_CARD};
    border: 1px solid {BORDER};
    border-radius: 4px;
    height: 8px;
    text-align: center;
    color: {TEXT_PRI};
}}
QProgressBar::chunk {{
    background: {ACCENT};
    border-radius: 4px;
}}
"""

# ─── Graphe pyqtgraph commun ────────────────────────────────────────────────
def make_plot(title="", y_label="", color=ACCENT):
    pg.setConfigOptions(antialias=True, background=BG_CARD, foreground=TEXT_SEC)
    plot = pg.PlotWidget(title=title)
    plot.setLabel('left', y_label)
    plot.showGrid(x=True, y=True, alpha=0.15)
    plot.getAxis('bottom').setStyle(showValues=False)
    plot.setMinimumHeight(140)
    return plot


# ═══════════════════════════════════════════════════════════════════════════════
# ONGLET RÉSEAU
# ═══════════════════════════════════════════════════════════════════════════════
class NetworkWorker(QThread):
    data_ready = pyqtSignal(dict)

    def __init__(self):
        super().__init__()
        self._running = True
        self._prev_bytes = psutil.net_io_counters()

    def run(self):
        while self._running:
            try:
                curr = psutil.net_io_counters()
                prev = self._prev_bytes
                conns = psutil.net_connections(kind='inet')

                http = https = other = 0
                conn_list = []
                for c in conns:
                    if c.status == 'ESTABLISHED':
                        raddr = c.raddr
                        if raddr:
                            port = raddr.port
                            if port == 80:    http += 1
                            elif port == 443: https += 1
                            else:             other += 1
                            try:
                                proc = psutil.Process(c.pid).name() if c.pid else "—"
                            except Exception:
                                proc = "—"
                            conn_list.append({
                                "laddr":  f"{c.laddr.ip}:{c.laddr.port}",
                                "raddr":  f"{raddr.ip}:{raddr.port}",
                                "port":   port,
                                "pid":    c.pid or 0,
                                "proc":   proc,
                                "proto":  "HTTPS" if port == 443 else ("HTTP" if port == 80 else f":{port}"),
                                "status": c.status,
                            })

                self._prev_bytes = curr
                self.data_ready.emit({
                    "sent_kbps":  (curr.bytes_sent - prev.bytes_sent) / 1024,
                    "recv_kbps":  (curr.bytes_recv - prev.bytes_recv) / 1024,
                    "total_conn": len([c for c in conns if c.status == 'ESTABLISHED']),
                    "http": http, "https": https, "other": other,
                    "conns": conn_list[:200],
                })
            except Exception as e:
                print(f"[Network] {e}")
            self.msleep(1000)

    def stop(self):
        self._running = False


class NetworkTab(QWidget):
    def __init__(self):
        super().__init__()
        self.history_len = 60
        self.sent_hist = deque([0.0] * self.history_len, maxlen=self.history_len)
        self.recv_hist = deque([0.0] * self.history_len, maxlen=self.history_len)
        self._build_ui()
        self.worker = NetworkWorker()
        self.worker.data_ready.connect(self._on_data)
        self.worker.start()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        kpi_row = QHBoxLayout()
        self.kpi_conn  = self._kpi_card("Connexions actives", "0", ACCENT)
        self.kpi_http  = self._kpi_card("HTTP (80)", "0", ACCENT2)
        self.kpi_https = self._kpi_card("HTTPS (443)", "0", GREEN)
        self.kpi_other = self._kpi_card("Autres ports", "0", ACCENT3)
        self.kpi_send  = self._kpi_card("↑ Envoi", "0 KB/s", RED)
        self.kpi_recv  = self._kpi_card("↓ Réception", "0 KB/s", ACCENT)
        for card in [self.kpi_conn, self.kpi_http, self.kpi_https,
                     self.kpi_other, self.kpi_send, self.kpi_recv]:
            kpi_row.addWidget(card)
        layout.addLayout(kpi_row)

        gbox = QGroupBox("Bande passante en temps réel (KB/s)")
        gbox_layout = QVBoxLayout(gbox)
        self.bw_plot = make_plot(y_label="KB/s")
        self.curve_sent = self.bw_plot.plot(pen=pg.mkPen(RED, width=2), name="↑ Envoi")
        self.curve_recv = self.bw_plot.plot(pen=pg.mkPen(ACCENT, width=2), name="↓ Réception")
        self.bw_plot.addLegend(offset=(10, 10))
        gbox_layout.addWidget(self.bw_plot)
        layout.addWidget(gbox)

        proto_row = QHBoxLayout()
        pb_box = QGroupBox("Répartition protocoles")
        pb_layout = QVBoxLayout(pb_box)
        self.bar_http  = self._proto_bar("HTTP",   ACCENT2)
        self.bar_https = self._proto_bar("HTTPS",  GREEN)
        self.bar_other = self._proto_bar("Autres", ACCENT3)
        for w in [self.bar_http, self.bar_https, self.bar_other]:
            pb_layout.addLayout(w[0])
        proto_row.addWidget(pb_box)

        pie_box = QGroupBox("Vue donut (connexions)")
        pie_layout = QVBoxLayout(pie_box)
        self.pie_plot = pg.PlotWidget()
        self.pie_plot.setMaximumHeight(160)
        self.pie_plot.hideAxis('bottom')
        self.pie_plot.hideAxis('left')
        self.pie_plot.setBackground(BG_CARD)
        self.pie_bars = pg.BarGraphItem(x=[1, 2, 3], height=[1, 1, 1], width=0.6,
                                        brushes=[ACCENT2, GREEN, ACCENT3])
        self.pie_plot.addItem(self.pie_bars)
        pie_layout.addWidget(self.pie_plot)
        proto_row.addWidget(pie_box)
        layout.addLayout(proto_row)

        tbox = QGroupBox("Connexions établies")
        tbox_layout = QVBoxLayout(tbox)
        filter_row = QHBoxLayout()
        self.conn_filter = QLineEdit()
        self.conn_filter.setPlaceholderText("Filtrer par IP, port, processus…")
        self.conn_filter.textChanged.connect(self._filter_conns)
        filter_row.addWidget(self.conn_filter)
        tbox_layout.addLayout(filter_row)

        self.conn_table = QTableWidget(0, 5)
        self.conn_table.setHorizontalHeaderLabels(
            ["Adresse locale", "Adresse distante", "Protocole", "PID", "Processus"])
        self.conn_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.conn_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.conn_table.setSortingEnabled(True)
        tbox_layout.addWidget(self.conn_table)
        layout.addWidget(tbox)
        self._all_conns = []

    def _kpi_card(self, label, value, color):
        frame = QFrame()
        frame.setStyleSheet(f"QFrame {{ background:{BG_CARD}; border:1px solid {BORDER}; border-radius:8px; }}")
        v = QVBoxLayout(frame)
        v.setContentsMargins(14, 10, 14, 10)
        lbl = QLabel(label)
        lbl.setStyleSheet(f"color:{TEXT_SEC}; font-size:11px; border:none;")
        val = QLabel(value)
        val.setStyleSheet(f"color:{color}; font-size:22px; font-weight:700; border:none;")
        v.addWidget(lbl)
        v.addWidget(val)
        frame._value_label = val
        return frame

    def _proto_bar(self, name, color):
        row = QHBoxLayout()
        lbl = QLabel(f"{name}:")
        lbl.setFixedWidth(55)
        lbl.setStyleSheet(f"color:{TEXT_SEC}; border:none;")
        bar = QProgressBar()
        bar.setRange(0, 100)
        bar.setValue(0)
        bar.setStyleSheet(f"QProgressBar::chunk {{ background:{color}; border-radius:4px; }}")
        cnt = QLabel("0")
        cnt.setFixedWidth(30)
        cnt.setStyleSheet(f"color:{TEXT_PRI}; border:none;")
        row.addWidget(lbl)
        row.addWidget(bar)
        row.addWidget(cnt)
        return row, bar, cnt

    def _on_data(self, data):
        self.sent_hist.append(data["sent_kbps"])
        self.recv_hist.append(data["recv_kbps"])
        self.curve_sent.setData(list(self.sent_hist))
        self.curve_recv.setData(list(self.recv_hist))
        self.kpi_conn._value_label.setText(str(data["total_conn"]))
        self.kpi_http._value_label.setText(str(data["http"]))
        self.kpi_https._value_label.setText(str(data["https"]))
        self.kpi_other._value_label.setText(str(data["other"]))
        self.kpi_send._value_label.setText(f"{data['sent_kbps']:.1f} KB/s")
        self.kpi_recv._value_label.setText(f"{data['recv_kbps']:.1f} KB/s")
        total = max(data["http"] + data["https"] + data["other"], 1)
        self.bar_http[1].setValue(int(data["http"] / total * 100))
        self.bar_http[2].setText(str(data["http"]))
        self.bar_https[1].setValue(int(data["https"] / total * 100))
        self.bar_https[2].setText(str(data["https"]))
        self.bar_other[1].setValue(int(data["other"] / total * 100))
        self.bar_other[2].setText(str(data["other"]))
        self.pie_bars.setOpts(height=[data["http"], data["https"], data["other"]])
        self._all_conns = data["conns"]
        self._populate_table(data["conns"])

    def _populate_table(self, conns):
        self.conn_table.setRowCount(len(conns))
        for r, c in enumerate(conns):
            self.conn_table.setItem(r, 0, QTableWidgetItem(c["laddr"]))
            self.conn_table.setItem(r, 1, QTableWidgetItem(c["raddr"]))
            proto_item = QTableWidgetItem(c["proto"])
            color = {"HTTPS": GREEN, "HTTP": ACCENT2}.get(c["proto"], ACCENT3)
            proto_item.setForeground(QColor(color))
            self.conn_table.setItem(r, 2, proto_item)
            self.conn_table.setItem(r, 3, QTableWidgetItem(str(c["pid"])))
            self.conn_table.setItem(r, 4, QTableWidgetItem(c["proc"]))

    def _filter_conns(self, text):
        text = text.lower()
        filtered = [c for c in self._all_conns if
                    text in c["laddr"].lower() or text in c["raddr"].lower() or
                    text in c["proc"].lower() or text in c["proto"].lower()]
        self._populate_table(filtered)

    def get_summary(self):
        return {"sent": list(self.sent_hist), "recv": list(self.recv_hist)}

    def stop(self):
        self.worker.stop()
        self.worker.wait()


# ═══════════════════════════════════════════════════════════════════════════════
# ONGLET APPLICATIONS — AppWorker via registre Windows (rapide)
# ═══════════════════════════════════════════════════════════════════════════════

_REG_PATHS = [
    (winreg.HKEY_LOCAL_MACHINE,
     r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
     winreg.KEY_READ | winreg.KEY_WOW64_64KEY),
    (winreg.HKEY_LOCAL_MACHINE,
     r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
     winreg.KEY_READ | winreg.KEY_WOW64_32KEY),
    (winreg.HKEY_CURRENT_USER,
     r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
     winreg.KEY_READ | winreg.KEY_WOW64_64KEY),
    (winreg.HKEY_CURRENT_USER,
     r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
     winreg.KEY_READ | winreg.KEY_WOW64_32KEY),
]


def _reg_value(key, name: str, default="") -> str:
    try:
        val, _ = winreg.QueryValueEx(key, name)
        return str(val).strip() if val else default
    except OSError:
        return default


def _read_registry_path(hive, path: str, flags: int) -> list:
    entries = []
    try:
        root = winreg.OpenKey(hive, path, 0, flags)
    except OSError:
        return entries
    idx = 0
    while True:
        try:
            subkey_name = winreg.EnumKey(root, idx)
        except OSError:
            break
        idx += 1
        try:
            subkey = winreg.OpenKey(root, subkey_name, 0, flags)
            name = _reg_value(subkey, "DisplayName")
            if not name:
                winreg.CloseKey(subkey)
                continue
            install_date_raw = _reg_value(subkey, "InstallDate")
            size_kb = 0
            try:
                size_kb, _ = winreg.QueryValueEx(subkey, "EstimatedSize")
            except OSError:
                pass
            entries.append({
                "name":         name,
                "vendor":       _reg_value(subkey, "Publisher"),
                "version":      _reg_value(subkey, "DisplayVersion"),
                "install_date": install_date_raw,   # brut, tel quel
                "install_loc":  _reg_value(subkey, "InstallLocation"),
                "size_kb":      int(size_kb) if isinstance(size_kb, int) else 0,
                "size_mb":      round(int(size_kb) / 1024, 1) if isinstance(size_kb, int) else 0.0,
                "mem_mb":       0.0,
                "key_id":       subkey_name,
            })
            winreg.CloseKey(subkey)
        except OSError:
            pass
    winreg.CloseKey(root)
    return entries


def _collect_process_memory() -> dict:
    mem_map = {}
    for proc in psutil.process_iter(["name", "memory_info"]):
        try:
            n = proc.info["name"]
            if not n:
                continue
            rss_mb = proc.info["memory_info"].rss / (1024 * 1024)
            key = n.lower().removesuffix(".exe")
            mem_map[key] = mem_map.get(key, 0.0) + rss_mb
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return mem_map


def _match_memory(app: dict, mem_map: dict) -> float:
    name_lower = app["name"].lower()
    first_word = name_lower.split()[0] if name_lower else ""
    if first_word in mem_map:
        return mem_map[first_word]
    for proc_name, mb in mem_map.items():
        if first_word and first_word in proc_name:
            return mb
    install_loc = app.get("install_loc", "").lower()
    if install_loc:
        for proc_name, mb in mem_map.items():
            if proc_name in install_loc:
                return mb
    return 0.0


class AppWorker(QThread):
    data_ready = pyqtSignal(list)
    progress   = pyqtSignal(int, str)

    def run(self):
        self.progress.emit(5, "Lecture du registre…")
        all_entries = []
        labels = ["HKLM x64", "HKLM x32", "HKCU x64", "HKCU x32"]
        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {
                pool.submit(_read_registry_path, hive, path, flags): label
                for (hive, path, flags), label in zip(_REG_PATHS, labels)
            }
            done = 0
            for future in as_completed(futures):
                done += 1
                label = futures[future]
                try:
                    results = future.result()
                    all_entries.extend(results)
                    self.progress.emit(10 + done * 15, f"{label} — {len(results)} entrées")
                except Exception as exc:
                    print(f"[Registry {label}] {exc}")

        self.progress.emit(75, "Déduplication…")
        seen = {}
        for entry in all_entries:
            kid = entry["key_id"]
            if kid not in seen or entry["size_kb"] > seen[kid]["size_kb"]:
                seen[kid] = entry
        apps = list(seen.values())

        self.progress.emit(85, "Collecte mémoire processus…")
        mem_map = _collect_process_memory()
        for app in apps:
            app["mem_mb"] = round(_match_memory(app, mem_map), 1)

        # Tri par défaut : avec date d'abord, sans date en bas, puis alphabétique
        apps.sort(key=lambda a: (
            a["install_date"] == "",
            a["name"].lower()
        ))

        self.progress.emit(100, f"{len(apps)} applications chargées")
        self.data_ready.emit(apps)


class AppsTab(QWidget):
    def __init__(self):
        super().__init__()
        self._apps = []
        self._build_ui()
        self._load_apps()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # ── Toolbar ──
        toolbar = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Rechercher une application…")
        self.search.textChanged.connect(self._filter)
        toolbar.addWidget(self.search)

        toolbar.addWidget(QLabel("Trier par :"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems([
            "Nom", "Mémoire active", "Taille disque", "Version"
        ])
        self.sort_combo.currentTextChanged.connect(self._sort)
        toolbar.addWidget(self.sort_combo)

        self.refresh_btn = QPushButton("Rafraîchir")
        self.refresh_btn.clicked.connect(self._load_apps)
        toolbar.addWidget(self.refresh_btn)

        self.export_btn = QPushButton("Export CSV")
        self.export_btn.clicked.connect(self._export_csv)
        toolbar.addWidget(self.export_btn)
        layout.addLayout(toolbar)

        # ── Barre de progression ──
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(6)
        layout.addWidget(self.progress_bar)

        # ── KPI ──
        kpi_row = QHBoxLayout()
        self.kpi_total  = self._mini_kpi("Applications installées", "—")
        self.kpi_active = self._mini_kpi("Actives en ce moment", "—")
        self.kpi_last   = self._mini_kpi("Dernière installation", "—")
        self.kpi_mem    = self._mini_kpi("Mémoire totale (actifs)", "—")
        for k in [self.kpi_total, self.kpi_active, self.kpi_last, self.kpi_mem]:
            kpi_row.addWidget(k)
        layout.addLayout(kpi_row)

        # ── Légende ──
        legend_row = QHBoxLayout()
        legend_row.addWidget(self._legend_dot(ORANGE,     "Application active en mémoire"))
        legend_row.addWidget(self._legend_dot(RED,        "Mémoire > 500 MB"))
        legend_row.addStretch()
        layout.addLayout(legend_row)

        # ── Tableau ──
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Nom", "Éditeur", "Version", "Date installation", "Taille disque", "Mémoire active"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSortingEnabled(True)
        layout.addWidget(self.table)

        self.status_lbl = QLabel("Chargement du registre…")
        self.status_lbl.setStyleSheet(f"color:{TEXT_SEC}; font-size:11px;")
        layout.addWidget(self.status_lbl)

    def _mini_kpi(self, label, value):
        frame = QFrame()
        frame.setStyleSheet(
            f"QFrame {{ background:{BG_CARD}; border:1px solid {BORDER}; border-radius:8px; }}")
        v = QVBoxLayout(frame)
        v.setContentsMargins(12, 8, 12, 8)
        lbl = QLabel(label)
        lbl.setStyleSheet(f"color:{TEXT_SEC}; font-size:11px; border:none;")
        val = QLabel(value)
        val.setStyleSheet(f"color:{ACCENT}; font-size:18px; font-weight:700; border:none;")
        val.setWordWrap(True)
        v.addWidget(lbl)
        v.addWidget(val)
        frame._val = val
        return frame

    def _legend_dot(self, color, label):
        frame = QFrame()
        h = QHBoxLayout(frame)
        h.setContentsMargins(0, 0, 12, 0)
        h.setSpacing(5)
        dot = QLabel("●")
        dot.setStyleSheet(f"color:{color}; font-size:14px; border:none;")
        lbl = QLabel(label)
        lbl.setStyleSheet(f"color:{TEXT_SEC}; font-size:11px; border:none;")
        h.addWidget(dot)
        h.addWidget(lbl)
        return frame

    def _load_apps(self):
        self.refresh_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_lbl.setText("Lecture du registre Windows…")
        self.worker = AppWorker()
        self.worker.data_ready.connect(self._on_apps)
        self.worker.progress.connect(self._on_progress)
        self.worker.start()

    def _on_progress(self, pct: int, msg: str):
        self.progress_bar.setValue(pct)
        self.status_lbl.setText(msg)
        if pct >= 100:
            self.progress_bar.setVisible(False)

    def _on_apps(self, apps):
        self._apps = apps
        self._populate(apps)
        self.refresh_btn.setEnabled(True)

        total_mem = sum(a.get("mem_mb", 0) for a in apps)
        active    = sum(1 for a in apps if a.get("mem_mb", 0) > 0)
        last      = next((a for a in apps if a.get("install_date")), None)

        self.kpi_total._val.setText(str(len(apps)))
        self.kpi_active._val.setText(str(active))
        self.kpi_mem._val.setText(f"{total_mem:.0f} MB")
        if last:
            self.kpi_last._val.setText(
                f"{last['install_date']}\n{last['name'][:22]}")
            self.status_lbl.setText(
                f"{len(apps)} applications — {active} actives en mémoire")
        else:
            self.kpi_last._val.setText("—")
            self.status_lbl.setText(f"{len(apps)} applications chargées")

    def _populate(self, apps):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(apps))

        for r, a in enumerate(apps):
            active = a.get("mem_mb", 0) > 0

            # Col 0 — Nom : orange si actif en mémoire
            name_item = QTableWidgetItem(a["name"])
            if active:
                name_item.setForeground(QColor(ORANGE))
            self.table.setItem(r, 0, name_item)

            # Col 1 — Éditeur
            self.table.setItem(r, 1, QTableWidgetItem(a.get("vendor", "—") or "—"))

            # Col 2 — Version
            self.table.setItem(r, 2, QTableWidgetItem(a.get("version", "—") or "—"))

            # Col 3 — Date installation : texte brut du registre, sans parsing
            date_raw = a.get("install_date", "") or "—"
            self.table.setItem(r, 3, QTableWidgetItem(date_raw))

            # Col 4 — Taille disque (depuis registre)
            size_mb = a.get("size_mb", 0)
            self.table.setItem(r, 4, QTableWidgetItem(
                f"{size_mb:.0f} MB" if size_mb > 0 else "—"))

            # Col 5 — Mémoire active (psutil)
            mem_mb = a.get("mem_mb", 0)
            mem_item = QTableWidgetItem(f"{mem_mb:.0f} MB" if mem_mb > 0 else "—")
            if mem_mb > 500:
                mem_item.setForeground(QColor(RED))
            self.table.setItem(r, 5, mem_item)

        self.table.setSortingEnabled(True)

    def _filter(self, text):
        text = text.lower()
        filtered = [a for a in self._apps if
                    text in a["name"].lower() or
                    text in (a.get("vendor") or "").lower() or
                    text in (a.get("version") or "").lower()]
        self._populate(filtered)

    def _sort(self, criterion):
        key_map = {
            "Nom":            lambda a: a["name"].lower(),
            "Mémoire active": lambda a: -a.get("mem_mb", 0),
            "Taille disque":  lambda a: -a.get("size_mb", 0),
            "Version":        lambda a: (a.get("version") or "").lower(),
        }
        fn = key_map.get(criterion, lambda a: a["name"].lower())
        self._populate(sorted(self._apps, key=fn))

    def _export_csv(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Exporter CSV", "applications.csv", "CSV (*.csv)")
        if path:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["Nom", "Éditeur", "Version", "Date installation",
                             "Taille (MB)", "Mémoire active (MB)"])
                for a in self._apps:
                    w.writerow([
                        a["name"], a.get("vendor", ""), a.get("version", ""),
                        a.get("install_date", ""),
                        a.get("size_mb", 0), a.get("mem_mb", 0),
                    ])
            QMessageBox.information(self, "Export", f"Exporté : {path}")

    def get_apps(self):
        return self._apps


# ═══════════════════════════════════════════════════════════════════════════════
# ONGLET PATCHES / WINDOWS UPDATE
# ═══════════════════════════════════════════════════════════════════════════════
class PatchWorker(QThread):
    data_ready = pyqtSignal(dict)

    def run(self):
        result = {
            "hotfixes": [], "os_caption": "—", "os_version": "—",
            "os_build": "—", "last_boot": "—",
            "pending": None, "pending_count": None,
        }
        try:
            import wmi
            c = wmi.WMI()
            for os_obj in c.Win32_OperatingSystem():
                result["os_caption"] = os_obj.Caption or "—"
                result["os_version"] = os_obj.Version or "—"
                result["os_build"]   = os_obj.BuildNumber or "—"
                try:
                    lb = os_obj.LastBootUpTime
                    if lb:
                        lb_dt = datetime.datetime.strptime(lb[:14], "%Y%m%d%H%M%S")
                        result["last_boot"] = lb_dt.strftime("%d/%m/%Y %H:%M")
                except Exception:
                    pass

            for hf in c.Win32_QuickFixEngineering():
                installed = hf.InstalledOn or ""
                try:
                    dt = datetime.datetime.strptime(installed, "%m/%d/%Y")
                    installed_fmt = dt.strftime("%d/%m/%Y")
                    dt_obj = dt
                except Exception:
                    installed_fmt = installed
                    dt_obj = None
                result["hotfixes"].append({
                    "kb": hf.HotFixID or "—", "desc": hf.Description or "—",
                    "date": installed_fmt, "date_obj": dt_obj,
                    "installed_by": hf.InstalledBy or "—",
                })
            result["hotfixes"].sort(
                key=lambda h: h["date_obj"] or datetime.datetime.min, reverse=True)
        except Exception as e:
            print(f"[Patch WMI] {e}")

        try:
            import win32com.client
            wua = win32com.client.Dispatch("Microsoft.Update.Session")
            searcher = wua.CreateUpdateSearcher()
            res_search = searcher.Search("IsInstalled=0 and Type='Software'")
            result["pending"]       = res_search.Updates.Count > 0
            result["pending_count"] = res_search.Updates.Count
        except Exception:
            result["pending"] = None

        self.data_ready.emit(result)


class PatchesTab(QWidget):
    def __init__(self):
        super().__init__()
        self._patches = []
        self._build_ui()
        self._load()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        kpi_row = QHBoxLayout()
        self.kpi_os      = self._kpi("Système",            "—", ACCENT)
        self.kpi_build   = self._kpi("Build Windows",      "—", ACCENT2)
        self.kpi_boot    = self._kpi("Dernier démarrage",  "—", ACCENT3)
        self.kpi_hotfix  = self._kpi("Hotfixes installés", "—", GREEN)
        self.kpi_pending = self._kpi("MAJ en attente",     "Vérification…", ACCENT3)
        for k in [self.kpi_os, self.kpi_build, self.kpi_boot,
                  self.kpi_hotfix, self.kpi_pending]:
            kpi_row.addWidget(k)
        layout.addLayout(kpi_row)

        status_box = QGroupBox("État des patchs")
        sb_layout = QHBoxLayout(status_box)
        self.status_icon = QLabel("●")
        self.status_icon.setStyleSheet("font-size:28px; color:gray; border:none;")
        self.status_text = QLabel("Analyse en cours…")
        self.status_text.setStyleSheet(f"font-size:15px; color:{TEXT_PRI}; border:none;")
        sb_layout.addWidget(self.status_icon)
        sb_layout.addWidget(self.status_text)
        sb_layout.addStretch()
        self.check_btn = QPushButton("Ouvrir Windows Update")
        self.check_btn.clicked.connect(lambda: os.startfile("ms-settings:windowsupdate"))
        sb_layout.addWidget(self.check_btn)
        layout.addWidget(status_box)

        timeline_box = QGroupBox("Timeline des 12 derniers mois (hotfixes)")
        tl_layout = QVBoxLayout(timeline_box)
        self.timeline_plot = pg.PlotWidget()
        self.timeline_plot.setBackground(BG_CARD)
        self.timeline_plot.setMinimumHeight(160)
        self.timeline_plot.setLabel('left', "Nombre de patchs")
        self.timeline_plot.showGrid(x=False, y=True, alpha=0.12)
        self.bar_item = pg.BarGraphItem(
            x=list(range(12)), height=[0]*12, width=0.7, brush=ACCENT)
        self.timeline_plot.addItem(self.bar_item)
        tl_layout.addWidget(self.timeline_plot)
        layout.addWidget(timeline_box)

        hf_box = QGroupBox("Hotfixes installés (Win32_QuickFixEngineering)")
        hf_layout = QVBoxLayout(hf_box)
        filter_row = QHBoxLayout()
        self.hf_filter = QLineEdit()
        self.hf_filter.setPlaceholderText("Rechercher un KB…")
        self.hf_filter.textChanged.connect(self._filter_hf)
        filter_row.addWidget(self.hf_filter)
        self.refresh_btn = QPushButton("Rafraîchir")
        self.refresh_btn.clicked.connect(self._load)
        filter_row.addWidget(self.refresh_btn)
        hf_layout.addLayout(filter_row)

        self.hf_table = QTableWidget(0, 4)
        self.hf_table.setHorizontalHeaderLabels(
            ["KB / ID", "Description", "Date installation", "Installé par"])
        self.hf_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.hf_table.setSortingEnabled(True)
        self.hf_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        hf_layout.addWidget(self.hf_table)
        layout.addWidget(hf_box)

        self.status_bar = QLabel("Chargement…")
        self.status_bar.setStyleSheet(f"color:{TEXT_SEC}; font-size:11px;")
        layout.addWidget(self.status_bar)

    def _kpi(self, label, value, color):
        frame = QFrame()
        frame.setStyleSheet(
            f"QFrame {{ background:{BG_CARD}; border:1px solid {BORDER}; border-radius:8px; }}")
        v = QVBoxLayout(frame)
        v.setContentsMargins(12, 8, 12, 8)
        lbl = QLabel(label)
        lbl.setStyleSheet(f"color:{TEXT_SEC}; font-size:11px; border:none;")
        val = QLabel(value)
        val.setStyleSheet(f"color:{color}; font-size:16px; font-weight:700; border:none;")
        val.setWordWrap(True)
        v.addWidget(lbl)
        v.addWidget(val)
        frame._val = val
        return frame

    def _load(self):
        self.refresh_btn.setEnabled(False)
        self.status_bar.setText("Analyse WMI en cours…")
        self.worker = PatchWorker()
        self.worker.data_ready.connect(self._on_data)
        self.worker.start()

    def _on_data(self, data):
        self._patches = data["hotfixes"]
        caption = data["os_caption"]
        self.kpi_os._val.setText(caption[:30] if len(caption) > 30 else caption)
        self.kpi_build._val.setText(data["os_build"])
        self.kpi_boot._val.setText(data["last_boot"])
        self.kpi_hotfix._val.setText(str(len(data["hotfixes"])))

        if data["pending"] is None:
            self.kpi_pending._val.setText("N/D")
            self.status_icon.setStyleSheet("font-size:28px; color:gray; border:none;")
            self.status_text.setText("Impossible de vérifier les MAJ en attente (wuapi)")
        elif data["pending"]:
            cnt = data.get("pending_count", "?")
            self.kpi_pending._val.setText(str(cnt))
            self.kpi_pending._val.setStyleSheet(
                f"color:{RED}; font-size:16px; font-weight:700; border:none;")
            self.status_icon.setStyleSheet(f"font-size:28px; color:{RED}; border:none;")
            self.status_text.setText(f"⚠ {cnt} mise(s) à jour en attente — système NON à jour")
        else:
            self.kpi_pending._val.setText("0")
            self.kpi_pending._val.setStyleSheet(
                f"color:{GREEN}; font-size:16px; font-weight:700; border:none;")
            self.status_icon.setStyleSheet(f"font-size:28px; color:{GREEN}; border:none;")
            self.status_text.setText("Système à jour ✓")

        now = datetime.date.today()
        monthly = [0] * 12
        for hf in data["hotfixes"]:
            d = hf["date_obj"]
            if d:
                d_date = d.date() if isinstance(d, datetime.datetime) else d
                delta = (now.year - d_date.year)*12 + (now.month - d_date.month)
                if 0 <= delta < 12:
                    monthly[11 - delta] += 1
        self.bar_item.setOpts(height=monthly)
        self._populate_hf(data["hotfixes"])
        self.refresh_btn.setEnabled(True)
        self.status_bar.setText(f"{len(data['hotfixes'])} hotfixes — {data['os_caption']}")

    def _populate_hf(self, patches):
        self.hf_table.setRowCount(len(patches))
        for r, p in enumerate(patches):
            self.hf_table.setItem(r, 0, QTableWidgetItem(p["kb"]))
            self.hf_table.setItem(r, 1, QTableWidgetItem(p["desc"]))
            self.hf_table.setItem(r, 2, QTableWidgetItem(p["date"]))
            self.hf_table.setItem(r, 3, QTableWidgetItem(p["installed_by"]))

    def _filter_hf(self, text):
        text = text.lower()
        filtered = [p for p in self._patches
                    if text in p["kb"].lower() or text in p["desc"].lower()]
        self._populate_hf(filtered)

    def get_summary(self):
        return {"hotfixes": self._patches}


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE D'ACCUEIL (Dashboard)
# ═══════════════════════════════════════════════════════════════════════════════
class DashboardTab(QWidget):
    def __init__(self, net_tab, app_tab, patch_tab):
        super().__init__()
        self.net_tab   = net_tab
        self.app_tab   = app_tab
        self.patch_tab = patch_tab
        self._build_ui()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._refresh)
        self.timer.start(2000)

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)

        title = QLabel("SysMonitor v0.1 — Vue d'ensemble")
        title.setStyleSheet(f"font-size:22px; font-weight:700; color:{ACCENT}; border:none;")
        layout.addWidget(title)

        sub = QLabel(f"Machine locale — {datetime.datetime.now().strftime('%A %d %B %Y')}")
        sub.setStyleSheet(f"font-size:12px; color:{TEXT_SEC}; border:none; margin-bottom:4px;")
        layout.addWidget(sub)

        kpi_row = QHBoxLayout()
        self.kpi_cpu  = self._stat_card("CPU",                "—", "%",         ACCENT)
        self.kpi_ram  = self._stat_card("RAM utilisée",       "—", "%",         ACCENT2)
        self.kpi_disk = self._stat_card("Disque C:",          "—", "% utilisé", ACCENT3)
        self.kpi_net  = self._stat_card("Connexions actives", "—", "",          GREEN)
        for k in [self.kpi_cpu, self.kpi_ram, self.kpi_disk, self.kpi_net]:
            kpi_row.addWidget(k)
        layout.addLayout(kpi_row)

        graphs_row = QHBoxLayout()

        net_box = QGroupBox("Réseau — Bande passante")
        nb_layout = QVBoxLayout(net_box)
        self.mini_net = make_plot(y_label="KB/s")
        self.mini_net.setMaximumHeight(180)
        self.mini_sent = self.mini_net.plot(pen=pg.mkPen(RED,    width=1.5), name="↑")
        self.mini_recv = self.mini_net.plot(pen=pg.mkPen(ACCENT, width=1.5), name="↓")
        nb_layout.addWidget(self.mini_net)
        graphs_row.addWidget(net_box)

        cpu_box = QGroupBox("CPU & RAM (60 s)")
        cb_layout = QVBoxLayout(cpu_box)
        self.mini_cpu_plot = make_plot(y_label="%")
        self.mini_cpu_plot.setMaximumHeight(180)
        self.mini_cpu_plot.setYRange(0, 100)
        self.cpu_hist = deque([0.0]*60, maxlen=60)
        self.ram_hist = deque([0.0]*60, maxlen=60)
        self.mini_cpu_curve = self.mini_cpu_plot.plot(
            pen=pg.mkPen(ACCENT2, width=1.5), name="CPU")
        self.mini_ram_curve = self.mini_cpu_plot.plot(
            pen=pg.mkPen(ACCENT3, width=1.5), name="RAM")
        cb_layout.addWidget(self.mini_cpu_plot)
        graphs_row.addWidget(cpu_box)

        patch_box = QGroupBox("Patchs — 12 derniers mois")
        pb_layout = QVBoxLayout(patch_box)
        self.mini_patch = pg.PlotWidget()
        self.mini_patch.setBackground(BG_CARD)
        self.mini_patch.setMaximumHeight(180)
        self.mini_patch.showGrid(x=False, y=True, alpha=0.12)
        self.mini_patch_bars = pg.BarGraphItem(
            x=list(range(12)), height=[0]*12, width=0.7, brush=GREEN)
        self.mini_patch.addItem(self.mini_patch_bars)
        pb_layout.addWidget(self.mini_patch)
        graphs_row.addWidget(patch_box)

        layout.addLayout(graphs_row)

        bottom_row = QHBoxLayout()

        patch_info_box = QGroupBox("État des mises à jour")
        pib_layout = QVBoxLayout(patch_info_box)
        self.dash_patch_status = QLabel("Analyse en cours…")
        self.dash_patch_status.setStyleSheet(f"font-size:14px; border:none;")
        self.dash_patch_status.setWordWrap(True)
        pib_layout.addWidget(self.dash_patch_status)
        bottom_row.addWidget(patch_info_box)

        apps_box = QGroupBox("Top 5 applications (mémoire active)")
        ab_layout = QVBoxLayout(apps_box)
        self.top_apps_table = QTableWidget(5, 2)
        self.top_apps_table.setHorizontalHeaderLabels(["Application", "Mémoire (MB)"])
        self.top_apps_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)
        self.top_apps_table.setMaximumHeight(180)
        ab_layout.addWidget(self.top_apps_table)
        bottom_row.addWidget(apps_box)

        layout.addLayout(bottom_row)
        layout.addStretch()

    def _stat_card(self, label, value, unit, color):
        frame = QFrame()
        frame.setStyleSheet(
            f"QFrame {{ background:{BG_CARD}; border:1px solid {BORDER}; border-radius:10px; }}")
        v = QVBoxLayout(frame)
        v.setContentsMargins(16, 12, 16, 12)
        lbl = QLabel(label)
        lbl.setStyleSheet(
            f"color:{TEXT_SEC}; font-size:11px; font-weight:600; "
            f"letter-spacing:1px; border:none;")
        val = QLabel(value)
        val.setStyleSheet(f"color:{color}; font-size:32px; font-weight:800; border:none;")
        unt = QLabel(unit)
        unt.setStyleSheet(f"color:{TEXT_SEC}; font-size:12px; border:none;")
        v.addWidget(lbl)
        v.addWidget(val)
        v.addWidget(unt)
        frame._val = val
        return frame

    def _refresh(self):
        cpu  = psutil.cpu_percent()
        ram  = psutil.virtual_memory().percent
        disk = psutil.disk_usage("C:/").percent
        self.kpi_cpu._val.setText(f"{cpu:.0f}")
        self.kpi_ram._val.setText(f"{ram:.0f}")
        self.kpi_disk._val.setText(f"{disk:.0f}")
        self.cpu_hist.append(cpu)
        self.ram_hist.append(ram)
        self.mini_cpu_curve.setData(list(self.cpu_hist))
        self.mini_ram_curve.setData(list(self.ram_hist))

        net = self.net_tab.get_summary()
        self.mini_sent.setData(net["sent"])
        self.mini_recv.setData(net["recv"])
        try:
            self.kpi_net._val.setText(
                str(len(psutil.net_connections(kind='inet'))))
        except Exception:
            pass

        patches = self.patch_tab.get_summary()
        now = datetime.date.today()
        monthly = [0] * 12
        for hf in patches["hotfixes"]:
            d = hf["date_obj"]
            if d:
                d_date = d.date() if isinstance(d, datetime.datetime) else d
                delta = (now.year - d_date.year)*12 + (now.month - d_date.month)
                if 0 <= delta < 12:
                    monthly[11 - delta] += 1
        self.mini_patch_bars.setOpts(height=monthly)
        if patches["hotfixes"]:
            last_hf = patches["hotfixes"][0]
            self.dash_patch_status.setText(
                f"<span style='color:{GREEN};font-weight:700'>"
                f"{len(patches['hotfixes'])}</span> hotfixes installés<br>"
                f"Dernier : {last_hf['kb']} — {last_hf['date']}"
            )
        else:
            self.dash_patch_status.setText("Chargement…")

        apps = sorted(self.app_tab.get_apps(), key=lambda a: -a.get("mem_mb", 0))[:5]
        for r, a in enumerate(apps):
            name_item = QTableWidgetItem(a["name"][:32])
            name_item.setForeground(QColor(ORANGE))
            self.top_apps_table.setItem(r, 0, name_item)
            mem_item = QTableWidgetItem(f"{a.get('mem_mb', 0):.0f}")
            if a.get("mem_mb", 0) > 500:
                mem_item.setForeground(QColor(RED))
            self.top_apps_table.setItem(r, 1, mem_item)


# ═══════════════════════════════════════════════════════════════════════════════
# FENÊTRE PRINCIPALE
# ═══════════════════════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SysMonitor v0.1 — Tableau de bord système")
        self.resize(1400, 900)
        self.setStyleSheet(STYLESHEET)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(0)

        tabs = QTabWidget()

        self.net_tab   = NetworkTab()
        self.app_tab   = AppsTab()
        self.patch_tab = PatchesTab()
        self.dash_tab  = DashboardTab(self.net_tab, self.app_tab, self.patch_tab)

        tabs.addTab(self.dash_tab,  "🏠  Vue d'ensemble")
        tabs.addTab(self.net_tab,   "🌐  Réseau")
        tabs.addTab(self.app_tab,   "📦  Applications")
        tabs.addTab(self.patch_tab, "🔒  Patches / MAJ")

        main_layout.addWidget(tabs)

    def closeEvent(self, event):
        self.net_tab.stop()
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
