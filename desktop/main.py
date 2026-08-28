# -*- coding: utf-8 -*-
"""Glasses Filler — desktop entry point.

Run from source:   "C:\\gv\\Scripts\\python.exe" desktop\\main.py
Build:             "C:\\gv\\Scripts\\pyinstaller.exe" desktop\\GlassesFiller.spec
"""
from __future__ import annotations

import os
import sys

# --- import bootstrap -------------------------------------------------------
# The shared logic (filler_core, ingest, dictionaries) lives in the repo root;
# the desktop modules import each other flat (`import theme`). Make both work
# whether running from source or from a PyInstaller bundle.
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
for _p in (_HERE, _ROOT):
    if _p not in sys.path:
        sys.path.insert(0, _p)
# ---------------------------------------------------------------------------

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtGui import QAction, QKeySequence  # noqa: E402
from PySide6.QtWidgets import (  # noqa: E402
    QApplication, QDockWidget, QLabel, QMainWindow, QMessageBox, QProgressBar,
    QTabWidget, QVBoxLayout, QWidget,
)

import data_source  # noqa: E402
import theme  # noqa: E402
import updater  # noqa: E402
from settings import Settings  # noqa: E402
from settings_dialog import SettingsDialog  # noqa: E402
from tabs import ALL_TABS  # noqa: E402
from version import APP_NAME, ORG_NAME, __version__  # noqa: E402
from workers import Worker  # noqa: E402


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = Settings()
        self._catalogue = None
        self._loader = None
        self._update_worker = None

        self.setWindowTitle(f"{APP_NAME} {__version__}")
        self.resize(1400, 880)

        # --- tabs ---
        self.tabs = QTabWidget()
        self.tab_widgets = []
        for cls in ALL_TABS:
            if getattr(cls, "NEEDS_ADMIN", False) and not self.settings.admin_enabled:
                continue
            tab = cls(self.settings, self)
            tab.status_message.connect(self.show_status)
            tab.busy.connect(self.set_busy)
            self.tabs.addTab(tab, cls.TITLE)
            self.tab_widgets.append(tab)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.setCentralWidget(self.tabs)

        # --- right dock: control panel ---
        self.dock = QDockWidget("Control panel", self)
        self.dock.setObjectName("control_panel")
        self.dock.setAllowedAreas(Qt.RightDockWidgetArea | Qt.LeftDockWidgetArea)
        self.dock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable)
        self._empty_dock = QWidget()
        QVBoxLayout(self._empty_dock)
        self.dock.setWidget(self._empty_dock)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock)
        self.resizeDocks([self.dock], [340], Qt.Horizontal)

        self._build_toolbar()
        self._build_statusbar()

        self._on_tab_changed(self.tabs.currentIndex())
        self.refresh_data()
        self._check_updates_quietly()

    # ------------------------------------------------------------- toolbar
    def _build_toolbar(self) -> None:
        tb = self.addToolBar("Main")
        tb.setObjectName("main_toolbar")
        tb.setMovable(False)

        def add(text, slot, shortcut=None, checkable=False):
            act = QAction(text, self)
            act.triggered.connect(slot)
            if shortcut:
                act.setShortcut(QKeySequence(shortcut))
            act.setCheckable(checkable)
            tb.addAction(act)
            return act

        self.act_open = add("📂 Open", self._route_open, "Ctrl+O")
        self.act_run = add("🪄 Fill columns", self._route_run, "Ctrl+R")
        self.act_save = add("💾 Save changes", self._route_save, "Ctrl+S")
        tb.addSeparator()
        self.act_refresh = add("☁️ Refresh data", lambda: self.refresh_data(force=True), "F5")
        tb.addSeparator()
        self.act_settings = add("⚙️ Settings", self.open_settings)
        self.act_update = add("⬆️ Check updates", self.check_updates)
        self.act_dark = add("🌙 Dark mode", self.toggle_dark, checkable=True)
        self.act_dark.setChecked(self.settings.dark_mode)

    def _build_statusbar(self) -> None:
        sb = self.statusBar()

        self.data_status = QLabel("Catalogue: loading…")
        self.data_status.setStyleSheet(theme.status_style("loading"))
        sb.addPermanentWidget(self.data_status)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)            # indeterminate
        self.progress.setMaximumWidth(180)
        self.progress.setVisible(False)
        sb.addPermanentWidget(self.progress)

        sb.showMessage("Ready.")

    # ----------------------------------------------------------------- tabs
    def current_tab(self):
        w = self.tabs.currentWidget()
        return w

    def _on_tab_changed(self, index: int) -> None:
        tab = self.tabs.widget(index)
        if tab is None:
            return

        panel = tab.control_panel()
        if panel is None:
            self.dock.setVisible(False)
        else:
            self.dock.setWidget(panel)
            self.dock.setVisible(True)

        supports = getattr(tab, "SUPPORTS", set())
        self.act_open.setEnabled("open" in supports)
        self.act_run.setEnabled("run" in supports)
        self.act_save.setEnabled("save" in supports)
        self._update_title()

    def _route(self, name: str) -> None:
        tab = self.current_tab()
        if tab is not None and name in getattr(tab, "SUPPORTS", set()):
            getattr(tab, {"open": "open_file", "run": "run_action", "save": "save_action"}[name])()

    def _route_open(self):
        self._route("open")

    def _route_run(self):
        self._route("run")

    def _route_save(self):
        self._route("save")

    # ------------------------------------------------------------ status bar
    def show_status(self, message: str) -> None:
        self.statusBar().showMessage(message, 15000)
        self._update_title()

    def set_busy(self, busy: bool) -> None:
        self.progress.setVisible(bool(busy))
        if busy:
            for act in (self.act_open, self.act_run, self.act_save, self.act_refresh):
                act.setEnabled(False)
        else:
            # Re-enable according to what the current tab supports.
            self.act_refresh.setEnabled(True)
            self._on_tab_changed(self.tabs.currentIndex())

    def _update_title(self) -> None:
        dirty = any(t.has_unsaved_changes() for t in self.tab_widgets)
        self.setWindowTitle(f"{APP_NAME} {__version__}" + (" •" if dirty else ""))

    # --------------------------------------------------------------- data
    def refresh_data(self, force: bool = False) -> None:
        self.data_status.setText("Catalogue: loading…")
        self.data_status.setStyleSheet(theme.status_style("loading"))
        self.set_busy(True)

        settings = self.settings
        self._loader = Worker(
            data_source.load_catalogue, settings, force_refresh=force, pass_progress=True
        )
        self._loader.progress.connect(lambda f, t: self.statusBar().showMessage(t, 10000))
        self._loader.done.connect(self._on_data_loaded)
        self._loader.failed.connect(self._on_data_failed)
        self._loader.start()

    def _on_data_loaded(self, data) -> None:
        self.set_busy(False)
        self._catalogue = data

        label = f"Catalogue: {len(data.master_db):,} products"
        if data.source == "cache":
            label += " (offline cache)"
        elif data.source == "database":
            label += " (live DB)"
        if data.generated_utc:
            label += f" · {data.generated_utc[:10]}"
        self.data_status.setText(label)
        self.data_status.setStyleSheet(
            theme.status_style("loading" if data.source == "cache" else "ready")
        )

        for tab in self.tab_widgets:
            tab.set_catalogue(data)

        if data.messages:
            self.statusBar().showMessage(" | ".join(data.messages), 20000)
        else:
            self.statusBar().showMessage(f"Catalogue ready — {len(data.master_db):,} products.", 8000)

    def _on_data_failed(self, message: str) -> None:
        self.set_busy(False)
        self.data_status.setText("Catalogue: not loaded")
        self.data_status.setStyleSheet(theme.status_style("error"))
        box = QMessageBox(QMessageBox.Warning, "Catalogue not loaded", message, parent=self)
        detail = getattr(self._loader, "error_detail", "")
        if detail:
            box.setDetailedText(detail)
        box.addButton("Open settings", QMessageBox.AcceptRole)
        box.addButton(QMessageBox.Close)
        box.exec()
        if box.clickedButton() and box.clickedButton().text() == "Open settings":
            self.open_settings()

    # ------------------------------------------------------------- settings
    def open_settings(self) -> None:
        dlg = SettingsDialog(self.settings, self)
        if dlg.exec():
            # refresh_data pushes the new catalogue to every tab when it lands.
            self.refresh_data(force=True)

    def toggle_dark(self, checked: bool) -> None:
        self.settings.dark_mode = checked
        theme.apply_theme(QApplication.instance(), checked)

    # -------------------------------------------------------------- updates
    def _check_updates_quietly(self) -> None:
        self._update_worker = Worker(updater.update_available)
        self._update_worker.done.connect(self._on_update_checked)
        self._update_worker.failed.connect(lambda _m: None)
        self._update_worker.start()

    def _on_update_checked(self, release) -> None:
        if not release:
            return
        self.act_update.setText(f"⬆️ Update to {release['version']}")
        self.show_status(f"Version {release['version']} is available — use ⬆️ to update.")

    def check_updates(self) -> None:
        self.set_busy(True)
        self.show_status("Checking for updates…")
        worker = Worker(updater.update_available)

        def done(release):
            self.set_busy(False)
            if not release:
                QMessageBox.information(
                    self, "Up to date", f"You're on the latest version ({__version__})."
                )
                return
            answer = QMessageBox.question(
                self, "Update available",
                f"Version {release['version']} is available (you have {__version__}).\n\n"
                "Download and install it now? The app will restart.",
            )
            if answer != QMessageBox.Yes:
                return
            self._do_update(release)

        worker.done.connect(done)
        worker.failed.connect(lambda m: (self.set_busy(False), QMessageBox.warning(
            self, "Could not check for updates", m)))
        worker.start()
        self._update_worker = worker

    def _do_update(self, release) -> None:
        self.set_busy(True)
        worker = Worker(updater.download_and_swap, release, pass_progress=True)
        worker.progress.connect(lambda f, t: self.statusBar().showMessage(t, 10000))

        def done(_path):
            self.set_busy(False)
            QMessageBox.information(
                self, "Restarting",
                "The update was downloaded. The app will now close and reopen.",
            )
            QApplication.instance().quit()

        worker.done.connect(done)
        worker.failed.connect(lambda m: (self.set_busy(False), QMessageBox.warning(
            self, "Update failed", m)))
        worker.start()
        self._update_worker = worker

    # ----------------------------------------------------------------- close
    def closeEvent(self, event) -> None:
        if any(t.has_unsaved_changes() for t in self.tab_widgets):
            answer = QMessageBox.question(
                self, "Unsaved changes",
                "A filled file hasn't been saved yet. Close anyway?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                event.ignore()
                return
        event.accept()


def _selftest(win) -> int:
    """--selftest: build the whole UI, exercise the tab/dock wiring, print a
    report and exit. Runs headless with QT_QPA_PLATFORM=offscreen, so it can be
    used as a smoke test without a display."""
    # The Windows console is cp1250 — emoji in tab titles would raise.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    problems = []
    print(f"{APP_NAME} {__version__}")
    print(f"tabs: {win.tabs.count()}")
    for i in range(win.tabs.count()):
        tab = win.tabs.widget(i)
        win.tabs.setCurrentIndex(i)
        QApplication.processEvents()
        panel = tab.control_panel()
        print(
            f"  [{i}] {win.tabs.tabText(i)!r:26} supports={sorted(tab.SUPPORTS)} "
            f"dock={'yes' if panel is not None else 'no'} "
            f"admin={getattr(tab, 'NEEDS_ADMIN', False)}"
        )
        for name in ("open_file", "run_action", "save_action", "has_unsaved_changes"):
            if not callable(getattr(tab, name, None)):
                problems.append(f"{tab.__class__.__name__} missing {name}")

    print(f"toolbar actions: open={win.act_open.isEnabled()} run={win.act_run.isEnabled()} "
          f"save={win.act_save.isEnabled()}")
    # isVisible() is always False while the window itself was never shown, so
    # ask whether the dock was explicitly hidden instead.
    print(f"dock shown on current tab: {not win.dock.isHidden()}")
    print(f"data status: {win.data_status.text()!r}")
    print(f"admin unlocked: {win.settings.admin_enabled}")
    print(f"snapshot configured: {bool(win.settings.snapshot_repo)}")

    if problems:
        print("PROBLEMS:")
        for p in problems:
            print("  -", p)
        return 1
    print("SELFTEST OK — UI built, all tabs satisfy the BaseTab contract.")
    return 0


def main() -> int:
    selftest = "--selftest" in sys.argv
    if selftest:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(ORG_NAME)
    app.setApplicationVersion(__version__)

    settings = Settings()
    theme.apply_theme(app, settings.dark_mode)

    win = MainWindow()
    if selftest:
        QApplication.processEvents()
        return _selftest(win)

    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
