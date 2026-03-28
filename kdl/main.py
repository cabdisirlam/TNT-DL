"""
NT DL Multipurpose Tool
Application entry point.
"""

import sys
import os
import ctypes

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QApplication, QSplashScreen
from PySide6.QtCore import Qt, QTimer, QRect
from PySide6.QtGui import QFont, QPixmap, QPainter, QColor, QLinearGradient, QIcon

from kdl import __app_name__, __display_name__, __version__
from kdl.main_window import MainWindow
from kdl.window.window_manager import WindowManager


_SINGLE_INSTANCE_MUTEX = "Local\\NT_DL_SingleInstance"
_ERROR_ALREADY_EXISTS = 183
_SW_RESTORE = 9


def resource_path(relative_path: str) -> str:
    """Return absolute path for source and PyInstaller runtime."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base_path, relative_path)


def create_splash_pixmap():
    """Create a cleaner, classic-style splash screen."""
    pixmap = QPixmap(760, 340)
    painter = QPainter(pixmap)

    # Main background
    painter.fillRect(pixmap.rect(), QColor("#111111"))

    # Left accent panel
    left_rect = QRect(0, 0, 220, 340)
    panel_gradient = QLinearGradient(0, 0, 220, 340)
    panel_gradient.setColorAt(0.0, QColor("#7FAEE0"))
    panel_gradient.setColorAt(1.0, QColor("#5B90C7"))
    painter.fillRect(left_rect, panel_gradient)

    # Draw app icon/logo on left
    icon_path = resource_path(os.path.join("kdl", "assets", "kdl_a.ico"))
    logo = QPixmap(icon_path) if os.path.exists(icon_path) else QPixmap()
    if not logo.isNull():
        logo_scaled = logo.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        x = (left_rect.width() - logo_scaled.width()) // 2
        y = 64
        painter.drawPixmap(x, y, logo_scaled)
    else:
        painter.setPen(QColor("#FFFFFF"))
        painter.setFont(QFont("Segoe UI", 46, QFont.Bold))
        painter.drawText(left_rect, Qt.AlignCenter, "NT DL")

    painter.setPen(QColor("#FFFFFF"))
    painter.setFont(QFont("Segoe UI", 10, QFont.Bold))
    painter.drawText(QRect(18, 266, 184, 34), Qt.AlignLeft | Qt.AlignVCenter, "NT DL Multipurpose Tool")
    painter.setFont(QFont("Segoe UI", 9))
    painter.drawText(QRect(18, 294, 184, 22), Qt.AlignLeft | Qt.AlignVCenter, "")

    # Right-side text block
    right_x = 246
    painter.setPen(QColor("#FF2D2D"))
    painter.setFont(QFont("Segoe UI", 34, QFont.Bold))
    painter.drawText(QRect(right_x, 30, 490, 56), Qt.AlignLeft | Qt.AlignVCenter, "NT DL")

    painter.setPen(QColor("#FFFFFF"))
    painter.setFont(QFont("Segoe UI", 22, QFont.Bold))
    painter.drawText(QRect(right_x, 88, 490, 44), Qt.AlignLeft | Qt.AlignVCenter, f"Version {__version__}")

    painter.setPen(QColor("#E2E2E2"))
    painter.setFont(QFont("Segoe UI", 14))
    painter.drawText(QRect(right_x, 146, 490, 34), Qt.AlignLeft | Qt.AlignVCenter, "Automation, Reports, and Utilities")
    painter.drawText(QRect(right_x, 184, 490, 34), Qt.AlignLeft | Qt.AlignVCenter, "Oracle, ERP, and workflow tools")

    painter.setPen(QColor("#B9B9B9"))
    painter.setFont(QFont("Segoe UI", 11))
    painter.drawText(QRect(right_x, 254, 490, 28), Qt.AlignLeft | Qt.AlignVCenter, "www.ntdl.local")
    painter.drawText(QRect(right_x, 282, 490, 28), Qt.AlignLeft | Qt.AlignVCenter, "Support: Workflow Operations")

    painter.end()
    return pixmap


def _install_exception_hook():
    """Install global exception handler so packaged app doesn't crash silently."""
    import traceback
    import logging

    logger = logging.getLogger("kdl")

    def _handle_exception(exc_type, exc_value, exc_tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logger.critical("Unhandled exception:\n%s", msg)
        # Show error dialog if QApplication exists
        try:
            from PySide6.QtWidgets import QMessageBox, QApplication as _App
            if _App.instance():
                QMessageBox.critical(
                    None, f"{__display_name__} - Unexpected Error",
                    f"An unexpected error occurred:\n\n{exc_value}\n\n"
                    "The application may be unstable. Please save your work.",
                )
        except Exception:
            pass

    sys.excepthook = _handle_exception


def _activate_existing_instance() -> bool:
    """Best-effort foreground activation for the already-running main window."""
    try:
        for hwnd, title in WindowManager.get_open_windows():
            text = (title or "").strip()
            if text != __display_name__ and not text.startswith(f"{__display_name__}  |"):
                continue
            if ctypes.windll.user32.IsIconic(hwnd):
                ctypes.windll.user32.ShowWindow(hwnd, _SW_RESTORE)
            WindowManager.activate_window(hwnd)
            return True
    except Exception:
        pass
    return False


def _acquire_single_instance_mutex() -> tuple[int | None, bool]:
    """
    Acquire a Windows named mutex for single-instance protection.
    Returns (handle, already_running).
    """
    if os.name != "nt":
        return None, False

    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, _SINGLE_INSTANCE_MUTEX)
    if not mutex:
        return None, False

    already_running = ctypes.windll.kernel32.GetLastError() == _ERROR_ALREADY_EXISTS
    return int(mutex), already_running


def main():
    """Main application entry point."""
    _install_exception_hook()

    mutex_handle, already_running = _acquire_single_instance_mutex()
    if already_running and _activate_existing_instance():
        return 0

    # High DPI support
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app._single_instance_mutex = mutex_handle
    if mutex_handle:
        app.aboutToQuit.connect(
            lambda handle=mutex_handle: ctypes.windll.kernel32.CloseHandle(handle)
        )

    # Application metadata
    app.setApplicationName(__app_name__)
    app.setApplicationDisplayName(__display_name__)
    app.setApplicationVersion(__version__)
    app.setOrganizationName(__app_name__)

    # App icon
    icon_path = resource_path(os.path.join("kdl", "assets", "kdl_a.ico"))
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    # Set default font
    font = QFont("Segoe UI", 11)
    app.setFont(font)

    # Splash screen
    splash_pixmap = create_splash_pixmap()
    splash = QSplashScreen(splash_pixmap)
    splash.show()
    app.processEvents()

    # Create main window
    window = MainWindow()
    if not app.windowIcon().isNull():
        window.setWindowIcon(app.windowIcon())

    def _activate_main_window():
        if window.isMinimized():
            window.showNormal()
        elif not window.isVisible():
            window.show()
        window.raise_()
        window.activateWindow()

    # Hide splash and show main window after a short delay
    QTimer.singleShot(1500, lambda: (splash.finish(window), window.show(), _activate_main_window()))

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
