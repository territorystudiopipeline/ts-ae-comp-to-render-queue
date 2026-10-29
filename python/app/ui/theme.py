# Copyright (c) 2026 Territory Studio
"""
Look and status widgets for the app, built on ts_qt_utils (StyleLoader + InfoTag).

ts_qt_utils comes from the After Effects rez environment (like scene_fx_manager), and is
imported through QtPy. QtPy must resolve to the same Qt binding Toolkit uses: widgets from
two bindings in one process crash. If ts_qt_utils is missing, or QtPy picked a different
binding, the app falls back to plain widgets and the default Toolkit look.
"""
import os

import sgtk
from sgtk.platform.qt import QtCore, QtGui

logger = sgtk.platform.get_logger(__name__)

ICON_DIR = os.path.join(os.path.dirname(__file__), "icons")
OVERRIDES_QSS = os.path.join(os.path.dirname(__file__), "app_overrides.qss")

_TS_QT_UTILS = None  # None = not tried yet, {} = unavailable


def _qt_api_name():
    """The QT_API value matching Toolkit's binding, e.g. 'pyside2'."""
    return QtCore.__name__.split(".")[0].lower()


def ts_qt_utils():
    """
        Import ts_qt_utils once.

        Returns:
            dict: {"InfoTag": cls, "StyleLoader": cls}, or None if unavailable.
    """
    global _TS_QT_UTILS
    if _TS_QT_UTILS is None:
        try:
            # Point QtPy at Toolkit's binding unless something already chose one.
            os.environ.setdefault("QT_API", _qt_api_name())
            from qtpy import QtWidgets as _QtPyWidgets
            if _QtPyWidgets.QWidget is not QtGui.QWidget:
                raise ImportError(
                    "QtPy resolved to %s but Toolkit uses %s"
                    % (_QtPyWidgets.__name__, QtCore.__name__.split(".")[0])
                )
            from ts_qt_utils.widgets.info_tag import InfoTag
            from ts_qt_utils.style_sheets.style_loader import StyleLoader
            _TS_QT_UTILS = {"InfoTag": InfoTag, "StyleLoader": StyleLoader}
        except Exception as e:
            logger.warning("ts_qt_utils not available, using the plain look: %s" % e)
            _TS_QT_UTILS = {}
    return _TS_QT_UTILS or None


def build_stylesheet(theme="dark"):
    """
        The ts_qt_utils stylesheet plus this app's overrides, or "" without ts_qt_utils.
    """
    lib = ts_qt_utils()
    if not lib:
        return ""
    try:
        extra = [OVERRIDES_QSS] if os.path.isfile(OVERRIDES_QSS) else None
        return lib["StyleLoader"].build(theme=theme, extra=extra)
    except Exception as e:
        logger.warning("Could not build the ts_qt_utils stylesheet: %s" % e)
        return ""


# Colours from ts_qt_utils dark tokens, for the app-specific states below.
_TOKENS = {
    "ELEVATED": "#323234", "ELEVATED2": "#3c3c3e", "BORDER_SOFT": "#48484a",
    "TEXT_SEC": "#a09d98", "TEXT_DIM": "#636360", "ACCENT_MID": "#7a5c2a", "ACCENT_LT": "#e0b85a",
    "ACCENT": "#d4a84b", "BG": "#1c1c1e",
}

# key -> (InfoTag state, default title, fallback icon file)
STATUSES = {
    "ready": ("info", "Ready", "clear_icon.png"),
    "needs_output": ("warning", "Needs output", "warning_icon.png"),
    "warning": ("warning", "Skipped", "warning_icon.png"),
    "error": ("error", "Error", "error_icon.png"),
    "busy": ("busy", "Working", "refresh_icon.png"),
    "submitted": ("submitted", "Submitted", "submit_icon.png"),
    "rendering": ("info", "Render ready", "clear_icon.png"),
}

# Extra InfoTag states this app uses (the built-in ones are critical/error/warning/info).
APP_STATE_STYLES = {
    "submitted": {"title_bg": "#0F5B7A", "title_fg": "#EAF8FF", "message_fg": _TOKENS["TEXT_SEC"]},
    # Gold like the primary button; the loading spinner is tinted with this colour
    "busy": {"title_bg": _TOKENS["ACCENT"], "title_fg": _TOKENS["BG"], "message_fg": _TOKENS["TEXT_SEC"]},
    "neutral": {"title_bg": _TOKENS["ELEVATED2"], "title_fg": _TOKENS["TEXT_SEC"], "message_fg": _TOKENS["TEXT_SEC"]},
}
# Message text for the built-in states: secondary text reads better inside a table.
for _state in ("critical", "error", "warning", "info"):
    APP_STATE_STYLES.setdefault(_state, {})["message_fg"] = _TOKENS["TEXT_SEC"]


def _icon(file_name):
    return QtGui.QIcon(os.path.join(ICON_DIR, file_name))


def _app_state_icons():
    return {
        "submitted": _icon("submit_icon.png"),
        "busy": _icon("refresh_icon.png"),
        "neutral": _icon("warning_icon.png"),
    }


def _make_info_tag(lib, framed, **kwargs):
    """
        An InfoTag whose frame style is scoped to the tag itself.

        InfoTag styles itself with a bare "QFrame { border ... }", which also matches its own
        QLabel children (QLabel is a QFrame), drawing a second border around the text when the
        frame is visible. Re-applying the frame style as "QFrame#InfoTag" avoids that.
    """
    InfoTag = lib["InfoTag"]

    class _ScopedInfoTag(InfoTag):
        def _apply_styles(self):
            super(_ScopedInfoTag, self)._apply_styles()
            self.setObjectName("InfoTag")
            self.setStyleSheet(
                "QFrame#InfoTag {"
                " background-color: %s; border: 1px solid %s; border-radius: %spx; }"
                % (self._frame_bg_colour, self._frame_border_colour, self._frame_radius)
            )

    options = dict(
        state_styles=APP_STATE_STYLES, state_icons=_app_state_icons(), enable_state_icons=True,
        selectable=False, icon_size=(14, 14), spacing=6, font_size=11, min_height=26,
        title_padding="1px 8px", message_padding="0px 2px", title_radius=9,
    )
    if framed:
        options.update(frame_bg_colour=_TOKENS["ELEVATED"], frame_border_colour=_TOKENS["BORDER_SOFT"],
                       margins=(6, 2, 6, 2), frame_radius=8, min_height=28, message_padding="0px 4px")
    else:
        options.update(frame_bg_colour="transparent", frame_border_colour="transparent", margins=(4, 1, 4, 1))
    options.update(kwargs)
    return _ScopedInfoTag(**options)


class StatusTag(QtGui.QWidget):
    """
        Status display: an InfoTag (icon, coloured title, message) when ts_qt_utils is
        available, otherwise an icon and a label. The full detail goes in the tooltip.

        Usage:
            tag = StatusTag()
            tag.set_status("needs_output", "Apply a template", tooltip="Template not applied")
    """

    def __init__(self, status="ready", message="", tooltip="", framed=False, parent=None):
        super(StatusTag, self).__init__(parent)
        self.status = None
        self._framed = framed
        layout = QtGui.QHBoxLayout(self)
        layout.setContentsMargins(4, 0, 4, 0)
        layout.setSpacing(6)

        lib = ts_qt_utils()
        if lib:
            self._tag = _make_info_tag(lib, framed)
            layout.addWidget(self._tag)
            self._icon_label = self._text_label = None
        else:
            self._tag = None
            self._icon_label = QtGui.QLabel(self)
            self._text_label = QtGui.QLabel(self)
            layout.addWidget(self._icon_label)
            layout.addWidget(self._text_label)
        layout.addStretch()
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        # Keep status colours when the table is locked (DraggableTableWidget.set_locked)
        self.setProperty("keepEnabledWhenLocked", True)
        self.set_status(status, message, tooltip)

    def set_status(self, status, message="", tooltip="", title=None):
        """
            Arguments:
                status (str): A key of STATUSES.
                message (str): Short text next to the title.
                tooltip (str): Full detail on hover.
                title (str): Overrides the status' default title.
        """
        state, default_title, icon_file = STATUSES.get(status, STATUSES["error"])
        title = default_title if title is None else title
        self.status = status
        if self._tag is not None:
            self._tag.state = state
            self._tag.title = title
            self._tag.message = message
            # Rotating loading icon while in progress (ts_qt_utils releases with InfoTag.loading;
            # older ones keep the static "busy" icon)
            if hasattr(type(self._tag), "loading"):
                self._tag.loading = state == "busy"
            self._tag.adjustSize()
        else:
            self._icon_label.setPixmap(_icon(icon_file).pixmap(16, 16))
            self._text_label.setText("%s  %s" % (title, message) if message else title)
        self.setToolTip(tooltip or (("%s: %s" % (title, message)) if message else title))
