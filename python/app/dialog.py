# Copyright (c) 2013 Shotgun Software Inc.
#
# CONFIDENTIAL AND PROPRIETARY
#
# This work is provided "AS IS" and subject to the Shotgun Pipeline Toolkit
# Source Code License included in this distribution package. See LICENSE.
# By accessing, using, copying or modifying this work you indicate your
# agreement to the Shotgun Pipeline Toolkit Source Code License. All rights
# not expressly granted therein are reserved by Shotgun Software Inc.

import sgtk
import os
import sys
import threading
import traceback

# by importing QT from sgtk rather than directly, we ensure that
# the code will be compatible with both PySide and PyQt.
from sgtk.platform.qt import QtCore, QtGui
from .ui.dialog import Ui_Dialog

# standard toolkit logger
logger = sgtk.platform.get_logger(__name__)


def show_dialog(app_instance):
    """
    Shows the main dialog window.
    """
    # in order to handle UIs seamlessly, each toolkit engine has methods for launching
    # different types of windows. By using these methods, your windows will be correctly
    # decorated and handled in a consistent fashion by the system.

    # we pass the dialog class to this method and leave the actual construction
    # to be carried out by toolkit.
    app_instance.engine.show_dialog("Add Selected Comps to Render Queue...", app_instance, AppDialog)


class AppDialog(QtGui.QWidget):
    """
    Main application dialog window
    """
    CUSTOM_TEXT = "Custom frame range"
    COMP_TEXT = 'Comp frame range'

    def __init__(self):
        """
        Constructor
        """
        # first, call the base class and let it do its thing.
        QtGui.QWidget.__init__(self)

        # now load in the UI that was created in the UI designer
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)

        # most of the useful accessors are available through the Application class instance
        # it is often handy to keep a reference to this. You can get it via the following method:
        self._app = sgtk.platform.current_bundle()

        self.adobe = self._app.engine.adobe


        # logging happens via a standard toolkit logger
        logger.info("Launching Add to render Queue Application...")

        # via the self._app handle we can for example access:
        # - The engine, via self._app.engine
        # - A Shotgun API instance, via self._app.shotgun
        # - An Sgtk API instance, via self._app.sgtk


        self.first_frame = self._app.get_setting('default_first_frame')
        self.last_frame = self._app.get_setting('default_last_frame')

        # lastly, set up our very basic UI
        # self.ui.context.setText("Current Context: %s" % self._app.context)
        self.populate_widgets()
        self.connect_signals_and_slots()
        self.alert_box("hello")

    def get_selected_comps(self):

        comps = []
        # ToDo: this method needs to comp items. Currently it returns dictionarys that represent comp items but that isnt enough
        # ToDo: A different adobe method needs to be used to get the full comp item so that we can read/write properties on the comp


        for selected_item in self.adobe.app.project.selection:
            if selected_item['instanceof'] == 'CompItem':
                comps.append(selected_item)
        return comps

    def populate_widgets(self):
        self.populate_frame_range()
        self.popular_frame_range_options()
        self.populate_presets()

    def populate_frame_range(self):
        self.ui.frameRangeLineEdit.setText("%d - %d"%(self.first_frame, self.last_frame))

    def popular_frame_range_options(self):
        self.ui.frameRangeComboBox.insertItems(0, [self.COMP_TEXT, self.CUSTOM_TEXT])
        self.ui.frameRangeLineEdit.setEnabled(False)

    def populate_presets(self):
        self.presets = {}
        for preset_item in self._app.get_setting("render_presets"):
            # use an internal method to resolve the path of the ae template files
            resolved_path = self._app._TankBundle__resolve_hook_expression(preset_item['name'], preset_item['path'])
            self.presets[preset_item['name']]=resolved_path
            self.ui.renderFormatDropdown.insertItems(-1,[preset_item['name']])


    def connect_signals_and_slots(self):
        self.ui.frameRangeComboBox.currentIndexChanged.connect(self.refresh_frame_range)
        self.ui.cancelButton.clicked.connect(self.close)
        self.ui.addButton.clicked.connect(self.create_render_queue_items)

    def refresh_frame_range(self):
        if self.ui.frameRangeComboBox.currentText() == self.CUSTOM_TEXT:
            self.ui.frameRangeLineEdit.setEnabled(True)
        else:
            self.ui.frameRangeLineEdit.setEnabled(False)

    def create_render_queue_items(self):
        selected_comps = self.get_selected_comps()

        if len(selected_comps) == 0:
            self.alert_box("No comps selected")
        for comp in selected_comps:
            frame_range = self.get_frame_range(comp)
            if frame_range[0] is None or frame_range[1] is None:
                self.alert_box("Bad frame range")
            self.create_render_queue_item_for_comp()

    def get_frame_range(self, comp_index):

        first_frame = None
        last_frame = None
        custom_first_frame = None
        custom_last_frame = None
        range_string = self.ui.frameRangeLineEdit.text()
        first_frame_str = range_string.split('-')[0].strip()
        if first_frame_str.isdigit():
            custom_first_frame = int(first_frame_str)
            last_frame_str = range_string.split('-')[-1].strip()
        if last_frame_str.isdigit():
            custom_last_frame = int(last_frame_str)

        # toDo: the current method for returning comp items does not give us the frame range data
        # comp_first_frame = comp.displayStartFrame

        if self.ui.frameRangeComboBox.currentText() == self.CUSTOM_TEXT:
            pass
            # use the custom frame range
            # make sure to not return a frame range that is outside the range of como
        else:
            pass
            # use the comp's inbuilt frame range
        return [first_frame, last_frame]

    def alert_box(self, text):

        QtGui.QMessageBox.critical(
            self,
            "!",
            str(text),
            buttons=QtGui.QMessageBox.Ok,
            defaultButton=QtGui.QMessageBox.Ok,
        )

    def create_render_queue_item_for_comp(self, comp, frame_range, render_queue_template):
        pass
        # Create a render queue item
        # for reference use: \\ldn-fs1\projects\__pipeline\software\TerritoryToolkit2Toolset\modules\ae\ae_renderItemFunctions.jsx
        # The above javascript file shows how a render queue item was created in javascript, the adobe wrapper should allow
        # the same functionality to be copied into python
