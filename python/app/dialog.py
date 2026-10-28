# Copyright (c) 2013 Shotgun Software Inc.
#
# CONFIDENTIAL AND PROPRIETARY
#
# This work is provided "AS IS" and subject to the Shotgun Pipeline Toolkit
# Source Code License included in this distribution package. See LICENSE.
# By accessing, using, copying or modifying this work you indicate your
# agreement to the Shotgun Pipeline Toolkit Source Code License. All rights
# not expressly granted therein are reserved by Shotgun Software Inc.
import getpass
import logging
import ntpath
import posixpath
import shutil
import subprocess

import sgtk
import os
import re
import time
from contextlib import contextmanager
import traceback

# by importing QT from sgtk rather than directly, we ensure that
# the code will be compatible with both PySide and PyQt.
from sgtk.platform.qt import QtCore, QtGui
from .ui.dialog import Ui_Dialog, ItemSelectionDialog
from . import afterfx

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
    WORK_AREA_TEXT = 'Work area frame range'
    SINGLE_FRAME_TEXT = 'Single frame'

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

        self.ae_version = self.adobe.app.version
        self.current_project = self._app.context.project
        self.project_id = self.current_project["id"]
        self.project_name = self.current_project["name"]
        self.project_code = None

        logger.debug(f"Current Project: {self.current_project}")
        logger.debug(f"Project ID: {self.project_id}")
        logger.debug(f"AE Version: {self.ae_version}")

        # logging happens via a standard toolkit logger
        logger.info("Launching Add to render Queue Application...")

        # via the self._app handle we can for example access:
        # - The engine, via self._app.engine
        # - A Shotgun API instance, via self._app.shotgun
        # - An Sgtk API instance, via self._app.sgtk

        # Get Settings
        self.first_frame = self._app.get_setting('default_first_frame')
        self.last_frame = self._app.get_setting('default_last_frame')
        self.deadline_defaults = self._app.get_setting('deadline_defaults')
        self.deadline_host = self._app.get_setting('deadline_host')
        self.deadline_port = self._app.get_setting('deadline_port')
        self.qsettings_ignore_list = self._app.get_setting('qsettings_ignore_list')
        self.render_preset_movie_formats = self._app.get_setting('render_preset_movie_formats')

        # After Effects script runs (AfterFX.com -ro). Both settings are optional.
        self._afterfx_path = None
        self.afterfx_timeout = self._app.get_setting('afterfx_timeout', 0) or afterfx.DEFAULT_TIMEOUT

        # Background work tracking (see is_busy / closeEvent)
        self._busy_count = 0
        self._submit_thread = None
        self._submit_worker = None
        self.deadline_progress_dialog = None
        self._submission_prep_errors = ""
        # Logger for worker threads; records are handed to the GUI thread (see GuiThreadLogRelay)
        self.thread_log = GuiThreadLogRelay(logger, parent=self)

        # Deadline variables
        self.deadline_error_message = ""
        self.ae_default_pool = None
        self.deadline_settings_initialized = False

        # lastly, set up our very basic UI
        # self.ui.context.setText("Current Context: %s" % self._app.context)
        self.populate_presets()
        self.populate_sg_fields()

        # Flags for multi-row propagation behavior.
        self._is_propagating_table_edit = False
        self._active_combo_selection_rows = []

        self.connect_signals_and_slots()

        # Create render queue items
        self.create_render_queue_items()

        # Set the window size
        self.resize(1000, self.ui.compTableWidget.sizeHint().height())

    def populate_sg_fields(self):
        """
            Get the required field info from the project
        """
        # Get the project data
        try:
            project_data = self._app.shotgun.find_one("Project", [["id", "is", self.project_id]], ["sg_ae_render_pool", 'code'])
        except Exception as e:
            logger.error(f"Error getting project data: {e}")
            project_data = {}

        self.ae_default_pool = project_data.get("sg_ae_render_pool", "none")
        self.project_code = project_data.get("code", None)

        if self.ae_default_pool == "none":
            logger.info("Project AE Render Pool Field is not set, defaulting to none")
        else:
            logger.info(f"Project AE Render Pool: {self.ae_default_pool}")

    def set_deadline_defaults(self):
        """
            Populate and Set the deadline defaults from the project config settings
        """

        self.ui.deadline_pool.setCurrentText(self.ae_default_pool)
        self.ui.deadline_secondary_pool.setCurrentText(self.deadline_defaults["secondary_pool"])
        self.ui.deadline_group.setCurrentText(self.deadline_defaults["group"])
        self.ui.deadline_frame_list.setText("%d - %d" % (self.first_frame, self.last_frame))
        self.ui.deadline_frames_per_task.setText(str(self.deadline_defaults["frames_per_task"]))
        self.ui.deadline_machine_limit.setText(str(self.deadline_defaults["machine_limit"]))
        self.ui.deadline_priority.setText(str(self.deadline_defaults["priority"]))
        self.ui.deadline_concurrent_tasks.setText(str(self.deadline_defaults["concurrent_tasks"]))
        self.ui.deadline_task_timeout.setText(str(self.deadline_defaults["task_timeout_minutes"]))

    def get_scene_name(self, sanitize=False):
        """
            Get the current scene name from the project file

            Arguments:
                sanitize (bool): If true, sanitize the scene name

            Returns:
                str: The current scene name
        """

        scene_name = self.adobe.app.project.file.name

        if not(sanitize):
            return scene_name

        else:
            # Remove Ext and version number _v###
            scene_name = os.path.splitext(scene_name)[0]

            scene_name = re.sub(r'_v\d{3}$', '', scene_name)

            return scene_name


    def get_selected_comps(self):
        """
            Get the selected comps in the project
        """
        comps = []

        # Have to search through all items in the scene because simply calling
        # project.selection returns read only objects with reduced properties
        # Loop through the items in the project

        item_collection = self.adobe.app.project.items
        for item in self._app.engine.iter_collection(item_collection):
            # Check if the item is selected and is a comp
            if item.selected and self._app.engine.is_item_of_type(item, "CompItem"):
                comps.append(item)

        return comps

    def create_table_entries(self):
        """
            Create table entries for each of the comps
        """
        # Clear the table

        self.ui.compTableWidget.setRowCount(0)

        render_queue = self.adobe.app.project.renderQueue

        # Get render queue items
        render_queue_items = self.adobe.app.project.renderQueue.items
        logger.debug("Render Queue Items: %s" % render_queue_items)

        # Check if there are any render queue items, if not no comps were selected before running the dialog
        if render_queue.numItems == 0:
            self.alert_box("Add Comp To Render Queue", "Please select some comps to add to the render queue")
            return

        # Filter out the render queue items by status
        # Should only include items that match NEEDS_OUTPUT and QUEUED
        filtered_render_queue_items = []
        for i in range(1, render_queue.numItems + 1):
            render_queue_item = self.adobe.app.project.renderQueue.item(i)
            logger.debug("Render Queue Item: %s" % render_queue_item)
            logger.debug("Render Queue Item Status: %s" % render_queue_item.status)

            # Check if the render queue item is queued or needs output
            if render_queue_item.status == self.adobe.RQItemStatus.QUEUED or render_queue_item.status == self.adobe.RQItemStatus.NEEDS_OUTPUT:
                filtered_render_queue_items.append(render_queue_item)

        # Check if any render queue items are selected
        if len(filtered_render_queue_items) == 0:
            self.alert_box("No Render Queue Items meet the criteria",
                           "Please add some render queue items to apply the changes to")
            return

        # Add the comps to the table
        for item in filtered_render_queue_items:
            self.add_table_row(item)

    def add_table_row(self, item):
        """
            Add a row to the table and populate it with the render queue item data

            :param item: The item to add to the table
        """
        # Get the current row position
        rowPosition = self.ui.compTableWidget.rowCount()

        # Insert a new row
        self.ui.compTableWidget.insertRow(rowPosition)

        ############################
        # Create Widgets
        ############################

        # Add the comp name
        comp_name = QtGui.QTableWidgetItem(item.comp.name)
        comp_name.setFlags(comp_name.flags() & ~QtCore.Qt.ItemIsEditable)
        comp_name.setData(QtCore.Qt.UserRole, item) # Store the render queue item in the user data
        self.ui.compTableWidget.setItem(rowPosition, 0, comp_name)

        # Add the status
        status = "Unknown"
        # Tooltip for the status
        tooltip = "Unknown | Possible Error"
        icon = self.ui.errorIcon

        if item.status == self.adobe.RQItemStatus.QUEUED:
            status = "Queued"
            tooltip = "Template Applied | Ready to Render"
            icon = self.ui.clearIcon

        elif item.status == self.adobe.RQItemStatus.NEEDS_OUTPUT:
            status = "Needs Output"
            tooltip = "Template Not Applied | Needs Output"
            icon = self.ui.warningIcon

        statusItem = QtGui.QTableWidgetItem(icon, "")
        statusItem.setToolTip(tooltip)
        statusItem.setFlags(statusItem.flags() & ~QtCore.Qt.ItemIsEditable)
        self.ui.compTableWidget.setItem(rowPosition, 1, statusItem)

        # Add Frame Range Input
        frameRangeLineEdit = QtGui.QLineEdit()
        self.ui.compTableWidget.setCellWidget(rowPosition, 2, frameRangeLineEdit)

        # Add the frame range ComboBox
        frameRangeComboBox = QtGui.QComboBox()
        frameRangeComboBox.insertItems(0,[self.WORK_AREA_TEXT, self.COMP_TEXT, self.SINGLE_FRAME_TEXT, self.CUSTOM_TEXT])
        frameRangeComboBox.setToolTip("Select the frame range to render")
        frameRangeComboBox.installEventFilter(self)
        self.ui.compTableWidget.setCellWidget(rowPosition, 3, frameRangeComboBox)

        # Add the render format dropdown
        renderFormatDropdown = QtGui.QComboBox()
        renderFormatDropdown.setToolTip("Select the render format template to use")
        renderFormatDropdown.installEventFilter(self)

        # Populate the render format dropdown
        if self.presets:
            renderFormatDropdown.insertItems(0, self.presets.keys())

        self.ui.compTableWidget.setCellWidget(rowPosition, 4, renderFormatDropdown)

        # Check if the current render format is one of the presets and sets the index
        # GetSettingsFormat.STRING
        logger.debug("Testing Module Settings")
        #currentTemplate = item.outputModule(1).getSettings(self.adobe.GetSettingsFormat.STRING).Name
        currentTemplate = item.outputModule(1).name
        logger.debug("Current Template: %s" % currentTemplate)
        if currentTemplate in self.presets:
            renderFormatDropdown.setCurrentIndex(renderFormatDropdown.findText(currentTemplate))

        # Add the use comp name checkbox
        useCompNameCheckBox = QtGui.QTableWidgetItem()
        useCompNameCheckBox.setCheckState(QtCore.Qt.Checked)
        useCompNameCheckBox.setToolTip("Use the comp name as the output file name")
        self.ui.compTableWidget.setItem(rowPosition, 5, useCompNameCheckBox)

        # Connect the signals and slots
        frameRangeComboBox.currentIndexChanged.connect(lambda: self.refresh_frame_range(frameRangeComboBox, frameRangeLineEdit, item))

        # Propagation signals
        frameRangeComboBox.currentIndexChanged[str].connect(
            lambda text, w=frameRangeComboBox, c=3: self._on_combo_changed_propagate(w, text, c)
        )
        renderFormatDropdown.currentIndexChanged[str].connect(
            lambda text, w=renderFormatDropdown, c=4: self._on_combo_changed_propagate(w, text, c)
        )
        frameRangeLineEdit.textEdited.connect(
            lambda text, w=frameRangeLineEdit, c=2: self._on_line_edit_edited(w, text, c)
        )

        # Trigger the signal to set the default frame range
        frameRangeComboBox.emit(QtCore.SIGNAL("currentIndexChanged(int)"), 0)

        # Include the render queue item checkbox
        includeCheckBox = QtGui.QTableWidgetItem()
        includeCheckBox.setCheckState(QtCore.Qt.Checked)
        includeCheckBox.setToolTip("Include this item in the render queue")
        self.ui.compTableWidget.setItem(rowPosition, 6, includeCheckBox)

    def populate_presets(self, widget=None):
        """
            Populate the render format dropdown with the available
            render presets from the app settings.

            :param widget: The widget to populate with the presets, if None,
        """
        self.presets = {}
        for preset_item in self._app.get_setting("render_presets"):
            # Check if path is a valid path
            resolved_path = (preset_item['name'], preset_item['path'])

            # Check if the path contains a hook expression that needs to be resolved
            if re.search(r'\{.*?\}', preset_item['path']):
                logger.debug("Preset path must be resolved: %s" % preset_item['path'])
                # use an internal method to resolve the path of the ae template files
                resolved_path = self._app._TankBundle__resolve_hook_expression(preset_item['name'], preset_item['path'])

                self.presets[preset_item['name']] = resolved_path[0]
            else:
                # When there are no tokens to resolve, just use the path as is
                # it should be a valid file path when it's a custom preset
                self.presets[preset_item['name']] = preset_item['path']
                logger.debug("Preset path %s already exists" % preset_item['path'])

            if widget:
                widget.insertItems(-1, [preset_item['name']])

    def connect_signals_and_slots(self):
        """
            Connect the signals and slots
        """
        # Connect the buttons
        self.ui.submitButton.clicked.connect(self.apply_and_submit)
        self.ui.renderButton.clicked.connect(self.render_current_queue_items)
        self.ui.addButton.clicked.connect(self.create_render_queue_items)
        self.ui.applyButton.clicked.connect(self.apply_to_render_queue_items)
        self.ui.refreshButton.clicked.connect(self.create_table_entries)
        self.ui.clearButton.clicked.connect(self.clear_render_queue_items)
        self.ui.cancelButton.clicked.connect(self.close)
        self.ui.deadlinePanel.toggle_button.clicked.connect(self.populate_deadline_settings)

        # Connect the deadline settings
        #Machine List
        self.ui.deadline_machine_list.button.clicked.connect(self.show_machine_list_dialog)
        self.ui.deadline_machine_list.button.setToolTip("Select the machines to submit to")

        # Limit Groups
        self.ui.deadline_limits.button.clicked.connect(self.show_limit_group_list_dialog)
        self.ui.deadline_limits.button.setToolTip("Select the limit groups to submit to")

        # Connect custom actions, pass the table cell widget user data from column 0
        self.ui.removeCompAction.triggered.connect(self.remove_comp)
        self.ui.jumpToCompAction.triggered.connect(self.jump_to_comp)
        self.ui.removeSelectedCompsAction.triggered.connect(self.remove_selected_comps)
        self.ui.matchSelectedCompsAction.triggered.connect(self.match_selected_to_current_row)
        self.ui.refreshAction.triggered.connect(self.create_table_entries)

        # Propagation signals
        self.ui.compTableWidget.itemChanged.connect(self._on_table_item_changed)

    def refresh_frame_range(self, frameRangeComboBox, frameRangeLineEdit, renderQueueItem):
        """
            Enable the frame range line edit if the custom option is selected

            Arguments:
                frameRangeComboBox: The frame range combo box
                frameRangeLineEdit: The frame range line edit
                renderQueueItem: The render queue item
        """
        if frameRangeComboBox.currentText() == self.CUSTOM_TEXT:
            # Clear the text
            frameRangeLineEdit.clear()

            # Remove Validation
            frameRangeLineEdit.setValidator(None)

            # Enable the frame range line edit
            frameRangeLineEdit.setEnabled(True)

            # set the text to the default frame range
            frameRangeLineEdit.setText("%d - %d" % (self.first_frame, self.last_frame))

        elif frameRangeComboBox.currentText() == self.SINGLE_FRAME_TEXT:

            # Clear the text
            frameRangeLineEdit.clear()

            # Remove Validation and set the validator to only allow digits
            frameRangeLineEdit.setValidator(None)
            frameRangeLineEdit.setValidator(QtGui.QIntValidator())

            # Enable the frame range line edit
            frameRangeLineEdit.setEnabled(True)

            # set the text to the default frame range
            frameRangeLineEdit.setText(str(self.first_frame))

        elif frameRangeComboBox.currentText() == self.WORK_AREA_TEXT:
            # Clear the text
            frameRangeLineEdit.clear()

            # Remove Validation
            frameRangeLineEdit.setValidator(None)

            # Disable the frame range line edit
            frameRangeLineEdit.setEnabled(False)

            # set the text to the default frame range
            startFrame = renderQueueItem.comp.workAreaStart
            endFrame = (startFrame + renderQueueItem.comp.workAreaDuration)

            # Convert to frame numbers
            startFrameNum = int(round((startFrame / renderQueueItem.comp.frameDuration))) + renderQueueItem.comp.displayStartFrame
            endFrameNum = int(round((endFrame / renderQueueItem.comp.frameDuration))) + renderQueueItem.comp.displayStartFrame
            frameRangeLineEdit.setText("%d - %d" % (startFrameNum, endFrameNum))

        elif frameRangeComboBox.currentText() == self.COMP_TEXT:
            # Clear the text
            frameRangeLineEdit.clear()

            # Remove Validation
            frameRangeLineEdit.setValidator(None)

            # Disable the frame range line edit
            frameRangeLineEdit.setEnabled(False)

            endFrame = renderQueueItem.comp.duration
            endFrameNum = int(round((endFrame / renderQueueItem.comp.frameDuration))) + renderQueueItem.comp.displayStartFrame

            # set the text to the current comp frame range
            frameRangeLineEdit.setText("%d - %d" % (renderQueueItem.comp.displayStartFrame, endFrameNum))

    def create_render_queue_items(self):
        """
            Create render queue items for each of the selected comps
            and display them in the table widget.

            This method is called when the Add button is clicked
        """
        # Add Selected Comps to Render Queue
        with self.supress_dialogs():
            self.toggle_buttons(False)
            self.adobe.app.executeCommand(self.adobe.app.findMenuCommandId("Add to Render Queue"))
            self.create_table_entries()
            self.toggle_buttons()

    def find_render_queue_item_for_row(self, row):
        """
            Fetch a fresh render queue item for a table row.

            The table only lists QUEUED / NEEDS_OUTPUT items, and rows can be removed from it,
            so table row N is not render queue item N+1 whenever the queue holds anything else
            (e.g. a DONE item from an earlier render). Use row+1 only when its comp matches the
            row, otherwise take the matching renderable item with that comp name.

            Returns:
                The render queue item, or None if it can't be found.
        """
        tableItem = self.ui.compTableWidget.item(row, 0)
        if tableItem is None:
            return None
        comp_name = tableItem.text()
        render_queue = self.adobe.app.project.renderQueue
        renderable = (self.adobe.RQItemStatus.QUEUED, self.adobe.RQItemStatus.NEEDS_OUTPUT)

        if row + 1 <= render_queue.numItems:
            candidate = render_queue.item(row + 1)
            if candidate.comp.name == comp_name and candidate.status in renderable:
                return candidate

        # Same comp may be queued more than once: count which occurrence this row is.
        occurrence = sum(
            1 for r in range(row)
            if self.ui.compTableWidget.item(r, 0) is not None
            and self.ui.compTableWidget.item(r, 0).text() == comp_name
        )
        seen = 0
        for i in range(1, render_queue.numItems + 1):
            candidate = render_queue.item(i)
            if candidate.status in renderable and candidate.comp.name == comp_name:
                if seen == occurrence:
                    return candidate
                seen += 1

        logger.warning("No render queue item found for row %s (%s)" % (row, comp_name))
        return None

    def refresh_table_item_data(self):
        """
        Refresh only the UserRole data for each table row, preserving status icons/tooltips.
        """
        for row in range(self.ui.compTableWidget.rowCount()):
            try:
                # Fetch the latest render queue item for this row
                render_queue_item = self.find_render_queue_item_for_row(row)
                # Update the UserRole data for the first column (comp name)
                tableItem = self.ui.compTableWidget.item(row, 0)
                if tableItem is not None and render_queue_item is not None:
                    tableItem.setData(QtCore.Qt.UserRole, render_queue_item)
            except Exception as e:
                logger.error(f"Failed to refresh UserRole data for row {row}: {e}")

    def apply_to_render_queue_items(self):
        """
            Apply the changes to the render queue items

            Arguments:
                None

            Returns:
                None
        """
        self.toggle_buttons(False)
        # Get the selected comps
        # Debugging time stamp for testing HH:MM:SS
        self.start_time = time.time()
        logger.debug("Start Render Queue Items Time: %s" % time.strftime("%H:%M:%S"))

        logger.debug("Applying to render queue items")
        try:
            if self.ui.compTableWidget.rowCount() == 0:
                self.alert_box("No render queue items", "Please add some render queue items to apply the changes to")
                return

            self.show_progress_bar(format_text="Applying to render queue items...")
            count = 0
            total_rows = self.ui.compTableWidget.rowCount()
            for row in range(total_rows):
                tableItem = self.ui.compTableWidget.item(row, 0)
                statusItem = self.ui.compTableWidget.item(row, 1)
                frameRangeComboBox = self.ui.compTableWidget.cellWidget(row, 3)
                renderFormatDropdown = self.ui.compTableWidget.cellWidget(row, 4)
                useCompNameCheckBox = self.ui.compTableWidget.item(row, 5)
                includeCheckBox = self.ui.compTableWidget.item(row, 6)

                compName = tableItem.text() if tableItem else f"Row {row+1}"
                progress = int((row / total_rows) * 100)
                self.show_progress_bar(format_text=f"Applying: {compName} ({row+1}/{total_rows})")
                self.update_progress_bar(progress)

                # Per-comp steps for secondary progress bar
                comp_steps = [
                    "Frame range check",
                    "Template check",
                    "Apply template",
                    "Set time span/output location"
                ]
                num_steps = len(comp_steps)
                self.show_progress_bar(format_text=f"{compName}: {comp_steps[0]}", max=num_steps, primary=False)
                self.update_progress_bar(0, primary=False)
                step_idx = 0

                if includeCheckBox.checkState() == QtCore.Qt.Checked:
                    try:
                        render_queue_item = tableItem.data(QtCore.Qt.UserRole)
                        comp = render_queue_item.comp

                        # Defensive type check for comp
                        if not hasattr(comp, 'name'):
                            logger.error(f"Row {row}: comp is not a valid comp object. Type: {type(comp)}, Value: {comp}")
                            statusItem.setIcon(self.ui.errorIcon)
                            statusItem.setToolTip("Template Not Applied | Error - Invalid comp object")
                            includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                            self.hide_progress_bar(primary=False)
                            continue

                        compName = comp.name
                        templateName = renderFormatDropdown.currentText()
                        render_queue_template = self.get_render_queue_template(row)
                        frame_range = self.get_frame_range(comp, row)

                        # Step 1: Frame range check
                        self.update_progress_bar_format(f"{compName}: {comp_steps[0]}", primary=False)
                        if frame_range[0] is None or frame_range[1] is None:
                            logger.debug("Bad frame range, skipping %s" % comp.name)
                            self.alert_box("Bad frame range", f"Please check the frame range for {comp.name}, Skipping")
                            statusItem.setIcon(self.ui.errorIcon)
                            statusItem.setToolTip("Template Not Applied | Error - Bad Frame Range")
                            includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                            self.hide_progress_bar(primary=False)
                            continue
                        step_idx += 1
                        self.update_progress_bar(step_idx, primary=False)

                        # Step 2: Template check
                        self.update_progress_bar_format(f"{compName}: {comp_steps[1]}", primary=False)

                        if not self.check_template_exists(render_queue_item, frame_range, render_queue_template, templateName):

                            self.alert_box("Error", "Something went wrong applying or locating an output template")

                            logger.debug("Error applying template %s" % templateName)
                            logger.debug(render_queue_template)
                            statusItem.setIcon(self.ui.errorIcon)
                            statusItem.setToolTip("Template Not Applied | Error - Template Not Found")
                            includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                            self.hide_progress_bar(primary=False)
                            continue

                        step_idx += 1
                        self.update_progress_bar(step_idx, primary=False)

                        # Step 3: Apply template
                        self.update_progress_bar_format(f"{compName}: {comp_steps[2]}", primary=False)
                        try:
                            try:
                                # Refetch item and comp to make sure its note a stale reference
                                render_queue_item = self.find_render_queue_item_for_row(row)
                                if render_queue_item is None:
                                    raise RuntimeError("Render queue item for '%s' not found" % compName)
                                comp = render_queue_item.comp

                            except Exception as refetch_exc:
                                logger.error("Failed to re-fetch render_queue_item: %s" % str(refetch_exc))
                                statusItem.setIcon(self.ui.errorIcon)
                                statusItem.setToolTip("Template Not Applied | Error - Could not re-fetch render_queue_item")
                                includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                                self.hide_progress_bar(primary=False)
                                continue

                            outputModuleAttr = getattr(render_queue_item, "outputModule", None)
                            if not callable(outputModuleAttr):
                                logger.error("render_queue_item.outputModule is not callable after refresh! Type: %s, Value: %s" % (type(outputModuleAttr), outputModuleAttr))
                                statusItem.setIcon(self.ui.errorIcon)
                                statusItem.setToolTip("Template Not Applied | Error - outputModule not callable after refresh")
                                includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                                self.hide_progress_bar(primary=False)
                                continue

                            if templateName in render_queue_item.outputModule(1).templates:
                                logger.debug("Applying template %s to comp %s" % (templateName, comp.name))
                                render_queue_item.outputModule(1).applyTemplate(templateName)
                            else:
                                logger.error("Template %s not found in output module templates after refresh" % templateName)
                                raise Exception("Template %s not found in output module templates after refresh" % templateName)

                        except Exception as e:
                            self.alert_box("Error",
                                           "There's some kind of issue with this template\n\n" + str(templateName) + '\n' + str(
                                               render_queue_template))
                            logger.debug("Error applying template %s" % templateName)
                            logger.debug("Exception: %s" % str(e))
                            logger.debug(traceback.format_exc())
                            statusItem.setIcon(self.ui.errorIcon)
                            statusItem.setToolTip("Template Not Applied | Error")
                            includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                            self.hide_progress_bar(primary=False)
                            continue

                        step_idx += 1
                        self.update_progress_bar(step_idx, primary=False)

                        # Step 4: Time span and output location
                        self.update_progress_bar_format(f"{compName}: {comp_steps[3]}", primary=False)
                        timespan = render_queue_item.getSetting("Time Span")
                        logger.debug("Time Span: %s" % timespan)

                        if frameRangeComboBox.currentText() == self.COMP_TEXT:
                            if timespan == 0:
                                pass
                            else:
                                render_queue_item.setSetting("Time Span", 0)
                                logger.debug("Setting time span to comp frame range")

                        elif frameRangeComboBox.currentText() == self.WORK_AREA_TEXT:
                            if timespan == 1:
                                pass
                            else:
                                render_queue_item.setSetting("Time Span", 1)
                                logger.debug("Setting time span to work area frame range")

                        elif frameRangeComboBox.currentText() == self.CUSTOM_TEXT:
                            render_queue_item.timeSpanStart = frame_range[0]
                            render_queue_item.timeSpanDuration = frame_range[1] - frame_range[0]
                            logger.debug("Setting time span to custom frame range")

                        elif frameRangeComboBox.currentText() == self.SINGLE_FRAME_TEXT:
                            render_queue_item.timeSpanStart = frame_range[0]
                            render_queue_item.timeSpanDuration = comp.frameDuration * 1
                            logger.debug("Setting time span to single frame")

                        if useCompNameCheckBox.checkState() == QtCore.Qt.Checked:
                            use_comp_name = True
                        else:
                            use_comp_name = False
                        logger.debug("Use Comp Name: %s" % use_comp_name)

                        outputLocation = self.get_shotgrid_template(render_queue_template, use_comp_name, compName)
                        folderPath = os.path.dirname(outputLocation)
                        if not os.path.exists(folderPath):
                            os.makedirs(folderPath)

                        logger.debug("Output location: %s" % outputLocation)
                        logger.debug("Output folder: %s" % folderPath)
                        logger.debug("Comp Name: %s" % comp.name)

                        if frameRangeComboBox.currentText() == self.SINGLE_FRAME_TEXT:
                            outputLocation = re.sub(r'\.\[?\#*\]?\.*', '.', outputLocation)

                        with self.supress_dialogs():
                            render_queue_item.outputModule(render_queue_item.numOutputModules).file = self.adobe.File(
                                outputLocation)
                            render_queue_item.outputModule(render_queue_item.numOutputModules).file = self.adobe.File(
                                outputLocation)

                        step_idx += 1
                        self.update_progress_bar(step_idx, primary=False)

                        logger.debug("Render Queue Item for: %s has been updated" % render_queue_item.comp.name)
                        count += 1
                        statusItem = self.ui.compTableWidget.item(row, 1)
                        statusItem.setIcon(self.ui.clearIcon)
                        statusItem.setToolTip("Template Applied | Ready to Render")

                    except Exception as item_exc:
                        logger.error(f"Exception while applying to row {row}: {item_exc}")
                        logger.error(traceback.format_exc())
                        statusItem.setIcon(self.ui.errorIcon)
                        statusItem.setToolTip("Template Not Applied | Error (row exception)")
                        includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                        self.hide_progress_bar(primary=False)
                        continue

                self.hide_progress_bar(primary=False)

            self.update_progress_bar(100)
            time.sleep(0.2)

        except Exception as main_exc:
            logger.error(f"Exception in apply_to_render_queue_items: {main_exc}")
            logger.error(traceback.format_exc())

        finally:
            self.hide_progress_bar()
            # Refresh only the UserRole data for each row, preserving status icons/tooltips
            self.refresh_table_item_data()
            self.toggle_buttons()

            logger.debug("Table item data refreshed after apply.")
            logger.debug("Finish Time: %s" % time.strftime("%H:%M:%S"))
            logger.debug("Total Time: %s" % (time.time() - self.start_time))
            logger.debug("Total Render Queue Items Updated: %s" % count)

    ####################################################################################################
    # After Effects scripts (AfterFX.com -ro)
    #
    # Anything touching self.adobe (the AE RPC bridge) or a widget must run on the GUI thread.
    # build_manifest_job() gathers everything up front on the GUI thread; run_manifest_job()
    # only writes files and runs AfterFX.com, so it is safe in a worker thread.
    ####################################################################################################
    JSX_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "jsx"))
    PROJECT_MANIFEST_JSX = "generate_manifest_for_all_comps.jsx"
    COMP_MANIFEST_JSX = "generate_manifest_from_comps.jsx"

    def get_afterfx_path(self):
        """
            The AfterFX.com matching the running After Effects (GUI thread, cached).

            Raises:
                afterfx.AfterFXError: If it can't be found.
        """
        if not self._afterfx_path:
            self._afterfx_path = afterfx.find_afterfx(self.ae_version, self._app.get_setting("afterfx_path", ""))
            logger.debug("Using AfterFX: %s" % self._afterfx_path)
        return self._afterfx_path

    def build_manifest_job(self, render_queue_item, render_scene_file_path, project_manifest=True, comp_manifest=False):
        """
            Collect what is needed to generate manifests for a render queue item. GUI thread only.

            Arguments:
                render_queue_item (RenderQueueItem): The render queue item.
                render_scene_file_path (str): The render scene the manifests describe; they are
                    written to its folder.
                project_manifest (bool): Generate project_manifest.json for all comps.
                comp_manifest (bool): Generate <comp>_manifest.json for this comp.

            Returns:
                dict: A plain-data job for run_manifest_job().
        """
        comp = render_queue_item.comp
        comp_identifiers = [{
            "name": comp.name,
            "id": getattr(comp, "id", None),
            "output_location": os.path.dirname(render_scene_file_path),
        }]

        scripts = []
        if project_manifest:
            scripts.append(os.path.join(self.JSX_DIR, self.PROJECT_MANIFEST_JSX))
        if comp_manifest:
            scripts.append(os.path.join(self.JSX_DIR, self.COMP_MANIFEST_JSX))

        return {
            "afterfx_path": self.get_afterfx_path(),
            "scripts": scripts,
            "args": {"comps": comp_identifiers},
            "timeout": self.afterfx_timeout,
        }

    @staticmethod
    def run_manifest_job(job, log=None):
        """
            Run the manifest scripts of a job from build_manifest_job(). Safe in any thread
            when given a thread-safe log (see GuiThreadLogRelay).

            Raises:
                afterfx.AfterFXError: If a script fails, never runs or times out.
        """
        for script_path in job["scripts"]:
            afterfx.run_script(job["afterfx_path"], script_path, job["args"], timeout=job["timeout"], log=log)

    def run_blocking_in_thread(self, func):
        """
            Run func() in a worker thread while keeping the UI responsive, and wait for it.
            func must not touch widgets or self.adobe. Its exception, if any, is re-raised here.

            Closing the window is refused while this runs (see closeEvent), so the thread
            can't outlive the widgets.
        """
        outcome = {"error": None, "traceback": ""}

        def _call():
            try:
                func()
            except Exception as exc:
                # No logging here: this runs in the worker thread.
                outcome["error"] = exc
                outcome["traceback"] = traceback.format_exc()

        thread = QtCore.QThread(self)
        worker = _CallableWorker(_call)
        worker.moveToThread(thread)
        loop = QtCore.QEventLoop(self)

        thread.started.connect(worker.run)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(loop.quit)

        self._busy_count += 1
        try:
            thread.start()
            loop.exec_()
        finally:
            thread.wait()
            thread.deleteLater()
            self._busy_count -= 1

        if outcome["error"] is not None:
            logger.error(outcome["traceback"])
            raise outcome["error"]

    def process_queue_items_for_render(self):
        """
            Process the render queue items for rendering locally
            Updates the status of the render queue items in the table and creates a copy
            of the project file for each render queue item to be rendered into the templated location.

            Arguments:
                None

            Returns:
                Boolean: True if successful, False if not
        """

        self.toggle_buttons(False)

        project_check = self.run_project_checks()
        if not project_check:
            return False

        # If there are rows in the table
        if self.ui.compTableWidget.rowCount() == 0:
            self.warning_box("No render queue items available",
                             "Please add some render queue items or refresh the table to Render the current items")
            return False

        logger.debug("All Checks passed, proceeding with Render")

        save_project = self.question_box(
            "Save Project",
            "Saving is required before rendering. Do you want to save the current project?",
        )

        if save_project:
            self.adobe.app.project.save()
            logger.info('Project saved')
        else:
            return False

        num_items = self.ui.compTableWidget.rowCount()
        steps_per_item = 2  # manifest + copy
        total_steps = num_items * steps_per_item
        current_step = 0

        self.show_progress_bar(format_text="Processing render queue items...")

        count = 0

        for row in range(num_items):
            tableItem = self.ui.compTableWidget.item(row, 0)
            statusItem = self.ui.compTableWidget.item(row, 1)
            renderFormatDropdown = self.ui.compTableWidget.cellWidget(row, 4)
            useCompNameCheckBox = self.ui.compTableWidget.item(row, 5)
            includeCheckBox = self.ui.compTableWidget.item(row, 6)

            render_queue_item = tableItem.data(QtCore.Qt.UserRole)
            comp = render_queue_item.comp
            compName = comp.name

            if includeCheckBox.checkState() == QtCore.Qt.Checked and render_queue_item.status != self.adobe.RQItemStatus.NEEDS_OUTPUT:
                render_queue_template = self.get_render_queue_template(row)
                frame_range = self.get_frame_range(comp, row)
                use_comp_name = useCompNameCheckBox.checkState() == QtCore.Qt.Checked

                logger.debug("Use Comp Name: %s" % use_comp_name)

                # Grab the output folder from templates
                render_scene_file_path = self.get_shotgrid_template(render_queue_template,
                                                                    use_comp_name,
                                                                    compName,
                                                                    True)

                # Create the output folder if it doesn't already exist
                render_scene_file_directory = os.path.dirname(render_scene_file_path)
                if not os.path.exists(render_scene_file_directory):
                    os.makedirs(render_scene_file_directory)

                # Debugging
                logger.debug("Render Scene File: %s" % render_scene_file_path)
                logger.debug("Render Scene File Directory: %s" % render_scene_file_directory)
                logger.debug("Comp Name: %s" % comp.name)

                # Manifest generation step
                current_step += 1

                self.update_progress_bar(int(current_step / total_steps * 100))

                try:
                    logger.debug("Generating project manifest file...")

                    self.update_progress_bar_format(f"Generating project manifest for {compName}...")
                    # Gather on the GUI thread, then run AfterFX.com off it so the UI stays live.
                    manifest_job = self.build_manifest_job(render_queue_item, render_scene_file_path)
                    self.run_blocking_in_thread(
                        lambda job=manifest_job: self.run_manifest_job(job, log=self.thread_log))
                    logger.debug("Project manifest file generated for render queue item: %s" % compName)

                except Exception as e:
                    logger.error("Failed to generate manifest file: %s" % e)
                    error = traceback.format_exc()
                    logger.error("%s" % error)
                    self.deadline_error_message += "Failed to generate manifest file: %s\n" % compName

                    # Update the status icon
                    statusItem.setIcon(self.ui.errorIcon)
                    statusItem.setToolTip("Error - Failed to generate manifest file")
                    includeCheckBox.setCheckState(QtCore.Qt.Unchecked)

                    continue

                self.update_progress_bar(int(current_step / total_steps * 100))

                # Project file copy step
                current_step += 1
                self.update_progress_bar_format(f"Copying project file for {compName}...")
                self.update_progress_bar(int(current_step / total_steps * 100))
                try:
                    logger.debug("Copying project file to render scene location: %s" % render_scene_file_path)
                    shutil.copy(self.adobe.app.project.file.fsName, render_scene_file_path)
                    logger.info('Copy created: %s' % render_scene_file_path)

                except Exception as e:
                    logger.error("Failed to create render scene backup: %s" % e)
                    error = traceback.format_exc()
                    logger.error("%s" % error)

                    self.alert_box("Error", "Failed to create render scene backup:\n\n%s" % e)

                    # Update the status icon
                    statusItem.setIcon(self.ui.errorIcon)
                    statusItem.setToolTip("Local Render | Error - Failed to create render scene backup")
                    includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                    continue

                self.update_progress_bar(int(current_step / total_steps * 100))

                # Log
                logger.debug("Render Queue Item for: %s has been processed" % render_queue_item.comp.name)
                count += 1

                # Update the status icon
                statusItem = self.ui.compTableWidget.item(row, 1)
                statusItem.setIcon(self.ui.clearIcon)
                statusItem.setToolTip("Local Render | Ready to Render - Render Scene Created")

            else:
                # Change the render queue item status to UNQUEUED if unchecked
                render_queue_item.render = False # Set to unqueued
                logger.debug("Render Queue Item for: %s has been set to UNQUEUED" % render_queue_item.comp.name)

                # Update the status icon
                statusItem = self.ui.compTableWidget.item(row, 1)
                statusItem.setIcon(self.ui.warningIcon)

                if render_queue_item.status == self.adobe.RQItemStatus.NEEDS_OUTPUT:
                    statusItem.setToolTip("NEEDS OUTPUT - Template Not Applied | Needs Output")
                else:
                    statusItem.setToolTip("UNQUEUED - Not Included in Render Queue")

        # Show progress at 100%
        self.update_progress_bar(100)
        self.update_progress_bar_format("Finished processing render queue items")
        time.sleep(0.2)
        self.hide_progress_bar()

        self.toggle_buttons()
        logger.debug("Total Render Queue Items Processed: %s" % count)

        return True

    def clear_render_queue_items(self):
        """
            Clear all render queue items from the After Effects render queue that have been completed
            This is to help clean up the render

        """
        logger.debug("Clearing completed render queue items")

        renderQueue = self.adobe.app.project.renderQueue

        with self.supress_dialogs():
            count = 0

            for i in range(renderQueue.numItems, 0, -1):
                item = renderQueue.item(i)
                if item.status == self.adobe.RQItemStatus.DONE or item.status == self.adobe.RQItemStatus.USER_STOPPED:
                    logger.debug("Removing render queue item: %s" % item.comp.name)
                    item.remove()
                    count += 1

        logger.debug("Clear Render Queue Items Processed: %s" % count)

        # Refresh Table
        self.create_table_entries()

        self.message_box("Clear Render Queue Items", "Cleared %s completed render queue items from the render queue." % count)

    def get_frame_range(self, comp, row):
        """
            Get the frame range to render

            Arguments:
                comp: The comp to get the frame range for
                row: The row in the table

            Returns:
                tuple: (startFrame, endFrame)
        """
        startFrame = None
        endFrame = None

        # Debug Info Report for Comp
        try:
            logger.debug("*" * 50)
            logger.debug(" Debug Info Report for Comp")
            logger.debug("*" * 50)
            logger.debug("Comp Name: %s" % comp.name)
            logger.debug("Comp Frame Rate: %s" % comp.frameRate)
            logger.debug("Comp Frame Duration: %s" % comp.frameDuration)
            logger.debug("Comp Display Start Frame: %s" % comp.displayStartFrame)
            logger.debug("Comp Display Start Time: %s" % comp.displayStartTime)
            logger.debug("Comp Duration: %s" % comp.duration)
            logger.debug("Comp Work Area Start: %s" % comp.workAreaStart)
            logger.debug("Comp Work Area Duration: %s" % comp.workAreaDuration)
            logger.debug("*" * 50)
        except Exception as e:
            logger.debug("Failed to get debug info for comp: %s" % e)

        # Use comp frame range (This is purely for debugging purposes)
        if self.ui.compTableWidget.cellWidget(row, 3).currentText() == self.COMP_TEXT:
            logger.debug("Using comp frame range")
            # Get the start and end frame
            startFrame = 0
            endFrame = comp.duration

            ############################
            # Debugging info
            ############################
            logger.debug("Start Time: %s" % startFrame)
            logger.debug("End Time: %s" % endFrame)

            # Convert to frame numbers
            # Check if the comp work area has a start frame of 0
            logger.debug("Checking if start time is 0")
            if int(startFrame) == 0:
                startFrameNum = comp.displayStartFrame

            else:
                startFrameNum = int(round((startFrame / comp.frameDuration))) + comp.displayStartFrame

            # Start frame
            logger.debug("Start Frame: %s" % startFrameNum)

            # End frame
            endFrameNum = int(round((endFrame / comp.frameDuration))) + comp.displayStartFrame
            logger.debug("End Frame: %s" % endFrameNum)

            #endFrame = int(comp.frameRate * comp.workAreaDuration)
            #endFrame = int(comp.frameRate * comp.duration + 0.0001)

        # Use work area frame range (This is purely for debugging purposes)
        elif self.ui.compTableWidget.cellWidget(row, 3).currentText() == self.WORK_AREA_TEXT:
            logger.debug("Using work area frame range")
            startFrame = comp.workAreaStart
            endFrame = (startFrame + comp.workAreaDuration)

            ############################
            # Debugging info
            ############################
            logger.debug("Start Time: %s" % startFrame)
            logger.debug("End Time: %s" % endFrame)

            # Convert to frame numbers
            startFrameNum = int(round((startFrame / comp.frameDuration))) + comp.displayStartFrame
            endFrameNum = int(round((endFrame / comp.frameDuration))) + comp.displayStartFrame

            # Start frame
            logger.debug("Start Frame: %s" % startFrameNum)

            # End frame
            logger.debug("End Frame: %s" % endFrameNum)

        # Use custom frame range
        elif self.ui.compTableWidget.cellWidget(row, 3).currentText() == self.CUSTOM_TEXT:

            rawText = self.ui.compTableWidget.cellWidget(row, 2).text()
            # Assumed pattern is {Digits}{NonDigitSeperator}{Digits} - e.g. 1001-1002
            match = re.match(r'(\d+)(\D+)(\d+)', rawText)
            logger.debug("Using custom frame range: %s" % rawText)
            if match:
                startFrame = match.group(1)
                endFrame = match.group(3)

                # Change frame number to Time and calculate the start and end time durations according to the start time
                logger.debug("Start Frame: %s" % startFrame)
                startFrame = (comp.frameDuration * int(startFrame)) - comp.displayStartTime
                logger.debug("Start Time: %s" % startFrame)

                logger.debug("End Frame: %s" % endFrame)
                endFrame = (comp.frameDuration * int(endFrame)) - comp.displayStartTime + 0.0001 # Add a small amount to ensure the last frame is included to avoid rounding errors
                logger.debug("End Time: %s" % endFrame)

        # Use single frame
        elif self.ui.compTableWidget.cellWidget(row, 3).currentText() == self.SINGLE_FRAME_TEXT:

            logger.debug("Using single frame range")
            rawText = self.ui.compTableWidget.cellWidget(row, 2).text()
            logger.debug("Single Frame: %s" % rawText)

            # Assumed pattern is {Digits} - e.g. 1001
            match = re.match(r'(\d+)', rawText)
            if match:
                startFrame = int(match.group(1))

                # Convert to frame number
                startFrame = (comp.frameDuration * int(startFrame)) - comp.displayStartTime
                endFrame = startFrame

                # Debugging info
                logger.debug("Start Frame: %s" % startFrame)
                logger.debug("End Frame: %s" % endFrame)

        return [startFrame, endFrame]

    def get_render_queue_template(self, row):
        """
            Get the render queue template to use for the render queue item

            Arguments:
                row: The row in the table

            Returns:
                The render queue template to use for the render queue item
        """
        render_queue_template = None

        userSelection = self.ui.compTableWidget.cellWidget(row, 4).currentText()
        if userSelection in self.presets:
            render_queue_template = self.presets[userSelection]

        return render_queue_template

    def render_current_queue_items(self):
        """
            Process and render the current render queue items

            Arguments:
                None

            returns: None
        """
        try:
            ready_to_render = self.process_queue_items_for_render()
        except Exception as e:
            logger.error(traceback.format_exc())
            self.hide_progress_bar()
            self.alert_box("Render", "Failed to prepare the render:\n\n%s" % e)
            ready_to_render = False

        if not ready_to_render:
            self.toggle_buttons()
            return

        # Bring Render Queue to front
        self.adobe.app.project.renderQueue.showWindow(True)

        # Render the current render queue items
        logger.debug("Starting render of current render queue items")
        self.adobe.app.project.renderQueue.renderAsync()

    def run_project_checks(self):
        """
            Runs basic project checks

            Arguments:
                None

            Returns: True if the project checks pass, False otherwise
        """
        logger.debug("Running project checks")

        # If there is an active project
        if not self.adobe.app.project:
            self.warning_box("No project open", "Please open a project before rendering")
            return False

        # If the project is not saved
        if not self.adobe.app.project.file:
            self.warning_box("Project not saved", "Please save the project before rendering")
            return False

        # If there are no render queue items
        if self.adobe.app.project.renderQueue.numItems == 0:
            self.warning_box("No render queue items", "Please add some render queue items to render")
            return False

        logger.debug("Project Checks passed, proceeding")

        return True

    ####################################################################################################
    # Dialog parenting
    #
    # Every dialog/message box this app creates is parented to the Toolkit window that
    # engine.show_dialog() built around this widget (or, before that window exists, to the
    # engine's dialog parent, the proxy window tk-aftereffects attaches to After Effects).
    # Never to None and never to a hidden, not-yet-parented widget: an ownerless top-level
    # window is what lets Qt / the Python process lose track of who owns a dialog, which can
    # take the Toolkit panel down with it. Widgets are also only ever created on the GUI thread.
    ####################################################################################################
    def dialog_parent(self):
        """
            The widget to parent dialogs to.
        """
        window = self.window()
        if window is not None and window is not self:
            return window
        try:
            engine_parent = self._app.engine._get_dialog_parent()
            if engine_parent is not None:
                return engine_parent
        except Exception as e:
            logger.debug("Engine dialog parent not available: %s" % e)
        return self

    @staticmethod
    def _on_gui_thread():
        app = QtCore.QCoreApplication.instance()
        return app is not None and QtCore.QThread.currentThread() == app.thread()

    def _exec_message_box(self, icon, title, text, buttons=QtGui.QMessageBox.Ok, default_button=QtGui.QMessageBox.Ok):
        """
            Show a parented, modal message box and return the clicked button.
            Off the GUI thread it only logs (creating widgets there crashes Qt).
        """
        if not self._on_gui_thread():
            logger.error("Message box requested off the GUI thread, not shown: %s - %s" % (title, text))
            return QtGui.QMessageBox.NoButton

        box = QtGui.QMessageBox(icon, title, str(text), buttons, self.dialog_parent())
        box.setDefaultButton(default_button)
        box.setWindowModality(QtCore.Qt.WindowModal)
        try:
            return box.exec_()
        finally:
            box.deleteLater()

    def alert_box(self, title, text):
        """
            Display an alert box

            Arguments:
                title: The title of the alert box
                text: The text of the alert box

            Returns: None
        """
        logger.debug("Displaying Alert Box: %s" % text)
        self._exec_message_box(QtGui.QMessageBox.Critical, title, text)

    def warning_box(self, title, text):
        """
            Display a warning box

            Arguments:
                title: The title of the warning box
                text: The text of the warning box

            Returns: None
        """
        logger.debug("Displaying Warning Box: %s" % text)
        self._exec_message_box(QtGui.QMessageBox.Warning, title, text)

    def message_box(self, title, text):
        """
            Display a message box

            Arguments:
                title: The title of the message box
                text: The text of the message box

            Returns: None
        """
        logger.debug("Displaying Message Box: %s" % text)
        self._exec_message_box(QtGui.QMessageBox.Information, title, text)

    def question_box(self, title, text):
        """
            Ask a Yes/No question (Yes is the default).

            Returns:
                bool: True if the user clicked Yes.
        """
        result = self._exec_message_box(
            QtGui.QMessageBox.Question, title, text,
            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, QtGui.QMessageBox.Yes,
        )
        return result == QtGui.QMessageBox.Yes

    def is_busy(self):
        """
            True while background work that references this window is still running.
        """
        submit_thread = getattr(self, "_submit_thread", None)
        try:
            submit_running = submit_thread is not None and submit_thread.isRunning()
        except RuntimeError:  # C++ object already deleted
            submit_running = False
        return self._busy_count > 0 or submit_running

    def closeEvent(self, event):
        """
            Refuse to close while a worker thread is running: destroying a running QThread,
            or the widgets its signals update, aborts the whole Python/Qt process.
            The Toolkit dialog wrapper honours an ignored close event.
        """
        if self.is_busy():
            self.warning_box(
                "Please wait",
                "Still working (submitting to Deadline or generating manifests).\n"
                "Close the window once it has finished.",
            )
            event.ignore()
            return
        super(AppDialog, self).closeEvent(event)

    def get_shotgrid_template(self, render_queue_template, use_comp_name=False, comp_name=None, render_scene=False):
        """
            Get the output location from the render queue template

            Arguments:
                render_queue_template: The render queue template to use
                use_comp_name: Whether to use the comp name in the output location
                comp_name: The name of the comp to use in the output location
                render_scene: Whether to get the render scene file path instead of the output location

            Returns:
                The output location for the render queue item or render scene file path
        """
        template_file_name = os.path.basename(render_queue_template)
        logger.debug("Using template file name: %s" % template_file_name)
        # Check if the template is a movie format or treat it as a sequence
        # Decide which template to use based on the use_comp_name and render_scene flags
        # This is easier to read and maintain than nested if statements
        if any(template_file_name.startswith(fmt) for fmt in self.render_preset_movie_formats):
            template_map = {
                (True, True): "comp_render_scene_template",
                (True, False): "mov_render_comp_template",
                (False, True): "render_scene_template",
                (False, False): "mov_render_template",
            }
        else:
            template_map = {
                (True, True): "comp_render_scene_template",
                (True, False): "seq_render_comp_template",
                (False, True): "render_scene_template",
                (False, False): "seq_render_template",
            }

        templateName = self._app.get_setting(template_map[(use_comp_name, render_scene)])
        logger.debug("Using template: %s for use_comp_name: %s and render_scene: %s" % (templateName, use_comp_name, render_scene))

        template = self._app.engine.get_template_by_name(templateName)

        # Apply context as base fields
        fields = self._app.context.as_template_fields(template)

        # Grab fields from filename
        fileName = self.adobe.app.project.file.name

        # EntityName _ Name _v VersionNumber FileExtension
        match = re.match(r'(.*)(_)(.*)(_v)(\d\d\d)(.*)', fileName)
        if not match:
            raise Exception("Couldn't retrieve info from filename, try saving your scene?")

        fields['name'] = match.group(3)
        fields['version'] = int(match.group(5))
        fields['ext'] = template_file_name.split("_")[0]
        fields['ae_comp_name'] = comp_name

        # Add in a %04d number if it's a sequence then strip it out to be [####] for AE
        if 'SEQ' in template.keys:
            fields['SEQ'] = 9999
            outputPath = template.apply_fields(fields)
            outputPath = outputPath.replace('.9999.', '.[####].')

        else:
            outputPath = template.apply_fields(fields)

        return outputPath

    def check_template_exists(self, render_queue_item, frame_range, render_queue_template, templateName):
        """
            Check that the template exists, if not create it

            Arguments:
                render_queue_item: The render queue item to check the template for
                frame_range: The frame range to use for creating the template if it doesn't exist
                render_queue_template: The template to use for the render queue item
                templateName: The name of the template to check for

            Returns: True if the template exists or was created, False otherwise
        """

        # If the output module template already exists, just apply it, otherwise import the preset project, save the new template, clean up, and then apply it
        if templateName in render_queue_item.outputModule(1).templates:
            logger.debug("Template %s already exists" % templateName)
            return True

        # Import the preset project
        logger.debug("Template %s does not exist, importing preset project to create template" % templateName)
        importedProject = self.import_preset_project(render_queue_template)

        if not importedProject:
            logger.debug("Failed to import preset project for template: %s" % templateName)
            return False

        logger.debug("Preset project imported for template: %s" % templateName)

        # Get preset render queue item
        presetRenderQueueItem = self.find_render_queue_item_by_comp_name('PRESET')
        if presetRenderQueueItem is None:
            logger.debug("No preset render queue item found in imported project")
            return False

        presetRenderQueueItem.outputModule(1).saveAsTemplate(templateName)
        logger.debug("Template %s has been created" % templateName)

        importedProject.remove()
        logger.debug("Removed imported project after creating template: %s" % templateName)

        return True

    def find_render_queue_item_by_comp_name(self, comp_name):
        """
            Find a render queue item by the comp name

            Arguments:
                comp_name: The name of the comp to find

            Returns: The render queue item
        """
        renderQueueItems = self.adobe.app.project.renderQueue.items
        for i in range(1, self.adobe.app.project.renderQueue.numItems+1):
            if renderQueueItems[i].comp.name == comp_name:
                return renderQueueItems[i]

        return None

    def import_preset_project(self, render_queue_template):
        """
            Import the preset project

            Arguments:
                render_queue_template: The template to use for the render queue item

            Returns: The imported project
        """
        importProjectFolder = None

        for i in range(1, self.adobe.app.project.numItems+1):
            projectItem = self.adobe.app.project.item(i)

            if projectItem['instanceof'] == 'FolderItem':
                if projectItem.name == render_queue_template.name:
                    importProjectFolder = projectItem
                    break

        if importProjectFolder is None:
            fileObject = self.adobe.File(render_queue_template)
            importOptions = self.adobe.ImportOptions()
            importOptions.file = fileObject

            importProjectFolder = self.adobe.app.project.importFile(importOptions)

        return importProjectFolder

    def toggle_buttons(self, state=True):
        """
            Toggle the buttons to be enabled or disabled

            Arguments:
                state: Whether to enable or disable the buttons

            Returns: None
        """
        self.ui.applyButton.setEnabled(state)
        self.ui.addButton.setEnabled(state)
        self.ui.submitButton.setEnabled(state)
        self.ui.renderButton.setEnabled(state)
        self.ui.refreshButton.setEnabled(state)
        self.ui.clearButton.setEnabled(state)
        # The table too: rows are addressed by index while work runs in the background.
        self.ui.compTableWidget.setEnabled(state)

    #########################
    # Propagation Methods
    #########################
    def _snapshot_selection_for_combo(self):
        """
            Capture selected rows before combo interaction mutates table selection.
        """
        self._active_combo_selection_rows = self._get_selected_rows()

    def _effective_selected_rows(self, source_row, preferred_rows=None):
        """
            Resolve selected rows using preferred rows for interaction-specific propagation.

            Arguments:
                source_row: The row to select from
                preferred_rows: The preferred rows to select from

            Returns:
                list: The effective selected rows
        """
        preferred_rows = sorted(preferred_rows or [])
        if len(preferred_rows) > 1 and source_row in preferred_rows:
            return preferred_rows

        current_rows = self._get_selected_rows()
        if len(current_rows) > 1 and source_row in current_rows:
            return current_rows

        fallback_rows = list(current_rows)
        if source_row >= 0 and source_row not in fallback_rows:
            fallback_rows.append(source_row)
        return sorted(fallback_rows)

    def _get_selected_rows(self):
        """
            Return selected row indices as a sorted list.

            Returns:
                list: Sorted list of selected row indices.
        """
        selection_model = self.ui.compTableWidget.selectionModel()
        if selection_model is not None:
            return sorted({idx.row() for idx in selection_model.selectedRows()})

        rows = {idx.row() for idx in self.ui.compTableWidget.selectedIndexes()}
        return sorted(rows)

    def _restore_selected_rows(self, rows):
        """
            Re-apply row selection after combo edits that alter table selection.

            Arguments:
                rows: The rows to restore selection from
        """
        table = self.ui.compTableWidget
        valid_rows = sorted({row for row in (rows or []) if 0 <= row < table.rowCount()})
        if not valid_rows:
            return

        last_col = max(0, table.columnCount() - 1)
        was_blocked = table.blockSignals(True)
        try:
            table.clearSelection()
            for row in valid_rows:
                row_range = QtGui.QTableWidgetSelectionRange(row, 0, row, last_col)
                table.setRangeSelected(row_range, True)
        finally:
            table.blockSignals(was_blocked)

    def _restore_selected_rows_deferred(self, rows):
        """Restore selection after combo popup/focus handling completes."""
        rows_to_restore = list(rows or [])
        QtCore.QTimer.singleShot(0, lambda: self._restore_selected_rows(rows_to_restore))

    def _propagate_cell_change(self, source_row, col, apply_fn, preferred_rows=None):
        """
            Propagate a cell edit to all selected rows in the same column.
            `apply_fn(target_row)` applies the value to one target row.

            Arguments:
                source_row: The row to propagate cell edit to.
                apply_fn: Function to apply the cell edit to `source_row`.

            Returns:
                None
        """
        if self._is_propagating_table_edit:
            return

        selected_rows = self._effective_selected_rows(source_row, preferred_rows=preferred_rows)

        # Only propagate when editing a selected row and multi-row selection exists.
        if len(selected_rows) <= 1 or source_row not in selected_rows:
            return

        self._is_propagating_table_edit = True
        try:
            for row in selected_rows:
                if row == source_row:
                    continue
                apply_fn(row)
        finally:
            self._is_propagating_table_edit = False

    def _on_line_edit_edited(self, source_widget, value_text, col):
        """
            Handle line edit edits and propagate changes to selected rows.

            Arguments:
                col: The column to propagate cell edit to.
                value_text: The value to edit.
        """
        source_row = self._row_for_cell_widget(source_widget)
        if source_row < 0:
            return

        def _apply(row):
            line_edit = self.ui.compTableWidget.cellWidget(row, col)
            if line_edit is not None:
                line_edit.setText(value_text)

        self._propagate_cell_change(source_row, col, _apply)

    def _row_for_cell_widget(self, widget):
        """
            Resolve current row for a cell widget.

            Arguments:
                widget: The cell widget to resolve.

            Returns:
                int: The row index of the cell widget, or -1 if not found.
        """
        table = self.ui.compTableWidget
        for row in range(table.rowCount()):
            for col in (2, 3, 4):
                if table.cellWidget(row, col) is widget:
                    return row
        return -1

    def _on_combo_changed_propagate(self, source_widget, value_text, col):
        """
            Propagate combo value in `col` from the edited row to other selected rows.

            Arguments:
                source_widget: The cell widget to edit.
                value_text: The value to edit.
                col: The column to propagate cell edit to.
        """
        if self._is_propagating_table_edit:
            return

        source_row = self._row_for_cell_widget(source_widget)
        if source_row < 0:
            return

        def _apply(row):
            """
                Apply `value_text` to `row`.
            """
            combo = self.ui.compTableWidget.cellWidget(row, col)
            if combo is None:
                return

            idx = combo.findText(value_text)
            if idx >= 0 and combo.currentIndex() != idx:
                combo.setCurrentIndex(idx)

        rows_to_restore = list(self._active_combo_selection_rows or [])
        if source_row not in rows_to_restore:
            rows_to_restore.append(source_row)
        rows_to_restore = sorted(set(rows_to_restore))

        self._propagate_cell_change(
            source_row,
            col,
            _apply,
            preferred_rows=rows_to_restore,
        )

        # Keep the original multi-row selection visible after combo commit.
        self._restore_selected_rows_deferred(rows_to_restore)
        self._active_combo_selection_rows = []

    def _on_table_item_changed(self, item):
        """
            Handle table item changes and propagate changes to selected rows.

            Arguments:
                item: The item that was changed.

            Returns:
                None
        """

        if self._is_propagating_table_edit:
            return

        col = item.column()
        # TODO: Make Sure to check this each time the columns change
        # Checkbox columns are QTableWidgetItem-based (5, 6)
        if col not in (5, 6):
            return

        state = item.checkState()
        source_row = item.row()

        def _apply(row):
            target = self.ui.compTableWidget.item(row, col)
            if target is not None:
                target.setCheckState(state)

        self._propagate_cell_change(source_row, col, _apply)
    #####################################################################################################
    # Events
    #####################################################################################################
    def eventFilter(self, obj, event):
        """
        Event filter to handle specific events for QComboBox and QWheelEvent.

        Arguments:
            obj: The object that received the event.
            event: The event that was received.
        Returns:
            bool: True if the event is handled, False otherwise.
        """
        if isinstance(obj, QtGui.QComboBox):
            if event.type() == QtCore.QEvent.MouseButtonPress:
                # Snapshot selection before the combo interaction can change row selection.
                self._snapshot_selection_for_combo()
            if event.type() == QtCore.QEvent.Wheel:
                return True  # Ignore the wheel event
        return super(AppDialog, self).eventFilter(obj, event)

    ####################################################################################################
    # Context Menu Actions
    ####################################################################################################
    def get_row_from_cursor(self):
        """
            Get the row under the mouse cursor

            Arguments:
                None
            Returns:
                int: The row under the mouse cursor
        """

        pos = QtGui.QCursor.pos()
        table_pos = self.ui.compTableWidget.viewport().mapFromGlobal(pos)
        current_row = self.ui.compTableWidget.rowAt(table_pos.y())

        logger.debug("Returning Current Row: %s" % current_row)
        return current_row

    def jump_to_comp(self):
        """
            Jump to the comp in the project
        """
        row = self.get_row_from_cursor()
        item = self.ui.compTableWidget.item(row, 0) # Get the item in the first column of the row

        if item:
            render_queue_item = item.data(QtCore.Qt.UserRole)

            # Jump to the comp
            logger.debug("Jumping to comp: %s" % render_queue_item.comp.name)
            render_queue_item.comp.openInViewer()

    def remove_comp(self):
        """
            Remove the comp from the render queue
            offset by -1 because the table is 0 indexed and the render queue is 1 indexed.
            This is still a bit of a hack, but it works for now and will be revised at in the future

        """
        row = self.get_row_from_cursor()
        if row == -1:
            logger.debug("Row is invalid")
            return

        item = self.ui.compTableWidget.item(row-1, 0)  # Get the item in the first column of the row

        if item:
            render_queue_item = item.data(QtCore.Qt.UserRole)

            # Remove the comp
            logger.debug("Removing comp: %s" % render_queue_item.comp.name)
            render_queue_item.remove()
            self.ui.compTableWidget.removeRow(row-1)

    def remove_selected_comps(self):
        """
            Remove the selected comps from the render queue
        """
        # Get the selected rows
        selected_rows = self.ui.compTableWidget.selectionModel().selectedRows()

        # Check if any rows are selected
        if not selected_rows:
            self.alert_box("No comps selected", "Please select some comps to remove")
            return

        logger.debug("Removing selected comps: %s" % selected_rows)
        # Remove the selected rows
        for row in selected_rows:
            item = self.ui.compTableWidget.item(row.row(), 0)
            render_queue_item = item.data(QtCore.Qt.UserRole)

            render_queue_item.remove()
            self.ui.compTableWidget.removeRow(row.row())

        logger.debug("Comps removed")

    def match_selected_to_current_row(self):
        """
            Match the selected rows to the row under the mouse cursor

            Returns:
                None
        """

        current_row = self.get_row_from_cursor()

        # Get current options for the current row
        current_frame_range = self.ui.compTableWidget.cellWidget(current_row, 3).currentText()

        current_frame_range_text = ""
        if current_frame_range == self.CUSTOM_TEXT or current_frame_range == self.SINGLE_FRAME_TEXT:
            current_frame_range_text = self.ui.compTableWidget.cellWidget(current_row, 2).text()

        current_render_format = self.ui.compTableWidget.cellWidget(current_row, 4).currentText()
        current_use_comp_name = self.ui.compTableWidget.item(current_row, 5).checkState()
        current_include = self.ui.compTableWidget.item(current_row, 6).checkState()

        # Get the selected rows
        selected_rows = self.ui.compTableWidget.selectionModel().selectedRows()

        # Check if any rows are selected
        if not selected_rows:
            self.alert_box("No comps selected", "Please select some comps to apply the changes to")
            return

        logger.debug("Updating selected rows to match current row: %s" % current_row)
        # Apply the changes to the selected rows
        for row in selected_rows:
            if row.row() != current_row:  # Avoid self-updating
                logger.debug("Updating row: %s" % row.row())
                # Set the frame range
                frame_range = self.ui.compTableWidget.cellWidget(row.row(), 3)
                frame_range.setCurrentIndex(frame_range.findText(current_frame_range))

                # Set the frame range text
                if current_frame_range == self.CUSTOM_TEXT or current_frame_range == self.SINGLE_FRAME_TEXT:
                    frame_range_text = self.ui.compTableWidget.cellWidget(row.row(), 2)
                    frame_range_text.setText(current_frame_range_text)

                # Set the render format
                render_format = self.ui.compTableWidget.cellWidget(row.row(), 4)
                render_format.setCurrentIndex(render_format.findText(current_render_format))

                # Set the use comp name checkbox
                use_comp_name = self.ui.compTableWidget.item(row.row(), 5)
                use_comp_name.setCheckState(current_use_comp_name)

                # Set the include checkbox
                include = self.ui.compTableWidget.item(row.row(), 6)
                include.setCheckState(current_include)

    ####################################################################################################
    # Context Manager
    ####################################################################################################
    @contextmanager
    def supress_dialogs(self):
        """
            Suppress dialogs

            Usage:
                with self.supress_dialogs():
                    # Code that may trigger dialogs
        """

        try:
            logger.debug("Suppressing dialogs")
            self.adobe.app.beginSuppressDialogs()
            yield
        finally:
            logger.debug("Ending dialog suppression")
            self.adobe.app.endSuppressDialogs(False)

    #####################################################################################################
    # Progress Bar
    #####################################################################################################

    def show_progress_bar(self, format_text=None, max=100,primary=True):
        """
            Show the progress bar

            Arguments:
                format_text: The text to display on the progress bar
                primary: Whether to show the primary progress bar or the secondary progress bar

            Returns: None
        """
        if primary:
            self.ui.progressBar.setVisible(True)
            self.ui.progressBar.setMaximum(max)
            self.ui.progressBar.setValue(0)
            if format_text:
                self.ui.progressBar.setFormat(f"{format_text} %p%")
        else:
            self.ui.secondaryProgressBar.setVisible(True)
            self.ui.secondaryProgressBar.setMaximum(max)
            self.ui.secondaryProgressBar.setValue(0)
            if format_text:
                self.ui.secondaryProgressBar.setFormat(f"{format_text} %p%")

    def hide_progress_bar(self, primary=True):
        """
            Hide the progress bar
        """
        if primary:
            self.ui.progressBar.setVisible(False)
            self.ui.progressBar.setValue(0)
            self.ui.progressBar.setFormat("%p%")
        else:
            self.ui.secondaryProgressBar.setVisible(False)
            self.ui.secondaryProgressBar.setValue(0)
            self.ui.secondaryProgressBar.setFormat("%p%")


    def update_progress_bar(self, value, primary=True):
        """
            Update the progress bar
            Arguments:
                value: The value to set the progress bar to
        """
        if primary:
            self.ui.progressBar.setValue(value)
        else:
            self.ui.secondaryProgressBar.setValue(value)

    def update_progress_bar_format(self, format_text, primary=True):
        """
            Update the progress bar format
            Arguments:
                format_text: The text to display on the progress bar
        """
        if primary:
            self.ui.progressBar.setFormat(f"{format_text} %p%")
        else:
            self.ui.secondaryProgressBar.setFormat(f"{format_text} %p%")

    ######################################################################################################
    # Deadline
    ######################################################################################################
    # Deadline Settings
    def get_deadline_settings(self):
        """
            Get the current Deadline settings from the UI

            Arguments:
                None

            Returns:
                dict: A dictionary containing the current Deadline settings
        """
        try:
            # Get required settings first to ensure they are always set and replace with defaults
            # if inputs arnt valid

            priority = self.ui.deadline_priority.text()
            if priority == '':
                priority = self.deadline_defaults['priority']

            group = self.ui.deadline_group.currentText()
            if group == '':
                group = self.deadline_defaults['group']

            chunk_size = self.ui.deadline_frames_per_task.text()
            if not chunk_size or not chunk_size.isdigit() or int(chunk_size) < 1:
                chunk_size = self.deadline_defaults['frames_per_task']

            task_timeout_minutes = self.ui.deadline_task_timeout.text()
            if task_timeout_minutes == '':
                task_timeout_minutes = self.deadline_defaults['task_timeout_minutes']

            concurrent_tasks = self.ui.deadline_concurrent_tasks.text()
            if concurrent_tasks == '':
                concurrent_tasks = self.deadline_defaults['concurrent_tasks']


            deadline_settings = {
                'priority': priority,
                'pool': self.ui.deadline_pool.currentText(),
                'secondary_pool': self.ui.deadline_secondary_pool.currentText(),
                'group': group,
                'chunk_size': chunk_size,
                'frame_list': self.ui.deadline_frame_list.text(),
                'submit_scene': self.ui.deadline_submit_project_file_with_job.isChecked(),
                'override_frame_list': self.ui.deadline_use_frame_list_from_comp.isChecked(),
                'task_timeout_minutes': task_timeout_minutes,
                'concurrent_tasks': concurrent_tasks,
                'limit_groups': self.ui.deadline_limits.get_text(),
                'machine_list': self.ui.deadline_machine_list.get_text(),
                'machine_limit': self.ui.deadline_machine_limit.text() or self.deadline_defaults.get('machine_limit', ''),
                'submit_allow_list_as_deny_list': self.ui.deadline_machine_list_deny.isChecked(),
                'submit_suspended': self.ui.deadline_submit_as_suspended.isChecked(),
                'submit_entire_render_queue': self.ui.deadline_submit_entire_render_queue.isChecked(),
                'multi_machine': self.ui.deadline_multi_machine_rendering.isChecked(),
                'multi_machine_tasks': self.ui.deadline_multi_machine_number_of_machines.get_value(),
                'file_size': self.ui.deadline_minimum_file_size.get_value(),
                'delete_files': self.ui.deadline_delete_files_under_minimum_size.isChecked(),
                'memory_management': self.ui.deadline_enable_memory_management.isChecked(),
                'image_cache_percentage': self.ui.deadline_image_cache.get_value(),
                'max_memory_percentage': self.ui.deadline_maximum_memory.get_value(),
                'use_comp_frame_list': self.ui.deadline_use_frame_list_from_comp.isChecked(),
                'first_and_last': self.ui.deadline_render_first_and_last_frames.isChecked(),
                'missing_layers': self.ui.deadline_ignore_missing_layer_dependencies.isChecked(),
                'missing_effects': self.ui.deadline_ignore_missing_effect_references.isChecked(),
                'fail_on_warnings': self.ui.deadline_fail_on_warning_messages.isChecked(),
                'local_rendering': self.ui.deadline_enable_local_rendering.isChecked(),
                'include_output_path': self.ui.deadline_include_output_file_path.isChecked(),
                'fail_on_existing_ae_process': self.ui.deadline_fail_on_existing_ae_process.isChecked(),
                'fail_on_missing_file': self.ui.deadline_fail_on_missing_output.isChecked(),
                'override_fail_on_existing_ae_process': self.ui.deadline_override_fail_on_existing_ae_process.isChecked(),
                'ignore_gpu_acceleration_warning': self.ui.deadline_ignore_gpu_acceleration_warning.isChecked(),
                'multi_process': self.ui.deadline_multi_process_rendering.isChecked(),
                'export_as_xml': self.ui.deadline_export_xml_project_file.isChecked(),
                'delete_tmp_xml': self.ui.deadline_delete_xml_file_after_export.isChecked(),
                'missing_footage': self.ui.deadline_continue_on_missing_footage.isChecked(),
                'job_dependencies': self.ui.deadline_dependencies.get_text(),
                'on_job_complete': self.ui.deadline_on_job_complete.currentText(),
                'comment': self.ui.deadline_comment.text(),
                'group_submissions' : self.ui.deadline_group_submissions.isChecked(),
            }

            return deadline_settings

        except Exception as e:
            logger.error(f"Failed to get Deadline settings: {e}")
            error = traceback.format_exc()
            logger.error(error)
            return None

    def populate_deadline_settings(self):
        """
            This Method populates the pool and group lists from Deadline and sets the default values.
            It also sets the default values for the Deadline settings.
        """
        if not self.deadline_settings_initialized:
            logger.debug("Populating Deadline settings")

            # Get the pool and group lists from Deadline
            pools = self.get_pool_list()
            groups = self.get_group_list()

            # Populate the UI with the pools and groups
            self.ui.deadline_pool.addItems(pools)
            self.ui.deadline_secondary_pool.addItems(pools)
            self.ui.deadline_group.addItems(groups)
            self.ui.deadline_on_job_complete.addItems(self.deadline_defaults["on_job_complete"])

            self.deadline_settings_initialized = True
            logger.debug("Deadline settings populated")

            self.set_deadline_defaults()

            settings = self.load_deadline_qsettings()
            self.apply_deadline_qsettings_to_ui(settings)

    # Deadline Command Line Tool
    @staticmethod
    def get_deadline_command():
        """
            Get the Deadline command line tool path.
            This method checks the DEADLINE_PATH environment variable and the default installation path for Deadline.

            Returns:
                str: The path to the Deadline command line tool.
        """
        deadlineBin = ""
        try:
            deadlineBin = os.environ['DEADLINE_PATH']
        except KeyError:
            # if the error is a key error it means that DEADLINE_PATH is not set. however Deadline command may be in the PATH or on OSX it could be in the file /Users/Shared/Thinkbox/DEADLINE_PATH
            pass

        # On OSX, we look for the DEADLINE_PATH file if the environment variable does not exist.
        if deadlineBin == "" and os.path.exists("/Users/Shared/Thinkbox/DEADLINE_PATH"):
            with open("/Users/Shared/Thinkbox/DEADLINE_PATH") as f:
                deadlineBin = f.read().strip()

        deadlineCommand = os.path.join(deadlineBin, "deadlinecommand")

        return deadlineCommand

    DEADLINE_COMMAND_TIMEOUT = 60

    def call_deadline_command(self, arguments, hideWindow=True, readStdout=True):
        """
        Calls the Deadline command with specified arguments.

        Args:
            arguments (list): List of arguments to pass to the Deadline command.
            hideWindow (bool): Whether to hide the command window (Windows only).
            readStdout (bool): Whether to read the standard output.

        Returns:
            str: The standard output from the command if readStdout is True.
        """
        deadlineCommand = self.get_deadline_command()
        if not deadlineCommand:
            logger.error("Deadline command not found.")
            return ""

        startupinfo = None
        creationflags = 0
        if os.name == 'nt' and hideWindow:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

        try:
            logger.debug("Calling Deadline command: %s %s" % (deadlineCommand, arguments))
            # Timeout so an unreachable repository can't freeze the UI (this runs on the GUI thread).
            result = subprocess.run(
                [deadlineCommand] + arguments,
                startupinfo=startupinfo,
                creationflags=creationflags,
                capture_output=readStdout,
                text=True,
                timeout=self.DEADLINE_COMMAND_TIMEOUT,
            )
            if result.returncode != 0:
                logger.warning("deadlinecommand %s exited with %s: %s" % (arguments, result.returncode, result.stderr))
            if readStdout:
                return result.stdout or ""
        except subprocess.TimeoutExpired:
            logger.error("deadlinecommand %s timed out after %ss" % (arguments, self.DEADLINE_COMMAND_TIMEOUT))
        except Exception as e:
            logger.error(f"Failed to execute Deadline command: {e}")
        # Callers parse the output, so never hand back None.
        return ""

    ####################################
    # QSettings Functions
    ###################################
    def save_deadline_qsettings(self):
        """
            Save the Deadline settings to a QSettings file.
        """
        scene_name = self.get_scene_name(sanitize= True)
        settings = self.get_deadline_settings()
        qsettings = QtCore.QSettings("Territory", "AfterEffectsDeadlineSubmission")
        qsettings.setValue(f"deadline_settings/{scene_name}", settings)

    def load_deadline_qsettings(self):
        """
            Load the Deadline settings from a QSettings file.
        """
        scene_name = self.get_scene_name(sanitize=True)
        qsettings = QtCore.QSettings("Territory", "AfterEffectsDeadlineSubmission")
        settings = qsettings.value(f"deadline_settings/{scene_name}")

        return settings

    def apply_deadline_qsettings_to_ui(self, settings):
        """
            Apply the loaded QSettings to the UI elements.
            Any settings in the ignore list will not be set

            Arguments:
                settings (QSettings): The loaded QSettings object.
        """
        if not settings:
            return

        # QSettings can return numeric values for saved text fields.
        # Ensure UI text-based widgets always receive strings.
        to_text = lambda v: "" if v is None else str(v)

        setters = {
            "priority": lambda v: self.ui.deadline_priority.setText(to_text(v)),
            "pool": lambda v: self.ui.deadline_pool.setCurrentText(to_text(v)),
            "secondary_pool": lambda v: self.ui.deadline_secondary_pool.setCurrentText(to_text(v)),
            "group": lambda v: self.ui.deadline_group.setCurrentText(to_text(v)),
            "on_job_complete": lambda v: self.ui.deadline_on_job_complete.setCurrentText(to_text(v)),
            "chunk_size": lambda v: self.ui.deadline_frames_per_task.setText(to_text(v)),
            "frame_list": lambda v: self.ui.deadline_frame_list.setText(to_text(v)),
            "submit_scene": lambda v: self.ui.deadline_submit_project_file_with_job.setChecked(v),
            "override_frame_list": lambda v: self.ui.deadline_use_frame_list_from_comp.setChecked(v),
            "task_timeout_minutes": lambda v: self.ui.deadline_task_timeout.setText(to_text(v)),
            "concurrent_tasks": lambda v: self.ui.deadline_concurrent_tasks.setText(to_text(v)),
            "limit_groups": lambda v: self.ui.deadline_limits.set_text(v),
            "machine_list": lambda v: self.ui.deadline_machine_list.set_text(v),
            "submit_allow_list_as_deny_list": lambda v: self.ui.deadline_machine_list_deny.setChecked(v),
            "submit_suspended": lambda v: self.ui.deadline_submit_as_suspended.setChecked(v),
            "submit_entire_render_queue": lambda v: self.ui.deadline_submit_entire_render_queue.setChecked(v),
            "multi_machine": lambda v: self.ui.deadline_multi_machine_rendering.setChecked(v),
            "multi_machine_tasks": lambda v: self.ui.deadline_multi_machine_number_of_machines.set_value(v),
            "file_size": lambda v: self.ui.deadline_minimum_file_size.set_value(v),
            "delete_files": lambda v: self.ui.deadline_delete_files_under_minimum_size.setChecked(v),
            "memory_management": lambda v: self.ui.deadline_enable_memory_management.setChecked(v),
            "image_cache_percentage": lambda v: self.ui.deadline_image_cache.set_value(v),
            "max_memory_percentage": lambda v: self.ui.deadline_maximum_memory.set_value(v),
            "use_comp_frame_list": lambda v: self.ui.deadline_use_frame_list_from_comp.setChecked(v),
            "first_and_last": lambda v: self.ui.deadline_render_first_and_last_frames.setChecked(v),
            "missing_layers": lambda v: self.ui.deadline_ignore_missing_layer_dependencies.setChecked(v),
            "missing_effects": lambda v: self.ui.deadline_ignore_missing_effect_references.setChecked(v),
            "fail_on_warnings": lambda v: self.ui.deadline_fail_on_warning_messages.setChecked(v),
            "local_rendering": lambda v: self.ui.deadline_enable_local_rendering.setChecked(v),
            "include_output_path": lambda v: self.ui.deadline_include_output_file_path.setChecked(v),
            "fail_on_existing_ae_process": lambda v: self.ui.deadline_fail_on_existing_ae_process.setChecked(v),
            "fail_on_missing_file": lambda v: self.ui.deadline_fail_on_missing_output.setChecked(v),
            "override_fail_on_existing_ae_process": lambda v: self.ui.deadline_override_fail_on_existing_ae_process.setChecked(v),
            "ignore_gpu_acceleration_warning": lambda v: self.ui.deadline_ignore_gpu_acceleration_warning.setChecked(v),
            "multi_process": lambda v: self.ui.deadline_multi_process_rendering.setChecked(v),
            "export_as_xml": lambda v: self.ui.deadline_export_xml_project_file.setChecked(v),
            "delete_tmp_xml": lambda v: self.ui.deadline_delete_xml_file_after_export.setChecked(v),
            "missing_footage": lambda v: self.ui.deadline_continue_on_missing_footage.setChecked(v),
            "job_dependencies": lambda v: self.ui.deadline_dependencies.set_text(v),
            "comment": lambda v: self.ui.deadline_comment.setText(to_text(v)),
            "group_submissions": lambda v: self.ui.deadline_group_submissions.setChecked(v),
        }

        for key, value in settings.items():
            if key in self.qsettings_ignore_list:
                continue

            setter = setters.get(key)
            if setter:
                setter(value)

    ####################################
    # Deadline Submission
    ####################################
    @staticmethod
    def build_deadline_output_attrs(output_paths):
        """
            Build the OutputDirectoryN / OutputFilenameN job info entries Deadline uses to
            show (and let you browse to) a job's output.

            Arguments:
                output_paths (iterable of str): Full native output paths, one per output module.
                    AE frame tokens such as [####] are converted to Deadline's ####.

            Returns:
                dict: e.g. {"OutputDirectory0": "X:/renders/v001", "OutputFilename0": "comp.####.exr"}
        """
        attrs = {}
        index = 0
        for path in output_paths:
            if not path:
                continue
            path = str(path).replace('[#', '#').replace('#]', '#')
            path_module = ntpath if ("\\" in path or re.match(r"^[a-zA-Z]:", path)) else posixpath
            directory, file_name = path_module.split(path)
            if not directory or not file_name:
                logger.warning("Skipping output path without a folder or file name: %s" % path)
                continue
            attrs[f"OutputDirectory{index}"] = directory
            attrs[f"OutputFilename{index}"] = file_name
            index += 1
        return attrs

    def build_deadline_job_and_plugin_dicts(self,render_queue_item, deadline_settings, project_path, layers, previous_job_id, version, render_file):
        """
        Builds the required job_attrs and plugin_attrs dictionaries for submitting a job to DeadlineConnect based on the provided render queue item and settings.

            Arguments:
                render_queue_item (RenderQueueItem): The render queue item to submit job for.
                deadline_settings (dict): The current Deadline settings from the UI.
                project_path (str): The file path of the current After Effects project.
                layers (bool): Whether to include layer information in the plugin attributes.
                previous_job_id (str): The job ID of the previously submitted job to set as a dependency, if any.
                version (str): The version of After Effects being used.
                render_file (str): The file path of the render scene to submit.

        Returns (job_attrs, plugin_attrs) for DeadlineConnect submission.
        """
        # Output Variables
        # Note: ExtendScript File.name is only the (URI-encoded) file name, with no folder,
        # so the full native path has to come from fsName.
        output_module = render_queue_item.outputModule(render_queue_item.numOutputModules)
        output_file_path = output_module.file.fsName
        output_file_name = os.path.basename(output_file_path)
        ae_project_name = os.path.basename(self.adobe.app.project.file.name)
        job_name = f"{self.project_name} - {ae_project_name} - {render_queue_item.comp.name} - {output_module.name}"
        start_frame = ""
        end_frame = ""
        frame_list = deadline_settings.get('frame_list', "")
        override_frame_list = deadline_settings.get('override_frame_list', False)
        first_and_last = deadline_settings.get('first_and_last', False)
        multi_machine = deadline_settings.get('multi_machine', False)
        submit_scene = deadline_settings.get('submit_scene', False)
        comp_name = render_queue_item.comp.name
        dependent_job_id = previous_job_id or ""
        dependent_comps = False
        is_movie = any(output_file_name.lower().endswith("." + ext.lower()) for ext in self.deadline_defaults["movie_formats"])

        if override_frame_list or multi_machine:
            # Get the frame duration and start/end times
            frame_duration = render_queue_item.comp.frameDuration
            frame_offset = self.adobe.app.project.displayStartFrame
            display_start_time = render_queue_item.comp.displayStartTime

            if display_start_time is not None:
                start_frame = frame_offset + round(display_start_time / frame_duration) + round(render_queue_item.timeSpanStart / frame_duration)
                end_frame = start_frame + round(render_queue_item.timeSpanDuration / frame_duration) - 1
                frame_list = f"{start_frame}-{end_frame}"

            if first_and_last and not multi_machine:
                frame_list = f"{start_frame},{end_frame},{frame_list}"

        current_job_dependencies = deadline_settings.get('job_dependencies', "")
        if dependent_comps and dependent_job_id != "":
            if current_job_dependencies == "":
                current_job_dependencies = dependent_job_id
            else:
                current_job_dependencies = f"{dependent_job_id},{current_job_dependencies}"

        if multi_machine:
            job_name = f"{job_name} (multi-machine rendering {frame_list})"

        # Build job_attrs dict
        job_attrs = {
            "Plugin": "AfterEffects",
            "Name": job_name,
            "UserName": getpass.getuser(),
            "Comment": deadline_settings.get('comment', ''),
            "Department": deadline_settings.get('department', self.project_name or ''),
            "Group": deadline_settings.get('group', 'ae'),
            "Pool": deadline_settings.get('pool', self.ae_default_pool),
            "SecondaryPool": deadline_settings.get('secondary_pool', self.deadline_defaults['secondary_pool']),
            "Priority": deadline_settings.get('priority', self.deadline_defaults['priority']),
            "TaskTimeoutMinutes": deadline_settings.get('task_timeout_minutes', self.deadline_defaults['task_timeout_minutes']),
            "LimitGroups": deadline_settings.get('limit_groups', ''),
            "ConcurrentTasks": deadline_settings.get('concurrent_tasks', self.deadline_defaults['concurrent_tasks']),
            "LimitConcurrentTasksToNumberOfCpus": deadline_settings.get('limit_concurrent_tasks_to_number_cpus', '0'),
            "JobDependencies": current_job_dependencies,
            "OnJobComplete": deadline_settings.get('on_job_complete', 'Nothing'),
            "FailureDetectionTaskErrors": deadline_settings.get('failure_detection_task_errors', 8),
            "FailureDetectionJobErrors": deadline_settings.get('failure_detection_job_errors', 20),
        }
        # Group Jobs into a batch
        if dependent_comps or deadline_settings.get('group_submissions', False):
            job_attrs["BatchName"] = f"{self.project_name} - {ae_project_name}"

        # Blacklist/Whitelist
        if deadline_settings.get('submit_allow_list_as_deny_list', False):
            job_attrs["Blacklist"] = deadline_settings.get('machine_list', '')
        else:
            job_attrs["Whitelist"] = deadline_settings.get('machine_list', '')

        # Submit Suspended
        if deadline_settings.get('submit_suspended', False):
            job_attrs["InitialStatus"] = "Suspended"

        # Frame List
        if not is_movie and multi_machine:
            job_attrs["Frames"] = f"1-{round(deadline_settings.get('multi_machine_tasks', False))}"
        else:
            job_attrs["Frames"] = frame_list

        # Output files
        # Deadline needs OutputDirectoryN and OutputFilenameN as separate keys. deadlinecommand
        # splits a full path given in OutputFilenameN for you, but the DeadlineConnect / Web
        # Service submission does not, which left "Output Directory" empty in the Monitor.
        output_paths = []
        for index in range(render_queue_item.numOutputModules):
            output_paths.append(render_queue_item.outputModule(index + 1).file.fsName)
        job_attrs.update(self.build_deadline_output_attrs(output_paths))

        # Movie/Chunk settings
        if is_movie:
            job_attrs["MachineLimit"] = 1
            job_attrs["ChunkSize"] = 1000000
        else:
            if multi_machine:
                job_attrs["MachineLimit"] = 0
                job_attrs["ChunkSize"] = 1
            else:
                job_attrs["MachineLimit"] = deadline_settings.get('machine_limit', 0)
                job_attrs["ChunkSize"] = deadline_settings.get('chunk_size', 15)
        if multi_machine:
            job_attrs["ExtraInfoKeyValue0"] = f"FrameRangeOverride={frame_list}"

        # Build plugin_attrs dict
        plugin_attrs = {
            "SceneFile": render_file,
            "Comp": comp_name,
            "Version": version,
            "SubmittedFromVersion": self.adobe.app.version,
            "IgnoreMissingLayerDependenciesErrors": deadline_settings.get('missing_layers', False),
            "IgnoreMissingEffectReferencesErrors": deadline_settings.get('missing_effects', False),
            "FailOnWarnings": deadline_settings.get('fail_on_warnings', False),
            "MultiProcess": deadline_settings.get('multi_process', False),
            "ContinueOnMissingFootage": deadline_settings.get('missing_footage', False),
        }
        if deadline_settings.get('include_output_path', False):
            # The AfterEffects plugin passes this to aerender -output, so it must be the full
            # output file path (as the stock Deadline submitter writes it), not a folder.
            plugin_attrs["Output"] = output_file_path
        if multi_machine:
            plugin_attrs["MultiMachineMode"] = True
            plugin_attrs["MultiMachineStartFrame"] = start_frame
            plugin_attrs["MultiMachineEndFrame"] = end_frame

        if not multi_machine:
            min_file_size = deadline_settings.get('file_size', 0)
            delete_files_under_min_size = deadline_settings.get('delete_files', False)
            if deadline_settings.get('fail_on_missing_file', False):
                min_file_size = max(1, round(deadline_settings.get('file_size', 0)))
                delete_files_under_min_size = True
            plugin_attrs["MinFileSize"] = min_file_size
            plugin_attrs["DeleteFilesUnderMinSize"] = delete_files_under_min_size
            if deadline_settings.get('include_output_path', False):
                plugin_attrs["LocalRendering"] = deadline_settings.get('local_rendering', False)

        plugin_attrs["OverrideFailOnExistingAEProcess"] = deadline_settings.get('override_fail_on_existing_ae_process', False)
        plugin_attrs["FailOnExistingAEProcess"] = deadline_settings.get('fail_on_existing_ae_process', False)
        plugin_attrs["MemoryManagement"] = deadline_settings.get('memory_management', False)
        plugin_attrs["ImageCachePercentage"] = round(deadline_settings.get('image_cache_percentage', 0))
        plugin_attrs["MaxMemoryPercentage"] = round(deadline_settings.get('max_memory_percentage', 0))
        return job_attrs, plugin_attrs

    def submit_render_queue_item_to_deadlineconnect(self, job_attrs, plugin_attrs):
        """
            Submits a job to Deadline using DeadlineConnect with the provided job and plugin dictionaries.

            Arguments:
                job_attrs (dict): The dictionary containing job attributes for Deadline submission.
                plugin_attrs (dict): The dictionary containing plugin attributes for Deadline submission.
        """
        return submit_job_to_deadlineconnect(self.deadline_host, self.deadline_port, job_attrs, plugin_attrs)

    def apply_and_submit(self):
        """
            Apply the settings and submit to Deadline
        """
        # Open dialog to ask the user if they want to save the project
        save_project = self.question_box(
            "Save Project",
            "Saving the project is required before submitting to Deadline. Do you want to save the project now?",
        )

        # Return if the user selects No
        if not save_project:
            return

        logger.debug("Applying settings and submitting to Deadline")
        self.apply_to_render_queue_items()

        # Save file
        self.adobe.app.project.save()
        logger.debug("Project saved")

        # Submit files and save current submission settings
        self.submit_to_deadline_threaded()
        self.save_deadline_qsettings()

    ####################################
    # Deadline Lists
    ####################################
    def get_machine_list(self):
        """
            Get the machine list from Deadline
            Returns: A list of machines
        """
        logger.debug("Getting machine list from Deadline")
        result = self.call_deadline_command(["-GetSlaveNames"])
        logger.debug("Machine list: %s" % result)

        #Parse the result, only add from the fourth line onwards
        machines = self.parse_non_empty_lines(result)
        logger.debug("Returning machine list: %s" % machines)
        return machines

    def get_pool_list(self):
        """
            Get the pool list from Deadline
            Returns: A list of pools
        """
        logger.debug("Getting pool list from Deadline")
        result = self.call_deadline_command(["-Pools"])
        logger.debug("Pool list: %s" % result)

        # Parse the result, only add from the fourth line onwards
        pools = self.parse_non_empty_lines(result)
        return pools

    def get_group_list(self):
        """
            Get the group list from Deadline
            Returns: A list of groups
        """
        logger.debug("Getting group list from Deadline")
        result = self.call_deadline_command(["-Groups"])
        logger.debug("Group list: %s" % result)

        # Parse the result, only add from the fourth line onwards
        groups = self.parse_non_empty_lines(result)
        return groups

    def get_limit_group_list(self):
        """
            Get the limit group list from Deadline
            :returns: A list of limit groups
        """
        logger.debug("Getting limit group list from Deadline")
        result = self.call_deadline_command(["-LimitGroups"])
        logger.debug("Limit group list: %s" % result)

        # Parse the result, only add from the fourth line onwards
        limit_groups = self.parse_non_empty_lines(result)
        return limit_groups

    @staticmethod
    def parse_non_empty_lines(result):
        """
        Parse non-empty lines from the given result.

        Args:
            result (str): The string to parse.

        Returns:
            list: A list of non-empty lines.
        """
        return [line.strip() for line in result.splitlines() if line.strip()]

    ####################################
    # Deadline Dialogs
    ####################################
    def _exec_item_selection_dialog(self, title, items, target_widget):
        """
            Show an ItemSelectionDialog parented to the Toolkit window, then delete it.
        """
        dialog = ItemSelectionDialog(self.dialog_parent(), item_list=items, widget=target_widget)
        dialog.setWindowTitle(title)
        dialog.setWindowModality(QtCore.Qt.WindowModal)
        try:
            dialog.exec_()
        finally:
            dialog.deleteLater()

    def show_machine_list_dialog(self):
        """
            Show the machine list dialog
        """
        logger.debug("Showing machine list dialog")
        machines = self.get_machine_list()

        if not machines:
            self.alert_box("No machines found", "No machines found in the Deadline machine list")
            return

        self._exec_item_selection_dialog("Select Machine List", machines, self.ui.deadline_machine_list)

    def show_limit_group_list_dialog(self):
        """
            Show the limit group list dialog
        """
        logger.debug("Showing limit group list dialog")
        limit_groups = self.get_limit_group_list()

        if not limit_groups:
            self.alert_box("No limit groups found", "No limit groups found in the Deadline limit group list")
            return

        self._exec_item_selection_dialog("Select Limit Groups", limit_groups, self.ui.deadline_limits)

    ########################################
    # Threaded functions
    ########################################
    def prepare_deadline_jobs(self):
        """
            Build everything needed to submit the included rows, on the GUI thread.

            All After Effects (self.adobe) and widget access happens here, so the submission
            worker only copies files, runs AfterFX.com and talks to Deadline over HTTP.

            Returns:
                tuple: (jobs, error_message). jobs is a list of plain dicts; rows that fail a
                check are marked in the table and described in error_message.
        """
        jobs = []
        error_message = ""

        if not self.deadline_settings_initialized:
            self.populate_deadline_settings()

        if not self.run_project_checks():
            return jobs, "Project checks failed"

        if self.ui.compTableWidget.rowCount() == 0:
            return jobs, "No render queue items available"

        deadline_settings = self.get_deadline_settings()
        if deadline_settings is None:
            return jobs, "Could not read the Deadline settings"

        app_version = self.adobe.app.version
        version = '.'.join(app_version.split('x')[0].split('.')[:2])
        project_path = self.adobe.app.project.file.fsName

        for row in range(self.ui.compTableWidget.rowCount()):
            tableItem = self.ui.compTableWidget.item(row, 0)
            includeCheckBox = self.ui.compTableWidget.item(row, 6)
            useCompNameCheckBox = self.ui.compTableWidget.item(row, 5)
            if not tableItem or not includeCheckBox or includeCheckBox.checkState() != QtCore.Qt.Checked:
                continue

            compName = tableItem.text()
            try:
                render_queue_item = tableItem.data(QtCore.Qt.UserRole)
                compName = render_queue_item.comp.name

                # Item Status Check
                if render_queue_item.status != self.adobe.RQItemStatus.QUEUED:
                    raise ValueError(f"Skipped because comp '{compName}' is not in QUEUED status (status={render_queue_item.status}).")

                # Item Spaces check
                if compName.startswith(" ") or compName.endswith(" "):
                    raise ValueError(f"Comp name '{compName}' has spaces at the front or back. Skipping.")

                # Item Special Character check
                if not re.match(r'^[a-zA-Z0-9_]+$', compName):
                    raise ValueError(f"Comp name '{compName}' has special characters. Only letters, numbers, and underscores are allowed. Skipping.")

                use_comp_name = useCompNameCheckBox.checkState() == QtCore.Qt.Checked
                render_queue_template = self.get_render_queue_template(row)
                render_scene_file_path = self.get_shotgrid_template(render_queue_template, use_comp_name, compName, True)

                job_attrs, plugin_attrs = self.build_deadline_job_and_plugin_dicts(
                    render_queue_item=render_queue_item,
                    deadline_settings=deadline_settings,
                    project_path=project_path,
                    layers=False,
                    previous_job_id="",
                    version=version,
                    render_file=render_scene_file_path
                )

                jobs.append({
                    "row": row,
                    "comp_name": compName,
                    "project_path": project_path,
                    "render_scene_file_path": render_scene_file_path,
                    "manifest_job": self.build_manifest_job(
                        render_queue_item, render_scene_file_path, project_manifest=True, comp_manifest=True),
                    "job_attrs": job_attrs,
                    "plugin_attrs": plugin_attrs,
                })

            except Exception as e:
                logger.debug(traceback.format_exc())
                msg = str(e) if isinstance(e, ValueError) else f"Failed to prepare comp '{compName}': {e}"
                error_message += msg + "\n"
                self.on_submission_item_update(row, msg, False)

        return jobs, error_message

    def submit_to_deadline_threaded(self):
        """
            Submits the render queue items to Deadline using a separate thread to avoid blocking the UI
            and to provide progress updates through a DeadlineProgressDialog.
        """
        if self.is_busy():
            self.warning_box("Deadline Submission", "A submission is already running.")
            return

        self.toggle_buttons(False)
        try:
            jobs, self._submission_prep_errors = self.prepare_deadline_jobs()
        except Exception as e:
            logger.error(traceback.format_exc())
            jobs, self._submission_prep_errors = [], f"Failed to prepare the submission: {e}"

        if not jobs:
            self.toggle_buttons(True)
            self.message_box("Deadline Submission",
                             f"Nothing was submitted.\n\n{self._submission_prep_errors}".strip())
            return

        self.show_progress_bar(format_text="Submitting to Deadline...")

        self.deadline_progress_dialog = DeadlineProgressDialog(
            [job["comp_name"] for job in jobs], logger=logger, parent=self.dialog_parent())
        self.deadline_progress_dialog.show()

        # Thread parented to this widget and kept on self (named _submit_thread so it doesn't
        # shadow QObject.thread()). closeEvent refuses to close while it runs.
        self._submit_thread = QtCore.QThread(self)
        self._submit_worker = DeadlineSubmissionWorker(jobs, self.deadline_host, self.deadline_port, self.thread_log)
        self._submit_worker.moveToThread(self._submit_thread)

        # Connections. The worker lives in the other thread, so these are queued connections
        # and every slot below runs on the GUI thread.
        self._submit_thread.started.connect(self._submit_worker.run)
        self._submit_worker.progress.connect(self.on_submission_progress)
        self._submit_worker.item_update.connect(self.on_submission_item_update)
        self._submit_worker.row_progress.connect(self.deadline_progress_dialog.update_progress)
        self._submit_worker.row_done.connect(self.deadline_progress_dialog.mark_done)
        self._submit_worker.finished.connect(self._submit_thread.quit)
        self._submit_worker.finished.connect(self._submit_worker.deleteLater)
        self._submit_worker.finished.connect(self.on_submission_finished)

        # Start Thread
        self._submit_thread.start()

    def on_submission_item_update(self, row, message, success=True):
        """
            Update the status icon and tooltip for a specific row in the comp table based on the submission result.

            Arguments:
                row (int): The row to update.
                message (str): The message to display.
                success (bool): Whether the row was successfully updated.

        """
        statusItem = self.ui.compTableWidget.item(row, 1)
        if statusItem is None:
            statusItem = QtGui.QTableWidgetItem("")
            self.ui.compTableWidget.setItem(row, 1, statusItem)
        if success:
            statusItem.setIcon(self.ui.submitIcon)
            statusItem.setToolTip(message or "Submitted to Deadline")
        else:
            statusItem.setIcon(self.ui.errorIcon)
            statusItem.setToolTip(message or "Submission Failed")

    def on_submission_progress(self, percent, message):
        """
            Update the overall submission progress bar and message.

            Arguments:
                percent (int): The percent to update.
                message (str): The message to display.
        """
        self.update_progress_bar(percent)
        self.update_progress_bar_format(message)

    def on_submission_finished(self, error_message, num_successful_submissions):
        """
            Handle the completion of the Deadline submission process, closing the progress dialog and showing a
            message box with the results.

            Arguments:
                error_message (str): The message to display.
                num_successful_submissions (int): The number of successful submissions.
        """
        # Let the thread finish before anything can close the window.
        if self._submit_thread is not None:
            self._submit_thread.wait()
            self._submit_thread.deleteLater()
            self._submit_thread = None
        self._submit_worker = None

        if self.deadline_progress_dialog is not None:
            self.deadline_progress_dialog.close()
            self.deadline_progress_dialog.deleteLater()
            self.deadline_progress_dialog = None

        self.hide_progress_bar()
        self.toggle_buttons(True)

        error_message = (self._submission_prep_errors or "") + (error_message or "")
        self._submission_prep_errors = ""

        if error_message:
            self.message_box(
                "Deadline Submission",
                f"{num_successful_submissions} job(s) submitted to Deadline.\n\nProblems:\n{error_message}")
        else:
            self.message_box("Deadline Submission",
                             f"Submission completed successfully.\n\n{num_successful_submissions} jobs submitted to Deadline.")
        self.activateWindow()


class GuiThreadLogRelay(QtCore.QObject):
    """
        A logger for worker threads that hands every record to the GUI thread.

        tk-aftereffects forwards each Toolkit log record to After Effects over the RPC
        socket (engine._emit_log_message -> adobe.log_message). That socket is shared with
        the GUI thread and is not thread safe, so logging straight from a worker thread can
        corrupt the connection and take the Toolkit panel down. Create this on the GUI thread.
    """
    record = QtCore.Signal(int, str)

    def __init__(self, target_logger, parent=None):
        super(GuiThreadLogRelay, self).__init__(parent)
        self._logger = target_logger
        self.record.connect(self._on_record, QtCore.Qt.QueuedConnection)

    @QtCore.Slot(int, str)
    def _on_record(self, level, message):
        self._logger.log(level, message)

    def log(self, level, message, *args):
        if args:
            message = message % args
        app = QtCore.QCoreApplication.instance()
        if app is not None and QtCore.QThread.currentThread() == app.thread():
            self._logger.log(level, message)
        else:
            self.record.emit(level, str(message))

    def debug(self, message, *args):
        self.log(logging.DEBUG, message, *args)

    def info(self, message, *args):
        self.log(logging.INFO, message, *args)

    def warning(self, message, *args):
        self.log(logging.WARNING, message, *args)

    def error(self, message, *args):
        self.log(logging.ERROR, message, *args)


class _CallableWorker(QtCore.QObject):
    """
        Runs one callable in a worker thread (see AppDialog.run_blocking_in_thread).
    """
    finished = QtCore.Signal()

    def __init__(self, func):
        super(_CallableWorker, self).__init__()
        self._func = func

    @QtCore.Slot()
    def run(self):
        try:
            self._func()
        finally:
            self.finished.emit()


def submit_job_to_deadlineconnect(host, port, job_attrs, plugin_attrs, log=None):
    """
        Submit one job through DeadlineConnect (HTTP to the Web Service). Safe in any thread.

        Returns:
            str: The new job ID.

        Raises:
            Exception: If Deadline did not return a job.
    """
    import Deadline.DeadlineConnect as Connect

    log = log or logger
    deadline = Connect.DeadlineCon(host, port)
    job = deadline.Jobs.SubmitJob(job_attrs, plugin_attrs)
    log.info(f"Deadline submission results: {job}")

    # If job is a dict and contains '_id', treat as success
    if isinstance(job, dict) and '_id' in job:
        log.info(f"Submitted Job ID: {job['_id']}")
        return job['_id']

    raise Exception(f"Deadline submission error: {job}")


class DeadlineSubmissionWorker(QtCore.QObject):
    """
        Worker class for handling the Deadline submission process in a separate thread, allowing for progress updates
        and UI responsiveness.

        It only gets plain data (from AppDialog.prepare_deadline_jobs) and never touches widgets
        or engine.adobe: the After Effects RPC bridge and Qt widgets are not thread safe.

        Signals:
            progress (int, str): Emitted to update overall submission progress (percent, message).
            finished (str, int): Emitted when submission is finished (error_message, num_successful_submissions).
            item_update (int, str, bool): Emitted to update a specific row's status (row, message, success).
            row_progress (int, int, str): Emitted to update a specific row's progress (row, percent, message).
            row_done (int, bool): Emitted when a specific row's submission is done (row, success).
    """
    progress = QtCore.Signal(int, str)  # percent, message
    finished = QtCore.Signal(str, int)  # error message, num_successful_submissions
    item_update = QtCore.Signal(int, str, bool)  # row, message, success
    row_progress = QtCore.Signal(int, int, str)  # progress index, percent, message
    row_done = QtCore.Signal(int, bool)  # progress index, success

    def __init__(self, jobs, deadline_host, deadline_port, log):
        super(DeadlineSubmissionWorker, self).__init__()
        self.jobs = jobs
        self.deadline_host = deadline_host
        self.deadline_port = deadline_port
        # Never the plain Toolkit logger here (see GuiThreadLogRelay)
        self.log = log

    @QtCore.Slot()
    def run(self):
        """
            Copy the render scene, generate the manifests and submit each prepared job.
        """
        error_message = ""
        num_successful_submissions = 0

        try:
            total = max(1, len(self.jobs))
            for idx, job in enumerate(self.jobs):
                row = job["row"]
                compName = job["comp_name"]
                render_scene_file_path = job["render_scene_file_path"]
                self.progress.emit(int(idx / total * 100), f"Submitting {compName} to Deadline...")
                self.row_progress.emit(idx, 0, "Starting submission...")

                try:
                    # Step 1: Copy project file
                    self.row_progress.emit(idx, 10, f"Copying project file for {compName}...")
                    os.makedirs(os.path.dirname(render_scene_file_path), exist_ok=True)
                    shutil.copy(job["project_path"], render_scene_file_path)
                except Exception as e:
                    self.log.error(traceback.format_exc())
                    self._fail(idx, row, f"Failed to create render scene backup for comp '{compName}': {e}")
                    error_message += f"Failed to create render scene backup for comp '{compName}': {e}\n"
                    continue

                try:
                    # Step 2: Generate manifests
                    self.row_progress.emit(idx, 30, f"Generating manifests for {compName}...")
                    AppDialog.run_manifest_job(job["manifest_job"], log=self.log)
                except Exception as e:
                    self.log.error(traceback.format_exc())
                    self._fail(idx, row, f"Failed to generate manifest file for comp '{compName}': {e}")
                    error_message += f"Failed to generate manifest file for comp '{compName}': {e}\n"
                    continue

                try:
                    # Step 3: Submit to Deadline
                    self.row_progress.emit(idx, 70, f"Submitting {compName}...")
                    job_id = submit_job_to_deadlineconnect(
                        self.deadline_host, self.deadline_port, job["job_attrs"], job["plugin_attrs"], log=self.log)
                    num_successful_submissions += 1
                    self.item_update.emit(row, f"Submitted comp '{compName}' to Deadline successfully ({job_id}).", True)
                    self.row_done.emit(idx, True)
                except Exception as e:
                    self.log.error(traceback.format_exc())
                    self._fail(idx, row, f"Failed to submit comp '{compName}' to Deadline: {e}")
                    error_message += f"Failed to submit comp '{compName}' to Deadline: {e}\n"
                    continue

            # Final progress
            self.progress.emit(100, "Finished submitting to Deadline")

        except Exception as e:
            self.log.error(traceback.format_exc())
            error_message += str(e)

        finally:
            # Always emit, so the UI is unlocked and the thread is stopped.
            self.finished.emit(error_message, num_successful_submissions)

    def _fail(self, idx, row, message):
        self.item_update.emit(row, message, False)
        self.row_done.emit(idx, False)


class DeadlineProgressDialog(QtGui.QDialog):
    """
        Dialog class for displaying the progress of Deadline submissions, with individual progress bars for each comp being submitted.
    """
    def __init__(self, comp_names, logger, parent=None):
        super().__init__(parent)
        self.logger = logger
        self.setWindowTitle("Deadline Submission Progress")
        # Window modal to its (Toolkit) parent rather than application modal, and no close
        # button: it is closed by the app once the submission thread has finished.
        self.setWindowModality(QtCore.Qt.WindowModal)
        self.setWindowFlags(
            (self.windowFlags() | QtCore.Qt.CustomizeWindowHint | QtCore.Qt.WindowTitleHint)
            & ~QtCore.Qt.WindowCloseButtonHint
            & ~QtCore.Qt.WindowContextHelpButtonHint
        )
        self.resize(600, 60 + 30 * len(comp_names))
        self.layout = QtGui.QVBoxLayout(self)
        self.progress_bars = {}

        for idx, name in enumerate(comp_names):
            row_widget = QtGui.QWidget(self)
            row_layout = QtGui.QHBoxLayout(row_widget)
            label = QtGui.QLabel(name, row_widget)
            progress = QtGui.QProgressBar(row_widget)
            progress.setMinimum(0)
            progress.setMaximum(100)
            progress.setValue(0)
            progress.setTextVisible(True)
            progress.setFormat("%p%")
            row_layout.addWidget(label)
            row_layout.addWidget(progress)
            self.layout.addWidget(row_widget)
            self.progress_bars[idx] = progress

    @QtCore.Slot(int, int, str)
    def update_progress(self, row, percent, message=""):
        """
            Update the progress bar for a specific row with the given percentage and message.

            Arguments:
                row (int): The index of the row to update.
                percent (int): The progress percentage to set (0-100).
                message (str): An optional message to display alongside the percentage.

        """
        self.logger.debug(f"update_progress called: row={row}, percent={percent}, message={message}")
        bar = self.progress_bars.get(row)
        if bar:
            bar.setValue(percent)
            if message:
                bar.setFormat(f"{message} %p%")
            else:
                bar.setFormat("%p%")
        else:
            self.logger.debug(f"No progress bar found for row {row}")

    @QtCore.Slot(int, bool)
    def mark_done(self, row, success=True):
        """
            Mark the progress for a specific row as done, setting the progress to 100% and updating the format based on success.

            Arguments:
                row (int): The index of the row to update.
                success (bool): Whether the progress should be marked as done.
        """
        self.logger.debug(f"mark_done called: row={row}, success={success}")
        bar = self.progress_bars.get(row)
        if bar:
            bar.setValue(100)
            if success:
                bar.setFormat("Done %p%")
            else:
                bar.setFormat("Error %p%")
        else:
            self.logger.debug(f"No progress bar found for row {row}")
