# Copyright (c) 2013 Shotgun Software Inc.
#
# CONFIDENTIAL AND PROPRIETARY
#
# This work is provided "AS IS" and subject to the Shotgun Pipeline Toolkit
# Source Code License included in this distribution package. See LICENSE.
# By accessing, using, copying or modifying this work you indicate your
# agreement to the Shotgun Pipeline Toolkit Source Code License. All rights
# not expressly granted therein are reserved by Shotgun Software Inc.
import shutil
import subprocess
import warnings
import functools

import sgtk
import os
import re
import time
import json
import sys
from contextlib import contextmanager
import traceback

# by importing QT from sgtk rather than directly, we ensure that
# the code will be compatible with both PySide and PyQt.
from sgtk.platform.qt import QtCore, QtGui
from .ui.dialog import Ui_Dialog, ItemSelectionDialog

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

def deprecated(reason=""):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            warnings.warn(
                f"Call to deprecated function {func.__name__}. {reason}",
                category=DeprecationWarning,
                stacklevel=2
            )
            return func(*args, **kwargs)
        return wrapper
    return decorator

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

        # Deadline variables
        self.deadline_error_message = ""
        self.ae_default_pool = None
        self.deadline_settings_initialized = False

        # lastly, set up our very basic UI
        # self.ui.context.setText("Current Context: %s" % self._app.context)
        self.populate_presets()
        self.populate_sg_fields()
        self.connect_signals_and_slots()

        # Create render queue items
        self.create_render_queue_items()

        # Set the window size
        self.resize(1000, self.ui.compTableWidget.sizeHint().height())

        # Manifest cache to avoid redundant analysis of comps
        self._manifest_cache = {}

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

    def refresh_table_item_data(self):
        """
        Refresh only the UserRole data for each table row, preserving status icons/tooltips.
        """
        for row in range(self.ui.compTableWidget.rowCount()):
            try:
                # Fetch the latest render queue item for this row
                render_queue_item = self.adobe.app.project.renderQueue.item(row + 1)
                # Update the UserRole data for the first column (comp name)
                tableItem = self.ui.compTableWidget.item(row, 0)
                if tableItem is not None:
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
                                render_queue_item = self.adobe.app.project.renderQueue.item(row + 1)
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

    def _run_jsx_manifest_generation(self, comp_identifiers, jsx_script_path):
        """
            Shared helper to write comp_identifiers to _comp_identifiers.json and run the specified JSX script.

            Arguments:
                comp_identifiers (list): A list of comp_identifiers to run the JSX script.
                jsx_script_path (str): The path to the JSX script to be run.

            Returns:
            None
        """
        def _after_effects_version_to_year(major_version):
            """
                Converts the major version number of After Effects to the corresponding year-based version string.
                 - For versions 24 and above, it converts to a year-based version (e.g., 24 to 2024).

                Arguments:
                    major_version (str): The major version number of After Effects as a string.
                Returns:
                    str: The year-based version string for After Effects if major version is 24 or above,
                    otherwise returns the original major version string.
            """
            try:
                major_version_int = int(major_version)
                if major_version_int >= 24:
                    return str(2000 + major_version_int)
                else:
                    return major_version
            except ValueError:
                logger.error(f"Unable to parse After Effects version: {self.ae_version}")
                return major_version

        # Validate comp_identifiers
        if not isinstance(comp_identifiers, list) or not comp_identifiers:
            logger.error("comp_identifiers must be a non-empty list.")
            return
        required_fields = ("name", "id", "output_location")
        for idx, comp in enumerate(comp_identifiers):
            if not isinstance(comp, dict):
                logger.error(f"comp_identifiers[{idx}] is not a dict: {comp}")
                return
            for field in required_fields:
                if field not in comp or comp[field] in (None, ""):
                    logger.error(f"comp_identifiers[{idx}] missing or empty required field '{field}': {comp}")
                    return

        # Check After Effects project file existence
        current_comp_file_path = getattr(self.adobe.app.project.file, 'fsName', None)
        if not current_comp_file_path or not os.path.isfile(current_comp_file_path):
            logger.error(f"After Effects project file not found or invalid: {current_comp_file_path}")
            return

        # Check directory writability for _comp_identifiers.json
        comp_id_dir = os.path.dirname(current_comp_file_path)
        if not os.path.isdir(comp_id_dir):
            logger.error(f"Directory for comp identifiers does not exist: {comp_id_dir}")
            return
        if not os.access(comp_id_dir, os.W_OK):
            logger.error(f"No write permission for directory: {comp_id_dir}")
            return

        # Write comp identifier JSON file next to the project file
        comp_id_json_path = os.path.join(comp_id_dir, "_comp_identifiers.json")
        try:
            with open(comp_id_json_path, "w") as f:
                json.dump(comp_identifiers, f, indent=4)
        except Exception as e:
            logger.error(f"Failed to write comp identifiers JSON: {e}")
            return

        major_version = str(self.ae_version).split(".")[0]
        after_effects_version = _after_effects_version_to_year(major_version)
        afterfx_path = r"C:\Program Files\Adobe\Adobe After Effects %s\Support Files\AfterFX.com" % after_effects_version

        # Check AfterFX executable existence
        if not os.path.isfile(afterfx_path):
            logger.error(f"AfterFX executable not found: {afterfx_path}")
            return

        command = [afterfx_path, "-ro", jsx_script_path]

        logger.info(f"Running command: {command}")

        try:
            result = subprocess.run(command, check=True, capture_output=True, text=True)
            if result.returncode != 0:
                logger.error(f"JSX script exited with non-zero code: {result.returncode}", file=sys.stderr)
                logger.error(f"Stdout: {result.stdout}", file=sys.stderr)
                logger.error(f"Stderr: {result.stderr}", file=sys.stderr)
        except subprocess.CalledProcessError as e:
            logger.error(f"JSX script failed to execute: {e}", file=sys.stderr)
            logger.error(f"Stdout: {e.stdout}", file=sys.stderr)
            logger.error(f"Stderr: {e.stderr}", file=sys.stderr)
        except OSError as e:
            logger.error(f"Failed to run After Effects JSX script: {e}", file=sys.stderr)
            logger.error(f"Command: {' '.join(command)}")

    def generate_manifest_file_for_queue_item_jsx(self, render_queue_item, render_scene_file_path):
        """
            Creates a JSON file with comp name and id, then executes the JSX script to generate the manifest file.

            Arguments:
                render_queue_item (RenderQueueItem): The render queue item to generate manifest file.
                render_scene_file_path (str): The path to the scene file to generate manifest file.
        """
        comp = render_queue_item.comp
        comp_identifier = {
            "name": comp.name,
            "id": getattr(comp, 'id', None),
            "output_location": os.path.dirname(render_scene_file_path)
        }
        jsx_script_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../jsx/generate_manifest_from_comps.jsx'))

        # Check jsx file exits
        if jsx_script_path is None or not os.path.exists(jsx_script_path):
            logger.error(f"JSX script not found at path: {jsx_script_path}")
            return

        self._run_jsx_manifest_generation([comp_identifier], jsx_script_path)

    def generate_project_manifest_file_jsx(self, render_queue_item, render_scene_file_path):
        """
            Creates a JSON file with comp name and id, then executes the JSX script to generate the manifest file for the entire project.

            Arguments:
                render_queue_item (RenderQueueItem): The render queue item to generate manifest file.
                render_scene_file_path (str): The path to the scene file to generate manifest file.
        """
        comp = render_queue_item.comp
        comp_identifier = {
            "name": comp.name,
            "id": getattr(comp, 'id', None),
            "output_location": os.path.dirname(render_scene_file_path)
        }
        jsx_script_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../../jsx/generate_manifest_for_all_comps.jsx'))

        if jsx_script_path is None or not os.path.exists(jsx_script_path):
            logger.error(f"JSX script not found at path: {jsx_script_path}")
            return

        self._run_jsx_manifest_generation([comp_identifier], jsx_script_path)

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

        save_project = QtGui.QMessageBox.question(
            self,
            "Save Project",
            "Saving is required before rendering. Do you want to save the current project?",
            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No,
            QtGui.QMessageBox.Yes,
        )

        if save_project == QtGui.QMessageBox.Yes:
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
                    self.generate_project_manifest_file_jsx(render_queue_item, render_scene_file_path)
                    logger.debug("Project manifest file generated for render queue item: %s" % compName)


                    logger.debug("Generating manifest file for render queue item: %s" % compName)
                    self.generate_manifest_file_for_queue_item_jsx(render_queue_item, render_scene_file_path)
                    logger.debug("Manifest file generated for render queue item: %s" % compName)

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
        ready_to_render = self.process_queue_items_for_render()

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

    def alert_box(self, title, text):
        """
            Display an alert box

            Arguments:
                title: The title of the alert box
                text: The text of the alert box

            Returns: None
        """
        QtGui.QMessageBox.critical(
            self,
            title,
            str(text),
            buttons=QtGui.QMessageBox.Ok,
            defaultButton=QtGui.QMessageBox.Ok,
        )

    def warning_box(self, title, text):
        """
            Display a warning box

            Arguments:
                title: The title of the warning box
                text: The text of the warning box

            Returns: None
        """
        logger.debug("Displaying Warning Box: %s" % text)
        # Display the warning box
        QtGui.QMessageBox.warning(
            self,
            title,
            str(text),
            buttons=QtGui.QMessageBox.Ok,
            defaultButton=QtGui.QMessageBox.Ok,
        )

    def message_box(self, title, text):
        """
            Display a message box

            Arguments:
                title: The title of the message box
                text: The text of the message box

            Returns: None
        """
        logger.debug("Displaying Message Box: %s" % text)
        # Display the message box
        QtGui.QMessageBox.information(
            self,
            title,
            str(text),
            buttons=QtGui.QMessageBox.Ok,
            defaultButton=QtGui.QMessageBox.Ok,
        )

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

    #####################################################################################################
    # Events
    #####################################################################################################
    def eventFilter(self, obj, event):
        """
        Event filter to void mouse scroll events for specific widgets.

        Arguments:
            obj: The object that received the event.
            event: The event that was received.
        Returns:
            bool: True if the event is handled, False otherwise.
        """
        if event.type() == QtCore.QEvent.Wheel and isinstance(obj, QtGui.QComboBox):
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
            self.deadline_settings = {
                'priority': self.ui.deadline_priority.text(),
                'pool': self.ui.deadline_pool.currentText(),
                'secondary_pool': self.ui.deadline_secondary_pool.currentText(),
                'group': self.ui.deadline_group.currentText(),
                'chunk_size': self.ui.deadline_frames_per_task.text(),
                'frame_list': self.ui.deadline_frame_list.text(),
                'submit_scene': self.ui.deadline_submit_project_file_with_job.isChecked(),
                'override_frame_list': self.ui.deadline_use_frame_list_from_comp.isChecked(),
                'task_timeout_minutes': self.ui.deadline_task_timeout.text(),
                'concurrent_tasks': self.ui.deadline_concurrent_tasks.text(),
                'limit_groups': self.ui.deadline_limits.get_text(),
                'machine_list': self.ui.deadline_machine_list.get_text(),
                'submit_allow_list_as_deny_list': self.ui.deadline_machine_list_deny.isChecked(),
                'submit_suspended': self.ui.deadline_submit_as_suspended.isChecked(),
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
        except Exception as e:
            logger.error(f"Failed to get Deadline settings: {e}")
            error = traceback.format_exc()
            logger.error(error)

        return self.deadline_settings

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
            return None

        startupinfo = None
        creationflags = 0
        if os.name == 'nt' and hideWindow:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NEW_CONSOLE

        try:
            logger.debug("Calling Deadline command: %s" % deadlineCommand)
            result = subprocess.run(
                [deadlineCommand] + arguments,
                startupinfo=startupinfo,
                creationflags=creationflags,
                capture_output=readStdout,
                text=True
            )
            if readStdout:
                return result.stdout
        except Exception as e:
            logger.error(f"Failed to execute Deadline command: {e}")
            return None

    ####################################
    # QSettings Functions
    ###################################
    def save_deadline_qsettings(self):
        """
            Save the Deadline settings to a QSettings file.
        """
        scene_path = self.adobe.app.project.file.fsName
        settings = self.get_deadline_settings()
        qsettings = QtCore.QSettings("Territory", "AfterEffectsDeadlineSubmission")
        qsettings.setValue(f"deadline_settings/{scene_path}", settings)

    def load_deadline_qsettings(self):
        """
            Load the Deadline settings from a QSettings file.
            It first tries to load settings specific to the current scene, and if not found, it loads the last used settings.
        """
        scene_path = self.adobe.app.project.file.fsName
        qsettings = QtCore.QSettings("Territory", "AfterEffectsDeadlineSubmission")
        settings = qsettings.value(f"deadline_settings/{scene_path}")

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

        setters = {
            "priority": lambda v: self.ui.deadline_priority.setText(v),
            "pool": lambda v: self.ui.deadline_pool.setCurrentText(v),
            "secondary_pool": lambda v: self.ui.deadline_secondary_pool.setCurrentText(v),
            "group": lambda v: self.ui.deadline_group.setCurrentText(v),
            "on_job_complete": lambda v: self.ui.deadline_on_job_complete.setCurrentText(v),
            "chunk_size": lambda v: self.ui.deadline_frames_per_task.setText(v),
            "frame_list": lambda v: self.ui.deadline_frame_list.setText(v),
            "submit_scene": lambda v: self.ui.deadline_submit_project_file_with_job.setChecked(v),
            "override_frame_list": lambda v: self.ui.deadline_use_frame_list_from_comp.setChecked(v),
            "task_timeout_minutes": lambda v: self.ui.deadline_task_timeout.setText(v),
            "concurrent_tasks": lambda v: self.ui.deadline_concurrent_tasks.setText(v),
            "limit_groups": lambda v: self.ui.deadline_limits.set_text(v),
            "machine_list": lambda v: self.ui.deadline_machine_list.set_text(v),
            "submit_allow_list_as_deny_list": lambda v: self.ui.deadline_machine_list_deny.setChecked(v),
            "submit_suspended": lambda v: self.ui.deadline_submit_as_suspended.setChecked(v),
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
            "comment": lambda v: self.ui.deadline_comment.setText(v),
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
    def build_deadline_job_and_plugin_dicts(self, render_queue_item, deadline_settings, project_path, layers, previous_job_id, version, render_file):
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
        output_module = render_queue_item.outputModule(render_queue_item.numOutputModules)
        output_file_name = output_module.file.name
        output_folder = os.path.dirname(output_file_name)
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
        is_movie = any(output_file_name.endswith(ext) for ext in self.deadline_defaults["movie_formats"])

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
            "UserName": os.getlogin(),
            "Comment": deadline_settings.get('comment', ''),
            "Department": deadline_settings.get('department', self.project_name or ''),
            "Group": deadline_settings.get('group', 'ae'),
            "Pool": deadline_settings.get('pool', 'none'),
            "SecondaryPool": deadline_settings.get('secondary_pool', 'none'),
            "Priority": deadline_settings.get('priority', 30),
            "TaskTimeoutMinutes": deadline_settings.get('task_timeout_minutes', '0'),
            "LimitGroups": deadline_settings.get('limit_groups', ''),
            "ConcurrentTasks": deadline_settings.get('concurrent_tasks', '1'),
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
            job_attrs["Blacklist"] = deadline_settings['machine_list']
        else:
            job_attrs["Whitelist"] = deadline_settings['machine_list']

        # Submit Suspended
        if deadline_settings.get('submit_suspended', False):
            job_attrs["InitialStatus"] = "Suspended"

        # Frame List
        if not is_movie and multi_machine:
            job_attrs["Frames"] = f"1-{round(deadline_settings.get('multi_machine_tasks', False))}"
        else:
            job_attrs["Frames"] = frame_list

        # Output files
        for index in range(render_queue_item.numOutputModules):
            output_module = render_queue_item.outputModule(index + 1)
            key = f"OutputFilename{index}"
            value = output_module.file.fsName.replace('[#', '#').replace('#]', '#')
            job_attrs[key] = value

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
            plugin_attrs["Output"] = output_folder
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
        try:
            import Deadline.DeadlineConnect as Connect
            host = self.deadline_host
            port = self.deadline_port
            deadline = Connect.DeadlineCon(host, port)
            job = deadline.Jobs.SubmitJob(job_attrs, plugin_attrs)
            logger.info(f"Deadline submission results: {job}")
            if isinstance(job, dict):
                    if '_id' in job:
                        logger.info(f"Submitted Job ID: {job['_id']}")
            else:
                logger.error(f"Deadline submission error: {job}")
        except Exception as e:
            logger.error(f"DeadlineConnect submission failed: {e}")
            logger.error(traceback.format_exc())
            raise

    def apply_and_submit(self):
        """
            Apply the settings and submit to Deadline
        """
        # Open dialog to ask the user if they want to save the project
        save_project = QtGui.QMessageBox.question(
            self,
            "Save Project",
            "Saving the project is required before submitting to Deadline. Do you want to save the project now?",
            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No,
            QtGui.QMessageBox.Yes,
        )

        # Return if the user selects No
        if save_project == QtGui.QMessageBox.No:
            return

        logger.debug("Applying settings and submitting to Deadline")
        self.apply_to_render_queue_items()
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
    def show_machine_list_dialog(self):

        """
            Show the machine list dialog
        """
        logger.debug("Showing machine list dialog")
        machines = self.get_machine_list()

        if not machines:
            self.alert_box("No machines found", "No machines found in the Deadline machine list")
            return

        # Create the dialog
        dialog = ItemSelectionDialog(self, item_list=machines, widget=self.ui.deadline_machine_list)
        dialog.setWindowTitle("Select Machine List")

        dialog.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Dialog)
        dialog.exec_()

    def show_limit_group_list_dialog(self):
        """
            Show the limit group list dialog
        """
        logger.debug("Showing limit group list dialog")
        limit_groups = self.get_limit_group_list()

        if not limit_groups:
            self.alert_box("No limit groups found", "No limit groups found in the Deadline limit group list")
            return

        # Create the dialog
        dialog = ItemSelectionDialog(self, item_list=limit_groups, widget=self.ui.deadline_limits)
        dialog.setWindowTitle("Select Limit Groups")

        dialog.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Dialog)
        dialog.exec_()

    ########################################
    # Threaded functions
    ########################################
    def submit_to_deadline_threaded(self):
        """
            Submits the render queue items to Deadline using a separate thread to avoid blocking the UI
            and to provide progress updates through a DeadlineProgressDialog.
        """

        comp_rows = []
        comp_names = []
        row_to_progress_idx = {}

        for idx, row in enumerate(range(self.ui.compTableWidget.rowCount())):
            item = self.ui.compTableWidget.item(row, 0)
            includeCheckBox = self.ui.compTableWidget.item(row, 6)
            if item and includeCheckBox and includeCheckBox.checkState() == QtCore.Qt.Checked:
                comp_rows.append(row)
                comp_names.append(item.text())
                row_to_progress_idx[row] = len(comp_rows) - 1
        self.deadline_progress_dialog = DeadlineProgressDialog(comp_names, logger=logger, parent=self)
        self.deadline_progress_dialog.show()

        # Create Thread For submissions
        self.thread = QtCore.QThread()
        self.worker = DeadlineSubmissionWorker(self, comp_rows, row_to_progress_idx)
        self.worker.moveToThread(self.thread)

        # Connections
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_submission_progress)
        self.worker.finished.connect(self.on_submission_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.item_update.connect(self.on_submission_item_update)
        self.worker.row_progress.connect(self.deadline_progress_dialog.update_progress)
        self.worker.row_done.connect(self.deadline_progress_dialog.mark_done)

        # Start Thread
        self.thread.start()

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
        if hasattr(self, "deadline_progress_dialog") and self.deadline_progress_dialog:
            self.deadline_progress_dialog.close()
            self.deadline_progress_dialog = None

        self.hide_progress_bar()
        self.toggle_buttons(True)

        if error_message:
            self.message_box("Deadline Submission", f"Submission failed: {error_message}")
        else:
            self.message_box("Deadline Submission",
                             f"Submission completed successfully.\n\n{num_successful_submissions} jobs submitted to Deadline.")
        self.activateWindow()


class DeadlineSubmissionWorker(QtCore.QObject):
    """
        Worker class for handling the Deadline submission process in a separate thread, allowing for progress updates
        and UI responsiveness.

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
    row_progress = QtCore.Signal(int, int, str)  # row, percent, message
    row_done = QtCore.Signal(int, bool)  # row, success

    def __init__(self, parent_dialog, comp_rows, row_to_progress_idx):
        super().__init__()
        self.dialog = parent_dialog
        self.comp_rows = comp_rows
        self.row_to_progress_idx = row_to_progress_idx

    @QtCore.Slot()
    def run(self):
        """
            Main function for the worker thread to handle the Deadline submission process,
            iterating through the specified render queue items, performing necessary checks, generating manifests,
            and submitting to Deadline while emitting progress updates  and handling errors appropriately.
        """
        error_message = ""
        num_successful_submissions = 0

        try:

            dialog = self.dialog

            if not dialog.deadline_settings_initialized:
                dialog.populate_deadline_settings()

            app_version = dialog.adobe.app.version
            version = '.'.join(app_version.split('x')[0].split('.')[:2])
            project_check = dialog.run_project_checks()

            if not project_check:
                self.finished.emit("Project checks failed", 0)
                return

            if dialog.ui.compTableWidget.rowCount() == 0:
                self.finished.emit("No render queue items available", 0)
                return

            current_deadline_settings = dialog.get_deadline_settings()

            deadline_error_message = ""
            previous_job_id = ""
            project_path = dialog.adobe.app.project.file.fsName
            num_rows = dialog.ui.compTableWidget.rowCount()

            for row in self.comp_rows:
                progress_idx = self.row_to_progress_idx.get(row, None)
                percent = int((row / max(1, num_rows)) * 100)
                if progress_idx is not None:
                    self.row_progress.emit(progress_idx, 0, "Starting submission...")

                render_queue_item = dialog.ui.compTableWidget.item(row, 0).data(QtCore.Qt.UserRole)
                includeCheckBox = dialog.ui.compTableWidget.item(row, 6)
                statusItem = dialog.ui.compTableWidget.item(row, 1)
                useCompNameCheckBox = dialog.ui.compTableWidget.item(row, 5)
                compName = render_queue_item.comp.name
                render_queue_template = dialog.get_render_queue_template(row)
                use_comp_name = useCompNameCheckBox.checkState() == QtCore.Qt.Checked

                # Checks

                # Item Included Check
                if includeCheckBox.checkState() != QtCore.Qt.Checked:
                    msg = f"Skipped because 'Include' checkbox is not checked for comp '{compName}'."
                    deadline_error_message += msg + "\n"
                    self.item_update.emit(row, msg, False)
                    if progress_idx is not None:
                        self.row_done.emit(progress_idx, False)
                    continue

                # Item Status Check
                if render_queue_item.status != dialog.adobe.RQItemStatus.QUEUED:
                    msg = f"Skipped because comp '{compName}' is not in QUEUED status (status={render_queue_item.status})."
                    deadline_error_message += msg + "\n"
                    self.item_update.emit(row, msg, False)
                    if progress_idx is not None:
                        self.row_done.emit(progress_idx, False)
                    continue

                # Item Spaces check
                if compName.startswith(" ") or compName.endswith(" "):
                    msg = f"Comp name '{compName}' has spaces at the front or back. Skipping."
                    deadline_error_message += msg + "\n"
                    self.item_update.emit(row, msg, False)
                    if progress_idx is not None:
                        self.row_done.emit(progress_idx, False)
                    continue

                # Iterm Special Character check
                if not re.match(r'^[a-zA-Z0-9_]+$', compName):
                    msg = f"Comp name '{compName}' has special characters. Only letters, numbers, and underscores are allowed. Skipping."
                    deadline_error_message += msg + "\n"
                    self.item_update.emit(row, msg, False)
                    if progress_idx is not None:
                        self.row_done.emit(progress_idx, False)
                    continue

                # Output folder
                render_scene_file_path = dialog.get_shotgrid_template(render_queue_template, use_comp_name, compName, True)
                render_scene_file_directory = os.path.dirname(render_scene_file_path)

                # Step 1: Copy project file
                try:
                    if progress_idx is not None:
                        self.row_progress.emit(progress_idx, 10, f"Copying project file for {compName}...")
                    if not os.path.exists(render_scene_file_directory):
                        os.makedirs(render_scene_file_directory, exist_ok=True)
                    shutil.copy(dialog.adobe.app.project.file.fsName, render_scene_file_path)

                except Exception as e:
                    msg = f"Failed to create render scene backup for comp '{compName}': {str(e)}"
                    deadline_error_message += msg + "\n"
                    self.item_update.emit(row, msg, False)
                    if progress_idx is not None:
                        self.row_done.emit(progress_idx, False)
                    continue

                # Step 2: Generate manifest
                try:
                    if progress_idx is not None:
                        self.row_progress.emit(progress_idx, 30, f"Generating project manifest for {compName}...")
                    dialog.generate_project_manifest_file_jsx(render_queue_item, render_scene_file_path)
                    if progress_idx is not None:
                        self.row_progress.emit(progress_idx, 50, f"Generating comp manifest for {compName}...")
                    dialog.generate_manifest_file_for_queue_item_jsx(render_queue_item, render_scene_file_path)

                except Exception as e:
                    msg = f"Failed to generate manifest file for comp '{compName}': {str(e)}"
                    deadline_error_message += msg + "\n"
                    self.item_update.emit(row, msg, False)
                    if progress_idx is not None:
                        self.row_done.emit(progress_idx, False)
                    continue

                # Step 3: Submit to Deadline
                try:
                    self.progress.emit(percent, f"Submitting {compName} to Deadline...")

                    job_attrs, plugin_attrs = dialog.build_deadline_job_and_plugin_dicts(
                        render_queue_item=render_queue_item,
                        deadline_settings=current_deadline_settings,
                        project_path=project_path,
                        layers=False,
                        previous_job_id=previous_job_id,
                        version=version,
                        render_file=render_scene_file_path
                    )
                    dialog.submit_render_queue_item_to_deadlineconnect(job_attrs, plugin_attrs)
                    num_successful_submissions += 1
                    self.item_update.emit(row, f"Submitted comp '{compName}' to Deadline successfully.", True)
                    if progress_idx is not None:
                        self.row_done.emit(progress_idx, True)

                except Exception as e:
                    msg = f"Failed to submit comp '{compName}' to Deadline: {str(e)}"
                    deadline_error_message += msg + "\n"
                    self.item_update.emit(row, msg, False)
                    if progress_idx is not None:
                        self.row_done.emit(progress_idx, False)
                    continue

            # Final progress
            self.progress.emit(100, "Finished submitting to Deadline")
            if deadline_error_message:
                self.finished.emit(deadline_error_message, num_successful_submissions)
            else:
                self.finished.emit("", num_successful_submissions)

        except Exception as e:
            self.finished.emit(str(e), num_successful_submissions)

class DeadlineProgressDialog(QtGui.QDialog):
    """
        Dialog class for displaying the progress of Deadline submissions, with individual progress bars for each comp being submitted.
    """
    def __init__(self, comp_names, logger, parent=None):
        super().__init__(parent)
        self.logger = logger
        self.setWindowTitle("Deadline Submission Progress")
        self.setModal(True)
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
