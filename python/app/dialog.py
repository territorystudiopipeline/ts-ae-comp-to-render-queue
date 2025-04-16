# Copyright (c) 2013 Shotgun Software Inc.
#
# CONFIDENTIAL AND PROPRIETARY
#
# This work is provided "AS IS" and subject to the Shotgun Pipeline Toolkit
# Source Code License included in this distribution package. See LICENSE.
# By accessing, using, copying or modifying this work you indicate your
# agreement to the Shotgun Pipeline Toolkit Source Code License. All rights
# not expressly granted therein are reserved by Shotgun Software Inc.
import datetime
import shutil
import subprocess
import sgtk
import os
import sys
import re
import time
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

        self.current_project = self._app.context.project
        self.project_id = self.current_project["id"]
        self.project_name = self.current_project["name"]
        self.project_code = None

        logger.debug(f"Current Project: {self.current_project}")
        logger.debug(f"Project ID: {self.project_id}")

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
            Populate the render format dropdown with the available presets
        """
        self.presets = {}
        for preset_item in self._app.get_setting("render_presets"):
            # use an internal method to resolve the path of the ae template files
            resolved_path = self._app._TankBundle__resolve_hook_expression(preset_item['name'], preset_item['path'])
            self.presets[preset_item['name']] = resolved_path[0]
            if widget:
                widget.insertItems(-1, [preset_item['name']])

    def connect_signals_and_slots(self):
        """
            Connect the signals and slots
        """
        # Connect the buttons
        self.ui.submitButton.clicked.connect(self.apply_and_submit)
        self.ui.addButton.clicked.connect(self.create_render_queue_items)
        self.ui.applyButton.clicked.connect(self.apply_to_render_queue_items)
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

    def apply_to_render_queue_items(self):
        """
            Apply the changes to the render queue items
        """
        self.toggle_buttons(False)
        # Get the selected comps
        # Debugging time stamp for testing HH:MM:SS
        self.start_time = time.time()
        logger.debug("Start Render Queue Items Time: %s" % time.strftime("%H:%M:%S"))

        logger.debug("Applying to render queue items")
        # Check if there are any render queue items
        if self.ui.compTableWidget.rowCount() == 0:
            self.alert_box("No render queue items", "Please add some render queue items to apply the changes to")
            return

        # Show the progress bar
        self.show_progress_bar(format_text="Applying to render queue items...")

        # Get the render queue items from the table
        count = 0
        for row in range(self.ui.compTableWidget.rowCount()):
            tableItem = self.ui.compTableWidget.item(row, 0)
            statusItem = self.ui.compTableWidget.item(row, 1)
            frameRangeLineEdit = self.ui.compTableWidget.cellWidget(row, 2)
            frameRangeComboBox = self.ui.compTableWidget.cellWidget(row, 3)
            renderFormatDropdown = self.ui.compTableWidget.cellWidget(row, 4)
            useCompNameCheckBox = self.ui.compTableWidget.item(row, 5)
            includeCheckBox = self.ui.compTableWidget.item(row, 6)

            # Emit progress
            self.update_progress_bar(int(row / self.ui.compTableWidget.rowCount() * 100))

            # Check if the item is checked
            if includeCheckBox.checkState() == QtCore.Qt.Checked:
                render_queue_item = tableItem.data(QtCore.Qt.UserRole)
                comp = render_queue_item.comp
                compName = comp.name
                templateName = renderFormatDropdown.currentText()
                render_queue_template = self.get_render_queue_template(row)
                frame_range = self.get_frame_range(comp, row)

                # Check if the frame range is valid
                if frame_range[0] is None or frame_range[1] is None:
                    logger.debug("Bad frame range, skipping %s" % comp.name)
                    self.alert_box("Bad frame range", "Please check the frame range for %s, Skipping" % comp.name)
                    # Update the status icon
                    statusItem.setIcon(self.ui.errorIcon)
                    statusItem.setToolTip("Template Not Applied | Error - Bad Frame Range")
                    includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                    continue

                # Check the template actually exists
                if not self.check_template_exists(comp, frame_range, render_queue_template, templateName):
                    self.alert_box("Error", "Something went wrong applying or locating an output template")
                    logger.debug("Error applying template %s" % templateName)
                    # Update the status icon
                    statusItem.setIcon(self.ui.errorIcon)
                    statusItem.setToolTip("Template Not Applied | Error - Template Not Found")
                    includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                    continue

                try:
                    render_queue_item.outputModule(render_queue_item.numOutputModules).applyTemplate(templateName)
                except:
                    self.alert_box("Error",
                                   "There's some kind of issue with this template\n\n" + str(templateName) + '\n' + str(
                                       render_queue_template))
                    logger.debug("Error applying template %s" % templateName)
                    # Update the status icon
                    statusItem.setIcon(self.ui.errorIcon)
                    statusItem.setToolTip("Template Not Applied | Error")
                    includeCheckBox.setCheckState(QtCore.Qt.Unchecked)
                    continue

                # Check current time span
                timespan = render_queue_item.getSetting("Time Span")
                logger.debug("Time Span: %s" % timespan)

                # Set the render to the start/end times
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

                # Grab the output folder from templates
                outputLocation = self.get_shotgrid_template(render_queue_template)

                # Create the output folder if it doesn't already exist
                folderPath = os.path.dirname(outputLocation)
                if not os.path.exists(folderPath):
                    os.makedirs(folderPath)

                # Debugging
                logger.debug("Output location: %s" % outputLocation)
                logger.debug("Output folder: %s" % folderPath)
                logger.debug("Comp Name: %s" % comp.name)

                # Replace output location with comp name if checkbox is checked
                if useCompNameCheckBox.checkState() == QtCore.Qt.Checked:
                    # Get the original output file and strip the folder path
                    originalOutputFile = outputLocation.replace(folderPath, '')

                    # Get the filename
                    fileName = self.adobe.app.project.file.name

                    # EntityName _ Name _v VersionNumber FileExtension
                    match = re.match(r'(.*)(_)(.*)(_v)(\d\d\d)(.*)', fileName)

                    name = match.group(3)
                    version = int(match.group(5))
                    compName = comp.name

                    # Join the comp name with the version number
                    newFileName = "%s_v%03d" % (compName, version)

                    # Replace the first group before the first . with the comp name
                    newOutputFile = re.sub(r'([^.]+)', newFileName, originalOutputFile, 1)

                    # Debugging
                    logger.debug("Original Output File: %s" % originalOutputFile)
                    logger.debug("New Output File: %s" % newOutputFile)

                    # Rebuild the output location
                    outputLocation = os.path.join(folderPath, compName, newOutputFile)
                    logger.debug("Output location: %s" % outputLocation)

                    # Create the output folder if it doesn't already exist
                    folderPath = os.path.dirname(outputLocation)
                    if not os.path.exists(folderPath):
                        os.makedirs(folderPath)

                # If Single Frame is selected, Remove the frame range from the output location
                if frameRangeComboBox.currentText() == self.SINGLE_FRAME_TEXT:
                    # Remove .[####]. from the output location
                    outputLocation = re.sub(r'\.\[?\#*\]?\.*', '.', outputLocation)

                # Set the filepath and name on the newly created output module
                # Do it twice because it sometimes fails the first time - Sean
                with self.supress_dialogs():
                    render_queue_item.outputModule(render_queue_item.numOutputModules).file = self.adobe.File(
                        outputLocation)
                    render_queue_item.outputModule(render_queue_item.numOutputModules).file = self.adobe.File(
                        outputLocation)

                # Log
                logger.debug("Render Queue Item for: %s has been updated" % render_queue_item.comp.name)
                count += 1

                # Update the status icon
                statusItem = self.ui.compTableWidget.item(row, 1)
                statusItem.setIcon(self.ui.clearIcon)
                statusItem.setToolTip("Template Applied | Ready to Render")

        # Show progress at 100%
        self.update_progress_bar(100)
        time.sleep(0.2)
        self.hide_progress_bar()

        self.toggle_buttons()

        # Debugging time stamp for testing HH:MM:SS
        logger.debug("Finish Time: %s" % time.strftime("%H:%M:%S"))
        logger.debug("Total Time: %s" % (time.time() - self.start_time))
        logger.debug("Total Render Queue Items Updated: %s" % count)

    def get_frame_range(self, comp, row):
        """
            Get the frame range to render

            :param comp: The comp to get the frame range for

            :returns: A list containing the start and end frame to render
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

            :returns: The render queue template to use for the render queue item
        """
        render_queue_template = None

        userSelection = self.ui.compTableWidget.cellWidget(row, 4).currentText()
        if userSelection in self.presets:
            render_queue_template = self.presets[userSelection]

        return render_queue_template

    def alert_box(self, title, text):
        """
            Display an alert box

            :param title: The title of the alert box
            :param text: The text of the alert box
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

            :param title: The title of the warning box
            :param text: The text of the warning box
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

            :param title: The title of the message box
            :param text: The text of the message box
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

    def get_shotgrid_template(self, render_queue_template):
        """
            Get the output location from the render queue template

            :param render_queue_template: The template to use for the render queue item

            :returns: The output location for the render queue item
        """
        template_file_name = os.path.basename(render_queue_template)

        if template_file_name.startswith('mov'):
            templateName = self._app.get_setting("mov_render_template")
        else:
            templateName = self._app.get_setting("seq_render_template")

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

        # Add in a %04d number if it's a sequence then strip it out to be [####] for AE
        if 'SEQ' in template.keys:
            fields['SEQ'] = 9999
            outputPath = template.apply_fields(fields)
            outputPath = outputPath.replace('.9999.', '.[####].')

        else:
            outputPath = template.apply_fields(fields)

        return outputPath

    def check_template_exists(self, comp, frame_range, render_queue_template, templateName):
        """
            Check that the template exists, if not create it
        """
        # Add the comp to the render queue
        renderQueueItem = self.adobe.app.project.renderQueue.items.add(comp)

        # If the output module template already exists, just apply it, otherwise import the prest project, save the new template, clean up, and then apply it
        if templateName in renderQueueItem.outputModule(renderQueueItem.numOutputModules).templates:
            renderQueueItem.outputModule(renderQueueItem.numOutputModules).applyTemplate(templateName)
            renderQueueItem.remove()
            return True

        # Remove the comp from the render
        renderQueueItem.remove()

        # Import the preset project
        importedProject = self.import_preset_project(render_queue_template)

        # Get preset render queue item
        presetRenderQueueItem = self.find_render_queue_item_by_comp_name('PRESET')
        if presetRenderQueueItem is None:
            return False

        presetRenderQueueItem.outputModule(presetRenderQueueItem.numOutputModules).saveAsTemplate(templateName)
        importedProject.remove()

        return True

    def find_render_queue_item_by_comp_name(self, comp_name):
        """
            Find a render queue item by the comp name

            :param comp_name: The name of the comp to search for

            :returns: The render queue item
        """
        renderQueueItems = self.adobe.app.project.renderQueue.items
        for i in range(1, self.adobe.app.project.renderQueue.numItems+1):
            if renderQueueItems[i].comp.name == comp_name:
                return renderQueueItems[i]

        return None

    def import_preset_project(self, render_queue_template):
        """
            Import the preset project

            :param render_queue_template: The template to use for the render queue item

            :returns: The imported project
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

            :param state: The state to set the buttons to
        """
        self.ui.applyButton.setEnabled(state)
        self.ui.addButton.setEnabled(state)
        self.ui.submitButton.setEnabled(state)

    #####################################################################################################
    # Events
    #####################################################################################################
    def eventFilter(self, obj, event):
        """
        Event filter to void mouse scroll events for specific widgets.

        :param obj: The object to filter events for
        :param event: The event to filter
        :returns: True if the event was handled, False otherwise
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

            :returns: The row under the mouse cursor
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
            Match the selected rows to the row
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

    def show_progress_bar(self, format_text=None):
        """
            Show the progress bar

            :param format_text: The text to display in the progress bar
        """
        self.ui.progressBar.setVisible(True)
        self.ui.progressBar.setValue(0)
        if format_text:
            self.ui.progressBar.setFormat(f"{format_text} %p%")

    def hide_progress_bar(self):
        """
            Hide the progress bar
        """
        self.ui.progressBar.setVisible(False)
        self.ui.progressBar.setValue(0)
        self.ui.progressBar.setFormat("%p%")

    def update_progress_bar(self, value):
        """
            Update the progress bar
            :param value: The value to set the progress bar to
        """
        self.ui.progressBar.setValue(value)

    ######################################################################################################
    # Deadline
    ######################################################################################################
    # Deadline Settings
    def get_deadline_settings(self):
        """
            Get the current Deadline settings from the UI

            :returns: A dictionary containing the Deadline settings
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
    # Deadline Submission
    ####################################
    def submit_to_deadline(self, save_project=True):
        """
            Submit the render queue items to Deadline
        """
        self.toggle_buttons(False)

        logger.info("Submitting to Deadline")

        if not self.deadline_settings_initialized:
            logger.debug("Lazy loading Deadline settings")
            self.populate_deadline_settings()

        app_version = self.adobe.app.version

        # Extract the major.minor version number
        version = '.'.join(app_version.split('x')[0].split('.')[:2])

        ## Run the project checks ##
        logger.debug("Running project checks")

        # If there is an active project
        if not self.adobe.app.project:
            self.warning_box("No project open", "Please open a project before submitting to Deadline")
            return

        # If the project is not saved
        if not self.adobe.app.project.file:
            self.warning_box("Project not saved", "Please save the project before submitting to Deadline")
            return

        # If there are no render queue items
        if self.adobe.app.project.renderQueue.numItems == 0:
            self.warning_box("No render queue items", "Please add some render queue items to submit to Deadline")
            return

        # If there are rows in the table
        if self.ui.compTableWidget.rowCount() == 0:
            self.warning_box("No render queue items available", "Please add some render queue items or refresh the table to submit to Deadline")
            return

        logger.debug("Project Checks passed, proceeding with submission")

        if save_project == QtGui.QMessageBox.Yes:
            # Save and create a copy of the current project
            self.adobe.app.project.save()
            logger.info('Project saved')
        else:
            logger.info('Project not saved')

        logger.debug("Creating backup of project")
        backup_location = os.path.join(os.path.dirname(self.adobe.app.project.file.fsName), "deadline_submission_backup")
        if not os.path.exists(backup_location):
            # Create the backup directory if it doesn't exist
            logger.debug("Creating backup directory: %s" % backup_location)
            os.makedirs(backup_location, exist_ok=True)

        backup_file = os.path.join(backup_location, self.adobe.app.project.file.name)
        backup_file = backup_file.replace('.aep', '_%s.aep' % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))

        # Copy the project to the backup location
        try:
            logger.debug("Copying project to backup location: %s" % backup_file)
            shutil.copy(self.adobe.app.project.file.fsName, backup_file)
            logger.info('Backup created: %s' % backup_file)

        except Exception as e:
            logger.error("Failed to create backup: %s" % e)
            error = traceback.format_exc()
            logger.error("%s" % error)

        # Get Deadline Settings
        current_deadline_settings = self.get_deadline_settings()
        logger.debug("Deadline settings: %s" % current_deadline_settings)

        # Submission variables
        self.deadline_error_message = ""
        total_jobs = self.ui.compTableWidget.rowCount()
        previous_job_id = ""
        num_successful_submissions = 0
        project_path = self.adobe.app.project.file.fsName

        self.show_progress_bar(format_text="Submitting to Deadline")

        # Submit the selected rows to Deadline
        for row in range(self.ui.compTableWidget.rowCount()):

            # Emit progress
            self.update_progress_bar(int(row / self.ui.compTableWidget.rowCount() * 100))

            render_queue_item = self.ui.compTableWidget.item(row, 0).data(QtCore.Qt.UserRole)
            includeCheckBox = self.ui.compTableWidget.item(row, 6)
            statusItem = self.ui.compTableWidget.item(row, 1)

            logger.debug("Running render queue item Checks: %s" % render_queue_item.comp.name)

            # Check if the render queue item is checked
            if includeCheckBox.checkState() != QtCore.Qt.Checked:
                logger.debug("Render queue item is not checked, skipping")
                continue

            # Check render queue item status
            if render_queue_item.status != self.adobe.RQItemStatus.QUEUED:
                logger.debug("Render queue item is not queued, skipping")
                continue

            # Check comp name for spaces trailing or leading
            if render_queue_item.comp.name.startswith(" ") or render_queue_item.comp.name.endswith(" "):
                logger.debug("Comp name has spaces at the front or back, skipping")
                self.deadline_error_message += "Comp name has spaces at the front or back, skipping - %s\n" % render_queue_item.comp.name

                # Update the status icon
                statusItem.setIcon(self.ui.errorIcon)
                statusItem.setToolTip("Comp name has spaces at the front or back - Cannot submit to Deadline")
                continue

            # Check comp name for special characters
            if not re.match(r'^[a-zA-Z0-9_]+$', render_queue_item.comp.name):
                logger.debug("Comp name has special characters, skipping")
                self.deadline_error_message += "Comp name has special characters, skipping - %s\n" % render_queue_item.comp.name

                # Update the status icon
                statusItem.setIcon(self.ui.errorIcon)
                statusItem.setToolTip("Comp name has special characters - Cannot submit to Deadline")
                continue

            logger.debug("Render queue item checks passed")

            try:
                # Submit the render queue item to Deadline
                logger.debug("Submitting render queue item: %s" % render_queue_item.comp.name)
                self.submit_render_queue_item_to_deadline(render_queue_item=render_queue_item,
                                                          deadline_settings=current_deadline_settings,
                                                          project_path=project_path,
                                                          layers=False,
                                                          previous_job_id=previous_job_id,
                                                          version=version)

                logger.debug("Render queue item submitted to Deadline: %s" % render_queue_item.comp.name)

                # Update Status
                statusItem = self.ui.compTableWidget.item(row, 1)
                statusItem.setIcon(self.ui.submitIcon)
                statusItem.setToolTip("Submitted to Deadline")
                num_successful_submissions += 1

            except Exception as e:
                logger.error("Failed to submit render queue item to Deadline: %s" % e)
                error = traceback.format_exc()
                logger.error(error)
                self.deadline_error_message += "Failed to submit render queue item to Deadline: %s\n" % render_queue_item.comp.name

                # Update Status
                statusItem.setIcon(self.ui.errorIcon)
                statusItem.setToolTip("Failed to submit to Deadline - %s" % e)

        # Show progress at 100%
        self.update_progress_bar(100)
        time.sleep(0.2)
        self.hide_progress_bar()

        self.toggle_buttons()
        logger.info("Finished submitting to Deadline")

        # Show all submission results on one message box
        if self.deadline_error_message:
            self.message_box("Deadline Submission", "Submission completed with errors:\n\n%s" % self.deadline_error_message)
        else:
            self.message_box("Deadline Submission", "Submission completed successfully.\n\n%d jobs submitted to Deadline." % num_successful_submissions)

    def submit_render_queue_item_to_deadline(self, render_queue_item, deadline_settings, project_path, layers, previous_job_id, version):
        """
            Submit the render queue item to Deadline

            Layer submission is not supported yet, so this is a placeholder for now
            Iv removed some of the checks and statements that were inplace for older after effects versions
            I have also kept the code and structure as close to the original submitter as possible for now

            :param render_queue_item: The render queue item to submit
            :param deadline_settings: The Deadline settings to use for the submission
            :param project_path: The path to the project file
            :param layers: Whether to submit layers or not
            :param previous_job_id: The previous job ID to use for the submission
            :param version: The version of After Effects
        """
        logger.debug("Submitting render queue item to Deadline: %s" % render_queue_item.comp.name)

        # Output Variables
        output_module = render_queue_item.outputModule(render_queue_item.numOutputModules)
        output_file_name = output_module.file.name
        output_folder = os.path.dirname(output_file_name)

        # Paths
        temp_folder = os.path.expanduser("~\\temp\\")
        os.makedirs(temp_folder, exist_ok=True)
        submit_info_path = os.path.join(temp_folder, "ae_submit_info.job")
        plugin_info_path = os.path.join(temp_folder, "ae_plugin_info.job")

        # Submission variables
        project_name = os.path.basename(self.adobe.app.project.file.name)
        job_name = f"{project_name} - {render_queue_item.comp.name}"
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

        # Check if the output module is rendering a movie
        is_movie = any(output_file_name.endswith(ext) for ext in self.deadline_defaults["movie_formats"])

        if override_frame_list or multi_machine:
            # Get the frame duration and start/end times
            frame_duration = render_queue_item.comp.frameDuration
            frame_offset = self.adobe.app.project.displayStartFrame
            display_start_time = render_queue_item.comp.displayStartTime

            if display_start_time:
                start_frame = frame_offset + round(display_start_time / frame_duration) + round(render_queue_item.timeSpanStart / frame_duration)
                end_frame = start_frame + round(render_queue_item.timeSpanDuration / frame_duration) - 1
                frame_list = f"{start_frame}-{end_frame}"

            if first_and_last and not multi_machine:
                frame_list = f"{start_frame},{end_frame},{frame_list}"

        current_job_dependencies = deadline_settings.get('job_dependencies', "")
        if dependent_comps and dependent_job_id is not "":
            if current_job_dependencies == "":
                current_job_dependencies = dependent_job_id
            else:
                current_job_dependencies = f"{dependent_job_id},{current_job_dependencies}"

        if multi_machine:
            job_name = f"{job_name} (multi-machine rendering {frame_list})"

        # Create submission info file
        with open(submit_info_path, "w") as submit_info:
            submit_info.write(f"Plugin=AfterEffects\n")
            submit_info.write(f"Name={job_name}\n")
            if dependent_comps:
                submit_info.write(f"BatchName={project_name}\n")
            submit_info.write(f"Comment={deadline_settings.get('comment', '')}\n")
            submit_info.write(f"Department={deadline_settings.get('department', self.project_name or '')}\n")
            submit_info.write(f"Group={deadline_settings.get('group', 'ae')}\n")
            submit_info.write(f"Pool={deadline_settings.get('pool', 'none')}\n")
            submit_info.write(f"SecondaryPool={deadline_settings.get('secondary_pool', 'none')}\n")
            submit_info.write(f"Priority={deadline_settings.get('priority', 30)}\n")
            submit_info.write(f"TaskTimeoutMinutes={deadline_settings.get('task_timeout_minutes', '0')}\n")
            submit_info.write(f"LimitGroups={deadline_settings.get('limit_groups', '')}\n")
            submit_info.write(f"ConcurrentTasks={deadline_settings.get('concurrent_tasks', '1')}\n")
            submit_info.write(f"LimitConcurrentTasksToNumberOfCpus={deadline_settings.get('limit_concurrent_tasks_to_number_cpus', '0')}\n")
            submit_info.write(f"JobDependencies={current_job_dependencies}\n")
            submit_info.write(f"OnJobComplete={deadline_settings.get('on_job_complete', 'Nothing')}\n")
            submit_info.write(f"FailureDetectionTaskErrors={deadline_settings.get('failure_detection_task_errors', 8)}\n")
            submit_info.write(f"FailureDetectionJobErrors={deadline_settings.get('failure_detection_job_errors', 20)}\n")

            # Blacklist
            if deadline_settings.get('submit_allow_list_as_deny_list', False):
                submit_info.write(f"Blacklist={deadline_settings['machine_list']}\n")
            else:
                submit_info.write(f"Whitelist={deadline_settings['machine_list']}\n")

            # Submit Suspended
            if deadline_settings.get('submit_suspended', False):
                submit_info.write(f"InitialStatus=Suspended\n")

            # Frame List
            if not is_movie and multi_machine:
                submit_info.write(f"Frames=1-{round(deadline_settings.get('multi_machine_tasks', False))}\n")
            else:
                submit_info.write(f"Frames={frame_list}\n")

            # Output file for all output modules
            index = 0
            for i in range(1, render_queue_item.numOutputModules + 1):
                output_module = render_queue_item.outputModule(i)
                output_file_name = output_module.file.name
                submit_info.write(f"OutputFilename{index}={render_queue_item.outputModule(i).file.fsName.replace('[#', '#').replace('#]', '#')}\n")

            # Movie settings
            if is_movie:
                # Override these settings for movie formats
                submit_info.write(f"MachineLimit=1\n")
                submit_info.write(f"ChunkSize=1000000\n")
            else:
                if multi_machine:
                    submit_info.write(f"MachineLimit=0\n")
                    submit_info.write(f"ChunkSize=1\n")
                else:
                    submit_info.write(f"MachineLimit={deadline_settings.get('machine_limit', 0)}\n")
                    submit_info.write(f"ChunkSize={deadline_settings.get('chunk_size', 15)}\n")

            if multi_machine:
                submit_info.write(f"ExtraInfoKeyValue0=FrameRangeOverride={frame_list}\n")

        # Create plugin info file
        with open(plugin_info_path, "w") as plugin_info:

            plugin_info.write(f"SceneFile={project_path}\n")
            plugin_info.write(f"Comp={comp_name}\n")

            if deadline_settings.get('include_output_path', False):
                plugin_info.write(f"Output={output_folder}\n")

            if multi_machine:
                plugin_info.write(f"MultiMachineMode=True\n")
                plugin_info.write(f"MultiMachineStartFrame={start_frame}\n")
                plugin_info.write(f"MultiMachineEndFrame={end_frame}\n")

            plugin_info.write(f"Version={version}\n")
            plugin_info.write(f"SubmittedFromVersion={self.adobe.app.version}\n")
            plugin_info.write(f"IgnoreMissingLayerDependenciesErrors={deadline_settings.get('missing_layers', False)}\n")
            plugin_info.write(f"IgnoreMissingEffectReferencesErrors={deadline_settings.get('missing_effects', False)}\n")
            plugin_info.write(f"FailOnWarnings={deadline_settings.get('fail_on_warnings', False)}\n")

            if not multi_machine:
                min_file_size = deadline_settings.get('file_size', 0)
                delete_files_under_min_size = deadline_settings.get('delete_files', False)

                if deadline_settings.get('fail_on_missing_file', False):
                    min_file_size = max(1, round(deadline_settings.get('file_size', 0)))
                    delete_files_under_min_size = True

                plugin_info.write(f"MinFileSize={min_file_size}\n")
                plugin_info.write(f"DeleteFilesUnderMinSize={delete_files_under_min_size}\n")

                if deadline_settings.get('include_output_path', False):
                    plugin_info.write(f"LocalRendering={deadline_settings.get('local_rendering', False)}\n")

            # Fail on existing AE process
            plugin_info.write(f"OverrideFailOnExistingAEProcess={deadline_settings.get('override_fail_on_existing_ae_process', False)}\n")
            plugin_info.write(f"FailOnExistingAEProcess={deadline_settings.get('fail_on_existing_ae_process', False)}\n")

            plugin_info.write(f"MemoryManagement={deadline_settings.get('memory_management', False)}\n")
            plugin_info.write(f"ImageCachePercentage={round(deadline_settings.get('image_cache_percentage', 0))}\n")
            plugin_info.write(f"MaxMemoryPercentage={round(deadline_settings.get('max_memory_percentage', 0))}\n")

            # Multi-process
            plugin_info.write(f"MultiProcess={deadline_settings.get('multi_process', False)}\n")

            plugin_info.write(f"ContinueOnMissingFootage={deadline_settings.get('missing_footage', False)}\n")


        # Submit to Deadline
        args = [submit_info_path, plugin_info_path]
        results = self.call_deadline_command(args)

        logger.info("Deadline submission results: %s" % results)

    def apply_and_submit(self):
        """
            Apply the settings and submit to Deadline
        """
        # Open dialog to ask the user if they want to save the project
        save_project = QtGui.QMessageBox.question(
            self,
            "Save Project",
            "Do you want to save the project before submitting to Deadline?",
            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No,
            QtGui.QMessageBox.Yes,
        )

        logger.debug("Applying settings and submitting to Deadline")
        self.apply_to_render_queue_items()
        self.submit_to_deadline(save_project=save_project)

    ####################################
    # Deadline Lists
    ####################################
    def get_machine_list(self):
        """
            Get the machine list from Deadline
            :returns: A list of machines
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
            :returns: A list of pools
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
            :returns: A list of groups
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
