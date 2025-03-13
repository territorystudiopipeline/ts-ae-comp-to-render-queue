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
import re
import time
from contextlib import contextmanager

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
        self.populate_presets()
        self.connect_signals_and_slots()

        # Create render queue items
        self.create_render_queue_items()

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

    def on_item_changed(self, item):
        """
        Slot to handle item changes and apply the changes to all selected rows.
        Currently not working
        """
        # Save the current selection
        selected_items = self.ui.compTableWidget.selectedItems()
        selected_rows = [item.row() for item in self.ui.compTableWidget.selectedItems()]
        logger.debug("Selected Rows: %s" % selected_rows)

        def update_item_signals(item, next_item, update_func):
            """
                Update the item signals
            """
            next_item.disableSignals()
            update_func()
            next_item.enableSignals()

        if len(selected_rows) > 1:
            # Apply the change to all selected rows
            for row in selected_rows:
                if row != item.row():
                    logger.debug("Applying changes to row: %s" % row)
                    if isinstance(item, QtGui.QTableWidgetItem):
                        logger.debug("Item is a QTableWidgetItem")
                        next_item = self.ui.compTableWidget.item(row, item.column())
                        if next_item:
                            # check if the item is using a check state else set the text
                            if item.checkState() == QtCore.Qt.Checked:
                                logger.debug("Updating Check State from %s to %s" % (next_item.checkState(), item.checkState()))
                                update_item_signals(item, next_item, lambda: next_item.setCheckState(item.checkState()))
                            elif item.checkState() == QtCore.Qt.Unchecked:
                                logger.debug("Updating Check State from %s to %s" % (item.checkState(), next_item.checkState()))
                                update_item_signals(item, next_item, lambda: next_item.setCheckState(item.checkState()))
                            else:
                                logger.debug("Updating Text from %s to %s" % (item.text(), next_item.text()))
                                update_item_signals(item, next_item, lambda: next_item.setText(item.text()))
                        else:
                            logger.debug("Next item is None for row: %s, column: %s" % (row, item.column()))
                    else:
                        cell_widget = self.ui.compTableWidget.cellWidget(item.row(), item.column())
                        if cell_widget:
                            if isinstance(cell_widget, QtGui.QLineEdit):
                                logger.debug("Cell Widget is a QLineEdit")
                                next_widget = self.ui.compTableWidget.cellWidget(row, item.column())
                                if next_widget:
                                    logger.debug("Updating Text from %s to %s" % (cell_widget.text(), next_widget.text()))
                                    update_item_signals(cell_widget, next_widget, lambda: next_widget.setText(cell_widget.text()))
                                else:
                                    logger.debug("Next widget is None for row: %s, column: %s" % (row, item.column()))

                            elif isinstance(cell_widget, QtGui.QComboBox):
                                logger.debug("Cell Widget is a QComboBox")
                                next_widget = self.ui.compTableWidget.cellWidget(row, item.column())
                                if next_widget:
                                    logger.debug("Updating Index from %s to %s" % (cell_widget.currentIndex(), next_widget.currentIndex()))
                                    update_item_signals(cell_widget, next_widget, lambda: next_widget.setCurrentIndex(cell_widget.currentIndex()))

                            elif isinstance(cell_widget, QtGui.QCheckBox):
                                logger.debug("Cell Widget is a QCheckBox")
                                next_widget = self.ui.compTableWidget.cellWidget(row, item.column())
                                if next_widget:
                                    logger.debug("Updating Check State from %s to %s" % (cell_widget.checkState(), next_widget.checkState()))
                                    update_item_signals(cell_widget, next_widget, lambda: next_widget.setCheckState(cell_widget.checkState()))

    def update_selected_rows(self):
        """
            Update the selected rows with the new item
            Currently not working
        """
        sender = self.sender()
        if not sender:
            return

        selected_rows = {index.row() for index in self.ui.compTableWidget.selectionModel().selectedRows()}
        if not selected_rows:
            return

        current_row = self.table.currentRow()
        current_col = self.table.currentColumn()

        # If the modified cell is a checkbox or text item
        if isinstance(sender, QtGui.QTableWidget):
            item = self.table.item(current_row, current_col)
            if item:
                new_value = item.text()
                new_check_state = item.checkState()

                for row in selected_rows:
                    if row != current_row:  # Avoid self-updating
                        update_item = self.table.item(row, current_col)
                        if update_item:
                            update_item.blockSignals(True)
                            update_item.setText(new_value)
                            update_item.setCheckState(new_check_state)
                            update_item.blockSignals(False)
        # If the modified cell is a combobox
        elif isinstance(sender, QtGui.QComboBox):
            new_text = sender.currentText()
            for row in selected_rows:
                if row != current_row:  # Avoid self-updating
                    combo = self.table.cellWidget(row, current_col)
                    if combo:
                        combo.blockSignals(True)
                        combo.setCurrentText(new_text)
                        combo.blockSignals(False)

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
            self.alert_box("No render queue items meet the criteria",
                           "Please add some render queue items to apply the changes to")
            return

        # Add the comps to the table
        for item in filtered_render_queue_items:
            self.add_table_row(item)

    def add_table_row(self, item):
        """
            Add a row to the table

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
        comp_name.setData(QtCore.Qt.UserRole, item)
        self.ui.compTableWidget.setItem(rowPosition, 0, comp_name)

        # Add the status
        status = "Unknown"
        icon = self.style().standardIcon(QtGui.QStyle.SP_MessageBoxCritical)

        if item.status == self.adobe.RQItemStatus.QUEUED:
            status = "Queued"
            icon = self.style().standardIcon(QtGui.QStyle.SP_DialogApplyButton)

        elif item.status == self.adobe.RQItemStatus.NEEDS_OUTPUT:
            status = "Needs Output"
            icon = self.style().standardIcon(QtGui.QStyle.SP_MessageBoxWarning)

        statusItem = QtGui.QTableWidgetItem(icon, status)
        statusItem.setFlags(statusItem.flags() & ~QtCore.Qt.ItemIsEditable)
        self.ui.compTableWidget.setItem(rowPosition, 1, statusItem)

        # Add Frame Range Input
        frameRangeLineEdit = QtGui.QLineEdit()
        self.ui.compTableWidget.setCellWidget(rowPosition, 2, frameRangeLineEdit)

        # Add the frame range ComboBox
        frameRangeComboBox = QtGui.QComboBox()
        frameRangeComboBox.insertItems(0,[self.WORK_AREA_TEXT, self.COMP_TEXT, self.SINGLE_FRAME_TEXT, self.CUSTOM_TEXT])
        self.ui.compTableWidget.setCellWidget(rowPosition, 3, frameRangeComboBox)

        # Add the render format dropdown
        renderFormatDropdown = QtGui.QComboBox()

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
        self.ui.compTableWidget.setItem(rowPosition, 5, useCompNameCheckBox)

        # Connect the signals and slots
        frameRangeComboBox.currentIndexChanged.connect(lambda: self.refresh_frame_range(frameRangeComboBox, frameRangeLineEdit, item))

        # Trigger the signal to set the default frame range
        frameRangeComboBox.emit(QtCore.SIGNAL("currentIndexChanged(int)"), 0)

        # Include the render queue item checkbox
        includeCheckBox = QtGui.QTableWidgetItem()
        includeCheckBox.setCheckState(QtCore.Qt.Checked)
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
        self.ui.addButton.clicked.connect(self.create_render_queue_items)
        self.ui.applyButton.clicked.connect(self.apply_to_render_queue_items)
        self.ui.refreshButton.clicked.connect(self.create_table_entries)
        self.ui.cancelButton.clicked.connect(self.close)
        #TODO: Get multiple selection updates working
        #self.ui.compTableWidget.itemChanged.connect(self.on_item_changed)

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
            self.adobe.app.executeCommand(self.adobe.app.findMenuCommandId("Add to Render Queue"))
            self.create_table_entries()
            # Resize UI to fit the table
            self.ui.compTableWidget.resizeColumnsToContents()

    def apply_to_render_queue_items(self):
        """
            Apply the changes to the render queue items
        """
        # Get the selected comps
        # Debugging time stamp for testing HH:MM:SS
        self.start_time = time.time()
        logger.debug("Start Render Queue Items Time: %s" % time.strftime("%H:%M:%S"))

        logger.debug("Applying to render queue items")
        # Check if there are any render queue items
        if self.ui.compTableWidget.rowCount() == 0:
            self.alert_box("No render queue items", "Please add some render queue items to apply the changes to")
            return

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
                    pass

                # Check the template actually exists
                if not self.check_template_exists(comp, frame_range, render_queue_template, templateName):
                    self.alert_box("Error", "Something went wrong applying or locating an output template")

                try:
                    render_queue_item.outputModule(render_queue_item.numOutputModules).applyTemplate(templateName)
                except:
                    self.alert_box("Error",
                                   "There's some kind of issue with this template\n\n" + str(templateName) + '\n' + str(
                                       render_queue_template))
                    return
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

                    # Get the comp name
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

        # Debugging time stamp for testing HH:MM:SS
        logger.debug("Finish Time: %s" % time.strftime("%H:%M:%S"))
        logger.debug("Total Time: %s" % (time.time() - self.start_time))

        self.message_box( 'Apply To Render Queue Items', 'Successfully updated %d render queue items' % count)
        self.close()

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
        if os.path.basename(render_queue_template).startswith('mov'):
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
        importedProject = self.importPresetProject(render_queue_template)

        # Get preset render queue item
        presetRenderQueueItem = self.findRenderQueueItemByCompName('PRESET')
        if presetRenderQueueItem is None:
            return False

        presetRenderQueueItem.outputModule(presetRenderQueueItem.numOutputModules).saveAsTemplate(templateName)
        importedProject.remove()

        return True

    def findRenderQueueItemByCompName(self, comp_name):
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

    def importPresetProject(self, render_queue_template):
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