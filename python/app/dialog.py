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

    def get_selected_comps(self):
        comps = []

        # Have to search through all items in the scene because simply calling
        # project.selection returns read only objects with reduced properties
        for i in range(1, self.adobe.app.project.numItems+1):
            theItem = self.adobe.app.project.item(i)
            if theItem.data['instanceof'] == 'CompItem' and theItem.selected:
                comps.append(theItem)

        return comps

    def populate_widgets(self):
        """
            Populate the widgets with the default values
        """
        self.populate_frame_range()
        self.popular_frame_range_options()
        self.populate_presets()

    def populate_frame_range(self):
        """
            Populate the frame range line edit with the default values
        """
        self.ui.frameRangeLineEdit.setText("%d - %d" % (self.first_frame, self.last_frame))

    def popular_frame_range_options(self):
        """
            Populate the frame range combo box with the default options
        """
        self.ui.frameRangeComboBox.insertItems(0, [self.COMP_TEXT, self.CUSTOM_TEXT])
        self.ui.frameRangeLineEdit.setEnabled(False)

    def populate_presets(self):
        """
            Populate the render format dropdown with the available presets
        """
        self.presets = {}
        for preset_item in self._app.get_setting("render_presets"):
            # use an internal method to resolve the path of the ae template files
            resolved_path = self._app._TankBundle__resolve_hook_expression(preset_item['name'], preset_item['path'])
            self.presets[preset_item['name']] = resolved_path[0]
            self.ui.renderFormatDropdown.insertItems(-1, [preset_item['name']])

    def connect_signals_and_slots(self):
        """
            Connect the signals and slots
        """
        self.ui.frameRangeComboBox.currentIndexChanged.connect(self.refresh_frame_range)
        self.ui.cancelButton.clicked.connect(self.close)
        self.ui.addButton.clicked.connect(self.create_render_queue_items)

    def refresh_frame_range(self):
        """
            Enable the frame range line edit if the custom option is selected
        """
        if self.ui.frameRangeComboBox.currentText() == self.CUSTOM_TEXT:
            self.ui.frameRangeLineEdit.setEnabled(True)
        else:
            self.ui.frameRangeLineEdit.setEnabled(False)

    def create_render_queue_items(self):
        """
            Create a render queue item for each of the selected comps
        """
        # Get the selected comps
        selected_comps = self.get_selected_comps()
        logger.debug("Selected comps: %s" % selected_comps)
        logger.debug("Selected comps: %s" % len(selected_comps))

        # Check if any comps are selected
        if len(selected_comps) == 0:
            self.alert_box("No comps selected", "Please select one or more comps to add to the render queue")
            return

        count = 0
        render_queue_template = self.get_render_queue_template()
        if render_queue_template is None:
            self.alert_box("Error", "Failed to find selected render preset")
            return
        # Suppress dialogs
        self.adobe.app.beginSuppressDialogs()

        for comp in selected_comps:

            # Get the frame range to render
            frame_range = self.get_frame_range(comp)
            if frame_range[0] is None or frame_range[1] is None:
                logger.debug("Bad frame range, skipping %s" % comp.name)
                self.alert_box("Bad frame range", "Please check the frame range for %s, Skipping" % comp.name)
                pass

            # Create a render queue item for each of the selected comps
            self.create_render_queue_item_for_comp(comp, frame_range, render_queue_template)
            count += 1

        # End Suppress Dialogs
        self.adobe.app.endSuppressDialogs()

        self.message_box( 'Add Comps To Render Queue', 'Successfully Added %d comps to the render queue' % count)
        self.close()

    def get_frame_range(self, comp):
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

        # Use comp frame range
        if self.ui.frameRangeComboBox.currentText() == self.COMP_TEXT:
            logger.debug("Using comp frame range")
            # Get the work area start and end times
            startFrame = comp.workAreaStart
            endFrame = (startFrame + comp.duration)

            # Debugging info
            logger.debug("Start Time: %s" % startFrame)
            logger.debug("End Time: %s" % endFrame)

            # Convert to frame numbers
            # Check if the comp work area has a start frame of 0
            logger.debug("Checking if start time is 0")
            if int(startFrame) == 0:
                startFrameNum = comp.displayStartFrame

            else:
                startFrameNum = int((startFrame / comp.frameDuration)) + comp.displayStartFrame

            # Start frame
            logger.debug("Start Frame: %s" % startFrameNum)

            # End frame
            endFrameNum = int((endFrame / comp.frameDuration)) + comp.displayStartFrame
            logger.debug("End Frame: %s" % endFrameNum)

            #endFrame = int(comp.frameRate * comp.workAreaDuration)
            #endFrame = int(comp.frameRate * comp.duration + 0.0001)
        else:
            rawText = self.ui.frameRangeLineEdit.text()
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


        return [startFrame, endFrame]

    def get_render_queue_template(self):
        """
            Get the render queue template to use for the render queue item

            :returns: The render queue template to use for the render queue item
        """
        render_queue_template = None

        userSelection = self.ui.renderFormatDropdown.currentText()
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

    def create_render_queue_item_for_comp(self, comp, frame_range, render_queue_template):
        """
            Create a render queue item for each of the selected comps

            :param comp: The comp to add to the render queue
            :param frame_range: The frame range to render
            :param render_queue_template: The template to use for the render queue item
        """
        templateName = self.ui.renderFormatDropdown.currentText()

        # Check the template actually exists
        if not self.check_template_exists(comp, frame_range, render_queue_template, templateName):
            self.alert_box("Error", "Something went wrong applying or locating an output template")

        renderQueueItem = self.adobe.app.project.renderQueue.items.add(comp)
        try:
            renderQueueItem.outputModule(renderQueueItem.numOutputModules).applyTemplate(templateName)
        except:
            self.alert_box("Error", "There's some kind of issue with this template\n\n" + str(templateName) + '\n' + str(render_queue_template))
            return

        # Set the render to the start/end times
        #self.adobe.app.beginSuppressDialogs()
        # Only set the time span if the frame range is not set to the comp frame range
        # This will automatically set the time span to the comp duration
        if self.ui.frameRangeComboBox.currentText() != self.COMP_TEXT:
            renderQueueItem.timeSpanStart = frame_range[0]
            renderQueueItem.timeSpanDuration = frame_range[1] - frame_range[0]
        #renderQueueItem.timeSpanDuration = comp.duration

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
        if self.ui.useCompNameCheckBox.isChecked():
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

            #Debugging
            logger.debug("Original Output File: %s" % originalOutputFile)
            logger.debug("New Output File: %s" % newOutputFile)

            # Rebuild the output location
            outputLocation = os.path.join(folderPath, compName, newOutputFile)
            logger.debug("Output location: %s" % outputLocation)

            # Create the output folder if it doesn't already exist
            folderPath = os.path.dirname(outputLocation)
            if not os.path.exists(folderPath):
                os.makedirs(folderPath)

        # Set the filepath and name on the newly created output module
        # Do it twice because it sometimes fails the first time - Sean
        renderQueueItem.outputModule(renderQueueItem.numOutputModules).file = self.adobe.File(outputLocation)
        renderQueueItem.outputModule(renderQueueItem.numOutputModules).file = self.adobe.File(outputLocation)

        # Log
        logger.debug("Comp: %s has been added to the render queue" % comp.name)
        #self.adobe.app.endSuppressDialogs(alert=False)

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
