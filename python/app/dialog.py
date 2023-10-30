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
        self.populate_frame_range()
        self.popular_frame_range_options()
        self.populate_presets()

    def populate_frame_range(self):
        self.ui.frameRangeLineEdit.setText("%d - %d" % (self.first_frame, self.last_frame))

    def popular_frame_range_options(self):
        self.ui.frameRangeComboBox.insertItems(0, [self.COMP_TEXT, self.CUSTOM_TEXT])
        self.ui.frameRangeLineEdit.setEnabled(False)

    def populate_presets(self):
        self.presets = {}
        for preset_item in self._app.get_setting("render_presets"):
            # use an internal method to resolve the path of the ae template files
            resolved_path = self._app._TankBundle__resolve_hook_expression(preset_item['name'], preset_item['path'])
            self.presets[preset_item['name']] = resolved_path[0]
            self.ui.renderFormatDropdown.insertItems(-1, [preset_item['name']])

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

        count = 0
        for comp in selected_comps:
            frame_range = self.get_frame_range(comp)
            if frame_range[0] is None or frame_range[1] is None:
                self.alert_box("Bad frame range")

            render_queue_template = self.get_render_queue_template()
            if render_queue_template is None:
                self.alert_box("Failed to find selected render preset")

            self.create_render_queue_item_for_comp(comp, frame_range, render_queue_template)
            count += 1

        self.close()
        self.alert_box('Successfully created ' + str(count) + ' item(s) in the render queue')

    def get_frame_range(self, comp):
        startFrame = None
        endFrame = None

        # Use comp frame range
        if self.ui.frameRangeComboBox.currentText() == self.COMP_TEXT:
            startFrame = comp.displayStartFrame
            endFrame = int(comp.frameRate * comp.duration + 0.0001)
        else:
            rawText = self.ui.frameRangeLineEdit.text()
            # Assumed pattern is {Digits}{NonDigitSeperator}{Digits} - e.g. 1001-1002
            match = re.match(r'(\d+)(\D+)(\d+)', rawText)
            if match:
                startFrame = match.group(1)
                endFrame = match.group(3)

        return [startFrame, endFrame]

    def get_render_queue_template(self):
        render_queue_template = None

        userSelection = self.ui.renderFormatDropdown.currentText()
        if userSelection in self.presets:
            render_queue_template = self.presets[userSelection]

        return render_queue_template

    def alert_box(self, text):
        QtGui.QMessageBox.critical(
            self,
            "!",
            str(text),
            buttons=QtGui.QMessageBox.Ok,
            defaultButton=QtGui.QMessageBox.Ok,
        )

    def create_render_queue_item_for_comp(self, comp, frame_range, render_queue_template):
        templateName = self.ui.renderFormatDropdown.currentText()

        # Check the template actually exists
        if not self.check_template_exists(comp, frame_range, render_queue_template, templateName):
            self.alert_box("Something went wrong applying or locating an output template")

        renderQueueItem = self.adobe.app.project.renderQueue.items.add(comp)
        try:
            renderQueueItem.outputModule(renderQueueItem.numOutputModules).applyTemplate(templateName)
        except:
            self.alert_box("There's some kind of issue with this template\n\n" + str(templateName) + '\n' + str(render_queue_template))
            return

        # Set the render to the start/end times
        self.adobe.app.beginSuppressDialogs()
        renderQueueItem.timeSpanStart = frame_range[0]
        renderQueueItem.timeSpanDuration = comp.duration
        self.adobe.app.beginSuppressDialogs(False)

        # Grab the output folder from templates
        outputLocation = self.get_shotgrid_template(render_queue_template)

        # Create the output folder if it doesn't already exist
        folderPath = os.path.dirname(outputLocation)
        if not os.path.exists(folderPath):
            os.makedirs(folderPath)

        # Set the filepath and name on the newly created output module
        # Do it twice because it sometimes fails the first time - Sean
        renderQueueItem.outputModule(renderQueueItem.numOutputModules).file = self.adobe.File(outputLocation)
        renderQueueItem.outputModule(renderQueueItem.numOutputModules).file = self.adobe.File(outputLocation)

    def get_shotgrid_template(self, render_queue_template):
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
        renderQueueItems = self.adobe.app.project.renderQueue.items
        for i in range(1, self.adobe.app.project.renderQueue.numItems+1):
            if renderQueueItems[i].comp.name == comp_name:
                return renderQueueItems[i]

        return None

    def importPresetProject(self, render_queue_template):
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
