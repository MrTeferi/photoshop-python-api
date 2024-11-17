"""The Adobe Photoshop CC application object.

Which is the root of the object model and provides access to all other
objects. This object provides application-wide information,
such as application defaults and available fonts. It provides many important
methods, such as those for opening files and loading documents.

app = Application()

app.documents.add(800, 600, 72, "docRef")

"""

# Import built-in modules
import os
from pathlib import Path
import time
from typing import List
from typing import Optional
from typing import Union

# Import third-party modules
from _ctypes import COMError

# Import local modules
from photoshop.api._core import Photoshop
from photoshop.api._document import Document
from photoshop.api._documents import Documents
from photoshop.api._measurement_log import MeasurementLog
from photoshop.api._notifiers import Notifiers
from photoshop.api._preferences import Preferences
from photoshop.api._text_fonts import TextFonts
from photoshop.api.enumerations import DialogModes
from photoshop.api.enumerations import PurgeTarget
from photoshop.api.errors import PhotoshopPythonAPIError
from photoshop.api.solid_color import SolidColor


class Application(Photoshop):
    """The Adobe Photoshop application object, which contains all other Adobe Photoshop objects.

    This is the root of the object model, and provides access to all other objects.
    To access the properties and methods, you can use the pre-defined global variable app.

    Args:
        version: Manually provide a Photoshop version to use when creating dispatch objects
            for Photoshop, e.g. 2023, 2024, 2025, etc.
        singleton: Pass False if you wish to create a new object instead of caching
            and reusing one Application object as a singleton. Defaults to True.

    """
    _app_initialized = False
    _app_instance = None

    app_methods = [
        "batch",
        "charIDToTypeID",
        "doAction",
        "doJavaScript",
        "eraseCustomOptions",
        "executeAction",
        "executeActionGet",
        "featureEnabled",
        "getCustomOptions",
        "isQuicktimeAvailable",
        "load",
        "open",
        "openDialog",
        "purge",
        "putCustomOptions",
        "refresh",
        "stringIDToTypeID",
        "toolSupportsBrushes",
        "toolSupportsBrushPresets",
        "typeIDToCharID",
        "typeIDToStringID",
    ]

    def __new__(cls, version: Optional[str] = None, singleton: bool = True):
        """Ensures that only one instance of `Application` exists at a time, unless
            False is passed to `singleton` in the constructor."""
        if singleton is False or Application._app_instance is None:
            return super().__new__(cls)
        return Application._app_instance

    def __init__(self, version: Optional[str] = None, singleton: bool = True):
        """Skips initialization on subsequent calls unless `False` passed to singleton."""
        if not singleton:
            super().__init__(ps_version=version)
        elif not self._app_initialized:
            super().__init__(ps_version=version)
            Application._app_initialized = True
            Application._app_instance = self

    @property
    def activeDocument(self) -> Document:
        """The front-most document.

        Setting this property is equivalent to clicking an open document in Photoshop
        to bring it to the front of the screen.

        """
        try:
            return Document(self.app.activeDocument)
        except COMError as e:
            if 'Invalid index' in e.text:
                raise PhotoshopPythonAPIError('There are no documents open in Photoshop.') from e
            raise PhotoshopPythonAPIError('Encountered an error while trying to access the active document.') from e

    @activeDocument.setter
    def activeDocument(self, document: Document):
        self.app.activeDocument = document

    @property
    def backgroundColor(self) -> SolidColor:
        """The default background color and color style for documents."""
        return SolidColor(self.app.backgroundColor)

    @backgroundColor.setter
    def backgroundColor(self, color: SolidColor):
        """Sets the default background color and color style for documents.

        Args:
            color: The SolidColor instance.

        """
        self.app.backgroundColor = color

    @property
    def build(self) -> str:
        """str: The information about the application."""
        return self.app.build

    @property
    def colorSettings(self) -> str:
        """str: The name of the currently selected color settings profile
        (selected with Edit > Color Settings).

        """
        return self.app.colorSettings

    @colorSettings.setter
    def colorSettings(self, settings: str):
        """Sets the currently selected color settings profile.

        Args:
            settings: The name of a color settings profile to select.

        """
        try:
            self.doJavaScript(f'app.colorSettings="{settings}"')
        except COMError as e:
            raise PhotoshopPythonAPIError(f"Invalid color profile provided: '{settings}'") from e

    @property
    def currentTool(self) -> str:
        """str: The name of the current tool selected."""
        return self.app.currentTool

    @currentTool.setter
    def currentTool(self, tool_name: str):
        """Sets the currently selected tool.

        Args:
            tool_name: The name of a tool to select..

        """
        self.app.currentTool = tool_name

    @property
    def displayDialogs(self) -> DialogModes:
        """The dialog mode for the document, which indicates whether
        Photoshop displays dialogs when the script runs."""
        return DialogModes(self.app.displayDialogs)

    @displayDialogs.setter
    def displayDialogs(self, dialog_mode: DialogModes):
        """The dialog mode for the document, which indicates whether
        Photoshop displays dialogs when the script runs.
        """
        self.app.displayDialogs = dialog_mode

    @property
    def documents(self) -> Documents:
        """._documents.Documents: The Documents instance."""
        return Documents(self.app.documents)

    @property
    def fonts(self) -> TextFonts:
        return TextFonts(self.app.fonts)

    @property
    def foregroundColor(self) -> SolidColor:
        """Get default foreground color.

        Used to paint, fill, and stroke selections.

        Returns:
            The SolidColor instance.

        """
        return SolidColor(parent=self.app.foregroundColor)

    @foregroundColor.setter
    def foregroundColor(self, color: SolidColor):
        """Set the `foregroundColor`.

        Args:
            color: The SolidColor instance.

        """
        self.app.foregroundColor = color

    @property
    def freeMemory(self) -> float:
        """The amount of unused memory available to ."""
        return self.app.freeMemory

    @property
    def locale(self) -> str:
        """The language locale of the application."""
        return self.app.locale

    @property
    def macintoshFileTypes(self) -> List[str]:
        """A list of the image file types Photoshop can open."""
        return self.app.macintoshFileTypes

    @property
    def measurementLog(self) -> MeasurementLog:
        """The log of measurements taken."""
        return MeasurementLog(self.app.measurementLog)

    @property
    def name(self) -> str:
        return self.app.name

    @property
    def notifiers(self) -> Notifiers:
        """The notifiers currently configured (in the Scripts Events Manager
        menu in the application)."""
        return Notifiers(self.app.notifiers)

    @property
    def notifiersEnabled(self) -> bool:
        """bool: If true, notifiers are enabled."""
        return self.app.notifiersEnabled

    @notifiersEnabled.setter
    def notifiersEnabled(self, value: bool):
        self.app.notifiersEnabled = value

    @property
    def path(self) -> Path:
        """str: The full path to the location of the Photoshop application."""
        return Path(self.app.path)

    @property
    def playbackDisplayDialogs(self):
        return self.doJavaScript("app.playbackDisplayDialogs")

    @property
    def playbackParameters(self):
        """Stores and retrieves parameters used as part of a recorded action."""
        return self.app.playbackParameters

    @playbackParameters.setter
    def playbackParameters(self, value):
        self.app.playbackParameters = value

    @property
    def preferences(self) -> Preferences:
        return Preferences(self.app.preferences)

    @property
    def preferencesFolder(self) -> Path:
        return Path(self.app.preferencesFolder)

    @property
    def recentFiles(self):
        return self.app.recentFiles

    @property
    def scriptingBuildDate(self):
        return self.app.scriptingBuildDate

    @property
    def scriptingVersion(self):
        return self.app.scriptingVersion

    @property
    def systemInformation(self):
        return self.app.systemInformation

    @property
    def version(self):
        return self.app.version

    @property
    def windowsFileTypes(self):
        return self.app.windowsFileTypes

    # Methods.
    def batch(self, files, actionName, actionSet, options):
        """Runs the batch automation routine.

        Similar to the **File** > **Automate** > **Batch** command.

        """
        self.app.batch(files, actionName, actionSet, options)

    def beep(self):
        """Causes a "beep" sound."""
        return self.eval_javascript("app.beep()")

    def bringToFront(self):
        return self.eval_javascript("app.bringToFront()")

    def changeProgressText(self, text):
        """Changes the text that appears in the progress window."""
        self.eval_javascript(f"app.changeProgressText('{text}')")

    def charIDToTypeID(self, char_id):
        return self.app.charIDToTypeID(char_id)

    @staticmethod
    def compareWithNumbers(first, second):
        return first > second

    def doAction(self, action, action_from="Default Actions"):
        """Plays the specified action from the Actions palette."""
        self.app.doAction(action, action_from)
        return True

    def doForcedProgress(self, title, javascript):
        script = "app.doForcedProgress('{}', '{}')".format(
            title,
            javascript,
        )
        self.eval_javascript(script)
        # Ensure the script execute success.
        time.sleep(1)

    def doProgress(self, title, javascript):
        """Performs a task with a progress bar. Other progress APIs must be
        called periodically to update the progress bar and allow cancelling.

        Args:
            title (str): String to show in the progress window.
            javascript (str): JavaScriptString to execute.

        """
        script = "app.doProgress('{}', '{}')".format(
            title,
            javascript,
        )
        self.eval_javascript(script)
        # Ensure the script executes successfully
        time.sleep(1)

    def doProgressSegmentTask(self, segmentLength, done, total, javascript):
        script = "app.doProgressSegmentTask({}, {}, {}, '{}');".format(
            segmentLength,
            done,
            total,
            javascript,
        )
        self.eval_javascript(script)
        # Ensure the script execute success.
        time.sleep(1)

    def doProgressSubTask(self, index, limit, javascript):
        script = "app.doProgressSubTask({}, {}, '{}');".format(
            index,
            limit,
            javascript,
        )
        self.eval_javascript(script)
        # Ensure the script execute success.
        time.sleep(1)

    def doProgressTask(self, index, javascript):
        """Sections off a portion of the unused progress bar for execution of
        a subtask. Returns false on cancel.

        """
        script = f"app.doProgressTask({index}, '{javascript}')"
        self.eval_javascript(script)
        # Ensure the script execute success.
        time.sleep(1)

    def eraseCustomOptions(self, key):
        """Removes the specified user objects from the Photoshop registry."""
        self.app.eraseCustomOptions(key)

    def executeAction(self, event_id, descriptor, display_dialogs=2):
        return self.app.executeAction(event_id, descriptor, display_dialogs)

    def executeActionGet(self, reference):
        return self.app.executeActionGet(reference)

    def featureEnabled(self, name):
        """Determines whether the feature

        specified by name is enabled.
        The following features are supported

        as values for name:
        "photoshop/extended"
        "photoshop/standard"
        "photoshop/trial

        """
        return self.app.featureEnabled(name)

    def getCustomOptions(self, key):
        """Retrieves user objects in the Photoshop registry for the ID with
        value key."""
        return self.app.getCustomOptions(key)

    def open(
        self,
        document_file_path,
        document_type: str = None,
        as_smart_object: bool = False,
    ) -> Document:
        document = self.app.open(document_file_path, document_type, as_smart_object)
        if not as_smart_object:
            return Document(document)
        return document

    def load(self, document_file_path: Union[str, os.PathLike]) -> Document:
        """Loads a supported Photoshop document."""
        self.app.load(str(document_file_path))
        return self.activeDocument

    def doJavaScript(self, javascript, Arguments=None, ExecutionMode=None):
        return self.app.doJavaScript(javascript, Arguments, ExecutionMode)

    def isQuicktimeAvailable(self) -> bool:
        return self.app.isQuicktimeAvailable

    def openDialog(self):
        return self.app.openDialog()

    def purge(self, target: PurgeTarget):
        """Purges one or more caches.

        Args:
            target:
                1: Clears the undo cache.
                2: Clears history states from the History palette.
                3: Clears the clipboard data.
                4: Clears all caches

        """
        self.app.purge(target)

    def putCustomOptions(self, key, custom_object, persistent):
        self.app.putCustomOptions(key, custom_object, persistent)

    def refresh(self):
        """Pauses the script while the application refreshes.

        Ues to slow down execution and show the results to the user as the
        script runs.
        Use carefully; your script runs much more slowly when using this
        method.

        """
        self.app.refresh()

    def refreshFonts(self):
        """Force the font list to get refreshed."""
        return self.eval_javascript("app.refreshFonts()")

    def runMenuItem(self, menu_id):
        """Run a menu item given the menu ID."""
        return self.eval_javascript(f"app.runMenuItem({menu_id})")

    def showColorPicker(self):
        """Returns false if dialog is cancelled, true otherwise."""
        return self.eval_javascript("app.showColorPicker()")

    def stringIDToTypeID(self, string_id):
        return self.app.stringIDToTypeID(string_id)

    def togglePalettes(self):
        """Toggle palette visibility."""
        return self.doJavaScript("app.togglePalettes()")

    def toolSupportsBrushes(self, tool):
        return self.app.toolSupportsBrushes(tool)

    def toolSupportsBrushPresets(self, tool):
        return self.app.toolSupportsPresets(tool)

    @staticmethod
    def system(command):
        os.system(command)

    def typeIDToStringID(self, type_id: int) -> str:
        return self.app.typeIDToStringID(type_id)

    def typeIDToCharID(self, type_id: int) -> str:
        return self.app.typeIDToCharID(type_id)

    def updateProgress(self, done, total):
        self.eval_javascript(f"app.updateProgress({done}, {total})")
