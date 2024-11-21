"""This class provides all photoshop API core functions."""

# Import future modules
from __future__ import annotations

# Import built-in modules
from functools import cache
from functools import cached_property
import logging
import os
from pathlib import Path
import platform
from typing import Any
from typing import TYPE_CHECKING
import winreg  # noqa

# Import third-party modules
from comtypes.client import CreateObject
from comtypes.client.dynamic import _Dispatch as FullyDynamicDispatch
from comtypes.client.lazybind import Dispatch

# Import local modules
from photoshop.api.constants import PHOTOSHOP_VERSION_MAPPINGS
from photoshop.api.constants import PHOTOSHOP_YEAR_MAPPINGS
from photoshop.api.errors import PhotoshopPythonAPIError


if TYPE_CHECKING:
    from photoshop.api.application import Application  # noqa:F401 isort:skip


class Photoshop:
    """Core API for all photoshop objects."""

    # Default object is Application
    object_name: str = "Application"

    # Members to flag as a method in the COM object
    app_methods: list[str] = []

    # Private attributes, do not modify
    _app_version = None
    _root = "Photoshop"
    _reg_path = "SOFTWARE\\Adobe\\Photoshop"

    def __init__(self, parent: Any = None, ps_version: str | None = None):
        """Initialize the core Photoshop object.

        Args:
            parent: Optional, parent instance to use as app object.
            ps_version: Optional, Photoshop version to look for explicitly in registry.
        """
        # Establish a version string if one was provided
        _version_given = os.environ.get("PS_VERSION", ps_version)
        if _version_given is not None:
            _version_formatted = str(_version_given).strip()
            if _version_formatted in PHOTOSHOP_VERSION_MAPPINGS:
                _version_given = PHOTOSHOP_VERSION_MAPPINGS[_version_formatted]

        # Establish initial state
        self.adobe, self.app = None, None
        self._has_parent = bool(parent is not None)

        # Establish the application object
        self.app = self._get_dispatch(version=_version_given, has_parent=self._has_parent)
        if self.app is None:
            raise PhotoshopPythonAPIError("Please check if you have Photoshop installed correctly.")

        # Substitute for parent object if provided
        if self._has_parent:
            self.adobe = self.app
            self.app = parent

        # Flag methods
        self._flag_methods()

    def __call__(self, *_args, **_kwargs):
        """Return COM object when called."""
        return self.app

    def __repr__(self):
        """Return a string representation of this Photoshop object."""
        _name = self.__class__.__name__
        if _name == "Application":
            return f"{_name}(version={self._app_year})"
        _parent = repr(self.app) if self._has_parent else None
        if _name == "Photoshop":
            return f"{_name}(parent={_parent}, ps_version={self._app_year})"
        return f"{_name}(parent={_parent})"

    def __str__(self):
        return f"{self.__class__.__name__} <{self.program_name}>"

    def __getattribute__(self, item):
        try:
            return super().__getattribute__(item)
        except AttributeError:
            return getattr(self.app, item)

    def eval_javascript(self, javascript: str, Arguments: Any = None, ExecutionMode: Any = None) -> str:
        """Instruct the application to execute javascript code."""
        executor = self.adobe if self._has_parent else self.app
        return executor.doJavaScript(javascript, Arguments, ExecutionMode)

    ###
    # * Private Properties
    ###

    @cached_property
    def _logger(self) -> logging.Logger:
        """Logger: Internal Logger object for debug output."""
        return self._get_logger()

    @cached_property
    def _app_year(self) -> str | None:
        """str: Year matching the Photoshop application version."""
        return PHOTOSHOP_YEAR_MAPPINGS.get(self._app_version)

    ###
    # * Public Properties
    ###

    @property
    def typename(self) -> str:
        """str: Current typename."""
        return self.__class__.__name__

    @property
    def program_name(self) -> str:
        """str: Formatted program name found in the Windows Classes registry, e.g. Photoshop.Application.140.

        Examples:
            - Photoshop.ActionDescriptor
            - Photoshop.ActionDescriptor.180
            - Photoshop.ActionDescriptor.190
        """
        return f"{self._root}.{self.object_name}.{self._app_version}"

    @property
    def app_id(self) -> str:
        """str: Current Photoshop application ID (maps to version), e.g. 170, 180, 190, etc."""
        return Photoshop._app_version

    @property
    def program_id(self) -> str:
        """str: Application ID as expressed in registry data, e.g. 170.0, 180.0, 190.0, etc."""
        return f"{Photoshop._app_version}.0"

    ###
    # * Private Methods
    ###

    def _get_progid(self, version: str | None = None) -> str:
        """Returns a comtypes progid for this object using a provided version string. If the version string is
            None, exclude it completely.

        Args:
            version: Version to add to progid if provided.

        Returns:
            str: The progid for comtypes to look for in registry.
        """
        _name = f"{self._root}.{self.object_name}"
        return _name if not version else f"{_name}.{version}"

    def _flag_methods(self):
        """Manually flags members of a COM object as methods.

        Notes:
            * Photoshop does not implement 'IDispatch::GetTypeInfo', so when
                getting a field from the COM object, comtypes will first try
                to fetch it as a property, then treat it as a method if it fails.
            * In this case, Photoshop does not return the proper error code, since it
                blindly treats the property getter as a method call. Fortunately, comtypes
                provides a way to explicitly flag methods.
        """
        if isinstance(self.app, FullyDynamicDispatch):
            for n in self.app_methods:
                try:
                    self.app._FlagAsMethod(n)
                except OSError:
                    self._logger.debug(f"Not a method: {n} | {str(self)}")

    def _get_photoshop_versions(self) -> list[str]:
        """Retrieve a list of Photoshop version ID's from registry."""
        try:
            key = self._open_key(self._reg_path)
            key_count = winreg.QueryInfoKey(key)[0]
            return sorted([winreg.EnumKey(key, i).split(".")[0] for i in range(key_count)], reverse=True)
        except (OSError, IndexError):
            self._logger.debug("Unable to find Photoshop version number in HKEY_LOCAL_MACHINE registry!")
        return []

    def _get_dispatch(
        self,
        version: str | None = None,
        has_parent: bool = False,
    ) -> FullyDynamicDispatch | Dispatch | None:
        """Try each version string until a valid Photoshop COM Dispatch object is returned.

        Args:
            version: Try this version first if provided by the user.
            has_parent: If this instance has a parent, use the middleman method to
                allow for cached dispatch objects and save on execution time.

        Returns:
            Photoshop application Dispatch object.

        Raises:
            OSError: If a Dispatch object wasn't resolved.
        """
        _func = self._get_cached_dispatch_object if has_parent else self._create_dispatch_object

        # Try existing version
        if Photoshop._app_version is not None:
            _obj = _func(progid=self._get_progid(Photoshop._app_version), logger=self._logger)
            if _obj is not None:
                return _obj

        # Try manual version
        if version is not None:
            _obj = _func(progid=self._get_progid(version), logger=self._logger)
            if _obj is not None:
                if Photoshop._app_version is None:
                    Photoshop._app_version = version
                return _obj

        # Try each version found in registry
        for _v in self._get_photoshop_versions():
            _obj = _func(progid=self._get_progid(_v), logger=self._logger)
            if _obj is not None:
                if Photoshop._app_version is None:
                    Photoshop._app_version = _v
                return _obj

        # Try with no version
        return _func(progid=self._get_progid(), logger=self._logger)

    ###
    # * Private Static Methods
    ###

    @staticmethod
    def _open_key(key: str) -> winreg.HKEYType:
        """Open the register key.

        Args:
            key: Photoshop application key path.

        Returns:
            The handle to the specified key.

        Raises:
            OSError: if registry key cannot be read.
        """
        machine_type = platform.machine()
        mappings = {"AMD64": winreg.KEY_WOW64_64KEY}
        access = winreg.KEY_READ | mappings.get(machine_type, winreg.KEY_WOW64_32KEY)
        try:
            return winreg.OpenKey(key=winreg.HKEY_LOCAL_MACHINE, sub_key=key, access=access)
        except FileNotFoundError as err:
            raise OSError(
                "Failed to read the registration: <{path}>\n"
                "Please check if you have Photoshop installed correctly.".format(path=f"HKEY_LOCAL_MACHINE\\{key}"),
            ) from err

    @staticmethod
    def _create_dispatch_object(
        progid: str | None = None,
        logger: logging.Logger | None = None,
    ) -> FullyDynamicDispatch | Dispatch | None:
        """Creates a COM object after injecting the provided version into a progid string.

        Args:
            progid: Program ID of the object being requested.
            logger: Optional Logger object to output any debugging information.

        Returns:
            A dynamic COM dispatch object if successful, otherwise None.
        """
        try:
            return CreateObject(progid, dynamic=True)
        except OSError:
            if logger:
                logger.debug(f"Unable to create COM object: <{progid}>")
        return

    @staticmethod
    @cache
    def _get_cached_dispatch_object(
        progid: str | None = None,
        logger: logging.Logger | None = None,
    ) -> FullyDynamicDispatch | Dispatch | None:
        """Middleware for `Photoshop._create_dispatch_object` that caches returned objects
        in cases where pulling a new instance is unnecessary.
        """
        return Photoshop._create_dispatch_object(progid=progid, logger=logger)

    @staticmethod
    @cache
    def _get_logger() -> logging.Logger:
        """Creates a Logger object with a log level optionally set by environment variable
            `PS_LOG_LEVEL`, caches it, and returns it.

        Returns:
            Logger: A `Logger` object.
        """
        logging.basicConfig(level=logging.CRITICAL)
        _log_level = {
            "DEBUG": logging.DEBUG,
            "INFO": logging.INFO,
            "WARN": logging.WARNING,
            "WARNING": logging.WARNING,
            "ERROR": logging.ERROR,
            "CRITICAL": logging.CRITICAL,
        }.get(os.environ.get("PS_LOG_LEVEL", "CRITICAL"), logging.CRITICAL)
        _logger = logging.getLogger("photoshop")
        _logger.setLevel(_log_level)
        return _logger

    ###
    # * Public Static Methods
    ###

    @staticmethod
    @cache
    def get_root_application() -> Application:
        """Application: Returns the main Application object. If one isn't cached, create one."""
        # Import local modules
        from photoshop.api.application import Application

        if Application._app_instance is None:
            return Application()
        return Application._app_instance

    @staticmethod
    @cache
    def get_application_path() -> Path:
        """Path: The absolute path of Photoshop installed location."""
        return Photoshop.get_root_application().path

    @staticmethod
    @cache
    def get_plugin_path() -> Path:
        """Path: The absolute path to the Photoshop plugins directory."""
        return Photoshop.get_root_application().path / "Plug-ins"

    @staticmethod
    @cache
    def get_presets_path() -> Path:
        """Path: The absolute path to the Photoshop presets directory."""
        return Photoshop.get_root_application().path / "Presets"

    @staticmethod
    @cache
    def get_script_path() -> Path:
        """Path: The absolute path to the Photoshop scripts directory."""
        return Photoshop.get_presets_path() / "Scripts"
