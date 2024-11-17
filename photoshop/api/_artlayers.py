# Import built-in modules
from contextlib import suppress
from typing import Any
from typing import TYPE_CHECKING
from typing import Union

# Import third-party modules
from _ctypes import COMError
from comtypes import ArgumentError
from comtypes.client.lazybind import Dispatch

# Import local modules
from photoshop.api._artlayer import ArtLayer
from photoshop.api._core import Photoshop


if TYPE_CHECKING:
    from photoshop.api._document import Document  # noqa:F401 isort:skip
    from photoshop.api._layerSet import LayerSet  # noqa:F401 isort:skip


# pylint: disable=too-many-public-methods
class ArtLayers(Photoshop):
    """A collection of `ArtLayer` objects in a `Document` or `LayerSet`."""

    app_methods = ["add"]

    def __len__(self):
        return self.length

    def __iter__(self):
        for layer in self.app:
            yield ArtLayer(layer)

    def __getitem__(self, key: str):
        """Access a given ArtLayer using dictionary key lookup, where name of the ArtLayer is key."""
        try:
            return ArtLayer(self.app[key])
        except (ArgumentError, COMError, KeyError):
            raise KeyError(f'Could not find an ArtLayer named "{key}"')

    @property
    def _layers(self) -> list[Dispatch]:
        """Private property that returns a list containing a Dispatch object for each
        `ArtLayer` in this collection."""
        return list(self.app)

    @property
    def length(self) -> int:
        """The number of `ArtLayer` objects in this collection."""
        return len(self._layers)

    @property
    def parent(self) -> Union["Document", "LayerSet"]:
        """The parent `Document` or `LayerSet` containing this `ArtLayers` collection."""
        _parent = self.app.parent
        try:
            # Parent is a Document
            _ = _parent.path
            from photoshop.api._document import Document  # noqa:F811 isort:skip

            return Document(_parent)
        except (COMError, NameError, OSError):
            # Parent is a LayerSet
            from photoshop.api._layerSet import LayerSet  # noqa:F811 isort:skip

            return LayerSet(_parent)

    def get(self, key: str, default: Any = None) -> ArtLayer:
        """Access a given ArtLayer using dictionary key lookup, where name of the ArtLayer is key. You can
            provide an optional default value to return if the key does not exist, otherwise will return None.

        Args:
            key: The name of the ArtLayer to lookup.
            default: The default value to return if the key does not exist, defaults to None.

        Returns:
            ArtLayer if key is found, otherwise default value.
        """
        try:
            return self[key]
        except KeyError:
            return default

    def add(self):
        """Adds an element."""
        return ArtLayer(self.app.add())

    def getByIndex(self, index: int):
        """Access ArtLayer using list index lookup."""
        return ArtLayer(self._layers[index])

    def getByName(self, name: str) -> ArtLayer:
        """Get the first element in the collection with the provided name.

        Args:
            name: The name of the ArtLayer to lookup.

        Raises:
            PhotoshopPythonAPIError: If an ArtLayer with the provided name couldn't be found.
        """
        return self[name]

    def removeAll(self):
        """Deletes all elements."""
        for n in self._layers:
            with suppress(Exception):
                n.remove()
