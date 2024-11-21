# Import future modules
from __future__ import annotations

# Import built-in modules
from typing import TYPE_CHECKING

# Import third-party modules
from _ctypes import COMError
from comtypes.client.lazybind import Dispatch

# Import local modules
from photoshop.api._artlayer import ArtLayer
from photoshop.api._artlayers import ArtLayers
from photoshop.api._core import Photoshop
from photoshop.api._layers import Layers
from photoshop.api.enumerations import AnchorPosition
from photoshop.api.enumerations import BlendMode
from photoshop.api.enumerations import ElementPlacement


if TYPE_CHECKING:
    from photoshop.api._document import Document  # noqa:F401 isort:skip
    from photoshop.api._layerSets import LayerSets  # noqa:F401 isort:skip


class LayerSet(Photoshop):
    """A group of layer objects, which can include art layer objects and other (nested) layer set objects.

    A single command or set of commands manipulates all layers in a layer set object.
    """

    app_methods = ["merge", "duplicate", "add", "delete", "link", "move", "resize", "rotate", "translate", "unlink"]

    def __getitem__(self, item: int | str) -> ArtLayer | LayerSet:
        """Retrieve an ArtLayer or LayerSet from this LayerSet by name (if provided
            as a string) or index (if provided as an integer) within the layers array.

        Args:
            item: The name of the layer to retrieve or index within the layers array.

        Returns:
            ArtLayer or LayerSet object.

        Raises:
            KeyError: If no ArtLayer or LayerSet is found with a provided
                name or index.
        """
        return self.layers[item]

    def __iter__(self):
        yield from self.layers

    def __len__(self):
        return self.length

    @property
    def _layers(self) -> list[Dispatch]:
        return list(self.layers._layers)

    @property
    def allLocked(self):
        return self.app.allLocked

    @allLocked.setter
    def allLocked(self, value):
        self.app.allLocked = value

    @property
    def artLayers(self):
        return ArtLayers(self.app.artLayers)

    @property
    def blendMode(self):
        return BlendMode(self.app.blendMode)

    @property
    def bounds(self):
        """The bounding rectangle of the layer set."""
        return self.app.bounds

    @property
    def enabledChannels(self):
        return self.app.enabledChannels

    @enabledChannels.setter
    def enabledChannels(self, value):
        self.app.enabledChannels = value

    @property
    def layers(self):
        return Layers(self.app.layers)

    @property
    def layerSets(self) -> LayerSets:
        # pylint: disable=import-outside-toplevel
        from ._layerSets import LayerSets

        return LayerSets(self.app.layerSets)

    @property
    def length(self) -> int:
        return len(self._layers)

    @property
    def linkedLayers(self):
        """The layers linked to this layerSet object."""
        return self.app.linkedLayers or []

    @property
    def name(self) -> str:
        return self.app.name

    @name.setter
    def name(self, value):
        """The name of this layer set."""
        self.app.name = value

    @property
    def opacity(self):
        """The master opacity of the set."""
        return round(self.app.opacity)

    @opacity.setter
    def opacity(self, value):
        self.app.opacity = value

    @property
    def parent(self) -> Document | LayerSet:
        """The parent `Document` or `LayerSet` containing this `LayerSet`."""
        _parent = self.app.parent
        try:
            # Parent is a Document
            _ = _parent.path
            from photoshop.api._document import Document  # noqa:F811 isort:skip

            return Document(_parent)
        except (COMError, NameError, OSError):
            # Parent is a LayerSet
            return LayerSet(_parent)

    @parent.setter
    def parent(self, parent: Document | LayerSet):
        """Set the object’s parent container."""
        if parent.__class__.__name__ == "Document":
            # Move to another Document
            _new = self.duplicate(parent, ElementPlacement.PlaceAtBeginning)
            self.remove()
            self.app = _new.app
        else:
            # Move to another LayerSet
            self.move(parent, ElementPlacement.PlaceInside)

    @property
    def visible(self) -> bool:
        return self.app.visible

    @visible.setter
    def visible(self, value: bool):
        self.app.visible = value

    def duplicate(
        self,
        relativeObject: ArtLayer | Document | LayerSet | None = None,
        insertionLocation: ElementPlacement | None = None,
    ):
        """Duplicate this `LayerSet` and place the duplicate relative to another `ArtLayer`,
            `Document`, or `LayerSet`.

        Note:
            LayerSet cannot be moved inside a target LayerSet, so we create a
                temporary layer for positioning.

        Args:
            relativeObject: An ArtLayer, Document, or LayerSet object to position the
                duplicate object relative to.
            insertionLocation: The placement strategy of the duplicate object relative to
                the relativeObject. Should be provided as an ElementPlacement enum.

        Returns:
            The duplicate `LayerSet` object.
        """
        if isinstance(relativeObject, LayerSet) and ElementPlacement not in [
            ElementPlacement.PlaceBefore,
            ElementPlacement.PlaceAfter,
        ]:
            _temp = relativeObject.add()
            if insertionLocation == ElementPlacement.PlaceAtEnd:
                _temp.move(relativeObject, ElementPlacement.PlaceAtEnd)
            _duplicate = LayerSet(self.app.duplicate(_temp.app, ElementPlacement.PlaceBefore))
            _temp.remove()
            return _duplicate
        return LayerSet(self.app.duplicate(relativeObject, insertionLocation))

    def link(self, with_layer):
        self.app.link(with_layer)

    def add(self, layer_set: bool = False):
        """Adds a layer object to this `LayerSet`.

        Args:
            layer_set: If True, add and return a new `LayerSet` object. Otherwise, add
                and return a new `ArtLayer`.
        """
        if layer_set:
            return self.layerSets.add()
        return self.artLayers.add()

    def merge(self) -> ArtLayer:
        """Merges the layer set."""
        return ArtLayer(self.app.merge())

    def move(
        self,
        relativeObject: ArtLayer | Document | LayerSet | None = None,
        insertionLocation: ElementPlacement | None = None,
    ) -> None:
        """Move this `LayerSet` relative to another `ArtLayer`, `Document`, or `LayerSet`.

        Note:
            LayerSet cannot be moved inside a target LayerSet, so we create a
                temporary layer for positioning.

        Args:
            relativeObject: An ArtLayer, Document, or LayerSet object to position this
                object relative to.
            insertionLocation: The placement strategy of this object relative to the
                relativeObject. Should be provided as an ElementPlacement enum.
        """
        if isinstance(relativeObject, LayerSet) and ElementPlacement not in [
            ElementPlacement.PlaceBefore,
            ElementPlacement.PlaceAfter,
        ]:
            _temp = relativeObject.add()
            if insertionLocation == ElementPlacement.PlaceAtEnd:
                _temp.move(relativeObject, ElementPlacement.PlaceAtEnd)
            self.app.move(_temp.app, ElementPlacement.PlaceBefore)
            _temp.remove()
        else:
            self.app.move(relativeObject, insertionLocation)

    def remove(self):
        """Remove this layer set from the document."""
        self.app.delete()

    def resize(self, horizontal=None, vertical=None, anchor: AnchorPosition = None):
        self.app.resize(horizontal, vertical, anchor)

    def rotate(self, angle, anchor=None):
        self.app.rotate(angle, anchor)

    def translate(self, delta_x, delta_y):
        """Moves the position relative to its current position."""
        self.app.translate(delta_x, delta_y)

    def unlink(self):
        """Unlinks the layer set."""
        self.app.unlink()
