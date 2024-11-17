# Import built-in modules
from contextlib import suppress
from typing import Union, TYPE_CHECKING, Iterator

# Import third-party modules
from _ctypes import COMError

# Import local modules
from photoshop.api._artlayer import ArtLayer
from photoshop.api._core import Photoshop
if TYPE_CHECKING:
    from photoshop.api._layerSet import LayerSet


# pylint: disable=too-many-public-methods
class Layers(Photoshop):
    """The layers collection in the document."""

    app_methods = ["add", "item"]

    def __getitem__(self, item: Union[int, str]) -> Union[ArtLayer, 'LayerSet']:
        """Retrieve an ArtLayer or LayerSet from this Layers collection by name
            (if provided as a string) or index (if provided as an integer) within
            the collection array.

        Args:
            item: The name of the layer to retrieve or index within the layers array.

        Returns:
            ArtLayer or LayerSet object.

        Raises:
            KeyError: If no ArtLayer or LayerSet is found matching a name provided.
            IndexError: If no ArtLayer or LayerSet is found matching an index provided.
        """
        # Photoshop arrays start at index of 1
        if isinstance(item, int):
            return self.item(item)
        return self.getByName(item)

    def __iter__(self) -> Iterator[Union[ArtLayer, 'LayerSet']]:
        for layer in self.app:
            try:
                _ = layer.kind
                yield ArtLayer(layer)
            except (COMError, NameError, OSError):
                from photoshop.api._layerSet import LayerSet
                yield LayerSet(layer)

    def __len__(self):
        return self.length

    @property
    def _layers(self):
        return list(self.app)

    @property
    def length(self):
        return len(self._layers)

    def getByName(self, name: str) -> Union[ArtLayer, 'LayerSet']:
        """Get the first element in the Layers collection with the provided name.

        Args:
            name: Name of the layer to look for.

        Returns:
            An `ArtLayer` or `LayerSet` object.

        Raises:
            KeyError: If no ArtLayer or LayerSet with the name provided is found.
        """
        try:
            layer = self.app[name]
            try:
                _ = layer.kind
                return ArtLayer(layer)
            except NameError:
                from photoshop.api._layerSet import LayerSet
                return LayerSet(layer)
        except (KeyError, COMError) as e:
            raise KeyError(f"Layer with name '{name}' not found in this layer collection.") from e

    def item(self, index: int) -> Union[ArtLayer, 'LayerSet']:
        """Retrieve an item from the Layers collection by index in the array.

        Args:
            index: Index in the array.

        Returns:
            An `ArtLayer` or `LayerSet` object.

        Raises:
            IndexError: If index provided is out of range.
        """
        try:
            # Arrays start at 1 in Photoshop
            layer = self.app.item(index + 1)
            try:
                _ = layer.kind
                return ArtLayer(layer)
            except NameError:
                from photoshop.api._layerSet import LayerSet
                return LayerSet(layer)
        except (IndexError, COMError) as e:
            raise IndexError(f"Index [{index}] is outside the range of this layer collection.") from e

    def removeAll(self):
        """Deletes all elements."""
        for layer in self.app:
            with suppress(Exception):
                layer.delete()
