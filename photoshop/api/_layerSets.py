# Import third-party modules
from _ctypes import COMError
from comtypes import ArgumentError

# Import local modules
from photoshop.api._core import Photoshop
from photoshop.api._layerSet import LayerSet
from photoshop.api.errors import PhotoshopPythonAPIError


class LayerSets(Photoshop):
    """The layer sets collection in the document."""

    app_methods = ["add", "item", "removeAll"]

    def __len__(self):
        return self.length

    def __iter__(self):
        for layer_set in self.app:
            yield LayerSet(layer_set)

    def __getitem__(self, key: str):
        """Access a given LayerSet using dictionary key lookup."""
        try:
            return LayerSet(self.app[key])
        except ArgumentError:
            raise PhotoshopPythonAPIError(f'Could not find a LayerSet named "{key}"')

    @property
    def _layerSets(self):
        return list(self.app)

    @property
    def length(self) -> int:
        """Number of elements in the collection."""
        return len(self._layerSets)

    def add(self):
        return LayerSet(self.app.add())

    def item(self, index: int) -> LayerSet:
        """Retrieves an item in this LayerSets collection by its index in the array.

        Note:
            Indexes start at 1 in Photoshop, we correct for this by adding 1 to the index
                provided so there's no confusion with Python conventions.

        Args:
            index: The index of the item to retrieve within the collection.

        Raises:
            IndexError: If the index provided is out of range.
        """
        try:
            return LayerSet(self.app.item(index + 1))
        except (IndexError, COMError) as e:
            raise IndexError(f"Index [{index}] is outside the range of this layer collection.") from e

    def removeAll(self):
        self.app.removeAll()

    def getByIndex(self, index: int):
        """Access LayerSet using list index lookup."""
        return self.item(index)

    def getByName(self, name: str) -> LayerSet:
        """Get the first element in the collection with the provided name."""
        return self[name]
