# Import built-in modules
from __future__ import annotations
from typing import TYPE_CHECKING

# Import third-party modules
from comtypes import ArgumentError

# Import local modules
from photoshop.api._core import Photoshop
from photoshop.api._layerComp import LayerComp
from photoshop.api.errors import PhotoshopPythonAPIError
if TYPE_CHECKING:
    from photoshop.api._document import Document


class LayerComps(Photoshop):
    """The layer comps collection in this document."""

    app_methods = ["add", "removeAll"]

    def __len__(self):
        return self.length

    def __iter__(self):
        for layer_comp in self.app:
            yield LayerComp(layer_comp)

    def __getitem__(self, key: str):
        """Access a given LayerComp using dictionary key lookup, where name of the LayerComp is key."""
        try:
            return LayerComp(self.app[key])
        except ArgumentError:
            raise PhotoshopPythonAPIError(f'Could not find a LayerComp named "{key}"')

    @property
    def length(self):
        """The number of `LayerComp` objects in this collection."""
        return len(self._layers)

    @property
    def _layers(self):
        """Private property that returns a list containing a Dispatch object for each
            `LayerComp` in this collection."""
        return list(self.app)

    @property
    def parent(self) -> 'Document':
        """The parent `Document` containing this `LayerComps` collection."""
        from photoshop.api._document import Document
        return Document(self.app.parent)

    def add(
        self,
        name,
        comment="No Comment.",
        appearance=True,
        position=True,
        visibility=True,
        childLayerCompStat=False,
    ):
        return LayerComp(self.app.add(name, comment, appearance, position, visibility, childLayerCompStat))

    def getByName(self, name: str):
        """Get the first element in the collection with the provided name.

        Args:
            name: The name of the LayerComp to lookup.

        Raises:
            PhotoshopPythonAPIError: If a LayerComp with the provided name couldn't be found.
        """
        return self[name]

    def removeAll(self):
        self.app.removeAll()
