# Import future modules
from __future__ import annotations

# Import built-in modules
from typing import TYPE_CHECKING

# Import local modules
from photoshop.api._core import Photoshop


if TYPE_CHECKING:
    from photoshop.api._document import Document  # noqa:F401 isort:skip


class LayerComp(Photoshop):
    """A snapshot of a state of the layers in a document (can be used to view
    different page layouts or compositions)."""

    app_methods = ["apply", "recapture", "remove", "resetfromComp"]

    def __len__(self):
        return self.length

    @property
    def appearance(self):
        return self.app.appearance

    @appearance.setter
    def appearance(self, value):
        self.app.appearance = value

    @property
    def childLayerCompState(self):
        return self.app.childLayerCompState

    @childLayerCompState.setter
    def childLayerCompState(self, value):
        self.app.childLayerCompState = value

    @property
    def comment(self):
        return self.app.comment

    @comment.setter
    def comment(self, text):
        self.app.comment = text

    @property
    def name(self):
        return self.app.name

    @name.setter
    def name(self, text):
        self.app.name = text

    @property
    def parent(self) -> Document:
        """The parent `Document` containing this `LayerComp`."""
        # Import local modules
        from photoshop.api._document import Document

        return Document(self.app.parent)

    @property
    def position(self):
        return self.app.position

    @position.setter
    def position(self, value):
        self.app.position = value

    @property
    def selected(self):
        """True if the layer comp is currently selected."""
        return self.app.selected

    @selected.setter
    def selected(self, value):
        self.app.selected = value

    @property
    def visibility(self):
        """True to use layer visibility settings."""
        return self.app.visibility

    @visibility.setter
    def visibility(self, value):
        self.app.visibility = value

    def apply(self):
        """Applies the layer comp to the document."""
        self.app.apply()

    def recapture(self):
        """Recaptures the current layer state(s) for this layer comp."""
        self.app.recapture()

    def remove(self):
        """Deletes the layerComp object."""
        self.app.remove()

    def resetfromComp(self):
        """Resets the layer comp state to the document state."""
        self.app.resetfromComp()
