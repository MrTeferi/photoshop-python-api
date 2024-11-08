"""Options for saving a document in BMO format."""

# Import local modules
from photoshop.api._core import Photoshop


class BMPSaveOptions(Photoshop):
    """Options for saving a document in BMP format."""

    object_name = "BMPSaveOptions"

    @property
    def alphaChannels(self) -> bool:
        """State to save the alpha channels."""
        return self.app.alphaChannels

    @alphaChannels.setter
    def alphaChannels(self, value: bool):
        """Sets whether to save the alpha channels or not.

        Args:
            value: Whether to save the alpha channels.
        """
        self.app.alphaChannels = value
