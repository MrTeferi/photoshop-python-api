# Import built-in modules
from ctypes import ArgumentError
from typing import Union

# Import third-party modules
from _ctypes import COMError

# Import local modules
from photoshop.api._core import Photoshop
from photoshop.api._document import Document
from photoshop.api.enumerations import BitsPerChannelType
from photoshop.api.enumerations import DocumentFill
from photoshop.api.enumerations import NewDocumentMode


# pylint: disable=too-many-public-methods, too-many-arguments
class Documents(Photoshop):
    """The collection of open documents."""

    app_methods = ["add"]

    def __getitem__(self, item: Union[int, str]) -> Document:
        """Retrieve a Document from the Documents collection with the provided name or index.

        Args:
            item (Union[int, str]): The index within the Documents collection or name of the
                Document to look for.

        Returns:
            A Document object.

        Raises:
            IndexError: If item was provided as an integer and is out of range of collection array.
            KeyError: If item was provided as a string and the Document name was not found.
        """
        if isinstance(item, int):
            return self.item(item)
        return self.getByName(item)

    def __iter__(self) -> Document:
        for doc in self.app:
            yield Document(doc)

    def __len__(self) -> int:
        return self.length

    @property
    def length(self) -> int:
        return len(list(self.app))

    def add(
        self,
        width: int = 960,
        height: int = 540,
        resolution: float = 72.0,
        name: str = None,
        mode: int = NewDocumentMode.NewRGB,
        initialFill: int = DocumentFill.White,
        pixelAspectRatio: float = 1.0,
        bitsPerChannel: int = BitsPerChannelType.Document8Bits,
        colorProfileName: str = None,
    ) -> Document:
        """Creates a new document object and adds it to this collections.

        Args:
            width (int): The width of the document.
            height (int): The height of the document.
            resolution (int): The resolution of the document (in pixels per inch)
            name (str): The name of the document.
            mode (): The document mode.
            initialFill : The initial fill of the document.
            pixelAspectRatio: The initial pixel aspect ratio of the document.
                                Default is `1.0`, the range is `0.1-10.00`.
            bitsPerChannel: The number of bits per channel.
            colorProfileName: The name of color profile for document.

        Returns:
            .Document: Document instance.

        """
        return Document(
            self.app.add(
                width,
                height,
                resolution,
                name,
                mode,
                initialFill,
                pixelAspectRatio,
                bitsPerChannel,
                colorProfileName,
            ),
        )

    def item(self, index: int) -> Document:
        """Retrieve an item from the Documents collection by index in the array.

        Args:
            index: Index in the array.

        Returns:
            A Document object.

        Raises:
            IndexError: If index provided is out of range.
        """
        try:
            return Document(self.app.item(index + 1))
        except (COMError, IndexError):
            raise IndexError(f"Index [{index}] out of range in Documents collection.")

    def getByName(self, name: str) -> Document:
        """Retrieve a Document from the Documents collection by name.

        Args:
            name (str): The name of the Document to look for.

        Returns:
            A Document object.

        Raises:
            KeyError: If no document with the given name exists.
        """
        try:
            return Document(self.app[name])
        except (ArgumentError, COMError, KeyError):
            raise KeyError(f"No document found with name '{name}'.")
