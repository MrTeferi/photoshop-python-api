"""The collection of Notifier objects in the document.

Access through the Application.notifiers collection property.

Examples:
    Notifiers must be enabled using the Application.notifiersEnabled property.
    ```javascript
    var notRef = app.notifiers.add("OnClickGoButton", eventFile)
    ```

"""
# Import built-in modules
from typing import Any, Iterator
from typing import Optional

# Import third-party modules
from comtypes import ArgumentError
from comtypes.client.lazybind import Dispatch

# Import local modules
from photoshop.api.errors import PhotoshopPythonAPIError
from photoshop.api._core import Photoshop
from photoshop.api._notifier import Notifier


class Notifiers(Photoshop):
    """An array collection of each `Notifier` currently configured (in the Scripts Events Manager
        menu in Photoshop)."""

    app_methods = ["add", "removeAll"]

    def __len__(self) -> int:
        return self.length

    def __iter__(self) -> Iterator[Notifier]:
        for notifier in self.app:
            yield Notifier(notifier)

    def __getitem__(self, item: int) -> Notifier:
        """Return a Notifier from the collection by index in the array."""
        try:
            return Notifier(self._notifiers[item])
        except ArgumentError:
            raise PhotoshopPythonAPIError(f'Could not find a Notifier with index {item}')

    @property
    def _notifiers(self) -> list[Dispatch]:
        """Private property that returns a list containing a Dispatch object for each
            `Notifier` in this collection."""
        return list(self.app)

    @property
    def length(self) -> int:
        """The number of `Notifier` objects in this collection."""
        return len(self._notifiers)

    def add(
        self, event: str, event_file: Optional[Any] = None,
        event_class: Optional[Any] = None
    ) -> Notifier:
        """Add a `Notifier` to the collection."""
        self.parent.notifiersEnabled = True
        return Notifier(self.app.add(event, event_file, event_class))

    def removeAll(self):
        """Remove all `Notifier` objects from the collection."""
        self.app.removeAll()
        self.parent.notifiersEnabled = False
