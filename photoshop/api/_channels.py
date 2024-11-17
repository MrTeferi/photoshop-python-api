# Import local modules
from photoshop.api._channel import Channel
from photoshop.api._core import Photoshop
from photoshop.api.errors import PhotoshopPythonAPIError


# pylint: disable=too-many-public-methods
class Channels(Photoshop):
    app_methods = ["add", "removeAll"]

    def __len__(self):
        return self.length

    def __iter__(self):
        for channel in self.app:
            yield Channel(channel)

    def __getitem__(self, index: int):
        """Access a given Channel using its index in the array."""
        return Channel(self.app[index])

    @property
    def _channels(self):
        return list(self.app)

    @property
    def length(self):
        return len(self._channels)

    def add(self):
        self.app.add()

    def removeAll(self):
        self.app.removeAll()

    def getByName(self, name) -> Channel:
        for channel in self._channels:
            if channel.name == name:
                return Channel(channel)
        raise PhotoshopPythonAPIError(f'Could not find a channel named "{name}"')
