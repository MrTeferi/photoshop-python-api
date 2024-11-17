# Import built-in modules
from contextlib import suppress
from typing import Union, TYPE_CHECKING, Optional

# Import third-party modules
from _ctypes import COMError

# Import local modules
from photoshop.api._core import Photoshop
from photoshop.api.enumerations import ElementPlacement
from photoshop.api.enumerations import LayerKind
from photoshop.api.enumerations import RasterizeType
from photoshop.api.text_item import TextItem
if TYPE_CHECKING:
    from photoshop.api._document import Document
    from photoshop.api._layerSet import LayerSet


# pylint: disable=too-many-public-methods, too-many-arguments
class ArtLayer(Photoshop):
    """An object within a document that contains the visual elements of the image
    (equivalent to a layer in the Adobe Photoshop application).
    """

    app_methods = [
        "adjustBrightnessContrast",
        "adjustColorBalance",
        "adjustCurves",
        "adjustLevels",
        "applyAddNoise",
        "applyAverage",
        "applyBlur",
        "applyBlurMore",
        "applyClouds",
        "applyCustomFilter",
        "applyDeInterlace",
        "applyDespeckle",
        "applyDifferenceClouds",
        "applyDiffuseGlow",
        "applyDisplace",
        "applyDustAndScratches",
        "applyGaussianBlur",
        "applyGlassEffect",
        "applyHighPass",
        "applyLensBlur",
        "applyLensFlare",
        "applyMaximum",
        "applyMedianNoise",
        "applyMinimum",
        "applyMotionBlur",
        "applyNTSC",
        "applyOceanRipple",
        "applyOffset",
        "applyPinch",
        "delete",
        "duplicate",
        "invert",
        "link",
        "merge",
        "move",
        "posterize",
        "rasterize",
        "unlink",
    ]

    @property
    def allLocked(self):
        return self.app.allLocked

    @allLocked.setter
    def allLocked(self, value):
        self.app.allLocked = value

    @property
    def blendMode(self):
        return self.app.blendMode

    @blendMode.setter
    def blendMode(self, mode):
        self.app.blendMode = mode

    @property
    def bounds(self):
        """The bounding rectangle of the layer."""
        return self.app.bounds

    @property
    def linkedLayers(self) -> list:
        """Get all layers linked to this layer.

        Returns:
            list: Layer objects"""
        return [ArtLayer(layer) for layer in self.app.linkedLayers]

    @property
    def name(self) -> str:
        return self.app.name

    @name.setter
    def name(self, text: str):
        self.app.name = text

    @property
    def fillOpacity(self):
        """The interior opacity of the layer. Range: 0.0 to 100.0."""
        return self.app.fillOpacity

    @fillOpacity.setter
    def fillOpacity(self, value):
        """The interior opacity of the layer. Range: 0.0 to 100.0."""
        self.app.fillOpacity = value

    @property
    def filterMaskDensity(self):
        return self.app.filterMaskDensity

    @filterMaskDensity.setter
    def filterMaskDensity(self, value):
        self.app.filterMaskDensity = value

    @property
    def filterMaskFeather(self):
        return self.app.filterMaskFeather

    @filterMaskFeather.setter
    def filterMaskFeather(self, value):
        self.app.filterMaskFeather = value

    @property
    def grouped(self) -> bool:
        """If true, the layer is grouped with the layer below."""
        return self.app.grouped

    @grouped.setter
    def grouped(self, value):
        self.app.grouped = value

    @property
    def isBackgroundLayer(self):
        """bool: If true, the layer is a background layer."""
        return self.app.isBackgroundLayer

    @isBackgroundLayer.setter
    def isBackgroundLayer(self, value):
        self.app.isBackgroundLayer = value

    @property
    def kind(self):
        """Sets the layer kind (such as ‘text layer’) for an empty layer.

        Valid only when the layer is empty and when `isBackgroundLayer` is
        false. You can use the ‘kind ‘ property to make a background layer a
         normal layer; however, to make a layer a background layer, you must
         set `isBackgroundLayer` to true.

        """
        return LayerKind(self.app.kind)

    @kind.setter
    def kind(self, layer_type):
        """set the layer kind."""
        self.app.kind = layer_type

    @property
    def layerMaskDensity(self):
        """The density of the layer mask (between 0.0 and 100.0)."""
        return self.app.layerMaskDensity

    @layerMaskDensity.setter
    def layerMaskDensity(self, value):
        self.app.layerMaskDensity = value

    @property
    def layerMaskFeather(self):
        """The feather of the layer mask (between 0.0 and 250.0)."""
        return self.app.layerMaskFeather

    @layerMaskFeather.setter
    def layerMaskFeather(self, value):
        self.app.layerMaskFeather = value

    @property
    def opacity(self):
        """The master opacity of the layer."""
        return round(self.app.opacity)

    @opacity.setter
    def opacity(self, value: Union[int, float]):
        """Set the master opacity of the layer.

        Args:
            value: Integer or float from 0-100.
        """
        self.app.opacity = value

    @property
    def parent(self) -> Union['Document', 'LayerSet']:
        """The parent `Document` or `LayerSet` containing this `ArtLayer`."""
        _parent = self.app.parent
        try:
            # Parent is a Document
            _ = _parent.path
            from photoshop.api._document import Document
            return Document(_parent)
        except (COMError, NameError, OSError):
            # Parent is a LayerSet
            from photoshop.api._layerSet import LayerSet
            return LayerSet(_parent)

    @parent.setter
    def parent(self, parent: Union['Document', 'LayerSet']):
        """Set the object’s container."""
        if parent.__class__.__name__ == 'Document':
            _new = self.duplicate(parent, ElementPlacement.PlaceAtBeginning)
            self.remove()
            self.app = _new.app
        else:
            self.move(parent, ElementPlacement.PlaceInside)

    @property
    def pixelsLocked(self) -> bool:
        """If true, the pixels in the layer’s image cannot be edited."""
        return self.app.pixelsLocked

    @pixelsLocked.setter
    def pixelsLocked(self, value: bool):
        self.app.pixelsLocked = value

    @property
    def positionLocked(self) -> bool:
        """bool: If true, the pixels in the layer’s image cannot be moved
        within the layer."""
        return self.app.positionLocked

    @positionLocked.setter
    def positionLocked(self, value: bool):
        self.app.positionLocked = value

    @property
    def textItem(self) -> TextItem:
        """The text that is associated with the layer. Valid only when ‘kind’
            is text layer. Note that some documentation sources insist this is a writable
            property, however with COM automation this doesn't appear to be possible."""
        return TextItem(self.app.textItem)

    @property
    def transparentPixelsLocked(self):
        return self.app.transparentPixelsLocked

    @transparentPixelsLocked.setter
    def transparentPixelsLocked(self, value):
        self.app.transparentPixelsLocked = value

    @property
    def vectorMaskDensity(self):
        return self.app.vectorMaskDensity

    @vectorMaskDensity.setter
    def vectorMaskDensity(self, value):
        self.app.vectorMaskDensity = value

    @property
    def vectorMaskFeather(self):
        return self.app.vectorMaskFeather

    @vectorMaskFeather.setter
    def vectorMaskFeather(self, value):
        self.app.vectorMaskFeather = value

    @property
    def visible(self):
        return self.app.visible

    @visible.setter
    def visible(self, value):
        self.app.visible = value

    @property
    def length(self):
        return len(list(self.app))

    def adjustBrightnessContrast(self, brightness, contrast) -> None:
        """Adjusts the brightness and contrast.

        Args:
            brightness (int): The brightness amount. Range: -100 to 100.
            contrast (int): The contrast amount. Range: -100 to 100.

        """
        return self.app.adjustBrightnessContrast(brightness, contrast)

    def adjustColorBalance(
        self,
        shadows,
        midtones,
        highlights,
        preserveLuminosity,
    ) -> None:
        """Adjusts the color balance of the layer’s component channels.

        Args:
            shadows: The adjustments for the shadows. The array must include
                     three values (in the range -100 to 100), which represent
                     cyan or red, magenta or green, and yellow or blue, when
                     the document mode is CMYK or RGB.
            midtones: The adjustments for the midtones. The array must include
                      three values (in the range -100 to 100), which represent
                      cyan or red, magenta or green, and yellow or blue, when
                      the document mode is CMYK or RGB.
            highlights: The adjustments for the highlights. The array must
                        include three values (in the range -100 to 100), which
                        represent cyan or red, magenta or green, and yellow or
                        blue, when the document mode is CMYK or RGB.
            preserveLuminosity: If true, luminosity is preserved.

        """
        return self.app.adjustColorBalance(
            shadows,
            midtones,
            highlights,
            preserveLuminosity,
        )

    def adjustCurves(self, curveShape) -> None:
        """Adjusts the tonal range of the selected channel using up to fourteen
        points.



        Args:
            curveShape: The curve points. The number of points must be between
                2 and 14.

        Returns:

        """
        return self.app.adjustCurves(curveShape)

    def adjustLevels(
        self,
        inputRangeStart,
        inputRangeEnd,
        inputRangeGamma,
        outputRangeStart,
        outputRangeEnd,
    ) -> None:
        """Adjusts levels of the selected channels.

        Args:
            inputRangeStart:
            inputRangeEnd:
            inputRangeGamma:
            outputRangeStart:
            outputRangeEnd:
        """
        return self.app.adjustLevels(
            inputRangeStart,
            inputRangeEnd,
            inputRangeGamma,
            outputRangeStart,
            outputRangeEnd,
        )

    def applyAddNoise(self, amount, distribution, monochromatic) -> None:
        return self.app.applyAddNoise(amount, distribution, monochromatic)

    def applyDiffuseGlow(self, graininess, amount, clear_amount) -> None:
        """Applies the diffuse glow filter.

        Args:
            graininess: The amount of graininess. Range: 0 to 10.
            amount: The glow amount. Range: 0 to 20.
            clear_amount: The clear amount. Range: 0 to 20.

        Returns:

        """
        return self.app.applyDiffuseGlow(graininess, amount, clear_amount)

    def applyAverage(self) -> None:
        """Applies the average filter."""
        return self.app.applyAverage()

    def applyBlur(self) -> None:
        """Applies the blur filter."""
        return self.app.applyBlur()

    def applyBlurMore(self) -> None:
        """Applies the blur more filter."""
        return self.app.applyBlurMore()

    def applyClouds(self) -> None:
        """Applies the clouds filter."""
        return self.app.applyClouds()

    def applyCustomFilter(self, characteristics, scale, offset) -> None:
        """Applies the custom filter."""
        return self.app.applyCustomFilter(characteristics, scale, offset)

    def applyDeInterlace(self, eliminateFields, createFields) -> None:
        """Applies the de-interlace filter."""
        return self.app.applyDeInterlace(eliminateFields, createFields)

    def applyDespeckle(self) -> None:
        return self.app.applyDespeckle()

    def applyDifferenceClouds(self) -> None:
        """Applies the difference clouds filter."""
        return self.app.applyDifferenceClouds()

    def applyDisplace(
        self,
        horizontalScale,
        verticalScale,
        displacementType,
        undefinedAreas,
        displacementMapFile,
    ) -> None:
        """Applies the displace filter."""
        return self.app.applyDisplace(
            horizontalScale,
            verticalScale,
            displacementType,
            undefinedAreas,
            displacementMapFile,
        )

    def applyDustAndScratches(self, radius, threshold) -> None:
        """Applies the dust and scratches filter."""
        return self.app.applyDustAndScratches(radius, threshold)

    def applyGaussianBlur(self, radius) -> None:
        """Applies the gaussian blur filter."""
        return self.app.applyGaussianBlur(radius)

    def applyGlassEffect(
        self,
        distortion,
        smoothness,
        scaling,
        invert,
        texture,
        textureFile,
    ) -> None:
        return self.app.applyGlassEffect(
            distortion,
            smoothness,
            scaling,
            invert,
            texture,
            textureFile,
        )

    def applyHighPass(self, radius) -> None:
        """Applies the high pass filter."""
        return self.app.applyHighPass(radius)

    def applyLensBlur(
        self,
        source,
        focalDistance,
        invertDepthMap,
        shape,
        radius,
        bladeCurvature,
        rotation,
        brightness,
        threshold,
        amount,
        distribution,
        monochromatic,
    ) -> None:
        """Apply the lens blur filter."""
        return self.app.applyLensBlur(
            source,
            focalDistance,
            invertDepthMap,
            shape,
            radius,
            bladeCurvature,
            rotation,
            brightness,
            threshold,
            amount,
            distribution,
            monochromatic,
        )

    def applyLensFlare(self, brightness, flareCenter, lensType) -> None:
        return self.app.applyLensFlare(brightness, flareCenter, lensType)

    def applyMaximum(self, radius) -> None:
        return self.app.applyMaximum(radius)

    def applyMedianNoise(self, radius) -> None:
        return self.app.applyMedianNoise(radius)

    def applyMinimum(self, radius) -> None:
        return self.app.applyMinimum(radius)

    def applyMotionBlur(self, angle, radius) -> None:
        return self.app.applyMotionBlur(angle, radius)

    def applyNTSC(self) -> None:
        return self.app.applyNTSC()

    def applyOceanRipple(self, size, magnitude) -> None:
        return self.app.applyOceanRipple(size, magnitude)

    def applyOffset(self, horizontal, vertical, undefinedAreas) -> None:
        return self.app.applyOffset(horizontal, vertical, undefinedAreas)

    def applyPinch(self, amount) -> None:
        return self.app.applyPinch(amount)

    def remove(self) -> None:
        """Removes this layer from the document and deletes the instance."""
        # Note: Always raises an exception, even when successfully deleted.
        with suppress(Exception):
            return self.app.delete()

    def rasterize(self, target: RasterizeType) -> None:
        self.app.rasterize(target)

    def posterize(self, levels) -> None:
        self.app.posterize(levels)

    def move(
        self,
        relativeObject: Optional[Union['ArtLayer', 'Document', 'LayerSet']] = None,
        insertionLocation: Optional[ElementPlacement] = None
    ) -> None:
        """Move this `ArtLayer` relative to another `ArtLayer`, `Document`, or `LayerSet`.

        Args:
            relativeObject: An ArtLayer, Document, or LayerSet object to position this
                object relative to.
            insertionLocation: The placement strategy of this object relative to the
                relativeObject. Should be provided as an ElementPlacement enum.
        """
        self.app.move(relativeObject, insertionLocation)

    def merge(self):
        return ArtLayer(self.app.merge())

    def link(self, with_layer):
        self.app.link(with_layer)

    def unlink(self):
        """Unlink this layer from any linked layers."""
        self.app.unlink()

    def invert(self):
        self.app.invert()

    def duplicate(
        self, relativeObject: Optional[Union['ArtLayer', 'Document', 'LayerSet']] = None,
        insertionLocation: Optional[ElementPlacement] = None
    ):
        """Duplicate this `ArtLayer` and place the duplicate relative to another `ArtLayer`,
            `Document`, or `LayerSet`.

        Args:
            relativeObject: An ArtLayer, Document, or LayerSet object to position the
                duplicate object relative to.
            insertionLocation: The placement strategy of the duplicate object relative to
                the relativeObject. Should be provided as an ElementPlacement enum.

        Returns:
            The duplicate `ArtLayer` object.
        """
        return ArtLayer(self.app.duplicate(relativeObject, insertionLocation))
