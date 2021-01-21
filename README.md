The *TopoXL* project is aimed to provide basic functionality for working with spatial data in a spreadsheet. This, in conjunction with basic data manipulation capabilities, expose a quick way of performing computation and analysis on data which is described by coordinates without the need of a graphical interface. The functionality is designed to be used as a user-defined function (UDF) library in *Microsoft Excel* environment. It is not designed to be a replacement for a CAD or GIS environment nor to provide the full functionality exposed by such environments. *TopoXL* deals with linear referencing and coordinate geometry problems but is limited to the capabilities of a spreadsheet. For a full description of its functionality please check the [Documentation](#Documentation) section of this document.

*TopoXL* is available under the [GNU General Public License](https://github.com/pastevalue/topoXL/blob/master/LICENSE)

# Installation

To be written

# Documentation
## Centerline Functions
###### Summary
Centerline referencing functions used to compute:

- XY coordinates based on measure and offset values
- Measure and offset values based on coordinates
- X coordinate at a given Y
- Y coordinate at a given X

### Centerline Point by Measure and Offset
###### Name
clPntByMeasOffset
###### Parameters
- *clName* (text): the name of the reference centerline used for computation
- *measure* (number): the measure value at which the output point coordinates are computed
- *offset* (number): the offset value at which the output point coordinates are computed

###### Result
Returns (*Excel* array – two numbers) the coordinates of the point  calculated at the given measure and offset

###### Errors returned:
- #N/A: no centerline with the name = clName was found
- #NUM!: measure value is out of range, that is, there is no *centerline element* in the specified centerline such that *start measure <= measure <= end measure*
