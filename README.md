The *TopoXL* project is aimed to provide basic functionality for working with spatial data in a spreadsheet. This, in conjunction with basic data manipulation capabilities, expose a quick way of performing computation and analysis on data which is described by coordinates without the need of a graphical interface. The functionality is designed to be used as a user-defined function (UDF) library in *Microsoft Excel* environment. It is not designed to be a replacement for a CAD or GIS environment nor to provide the full functionality exposed by such environments. *TopoXL* deals with linear referencing and coordinate geometry problems but is limited to the capabilities of a spreadsheet. For a full description of its functionality please check the [Documentation](#Documentation) section of this document.

*TopoXL* is available under the [GNU General Public License](https://github.com/pastevalue/topoXL/blob/master/LICENSE)

# Installation

To be written

# Documentation

## General Notes
- The definitions, notations and concepts used by this library should be understood in the context of this library only.
- A point represents a location and is defined by coordinates XY(Z), where X is the *abscissa* and Y is the *ordinate* 
- An *Excel Array* is defined as a list of values. Some functions of this library return *Excel Arrays*. These functions should be used as *Array Formula* (*Ctrl + Shift + Enter*)


## Centerline Functions

### Terms and Concepts

A *Centerline*  is an imaginary line along the center of a road, railway, culvert, etc. It's made of one or more geometry elements of the following type: line segment, circular arc, Cornu/Euler spiral.
A geometry element has a starting point and a ending point defined.
The *measure* of a point along a geometry element is the distance measured from the start (or the end if the element is defined as *reversed*) to that point.
The *offset* of a point along a geometry element is the perpendicular distance between the point and the element. Offset values are negative if the points are on the left hand side (LHS) along the geometry element and positive if they are on the right hand side (RHS).

### Summary

*Centerline* referencing functions of this library deal with three scenarios:

- for a given *centerline* and *Point*, find the *measure* and the *offset*
- for a given *centerline*, *measure* and *offset*, find the *Point*
- for a given *centerline* and coordinate (X/Y), find the other coordinate (Y/X)

### *clPntByMeasOffset* - Centerline Point by Measure and Offset

#### Parameters

- *clName* (text): the name of the reference *centerline* used for computation
- *measure* (number): the *measure* value at which the output point coordinates will be computed
- *offset* (number): the *offset* value at which the output point coordinates will be computed

#### Result

Result type: *Excel Array* – two numbers

Returns the coordinates of the point calculated at the given *measure* and *offset*

#### Error(s) returned:

- *#N/A*: no *centerline* with the name *clName* was found
- *#NUM!*: *measure* value is out of range, that is, the *centerline* has no element such that *start measure <= measure <= end measure*

### *clMeasOffsetOfPnt* - Centerline Measure and Offset of Point

#### Parameters

- *clName* (text): the name of the reference *centerline* used for computation
- *X* (number): the X value of the input location
- *Y* (number): the Y value of the input location

#### Result

Result type: *Excel Array* – two numbers

Returns the centerline reference, *measure* and *offset*, of the given coordinates

#### Error(s) returned:

- *#N/A*: no *centerline* with the name *clName* was found
- *#NUM!*: coordinates not covered by the *centerline*, that is, the *centerline* has no element such that a perpendicular from the given coordinates to one of its elements exists

### *clYatX* - Centerline Y Value at X Value

#### Parameters

- *clName* (text): the name of the reference *centerline* used for computation
- *X* (number): the X value for which Y value must be computed

#### Result

Result type: number

Returns the Y value computed at the given X value. In the scenario of multiple *centerline* elements which satisfy the condition *Xmin <= X <= Xmax*, the first one found is used for computation.

#### Error(s) returned:

- *#N/A*: no centerline with the name *clName* was found
- *#NUM!*: X value not covered by the *centerline*, that is, the *centerline* has no element such that *Xmin <= X <= Xmax*

### *clXatY* - Centerline X Value at Y Value

#### Parameters

- *clName* (text): the name of the reference *centerline* used for computation
- *Y* (number): the Y value for which X value must be computed

#### Result

Result type: number

Returns the X value computed at the given Y value. In the scenario of multiple *centerline* elements which satisfy the condition *Ymin <= Y <= Ymax*, the first one found is used for computation.

#### Error(s) returned:

- *#N/A*: no centerline with the name *clName* was found
- *#NUM!*: Y value not covered by the *centerline*, that is, the *centerline* has no element such that *Ymin <= Y <= Ymax*

