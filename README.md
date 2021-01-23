The *TopoXL* project is aimed to provide basic functionality for working with spatial data in a spreadsheet. This, in conjunction with basic data manipulation capabilities, expose a quick way of performing computation and analysis on data which is described by coordinates without the need of a graphical interface. The functionality is designed to be used as a user-defined function (UDF) library in *Microsoft Excel* environment. It is not designed to be a replacement for a CAD or GIS environment nor to provide the full functionality exposed by such environments. *TopoXL* deals with linear referencing and coordinate geometry problems but is limited to the capabilities of a spreadsheet. For a full description of its functionality please check the [Documentation](#Documentation) section of this document.

*TopoXL* is available under the [GNU General Public License](https://github.com/pastevalue/topoXL/blob/master/LICENSE)

# Table of Contents

- [Installation](#installation)
- [UDF Documentation](#udf-documentation)
  * [General Notes](#general-notes)
  * [Centerline Functions](#centerline-functions)
    + [Terms and Concepts](#terms-and-concepts)
    + [Summary](#summary)
    + [Initialize *Centerline* Objects](#initialize--centerline--objects)
      - [Initialisation Table Description](#initialisation-table-description)
    + [Functions Help](#functions-help)
      - [*clPntByMeasOffset* - Centerline Point by Measure and Offset](#-clpntbymeasoffset----centerline-point-by-measure-and-offset)
        * [Parameters](#parameters)
        * [Result](#result)
        * [Error(s) returned:](#error-s--returned-)
      - [*clMeasOffsetOfPnt* - Centerline Measure and Offset of Point](#-clmeasoffsetofpnt----centerline-measure-and-offset-of-point)
        * [Parameters](#parameters-1)
        * [Result](#result-1)
        * [Error(s) returned:](#error-s--returned--1)
      - [*clYatX* - Centerline Y Value at X Value](#-clyatx----centerline-y-value-at-x-value)
        * [Parameters](#parameters-2)
        * [Result](#result-2)
        * [Error(s) returned:](#error-s--returned--2)
      - [*clXatY* - Centerline X Value at Y Value](#-clxaty----centerline-x-value-at-y-value)
        * [Parameters](#parameters-3)
        * [Result](#result-3)
        * [Error(s) returned:](#error-s--returned--3)


# Installation

To be written

# UDF Documentation
## General Notes
- The definitions, notations and concepts used by this library should be understood in the context of this library only.
- A point represents a location is defined by coordinates XY(Z), where X is the *abscissa* and Y is the *ordinate* 
- An *Excel Array* is defined as a list of values. Some functions of this library return *Excel Arrays*. These functions should be used as *Array Formula* (*Ctrl + Shift + Enter*)


## Centerline Functions
### Terms and Concepts
A *Centerline*  is an imaginary line along the center of a road, railway, culvert, etc. It's made of one or more geometry elements of the following type: line segment, circular arc, Cornu/Euler spiral.

A geometry element has a starting point and an ending point defined which relate to the start and end of its geometry.

The *measure* of a point along a geometry element is the distance to the point measured from the start point (or the end point if the element is defined as *reversed*).

The *offset* of a point along a geometry element is the perpendicular distance between the point and the element. Offset values are negative if the points are on the left hand side (LHS) along the geometry element and positive if they are on the right hand side (RHS).

The geometry element, measure, offset and their properties are shown in the figure below.
![](https://github.com/pastevalue/topoXL/blob/develop/docs/udf_cl/geom_elem.png)


### Summary
*Centerline* referencing functions of this library deal with three scenarios:

- for a given *centerline* and *Point*, calculate the *measure* and the *offset*
- for a given *centerline*, *measure* and *offset*, calculate the *Point*
- for a given *centerline* and coordinate (X/Y), calculate the other coordinate (Y/X)

*Centerline* referencing functions work only if a *centerline* object has been successfully initialized prior to the function call (see [Initialize *Centerline* Objects](###Initialize *Centerline* Objects)).

### Initialize *Centerline* Objects
A *Centerline* object can be initialized from an *Excel table object*/VBA ListObject (not *Excel data range*). The initialisation is done when the workbook is opened (triggered by the workbook open event).

#### Initialisation Table Description
Each row of the table represents a *centerline* element.

The table must meet the following requirements:

- table name starts with *tblCL*
- the table header must include *GeomType*, *InitType*, *Reversed* and *Measure* columns. Valid values are required for all of these columns
- the table header must include the columns relevant for the geometry properties definition. Based on the value of  *GeomType*, the following columns and valid values are required:
	+ *LineSegment*: *StartX*, *StartY*, *EndX*, *EndY*
	+ *CircularArc*:
		* *StartX*, *StartY*, *EndX*, *EndY*, *Radius*, *CurveDirection*
		* *StartX*, *StartY*, *CenterX*, *CenterY*, *Length*, *CurveDirection*
	+ *ClothoidArc*: *StartX*, *StartY*, *Length*, *Radius*, *CurveDirection*, StartTheta

The meaning and expected values of each column is detailed below:

- *GeomType* is the *centerline* geometry type. The accepted values are: *LineSegment*, *CircularArc*, *ClothoidArc*;
- *InitType* is an indicator for what geometry properties will be used to initialize the *centerline* element. Based on the *GeomType* attribute, the accepted values are as follows:
	+ *LineSegment*: *SE* (Start - End)
	+ *CircularArc*:
		* *SERD* (Start, End, Radius, Curve direction)
		* *SCLD* (Start, Center, Length, Curve direction)
	+ ClothoidArc: SLRDT (Start, Length, Radius, Curve direction, Start theta)
- *StartX* is the X coordinate of the geometry's start point
- *StartY* is the Y coordinate of the geometry's start point
- *EndX* is the X coordinate of the geometry's end point
- *EndY* is the Y coordinate of the geometry's end point
- *CenterX* is the X coordinate of the geometry's center point
- *CenterY* is the Y coordinate of the geometry's center point
- *Length* is the length of the geometry
- *Radius* is the radius length of the geometry
- *CurveDirection* is the direction of the curve. The accepted values are *CW* (clocwise) and *CCW* (counter-clocwise)
- *StartTheta* is the angle in radians, counter-clocwise measured, measured between the tangent at entrance on a spiral and Ox axis
- *Reversed* is an indicator which defines where the *measure* of a geometry starts from. Accepted values are *TRUE* or *FALSE*. If a *centerline* element is reversed, the *measure* is applied starting from the *end* of the geometry
- *Measure* is a number which indicates the starting value for measuring lengths along geometries

An example of a valid *centerline* input table is provided below.

|GeomType|InitType|StartX|StartY|EndX|EndY|CenterX|CenterY|Length|Radius|CurveDirection|StartTheta|Reversed|Measure
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
|LineSegment|SE|198764.3459|304156.5693|198602.8152|304353.6703|||||||FALSE|115239.0686
|ClothoidArc|SLRDT|198602.8152|304353.6703|||||100.0000|1500.0000|CCW|2.25733470|FALSE|115493.9037
|CircularArc|SERD|198538.5765|304430.3020|198536.2834|304432.9111||||1500.0000|CCW||FALSE|115593.9037


### Functions Help
The *centerline* functions are stored in the *UDF_CL* module. Their names are prefixed with *cl*.
#### *clPntByMeasOffset* - Centerline Point by Measure and Offset

##### Parameters
- *clName* (text): the name of the reference *centerline* used for computation
- *measure* (number): the *measure* value at which the output point coordinates will be computed
- *offset* (number): the *offset* value at which the output point coordinates will be computed

##### Result

Result type: *Excel Array* – two numbers

Returns the coordinates of the point calculated at the given *measure* and *offset*
##### Errors returned:
- *#N/A*: no *centerline* with the name *clName* was found
- *#NUM!*: *measure* value is out of range, that is, the *centerline* has no element such that *start measure <= measure <= end measure*

#### *clMeasOffsetOfPnt* - Centerline Measure and Offset of Point
##### Parameters
- *clName* (text): the name of the reference *centerline* used for computation
- *X* (number): the X value of the input location
- *Y* (number): the Y value of the input location

##### Result
Result type: *Excel Array* – two numbers

Returns the centerline reference, *measure* and *offset*, of the given coordinates

##### Errors returned:
- *#N/A*: no *centerline* with the name *clName* was found
- *#NUM!*: coordinates not covered by the *centerline*, that is, the *centerline* has no element such that a perpendicular from the given coordinates to one of its elements exists

#### *clYatX* - Centerline Y Value at X Value
##### Parameters
- *clName* (text): the name of the reference *centerline* used for computation
- *X* (number): the X value for which Y value must be computed

##### Result

Result type: number

Returns the Y value computed at the given X value. In the scenario of multiple *centerline* elements which satisfy the condition *Xmin <= X <= Xmax*, the first one found is used for computation.

##### Errors returned:
- *#N/A*: no centerline with the name *clName* was found
- *#NUM!*: X value not covered by the *centerline*, that is, the *centerline* has no element such that *Xmin <= X <= Xmax*

#### *clXatY* - Centerline X Value at Y Value
##### Parameters
- *clName* (text): the name of the reference *centerline* used for computation
- *Y* (number): the Y value for which X value must be computed

##### Result

Result type: number

Returns the X value computed at the given Y value. In the scenario of multiple *centerline* elements which satisfy the condition *Ymin <= Y <= Ymax*, the first one found is used for computation.

##### Errors returned:
- *#N/A*: no centerline with the name *clName* was found
- *#NUM!*: Y value not covered by the *centerline*, that is, the *centerline* has no element such that *Ymin <= Y <= Ymax*



