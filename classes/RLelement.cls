VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RLelement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'abstract function to get a set of coordinates from a RLelement at a specified station
Public Function getCoo(station As Double, offset As Double) As Coo2D
End Function

'abstract function to get abscissa from a RLelement at a specified ordinate
Public Function getX(y As Double) As Double
End Function
'abstract function to get ordinate from a RLelement at a specified abscissa
Public Function getY(x As Double) As Double
End Function

'abstract function to get station and perpendicular/radial distance from a specified set of coordinates to the RLelement
Public Function getPerpOrRadFromPoint(px As Double, py As Double, station As Double, distFromElem As Double) As Boolean
End Function

'abstract method to set the stations of the Red Line Element
Public Sub setStations(startStation As Double)
End Sub

'abstract property to get start station of a RLelement
Public Property Get startStation() As Double
End Property

'abstract property to get end station of a RLelement
Public Property Get endStation() As Double
End Property

'abstract property to get start X of a RLelement
Public Property Get startX() As Double
End Property

'abstract property to get start Y of a RLelement
Public Property Get startY() As Double
End Property

'abstract property to get end X of a RLelement
Public Property Get endX() As Double
End Property

'abstract property to get end X of a RLelement
Public Property Get endY() As Double
End Property

'abstract property to get the length of a RLelement
Public Property Get getLength() As Double
End Property

'abstract property to get the type of a RLelement
Public Property Get getElemType() As Double
End Property

