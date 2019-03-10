Attribute VB_Name = "GeomRndGenerator"
''' TopoXL: Excel UDF library for land surveyors
''' Copyright (C) 2019 Bogdan Morosanu and Cristian Buse
''' This program is free software: you can redistribute it and/or modify
''' it under the terms of the GNU General Public License as published by
''' the Free Software Foundation, either version 3 of the License, or
''' (at your option) any later version.
'''
''' This program is distributed in the hope that it will be useful,
''' but WITHOUT ANY WARRANTY; without even the implied warranty of
''' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''' GNU General Public License for more details.
'''
''' You should have received a copy of the GNU General Public License
''' along with this program.  If not, see <https://www.gnu.org/licenses/>.

''========================================================================
'' Description:
'' Generates random geom objects
'' The main purpose of this class is to generate geom objects for test
'' and comparison against the geom elements before refactoring process
''========================================================================

'@Folder("TopoXL")

Option Explicit
Option Private Module

' Generates a random Point
Public Function PointRnd(ByVal minCoo As Double, _
                         ByVal maxCoo As Double) As Point
    Set PointRnd = New Point
    PointRnd.init MathLib.rndBetween(minCoo, maxCoo), MathLib.rndBetween(minCoo, maxCoo)
End Function

' Generates a random LineSegment
Public Function LineSegmentRnd(ByVal minMidCoo As Double, ByVal maxMidCoo As Double, _
                               ByVal minTheta As Double, ByVal maxTheta As Double, _
                               ByVal minLength As Double, ByVal maxLength As Double) As LineSegment
    Dim midX As Double
    Dim midY As Double
    Dim length As Double
    Dim theta As Double
    Dim dx_half As Double
    Dim dy_half As Double
    
    midX = MathLib.rndBetween(minMidCoo, maxMidCoo)
    midY = MathLib.rndBetween(minMidCoo, maxMidCoo)
    length = MathLib.rndBetween(minLength, maxLength)
    theta = MathLib.rndBetween(GeomLib.normalizeAngle(minTheta, 0), GeomLib.normalizeAngle(maxTheta, 0))
    
    dx_half = length * Math.Cos(theta)
    dy_half = length * Math.Sin(theta)
    
    Set LineSegmentRnd = New LineSegment
    LineSegmentRnd.init midX - dx_half, midY - dy_half, midX + dx_half, midY + dy_half
End Function

' Generates a random CircularArc
Public Function CircularArcRnd(ByVal minCenCoo As Double, ByVal maxCenCoo As Double, _
                               ByVal minStartTheta As Double, ByVal maxStartTheta As Double, _
                               ByVal minRadius As Double, ByVal maxRadius As Double, _
                               ByVal minLength As Double, ByVal maxLength As Double, _
                               ByVal curveDirection As CURVE_DIR) As CircularArc
    Dim cenX As Double
    Dim cenY As Double
    Dim sTheta As Double
    Dim radius As Double
    Dim length As Double
   
    Dim sX As Double
    Dim sY As Double
    
    cenX = MathLib.rndBetween(minCenCoo, maxCenCoo)
    cenY = MathLib.rndBetween(minCenCoo, maxCenCoo)
    
    sTheta = MathLib.rndBetween(GeomLib.normalizeAngle(minStartTheta, 0), _
                                GeomLib.normalizeAngle(maxStartTheta, 0))
    radius = MathLib.rndBetween(minRadius, maxRadius)
    length = MathLib.rndBetween(minLength, maxLength)
    
    sX = cenX + radius * Math.Cos(sTheta)
    sY = cenY + radius * Math.Sin(sTheta)
    
    Set CircularArcRnd = New CircularArc
    CircularArcRnd.initFromSCLD sX, sY, cenX, cenY, length, curveDirection
End Function




