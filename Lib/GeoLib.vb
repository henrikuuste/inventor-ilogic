' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' GeoLib - Geometric Calculations Library
' 
' Functions for geometric calculations, measurements, and projections.
' Wraps and extends UtilsLib geometry functions with additional capabilities.
'
' Depends on: UtilsLib.vb
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/GeoLib.vb"
'
' ============================================================================

Option Strict Off
Imports Inventor

Public Module GeoLib

    ' ============================================================================
    ' SECTION 1: Plane Geometry (wrappers for UtilsLib + extensions)
    ' ============================================================================
    
    ''' <summary>
    ''' Extract Plane from various geometry objects.
    ''' Wrapper for UtilsLib.GetPlaneGeometry.
    ''' </summary>
    Public Function GetPlaneGeometry(obj As Object) As Plane
        Return UtilsLib.GetPlaneGeometry(obj)
    End Function
    
    ''' <summary>
    ''' Get plane normal vector.
    ''' Wrapper for UtilsLib.GetPlaneNormal.
    ''' </summary>
    Public Function GetPlaneNormal(planeObj As Object) As UnitVector
        Return UtilsLib.GetPlaneNormal(planeObj)
    End Function
    
    ''' <summary>
    ''' Get plane root point (a point on the plane).
    ''' </summary>
    Public Function GetPlaneRootPoint(planeObj As Object) As Point
        Dim plane As Plane = GetPlaneGeometry(planeObj)
        If plane IsNot Nothing Then
            Return plane.RootPoint
        End If
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Calculate distance between two parallel planes.
    ''' Returns distance in internal units (cm).
    ''' Wrapper for UtilsLib.MeasurePlaneDistance.
    ''' </summary>
    Public Function MeasurePlaneDistance(plane1 As Object, plane2 As Object) As Double
        Return UtilsLib.MeasurePlaneDistance(plane1, plane2)
    End Function

    ' ============================================================================
    ' SECTION 2: Axis Geometry
    ' ============================================================================
    
    ''' <summary>
    ''' Get axis direction from work axis or edge.
    ''' </summary>
    Public Function GetAxisDirection(axisObj As Object) As UnitVector
        If axisObj Is Nothing Then Return Nothing
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(axisObj)
        
        Select Case typeName
            Case "WorkAxis"
                Return CType(axisObj, WorkAxis).Line.Direction
            Case "WorkAxisProxy"
                Return CType(axisObj, WorkAxisProxy).Line.Direction
            Case Else
                ' Use UtilsLib for edges
                Return UtilsLib.GetAxisDirection(axisObj)
        End Select
    End Function
    
    ''' <summary>
    ''' Get axis origin point (root point of the line).
    ''' </summary>
    Public Function GetAxisOrigin(axisObj As Object) As Point
        If axisObj Is Nothing Then Return Nothing
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(axisObj)
        
        Select Case typeName
            Case "WorkAxis"
                Return CType(axisObj, WorkAxis).Line.RootPoint
            Case "WorkAxisProxy"
                Return CType(axisObj, WorkAxisProxy).Line.RootPoint
            Case Else
                ' Use UtilsLib for edges
                Dim pt As Point = Nothing
                Dim dir As UnitVector = Nothing
                If UtilsLib.GetAxisProperties(axisObj, pt, dir) Then
                    Return pt
                End If
                Return Nothing
        End Select
    End Function

    ' ============================================================================
    ' SECTION 3: Point Projections
    ' ============================================================================
    
    ''' <summary>
    ''' Project a point onto an axis line.
    ''' Returns the projected point on the axis.
    ''' </summary>
    Public Function ProjectPointOntoAxis(app As Inventor.Application, _
                                          point As Point, _
                                          axisOrigin As Point, _
                                          axisDirection As UnitVector) As Point
        If point Is Nothing OrElse axisOrigin Is Nothing OrElse axisDirection Is Nothing Then
            Return Nothing
        End If
        
        ' Vector from axis origin to point
        Dim toPointX As Double = point.X - axisOrigin.X
        Dim toPointY As Double = point.Y - axisOrigin.Y
        Dim toPointZ As Double = point.Z - axisOrigin.Z
        
        ' Project onto axis direction (dot product)
        Dim distance As Double = toPointX * axisDirection.X + _
                                  toPointY * axisDirection.Y + _
                                  toPointZ * axisDirection.Z
        
        ' Projected point = origin + distance * direction
        Return app.TransientGeometry.CreatePoint( _
            axisOrigin.X + distance * axisDirection.X, _
            axisOrigin.Y + distance * axisDirection.Y, _
            axisOrigin.Z + distance * axisDirection.Z)
    End Function
    
    ''' <summary>
    ''' Get the signed distance of a point along an axis from the axis origin.
    ''' Positive = in direction of axis, Negative = opposite direction.
    ''' Returns distance in internal units (cm).
    ''' </summary>
    Public Function GetPointDistanceAlongAxis(point As Point, _
                                               axisOrigin As Point, _
                                               axisDirection As UnitVector) As Double
        If point Is Nothing OrElse axisOrigin Is Nothing OrElse axisDirection Is Nothing Then
            Return 0
        End If
        
        ' Vector from axis origin to point
        Dim toPointX As Double = point.X - axisOrigin.X
        Dim toPointY As Double = point.Y - axisOrigin.Y
        Dim toPointZ As Double = point.Z - axisOrigin.Z
        
        ' Dot product = signed distance along axis
        Return toPointX * axisDirection.X + _
               toPointY * axisDirection.Y + _
               toPointZ * axisDirection.Z
    End Function

    ' ============================================================================
    ' SECTION 4: Occurrence Geometry
    ' ============================================================================
    
    ''' <summary>
    ''' Get occurrence position (transformation origin) as a Point.
    ''' </summary>
    Public Function GetOccurrencePosition(app As Inventor.Application, _
                                           occ As ComponentOccurrence) As Point
        If occ Is Nothing Then Return Nothing
        
        Try
            Dim matrix As Matrix = occ.Transformation
            Return app.TransientGeometry.CreatePoint( _
                matrix.Cell(1, 4), _
                matrix.Cell(2, 4), _
                matrix.Cell(3, 4))
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Get occurrence position along an axis (signed distance from axis origin).
    ''' Returns distance in internal units (cm).
    ''' </summary>
    Public Function GetOccurrencePositionAlongAxis(app As Inventor.Application, _
                                                    occ As ComponentOccurrence, _
                                                    axisOrigin As Point, _
                                                    axisDirection As UnitVector) As Double
        Dim occPos As Point = GetOccurrencePosition(app, occ)
        If occPos Is Nothing Then Return 0
        
        Return GetPointDistanceAlongAxis(occPos, axisOrigin, axisDirection)
    End Function

    ' ============================================================================
    ' SECTION 5: Principal Plane Detection
    ' ============================================================================
    
    ''' <summary>
    ''' Find the principal plane (origin work plane) of a part that is most
    ''' perpendicular to the given axis direction.
    ''' 
    ''' Returns the work plane proxy in assembly context, or Nothing if failed.
    ''' 
    ''' planeIndex meanings: 1=YZ, 2=XZ, 3=XY
    ''' For axis in X direction: YZ plane (index 1) is perpendicular
    ''' For axis in Y direction: XZ plane (index 2) is perpendicular
    ''' For axis in Z direction: XY plane (index 3) is perpendicular
    ''' </summary>
    Public Function FindPrincipalPlane(occ As ComponentOccurrence, _
                                        axisDirection As UnitVector) As Object
        If occ Is Nothing OrElse axisDirection Is Nothing Then Return Nothing
        
        Try
            Dim compDef As ComponentDefinition = occ.Definition
            Dim bestPlane As Object = Nothing
            Dim bestDot As Double = -1
            
            ' Check all three origin planes (YZ=1, XZ=2, XY=3)
            For planeIdx As Integer = 1 To 3
                Try
                    Dim wp As WorkPlane = compDef.WorkPlanes.Item(planeIdx)
                    
                    ' Create proxy in assembly context
                    Dim proxyResult As Object = Nothing
                    occ.CreateGeometryProxy(wp, proxyResult)
                    
                    If proxyResult IsNot Nothing Then
                        ' Get plane normal in assembly coordinates
                        Dim proxyGeom As Plane = GetPlaneGeometry(proxyResult)
                        If proxyGeom IsNot Nothing Then
                            Dim wpNormal As UnitVector = proxyGeom.Normal
                            
                            ' Check how parallel the normal is to the axis
                            ' (perpendicular plane has normal parallel to axis)
                            Dim dot As Double = Math.Abs( _
                                wpNormal.X * axisDirection.X + _
                                wpNormal.Y * axisDirection.Y + _
                                wpNormal.Z * axisDirection.Z)
                            
                            If dot > bestDot Then
                                bestDot = dot
                                bestPlane = proxyResult
                            End If
                        End If
                    End If
                Catch
                End Try
            Next
            
            ' Only return if we found a reasonably perpendicular plane
            If bestDot > 0.9 Then
                Return bestPlane
            End If
            
        Catch
        End Try
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Get the index of the origin work plane most perpendicular to axis.
    ''' Returns: 1=YZ, 2=XZ, 3=XY, or 0 if can't determine.
    ''' </summary>
    Public Function GetPrincipalPlaneIndex(occ As ComponentOccurrence, _
                                            axisDirection As UnitVector) As Integer
        If occ Is Nothing OrElse axisDirection Is Nothing Then Return 0
        
        Try
            Dim compDef As ComponentDefinition = occ.Definition
            Dim bestIndex As Integer = 0
            Dim bestDot As Double = -1
            
            For planeIdx As Integer = 1 To 3
                Try
                    Dim wp As WorkPlane = compDef.WorkPlanes.Item(planeIdx)
                    
                    ' Create proxy
                    Dim proxyResult As Object = Nothing
                    occ.CreateGeometryProxy(wp, proxyResult)
                    
                    If proxyResult IsNot Nothing Then
                        Dim proxyGeom As Plane = GetPlaneGeometry(proxyResult)
                        If proxyGeom IsNot Nothing Then
                            Dim dot As Double = Math.Abs( _
                                proxyGeom.Normal.X * axisDirection.X + _
                                proxyGeom.Normal.Y * axisDirection.Y + _
                                proxyGeom.Normal.Z * axisDirection.Z)
                            
                            If dot > bestDot Then
                                bestDot = dot
                                bestIndex = planeIdx
                            End If
                        End If
                    End If
                Catch
                End Try
            Next
            
            If bestDot > 0.9 Then
                Return bestIndex
            End If
        Catch
        End Try
        
        Return 0
    End Function

    ' ============================================================================
    ' SECTION 6: Direction Utilities
    ' ============================================================================
    
    ''' <summary>
    ''' Get plane normal oriented to point from one plane toward another.
    ''' Returns the normal of plane1, flipped if needed to point toward plane2.
    ''' Wrapper for UtilsLib.GetPlaneNormalTowardPlane.
    ''' </summary>
    Public Function GetPlaneNormalTowardPlane(app As Inventor.Application, _
                                               plane1 As Object, _
                                               plane2 As Object) As UnitVector
        Return UtilsLib.GetPlaneNormalTowardPlane(app, plane1, plane2)
    End Function
    
    ''' <summary>
    ''' Create a unit vector from components.
    ''' </summary>
    Public Function CreateUnitVector(app As Inventor.Application, _
                                      x As Double, y As Double, z As Double) As UnitVector
        Dim length As Double = Math.Sqrt(x * x + y * y + z * z)
        If length < 0.0001 Then Return Nothing
        
        Return app.TransientGeometry.CreateUnitVector(x / length, y / length, z / length)
    End Function
    
    ''' <summary>
    ''' Flip a unit vector direction.
    ''' </summary>
    Public Function FlipDirection(app As Inventor.Application, _
                                   dir As UnitVector) As UnitVector
        If dir Is Nothing Then Return Nothing
        Return app.TransientGeometry.CreateUnitVector(-dir.X, -dir.Y, -dir.Z)
    End Function
    
    ''' <summary>
    ''' Check if two vectors are parallel (dot product close to ±1).
    ''' </summary>
    Public Function AreVectorsParallel(v1 As UnitVector, v2 As UnitVector, _
                                        Optional tolerance As Double = 0.99) As Boolean
        If v1 Is Nothing OrElse v2 Is Nothing Then Return False
        
        Dim dot As Double = Math.Abs(v1.X * v2.X + v1.Y * v2.Y + v1.Z * v2.Z)
        Return dot > tolerance
    End Function
    
    ''' <summary>
    ''' Check if two vectors point in the same direction (dot product close to +1).
    ''' </summary>
    Public Function AreVectorsSameDirection(v1 As UnitVector, v2 As UnitVector, _
                                             Optional tolerance As Double = 0.99) As Boolean
        If v1 Is Nothing OrElse v2 Is Nothing Then Return False
        
        Dim dot As Double = v1.X * v2.X + v1.Y * v2.Y + v1.Z * v2.Z
        Return dot > tolerance
    End Function

End Module
