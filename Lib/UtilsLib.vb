' ============================================================================
' UtilsLib - Generic Utility Functions for Inventor iLogic
' 
' Reusable geometry, measurement, and UI utility functions.
' These functions have no dependencies on specific features or workflows.
'
' Usage: AddVbFile "Lib/UtilsLib.vb"
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

Imports Inventor

Public Module UtilsLib

    ' ============================================================================
    ' SECTION 1: Geometry Extraction
    ' ============================================================================

    ''' <summary>
    ''' Extract Point from various geometry objects.
    ''' Supports: WorkPoint, WorkPointProxy, Vertex, VertexProxy, SketchPoint, Point
    ''' </summary>
    Public Function GetPointGeometry(obj As Object) As Point
        If obj Is Nothing Then Return Nothing
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(obj)
        
        Select Case typeName
            Case "WorkPoint"
                Return CType(obj, WorkPoint).Point
            Case "WorkPointProxy"
                Return CType(obj, WorkPointProxy).Point
            Case "Vertex"
                Return CType(obj, Vertex).Point
            Case "VertexProxy"
                Return CType(obj, VertexProxy).Point
            Case "SketchPoint"
                Return CType(obj, SketchPoint).Geometry3d
            Case "Point"
                Return CType(obj, Point)
        End Select
        Return Nothing
    End Function

    ''' <summary>
    ''' Extract Plane from various geometry objects.
    ''' Supports: WorkPlane, WorkPlaneProxy, Face, FaceProxy (planar only)
    ''' </summary>
    Public Function GetPlaneGeometry(obj As Object) As Plane
        If obj Is Nothing Then Return Nothing
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(obj)
        
        Select Case typeName
            Case "WorkPlane"
                Return CType(obj, WorkPlane).Plane
            Case "WorkPlaneProxy"
                Return CType(obj, WorkPlaneProxy).Plane
            Case "Face"
                Dim geom As Object = CType(obj, Face).Geometry
                If TypeOf geom Is Plane Then Return CType(geom, Plane)
            Case "FaceProxy"
                Dim geom As Object = CType(obj, FaceProxy).Geometry
                If TypeOf geom Is Plane Then Return CType(geom, Plane)
        End Select
        Return Nothing
    End Function

    ''' <summary>
    ''' Extract Line from various geometry objects.
    ''' Supports: WorkAxis, WorkAxisProxy, Edge, EdgeProxy (linear only)
    ''' Note: Returns Nothing for LineSegment edges - use GetAxisProperties instead.
    ''' </summary>
    Public Function GetLineGeometry(obj As Object) As Line
        If obj Is Nothing Then Return Nothing
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(obj)
        
        Select Case typeName
            Case "WorkAxis"
                Return CType(obj, WorkAxis).Line
            Case "WorkAxisProxy"
                Return CType(obj, WorkAxisProxy).Line
            Case "Edge"
                Dim geom As Object = CType(obj, Edge).Geometry
                If TypeOf geom Is Line Then Return CType(geom, Line)
            Case "EdgeProxy"
                Dim geom As Object = CType(obj, EdgeProxy).Geometry
                If TypeOf geom Is Line Then Return CType(geom, Line)
        End Select
        Return Nothing
    End Function

    ''' <summary>
    ''' Get axis properties (point and direction) from any linear object.
    ''' Supports: WorkAxis, WorkAxisProxy, Edge, EdgeProxy (with Line or LineSegment geometry)
    ''' </summary>
    Public Function GetAxisProperties(axis As Object, ByRef point As Point, ByRef direction As UnitVector) As Boolean
        point = Nothing
        direction = Nothing
        
        If axis Is Nothing Then Return False
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(axis)
        
        Select Case typeName
            Case "WorkAxis"
                Dim line As Line = CType(axis, WorkAxis).Line
                point = line.RootPoint
                direction = line.Direction
                Return True
                
            Case "WorkAxisProxy"
                Dim line As Line = CType(axis, WorkAxisProxy).Line
                point = line.RootPoint
                direction = line.Direction
                Return True
                
            Case "Edge"
                Dim geom As Object = CType(axis, Edge).Geometry
                If TypeOf geom Is Line Then
                    Dim line As Line = CType(geom, Line)
                    point = line.RootPoint
                    direction = line.Direction
                    Return True
                ElseIf TypeOf geom Is LineSegment Then
                    Dim seg As LineSegment = CType(geom, LineSegment)
                    point = seg.StartPoint
                    direction = seg.Direction
                    Return True
                End If
                
            Case "EdgeProxy"
                Dim geom As Object = CType(axis, EdgeProxy).Geometry
                If TypeOf geom Is Line Then
                    Dim line As Line = CType(geom, Line)
                    point = line.RootPoint
                    direction = line.Direction
                    Return True
                ElseIf TypeOf geom Is LineSegment Then
                    Dim seg As LineSegment = CType(geom, LineSegment)
                    point = seg.StartPoint
                    direction = seg.Direction
                    Return True
                End If
        End Select
        
        Return False
    End Function

    ''' <summary>
    ''' Get direction from axis/edge object.
    ''' </summary>
    Public Function GetAxisDirection(axis As Object) As UnitVector
        Dim pt As Point = Nothing
        Dim dir As UnitVector = Nothing
        
        If GetAxisProperties(axis, pt, dir) Then
            Return dir
        End If
        
        Return Nothing
    End Function

    ''' <summary>
    ''' Get plane normal as direction.
    ''' </summary>
    Public Function GetPlaneNormal(planeObj As Object) As UnitVector
        Dim plane As Plane = GetPlaneGeometry(planeObj)
        If plane IsNot Nothing Then
            Return plane.Normal
        End If
        Return Nothing
    End Function

    ' ============================================================================
    ' SECTION 2: Distance and Direction Calculations
    ' ============================================================================

    ''' <summary>
    ''' Calculate distance between two parallel planes.
    ''' Returns distance in internal units (cm).
    ''' </summary>
    Public Function MeasurePlaneDistance(plane1 As Object, plane2 As Object) As Double
        Dim p1 As Plane = GetPlaneGeometry(plane1)
        Dim p2 As Plane = GetPlaneGeometry(plane2)
        
        If p1 Is Nothing OrElse p2 Is Nothing Then Return 0
        
        ' Calculate signed distance from plane1's root point to plane2
        Dim rootPoint As Point = p1.RootPoint
        Dim normal As UnitVector = p2.Normal
        Dim d As Double = (rootPoint.X - p2.RootPoint.X) * normal.X + _
                          (rootPoint.Y - p2.RootPoint.Y) * normal.Y + _
                          (rootPoint.Z - p2.RootPoint.Z) * normal.Z
        
        Return Math.Abs(d)
    End Function

    ''' <summary>
    ''' Calculate distance between two points.
    ''' Returns distance in internal units (cm).
    ''' </summary>
    Public Function MeasurePointDistance(point1 As Object, point2 As Object) As Double
        Dim pt1 As Point = GetPointGeometry(point1)
        Dim pt2 As Point = GetPointGeometry(point2)
        
        If pt1 Is Nothing OrElse pt2 Is Nothing Then Return 0
        
        Dim dx As Double = pt2.X - pt1.X
        Dim dy As Double = pt2.Y - pt1.Y
        Dim dz As Double = pt2.Z - pt1.Z
        
        Return Math.Sqrt(dx * dx + dy * dy + dz * dz)
    End Function

    ''' <summary>
    ''' Get direction from point1 to point2.
    ''' </summary>
    Public Function GetDirectionBetweenPoints(app As Inventor.Application, point1 As Object, point2 As Object) As UnitVector
        Dim pt1 As Point = GetPointGeometry(point1)
        Dim pt2 As Point = GetPointGeometry(point2)
        
        If pt1 Is Nothing OrElse pt2 Is Nothing Then Return Nothing
        
        Dim dx As Double = pt2.X - pt1.X
        Dim dy As Double = pt2.Y - pt1.Y
        Dim dz As Double = pt2.Z - pt1.Z
        
        Dim len As Double = Math.Sqrt(dx * dx + dy * dy + dz * dz)
        If len < 0.0001 Then Return Nothing
        
        Return app.TransientGeometry.CreateUnitVector(dx / len, dy / len, dz / len)
    End Function

    ''' <summary>
    ''' Get intersection point of axis and plane.
    ''' </summary>
    Public Function GetAxisPlaneIntersection(app As Inventor.Application, axis As Object, planeObj As Object) As Point
        Dim p0 As Point = Nothing
        Dim d As UnitVector = Nothing
        
        If Not GetAxisProperties(axis, p0, d) Then Return Nothing
        
        Dim plane As Plane = GetPlaneGeometry(planeObj)
        If plane Is Nothing Then Return Nothing
        
        Dim n As UnitVector = plane.Normal
        Dim q As Point = plane.RootPoint
        
        Dim denom As Double = n.X * d.X + n.Y * d.Y + n.Z * d.Z
        If Math.Abs(denom) < 0.0001 Then
            ' Line is parallel to plane
            Return GetClosestPointOnAxisToPlane(app, axis, planeObj)
        End If
        
        Dim numer As Double = n.X * (q.X - p0.X) + n.Y * (q.Y - p0.Y) + n.Z * (q.Z - p0.Z)
        Dim t As Double = numer / denom
        
        Return app.TransientGeometry.CreatePoint( _
            p0.X + t * d.X, _
            p0.Y + t * d.Y, _
            p0.Z + t * d.Z)
    End Function

    ''' <summary>
    ''' Get closest point on axis to a plane (for parallel axis/plane cases).
    ''' </summary>
    Public Function GetClosestPointOnAxisToPlane(app As Inventor.Application, axis As Object, planeObj As Object) As Point
        Dim p0 As Point = Nothing
        Dim axisDir As UnitVector = Nothing
        
        If Not GetAxisProperties(axis, p0, axisDir) Then Return Nothing
        
        Dim plane As Plane = GetPlaneGeometry(planeObj)
        If plane Is Nothing Then Return Nothing
        
        Dim n As UnitVector = plane.Normal
        Dim q As Point = plane.RootPoint
        
        ' Distance from p0 to plane along normal
        Dim dist As Double = (p0.X - q.X) * n.X + (p0.Y - q.Y) * n.Y + (p0.Z - q.Z) * n.Z
        
        ' Project p0 onto plane
        Return app.TransientGeometry.CreatePoint( _
            p0.X - dist * n.X, _
            p0.Y - dist * n.Y, _
            p0.Z - dist * n.Z)
    End Function

    ''' <summary>
    ''' Project a point onto a plane along the plane's normal.
    ''' </summary>
    Public Function ProjectPointOntoPlane(app As Inventor.Application, pt As Point, planeObj As Object) As Point
        If pt Is Nothing Then Return Nothing
        
        Dim plane As Plane = GetPlaneGeometry(planeObj)
        If plane Is Nothing Then Return Nothing
        
        Dim n As UnitVector = plane.Normal
        Dim q As Point = plane.RootPoint
        
        ' Distance from point to plane along normal
        Dim dist As Double = (pt.X - q.X) * n.X + (pt.Y - q.Y) * n.Y + (pt.Z - q.Z) * n.Z
        
        ' Project point onto plane
        Return app.TransientGeometry.CreatePoint( _
            pt.X - dist * n.X, _
            pt.Y - dist * n.Y, _
            pt.Z - dist * n.Z)
    End Function

    ''' <summary>
    ''' Get plane normal oriented toward another plane.
    ''' Returns the normal of plane1, flipped if needed to point toward plane2.
    ''' </summary>
    Public Function GetPlaneNormalTowardPlane(app As Inventor.Application, plane1 As Object, plane2 As Object) As UnitVector
        Dim p1 As Plane = GetPlaneGeometry(plane1)
        Dim p2 As Plane = GetPlaneGeometry(plane2)
        
        If p1 Is Nothing OrElse p2 Is Nothing Then Return Nothing
        
        Dim normal As UnitVector = p1.Normal
        
        ' Check if normal points toward plane2
        Dim toPlane2 As Vector = app.TransientGeometry.CreateVector( _
            p2.RootPoint.X - p1.RootPoint.X, _
            p2.RootPoint.Y - p1.RootPoint.Y, _
            p2.RootPoint.Z - p1.RootPoint.Z)
        
        Dim dot As Double = normal.X * toPlane2.X + normal.Y * toPlane2.Y + normal.Z * toPlane2.Z
        
        If dot < 0 Then
            normal = app.TransientGeometry.CreateUnitVector(-normal.X, -normal.Y, -normal.Z)
        End If
        
        Return normal
    End Function

    ' ============================================================================
    ' SECTION 3: Object Picking (Work Features Only)
    ' ============================================================================

    ''' <summary>
    ''' Pick a WorkPoint only.
    ''' </summary>
    Public Function PickPoint(app As Inventor.Application, prompt As String) As Object
        Try
            Return app.CommandManager.Pick(SelectionFilterEnum.kWorkPointFilter, prompt)
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Pick a WorkAxis only.
    ''' </summary>
    Public Function PickAxis(app As Inventor.Application, prompt As String) As Object
        Try
            Return app.CommandManager.Pick(SelectionFilterEnum.kWorkAxisFilter, prompt)
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Pick a WorkPlane only.
    ''' </summary>
    Public Function PickPlane(app As Inventor.Application, prompt As String) As Object
        Try
            Return app.CommandManager.Pick(SelectionFilterEnum.kWorkPlaneFilter, prompt)
        Catch
            Return Nothing
        End Try
    End Function

    ' ============================================================================
    ' SECTION 4: Display Names
    ' ============================================================================

    ''' <summary>
    ''' Get a human-readable display name for a geometry object.
    ''' </summary>
    Public Function GetObjectDisplayName(obj As Object) As String
        If obj Is Nothing Then Return "(not selected)"
        
        Try
            Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(obj)
            
            ' Try to get the Name property (work features have names)
            Dim objName As String = ""
            Try
                objName = CStr(CallByName(obj, "Name", CallType.Get))
            Catch
            End Try
            
            ' Clean up proxy suffix for display
            Dim displayType As String = typeName
            If displayType.EndsWith("Proxy") Then
                displayType = displayType.Substring(0, displayType.Length - 5)
            End If
            
            ' Return name if available
            If objName <> "" Then
                Return displayType & ": " & objName
            End If
            
            ' For unnamed geometry, add identifying info
            Select Case displayType
                Case "Vertex"
                    Try
                        Dim pt As Point = GetPointGeometry(obj)
                        If pt IsNot Nothing Then
                            Return "Vertex @ (" & Math.Round(pt.X, 2).ToString() & ", " & _
                                   Math.Round(pt.Y, 2).ToString() & ", " & _
                                   Math.Round(pt.Z, 2).ToString() & ")"
                        End If
                    Catch
                    End Try
                    Return "Vertex"
                Case "Face"
                    Return "Face (planar)"
                Case "Edge"
                    Return "Edge (linear)"
                Case Else
                    Return displayType
            End Select
        Catch
            Return "(unknown)"
        End Try
    End Function

    ''' <summary>
    ''' Get the name of a named object (work features, etc).
    ''' Returns empty string if object has no name.
    ''' </summary>
    Public Function GetObjectName(obj As Object) As String
        If obj Is Nothing Then Return ""
        Try
            Return CStr(CallByName(obj, "Name", CallType.Get))
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Get a stable reference string for a work feature.
    ''' For proxies (work features inside components), includes occurrence path using InternalNames.
    ''' Format: "@OccInternalName1/OccInternalName2|WorkFeatureName" or just "WorkFeatureName" for assembly-level.
    ''' Uses InternalName which survives occurrence renames.
    ''' </summary>
    Public Function GetWorkFeatureReference(obj As Object) As String
        If obj Is Nothing Then Return ""
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(obj)
        
        ' Check if it's a proxy (work feature inside a component)
        If typeName.EndsWith("Proxy") Then
            Try
                ' Get the containing occurrence from the proxy
                Dim containingOcc As Object = CallByName(obj, "ContainingOccurrence", CallType.Get)
                If containingOcc IsNot Nothing Then
                    ' Build occurrence path using InternalNames
                    Dim occPath As String = GetOccurrenceInternalPath(containingOcc)
                    Dim wfName As String = CStr(CallByName(obj, "Name", CallType.Get))
                    
                    ' Only use path-based format if we got a valid path
                    If occPath <> "" AndAlso wfName <> "" Then
                        Return "@" & occPath & "|" & wfName
                    ElseIf wfName <> "" Then
                        ' Fallback to just the name if path couldn't be built
                        Return wfName
                    End If
                End If
            Catch
            End Try
        End If
        
        ' Assembly-level or fallback - just return name
        Return GetObjectName(obj)
    End Function

    ''' <summary>
    ''' Build the occurrence path using InternalNames (stable across renames).
    ''' Returns path like "InternalName1/InternalName2" from root to the given occurrence.
    ''' </summary>
    Private Function GetOccurrenceInternalPath(occ As Object) As String
        Dim path As New System.Collections.Generic.List(Of String)
        Dim current As Object = occ
        
        Do While current IsNot Nothing
            Try
                ' Try direct cast to ComponentOccurrence first
                Dim compOcc As Inventor.ComponentOccurrence = Nothing
                Try
                    compOcc = CType(current, Inventor.ComponentOccurrence)
                Catch
                    ' Maybe it's a proxy - try to get the native object
                    Try
                        compOcc = CType(CallByName(current, "NativeObject", CallType.Get), Inventor.ComponentOccurrence)
                    Catch
                    End Try
                End Try
                
                If compOcc IsNot Nothing Then
                    path.Insert(0, compOcc.InternalName)
                    
                    ' Try to get parent occurrence
                    Try
                        current = compOcc.ParentOccurrence
                    Catch
                        current = Nothing
                    End Try
                Else
                    ' Fallback: try CallByName
                    Dim internalName As String = CStr(CallByName(current, "InternalName", CallType.Get))
                    If internalName <> "" Then
                        path.Insert(0, internalName)
                    End If
                    
                    Try
                        current = CallByName(current, "ParentOccurrence", CallType.Get)
                    Catch
                        current = Nothing
                    End Try
                End If
            Catch
                Exit Do
            End Try
        Loop
        
        Return String.Join("/", path.ToArray())
    End Function

    ' ============================================================================
    ' SECTION 5: Type Checking
    ' ============================================================================

    ''' <summary>
    ''' Check if object is a work feature (WorkPoint, WorkAxis, WorkPlane or their proxies).
    ''' </summary>
    Public Function IsWorkFeature(obj As Object) As Boolean
        If obj Is Nothing Then Return False
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(obj)
        
        Select Case typeName
            Case "WorkPoint", "WorkPointProxy", _
                 "WorkAxis", "WorkAxisProxy", _
                 "WorkPlane", "WorkPlaneProxy"
                Return True
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Check if object is a point-like geometry (WorkPoint, Vertex, etc).
    ''' </summary>
    Public Function IsPointGeometry(obj As Object) As Boolean
        Return GetPointGeometry(obj) IsNot Nothing
    End Function

    ''' <summary>
    ''' Check if object is a plane-like geometry (WorkPlane, planar Face).
    ''' </summary>
    Public Function IsPlaneGeometry(obj As Object) As Boolean
        Return GetPlaneGeometry(obj) IsNot Nothing
    End Function

    ''' <summary>
    ''' Check if object is a linear geometry (WorkAxis, linear Edge).
    ''' </summary>
    Public Function IsAxisGeometry(obj As Object) As Boolean
        Dim pt As Point = Nothing
        Dim dir As UnitVector = Nothing
        Return GetAxisProperties(obj, pt, dir)
    End Function

    ' ============================================================================
    ' SECTION 6: Value Formatting
    ' ============================================================================

    ''' <summary>
    ''' Format a dimension value in millimeters with 3 decimal places.
    ''' Returns a string like "123.000 mm" for use in iProperties.
    ''' </summary>
    Public Function FormatDimensionMm(valueMm As Double) As String
        Return valueMm.ToString("0.000") & " mm"
    End Function

    ''' <summary>
    ''' Format a dimension value in centimeters (internal units) to mm with 3 decimal places.
    ''' Converts from cm to mm and formats as "123.000 mm".
    ''' </summary>
    Public Function FormatDimensionCmToMm(valueCm As Double) As String
        Dim valueMm As Double = valueCm * 10.0
        Return valueMm.ToString("0.000") & " mm"
    End Function

End Module

