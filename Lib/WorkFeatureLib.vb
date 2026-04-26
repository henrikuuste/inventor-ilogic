' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' WorkFeatureLib - Work Feature Creation and Management
' 
' Functions for creating, finding, and managing work planes, axes, and points
' in assemblies with support for associativity.
'
' Depends on: UtilsLib.vb, GeoLib.vb
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/GeoLib.vb"
'   AddVbFile "Lib/WorkFeatureLib.vb"
'
' ============================================================================

Option Strict Off
Imports Inventor

Public Module WorkFeatureLib

    ' ============================================================================
    ' CONSTANTS
    ' ============================================================================
    
    Public Const GEOM_TYPE_UNKNOWN As Integer = 0
    Public Const GEOM_TYPE_PLANE As Integer = 1
    Public Const GEOM_TYPE_POINT As Integer = 2
    Public Const GEOM_TYPE_EDGE As Integer = 3
    Public Const GEOM_TYPE_VERTEX As Integer = 4
    Public Const GEOM_TYPE_FACE_NONPLANAR As Integer = 5

    ' ============================================================================
    ' SECTION 1: Find Work Features by Name
    ' ============================================================================
    
    ''' <summary>
    ''' Find a work plane by name in the assembly.
    ''' </summary>
    Public Function FindWorkPlaneByName(asmDef As AssemblyComponentDefinition, _
                                         name As String) As WorkPlane
        If asmDef Is Nothing OrElse String.IsNullOrEmpty(name) Then Return Nothing
        
        Try
            For Each wp As WorkPlane In asmDef.WorkPlanes
                If wp.Name = name Then Return wp
            Next
        Catch
        End Try
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Find a work axis by name in the assembly.
    ''' </summary>
    Public Function FindWorkAxisByName(asmDef As AssemblyComponentDefinition, _
                                        name As String) As WorkAxis
        If asmDef Is Nothing OrElse String.IsNullOrEmpty(name) Then Return Nothing
        
        Try
            For Each wa As WorkAxis In asmDef.WorkAxes
                If wa.Name = name Then Return wa
            Next
        Catch
        End Try
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Find a work point by name in the assembly.
    ''' </summary>
    Public Function FindWorkPointByName(asmDef As AssemblyComponentDefinition, _
                                         name As String) As WorkPoint
        If asmDef Is Nothing OrElse String.IsNullOrEmpty(name) Then Return Nothing
        
        Try
            For Each wp As WorkPoint In asmDef.WorkPoints
                If wp.Name = name Then Return wp
            Next
        Catch
        End Try
        Return Nothing
    End Function

    ' ============================================================================
    ' SECTION 2: Create Associative Work Planes
    ' ============================================================================
    
    ''' <summary>
    ''' Create an associative work plane from geometry (face, plane, work plane).
    ''' If geometry is already an assembly-level work plane, returns it.
    ''' For faces/proxies, creates a work plane using AddByPlaneAndOffset(geometry, 0).
    ''' </summary>
    Public Function CreateAssociativeWorkPlane(app As Inventor.Application, _
                                                asmDoc As AssemblyDocument, _
                                                geometry As Object, _
                                                planeName As String) As WorkPlane
        If geometry Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(geometry)
        
        ' If it's already an assembly-level work plane, just rename and return
        If typeName = "WorkPlane" Then
            Dim wp As WorkPlane = CType(geometry, WorkPlane)
            If wp.Parent Is asmDef Then
                Try
                    wp.Name = planeName
                Catch
                End Try
                Return wp
            End If
        End If
        
        ' Check if we already have a work plane with this name
        Dim existingPlane As WorkPlane = FindWorkPlaneByName(asmDef, planeName)
        If existingPlane IsNot Nothing Then
            Return existingPlane
        End If
        
        ' Try to create associative work plane using AddByPlaneAndOffset
        ' This works in assemblies for faces, work plane proxies, etc.
        Try
            Dim newPlane As WorkPlane = asmDef.WorkPlanes.AddByPlaneAndOffset(geometry, 0)
            newPlane.Name = planeName
            newPlane.Visible = False
            Return newPlane
        Catch
            ' Fallback to fixed work plane with flush constraint
            Return CreateFixedWorkPlaneWithConstraint(app, asmDoc, geometry, planeName)
        End Try
    End Function
    
    ''' <summary>
    ''' Fallback: create a fixed work plane and use Flush constraint for associativity.
    ''' </summary>
    Private Function CreateFixedWorkPlaneWithConstraint(app As Inventor.Application, _
                                                         asmDoc As AssemblyDocument, _
                                                         geometry As Object, _
                                                         planeName As String) As WorkPlane
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim planeGeom As Plane = GeoLib.GetPlaneGeometry(geometry)
        If planeGeom Is Nothing Then Return Nothing
        
        Try
            Dim tg As TransientGeometry = app.TransientGeometry
            Dim origin As Point = planeGeom.RootPoint
            Dim normal As UnitVector = planeGeom.Normal
            
            ' Create X and Y axes perpendicular to normal
            Dim refVec As Vector
            If Math.Abs(normal.Z) < 0.9 Then
                refVec = tg.CreateVector(0, 0, 1)
            Else
                refVec = tg.CreateVector(1, 0, 0)
            End If
            
            Dim xVec As Vector = refVec.CrossProduct(tg.CreateVector(normal.X, normal.Y, normal.Z))
            xVec.Normalize()
            Dim xAxis As UnitVector = tg.CreateUnitVector(xVec.X, xVec.Y, xVec.Z)
            
            Dim yVec As Vector = tg.CreateVector(normal.X, normal.Y, normal.Z).CrossProduct(xVec)
            yVec.Normalize()
            Dim yAxis As UnitVector = tg.CreateUnitVector(yVec.X, yVec.Y, yVec.Z)
            
            Dim newPlane As WorkPlane = asmDef.WorkPlanes.AddFixed(origin, xAxis, yAxis)
            newPlane.Name = planeName
            newPlane.Visible = False
            
            ' Add Flush constraint to geometry with 0 offset to make it associative
            Try
                asmDef.Constraints.AddFlushConstraint(geometry, newPlane, 0)
            Catch
                ' Constraint may fail, but plane is still at correct position
            End Try
            
            Return newPlane
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Get or create an associative work plane. Returns existing if found by name.
    ''' </summary>
    Public Function GetOrCreateWorkPlane(app As Inventor.Application, _
                                          asmDoc As AssemblyDocument, _
                                          geometry As Object, _
                                          planeName As String) As WorkPlane
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Check if already exists
        Dim existing As WorkPlane = FindWorkPlaneByName(asmDef, planeName)
        If existing IsNot Nothing Then
            Return existing
        End If
        
        ' Create new
        Return CreateAssociativeWorkPlane(app, asmDoc, geometry, planeName)
    End Function

    ' ============================================================================
    ' SECTION 3: Create Associative Work Axis
    ' ============================================================================
    
    ''' <summary>
    ''' Create an associative work axis perpendicular to the start plane.
    ''' The axis direction is the start plane's normal, oriented to point toward end plane.
    ''' </summary>
    Public Function CreateAssociativeWorkAxis(app As Inventor.Application, _
                                               asmDoc As AssemblyDocument, _
                                               startPlane As WorkPlane, _
                                               endPlane As WorkPlane, _
                                               axisName As String) As WorkAxis
        If startPlane Is Nothing OrElse endPlane Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Check if axis already exists
        Dim existingAxis As WorkAxis = FindWorkAxisByName(asmDef, axisName)
        If existingAxis IsNot Nothing Then
            Return existingAxis
        End If
        
        ' Get plane geometry
        Dim plane1 As Plane = GeoLib.GetPlaneGeometry(startPlane)
        Dim plane2 As Plane = GeoLib.GetPlaneGeometry(endPlane)
        If plane1 Is Nothing OrElse plane2 Is Nothing Then Return Nothing
        
        Dim tg As TransientGeometry = app.TransientGeometry
        
        ' Direction is the start plane's normal
        Dim direction As UnitVector = plane1.Normal
        Dim startPt As Point = plane1.RootPoint
        Dim endPt As Point = plane2.RootPoint
        
        ' Check if direction points from start to end
        Dim toEnd As Double = (endPt.X - startPt.X) * direction.X + _
                              (endPt.Y - startPt.Y) * direction.Y + _
                              (endPt.Z - startPt.Z) * direction.Z
        
        If toEnd < 0 Then
            ' Flip direction to point toward end
            direction = tg.CreateUnitVector(-direction.X, -direction.Y, -direction.Z)
        End If
        
        ' Try to create using AddByNormalToSurface (more associative)
        Try
            Dim newAxis As WorkAxis = asmDef.WorkAxes.AddByNormalToSurface(startPlane, startPt)
            newAxis.Name = axisName
            newAxis.Visible = False
            
            ' Verify direction and return
            Return newAxis
        Catch
            ' Fallback to AddFixed
        End Try
        
        ' Create fixed axis at start plane origin with normal direction
        Try
            Dim newAxis As WorkAxis = asmDef.WorkAxes.AddFixed(startPt, direction)
            newAxis.Name = axisName
            newAxis.Visible = False
            Return newAxis
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Get or create an associative work axis. Returns existing if found by name.
    ''' </summary>
    Public Function GetOrCreateWorkAxis(app As Inventor.Application, _
                                         asmDoc As AssemblyDocument, _
                                         startPlane As WorkPlane, _
                                         endPlane As WorkPlane, _
                                         axisName As String) As WorkAxis
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Check if already exists
        Dim existing As WorkAxis = FindWorkAxisByName(asmDef, axisName)
        If existing IsNot Nothing Then
            Return existing
        End If
        
        ' Create new
        Return CreateAssociativeWorkAxis(app, asmDoc, startPlane, endPlane, axisName)
    End Function

    ' ============================================================================
    ' SECTION 4: Work Plane Proxies
    ' ============================================================================
    
    ''' <summary>
    ''' Create a proxy for a part's work plane in assembly context.
    ''' This allows constraining assembly features to the part's internal planes.
    ''' </summary>
    Public Function CreateWorkPlaneProxy(occ As ComponentOccurrence, _
                                          planeIndex As Integer) As Object
        If occ Is Nothing Then Return Nothing
        
        Try
            Dim compDef As ComponentDefinition = occ.Definition
            Dim wp As WorkPlane = compDef.WorkPlanes.Item(planeIndex)
            
            ' Create proxy for assembly context
            Dim proxyResult As Object = Nothing
            occ.CreateGeometryProxy(wp, proxyResult)
            Return proxyResult
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Create a proxy for the principal plane of a part (most perpendicular to axis).
    ''' </summary>
    Public Function CreatePrincipalPlaneProxy(occ As ComponentOccurrence, _
                                               axisDirection As UnitVector) As Object
        If occ Is Nothing OrElse axisDirection Is Nothing Then Return Nothing
        
        Dim planeIndex As Integer = GeoLib.GetPrincipalPlaneIndex(occ, axisDirection)
        If planeIndex = 0 Then Return Nothing
        
        Return CreateWorkPlaneProxy(occ, planeIndex)
    End Function

    ' ============================================================================
    ' SECTION 5: Delete Work Features
    ' ============================================================================
    
    ''' <summary>
    ''' Delete a work plane by name if it exists.
    ''' </summary>
    Public Function DeleteWorkPlane(asmDef As AssemblyComponentDefinition, _
                                     name As String) As Boolean
        Dim wp As WorkPlane = FindWorkPlaneByName(asmDef, name)
        If wp IsNot Nothing Then
            Try
                wp.Delete()
                Return True
            Catch
            End Try
        End If
        Return False
    End Function
    
    ''' <summary>
    ''' Delete a work axis by name if it exists.
    ''' </summary>
    Public Function DeleteWorkAxis(asmDef As AssemblyComponentDefinition, _
                                    name As String) As Boolean
        Dim wa As WorkAxis = FindWorkAxisByName(asmDef, name)
        If wa IsNot Nothing Then
            Try
                wa.Delete()
                Return True
            Catch
            End Try
        End If
        Return False
    End Function

    ' ============================================================================
    ' SECTION 6: Geometry Type Detection
    ' ============================================================================
    
    ''' <summary>
    ''' Detect the type of geometry for work feature creation.
    ''' Returns one of the GEOM_TYPE_* constants.
    ''' </summary>
    Public Function GetGeometryType(geometry As Object) As Integer
        If geometry Is Nothing Then Return GEOM_TYPE_UNKNOWN
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(geometry)
        
        Select Case typeName
            Case "WorkPlane", "WorkPlaneProxy", "FaceProxy", "Face"
                ' Check if face is planar
                If typeName = "FaceProxy" OrElse typeName = "Face" Then
                    Try
                        Dim face As Face = Nothing
                        If typeName = "FaceProxy" Then
                            face = CType(geometry, FaceProxy).NativeObject
                        Else
                            face = CType(geometry, Face)
                        End If
                        
                        If face.SurfaceType = SurfaceTypeEnum.kPlaneSurface Then
                            Return GEOM_TYPE_PLANE
                        Else
                            Return GEOM_TYPE_FACE_NONPLANAR
                        End If
                    Catch
                        Return GEOM_TYPE_PLANE ' Assume planar if can't determine
                    End Try
                End If
                Return GEOM_TYPE_PLANE
                
            Case "WorkPoint", "WorkPointProxy"
                Return GEOM_TYPE_POINT
                
            Case "Vertex", "VertexProxy"
                Return GEOM_TYPE_VERTEX
                
            Case "Edge", "EdgeProxy"
                Return GEOM_TYPE_EDGE
                
            Case "Point"
                Return GEOM_TYPE_POINT
                
            Case Else
                Return GEOM_TYPE_UNKNOWN
        End Select
    End Function
    
    ''' <summary>
    ''' Check if geometry is a planar type (plane or planar face).
    ''' </summary>
    Public Function IsPlanarGeometry(geometry As Object) As Boolean
        Return GetGeometryType(geometry) = GEOM_TYPE_PLANE
    End Function

    ' ============================================================================
    ' SECTION 7: Extended Work Point Creation
    ' ============================================================================
    
    ''' <summary>
    ''' Create an associative work point from various geometry types.
    ''' Supports: Vertex, Edge (midpoint), Face (centroid), WorkPoint.
    ''' All methods maintain associativity to the source geometry.
    ''' </summary>
    Public Function CreateAssociativeWorkPoint(app As Inventor.Application, _
                                                asmDoc As AssemblyDocument, _
                                                geometry As Object, _
                                                pointName As String) As WorkPoint
        If geometry Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(geometry)
        
        ' Check if work point with this name already exists
        Dim existing As WorkPoint = FindWorkPointByName(asmDef, pointName)
        If existing IsNot Nothing Then
            Return existing
        End If
        
        ' If it's already an assembly-level work point, return it
        If typeName = "WorkPoint" Then
            Dim wp As WorkPoint = CType(geometry, WorkPoint)
            If wp.Parent Is asmDef Then
                Try
                    wp.Name = pointName
                Catch
                End Try
                Return wp
            End If
        End If
        
        Try
            Dim newPoint As WorkPoint = Nothing
            
            Select Case typeName
                Case "WorkPoint", "WorkPointProxy"
                    ' Create offset work point at same location (associative)
                    newPoint = asmDef.WorkPoints.AddByPoint(geometry)
                    
                Case "Vertex", "VertexProxy"
                    ' Create work point at vertex (associative)
                    newPoint = asmDef.WorkPoints.AddByPoint(geometry)
                    
                Case "Edge", "EdgeProxy"
                    ' Create work point at edge midpoint (associative)
                    newPoint = asmDef.WorkPoints.AddAtMidPoint(geometry)
                    
                Case "FaceProxy", "Face"
                    ' Create work point at face centroid (associative)
                    newPoint = asmDef.WorkPoints.AddAtCentroid(geometry)
                    
                Case "Point"
                    ' Transient point - create fixed work point (NOT associative)
                    Dim pt As Point = CType(geometry, Point)
                    newPoint = asmDef.WorkPoints.AddFixed(pt)
                    
            End Select
            
            If newPoint IsNot Nothing Then
                newPoint.Name = pointName
                newPoint.Visible = False
                Return newPoint
            End If
        Catch
        End Try
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Get a Point (coordinates) from various geometry types.
    ''' Used for fallback fixed work feature creation.
    ''' </summary>
    Public Function GetPointFromGeometry(app As Inventor.Application, _
                                          geometry As Object) As Point
        If geometry Is Nothing Then Return Nothing
        
        Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(geometry)
        Dim tg As TransientGeometry = app.TransientGeometry
        
        Try
            Select Case typeName
                Case "WorkPoint"
                    Return CType(geometry, WorkPoint).Point
                    
                Case "WorkPointProxy"
                    Return CType(geometry, WorkPointProxy).Point
                    
                Case "Point"
                    Return CType(geometry, Point)
                    
                Case "Vertex"
                    Return CType(geometry, Vertex).Point
                    
                Case "VertexProxy"
                    Return CType(geometry, VertexProxy).Point
                    
                Case "Edge", "EdgeProxy"
                    ' Get midpoint of edge
                    Dim edge As Object = geometry
                    Dim evaluator As Object = Nothing
                    If typeName = "EdgeProxy" Then
                        evaluator = CType(geometry, EdgeProxy).Geometry.Evaluator
                    Else
                        evaluator = CType(geometry, Edge).Geometry.Evaluator
                    End If
                    
                    Dim minParam As Double = 0
                    Dim maxParam As Double = 0
                    evaluator.GetParamExtents(minParam, maxParam)
                    
                    Dim midParam As Double = (minParam + maxParam) / 2
                    Dim pts(2) As Double
                    evaluator.GetPointAtParam(midParam, pts)
                    Return tg.CreatePoint(pts(0), pts(1), pts(2))
                    
                Case "FaceProxy", "Face"
                    ' Get centroid of face (approximate with bounding box center)
                    Dim box As Box = Nothing
                    If typeName = "FaceProxy" Then
                        box = CType(geometry, FaceProxy).RangeBox
                    Else
                        box = CType(geometry, Face).RangeBox
                    End If
                    
                    Return tg.CreatePoint( _
                        (box.MinPoint.X + box.MaxPoint.X) / 2, _
                        (box.MinPoint.Y + box.MaxPoint.Y) / 2, _
                        (box.MinPoint.Z + box.MaxPoint.Z) / 2)
                    
            End Select
        Catch
        End Try
        
        Return Nothing
    End Function

    ' ============================================================================
    ' SECTION 8: Extended Work Plane Creation (with axis direction)
    ' ============================================================================
    
    ''' <summary>
    ''' Create an associative work plane from any geometry type.
    ''' For point-based geometry (vertices, points, non-planar faces),
    ''' creates a plane perpendicular to the given axis direction.
    ''' 
    ''' Maintains associativity where possible:
    ''' - Planar faces/planes: AddByPlaneAndOffset (associative)
    ''' - Vertices: AddByPoint → AddByPointNormalToLine (associative)
    ''' - Edges: AddAtMidPoint → AddByPointNormalToLine (associative)
    ''' - Non-planar faces: AddAtCentroid → AddByPointNormalToLine (associative)
    ''' - WorkPoints: AddByPointNormalToLine (associative)
    ''' </summary>
    Public Function CreateAssociativeWorkPlaneEx(app As Inventor.Application, _
                                                  asmDoc As AssemblyDocument, _
                                                  geometry As Object, _
                                                  axisOrDirection As Object, _
                                                  planeName As String) As WorkPlane
        If geometry Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim geomType As Integer = GetGeometryType(geometry)
        
        ' Check if already exists
        Dim existingPlane As WorkPlane = FindWorkPlaneByName(asmDef, planeName)
        If existingPlane IsNot Nothing Then
            Return existingPlane
        End If
        
        ' Handle planar geometry - use existing function
        If geomType = GEOM_TYPE_PLANE Then
            Return CreateAssociativeWorkPlane(app, asmDoc, geometry, planeName)
        End If
        
        ' For point-based geometry, we need an axis or direction
        Dim workAxis As WorkAxis = Nothing
        Dim axisDirection As UnitVector = Nothing
        
        ' Determine axis from axisOrDirection parameter
        If axisOrDirection IsNot Nothing Then
            Dim axisTypeName As String = Microsoft.VisualBasic.Information.TypeName(axisOrDirection)
            
            If axisTypeName = "WorkAxis" OrElse axisTypeName = "WorkAxisProxy" Then
                workAxis = CType(axisOrDirection, WorkAxis)
                axisDirection = GeoLib.GetAxisDirection(axisOrDirection)
            ElseIf axisTypeName = "UnitVector" Then
                axisDirection = CType(axisOrDirection, UnitVector)
            ElseIf axisTypeName = "Edge" OrElse axisTypeName = "EdgeProxy" Then
                ' Linear edge as axis
                axisDirection = GeoLib.GetAxisDirection(axisOrDirection)
            End If
        End If
        
        If axisDirection Is Nothing Then
            ' Can't create plane without axis direction for non-planar geometry
            Return Nothing
        End If
        
        ' Create work point from geometry if needed
        Dim workPoint As WorkPoint = Nothing
        Dim pointName As String = planeName & "_Punkt"
        
        Select Case geomType
            Case GEOM_TYPE_POINT
                ' Already a work point or need to create from it
                Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(geometry)
                If typeName = "WorkPoint" Then
                    workPoint = CType(geometry, WorkPoint)
                Else
                    workPoint = CreateAssociativeWorkPoint(app, asmDoc, geometry, pointName)
                End If
                
            Case GEOM_TYPE_VERTEX
                workPoint = CreateAssociativeWorkPoint(app, asmDoc, geometry, pointName)
                
            Case GEOM_TYPE_EDGE
                workPoint = CreateAssociativeWorkPoint(app, asmDoc, geometry, pointName)
                
            Case GEOM_TYPE_FACE_NONPLANAR
                workPoint = CreateAssociativeWorkPoint(app, asmDoc, geometry, pointName)
        End Select
        
        If workPoint Is Nothing Then
            Return Nothing
        End If
        
        ' Create plane through work point, perpendicular to axis
        Try
            Dim newPlane As WorkPlane = Nothing
            
            ' Try AddByPointNormalToLine if we have a work axis
            If workAxis IsNot Nothing Then
                Try
                    newPlane = asmDef.WorkPlanes.AddByPointNormalToLine(workPoint, workAxis)
                Catch
                End Try
            End If
            
            ' Fallback: create fixed plane at work point with normal = axis direction
            If newPlane Is Nothing Then
                Dim tg As TransientGeometry = app.TransientGeometry
                Dim origin As Point = workPoint.Point
                
                ' Create perpendicular X and Y axes
                Dim refVec As Vector
                If Math.Abs(axisDirection.Z) < 0.9 Then
                    refVec = tg.CreateVector(0, 0, 1)
                Else
                    refVec = tg.CreateVector(1, 0, 0)
                End If
                
                Dim normalVec As Vector = tg.CreateVector(axisDirection.X, axisDirection.Y, axisDirection.Z)
                Dim xVec As Vector = refVec.CrossProduct(normalVec)
                xVec.Normalize()
                Dim xAxis As UnitVector = tg.CreateUnitVector(xVec.X, xVec.Y, xVec.Z)
                
                Dim yVec As Vector = normalVec.CrossProduct(xVec)
                yVec.Normalize()
                Dim yAxis As UnitVector = tg.CreateUnitVector(yVec.X, yVec.Y, yVec.Z)
                
                newPlane = asmDef.WorkPlanes.AddFixed(origin, xAxis, yAxis)
            End If
            
            If newPlane IsNot Nothing Then
                newPlane.Name = planeName
                newPlane.Visible = False
                Return newPlane
            End If
        Catch
        End Try
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Get or create an associative work plane with axis direction support.
    ''' </summary>
    Public Function GetOrCreateWorkPlaneEx(app As Inventor.Application, _
                                            asmDoc As AssemblyDocument, _
                                            geometry As Object, _
                                            axisOrDirection As Object, _
                                            planeName As String) As WorkPlane
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Check if already exists
        Dim existing As WorkPlane = FindWorkPlaneByName(asmDef, planeName)
        If existing IsNot Nothing Then
            Return existing
        End If
        
        ' Create new
        Return CreateAssociativeWorkPlaneEx(app, asmDoc, geometry, axisOrDirection, planeName)
    End Function

    ' ============================================================================
    ' SECTION 9: Extended Work Axis Creation
    ' ============================================================================
    
    ''' <summary>
    ''' Create a work axis from two points (vertices, work points, or points).
    ''' Maintains associativity when using vertices or work points.
    ''' </summary>
    Public Function CreateWorkAxisFromTwoPoints(app As Inventor.Application, _
                                                 asmDoc As AssemblyDocument, _
                                                 point1 As Object, _
                                                 point2 As Object, _
                                                 axisName As String) As WorkAxis
        If point1 Is Nothing OrElse point2 Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Check if already exists
        Dim existing As WorkAxis = FindWorkAxisByName(asmDef, axisName)
        If existing IsNot Nothing Then
            Return existing
        End If
        
        ' Create work points if needed
        Dim wp1 As WorkPoint = Nothing
        Dim wp2 As WorkPoint = Nothing
        
        Dim type1 As String = Microsoft.VisualBasic.Information.TypeName(point1)
        Dim type2 As String = Microsoft.VisualBasic.Information.TypeName(point2)
        
        ' Handle first point
        If type1 = "WorkPoint" Then
            wp1 = CType(point1, WorkPoint)
        ElseIf type1 = "Vertex" OrElse type1 = "VertexProxy" Then
            wp1 = CreateAssociativeWorkPoint(app, asmDoc, point1, axisName & "_P1")
        Else
            wp1 = CreateAssociativeWorkPoint(app, asmDoc, point1, axisName & "_P1")
        End If
        
        ' Handle second point
        If type2 = "WorkPoint" Then
            wp2 = CType(point2, WorkPoint)
        ElseIf type2 = "Vertex" OrElse type2 = "VertexProxy" Then
            wp2 = CreateAssociativeWorkPoint(app, asmDoc, point2, axisName & "_P2")
        Else
            wp2 = CreateAssociativeWorkPoint(app, asmDoc, point2, axisName & "_P2")
        End If
        
        If wp1 Is Nothing OrElse wp2 Is Nothing Then
            ' Fallback to fixed axis
            Dim pt1 As Point = GetPointFromGeometry(app, point1)
            Dim pt2 As Point = GetPointFromGeometry(app, point2)
            If pt1 IsNot Nothing AndAlso pt2 IsNot Nothing Then
                Return CreateFixedAxisBetweenPoints(app, asmDoc, pt1, pt2, axisName)
            End If
            Return Nothing
        End If
        
        ' Create axis between work points (associative)
        Try
            Dim newAxis As WorkAxis = asmDef.WorkAxes.AddByTwoPoints(wp1, wp2)
            newAxis.Name = axisName
            newAxis.Visible = False
            Return newAxis
        Catch
            ' Fallback
            Dim pt1 As Point = wp1.Point
            Dim pt2 As Point = wp2.Point
            Return CreateFixedAxisBetweenPoints(app, asmDoc, pt1, pt2, axisName)
        End Try
    End Function
    
    ''' <summary>
    ''' Create a fixed (non-associative) work axis between two points.
    ''' </summary>
    Private Function CreateFixedAxisBetweenPoints(app As Inventor.Application, _
                                                   asmDoc As AssemblyDocument, _
                                                   pt1 As Point, _
                                                   pt2 As Point, _
                                                   axisName As String) As WorkAxis
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim tg As TransientGeometry = app.TransientGeometry
        
        ' Calculate direction from pt1 to pt2
        Dim dx As Double = pt2.X - pt1.X
        Dim dy As Double = pt2.Y - pt1.Y
        Dim dz As Double = pt2.Z - pt1.Z
        Dim length As Double = Math.Sqrt(dx * dx + dy * dy + dz * dz)
        
        If length < 0.0001 Then Return Nothing
        
        Dim direction As UnitVector = tg.CreateUnitVector(dx / length, dy / length, dz / length)
        
        Try
            Dim newAxis As WorkAxis = asmDef.WorkAxes.AddFixed(pt1, direction)
            newAxis.Name = axisName
            newAxis.Visible = False
            Return newAxis
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Create a work axis from a planar geometry's normal direction.
    ''' </summary>
    Public Function CreateWorkAxisFromPlaneNormal(app As Inventor.Application, _
                                                   asmDoc As AssemblyDocument, _
                                                   planeGeometry As Object, _
                                                   axisName As String) As WorkAxis
        If planeGeometry Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Check if already exists
        Dim existing As WorkAxis = FindWorkAxisByName(asmDef, axisName)
        If existing IsNot Nothing Then
            Return existing
        End If
        
        ' Get plane geometry first
        Dim planeGeom As Plane = GeoLib.GetPlaneGeometry(planeGeometry)
        If planeGeom Is Nothing Then Return Nothing
        
        ' Try to create using AddByNormalToSurface (associative for faces/work planes)
        Try
            Dim newAxis As WorkAxis = asmDef.WorkAxes.AddByNormalToSurface(planeGeometry, planeGeom.RootPoint)
            newAxis.Name = axisName
            newAxis.Visible = False
            Return newAxis
        Catch
        End Try
        
        ' Fallback to fixed axis
        If planeGeom IsNot Nothing Then
            Try
                Dim newAxis As WorkAxis = asmDef.WorkAxes.AddFixed(planeGeom.RootPoint, planeGeom.Normal)
                newAxis.Name = axisName
                newAxis.Visible = False
                Return newAxis
            Catch
            End Try
        End If
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Determine axis direction from boundary selections.
    ''' Logic:
    ''' 1. If explicit axis provided → use it
    ''' 2. If two points/vertices → axis = line between them
    ''' 3. If point/vertex + plane/face → axis = plane normal
    ''' 4. If two parallel planes → axis = plane normal (existing behavior)
    ''' </summary>
    Public Function DetermineAxisFromBoundaries(app As Inventor.Application, _
                                                 asmDoc As AssemblyDocument, _
                                                 startGeometry As Object, _
                                                 endGeometry As Object, _
                                                 explicitAxis As Object, _
                                                 axisName As String) As WorkAxis
        ' If explicit axis provided, use it
        If explicitAxis IsNot Nothing Then
            Dim typeName As String = Microsoft.VisualBasic.Information.TypeName(explicitAxis)
            If typeName = "WorkAxis" Then
                Return CType(explicitAxis, WorkAxis)
            ElseIf typeName = "Edge" OrElse typeName = "EdgeProxy" Then
                ' Create axis from linear edge
                Return CreateWorkAxisFromLinearEdge(app, asmDoc, explicitAxis, axisName)
            End If
        End If
        
        Dim startType As Integer = GetGeometryType(startGeometry)
        Dim endType As Integer = GetGeometryType(endGeometry)
        
        ' Both are points/vertices → axis between them
        If (startType = GEOM_TYPE_POINT OrElse startType = GEOM_TYPE_VERTEX) AndAlso _
           (endType = GEOM_TYPE_POINT OrElse endType = GEOM_TYPE_VERTEX) Then
            Return CreateWorkAxisFromTwoPoints(app, asmDoc, startGeometry, endGeometry, axisName)
        End If
        
        ' One is plane, one is point → use plane normal
        If startType = GEOM_TYPE_PLANE AndAlso _
           (endType = GEOM_TYPE_POINT OrElse endType = GEOM_TYPE_VERTEX OrElse endType = GEOM_TYPE_FACE_NONPLANAR) Then
            Return CreateWorkAxisFromPlaneNormal(app, asmDoc, startGeometry, axisName)
        End If
        
        If endType = GEOM_TYPE_PLANE AndAlso _
           (startType = GEOM_TYPE_POINT OrElse startType = GEOM_TYPE_VERTEX OrElse startType = GEOM_TYPE_FACE_NONPLANAR) Then
            Return CreateWorkAxisFromPlaneNormal(app, asmDoc, endGeometry, axisName)
        End If
        
        ' Both are planes - return Nothing, let caller handle with existing CreateAssociativeWorkAxis
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Create work axis from a linear edge.
    ''' </summary>
    Public Function CreateWorkAxisFromLinearEdge(app As Inventor.Application, _
                                                  asmDoc As AssemblyDocument, _
                                                  edge As Object, _
                                                  axisName As String) As WorkAxis
        If edge Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Check if already exists
        Dim existing As WorkAxis = FindWorkAxisByName(asmDef, axisName)
        If existing IsNot Nothing Then
            Return existing
        End If
        
        ' Try to create axis through edge (associative)
        Try
            Dim newAxis As WorkAxis = asmDef.WorkAxes.AddByLine(edge)
            newAxis.Name = axisName
            newAxis.Visible = False
            Return newAxis
        Catch
        End Try
        
        ' Fallback: get edge direction and create fixed axis
        Dim edgeDir As UnitVector = GeoLib.GetAxisDirection(edge)
        Dim edgePt As Point = GetPointFromGeometry(app, edge)
        
        If edgeDir IsNot Nothing AndAlso edgePt IsNot Nothing Then
            Try
                Dim newAxis As WorkAxis = asmDef.WorkAxes.AddFixed(edgePt, edgeDir)
                newAxis.Name = axisName
                newAxis.Visible = False
                Return newAxis
            Catch
            End Try
        End If
        
        Return Nothing
    End Function

End Module
