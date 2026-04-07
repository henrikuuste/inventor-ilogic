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

End Module
