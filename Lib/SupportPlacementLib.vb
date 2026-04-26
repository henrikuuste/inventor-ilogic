' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' SupportPlacementLib - Library for Parametric Birch Support Placement
' 
' Support-specific functions for placing Kask24 supports in assemblies:
' - Part file management (template, creation, reuse)
' - Placement matrix creation
' - Occurrence reference storage and retrieval
' - iProperty updates for BOM compatibility
'
' Depends on: UtilsLib.vb (geometry and utility functions)
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/SupportPlacementLib.vb"
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

Imports Inventor

Public Module SupportPlacementLib

    ' ============================================================================
    ' Constants
    ' ============================================================================

    ' Available beam widths (X axis dimension in mm)
    Public ReadOnly SUPPORT_WIDTHS As Integer() = {24, 35, 45, 57, 70}
    
    ' Fixed thickness (Y axis dimension in mm)
    Public Const SUPPORT_THICKNESS_MM As Double = 24.0
    
    ' Template filename
    Public Const TEMPLATE_FILENAME As String = "Kask24_Template.ipt"
    
    ' Attribute set name for occurrence storage
    Private Const ATTR_SET_NAME As String = "SupportPlacement"

    ' ============================================================================
    ' SECTION 1: Template Work Points
    ' ============================================================================

    ''' <summary>
    ''' Get list of work point names from the template.
    ''' Opens template invisibly, reads work points, closes without saving.
    ''' </summary>
    Public Function GetTemplateWorkPoints(app As Inventor.Application, templatePath As String) As String()
        Dim workPointNames As New System.Collections.Generic.List(Of String)
        
        Dim partDoc As PartDocument = Nothing
        Try
            partDoc = CType(app.Documents.Open(templatePath, False), PartDocument)
            
            Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
            For Each wp As WorkPoint In partDef.WorkPoints
                Try
                    Dim wpName As String = wp.Name
                    If wpName <> "" AndAlso Not wpName.StartsWith("Work Point") Then
                        workPointNames.Add(wpName)
                    End If
                Catch
                End Try
            Next
        Catch
        Finally
            If partDoc IsNot Nothing Then
                Try
                    partDoc.Close(True)
                Catch
                End Try
            End If
        End Try
        
        If workPointNames.Count = 0 Then
            workPointNames.Add("Origin")
        End If
        
        Return workPointNames.ToArray()
    End Function

    ' ============================================================================
    ' SECTION 2: Part File Management
    ' ============================================================================

    ''' <summary>
    ''' Find the template file path, searching common locations.
    ''' </summary>
    Public Function FindTemplatePath(asmDoc As AssemblyDocument) As String
        Dim asmFolder As String = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
        
        Dim searchPaths As String() = { _
            System.IO.Path.Combine(asmFolder, TEMPLATE_FILENAME), _
            System.IO.Path.Combine(asmFolder, "Templates", TEMPLATE_FILENAME), _
            System.IO.Path.Combine(System.IO.Path.GetDirectoryName(asmFolder), TEMPLATE_FILENAME), _
            System.IO.Path.Combine(System.IO.Path.GetDirectoryName(asmFolder), "Templates", TEMPLATE_FILENAME) _
        }
        
        For Each path As String In searchPaths
            If System.IO.File.Exists(path) Then
                Return path
            End If
        Next
        
        Return ""
    End Function

    ''' <summary>
    ''' Get the Kask subfolder path, creating it if it doesn't exist.
    ''' </summary>
    Private Function GetKaskFolder(asmDoc As AssemblyDocument) As String
        Dim asmFolder As String = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
        Dim kaskFolder As String = System.IO.Path.Combine(asmFolder, "Kask")
        
        If Not System.IO.Directory.Exists(kaskFolder) Then
            System.IO.Directory.CreateDirectory(kaskFolder)
        End If
        
        Return kaskFolder
    End Function

    ''' <summary>
    ''' Find an existing support file with matching width and length.
    ''' Returns empty string if not found.
    ''' </summary>
    Public Function FindExistingSupportFile(app As Inventor.Application, asmDoc As AssemblyDocument, widthMm As Integer, lengthMm As Integer) As String
        Dim kaskFolder As String = GetKaskFolder(asmDoc)
        Dim pattern As String = "Kask24x" & widthMm.ToString() & "-*.ipt"
        
        Try
            Dim files As String() = System.IO.Directory.GetFiles(kaskFolder, pattern)
            
            For Each filePath As String In files
                Try
                    Dim partDoc As PartDocument = CType(app.Documents.Open(filePath, False), PartDocument)
                    Try
                        Dim params As Parameters = partDoc.ComponentDefinition.Parameters
                        Dim fileWidth As Double = params.Item("Width").Value * 10  ' cm to mm
                        Dim fileLength As Double = params.Item("Length").Value * 10  ' cm to mm
                        
                        If Math.Abs(fileWidth - widthMm) < 0.1 AndAlso Math.Abs(fileLength - lengthMm) < 0.1 Then
                            partDoc.Close(True)
                            Return filePath
                        End If
                    Finally
                        partDoc.Close(True)
                    End Try
                Catch
                End Try
            Next
        Catch
        End Try
        
        Return ""
    End Function

    ''' <summary>
    ''' Generate a unique filename for a new support in the Kask subfolder.
    ''' If customName is provided, uses that as the base name instead of auto-generated.
    ''' </summary>
    Public Function GenerateSupportFileName(asmDoc As AssemblyDocument, widthMm As Integer, _
                                             Optional customName As String = "") As String
        Dim kaskFolder As String = GetKaskFolder(asmDoc)
        Dim baseName As String
        
        If customName <> "" Then
            ' Use custom name - sanitize for filename
            baseName = SanitizeFileName(customName)
        Else
            baseName = "Kask24x" & widthMm.ToString()
        End If
        
        ' First try without counter
        Dim fullPath As String = System.IO.Path.Combine(kaskFolder, baseName & ".ipt")
        If Not System.IO.File.Exists(fullPath) Then
            Return fullPath
        End If
        
        ' Add counter to make unique
        Dim counter As Integer = 1
        Dim fileName As String
        Do
            fileName = baseName & "-" & counter.ToString() & ".ipt"
            fullPath = System.IO.Path.Combine(kaskFolder, fileName)
            If Not System.IO.File.Exists(fullPath) Then
                Return fullPath
            End If
            counter += 1
        Loop While counter < 1000
        
        Return System.IO.Path.Combine(kaskFolder, baseName & "-" & Guid.NewGuid().ToString().Substring(0, 8) & ".ipt")
    End Function

    ''' <summary>
    ''' Sanitize a string for use as a filename.
    ''' </summary>
    Private Function SanitizeFileName(name As String) As String
        Dim invalidChars As Char() = System.IO.Path.GetInvalidFileNameChars()
        Dim result As String = name
        For Each c As Char In invalidChars
            result = result.Replace(c, "_"c)
        Next
        Return result.Trim()
    End Function

    ''' <summary>
    ''' Create a new support part from template with specified width and length.
    ''' </summary>
    Public Function CreateSupportPart(app As Inventor.Application, templatePath As String, _
                                       targetPath As String, widthMm As Integer, lengthCm As Double) As Boolean
        Try
            System.IO.File.Copy(templatePath, targetPath, True)
            
            Dim partDoc As PartDocument = CType(app.Documents.Open(targetPath, False), PartDocument)
            Try
                Dim params As Parameters = partDoc.ComponentDefinition.Parameters
                
                ' Set width (cm)
                Dim widthParam As Parameter = params.Item("Width")
                widthParam.Expression = (widthMm / 10.0).ToString() & " cm"
                
                ' Set length (cm)
                Dim lengthParam As Parameter = params.Item("Length")
                lengthParam.Expression = lengthCm.ToString() & " cm"
                
                partDoc.Update()
                partDoc.Save()
            Finally
                partDoc.Close()
            End Try
            
            Return True
        Catch
            Return False
        End Try
    End Function

    ' ============================================================================
    ' SECTION 3: Placement Matrix Creation
    ' ============================================================================

    ''' <summary>
    ''' Create placement matrix from start point and direction.
    ''' </summary>
    Public Function CreatePlacementMatrix(app As Inventor.Application, _
                                           startPoint As Point, direction As UnitVector) As Matrix
        Dim matrix As Matrix = app.TransientGeometry.CreateMatrix()
        
        ' Z axis = length direction
        Dim zAxis As UnitVector = direction
        
        ' Create perpendicular X and Y axes
        Dim xAxis As UnitVector
        Dim yAxis As UnitVector
        
        ' Find a vector not parallel to Z
        Dim tempVec As Vector
        If Math.Abs(zAxis.Z) < 0.9 Then
            tempVec = app.TransientGeometry.CreateVector(0, 0, 1)
        Else
            tempVec = app.TransientGeometry.CreateVector(1, 0, 0)
        End If
        
        ' X = temp × Z (perpendicular to Z)
        Dim xVec As Vector = tempVec.CrossProduct(app.TransientGeometry.CreateVector(zAxis.X, zAxis.Y, zAxis.Z))
        xVec.Normalize()
        xAxis = app.TransientGeometry.CreateUnitVector(xVec.X, xVec.Y, xVec.Z)
        
        ' Y = Z × X
        Dim yVec As Vector = app.TransientGeometry.CreateVector(zAxis.X, zAxis.Y, zAxis.Z).CrossProduct(xVec)
        yVec.Normalize()
        yAxis = app.TransientGeometry.CreateUnitVector(yVec.X, yVec.Y, yVec.Z)
        
        ' Set rotation part of matrix
        matrix.Cell(1, 1) = xAxis.X : matrix.Cell(2, 1) = xAxis.Y : matrix.Cell(3, 1) = xAxis.Z
        matrix.Cell(1, 2) = yAxis.X : matrix.Cell(2, 2) = yAxis.Y : matrix.Cell(3, 2) = yAxis.Z
        matrix.Cell(1, 3) = zAxis.X : matrix.Cell(2, 3) = zAxis.Y : matrix.Cell(3, 3) = zAxis.Z
        
        ' Set translation
        matrix.Cell(1, 4) = startPoint.X
        matrix.Cell(2, 4) = startPoint.Y
        matrix.Cell(3, 4) = startPoint.Z
        
        Return matrix
    End Function

    ''' <summary>
    ''' Create full placement matrix including align point offset and orientation.
    ''' Order of operations:
    ''' 1. Create basic matrix at startPoint with Z along direction
    ''' 2. Apply orientation (rotates X/Y axes around Z axis)
    ''' 3. Apply align point offset (uses the rotated X/Y axes)
    ''' 4. Apply custom X/Y/Z offset (in local coordinates, already in cm)
    ''' </summary>
    Public Function CreateFullPlacementMatrix(app As Inventor.Application, _
                                               startPoint As Point, direction As UnitVector, _
                                               alignPoint As String, widthMm As Integer, _
                                               orientMode As String, orientRef As Object, _
                                               Optional offsetXCm As Double = 0, _
                                               Optional offsetYCm As Double = 0, _
                                               Optional offsetZCm As Double = 0) As Matrix
        
        ' Start with basic matrix
        Dim matrix As Matrix = CreatePlacementMatrix(app, startPoint, direction)
        
        ' Apply orientation FIRST (rotates around Z/length axis)
        If orientMode <> "" AndAlso orientMode <> "NONE" AndAlso orientRef IsNot Nothing Then
            ApplyOrientation(app, matrix, orientMode, orientRef, direction)
        End If
        
        ' Apply align point offset AFTER orientation (uses rotated X/Y axes)
        ApplyAlignPointOffset(app, matrix, alignPoint, widthMm)
        
        ' Apply custom X/Y/Z offset (in local coordinates)
        If offsetXCm <> 0 OrElse offsetYCm <> 0 OrElse offsetZCm <> 0 Then
            ApplyCustomOffset(app, matrix, offsetXCm, offsetYCm, offsetZCm)
        End If
        
        Return matrix
    End Function

    ''' <summary>
    ''' Apply custom X/Y/Z offset to placement matrix (in global coordinates).
    ''' X, Y, Z correspond to world axes.
    ''' </summary>
    Private Sub ApplyCustomOffset(app As Inventor.Application, matrix As Matrix, _
                                   offsetXCm As Double, offsetYCm As Double, offsetZCm As Double)
        ' Apply offset in global coordinates (X, Y, Z world axes)
        matrix.Cell(1, 4) = matrix.Cell(1, 4) + offsetXCm  ' Global X
        matrix.Cell(2, 4) = matrix.Cell(2, 4) + offsetYCm  ' Global Y
        matrix.Cell(3, 4) = matrix.Cell(3, 4) + offsetZCm  ' Global Z
    End Sub

    ''' <summary>
    ''' Apply align point offset to placement matrix.
    ''' </summary>
    Private Sub ApplyAlignPointOffset(app As Inventor.Application, matrix As Matrix, _
                                       alignPoint As String, widthMm As Integer)
        If alignPoint = "" OrElse alignPoint = "Origin" Then Exit Sub
        
        ' Calculate offset in local coordinates (cm)
        Dim offsetX As Double = 0
        Dim offsetY As Double = 0
        
        Dim halfWidth As Double = (widthMm / 10.0) / 2.0  ' half width in cm
        Dim halfThickness As Double = (SUPPORT_THICKNESS_MM / 10.0) / 2.0  ' half thickness in cm
        
        Select Case alignPoint
            Case "EndCenter"
                ' No offset needed
            Case "Corner_XpYn"
                offsetX = -halfWidth
                offsetY = halfThickness
            Case "Corner_XpYp"
                offsetX = -halfWidth
                offsetY = -halfThickness
            Case "Corner_XnYn"
                offsetX = halfWidth
                offsetY = halfThickness
            Case "Corner_XnYp"
                offsetX = halfWidth
                offsetY = -halfThickness
            Case "Edge_Xp"
                offsetX = -halfWidth
            Case "Edge_Xn"
                offsetX = halfWidth
            Case "Edge_Yp"
                offsetY = -halfThickness
            Case "Edge_Yn"
                offsetY = halfThickness
        End Select
        
        ' Transform offset to world coordinates
        Dim xAxis As Vector = app.TransientGeometry.CreateVector(matrix.Cell(1, 1), matrix.Cell(2, 1), matrix.Cell(3, 1))
        Dim yAxis As Vector = app.TransientGeometry.CreateVector(matrix.Cell(1, 2), matrix.Cell(2, 2), matrix.Cell(3, 2))
        
        ' Apply offset
        matrix.Cell(1, 4) = matrix.Cell(1, 4) + offsetX * xAxis.X + offsetY * yAxis.X
        matrix.Cell(2, 4) = matrix.Cell(2, 4) + offsetX * xAxis.Y + offsetY * yAxis.Y
        matrix.Cell(3, 4) = matrix.Cell(3, 4) + offsetX * xAxis.Z + offsetY * yAxis.Z
    End Sub

    ''' <summary>
    ''' Apply orientation adjustment to placement matrix.
    ''' </summary>
    Private Sub ApplyOrientation(app As Inventor.Application, matrix As Matrix, _
                                  orientMode As String, orientRef As Object, direction As UnitVector)
        If orientRef Is Nothing Then Exit Sub
        
        Dim refNormal As UnitVector = Nothing
        
        Select Case orientMode
            Case "ALIGN_AXIS"
                refNormal = UtilsLib.GetAxisDirection(orientRef)
            Case "ALIGN_BOTTOM", "ALIGN_SIDE"
                refNormal = UtilsLib.GetPlaneNormal(orientRef)
        End Select
        
        If refNormal Is Nothing Then Exit Sub
        
        ' Current Z axis (length direction)
        Dim zAxis As UnitVector = direction
        
        ' Project refNormal onto plane perpendicular to Z
        Dim dotZ As Double = refNormal.X * zAxis.X + refNormal.Y * zAxis.Y + refNormal.Z * zAxis.Z
        Dim projX As Double = refNormal.X - dotZ * zAxis.X
        Dim projY As Double = refNormal.Y - dotZ * zAxis.Y
        Dim projZ As Double = refNormal.Z - dotZ * zAxis.Z
        
        Dim projLen As Double = Math.Sqrt(projX * projX + projY * projY + projZ * projZ)
        If projLen < 0.0001 Then Exit Sub
        
        ' Normalize projected vector
        projX /= projLen
        projY /= projLen
        projZ /= projLen
        
        Dim newAxis As UnitVector
        If orientMode = "ALIGN_BOTTOM" Then
            ' Y axis should align with reference
            newAxis = app.TransientGeometry.CreateUnitVector(projX, projY, projZ)
            Dim xVec As Vector = app.TransientGeometry.CreateVector(newAxis.X, newAxis.Y, newAxis.Z).CrossProduct( _
                app.TransientGeometry.CreateVector(zAxis.X, zAxis.Y, zAxis.Z))
            xVec.Normalize()
            
            matrix.Cell(1, 1) = xVec.X : matrix.Cell(2, 1) = xVec.Y : matrix.Cell(3, 1) = xVec.Z
            matrix.Cell(1, 2) = newAxis.X : matrix.Cell(2, 2) = newAxis.Y : matrix.Cell(3, 2) = newAxis.Z
        Else
            ' X axis should align with reference (ALIGN_SIDE or ALIGN_AXIS)
            newAxis = app.TransientGeometry.CreateUnitVector(projX, projY, projZ)
            Dim yVec As Vector = app.TransientGeometry.CreateVector(zAxis.X, zAxis.Y, zAxis.Z).CrossProduct( _
                app.TransientGeometry.CreateVector(newAxis.X, newAxis.Y, newAxis.Z))
            yVec.Normalize()
            
            matrix.Cell(1, 1) = newAxis.X : matrix.Cell(2, 1) = newAxis.Y : matrix.Cell(3, 1) = newAxis.Z
            matrix.Cell(1, 2) = yVec.X : matrix.Cell(2, 2) = yVec.Y : matrix.Cell(3, 2) = yVec.Z
        End If
    End Sub

    ' ============================================================================
    ' SECTION 4: Placement Calculation
    ' ============================================================================

    ''' <summary>
    ''' Calculate placement from geometry references.
    ''' Returns True on success, populates startPoint, direction, and length.
    ''' </summary>
    Public Function CalculatePlacement(app As Inventor.Application, mode As String, _
                                        ref1 As Object, ref2 As Object, ref3 As Object, _
                                        manualLen As Double, flipDir As Boolean, _
                                        ByRef startPoint As Point, ByRef direction As UnitVector, _
                                        ByRef length As Double, ByRef errorMsg As String) As Boolean
        
        startPoint = Nothing
        direction = Nothing
        length = 0
        errorMsg = ""
        
        Select Case mode
            Case "TWO_POINTS"
                If ref1 Is Nothing OrElse ref2 Is Nothing Then
                    errorMsg = "Two points are required"
                    Return False
                End If
                startPoint = UtilsLib.GetPointGeometry(ref1)
                direction = UtilsLib.GetDirectionBetweenPoints(app, ref1, ref2)
                length = UtilsLib.MeasurePointDistance(ref1, ref2)
                
            Case "AXIS_TWO_PLANES"
                If ref1 Is Nothing OrElse ref2 Is Nothing OrElse ref3 Is Nothing Then
                    errorMsg = "Axis and two planes are required"
                    Return False
                End If
                direction = UtilsLib.GetAxisDirection(ref1)
                startPoint = UtilsLib.GetAxisPlaneIntersection(app, ref1, ref2)
                length = UtilsLib.MeasurePlaneDistance(ref2, ref3)
                
                ' Ensure direction points toward end plane
                If startPoint IsNot Nothing AndAlso direction IsNot Nothing Then
                    Dim endPoint As Point = UtilsLib.GetAxisPlaneIntersection(app, ref1, ref3)
                    If endPoint IsNot Nothing Then
                        Dim toEnd As UnitVector = UtilsLib.GetDirectionBetweenPoints(app, startPoint, endPoint)
                        If toEnd IsNot Nothing Then
                            Dim dot As Double = direction.X * toEnd.X + direction.Y * toEnd.Y + direction.Z * toEnd.Z
                            If dot < 0 Then
                                direction = app.TransientGeometry.CreateUnitVector(-direction.X, -direction.Y, -direction.Z)
                            End If
                        End If
                    End If
                End If
                
            Case "PLANE_AXIS_LENGTH"
                If ref1 Is Nothing OrElse ref2 Is Nothing Then
                    errorMsg = "Plane and axis are required"
                    Return False
                End If
                direction = UtilsLib.GetAxisDirection(ref2)
                startPoint = UtilsLib.GetAxisPlaneIntersection(app, ref2, ref1)
                If startPoint Is Nothing Then
                    startPoint = UtilsLib.GetClosestPointOnAxisToPlane(app, ref2, ref1)
                End If
                length = manualLen
                
            Case "POINT_AXIS_LENGTH"
                If ref1 Is Nothing OrElse ref2 Is Nothing Then
                    errorMsg = "Point and axis are required"
                    Return False
                End If
                startPoint = UtilsLib.GetPointGeometry(ref1)
                direction = UtilsLib.GetAxisDirection(ref2)
                length = manualLen
                
            Case "TWO_PLANES_POINT"
                If ref1 Is Nothing OrElse ref2 Is Nothing OrElse ref3 Is Nothing Then
                    errorMsg = "Two planes and a point are required"
                    Return False
                End If
                
                Dim pickedPoint As Point = UtilsLib.GetPointGeometry(ref3)
                If pickedPoint Is Nothing Then
                    errorMsg = "Could not get point geometry"
                    Return False
                End If
                
                direction = UtilsLib.GetPlaneNormalTowardPlane(app, ref1, ref2)
                If direction Is Nothing Then
                    errorMsg = "Could not calculate direction from planes"
                    Return False
                End If
                
                startPoint = UtilsLib.ProjectPointOntoPlane(app, pickedPoint, ref1)
                length = UtilsLib.MeasurePlaneDistance(ref1, ref2)
                
            Case Else
                errorMsg = "Unknown placement mode: " & mode
                Return False
        End Select
        
        ' Apply flip if requested
        If flipDir AndAlso direction IsNot Nothing Then
            direction = app.TransientGeometry.CreateUnitVector(-direction.X, -direction.Y, -direction.Z)
        End If
        
        ' Validate results
        If startPoint Is Nothing Then
            errorMsg = "Could not calculate start point"
            Return False
        End If
        If direction Is Nothing Then
            errorMsg = "Could not calculate direction"
            Return False
        End If
        If length <= 0 Then
            errorMsg = "Length must be positive"
            Return False
        End If
        
        Return True
    End Function

    ' ============================================================================
    ' SECTION 5: Occurrence Reference Storage
    ' ============================================================================

    ''' <summary>
    ''' Store geometry references on an occurrence for later updates.
    ''' References are stored as work feature names (strings).
    ''' LengthInput: Either a numeric value (mm) or a parameter name.
    ''' OffsetX/Y/Z: Either a numeric value (mm) or a parameter name.
    ''' </summary>
    Public Sub StoreOccurrenceReferences(occ As ComponentOccurrence, _
                                          mode As String, _
                                          ref1Name As String, ref2Name As String, ref3Name As String, _
                                          alignPoint As String, _
                                          orientMode As String, orientRefName As String, _
                                          lengthInput As String, _
                                          flipDirection As Boolean, _
                                          Optional offsetX As String = "", _
                                          Optional offsetY As String = "", _
                                          Optional offsetZ As String = "", _
                                          Optional customName As String = "")
        
        Dim attrSet As AttributeSet = Nothing
        Try
            attrSet = occ.AttributeSets.Item(ATTR_SET_NAME)
        Catch
            attrSet = occ.AttributeSets.Add(ATTR_SET_NAME)
        End Try
        
        SetAttribute(attrSet, "Mode", mode)
        SetAttribute(attrSet, "Ref1", ref1Name)
        SetAttribute(attrSet, "Ref2", ref2Name)
        SetAttribute(attrSet, "Ref3", ref3Name)
        SetAttribute(attrSet, "AlignPoint", alignPoint)
        SetAttribute(attrSet, "OrientMode", orientMode)
        SetAttribute(attrSet, "OrientRef", orientRefName)
        SetAttribute(attrSet, "LengthInput", If(lengthInput, ""))
        SetAttribute(attrSet, "FlipDirection", flipDirection.ToString())
        SetAttribute(attrSet, "OffsetX", If(offsetX, ""))
        SetAttribute(attrSet, "OffsetY", If(offsetY, ""))
        SetAttribute(attrSet, "OffsetZ", If(offsetZ, ""))
        SetAttribute(attrSet, "CustomName", If(customName, ""))
    End Sub

    ''' <summary>
    ''' Read stored references from an occurrence.
    ''' </summary>
    Public Function GetOccurrenceReferences(occ As ComponentOccurrence) As Dictionary(Of String, String)
        Dim refs As New Dictionary(Of String, String)
        
        Try
            Dim attrSet As AttributeSet = occ.AttributeSets.Item(ATTR_SET_NAME)
            
            refs("Mode") = GetAttribute(attrSet, "Mode", "")
            refs("Ref1") = GetAttribute(attrSet, "Ref1", "")
            refs("Ref2") = GetAttribute(attrSet, "Ref2", "")
            refs("Ref3") = GetAttribute(attrSet, "Ref3", "")
            refs("AlignPoint") = GetAttribute(attrSet, "AlignPoint", "")
            refs("OrientMode") = GetAttribute(attrSet, "OrientMode", "")
            refs("OrientRef") = GetAttribute(attrSet, "OrientRef", "")
            refs("LengthInput") = GetAttribute(attrSet, "LengthInput", "")
            ' Backwards compatibility: try old ManualLength attribute
            If refs("LengthInput") = "" Then
                Dim oldManual As String = GetAttribute(attrSet, "ManualLength", "")
                If oldManual <> "" AndAlso oldManual <> "0" Then
                    ' Convert old cm value to mm string
                    Dim oldVal As Double
                    If Double.TryParse(oldManual, oldVal) Then
                        refs("LengthInput") = (oldVal * 10).ToString()
                    End If
                End If
            End If
            refs("FlipDirection") = GetAttribute(attrSet, "FlipDirection", "False")
            refs("OffsetX") = GetAttribute(attrSet, "OffsetX", "")
            refs("OffsetY") = GetAttribute(attrSet, "OffsetY", "")
            refs("OffsetZ") = GetAttribute(attrSet, "OffsetZ", "")
            refs("CustomName") = GetAttribute(attrSet, "CustomName", "")
        Catch
        End Try
        
        Return refs
    End Function

    ''' <summary>
    ''' Check if an occurrence has support placement data.
    ''' </summary>
    Public Function HasPlacementData(occ As ComponentOccurrence) As Boolean
        Try
            Dim attrSet As AttributeSet = occ.AttributeSets.Item(ATTR_SET_NAME)
            Dim mode As String = GetAttribute(attrSet, "Mode", "")
            Return mode <> ""
        Catch
            Return False
        End Try
    End Function

    Private Sub SetAttribute(attrSet As AttributeSet, name As String, value As String)
        Try
            attrSet.Item(name).Value = value
        Catch
            attrSet.Add(name, ValueTypeEnum.kStringType, value)
        End Try
    End Sub

    Private Function GetAttribute(attrSet As AttributeSet, name As String, defaultValue As String) As String
        Try
            Return CStr(attrSet.Item(name).Value)
        Catch
            Return defaultValue
        End Try
    End Function

    ' ============================================================================
    ' SECTION 6: Work Feature Resolution
    ' ============================================================================

    ''' <summary>
    ''' Resolve a work feature by name in the assembly.
    ''' Returns the WorkPoint, WorkAxis, or WorkPlane with the given name, or Nothing if not found.
    ''' </summary>
    Public Function ResolveWorkFeatureByName(asmDoc As AssemblyDocument, name As String) As Object
        If name = "" Then Return Nothing
        
        Try
            Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            
            ' Check if this is a path-based reference (starts with @)
            If name.StartsWith("@") Then
                Return ResolveWorkFeatureByPath(asmDoc, name)
            End If
            
            ' Legacy/simple name - try assembly-level first
            Try
                For Each wp As WorkPoint In asmDef.WorkPoints
                    If wp.Name = name Then Return wp
                Next
            Catch
            End Try
            
            Try
                For Each wa As WorkAxis In asmDef.WorkAxes
                    If wa.Name = name Then Return wa
                Next
            Catch
            End Try
            
            Try
                For Each wpl As WorkPlane In asmDef.WorkPlanes
                    If wpl.Name = name Then Return wpl
                Next
            Catch
            End Try
            
            ' Fallback: search inside component occurrences (for backwards compatibility)
            Try
                Dim result As Object = SearchWorkFeatureByNameInOccurrences(asmDef.Occurrences, name)
                If result IsNot Nothing Then Return result
            Catch
            End Try
            
            Return Nothing
        Catch
            ' During event triggers, some API calls may fail - return Nothing gracefully
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Resolve a work feature using path-based reference.
    ''' Format: "@OccInternalName1/OccInternalName2|WorkFeatureName"
    ''' </summary>
    Private Function ResolveWorkFeatureByPath(asmDoc As AssemblyDocument, pathRef As String) As Object
        ' Parse the reference: "@path/to/occ|WorkFeatureName"
        If Not pathRef.StartsWith("@") Then Return Nothing
        
        Dim content As String = pathRef.Substring(1) ' Remove @
        Dim pipePos As Integer = content.LastIndexOf("|")
        If pipePos < 0 Then Return Nothing
        
        Dim occPath As String = content.Substring(0, pipePos)
        Dim wfName As String = content.Substring(pipePos + 1)
        
        If wfName = "" Then Return Nothing
        
        ' If occPath is empty, this is a malformed reference - fallback to name search
        If occPath = "" Then
            Return SearchWorkFeatureByNameInOccurrences(asmDoc.ComponentDefinition.Occurrences, wfName)
        End If
        
        ' Navigate to the occurrence using InternalNames
        Dim occ As ComponentOccurrence = FindOccurrenceByInternalPath(asmDoc.ComponentDefinition.Occurrences, occPath)
        If occ Is Nothing Then Return Nothing
        
        ' Find the work feature in this occurrence and create proxy
        Dim compDef As ComponentDefinition = occ.Definition
        
        Dim proxyResult As Object = Nothing
        
        ' Try WorkPoints
        Try
            For Each wp As WorkPoint In compDef.WorkPoints
                If wp.Name = wfName Then
                    occ.CreateGeometryProxy(wp, proxyResult)
                    Return proxyResult
                End If
            Next
        Catch
        End Try
        
        ' Try WorkAxes
        Try
            For Each wa As WorkAxis In compDef.WorkAxes
                If wa.Name = wfName Then
                    occ.CreateGeometryProxy(wa, proxyResult)
                    Return proxyResult
                End If
            Next
        Catch
        End Try
        
        ' Try WorkPlanes
        Try
            For Each wpl As WorkPlane In compDef.WorkPlanes
                If wpl.Name = wfName Then
                    occ.CreateGeometryProxy(wpl, proxyResult)
                    Return proxyResult
                End If
            Next
        Catch
        End Try
        
        Return Nothing
    End Function

    ''' <summary>
    ''' Find an occurrence by following a path of InternalNames.
    ''' Path format: "InternalName1/InternalName2/InternalName3"
    ''' </summary>
    Private Function FindOccurrenceByInternalPath(occurrences As ComponentOccurrences, path As String) As ComponentOccurrence
        If path = "" Then Return Nothing
        
        Dim parts As String() = path.Split("/"c)
        Dim currentOccurrences As ComponentOccurrences = occurrences
        Dim foundOcc As ComponentOccurrence = Nothing
        
        For Each internalName As String In parts
            foundOcc = Nothing
            
            ' Find occurrence with this internal name
            For Each occ As ComponentOccurrence In currentOccurrences
                Try
                    If occ.InternalName = internalName Then
                        foundOcc = occ
                        Exit For
                    End If
                Catch
                End Try
            Next
            
            If foundOcc Is Nothing Then Return Nothing
            
            ' If there are more parts, navigate into sub-assembly
            If Array.IndexOf(parts, internalName) < parts.Length - 1 Then
                If foundOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Try
                        Dim subAsmDef As AssemblyComponentDefinition = CType(foundOcc.Definition, AssemblyComponentDefinition)
                        currentOccurrences = subAsmDef.Occurrences
                    Catch
                        Return Nothing
                    End Try
                Else
                    Return Nothing ' Can't navigate into a part
                End If
            End If
        Next
        
        Return foundOcc
    End Function

    ''' <summary>
    ''' Recursively search for a work feature by name inside occurrences.
    ''' Used for backwards compatibility with old simple-name references.
    ''' Returns the proxy object if found.
    ''' </summary>
    Private Function SearchWorkFeatureByNameInOccurrences(occurrences As ComponentOccurrences, name As String) As Object
        For Each occ As ComponentOccurrence In occurrences
            Try
                ' Get the component definition
                Dim compDef As ComponentDefinition = occ.Definition
                
                Dim proxyResult As Object = Nothing
                
                ' Search WorkPoints in this occurrence
                Try
                    For Each wp As WorkPoint In compDef.WorkPoints
                        If wp.Name = name Then
                            ' Return the proxy (work feature in assembly context)
                            occ.CreateGeometryProxy(wp, proxyResult)
                            Return proxyResult
                        End If
                    Next
                Catch
                End Try
                
                ' Search WorkAxes in this occurrence
                Try
                    For Each wa As WorkAxis In compDef.WorkAxes
                        If wa.Name = name Then
                            occ.CreateGeometryProxy(wa, proxyResult)
                            Return proxyResult
                        End If
                    Next
                Catch
                End Try
                
                ' Search WorkPlanes in this occurrence
                Try
                    For Each wpl As WorkPlane In compDef.WorkPlanes
                        If wpl.Name = name Then
                            occ.CreateGeometryProxy(wpl, proxyResult)
                            Return proxyResult
                        End If
                    Next
                Catch
                End Try
                
                ' If this is a sub-assembly, search recursively
                If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Try
                        Dim subAsmDef As AssemblyComponentDefinition = CType(compDef, AssemblyComponentDefinition)
                        Dim result As Object = SearchWorkFeatureByNameInOccurrences(subAsmDef.Occurrences, name)
                        If result IsNot Nothing Then Return result
                    Catch
                    End Try
                End If
            Catch
            End Try
        Next
        
        Return Nothing
    End Function

    ' ============================================================================
    ' SECTION 7: Parameter Handling
    ' ============================================================================

    ''' <summary>
    ''' Get list of parameter names from the assembly.
    ''' Includes user parameters, linked parameters, and reference parameters.
    ''' Excludes model parameters (internal feature parameters).
    ''' </summary>
    Public Function GetUserParameterNames(asmDoc As AssemblyDocument) As String()
        Dim names As New System.Collections.Generic.List(Of String)
        
        Try
            Dim allParams As Parameters = asmDoc.ComponentDefinition.Parameters
            
            For Each param As Parameter In allParams
                Try
                    ' Skip model parameters (internal feature parameters like "d0", "d1", etc.)
                    If param.ParameterType = ParameterTypeEnum.kModelParameter Then
                        Continue For
                    End If
                    
                    ' Include user parameters, linked parameters, reference parameters, etc.
                    If Not names.Contains(param.Name) Then
                        names.Add(param.Name)
                    End If
                Catch
                    ' If we can't check type, try to include it anyway
                    If Not names.Contains(param.Name) Then
                        names.Add(param.Name)
                    End If
                End Try
            Next
        Catch
        End Try
        
        names.Sort()
        Return names.ToArray()
    End Function

    ''' <summary>
    ''' Resolve a parameter value by name. Returns value in cm (internal units).
    ''' If paramName is empty or not found, returns the fallbackValue.
    ''' </summary>
    Public Function ResolveParameterLength(asmDoc As AssemblyDocument, paramName As String, fallbackValue As Double) As Double
        If paramName = "" Then Return fallbackValue
        
        Try
            Dim params As Parameters = asmDoc.ComponentDefinition.Parameters
            Dim param As Parameter = params.Item(paramName)
            If param IsNot Nothing Then
                ' param.Value is in internal units (cm)
                Return param.Value
            End If
        Catch
        End Try
        
        Return fallbackValue
    End Function

    ''' <summary>
    ''' Resolve length input to a value in cm.
    ''' Input can be a numeric value (in mm) or a parameter name.
    ''' Returns the resolved length in cm, or 0 if it can't be resolved.
    ''' </summary>
    Public Function ResolveLengthInput(asmDoc As AssemblyDocument, input As String) As Double
        If input = "" Then Return 0
        
        ' Try to parse as number (assumed to be in mm)
        Dim valueMm As Double
        If Double.TryParse(input.Trim(), valueMm) Then
            Return valueMm / 10.0  ' Convert mm to cm
        End If
        
        ' It's a parameter name - resolve from assembly
        Return ResolveParameterLength(asmDoc, input.Trim(), 0)
    End Function

    ''' <summary>
    ''' Check if length input is a parameter name (vs numeric value).
    ''' </summary>
    Public Function IsParameterInput(input As String) As Boolean
        If input = "" Then Return False
        Dim valueMm As Double
        Return Not Double.TryParse(input.Trim(), valueMm)
    End Function

    ''' <summary>
    ''' Resolve offset input to a value in cm.
    ''' Input can be a numeric value (in mm) or a parameter name.
    ''' Returns the resolved offset in cm, or 0 if empty/invalid.
    ''' </summary>
    Public Function ResolveOffsetInput(asmDoc As AssemblyDocument, input As String) As Double
        If input = "" OrElse input = "0" Then Return 0
        
        ' Try to parse as number (assumed to be in mm)
        Dim valueMm As Double
        If Double.TryParse(input.Trim(), valueMm) Then
            Return valueMm / 10.0  ' Convert mm to cm
        End If
        
        ' It's a parameter name - resolve from assembly
        Return ResolveParameterLength(asmDoc, input.Trim(), 0)
    End Function

    ' ============================================================================
    ' SECTION 8: iProperty Management
    ' ============================================================================

    ''' <summary>
    ''' Update iProperties on a support part for BOM compatibility.
    ''' Tracks AutoPartNumber to detect user customization of Part Number.
    ''' If customName is provided, uses it as the Part Number (overrides auto-generation).
    ''' </summary>
    Public Sub UpdateSupportiProperties(partDoc As PartDocument, widthMm As Integer, lengthMm As Integer, _
                                         Optional customName As String = "")
        Try
            partDoc.Rebuild()
            partDoc.Update()
        Catch
        End Try
        
        Try
            Dim customProps As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
            
            ' Update custom iProperties for BOM
            ' Width = length (longest dimension), Height = beam width, Thickness = 24mm fixed
            ' Use UtilsLib formatting for consistent 3 decimal place precision
            SetProperty(customProps, "Width", UtilsLib.FormatDimensionMm(CDbl(lengthMm)))
            SetProperty(customProps, "Thickness", UtilsLib.FormatDimensionMm(SUPPORT_THICKNESS_MM))
            SetProperty(customProps, "Height", UtilsLib.FormatDimensionMm(CDbl(widthMm)))
            
            ' Calculate what the auto-generated Part Number would be
            Dim autoPartNumber As String = "Kask24x" & widthMm.ToString() & "x" & lengthMm.ToString()
            
            ' Determine the Part Number to use
            Dim targetPartNumber As String = autoPartNumber
            If customName <> "" Then
                ' User provided a custom name - use it
                targetPartNumber = customName
            End If
            
            ' Get current Part Number
            Dim currentPartNumber As String = ""
            Try
                currentPartNumber = CStr(designProps.Item("Part Number").Value)
            Catch
            End Try
            
            ' Get stored auto Part Number (what we last generated)
            Dim storedAutoPartNumber As String = ""
            Try
                storedAutoPartNumber = CStr(customProps.Item("AutoPartNumber").Value)
            Catch
            End Try
            
            ' Get stored custom name (what user provided before)
            Dim storedCustomName As String = ""
            Try
                storedCustomName = CStr(customProps.Item("CustomName").Value)
            Catch
            End Try
            
            ' Decide whether to update Part Number
            ' If custom name is provided, always use it
            ' Otherwise: Update if no stored auto (first time), or current matches stored (user hasn't customized)
            If customName <> "" Then
                ' User provided custom name - use it
                Try
                    designProps.Item("Part Number").Value = targetPartNumber
                Catch
                End Try
            ElseIf storedAutoPartNumber = "" OrElse currentPartNumber = storedAutoPartNumber OrElse currentPartNumber = storedCustomName Then
                ' Auto mode: user hasn't manually customized, update Part Number
                Try
                    designProps.Item("Part Number").Value = targetPartNumber
                Catch
                End Try
            End If
            ' Else: Custom mode - user has renamed in iProperties, keep their Part Number
            
            ' Always update stored values for tracking
            SetProperty(customProps, "AutoPartNumber", autoPartNumber)
            SetProperty(customProps, "CustomName", customName)
            
            partDoc.Save()
        Catch
        End Try
    End Sub

    Private Sub SetProperty(propSet As PropertySet, name As String, value As String)
        Try
            propSet.Item(name).Value = value
        Catch
            Try
                propSet.Add(value, name)
            Catch
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' Get the Part Number from a part document.
    ''' </summary>
    Public Function GetPartNumber(partDoc As PartDocument) As String
        Try
            Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
            Return CStr(designProps.Item("Part Number").Value)
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Get the instance number from an occurrence name (e.g., "PartName:1" returns "1").
    ''' </summary>
    Public Function GetInstanceNumber(occName As String) As String
        Dim colonPos As Integer = occName.LastIndexOf(":")
        If colonPos >= 0 AndAlso colonPos < occName.Length - 1 Then
            Return occName.Substring(colonPos + 1)
        End If
        Return "1"
    End Function

    ''' <summary>
    ''' Get the base name from an occurrence name (e.g., "PartName:1" returns "PartName").
    ''' </summary>
    Public Function GetOccurrenceBaseName(occName As String) As String
        Dim colonPos As Integer = occName.LastIndexOf(":")
        If colonPos > 0 Then
            Return occName.Substring(0, colonPos)
        End If
        Return occName
    End Function

    ''' <summary>
    ''' Sync all occurrence names of a part to match its Part Number.
    ''' </summary>
    Public Sub SyncOccurrenceNames(asmDoc As AssemblyDocument, partDoc As PartDocument)
        Dim partNumber As String = GetPartNumber(partDoc)
        If partNumber = "" Then Exit Sub
        
        ' Find all occurrences of this part and rename them
        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            Try
                If occ.Definition.Document Is partDoc Then
                    ' Keep instance number suffix
                    Dim instanceNum As String = GetInstanceNumber(occ.Name)
                    Dim newName As String = partNumber & ":" & instanceNum
                    If occ.Name <> newName Then
                        occ.Name = newName
                    End If
                End If
            Catch
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Check if any occurrence of a part was renamed by the user.
    ''' If so, adopt that name as the Part Number.
    ''' Call this BEFORE UpdateSupportiProperties to detect user renames.
    ''' </summary>
    Public Sub DetectAndAdoptOccurrenceRename(asmDoc As AssemblyDocument, partDoc As PartDocument)
        Try
            Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
            Dim customProps As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            
            ' Get current Part Number
            Dim currentPartNumber As String = ""
            Try
                currentPartNumber = CStr(designProps.Item("Part Number").Value)
            Catch
            End Try
            
            ' Get stored auto Part Number
            Dim storedAutoPartNumber As String = ""
            Try
                storedAutoPartNumber = CStr(customProps.Item("AutoPartNumber").Value)
            Catch
            End Try
            
            ' Check each occurrence of this part
            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                Try
                    If occ.Definition.Document Is partDoc Then
                        Dim occBaseName As String = GetOccurrenceBaseName(occ.Name)
                        
                        ' If occurrence base name differs from both current Part Number and stored auto,
                        ' user has renamed the occurrence - adopt that name
                        If occBaseName <> currentPartNumber AndAlso occBaseName <> storedAutoPartNumber Then
                            ' User renamed occurrence, adopt that name as Part Number
                            designProps.Item("Part Number").Value = occBaseName
                            Exit For ' Only need to find one renamed occurrence
                        End If
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

    ' ============================================================================
    ' SECTION 8: Support Identification
    ' ============================================================================

    ''' <summary>
    ''' Check if a part document is a Kask24 support.
    ''' </summary>
    Public Function IsKask24Support(partDoc As PartDocument) As Boolean
        Try
            Dim fileName As String = System.IO.Path.GetFileName(partDoc.FullFileName)
            
            ' Check if filename starts with Kask24 (traditional naming)
            If fileName.StartsWith("Kask24") Then Return True
            
            ' Check if the part is in the Kask subfolder
            Dim folderPath As String = System.IO.Path.GetDirectoryName(partDoc.FullFileName)
            Dim folderName As String = System.IO.Path.GetFileName(folderPath)
            If folderName = "Kask" Then Return True
            
            ' Check if the part has Width and Length parameters (signature of a Kask support)
            Try
                Dim params As Parameters = partDoc.ComponentDefinition.Parameters
                Dim widthParam As Parameter = params.Item("Width")
                Dim lengthParam As Parameter = params.Item("Length")
                ' Both parameters exist - likely a Kask support
                Return True
            Catch
            End Try
            
            Return False
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Update support length parameter.
    ''' </summary>
    Public Sub UpdateSupportLength(partDoc As PartDocument, newLengthCm As Double)
        Try
            Dim params As Parameters = partDoc.ComponentDefinition.Parameters
            Dim lengthParam As Parameter = params.Item("Length")
            lengthParam.Value = newLengthCm
            partDoc.Update()
        Catch
        End Try
    End Sub

End Module
