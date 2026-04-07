' ============================================================================
' SheetMetalLib - Sheet metal conversion with auto A-side detection
' 
' Provides functions to:
' - Detect thickness direction for a solid body
' - Find A-side face by normal vector
' - Convert a part to sheet metal
' - Create flat pattern
'
' Based on Lehtmetall.vb but with automated A-side detection.
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/SheetMetalLib.vb"
'   UtilsLib.SetLogger(Logger) ' In Sub Main
'
' Dependencies: UtilsLib (for logging)
' ============================================================================

Imports Inventor

Public Module SheetMetalLib

    Public Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    
    ' ============================================================================
    ' Thickness Detection
    ' ============================================================================
    
    Public Function DetectThicknessVector(body As SurfaceBody, _
                                          ByRef thicknessValue As Double) As String
        Dim bestNormalX As Double = 0, bestNormalY As Double = 0, bestNormalZ As Double = 0
        Dim minExtent As Double = Double.MaxValue
        Dim checkedNormals As New System.Collections.Generic.List(Of String)
        
        For Each face As Face In body.Faces
            Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
            If GetFaceNormal(face, nx, ny, nz) Then
                ' Normalize
                Dim len As Double = Math.Sqrt(nx * nx + ny * ny + nz * nz)
                If len > 0.0001 Then
                    nx /= len : ny /= len : nz /= len
                End If
                
                ' Make normal canonical (always point in "positive" direction)
                If nx < -0.0001 OrElse (Math.Abs(nx) < 0.0001 AndAlso ny < -0.0001) OrElse _
                   (Math.Abs(nx) < 0.0001 AndAlso Math.Abs(ny) < 0.0001 AndAlso nz < -0.0001) Then
                    nx = -nx : ny = -ny : nz = -nz
                End If
                
                Dim normalKey As String = Math.Round(nx, 3).ToString() & "," & _
                                          Math.Round(ny, 3).ToString() & "," & _
                                          Math.Round(nz, 3).ToString()
                
                If checkedNormals.Contains(normalKey) Then Continue For
                checkedNormals.Add(normalKey)
                
                Dim extent As Double = GetOrientedExtentForBody(body, nx, ny, nz)
                
                If extent > 0 AndAlso extent < minExtent Then
                    minExtent = extent
                    bestNormalX = nx
                    bestNormalY = ny
                    bestNormalZ = nz
                End If
            End If
        Next
        
        If minExtent = Double.MaxValue Then
            UtilsLib.LogWarn("SheetMetalLib: Could not detect thickness direction")
            thicknessValue = 0
            Return ""
        End If
        
        thicknessValue = minExtent
        UtilsLib.LogInfo("SheetMetalLib: Detected thickness: " & FormatNumber(minExtent * 10, 2) & " mm")
        
        Return VectorToString(bestNormalX, bestNormalY, bestNormalZ)
    End Function
    
    ' ============================================================================
    ' Face Finding
    ' ============================================================================
    
    Public Function FindFaceByNormal(partDoc As PartDocument, thicknessVector As String) As Face
        Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
        If Not ParseVectorComponents(thicknessVector, nx, ny, nz) Then
            UtilsLib.LogWarn("SheetMetalLib: Invalid thickness vector format")
            Return Nothing
        End If
        
        For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
            For Each face As Face In body.Faces
                Dim fx As Double = 0, fy As Double = 0, fz As Double = 0
                If GetFaceNormal(face, fx, fy, fz) Then
                    Dim len As Double = Math.Sqrt(fx * fx + fy * fy + fz * fz)
                    If len > 0.0001 Then
                        fx /= len : fy /= len : fz /= len
                    End If
                    
                    Dim dot As Double = nx * fx + ny * fy + nz * fz
                    If Math.Abs(Math.Abs(dot) - 1) < 0.001 Then
                        UtilsLib.LogInfo("SheetMetalLib: Found A-side face")
                        Return face
                    End If
                End If
            Next
        Next
        
        UtilsLib.LogWarn("SheetMetalLib: Could not find A-side face")
        Return Nothing
    End Function
    
    Public Function FindFaceByNormalOnBody(body As SurfaceBody, thicknessVector As String) As Face
        Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
        If Not ParseVectorComponents(thicknessVector, nx, ny, nz) Then
            Return Nothing
        End If
        
        For Each face As Face In body.Faces
            Dim fx As Double = 0, fy As Double = 0, fz As Double = 0
            If GetFaceNormal(face, fx, fy, fz) Then
                Dim len As Double = Math.Sqrt(fx * fx + fy * fy + fz * fz)
                If len > 0.0001 Then
                    fx /= len : fy /= len : fz /= len
                End If
                
                Dim dot As Double = nx * fx + ny * fy + nz * fz
                If Math.Abs(Math.Abs(dot) - 1) < 0.001 Then
                    Return face
                End If
            End If
        Next
        
        Return Nothing
    End Function
    
    ' ============================================================================
    ' Sheet Metal Conversion
    ' ============================================================================
    
    Public Function IsSheetMetal(partDoc As PartDocument) As Boolean
        Return partDoc.SubType = SHEET_METAL_GUID
    End Function
    
    Public Function ConvertToSheetMetal(partDoc As PartDocument, _
                                        thicknessVector As String, _
                                        thickness As Double) As Boolean
        If IsSheetMetal(partDoc) Then
            UtilsLib.LogInfo("SheetMetalLib: Part is already sheet metal")
            Return False
        End If
        
        Try
            partDoc.SubType = SHEET_METAL_GUID
            partDoc.Update()
            UtilsLib.LogInfo("SheetMetalLib: Converted to sheet metal subtype")
        Catch ex As Exception
            UtilsLib.LogWarn("SheetMetalLib: Failed to convert to sheet metal: " & ex.Message)
            Return False
        End Try
        
        Dim smCompDef As SheetMetalComponentDefinition = partDoc.ComponentDefinition
        
        SetSheetMetalStyle(smCompDef, "Default_mm")
        SetThickness(smCompDef, thickness)
        ExportThicknessAsProperty(smCompDef)
        SetWidthLengthProperties(partDoc)
        
        Dim aSideFace As Face = FindFaceByNormal(partDoc, thicknessVector)
        If aSideFace IsNot Nothing Then
            CreateFlatPattern(smCompDef, aSideFace)
        Else
            UtilsLib.LogWarn("SheetMetalLib: Could not create flat pattern - A-side face not found")
        End If
        
        partDoc.Update()
        UtilsLib.LogInfo("SheetMetalLib: Sheet metal conversion complete")
        Return True
    End Function
    
    Public Sub SetSheetMetalStyle(smCompDef As SheetMetalComponentDefinition, styleName As String)
        Try
            Dim style As SheetMetalStyle = smCompDef.SheetMetalStyles.Item(styleName)
            style.Activate()
            UtilsLib.LogInfo("SheetMetalLib: Set style to " & styleName)
        Catch
            UtilsLib.LogInfo("SheetMetalLib: Style '" & styleName & "' not found, using default")
        End Try
    End Sub
    
    Public Sub SetThickness(smCompDef As SheetMetalComponentDefinition, thickness As Double)
        Try
            smCompDef.UseSheetMetalStyleThickness = False
            smCompDef.Thickness.Value = thickness
            UtilsLib.LogInfo("SheetMetalLib: Set thickness to " & FormatNumber(thickness * 10, 2) & " mm")
        Catch ex As Exception
            UtilsLib.LogWarn("SheetMetalLib: Could not set thickness: " & ex.Message)
        End Try
    End Sub
    
    Public Sub ExportThicknessAsProperty(smCompDef As SheetMetalComponentDefinition)
        Try
            Dim thicknessParam As Parameter = smCompDef.Thickness
            thicknessParam.ExposedAsProperty = True
            thicknessParam.CustomPropertyFormat.PropertyType = CustomPropertyTypeEnum.kTextPropertyType
            thicknessParam.CustomPropertyFormat.ShowUnitsString = True
            thicknessParam.CustomPropertyFormat.Units = "mm"
            UtilsLib.LogInfo("SheetMetalLib: Exported Thickness as iProperty")
        Catch ex As Exception
            UtilsLib.LogWarn("SheetMetalLib: Could not export Thickness: " & ex.Message)
        End Try
    End Sub
    
    Public Sub SetWidthLengthProperties(partDoc As PartDocument)
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            SetOrAddProperty(propSet, "Width", "=<Sheet Metal Width>")
            SetOrAddProperty(propSet, "Length", "=<Sheet Metal Length>")
            UtilsLib.LogInfo("SheetMetalLib: Set Width/Length properties")
        Catch ex As Exception
            UtilsLib.LogWarn("SheetMetalLib: Could not set Width/Length: " & ex.Message)
        End Try
    End Sub
    
    Public Sub CreateFlatPattern(smCompDef As SheetMetalComponentDefinition, aSideFace As Face)
        Try
            smCompDef.ASideFace = aSideFace
            smCompDef.Unfold()
            UtilsLib.LogInfo("SheetMetalLib: Flat pattern created")
            
            If smCompDef.HasFlatPattern Then
                smCompDef.FlatPattern.ExitEdit()
                UtilsLib.LogInfo("SheetMetalLib: Returned to folded view")
            End If
        Catch ex As Exception
            UtilsLib.LogWarn("SheetMetalLib: Could not create flat pattern: " & ex.Message)
        End Try
    End Sub
    
    ' ============================================================================
    ' Helper Functions
    ' ============================================================================
    
    Private Sub SetOrAddProperty(propSet As PropertySet, propName As String, propValue As String)
        Try
            propSet.Item(propName).Value = propValue
        Catch
            Try
                propSet.Add(propValue, propName)
            Catch
            End Try
        End Try
    End Sub
    
    Public Function GetFaceNormal(face As Face, ByRef nx As Double, ByRef ny As Double, ByRef nz As Double) As Boolean
        Try
            Dim geom As Object = face.Geometry
            If TypeOf geom Is Plane Then
                Dim plane As Plane = CType(geom, Plane)
                Dim normal As UnitVector = plane.Normal
                nx = normal.X
                ny = normal.Y
                nz = normal.Z
                Return True
            End If
        Catch
        End Try
        Return False
    End Function
    
    Public Function GetOrientedExtentForBody(body As SurfaceBody, dirX As Double, dirY As Double, dirZ As Double) As Double
        Dim minProj As Double = Double.MaxValue
        Dim maxProj As Double = Double.MinValue
        
        Try
            For Each vertex As Vertex In body.Vertices
                Dim pt As Point = vertex.Point
                Dim proj As Double = pt.X * dirX + pt.Y * dirY + pt.Z * dirZ
                If proj < minProj Then minProj = proj
                If proj > maxProj Then maxProj = proj
            Next
        Catch
        End Try
        
        If minProj = Double.MaxValue Then Return 0
        Return maxProj - minProj
    End Function
    
    Public Function VectorToString(vx As Double, vy As Double, vz As Double) As String
        Return "V:" & vx.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                      vy.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                      vz.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture)
    End Function
    
    Public Function ParseVectorComponents(axis As String, ByRef vx As Double, ByRef vy As Double, ByRef vz As Double) As Boolean
        If String.IsNullOrEmpty(axis) OrElse Not axis.StartsWith("V:") Then Return False
        Try
            Dim parts() As String = axis.Substring(2).Split(","c)
            If parts.Length <> 3 Then Return False
            vx = Double.Parse(parts(0), System.Globalization.CultureInfo.InvariantCulture)
            vy = Double.Parse(parts(1), System.Globalization.CultureInfo.InvariantCulture)
            vz = Double.Parse(parts(2), System.Globalization.CultureInfo.InvariantCulture)
            Return True
        Catch
            Return False
        End Try
    End Function

End Module
