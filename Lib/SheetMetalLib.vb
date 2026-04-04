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
' Usage: AddVbFile "Lib/SheetMetalLib.vb"
'
' Note: Logger is not available in library modules.
'       Pass a List(Of String) to collect log messages.
' ============================================================================

Imports Inventor

Public Module SheetMetalLib

    Public Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    
    ' ============================================================================
    ' Thickness Detection
    ' ============================================================================
    
    ' Detect thickness direction for a solid body
    ' Returns the thickness vector as a string (e.g., "V:0,0,1" or "Z")
    Public Function DetectThicknessVector(body As SurfaceBody, _
                                          ByRef thicknessValue As Double, _
                                          logs As System.Collections.Generic.List(Of String)) As String
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
                
                ' Create key for this normal direction
                Dim normalKey As String = Math.Round(nx, 3).ToString() & "," & _
                                          Math.Round(ny, 3).ToString() & "," & _
                                          Math.Round(nz, 3).ToString()
                
                If checkedNormals.Contains(normalKey) Then Continue For
                checkedNormals.Add(normalKey)
                
                ' Calculate extent along this normal
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
            logs.Add("SheetMetalLib: Could not detect thickness direction")
            thicknessValue = 0
            Return ""
        End If
        
        thicknessValue = minExtent
        logs.Add("SheetMetalLib: Detected thickness: " & FormatNumber(minExtent * 10, 2) & " mm")
        
        Return VectorToString(bestNormalX, bestNormalY, bestNormalZ)
    End Function
    
    ' ============================================================================
    ' Face Finding
    ' ============================================================================
    
    ' Find a face whose normal matches the given vector (for A-side)
    Public Function FindFaceByNormal(partDoc As PartDocument, thicknessVector As String, _
                                     logs As System.Collections.Generic.List(Of String)) As Face
        Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
        If Not ParseVectorComponents(thicknessVector, nx, ny, nz) Then
            logs.Add("SheetMetalLib: Invalid thickness vector format")
            Return Nothing
        End If
        
        For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
            For Each face As Face In body.Faces
                Dim fx As Double = 0, fy As Double = 0, fz As Double = 0
                If GetFaceNormal(face, fx, fy, fz) Then
                    ' Normalize
                    Dim len As Double = Math.Sqrt(fx * fx + fy * fy + fz * fz)
                    If len > 0.0001 Then
                        fx /= len : fy /= len : fz /= len
                    End If
                    
                    ' Check if normals are parallel (dot product ~1 or ~-1)
                    Dim dot As Double = nx * fx + ny * fy + nz * fz
                    If Math.Abs(Math.Abs(dot) - 1) < 0.001 Then
                        logs.Add("SheetMetalLib: Found A-side face")
                        Return face
                    End If
                End If
            Next
        Next
        
        logs.Add("SheetMetalLib: Could not find A-side face")
        Return Nothing
    End Function
    
    ' Find A-side face on a specific body
    Public Function FindFaceByNormalOnBody(body As SurfaceBody, thicknessVector As String, _
                                           logs As System.Collections.Generic.List(Of String)) As Face
        Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
        If Not ParseVectorComponents(thicknessVector, nx, ny, nz) Then
            logs.Add("SheetMetalLib: Invalid thickness vector format")
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
    
    ' Check if a part is already sheet metal
    Public Function IsSheetMetal(partDoc As PartDocument) As Boolean
        Return partDoc.SubType = SHEET_METAL_GUID
    End Function
    
    ' Convert a part to sheet metal with auto-detected A-side
    Public Function ConvertToSheetMetal(partDoc As PartDocument, _
                                        thicknessVector As String, _
                                        thickness As Double, _
                                        logs As System.Collections.Generic.List(Of String)) As Boolean
        ' Check if already sheet metal
        If IsSheetMetal(partDoc) Then
            logs.Add("SheetMetalLib: Part is already sheet metal")
            Return False
        End If
        
        ' Convert to sheet metal subtype
        Try
            partDoc.SubType = SHEET_METAL_GUID
            partDoc.Update()
            logs.Add("SheetMetalLib: Converted to sheet metal subtype")
        Catch ex As Exception
            logs.Add("SheetMetalLib: Failed to convert to sheet metal: " & ex.Message)
            Return False
        End Try
        
        ' Get sheet metal component definition
        Dim smCompDef As SheetMetalComponentDefinition = partDoc.ComponentDefinition
        
        ' Set sheet metal style
        SetSheetMetalStyle(smCompDef, "Default_mm", logs)
        
        ' Set thickness
        SetThickness(smCompDef, thickness, logs)
        
        ' Export thickness as iProperty
        ExportThicknessAsProperty(smCompDef, logs)
        
        ' Set Width/Length properties
        SetWidthLengthProperties(partDoc, logs)
        
        ' Find A-side face and create flat pattern
        Dim aSideFace As Face = FindFaceByNormal(partDoc, thicknessVector, logs)
        If aSideFace IsNot Nothing Then
            CreateFlatPattern(smCompDef, aSideFace, logs)
        Else
            logs.Add("SheetMetalLib: Could not create flat pattern - A-side face not found")
        End If
        
        partDoc.Update()
        logs.Add("SheetMetalLib: Sheet metal conversion complete")
        Return True
    End Function
    
    ' Set sheet metal style
    Public Sub SetSheetMetalStyle(smCompDef As SheetMetalComponentDefinition, styleName As String, _
                                  logs As System.Collections.Generic.List(Of String))
        Try
            Dim style As SheetMetalStyle = smCompDef.SheetMetalStyles.Item(styleName)
            style.Activate()
            logs.Add("SheetMetalLib: Set style to " & styleName)
        Catch
            logs.Add("SheetMetalLib: Style '" & styleName & "' not found, using default")
        End Try
    End Sub
    
    ' Set thickness value
    Public Sub SetThickness(smCompDef As SheetMetalComponentDefinition, thickness As Double, _
                            logs As System.Collections.Generic.List(Of String))
        Try
            smCompDef.UseSheetMetalStyleThickness = False
            smCompDef.Thickness.Value = thickness
            logs.Add("SheetMetalLib: Set thickness to " & FormatNumber(thickness * 10, 2) & " mm")
        Catch ex As Exception
            logs.Add("SheetMetalLib: Could not set thickness: " & ex.Message)
        End Try
    End Sub
    
    ' Export thickness parameter as iProperty
    Public Sub ExportThicknessAsProperty(smCompDef As SheetMetalComponentDefinition, _
                                         logs As System.Collections.Generic.List(Of String))
        Try
            Dim thicknessParam As Parameter = smCompDef.Thickness
            thicknessParam.ExposedAsProperty = True
            thicknessParam.CustomPropertyFormat.PropertyType = CustomPropertyTypeEnum.kTextPropertyType
            thicknessParam.CustomPropertyFormat.ShowUnitsString = True
            thicknessParam.CustomPropertyFormat.Units = "mm"
            logs.Add("SheetMetalLib: Exported Thickness as iProperty")
        Catch ex As Exception
            logs.Add("SheetMetalLib: Could not export Thickness: " & ex.Message)
        End Try
    End Sub
    
    ' Set Width and Length custom properties with sheet metal expressions
    Public Sub SetWidthLengthProperties(partDoc As PartDocument, _
                                        logs As System.Collections.Generic.List(Of String))
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            SetOrAddProperty(propSet, "Width", "=<Sheet Metal Width>")
            SetOrAddProperty(propSet, "Length", "=<Sheet Metal Length>")
            logs.Add("SheetMetalLib: Set Width/Length properties")
        Catch ex As Exception
            logs.Add("SheetMetalLib: Could not set Width/Length: " & ex.Message)
        End Try
    End Sub
    
    ' Create flat pattern and return to folded view
    Public Sub CreateFlatPattern(smCompDef As SheetMetalComponentDefinition, aSideFace As Face, _
                                 logs As System.Collections.Generic.List(Of String))
        Try
            smCompDef.ASideFace = aSideFace
            smCompDef.Unfold()
            logs.Add("SheetMetalLib: Flat pattern created")
            
            ' Return to folded model view (required before save)
            ' Use FlatPattern.ExitEdit instead of FoldedModel
            If smCompDef.HasFlatPattern Then
                smCompDef.FlatPattern.ExitEdit()
                logs.Add("SheetMetalLib: Returned to folded view")
            End If
        Catch ex As Exception
            logs.Add("SheetMetalLib: Could not create flat pattern: " & ex.Message)
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
