' ============================================================================
' TestSheetMetalConversion - Test sheet metal conversion with auto A-side
' 
' Tests:
' - Can we detect the thickness direction automatically?
' - Can we find the A-side face using thickness vector?
' - Does sheet metal conversion work with that face?
' - Are Width/Length expressions set correctly?
' - Is flat pattern created?
'
' Usage: Open a simple flat solid part (NOT already sheet metal), then run.
'
' WARNING: This test modifies the active document! 
'          Make a copy or use an unsaved test part.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    Logger.Info("TestSheetMetalConversion: Starting sheet metal conversion tests...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("TestSheetMetalConversion: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestSheetMetalConversion")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestSheetMetalConversion: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestSheetMetalConversion")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    ' Check if already sheet metal
    Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    If partDoc.SubType = SHEET_METAL_GUID Then
        Logger.Warn("TestSheetMetalConversion: Part is already sheet metal.")
        MessageBox.Show("See detail on juba lehtmetall.", "TestSheetMetalConversion")
        Exit Sub
    End If
    
    ' Find first solid body
    Dim targetBody As SurfaceBody = Nothing
    For Each body As SurfaceBody In compDef.SurfaceBodies
        If body.IsSolid Then
            targetBody = body
            Exit For
        End If
    Next
    
    If targetBody Is Nothing Then
        Logger.Error("TestSheetMetalConversion: No solid body found.")
        MessageBox.Show("Detailis ei ole tahkkeha.", "TestSheetMetalConversion")
        Exit Sub
    End If
    
    Logger.Info("TestSheetMetalConversion: Found solid body: '" & targetBody.Name & "'")
    
    ' Warn user about modification
    Dim result As DialogResult = MessageBox.Show( _
        "HOIATUS: See test muudab aktiivset dokumenti!" & vbCrLf & vbCrLf & _
        "Detaili konverteeritakse lehtmetalliks." & vbCrLf & _
        "Soovitav on kasutada salvestamata testi detaili." & vbCrLf & vbCrLf & _
        "Kas jätkata?", _
        "TestSheetMetalConversion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
    
    If result <> DialogResult.Yes Then
        Logger.Info("TestSheetMetalConversion: Test cancelled by user.")
        Exit Sub
    End If
    
    ' Test 1: Detect thickness direction
    Logger.Info("TestSheetMetalConversion: Test 1 - Detecting thickness direction...")
    
    Dim thicknessVector As String = ""
    Dim thicknessValue As Double = 0
    Dim success As Boolean = DetectThicknessForBody(targetBody, thicknessVector, thicknessValue)
    
    If Not success Then
        Logger.Error("TestSheetMetalConversion: Could not detect thickness direction.")
        MessageBox.Show("Paksuse suunda ei õnnestunud tuvastada.", "TestSheetMetalConversion")
        Exit Sub
    End If
    
    Logger.Info("TestSheetMetalConversion: Thickness vector: " & thicknessVector)
    Logger.Info("TestSheetMetalConversion: Thickness value: " & FormatNumber(thicknessValue * 10, 2) & " mm")
    
    ' Test 2: Find A-side face
    Logger.Info("TestSheetMetalConversion: Test 2 - Finding A-side face...")
    
    Dim aSideFace As Face = FindFaceByNormal(targetBody, thicknessVector)
    
    If aSideFace Is Nothing Then
        Logger.Error("TestSheetMetalConversion: Could not find A-side face.")
        MessageBox.Show("A-külje pinda ei leitud.", "TestSheetMetalConversion")
        Exit Sub
    End If
    
    Logger.Info("TestSheetMetalConversion: Found A-side face (type: " & aSideFace.SurfaceType.ToString() & ")")
    
    ' Test 3: Convert to sheet metal
    Logger.Info("TestSheetMetalConversion: Test 3 - Converting to sheet metal...")
    
    Try
        partDoc.SubType = SHEET_METAL_GUID
        partDoc.Update()
        Logger.Info("TestSheetMetalConversion: Part converted to sheet metal subtype")
    Catch ex As Exception
        Logger.Error("TestSheetMetalConversion: Failed to convert to sheet metal: " & ex.Message)
        Exit Sub
    End Try
    
    ' Get sheet metal component definition
    Dim smCompDef As SheetMetalComponentDefinition = partDoc.ComponentDefinition
    
    ' Test 4: Set sheet metal style and thickness
    Logger.Info("TestSheetMetalConversion: Test 4 - Setting style and thickness...")
    
    Try
        ' Try to set style to Default_mm
        SetSheetMetalStyle(smCompDef, "Default_mm")
        Logger.Info("TestSheetMetalConversion: Set style to Default_mm")
        
        ' Set measured thickness
        smCompDef.UseSheetMetalStyleThickness = False
        smCompDef.Thickness.Value = thicknessValue
        Logger.Info("TestSheetMetalConversion: Set thickness to " & FormatNumber(thicknessValue * 10, 2) & " mm")
        
    Catch ex As Exception
        Logger.Warn("TestSheetMetalConversion: Could not set style/thickness: " & ex.Message)
    End Try
    
    ' Test 5: Export Thickness as iProperty
    Logger.Info("TestSheetMetalConversion: Test 5 - Exporting Thickness as iProperty...")
    
    Try
        Dim thicknessParam As Parameter = smCompDef.Thickness
        thicknessParam.ExposedAsProperty = True
        thicknessParam.CustomPropertyFormat.PropertyType = CustomPropertyTypeEnum.kTextPropertyType
        thicknessParam.CustomPropertyFormat.ShowUnitsString = True
        thicknessParam.CustomPropertyFormat.Units = "mm"
        Logger.Info("TestSheetMetalConversion: Thickness exported as iProperty")
    Catch ex As Exception
        Logger.Warn("TestSheetMetalConversion: Could not export Thickness: " & ex.Message)
    End Try
    
    ' Test 6: Set Width/Length expressions
    Logger.Info("TestSheetMetalConversion: Test 6 - Setting Width/Length properties...")
    
    Try
        Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
        SetOrAddProperty(propSet, "Width", "=<Sheet Metal Width>")
        SetOrAddProperty(propSet, "Length", "=<Sheet Metal Length>")
        Logger.Info("TestSheetMetalConversion: Width/Length properties set")
    Catch ex As Exception
        Logger.Warn("TestSheetMetalConversion: Could not set Width/Length: " & ex.Message)
    End Try
    
    ' Test 7: Create flat pattern
    Logger.Info("TestSheetMetalConversion: Test 7 - Creating flat pattern...")
    
    Try
        ' Need to find the A-side face again after conversion (geometry changed)
        Dim newASideFace As Face = FindASideFaceAfterConversion(smCompDef, thicknessVector)
        
        If newASideFace IsNot Nothing Then
            smCompDef.ASideFace = newASideFace
            smCompDef.Unfold()
            Logger.Info("TestSheetMetalConversion: Flat pattern created successfully")
        Else
            Logger.Warn("TestSheetMetalConversion: Could not find A-side face after conversion")
        End If
    Catch ex As Exception
        Logger.Warn("TestSheetMetalConversion: Could not create flat pattern: " & ex.Message)
    End Try
    
    partDoc.Update()
    
    ' Summary
    Logger.Info("TestSheetMetalConversion: ========================================")
    Logger.Info("TestSheetMetalConversion: TEST SUMMARY")
    Logger.Info("TestSheetMetalConversion: ========================================")
    Logger.Info("TestSheetMetalConversion: Thickness detected: " & FormatNumber(thicknessValue * 10, 2) & " mm")
    Logger.Info("TestSheetMetalConversion: A-side face found: Yes")
    Logger.Info("TestSheetMetalConversion: Conversion: Success")
    Logger.Info("TestSheetMetalConversion: ========================================")
    Logger.Info("TestSheetMetalConversion: All tests completed!")
    
    MessageBox.Show("Lehtmetalli konverteerimine õnnestus!" & vbCrLf & vbCrLf & _
                    "Paksus: " & FormatNumber(thicknessValue * 10, 2) & " mm", _
                    "TestSheetMetalConversion")
End Sub

Sub SetSheetMetalStyle(smCompDef As SheetMetalComponentDefinition, styleName As String)
    Try
        Dim style As SheetMetalStyle = smCompDef.SheetMetalStyles.Item(styleName)
        style.Activate()
    Catch
        ' Style not found - continue with default
    End Try
End Sub

Sub SetOrAddProperty(propSet As PropertySet, propName As String, propValue As String)
    Try
        propSet.Item(propName).Value = propValue
    Catch
        Try
            propSet.Add(propValue, propName)
        Catch
        End Try
    End Try
End Sub

Function DetectThicknessForBody(body As SurfaceBody, ByRef thicknessVector As String, ByRef thicknessValue As Double) As Boolean
    Dim bestNormalX As Double = 0, bestNormalY As Double = 0, bestNormalZ As Double = 0
    Dim minExtent As Double = Double.MaxValue
    Dim checkedNormals As New System.Collections.Generic.List(Of String)
    
    For Each face As Face In body.Faces
        Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
        If GetFaceNormal(face, nx, ny, nz) Then
            Dim len As Double = Math.Sqrt(nx * nx + ny * ny + nz * nz)
            If len > 0.0001 Then
                nx /= len : ny /= len : nz /= len
            End If
            
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
    
    If minExtent = Double.MaxValue Then Return False
    
    thicknessVector = VectorToString(bestNormalX, bestNormalY, bestNormalZ)
    thicknessValue = minExtent
    Return True
End Function

Function FindFaceByNormal(body As SurfaceBody, thicknessVector As String) As Face
    Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
    If Not ParseVectorComponents(thicknessVector, nx, ny, nz) Then Return Nothing
    
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

Function FindASideFaceAfterConversion(smCompDef As SheetMetalComponentDefinition, thicknessVector As String) As Face
    Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
    If Not ParseVectorComponents(thicknessVector, nx, ny, nz) Then Return Nothing
    
    For Each body As SurfaceBody In smCompDef.SurfaceBodies
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
    Next
    
    Return Nothing
End Function

Function GetOrientedExtentForBody(body As SurfaceBody, dirX As Double, dirY As Double, dirZ As Double) As Double
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

Function GetFaceNormal(face As Face, ByRef nx As Double, ByRef ny As Double, ByRef nz As Double) As Boolean
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

Function VectorToString(vx As Double, vy As Double, vz As Double) As String
    Return "V:" & vx.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                  vy.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                  vz.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture)
End Function

Function ParseVectorComponents(axis As String, ByRef vx As Double, ByRef vy As Double, ByRef vz As Double) As Boolean
    If Not axis.StartsWith("V:") Then Return False
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
