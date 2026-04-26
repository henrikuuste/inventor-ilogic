' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' TestDerivedPart - Test creating a derived part from a single solid body
' 
' Tests:
' - Can we create a new part document (from default template)?
' - Can we derive from a specific body (exclude others)?
' - Does the derived part have exactly one solid body?
' - Can we find faces by normal in the derived part?
'
' Usage: Open a multi-body part with 2+ solid bodies, then run this rule.
'        The first solid body will be derived into a new part.
'
' Note: Creates a temporary unsaved part document for testing.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    Logger.Info("TestDerivedPart: Starting derived part creation tests...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("TestDerivedPart: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestDerivedPart")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestDerivedPart: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestDerivedPart")
        Exit Sub
    End If
    
    Dim masterDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = masterDoc.ComponentDefinition
    
    ' Find solid bodies
    Dim solidBodies As New System.Collections.Generic.List(Of SurfaceBody)
    For Each body As SurfaceBody In compDef.SurfaceBodies
        If body.IsSolid Then
            solidBodies.Add(body)
        End If
    Next
    
    If solidBodies.Count = 0 Then
        Logger.Error("TestDerivedPart: No solid bodies found in part.")
        MessageBox.Show("Detailis ei ole tahkkehasid.", "TestDerivedPart")
        Exit Sub
    End If
    
    Logger.Info("TestDerivedPart: Found " & solidBodies.Count & " solid body(ies)")
    
    ' Use the first body for testing
    Dim targetBody As SurfaceBody = solidBodies(0)
    Logger.Info("TestDerivedPart: Will derive body: '" & targetBody.Name & "'")
    
    ' Detect thickness vector for the target body (for later face finding test)
    Dim thicknessVector As String = DetectThicknessVector(targetBody)
    Logger.Info("TestDerivedPart: Detected thickness vector: " & thicknessVector)
    
    ' Test 1: Create new part document from template
    Logger.Info("TestDerivedPart: Test 1 - Creating new part document from template...")
    Dim newPart As PartDocument = Nothing
    
    ' Find template path
    Dim templatePath As String = FindPartTemplate(app)
    If String.IsNullOrEmpty(templatePath) Then
        Logger.Warn("TestDerivedPart: No template found, using default")
    Else
        Logger.Info("TestDerivedPart: Using template: " & templatePath)
    End If
    
    Try
        If String.IsNullOrEmpty(templatePath) Then
            newPart = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject)
        Else
            newPart = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject, templatePath, True)
        End If
        Logger.Info("TestDerivedPart: New part document created successfully")
    Catch ex As Exception
        Logger.Error("TestDerivedPart: Failed to create new part: " & ex.Message)
        Exit Sub
    End Try
    
    ' Check if master document is saved
    If String.IsNullOrEmpty(masterDoc.FullDocumentName) OrElse Not System.IO.File.Exists(masterDoc.FullDocumentName) Then
        Logger.Error("TestDerivedPart: Master document must be saved before deriving.")
        MessageBox.Show("Põhidokument peab olema salvestatud enne tuletamist.", "TestDerivedPart")
        newPart.Close(True)
        Exit Sub
    End If
    
    Logger.Info("TestDerivedPart: Master document path: " & masterDoc.FullDocumentName)
    
    ' Test 2: Create derived part definition
    Logger.Info("TestDerivedPart: Test 2 - Creating derived part from body...")
    
    Try
        Dim dpcs As DerivedPartComponents = newPart.ComponentDefinition.ReferenceComponents.DerivedPartComponents
        Dim dpd As DerivedPartUniformScaleDef = dpcs.CreateUniformScaleDef(masterDoc.FullDocumentName)
        
        Logger.Info("TestDerivedPart: DerivedPartUniformScaleDef created")
        Logger.Info("TestDerivedPart: Number of solids in definition: " & dpd.Solids.Count)
        
        ' Log available solids
        For Each dpe As DerivedPartEntity In dpd.Solids
            Dim refEntity As Object = dpe.ReferencedEntity
            Dim bodyName As String = ""
            If TypeOf refEntity Is SurfaceBody Then
                bodyName = CType(refEntity, SurfaceBody).Name
            End If
            Logger.Info("TestDerivedPart:   Solid entity: '" & bodyName & "'")
        Next
        
        ' Exclude all bodies except the target (compare by name)
        Dim includedCount As Integer = 0
        Dim excludedCount As Integer = 0
        Dim targetBodyName As String = targetBody.Name
        
        For Each dpe As DerivedPartEntity In dpd.Solids
            Dim refEntity As Object = dpe.ReferencedEntity
            Dim bodyName As String = ""
            If TypeOf refEntity Is SurfaceBody Then
                bodyName = CType(refEntity, SurfaceBody).Name
            End If
            
            If bodyName = targetBodyName Then
                dpe.IncludeEntity = True
                includedCount += 1
                Logger.Info("TestDerivedPart: Including body: '" & bodyName & "'")
            Else
                dpe.IncludeEntity = False
                excludedCount += 1
                Logger.Info("TestDerivedPart: Excluding body: '" & bodyName & "'")
            End If
        Next
        
        Logger.Info("TestDerivedPart: Included: " & includedCount & ", Excluded: " & excludedCount)
        
        If includedCount = 0 Then
            Logger.Error("TestDerivedPart: No bodies matched - cannot derive")
            newPart.Close(True)
            Exit Sub
        End If
        
        ' Configure derivation options - use solid body style
        Logger.Info("TestDerivedPart: Setting derive style to single solid body...")
        dpd.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyWithSeams
        
        ' Add the derivation
        Logger.Info("TestDerivedPart: Adding derivation...")
        Dim derivedComp As DerivedPartComponent = dpcs.Add(dpd)
        Logger.Info("TestDerivedPart: Derivation added successfully")
        
        ' Update the document
        newPart.Update()
        
    Catch ex As Exception
        Logger.Error("TestDerivedPart: Failed to create derivation: " & ex.Message)
        newPart.Close(True)  ' Close without saving
        Exit Sub
    End Try
    
    ' Test 3: Verify derived part has exactly one solid body
    Logger.Info("TestDerivedPart: Test 3 - Verifying derived part geometry...")
    
    Dim newCompDef As PartComponentDefinition = newPart.ComponentDefinition
    Dim derivedBodyCount As Integer = 0
    
    For Each body As SurfaceBody In newCompDef.SurfaceBodies
        If body.IsSolid Then
            derivedBodyCount += 1
            Logger.Info("TestDerivedPart: Found derived solid body: '" & body.Name & "'")
        End If
    Next
    
    If derivedBodyCount = 1 Then
        Logger.Info("TestDerivedPart: SUCCESS - Derived part has exactly 1 solid body")
    Else
        Logger.Warn("TestDerivedPart: Derived part has " & derivedBodyCount & " solid bodies (expected 1)")
    End If
    
    ' Test 4: Find face by normal in derived part
    Logger.Info("TestDerivedPart: Test 4 - Finding face by normal vector...")
    
    Dim foundFace As Face = FindFaceByNormal(newPart, thicknessVector)
    
    If foundFace IsNot Nothing Then
        Logger.Info("TestDerivedPart: SUCCESS - Found face matching thickness normal")
        Logger.Info("TestDerivedPart: Face type: " & foundFace.SurfaceType.ToString())
    Else
        Logger.Warn("TestDerivedPart: Could not find face matching thickness normal")
    End If
    
    ' Summary
    Logger.Info("TestDerivedPart: ========================================")
    Logger.Info("TestDerivedPart: TEST SUMMARY")
    Logger.Info("TestDerivedPart: ========================================")
    Logger.Info("TestDerivedPart: New part created: Yes")
    Logger.Info("TestDerivedPart: Derivation successful: Yes")
    Logger.Info("TestDerivedPart: Derived body count: " & derivedBodyCount)
    Logger.Info("TestDerivedPart: Face by normal found: " & If(foundFace IsNot Nothing, "Yes", "No"))
    Logger.Info("TestDerivedPart: ========================================")
    
    ' Ask user what to do with the test part
    Dim result As DialogResult = MessageBox.Show( _
        "Derived part test lõpetatud edukalt!" & vbCrLf & vbCrLf & _
        "Kas soovid testidetaili sulgeda ilma salvestamata?", _
        "TestDerivedPart", MessageBoxButtons.YesNo)
    
    If result = DialogResult.Yes Then
        newPart.Close(True)  ' Close without saving
        Logger.Info("TestDerivedPart: Test part closed without saving")
    Else
        Logger.Info("TestDerivedPart: Test part left open for inspection")
    End If
    
    Logger.Info("TestDerivedPart: All tests completed!")
End Sub

Function DetectThicknessVector(body As SurfaceBody) As String
    ' Find the face normal with smallest extent (thickness direction)
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
            
            ' Make canonical
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
    
    ' Return as vector string
    Return VectorToString(bestNormalX, bestNormalY, bestNormalZ)
End Function

Function FindFaceByNormal(partDoc As PartDocument, thicknessVector As String) As Face
    Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
    If Not ParseVectorComponents(thicknessVector, nx, ny, nz) Then
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

Function FindPartTemplate(app As Inventor.Application) As String
    ' Try to find Part.ipt or Standard.ipt template
    ' Check multiple possible locations
    
    Dim templateNames() As String = {"Part.ipt", "Standard.ipt", "Metric\Standard (mm).ipt"}
    
    ' Get template paths from Inventor options
    Try
        Dim templatesPath As String = app.DesignProjectManager.ActiveDesignProject.TemplatesPath
        Logger.Info("TestDerivedPart: Templates folder: " & templatesPath)
        
        ' Try each template name
        For Each templateName As String In templateNames
            Dim fullPath As String = System.IO.Path.Combine(templatesPath, templateName)
            If System.IO.File.Exists(fullPath) Then
                Logger.Info("TestDerivedPart: Found template: " & fullPath)
                Return fullPath
            End If
        Next
        
        ' Also try en-US subfolder
        Dim enUSPath As String = System.IO.Path.Combine(templatesPath, "en-US")
        If System.IO.Directory.Exists(enUSPath) Then
            For Each templateName As String In templateNames
                Dim fullPath As String = System.IO.Path.Combine(enUSPath, templateName)
                If System.IO.File.Exists(fullPath) Then
                    Logger.Info("TestDerivedPart: Found template: " & fullPath)
                    Return fullPath
                End If
            Next
        End If
        
        ' List available templates for debugging
        Logger.Info("TestDerivedPart: Listing available .ipt files in templates folder...")
        If System.IO.Directory.Exists(templatesPath) Then
            Dim iptFiles() As String = System.IO.Directory.GetFiles(templatesPath, "*.ipt", System.IO.SearchOption.AllDirectories)
            For i As Integer = 0 To Math.Min(iptFiles.Length - 1, 9)
                Logger.Info("TestDerivedPart:   " & iptFiles(i))
            Next
            If iptFiles.Length > 10 Then
                Logger.Info("TestDerivedPart:   ... and " & (iptFiles.Length - 10) & " more")
            End If
        End If
        
    Catch ex As Exception
        Logger.Warn("TestDerivedPart: Error finding template: " & ex.Message)
    End Try
    
    Return ""
End Function
