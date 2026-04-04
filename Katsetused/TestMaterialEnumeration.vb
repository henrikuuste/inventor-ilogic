' ============================================================================
' TestMaterialEnumeration - Test enumerating available materials
' 
' Tests:
' - Can we enumerate partDoc.Materials?
' - What properties are available on Material objects?
' - Can we get the current material?
' - Can we set material on a part?
'
' Usage: Open any part document, then run this rule.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    Logger.Info("TestMaterialEnumeration: Starting material enumeration tests...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("TestMaterialEnumeration: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestMaterialEnumeration")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestMaterialEnumeration: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestMaterialEnumeration")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    ' Test 1: Get current material
    Logger.Info("TestMaterialEnumeration: Test 1 - Getting current material...")
    
    Dim currentMaterial As Material = Nothing
    Dim currentMaterialName As String = ""
    Try
        currentMaterial = compDef.Material
        If currentMaterial IsNot Nothing Then
            currentMaterialName = currentMaterial.Name
            Logger.Info("TestMaterialEnumeration: Current material: '" & currentMaterialName & "'")
        Else
            Logger.Info("TestMaterialEnumeration: No material assigned")
        End If
    Catch ex As Exception
        Logger.Warn("TestMaterialEnumeration: Could not get current material: " & ex.Message)
    End Try
    
    ' Test 2: Enumerate materials from document
    Logger.Info("TestMaterialEnumeration: Test 2 - Enumerating materials from document...")
    
    Dim materialCount As Integer = 0
    Dim materialNames As New System.Collections.Generic.List(Of String)
    
    Try
        For Each mat As Material In partDoc.Materials
            materialCount += 1
            materialNames.Add(mat.Name)
            
            ' Log first 20 materials
            If materialCount <= 20 Then
                Logger.Info("TestMaterialEnumeration:   " & materialCount & ". '" & mat.Name & "'")
            End If
        Next
        
        If materialCount > 20 Then
            Logger.Info("TestMaterialEnumeration:   ... and " & (materialCount - 20) & " more materials")
        End If
        
        Logger.Info("TestMaterialEnumeration: Total materials in document: " & materialCount)
        
    Catch ex As Exception
        Logger.Error("TestMaterialEnumeration: Exception enumerating materials: " & ex.Message)
    End Try
    
    ' Test 3: Try to access material by name
    Logger.Info("TestMaterialEnumeration: Test 3 - Accessing material by name...")
    
    ' Try some common material names
    Dim testNames() As String = {"Default", "Steel", "Aluminum", "Wood", "Generic"}
    
    For Each testName As String In testNames
        Try
            Dim mat As Material = partDoc.Materials.Item(testName)
            If mat IsNot Nothing Then
                Logger.Info("TestMaterialEnumeration: Found material '" & testName & "'")
            End If
        Catch
            ' Material not found - this is expected for most names
        End Try
    Next
    
    ' Test 4: Material library access
    Logger.Info("TestMaterialEnumeration: Test 4 - Checking material libraries...")
    
    Dim softcomLibrary As AssetLibrary = Nothing
    
    Try
        Dim assetLibs As AssetLibraries = app.AssetLibraries
        Logger.Info("TestMaterialEnumeration: Number of asset libraries: " & assetLibs.Count)
        
        For i As Integer = 1 To assetLibs.Count
            Dim assetLib As AssetLibrary = assetLibs.Item(i)
            Logger.Info("TestMaterialEnumeration:   Library: '" & assetLib.DisplayName & "' (Internal: " & assetLib.InternalName & ")")
            
            ' Check for SoftcomMaterials library
            If assetLib.DisplayName = "SoftcomMaterials" Then
                softcomLibrary = assetLib
            End If
        Next
        
    Catch ex As Exception
        Logger.Warn("TestMaterialEnumeration: Could not enumerate asset libraries: " & ex.Message)
    End Try
    
    ' Test 5: Try to access materials via app.ActiveMaterialLibrary or similar
    Logger.Info("TestMaterialEnumeration: Test 5 - Testing material assignment...")
    
    ' For the dropdown, we'll use partDoc.Materials which contains materials already
    ' available to the document. To add materials from a library, user needs to
    ' add them via Inventor UI first, or we can use Assets API.
    
    ' Test if we can get materials from app.Assets
    Try
        Logger.Info("TestMaterialEnumeration: Checking app.Assets for materials...")
        Dim matAssets As AssetsEnumerator = app.Assets(AssetTypeEnum.kAssetTypeMaterial)
        Dim assetCount As Integer = 0
        
        For Each asset As Asset In matAssets
            assetCount += 1
            If assetCount <= 30 Then
                Logger.Info("TestMaterialEnumeration:   Asset: '" & asset.DisplayName & "' from " & asset.LibraryName)
            End If
            
            ' Add to list
            If Not materialNames.Contains(asset.DisplayName) Then
                materialNames.Add(asset.DisplayName)
            End If
        Next
        
        If assetCount > 30 Then
            Logger.Info("TestMaterialEnumeration:   ... and " & (assetCount - 30) & " more assets")
        End If
        
        Logger.Info("TestMaterialEnumeration: Total material assets: " & assetCount)
        
    Catch ex As Exception
        Logger.Warn("TestMaterialEnumeration: Could not enumerate via app.Assets: " & ex.Message)
    End Try
    
    ' Summary
    Logger.Info("TestMaterialEnumeration: ========================================")
    Logger.Info("TestMaterialEnumeration: TEST SUMMARY")
    Logger.Info("TestMaterialEnumeration: ========================================")
    Logger.Info("TestMaterialEnumeration: Current material: " & If(currentMaterialName <> "", currentMaterialName, "None"))
    Logger.Info("TestMaterialEnumeration: Materials in document: " & materialCount)
    Logger.Info("TestMaterialEnumeration: ========================================")
    
    ' Show summary dialog
    Dim summaryText As String = "Materjalide loend:" & vbCrLf & vbCrLf
    summaryText &= "Praegune materjal: " & If(currentMaterialName <> "", currentMaterialName, "Puudub") & vbCrLf
    summaryText &= "Materjale dokumendis: " & materialCount & vbCrLf & vbCrLf
    
    If materialNames.Count > 0 Then
        summaryText &= "Esimesed materjalid:" & vbCrLf
        For i As Integer = 0 To Math.Min(materialNames.Count - 1, 9)
            summaryText &= "  - " & materialNames(i) & vbCrLf
        Next
    End If
    
    summaryText &= vbCrLf & "Vaata iLogic logi täieliku nimekirja jaoks."
    
    MessageBox.Show(summaryText, "TestMaterialEnumeration")
    
    Logger.Info("TestMaterialEnumeration: All tests completed!")
End Sub
