' Copyright (c) 2026 Henri Kuuste
' Test15_DerivedPartReplace.vb
' PURPOSE: Test replacing derived base part reference with a different file
' 
' GOAL: Use SaveAs to create a new part with new GUID, then change the 
' derivation reference to point to a different base part that updates
' when the new base changes geometry.
'
' KEY INSIGHTS FROM TESTING:
' - DerivedPartComponent.Replace(path, options) works with ANY file
' - BUT Replace() brings in ALL bodies from new base (doesn't preserve config)
' - SaveAs preserves InternalName (same GUID)
'
' SOLUTION: Delete old derivation + Create new derivation with body selection
' - This preserves the ability to select specific bodies
' - We can match bodies by name or index
'
' TESTS:
' 1. Can we record the current derivation configuration (included bodies)?
' 2. Can we delete the old DerivedPartComponent?
' 3. Can we create a new derivation with specific body selection?
' 4. Does the new derivation update from the new base?
'
' RUN: Open a DERIVED part file, then run this rule
' REQUIRES: At least one other part file that could serve as alternative base

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("Test15: Open a derived part document first")
        MessageBox.Show("Ava esmalt tuletatud detaili fail (.ipt)", "Test15")
        Return
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    Logger.Info("=== Test15_DerivedPartReplace: Starting ===")
    Logger.Info("Part: " & partDoc.DisplayName)
    Logger.Info("Full path: " & partDoc.FullFileName)
    Logger.Info("")
    
    ' === TEST 1: Check if this is a derived part ===
    Logger.Info("--- TEST 1: Verify derived part ---")
    
    Dim dpcs As DerivedPartComponents = compDef.ReferenceComponents.DerivedPartComponents
    Logger.Info("DerivedPartComponents count: " & dpcs.Count)
    
    If dpcs.Count = 0 Then
        Logger.Warn("This part has NO derivation - cannot test replacement")
        Logger.Info("Open a part that was created via derivation")
        Return
    End If
    
    ' Get the first DerivedPartComponent
    Dim dpc As DerivedPartComponent = dpcs.Item(1)
    Dim originalBase As String = ""
    
    Try
        originalBase = dpc.ReferencedFile.FullFileName
        Logger.Info("Current base part: " & originalBase)
    Catch ex As Exception
        Logger.Warn("Could not get current base path: " & ex.Message)
        originalBase = "(unresolved)"
    End Try
    
    Logger.Info("DerivedPartComponent name: " & dpc.Name)
    
    ' === TEST 2: Record current derivation configuration ===
    Logger.Info("")
    Logger.Info("--- TEST 2: Record derivation configuration ---")
    
    Dim includedBodyNames As New List(Of String)
    Dim deriveStyle As DerivedComponentStyleEnum = DerivedComponentStyleEnum.kDeriveAsSingleBodyWithSeams
    
    Try
        deriveStyle = dpc.DeriveStyle
        Logger.Info("DeriveStyle: " & deriveStyle.ToString())
    Catch ex As Exception
        Logger.Warn("Could not get DeriveStyle: " & ex.Message)
    End Try
    
    ' Try to get included entities (bodies)
    Try
        Dim includedEntities As ObjectCollection = dpc.IncludedEntities
        Logger.Info("IncludedEntities count: " & includedEntities.Count)
        
        For i As Integer = 1 To includedEntities.Count
            Dim entity As Object = includedEntities.Item(i)
            Dim bodyName As String = "(unknown)"
            
            If TypeOf entity Is SurfaceBody Then
                bodyName = CType(entity, SurfaceBody).Name
                includedBodyNames.Add(bodyName)
            End If
            
            Logger.Info("  Included entity " & i & ": " & entity.GetType().Name & " - " & bodyName)
        Next
    Catch ex As Exception
        Logger.Warn("Could not enumerate IncludedEntities: " & ex.Message)
    End Try
    
    ' Count current bodies
    Dim bodyCountBefore As Integer = 0
    For Each body As SurfaceBody In compDef.SurfaceBodies
        If body.IsSolid Then bodyCountBefore += 1
    Next
    Logger.Info("Current solid body count: " & bodyCountBefore)
    
    ' === TEST 3: Find alternative base part ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Find alternative base part ---")
    
    Dim alternativeBase As String = FindAlternativeBasePart(app, partDoc.FullFileName, originalBase)
    
    If String.IsNullOrEmpty(alternativeBase) Then
        Logger.Warn("No alternative base part found")
        Logger.Info("Will test with user-selected file...")
        
        Dim fileDlg As Inventor.FileDialog = Nothing
        app.CreateFileDialog(fileDlg)
        fileDlg.Filter = "Inventor Parts (*.ipt)|*.ipt"
        fileDlg.FilterIndex = 1
        fileDlg.DialogTitle = "Select alternative base part"
        
        If Not String.IsNullOrEmpty(originalBase) AndAlso System.IO.File.Exists(originalBase) Then
            fileDlg.InitialDirectory = System.IO.Path.GetDirectoryName(originalBase)
        End If
        
        Try
            fileDlg.ShowOpen()
            If String.IsNullOrEmpty(fileDlg.FileName) Then
                Logger.Info("User cancelled file selection")
                Return
            End If
            alternativeBase = fileDlg.FileName
        Catch
            Logger.Info("File dialog cancelled")
            Return
        End Try
    End If
    
    Logger.Info("Alternative base part: " & alternativeBase)
    
    If alternativeBase.Equals(originalBase, StringComparison.OrdinalIgnoreCase) Then
        Logger.Warn("Alternative is same as original - test would be meaningless")
        Return
    End If
    
    If Not System.IO.File.Exists(alternativeBase) Then
        Logger.Error("Alternative base file not found: " & alternativeBase)
        Return
    End If
    
    ' === TEST 4: Inspect alternative base bodies ===
    Logger.Info("")
    Logger.Info("--- TEST 4: Inspect alternative base bodies ---")
    
    Dim altBodyNames As List(Of String) = GetBodyNamesFromPart(app, alternativeBase)
    Logger.Info("Alternative base has " & altBodyNames.Count & " solid bodies:")
    For Each bn As String In altBodyNames
        Logger.Info("  - " & bn)
    Next
    
    ' Find matching body name (if original had specific body)
    Dim targetBodyName As String = ""
    If includedBodyNames.Count > 0 Then
        targetBodyName = includedBodyNames(0)
        Logger.Info("Original derived body name: " & targetBodyName)
        
        If altBodyNames.Contains(targetBodyName) Then
            Logger.Info("Matching body found in alternative base!")
        Else
            Logger.Warn("No matching body name in alternative - will use first body")
            If altBodyNames.Count > 0 Then
                targetBodyName = altBodyNames(0)
            End If
        End If
    Else
        ' Default to first body
        If altBodyNames.Count > 0 Then
            targetBodyName = altBodyNames(0)
        End If
    End If
    
    Logger.Info("Target body for new derivation: " & targetBodyName)
    
    ' === TEST 5: Choose test method ===
    Logger.Info("")
    Logger.Info("--- TEST 5: Choose replacement method ---")
    
    Dim methodChoice As DialogResult = MessageBox.Show(
        "Choose replacement method:" & vbCrLf & vbCrLf &
        "YES = Delete + Recreate derivation (preserves body selection)" & vbCrLf &
        "NO = Use Replace() method (brings ALL bodies)" & vbCrLf &
        "CANCEL = Abort test" & vbCrLf & vbCrLf &
        "Current body count: " & bodyCountBefore & vbCrLf &
        "Alternative base body count: " & altBodyNames.Count & vbCrLf &
        "Target body: " & targetBodyName,
        "Test15_DerivedPartReplace",
        MessageBoxButtons.YesNoCancel)
    
    If methodChoice = DialogResult.Cancel Then
        Logger.Info("User cancelled")
        Return
    End If
    
    Dim fpBefore As String = ComputePartFingerprint(partDoc)
    Logger.Info("Fingerprint before: " & fpBefore)
    
    Dim success As Boolean = False
    Dim errorMsg As String = ""
    
    If methodChoice = DialogResult.Yes Then
        ' Method 1: Delete old + Create new with body selection
        success = TestDeleteAndRecreate(app, partDoc, alternativeBase, targetBodyName, deriveStyle, errorMsg)
    Else
        ' Method 2: Use Replace() (brings all bodies)
        success = TestReplaceMethod(app, partDoc, dpc, alternativeBase, errorMsg)
    End If
    
    ' === TEST 6: Verify results ===
    Logger.Info("")
    Logger.Info("--- TEST 6: Verify results ---")
    
    If success Then
        partDoc.Update()
        
        Dim bodyCountAfter As Integer = 0
        For Each body As SurfaceBody In compDef.SurfaceBodies
            If body.IsSolid Then bodyCountAfter += 1
        Next
        Logger.Info("Body count after: " & bodyCountAfter)
        
        Dim fpAfter As String = ComputePartFingerprint(partDoc)
        Logger.Info("Fingerprint after: " & fpAfter)
    End If
    
    ' === SUMMARY ===
    Logger.Info("")
    Logger.Info("========================================")
    Logger.Info("TEST SUMMARY")
    Logger.Info("========================================")
    Logger.Info("Method: " & If(methodChoice = DialogResult.Yes, "Delete + Recreate", "Replace()"))
    Logger.Info("Original base: " & System.IO.Path.GetFileName(originalBase))
    Logger.Info("Alternative base: " & System.IO.Path.GetFileName(alternativeBase))
    Logger.Info("Target body: " & targetBodyName)
    Logger.Info("Success: " & If(success, "YES", "NO - " & errorMsg))
    Logger.Info("========================================")
    
    If success Then
        MessageBox.Show(
            "Derivation replacement SUCCESS!" & vbCrLf & vbCrLf &
            "Method: " & If(methodChoice = DialogResult.Yes, "Delete + Recreate", "Replace()") & vbCrLf &
            "New base: " & System.IO.Path.GetFileName(alternativeBase) & vbCrLf & vbCrLf &
            "Document is dirty - save to keep or Ctrl+Z to undo.",
            "Test15 - SUCCESS")
    Else
        MessageBox.Show(
            "Derivation replacement FAILED:" & vbCrLf & vbCrLf &
            errorMsg & vbCrLf & vbCrLf &
            "Check log for details.",
            "Test15 - FAILED")
    End If
End Sub

Function TestDeleteAndRecreate(app As Inventor.Application, partDoc As PartDocument, _
                                newBasePath As String, targetBodyName As String, _
                                deriveStyle As DerivedComponentStyleEnum, _
                                ByRef errorMsg As String) As Boolean
    Logger.Info("")
    Logger.Info("=== Method: Delete + Recreate with body selection ===")
    
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim dpcs As DerivedPartComponents = compDef.ReferenceComponents.DerivedPartComponents
    
    ' Step 1: Delete existing derivation
    Logger.Info("Step 1: Deleting existing DerivedPartComponent...")
    
    Try
        Dim dpcToDelete As DerivedPartComponent = dpcs.Item(1)
        dpcToDelete.Delete()
        Logger.Info("Deleted existing derivation")
    Catch ex As Exception
        errorMsg = "Failed to delete existing derivation: " & ex.Message
        Logger.Error(errorMsg)
        Return False
    End Try
    
    ' Step 2: Create new derivation definition
    Logger.Info("Step 2: Creating new DerivedPartUniformScaleDef...")
    
    Try
        Dim dpDef As DerivedPartUniformScaleDef = dpcs.CreateUniformScaleDef(newBasePath)
        Logger.Info("Definition created, solids count: " & dpDef.Solids.Count)
        
        ' Step 3: Configure body selection
        Logger.Info("Step 3: Configuring body selection...")
        
        Dim includedCount As Integer = 0
        Dim excludedCount As Integer = 0
        
        For Each dpe As DerivedPartEntity In dpDef.Solids
            Dim bodyName As String = GetBodyNameFromEntity(dpe)
            
            If bodyName = targetBodyName Then
                dpe.IncludeEntity = True
                includedCount += 1
                Logger.Info("  Including: " & bodyName)
            Else
                dpe.IncludeEntity = False
                excludedCount += 1
                Logger.Info("  Excluding: " & bodyName)
            End If
        Next
        
        Logger.Info("Included: " & includedCount & ", Excluded: " & excludedCount)
        
        If includedCount = 0 Then
            ' Fallback: include first body
            Logger.Warn("No matching body found, including first body...")
            For Each dpe As DerivedPartEntity In dpDef.Solids
                dpe.IncludeEntity = True
                Exit For
            Next
        End If
        
        ' Exclude sketches, work features, parameters
        ExcludeNonSolidEntities(dpDef)
        
        ' Set derive style
        dpDef.DeriveStyle = deriveStyle
        Logger.Info("DeriveStyle: " & deriveStyle.ToString())
        
        ' Step 4: Add the derivation
        Logger.Info("Step 4: Adding new derivation...")
        
        Dim newDpc As DerivedPartComponent = dpcs.Add(dpDef)
        Logger.Info("New derivation added: " & newDpc.Name)
        
        Return True
        
    Catch ex As Exception
        errorMsg = "Failed to create new derivation: " & ex.Message
        Logger.Error(errorMsg)
        Return False
    End Try
End Function

Function TestReplaceMethod(app As Inventor.Application, partDoc As PartDocument, _
                           dpc As DerivedPartComponent, newBasePath As String, _
                           ByRef errorMsg As String) As Boolean
    Logger.Info("")
    Logger.Info("=== Method: Replace() (brings all bodies) ===")
    
    Try
        Dim options As NameValueMap = app.TransientObjects.CreateNameValueMap()
        
        Dim altModelState As String = GetPrimaryModelState(app, newBasePath)
        If Not String.IsNullOrEmpty(altModelState) Then
            Logger.Info("Using model state: " & altModelState)
            options.Add("ModelState", altModelState)
        End If
        
        dpc.Replace(newBasePath, options)
        Logger.Info("Replace: SUCCESS!")
        
        Return True
        
    Catch ex As Exception
        errorMsg = "Replace failed: " & ex.Message
        Logger.Error(errorMsg)
        Return False
    End Try
End Function

Sub ExcludeNonSolidEntities(dpDef As DerivedPartUniformScaleDef)
    ' Exclude all non-solid entities (sketches, work features, parameters, surfaces)
    Try
        For Each dpe As DerivedPartEntity In dpDef.Sketches3D : dpe.IncludeEntity = False : Next
    Catch : End Try
    Try
        For Each dpe As DerivedPartEntity In dpDef.Sketches : dpe.IncludeEntity = False : Next
    Catch : End Try
    Try
        For Each dpe As DerivedPartEntity In dpDef.WorkFeatures : dpe.IncludeEntity = False : Next
    Catch : End Try
    Try
        For Each dpe As DerivedPartEntity In dpDef.Surfaces : dpe.IncludeEntity = False : Next
    Catch : End Try
    Try
        For Each dpe As DerivedPartEntity In dpDef.Parameters : dpe.IncludeEntity = False : Next
    Catch : End Try
End Sub

Function GetBodyNameFromEntity(dpe As DerivedPartEntity) As String
    Try
        Dim refEntity As Object = dpe.ReferencedEntity
        If TypeOf refEntity Is SurfaceBody Then
            Return CType(refEntity, SurfaceBody).Name
        End If
    Catch : End Try
    Return "(unknown)"
End Function

Function GetBodyNamesFromPart(app As Inventor.Application, partPath As String) As List(Of String)
    Dim names As New List(Of String)
    
    Try
        Dim tempDoc As PartDocument = CType(app.Documents.Open(partPath, False), PartDocument)
        
        For Each body As SurfaceBody In tempDoc.ComponentDefinition.SurfaceBodies
            If body.IsSolid Then
                names.Add(body.Name)
            End If
        Next
        
        tempDoc.Close(True)
    Catch ex As Exception
        Logger.Warn("Could not get body names from " & partPath & ": " & ex.Message)
    End Try
    
    Return names
End Function

Function FindAlternativeBasePart(app As Inventor.Application, currentPart As String, _
                                  currentBase As String) As String
    ' Try to find another .ipt file in the same folder as the current base
    If String.IsNullOrEmpty(currentBase) OrElse Not System.IO.File.Exists(currentBase) Then
        Return ""
    End If
    
    Dim baseFolder As String = System.IO.Path.GetDirectoryName(currentBase)
    
    Try
        Dim iptFiles() As String = System.IO.Directory.GetFiles(baseFolder, "*.ipt")
        
        For Each iptFile As String In iptFiles
            ' Skip current base and current part
            If iptFile.Equals(currentBase, StringComparison.OrdinalIgnoreCase) Then Continue For
            If iptFile.Equals(currentPart, StringComparison.OrdinalIgnoreCase) Then Continue For
            
            ' Return first alternative found
            Logger.Info("Found alternative: " & iptFile)
            Return iptFile
        Next
    Catch ex As Exception
        Logger.Warn("Error searching for alternatives: " & ex.Message)
    End Try
    
    Return ""
End Function

Function GetPrimaryModelState(app As Inventor.Application, partPath As String) As String
    ' Get the primary/active model state from a part file
    Try
        ' Open the part silently
        Dim altDoc As PartDocument = CType(app.Documents.Open(partPath, False), PartDocument)
        
        ' Get model states
        Dim modelStates As Object = altDoc.ComponentDefinition.ModelStates
        Dim primaryState As String = ""
        
        For Each ms As Object In modelStates
            If ms.Name.Contains("[Primary]") OrElse ms.Name = "Primary" Then
                primaryState = ms.Name
                Exit For
            End If
        Next
        
        ' If no explicit primary, get first one
        If String.IsNullOrEmpty(primaryState) Then
            For Each ms As Object In modelStates
                primaryState = ms.Name
                Exit For
            Next
        End If
        
        altDoc.Close(True)  ' Close without saving
        
        Return primaryState
    Catch ex As Exception
        Logger.Warn("Could not get model state from " & partPath & ": " & ex.Message)
        Return ""
    End Try
End Function

Function ComputePartFingerprint(partDoc As PartDocument) As String
    Try
        Dim bodyFps As New List(Of String)
        
        For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
            If body.IsSolid Then
                Dim tol As Double = 0.001
                Dim vol As Double = 0
                Dim area As Double = 0
                Dim faceCount As Integer = 0
                
                Try : vol = Math.Round(body.Volume(tol), 6) : Catch : End Try
                Try
                    For Each face As Face In body.Faces
                        area += face.Evaluator.Area
                    Next
                    area = Math.Round(area, 6)
                    faceCount = body.Faces.Count
                Catch : End Try
                
                Dim bb As Box = body.RangeBox
                Dim dims() As Double = {
                    Math.Round(bb.MaxPoint.X - bb.MinPoint.X, 4),
                    Math.Round(bb.MaxPoint.Y - bb.MinPoint.Y, 4),
                    Math.Round(bb.MaxPoint.Z - bb.MinPoint.Z, 4)
                }
                Array.Sort(dims)
                
                Dim fp As String = String.Format("V:{0}|A:{1}|F:{2}|BB:{3}x{4}x{5}",
                    vol.ToString("F6", System.Globalization.CultureInfo.InvariantCulture),
                    area.ToString("F6", System.Globalization.CultureInfo.InvariantCulture),
                    faceCount,
                    dims(0).ToString("F4", System.Globalization.CultureInfo.InvariantCulture),
                    dims(1).ToString("F4", System.Globalization.CultureInfo.InvariantCulture),
                    dims(2).ToString("F4", System.Globalization.CultureInfo.InvariantCulture))
                bodyFps.Add(fp)
            End If
        Next
        
        If bodyFps.Count = 0 Then
            Return "NO_BODIES"
        End If
        
        bodyFps.Sort()
        Return String.Join("|", bodyFps.ToArray())
    Catch ex As Exception
        Return "ERROR:" & ex.Message
    End Try
End Function
