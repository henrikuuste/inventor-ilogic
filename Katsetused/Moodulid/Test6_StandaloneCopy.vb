' Copyright (c) 2026 Henri Kuuste
' Test6_StandaloneCopy.vb
' PURPOSE: Test creating standalone part copies (alternative to BreakLinkToFile)
' 
' If BreakLinkToFile doesn't work, we need an alternative:
' 1. SaveAs/Copy the derived part to new location
' 2. Open the copy
' 3. Delete the DerivedPartComponent feature
' 4. Save - now it's standalone with no references
'
' TESTS:
' 1. Can we SaveAs a derived part?
' 2. Can we delete DerivedPartComponent from the copy?
' 3. Does geometry survive after deleting the derivation?
' 4. Is the copy truly standalone?
'
' RUN: Open a DERIVED part file, then run this rule

AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("Test6_StandaloneCopy: Open a derived part document first")
        MessageBox.Show("Ava esmalt tuletatud detaili fail (.ipt)", "Test6")
        Return
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    Logger.Info("=== Test6_StandaloneCopy: Starting ===")
    Logger.Info("Part: " & partDoc.DisplayName)
    Logger.Info("Full path: " & partDoc.FullFileName)
    Logger.Info("")
    
    ' === TEST 1: Check if this is a derived part ===
    Logger.Info("--- TEST 1: Verify derived part ---")
    
    Dim dpcs As DerivedPartComponents = compDef.ReferenceComponents.DerivedPartComponents
    Logger.Info("DerivedPartComponents count: " & dpcs.Count)
    
    If dpcs.Count = 0 Then
        Logger.Warn("This part has NO derivation - nothing to test")
        Logger.Info("Open a part created with 'Loo komponendid.vb'")
        Return
    End If
    
    ' Log source reference
    For Each dpc As DerivedPartComponent In dpcs
        Logger.Info("DerivedPartComponent: " & dpc.Name)
        Try
            Logger.Info("  Source: " & dpc.ReferencedFile.FullFileName)
        Catch
            Logger.Info("  Source: (could not read)")
        End Try
    Next
    
    ' === TEST 2: Fingerprint original ===
    Logger.Info("")
    Logger.Info("--- TEST 2: Fingerprint original ---")
    
    Dim fpOriginal As String = ComputePartFingerprint(partDoc)
    Logger.Info("Fingerprint: " & fpOriginal)
    
    ' === TEST 3: Prepare copy path ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Prepare copy location ---")
    
    Dim originalPath As String = partDoc.FullFileName
    Dim originalFolder As String = System.IO.Path.GetDirectoryName(originalPath)
    Dim originalName As String = System.IO.Path.GetFileNameWithoutExtension(originalPath)
    
    ' Create test folder
    Dim testFolder As String = System.IO.Path.Combine(originalFolder, "_StandaloneTest")
    If Not System.IO.Directory.Exists(testFolder) Then
        System.IO.Directory.CreateDirectory(testFolder)
        Logger.Info("Created test folder: " & testFolder)
    End If
    
    Dim copyPath As String = System.IO.Path.Combine(testFolder, originalName & "_Standalone.ipt")
    Logger.Info("Copy path: " & copyPath)
    
    ' Delete existing copy if present
    If System.IO.File.Exists(copyPath) Then
        Try
            System.IO.File.Delete(copyPath)
            Logger.Info("Deleted existing copy")
        Catch ex As Exception
            Logger.Error("Could not delete existing copy: " & ex.Message)
            Return
        End Try
    End If
    
    ' === TEST 4: SaveAs copy ===
    Logger.Info("")
    Logger.Info("--- TEST 4: Create copy via SaveAs ---")
    
    Dim confirmResult As DialogResult = MessageBox.Show(
        "This will create a standalone copy:" & vbCrLf & vbCrLf &
        "Source: " & originalPath & vbCrLf &
        "Copy: " & copyPath & vbCrLf & vbCrLf &
        "The copy will have derivation removed." & vbCrLf & vbCrLf &
        "Continue?",
        "Test6_StandaloneCopy",
        MessageBoxButtons.YesNo)
    
    If confirmResult <> DialogResult.Yes Then
        Logger.Info("User cancelled")
        Return
    End If
    
    ' Method 1: SaveAs with saveCopyAs=True
    Try
        partDoc.SaveAs(copyPath, True)  ' True = save copy (original stays open)
        Logger.Info("SaveAs completed")
    Catch ex As Exception
        Logger.Error("SaveAs failed: " & ex.Message)
        Return
    End Try
    
    ' Verify copy exists
    If Not System.IO.File.Exists(copyPath) Then
        Logger.Error("Copy file does not exist after SaveAs!")
        Return
    End If
    Logger.Info("Copy file exists: " & copyPath)
    
    ' === TEST 5: Open the copy ===
    Logger.Info("")
    Logger.Info("--- TEST 5: Open copy and remove derivation ---")
    
    Dim copyDoc As PartDocument = Nothing
    Try
        copyDoc = CType(app.Documents.Open(copyPath, True), PartDocument)  ' True = visible
        Logger.Info("Copy opened successfully")
    Catch ex As Exception
        Logger.Error("Failed to open copy: " & ex.Message)
        Return
    End Try
    
    ' Check derivation in copy
    Dim copyDpcs As DerivedPartComponents = copyDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
    Logger.Info("Copy has " & copyDpcs.Count & " DerivedPartComponents")
    
    ' === TEST 6: Delete DerivedPartComponents ===
    Logger.Info("")
    Logger.Info("--- TEST 6: Delete derivation features ---")
    
    Dim deleteSuccess As Boolean = True
    Dim deleteError As String = ""
    
    ' Collect to list first (can't modify while iterating)
    Dim dpcList As New List(Of DerivedPartComponent)
    For Each dpc As DerivedPartComponent In copyDpcs
        dpcList.Add(dpc)
    Next
    
    For Each dpc As DerivedPartComponent In dpcList
        Try
            Logger.Info("Deleting: " & dpc.Name)
            dpc.Delete()
            Logger.Info("  Deleted successfully")
        Catch ex As Exception
            Logger.Error("  Delete failed: " & ex.Message)
            deleteSuccess = False
            deleteError = ex.Message
        End Try
    Next
    
    ' Update after delete
    Try
        copyDoc.Update()
        Logger.Info("Copy updated after delete")
    Catch ex As Exception
        Logger.Warn("Update had issues: " & ex.Message)
    End Try
    
    ' === TEST 7: Check geometry survived ===
    Logger.Info("")
    Logger.Info("--- TEST 7: Check geometry after delete ---")
    
    Dim fpCopy As String = ComputePartFingerprint(copyDoc)
    Logger.Info("Copy fingerprint: " & fpCopy)
    
    Dim geometryPreserved As Boolean = (fpOriginal = fpCopy)
    If geometryPreserved Then
        Logger.Info("PASS: Geometry PRESERVED after deleting derivation")
    Else
        Logger.Error("FAIL: Geometry DIFFERS after deleting derivation!")
        Logger.Info("  Original: " & fpOriginal)
        Logger.Info("  Copy:     " & fpCopy)
        
        ' Check if copy has any bodies
        Dim bodyCount As Integer = 0
        For Each body As SurfaceBody In copyDoc.ComponentDefinition.SurfaceBodies
            If body.IsSolid Then bodyCount += 1
        Next
        Logger.Info("  Copy has " & bodyCount & " solid bodies")
        
        If bodyCount = 0 Then
            Logger.Error("  CRITICAL: Deleting derivation removed all geometry!")
            Logger.Info("  Need to use BreakLinkToFile or BaseFeature.Adaptive=False instead")
        End If
    End If
    
    ' === TEST 8: Check references ===
    Logger.Info("")
    Logger.Info("--- TEST 8: Check copy references ---")
    
    Dim refCount As Integer = copyDoc.ReferencedDocuments.Count
    Logger.Info("Copy ReferencedDocuments: " & refCount)
    
    For Each refDoc As Document In copyDoc.ReferencedDocuments
        Logger.Info("  Still references: " & refDoc.FullFileName)
    Next
    
    Dim copyDpcsAfter As DerivedPartComponents = copyDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
    Logger.Info("Copy DerivedPartComponents after delete: " & copyDpcsAfter.Count)
    
    Dim isStandalone As Boolean = (refCount = 0 AndAlso copyDpcsAfter.Count = 0)
    
    ' === TEST 9: Save the standalone copy ===
    Logger.Info("")
    Logger.Info("--- TEST 9: Save standalone copy ---")
    
    If deleteSuccess AndAlso geometryPreserved Then
        Try
            copyDoc.Save()
            Logger.Info("Copy saved")
        Catch ex As Exception
            Logger.Error("Save failed: " & ex.Message)
        End Try
    Else
        Logger.Info("Not saving due to issues")
    End If
    
    ' === SUMMARY ===
    Logger.Info("")
    Logger.Info("========================================")
    Logger.Info("TEST SUMMARY")
    Logger.Info("========================================")
    Logger.Info("SaveAs copy: SUCCESS")
    Logger.Info("Delete derivation: " & If(deleteSuccess, "SUCCESS", "FAILED - " & deleteError))
    Logger.Info("Geometry preserved: " & If(geometryPreserved, "YES", "NO"))
    Logger.Info("References removed: " & If(refCount = 0, "YES", "NO (" & refCount & " remaining)"))
    Logger.Info("Is standalone: " & If(isStandalone AndAlso geometryPreserved, "YES", "NO"))
    Logger.Info("========================================")
    
    If deleteSuccess AndAlso geometryPreserved AndAlso isStandalone Then
        Logger.Info("OVERALL: SaveAs + Delete approach WORKS!")
        MessageBox.Show(
            "SaveAs + Delete approach WORKS!" & vbCrLf & vbCrLf &
            "- Copy created: " & copyPath & vbCrLf &
            "- Derivation deleted" & vbCrLf &
            "- Geometry preserved" & vbCrLf &
            "- No references remain" & vbCrLf & vbCrLf &
            "This is a viable alternative to BreakLinkToFile!",
            "Test6_StandaloneCopy - SUCCESS")
    ElseIf Not geometryPreserved Then
        Logger.Error("OVERALL: Delete approach REMOVES GEOMETRY")
        MessageBox.Show(
            "WARNING: Deleting derivation removes geometry!" & vbCrLf & vbCrLf &
            "The copy has " & (If(fpCopy.StartsWith("ERROR"), "no bodies", "different geometry")) & vbCrLf & vbCrLf &
            "Need to use BreakLinkToFile instead, or" & vbCrLf &
            "set BaseFeature.Adaptive = False" & vbCrLf & vbCrLf &
            "Copy left open for inspection.",
            "Test6_StandaloneCopy - FAILED")
    Else
        MessageBox.Show(
            "Partial success:" & vbCrLf & vbCrLf &
            "Delete: " & deleteSuccess.ToString() & vbCrLf &
            "Geometry: " & geometryPreserved.ToString() & vbCrLf &
            "Standalone: " & isStandalone.ToString() & vbCrLf & vbCrLf &
            "Check log for details. Copy left open.",
            "Test6_StandaloneCopy - PARTIAL")
    End If
    
    ' Ask if user wants to close the copy
    Dim closeResult As DialogResult = MessageBox.Show(
        "Close the test copy?",
        "Test6_StandaloneCopy",
        MessageBoxButtons.YesNo)
    
    If closeResult = DialogResult.Yes Then
        copyDoc.Close(True)  ' Close without saving changes
        Logger.Info("Copy closed")
    Else
        Logger.Info("Copy left open for inspection")
    End If
End Sub

' Compute fingerprint for whole part
' NOTE: SurfaceBody.SurfaceArea does NOT exist - use sum of face areas
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
