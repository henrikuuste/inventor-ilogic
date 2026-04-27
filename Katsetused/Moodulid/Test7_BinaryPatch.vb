' Copyright (c) 2026 Henri Kuuste
' Test7_BinaryPatch.vb
' PURPOSE: Validate binary patching for drawing reference updates
' 
' Since PutLogicalFileName doesn't exist in iLogic, we must use binary patching.
' This test validates the approach used in VariantReleaseLib/BinaryReferenceUpdateLib.
'
' TESTS:
' 1. Can we read references from a drawing?
' 2. Can we binary patch the closed file?
' 3. Does the drawing open correctly with new reference?
' 4. Does the drawing update/display correctly?
'
' RUN: Open a drawing file (.idw), then run this rule
' NOTE: The drawing will be CLOSED, patched, and reopened

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/BinaryReferenceUpdateLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        Logger.Error("Test7_BinaryPatch: Open a drawing document first")
        MessageBox.Show("Ava esmalt joonise fail (.idw)", "Test7")
        Return
    End If
    
    Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
    Dim drawPath As String = drawDoc.FullFileName
    
    Logger.Info("=== Test7_BinaryPatch: Starting ===")
    Logger.Info("Drawing: " & drawDoc.DisplayName)
    Logger.Info("Full path: " & drawPath)
    Logger.Info("")
    
    ' === TEST 1: Find a reference to patch ===
    Logger.Info("--- TEST 1: Find reference to patch ---")
    
    Dim originalRefPath As String = ""
    For Each refDoc As Document In drawDoc.ReferencedDocuments
        If refDoc.FullFileName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
            originalRefPath = refDoc.FullFileName
            Exit For
        End If
    Next
    
    If String.IsNullOrEmpty(originalRefPath) Then
        Logger.Error("No .ipt reference found in drawing")
        Return
    End If
    
    Logger.Info("Original reference: " & originalRefPath)
    
    ' === TEST 2: Create length-matched copy path ===
    Logger.Info("")
    Logger.Info("--- TEST 2: Create length-matched copy path ---")
    
    ' For binary patching, new path must be <= old path length
    ' Strategy: Create copy in SAME folder with modified name (same length)
    Dim originalFolder As String = System.IO.Path.GetDirectoryName(originalRefPath)
    Dim originalName As String = System.IO.Path.GetFileNameWithoutExtension(originalRefPath)
    Dim originalLen As Integer = originalRefPath.Length
    
    Logger.Info("Original path: " & originalRefPath)
    Logger.Info("Original path length: " & originalLen)
    
    ' Create new filename with SAME length by replacing characters
    ' e.g., "00012.ipt" -> "00012X.ipt" won't work (longer)
    ' e.g., "00012.ipt" -> "X0012.ipt" works (same length)
    Dim newRefPath As String = ""
    
    ' Strategy 1: Replace first character with 'X' (same length)
    If originalName.Length >= 1 Then
        Dim newName As String = "X" & originalName.Substring(1)
        newRefPath = System.IO.Path.Combine(originalFolder, newName & ".ipt")
        
        ' Make sure it's not the same file
        If newRefPath.Equals(originalRefPath, StringComparison.OrdinalIgnoreCase) Then
            ' Try different prefix
            newName = "Y" & originalName.Substring(1)
            newRefPath = System.IO.Path.Combine(originalFolder, newName & ".ipt")
        End If
    End If
    
    Logger.Info("New reference path: " & newRefPath)
    Logger.Info("New path length: " & newRefPath.Length)
    
    If newRefPath.Length <> originalLen Then
        Logger.Warn("Path lengths don't match! Binary patching may fail.")
        Logger.Info("Difference: " & (newRefPath.Length - originalLen))
    Else
        Logger.Info("Path lengths MATCH - good for binary patching")
    End If
    
    ' === TEST 3: Create the test file ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Create test copy ---")
    
    ' Verify path lengths match
    If newRefPath.Length <> originalLen Then
        Logger.Error("Path length mismatch! Original: " & originalLen & ", New: " & newRefPath.Length)
        Logger.Error("Binary patching requires exact length match")
        Return
    End If
    
    Logger.Info("Path lengths match: " & originalLen & " characters")
    
    ' Copy the original part to the new path (same folder, different name)
    Try
        System.IO.File.Copy(originalRefPath, newRefPath, True)
        Logger.Info("Created test copy: " & newRefPath)
    Catch ex As Exception
        Logger.Error("Failed to create test copy: " & ex.Message)
        Return
    End Try
    
    ' === TEST 4: Confirm with user ===
    Dim confirmResult As DialogResult = MessageBox.Show(
        "This test will:" & vbCrLf & vbCrLf &
        "1. CLOSE the drawing (required for binary patching)" & vbCrLf &
        "2. Patch the file to use: " & System.IO.Path.GetFileName(newRefPath) & vbCrLf &
        "3. Reopen the drawing" & vbCrLf & vbCrLf &
        "Original will be backed up as .backup file." & vbCrLf & vbCrLf &
        "Continue?",
        "Test7_BinaryPatch",
        MessageBoxButtons.YesNo)
    
    If confirmResult <> DialogResult.Yes Then
        Logger.Info("User cancelled")
        Return
    End If
    
    ' === TEST 5: Close the drawing ===
    Logger.Info("")
    Logger.Info("--- TEST 5: Close drawing for patching ---")
    
    Dim wasDirty As Boolean = drawDoc.Dirty
    If wasDirty Then
        Logger.Warn("Drawing has unsaved changes - saving first")
        drawDoc.Save()
    End If
    
    drawDoc.Close(False)  ' Close without prompt (already saved if needed)
    Logger.Info("Drawing closed")
    
    ' === TEST 6: Binary patch ===
    Logger.Info("")
    Logger.Info("--- TEST 6: Binary patch ---")
    
    ' Build path map
    Dim pathMap As New Dictionary(Of String, String)
    pathMap.Add(originalRefPath, newRefPath)
    
    Dim logs As New List(Of String)
    Dim success As Boolean = BinaryReferenceUpdateLib.UpdateFileReferencesBinary(drawPath, pathMap, logs)
    
    For Each logMsg As String In logs
        Logger.Info(logMsg)
    Next
    
    If success Then
        Logger.Info("Binary patch: SUCCESS")
    Else
        Logger.Error("Binary patch: FAILED")
    End If
    
    ' === TEST 7: Reopen and verify ===
    Logger.Info("")
    Logger.Info("--- TEST 7: Reopen and verify ---")
    
    Dim reopenedDoc As DrawingDocument = Nothing
    Try
        reopenedDoc = CType(app.Documents.Open(drawPath, True), DrawingDocument)
        Logger.Info("Drawing reopened successfully")
    Catch ex As Exception
        Logger.Error("Failed to reopen: " & ex.Message)
        Return
    End Try
    
    ' Check what references it has now
    Logger.Info("Current references after patch:")
    Dim foundNewRef As Boolean = False
    For Each refDoc As Document In reopenedDoc.ReferencedDocuments
        Logger.Info("  " & refDoc.FullFileName)
        If refDoc.FullFileName.Equals(newRefPath, StringComparison.OrdinalIgnoreCase) Then
            foundNewRef = True
        End If
    Next
    
    ' === TEST 8: Check views ===
    Logger.Info("")
    Logger.Info("--- TEST 8: Check drawing views ---")
    
    Dim viewCount As Integer = 0
    Dim errorViews As Integer = 0
    For Each sheet As Sheet In reopenedDoc.Sheets
        For Each view As DrawingView In sheet.DrawingViews
            viewCount += 1
            Try
                ' Try to access view properties to check if it's valid
                Dim modelPath As String = ""
                If view.ReferencedDocumentDescriptor IsNot Nothing Then
                    modelPath = view.ReferencedDocumentDescriptor.FullDocumentName
                End If
                Logger.Info("  View: " & view.Name & " -> " & System.IO.Path.GetFileName(modelPath))
            Catch ex As Exception
                errorViews += 1
                Logger.Warn("  View " & view.Name & " has issues: " & ex.Message)
            End Try
        Next
    Next
    
    Logger.Info("Total views: " & viewCount & ", Error views: " & errorViews)
    
    ' === SUMMARY ===
    Logger.Info("")
    Logger.Info("========================================")
    Logger.Info("TEST SUMMARY")
    Logger.Info("========================================")
    Logger.Info("Binary patch executed: " & If(success, "SUCCESS", "FAILED"))
    Logger.Info("Drawing reopened: YES")
    Logger.Info("New reference found: " & If(foundNewRef, "YES", "NO"))
    Logger.Info("Views: " & viewCount & " total, " & errorViews & " with errors")
    Logger.Info("========================================")
    
    If success AndAlso foundNewRef AndAlso errorViews = 0 Then
        Logger.Info("OVERALL: Binary patching WORKS!")
        MessageBox.Show(
            "Binary patching WORKS!" & vbCrLf & vbCrLf &
            "- Drawing patched successfully" & vbCrLf &
            "- New reference: " & System.IO.Path.GetFileName(newRefPath) & vbCrLf &
            "- All views intact" & vbCrLf & vbCrLf &
            "Backup saved as: " & drawPath & ".backup" & vbCrLf & vbCrLf &
            "To restore original reference, either:" & vbCrLf &
            "- Rename backup to .idw" & vbCrLf &
            "- Or run test again with original path",
            "Test7_BinaryPatch - SUCCESS")
    Else
        MessageBox.Show(
            "Binary patching issues:" & vbCrLf & vbCrLf &
            "Patch executed: " & success.ToString() & vbCrLf &
            "New ref found: " & foundNewRef.ToString() & vbCrLf &
            "Error views: " & errorViews & vbCrLf & vbCrLf &
            "Check log for details.",
            "Test7_BinaryPatch - ISSUES")
    End If
End Sub
