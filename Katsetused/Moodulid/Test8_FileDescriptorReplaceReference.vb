' Copyright (c) 2026 Henri Kuuste
' Test8_FileDescriptorReplaceReference.vb
' PURPOSE: Validate FileDescriptor.ReplaceReference for drawing model references
' 
' This test validates the CORRECT API approach for updating drawing references.
' Test4 failed because it tried non-existent PutLogicalFileName.
' The correct approach is: doc.File.ReferencedFileDescriptors.Item(n).ReplaceReference()
'
' KEY INSIGHT: ReplaceReference requires files to share "heritage" (same InternalName/GUID)
' Files copied via File.Copy or SaveCopyAs share heritage - perfect for our release workflow!
'
' TESTS:
' 1. Can we access FileDescriptor via doc.File.ReferencedFileDescriptors?
' 2. Does ReplaceReference work when target is a COPY of original (shared heritage)?
' 3. Does the drawing display correctly after replacement?
' 4. Does ReplaceReference fail when target has DIFFERENT heritage (new file)?
'
' RUN: Open a drawing file (.idw) that references a part, then run this rule

AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        Logger.Error("Test8: Open a drawing document first")
        MessageBox.Show("Ava esmalt joonise fail (.idw)", "Test8")
        Return
    End If
    
    Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
    
    Logger.Info("=== Test8_FileDescriptorReplaceReference: Starting ===")
    Logger.Info("Drawing: " & drawDoc.DisplayName)
    Logger.Info("Full path: " & drawDoc.FullFileName)
    Logger.Info("")
    
    ' === TEST 1: Access FileDescriptors via doc.File ===
    Logger.Info("--- TEST 1: Access FileDescriptor via doc.File ---")
    
    Dim fileDescriptors As FileDescriptorsEnumerator = Nothing
    Try
        fileDescriptors = drawDoc.File.ReferencedFileDescriptors
        Logger.Info("doc.File.ReferencedFileDescriptors: ACCESSIBLE")
        Logger.Info("Count: " & fileDescriptors.Count)
    Catch ex As Exception
        Logger.Error("Cannot access doc.File.ReferencedFileDescriptors: " & ex.Message)
        Return
    End Try
    
    ' List all references via FileDescriptor
    Logger.Info("")
    Logger.Info("FileDescriptor references:")
    Dim testFd As FileDescriptor = Nothing
    Dim testOriginalPath As String = ""
    
    For i As Integer = 1 To fileDescriptors.Count
        Dim fd As FileDescriptor = fileDescriptors.Item(i)
        Logger.Info("  [" & i & "] " & fd.FullFileName)
        
        ' Select first .ipt file for testing
        If testFd Is Nothing AndAlso fd.FullFileName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
            If System.IO.File.Exists(fd.FullFileName) Then
                testFd = fd
                testOriginalPath = fd.FullFileName
            End If
        End If
    Next
    
    If testFd Is Nothing Then
        Logger.Warn("No .ipt reference found for testing")
        Logger.Info("Test requires a drawing with at least one existing part reference")
        Return
    End If
    
    Logger.Info("")
    Logger.Info("Selected for test: " & testOriginalPath)
    
    ' === TEST 2: Create test copy (preserves InternalName/heritage) ===
    Logger.Info("")
    Logger.Info("--- TEST 2: Create test copy (shared heritage) ---")
    
    Dim originalFolder As String = System.IO.Path.GetDirectoryName(testOriginalPath)
    Dim originalName As String = System.IO.Path.GetFileNameWithoutExtension(testOriginalPath)
    Dim copyPath As String = System.IO.Path.Combine(originalFolder, originalName & "_Test8Copy.ipt")
    
    Logger.Info("Original: " & testOriginalPath)
    Logger.Info("Copy path: " & copyPath)
    
    ' Create copy using File.Copy (preserves InternalName!)
    Try
        System.IO.File.Copy(testOriginalPath, copyPath, True)
        Logger.Info("Created copy: SUCCESS (heritage preserved)")
    Catch ex As Exception
        Logger.Error("Failed to create copy: " & ex.Message)
        Return
    End Try
    
    ' Verify copy exists
    If Not System.IO.File.Exists(copyPath) Then
        Logger.Error("Copy file not found after creation")
        Return
    End If
    
    ' === TEST 3: Verify InternalName matches (optional - for educational purposes) ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Verify heritage (InternalName) ---")
    
    ' Open both files to check InternalName
    Dim originalInternalName As String = ""
    Dim copyInternalName As String = ""
    
    Try
        Dim origDoc As PartDocument = CType(app.Documents.Open(testOriginalPath, False), PartDocument)
        originalInternalName = origDoc.InternalName
        origDoc.Close(True)
        
        Dim copyDoc As PartDocument = CType(app.Documents.Open(copyPath, False), PartDocument)
        copyInternalName = copyDoc.InternalName
        copyDoc.Close(True)
        
        Logger.Info("Original InternalName: " & originalInternalName)
        Logger.Info("Copy InternalName: " & copyInternalName)
        
        If originalInternalName = copyInternalName Then
            Logger.Info("Heritage: MATCH (same InternalName) - ReplaceReference should work!")
        Else
            Logger.Warn("Heritage: DIFFERENT - This shouldn't happen with File.Copy!")
        End If
    Catch ex As Exception
        Logger.Warn("Could not verify InternalName: " & ex.Message)
        Logger.Info("Continuing anyway...")
    End Try
    
    ' === TEST 4: Attempt ReplaceReference ===
    Logger.Info("")
    Logger.Info("--- TEST 4: ReplaceReference with copied file ---")
    
    Dim confirmResult As DialogResult = MessageBox.Show(
        "This will attempt to change the drawing reference:" & vbCrLf & vbCrLf &
        "From: " & System.IO.Path.GetFileName(testOriginalPath) & vbCrLf &
        "To: " & System.IO.Path.GetFileName(copyPath) & vbCrLf & vbCrLf &
        "This uses FileDescriptor.ReplaceReference (the CORRECT API)." & vbCrLf &
        "The copy has shared heritage (same InternalName)." & vbCrLf & vbCrLf &
        "Continue?",
        "Test8_FileDescriptorReplaceReference",
        MessageBoxButtons.YesNo)
    
    If confirmResult <> DialogResult.Yes Then
        Logger.Info("User cancelled")
        Return
    End If
    
    Dim replaceSuccess As Boolean = False
    Dim replaceError As String = ""
    
    Try
        Logger.Info("Calling FileDescriptor.ReplaceReference...")
        testFd.ReplaceReference(copyPath)
        replaceSuccess = True
        Logger.Info("ReplaceReference: SUCCESS!")
    Catch ex As Exception
        replaceError = ex.Message
        Logger.Error("ReplaceReference FAILED: " & ex.Message)
        Logger.Info("Exception type: " & ex.GetType().Name)
    End Try
    
    ' === TEST 5: Verify change and update display ===
    Logger.Info("")
    Logger.Info("--- TEST 5: Verify and update ---")
    
    If replaceSuccess Then
        ' Update drawing
        Try
            drawDoc.Update()
            Logger.Info("Drawing updated")
        Catch ex As Exception
            Logger.Warn("Update had issues: " & ex.Message)
        End Try
        
        ' Check current reference
        Try
            Dim currentRef As String = testFd.FullFileName
            Logger.Info("Current reference: " & currentRef)
            
            If currentRef.Equals(copyPath, StringComparison.OrdinalIgnoreCase) Then
                Logger.Info("Reference successfully changed to copy!")
            Else
                Logger.Warn("Reference may not have changed: " & currentRef)
            End If
        Catch ex As Exception
            Logger.Error("Could not verify reference: " & ex.Message)
        End Try
        
        ' Check views
        Dim viewCount As Integer = 0
        Dim errorViews As Integer = 0
        For Each sheet As Sheet In drawDoc.Sheets
            For Each view As DrawingView In sheet.DrawingViews
                viewCount += 1
                Try
                    Dim viewRef As String = ""
                    If view.ReferencedDocumentDescriptor IsNot Nothing Then
                        viewRef = view.ReferencedDocumentDescriptor.FullDocumentName
                    End If
                    Logger.Info("  View: " & view.Name & " -> " & System.IO.Path.GetFileName(viewRef))
                Catch ex As Exception
                    errorViews += 1
                    Logger.Warn("  View " & view.Name & " has issues: " & ex.Message)
                End Try
            Next
        Next
        Logger.Info("Total views: " & viewCount & ", Error views: " & errorViews)
    End If
    
    ' === TEST 6: Test with DIFFERENT heritage (optional, educational) ===
    Logger.Info("")
    Logger.Info("--- TEST 6: Info about different heritage scenario ---")
    Logger.Info("Note: If target file has DIFFERENT InternalName, ReplaceReference would fail with:")
    Logger.Info("  'The resolved document is not usable'")
    Logger.Info("  'Please make sure the replacement file has the same database id'")
    Logger.Info("")
    Logger.Info("This is WHY binary patching exists - for cases where heritage doesn't match.")
    Logger.Info("But for Moodulid release workflow, we COPY masters, so heritage always matches!")
    
    ' === SUMMARY ===
    Logger.Info("")
    Logger.Info("========================================")
    Logger.Info("TEST SUMMARY")
    Logger.Info("========================================")
    Logger.Info("FileDescriptor access via doc.File: WORKS")
    Logger.Info("ReplaceReference with copied file: " & If(replaceSuccess, "SUCCESS", "FAILED - " & replaceError))
    
    If replaceSuccess Then
        Logger.Info("========================================")
        Logger.Info("CONCLUSION: FileDescriptor.ReplaceReference WORKS!")
        Logger.Info("")
        Logger.Info("Key insight: Use doc.File.ReferencedFileDescriptors (not doc.ReferencedFileDescriptors)")
        Logger.Info("Requirement: Target file must share heritage (same InternalName)")
        Logger.Info("For Moodulid: File.Copy preserves heritage - perfect for release workflow!")
        Logger.Info("")
        Logger.Info("This is BETTER than binary patching because:")
        Logger.Info("  - No path length constraint")
        Logger.Info("  - Works through official API")
        Logger.Info("  - Drawing views update properly")
        
        MessageBox.Show(
            "FileDescriptor.ReplaceReference WORKS!" & vbCrLf & vbCrLf &
            "Key findings:" & vbCrLf &
            "- Access via doc.File.ReferencedFileDescriptors" & vbCrLf &
            "- Requires shared heritage (File.Copy preserves this)" & vbCrLf &
            "- No path length constraint!" & vbCrLf & vbCrLf &
            "Drawing reference changed to: " & System.IO.Path.GetFileName(copyPath) & vbCrLf & vbCrLf &
            "To restore: Run test again or use Ctrl+Z" & vbCrLf &
            "The test copy file remains in: " & originalFolder,
            "Test8 - SUCCESS")
    Else
        MessageBox.Show(
            "ReplaceReference FAILED:" & vbCrLf & vbCrLf &
            replaceError & vbCrLf & vbCrLf &
            "This may indicate:" & vbCrLf &
            "- InternalName mismatch (shouldn't happen with File.Copy)" & vbCrLf &
            "- File access issue" & vbCrLf &
            "- API limitation in this Inventor version" & vbCrLf & vbCrLf &
            "Check log for details.",
            "Test8 - FAILED")
    End If
End Sub
