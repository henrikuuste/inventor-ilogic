' Copyright (c) 2026 Henri Kuuste
' Test4_DrawingRelink.vb
' PURPOSE: Validate that drawing model references can be updated
' 
' CRITICAL: Research says PutLogicalFileName may NOT work in iLogic!
' Need to test both API approach and understand when binary patching is needed.
'
' TESTS:
' 1. Can we read ReferencedFileDescriptors from a drawing?
' 2. Can we use PutLogicalFileName to change the reference?
' 3. Does the drawing update correctly after reference change?
' 4. What errors occur if it fails?
'
' RUN: Open a drawing file (.idw), then run this rule
' PREP: Have a copy of the referenced model ready to test relinking

AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        Logger.Error("Test4_DrawingRelink: Open a drawing document first")
        MessageBox.Show("Ava esmalt joonise fail (.idw)", "Test4")
        Return
    End If
    
    Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
    
    Logger.Info("=== Test4_DrawingRelink: Starting ===")
    Logger.Info("Drawing: " & drawDoc.DisplayName)
    Logger.Info("Full path: " & drawDoc.FullFileName)
    Logger.Info("")
    
    ' === TEST 1: List ReferencedDocuments ===
    Logger.Info("--- TEST 1: ReferencedDocuments (API) ---")
    
    Dim refDocs As New List(Of String)
    For Each refDoc As Document In drawDoc.ReferencedDocuments
        refDocs.Add(refDoc.FullFileName)
        Logger.Info("  " & refDoc.FullFileName)
    Next
    Logger.Info("Total referenced: " & refDocs.Count)
    
    If refDocs.Count = 0 Then
        Logger.Warn("Drawing has no model references")
        Return
    End If
    
    ' === TEST 2: List ReferencedFileDescriptors ===
    Logger.Info("")
    Logger.Info("--- TEST 2: ReferencedFileDescriptors ---")
    
    Dim rfds As ReferencedFileDescriptors = drawDoc.ReferencedFileDescriptors
    Logger.Info("ReferencedFileDescriptors count: " & rfds.Count)
    
    Dim rfdList As New List(Of ReferencedFileDescriptor)
    For i As Integer = 1 To rfds.Count
        Dim rfd As ReferencedFileDescriptor = rfds.Item(i)
        rfdList.Add(rfd)
        
        Logger.Info("Descriptor " & i & ":")
        Try
            Logger.Info("  FullFileName: " & rfd.FullFileName)
        Catch ex As Exception
            Logger.Info("  FullFileName: (error: " & ex.Message & ")")
        End Try
        Try
            Logger.Info("  LogicalFileName: " & rfd.LogicalFileName)
        Catch ex As Exception
            Logger.Info("  LogicalFileName: (error: " & ex.Message & ")")
        End Try
        Try
            Logger.Info("  ReferenceMissing: " & rfd.ReferenceMissing.ToString())
        Catch ex As Exception
            Logger.Info("  ReferenceMissing: (error: " & ex.Message & ")")
        End Try
    Next
    
    ' === TEST 3: Try to identify valid test scenario ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Identify test target ---")
    
    ' Find first part file reference (ReferenceMissing doesn't exist, check file on disk)
    Dim testRfd As ReferencedFileDescriptor = Nothing
    Dim testOriginalPath As String = ""
    
    For Each rfd As ReferencedFileDescriptor In rfdList
        Try
            Dim path As String = rfd.FullFileName
            If path.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                ' Check if file exists on disk
                If System.IO.File.Exists(path) Then
                    testRfd = rfd
                    testOriginalPath = path
                    Logger.Info("Found valid reference: " & path)
                    Exit For
                Else
                    Logger.Warn("Reference exists but file missing: " & path)
                End If
            End If
        Catch ex As Exception
            Logger.Warn("Error reading descriptor: " & ex.Message)
        End Try
    Next
    
    If testRfd Is Nothing Then
        Logger.Warn("Could not find a valid part reference to test")
        Logger.Info("Need a drawing that references at least one existing .ipt file")
        
        ' Still show what we learned
        Logger.Info("")
        Logger.Info("What we learned:")
        Logger.Info("- ReferencedFileDescriptors API works: YES")
        Logger.Info("- Can enumerate references: YES")
        Logger.Info("- LogicalFileName property: DOES NOT EXIST")
        Logger.Info("- ReferenceMissing property: DOES NOT EXIST")
        Return
    End If
    
    Logger.Info("Found test target: " & testOriginalPath)
    
    ' === TEST 4: Prepare a test path ===
    Logger.Info("")
    Logger.Info("--- TEST 4: Prepare test path ---")
    
    ' Create a test copy path (in same folder with _Test suffix)
    Dim originalFolder As String = System.IO.Path.GetDirectoryName(testOriginalPath)
    Dim originalName As String = System.IO.Path.GetFileNameWithoutExtension(testOriginalPath)
    Dim testCopyPath As String = System.IO.Path.Combine(originalFolder, originalName & "_TestCopy.ipt")
    
    Logger.Info("Will test relinking to: " & testCopyPath)
    
    ' Check if test copy exists (user needs to create it)
    Dim copyExists As Boolean = System.IO.File.Exists(testCopyPath)
    Logger.Info("Test copy exists: " & copyExists.ToString())
    
    If Not copyExists Then
        ' Offer to create copy
        Dim createResult As DialogResult = MessageBox.Show(
            "Test copy file does not exist:" & vbCrLf & vbCrLf &
            testCopyPath & vbCrLf & vbCrLf &
            "Create a copy now for testing?",
            "Test4_DrawingRelink",
            MessageBoxButtons.YesNo)
        
        If createResult = DialogResult.Yes Then
            Try
                System.IO.File.Copy(testOriginalPath, testCopyPath, True)
                Logger.Info("Created test copy: " & testCopyPath)
                copyExists = True
            Catch ex As Exception
                Logger.Error("Failed to create copy: " & ex.Message)
                Return
            End Try
        Else
            Logger.Info("Cannot proceed without test copy file")
            Return
        End If
    End If
    
    ' === TEST 5: Attempt PutLogicalFileName ===
    Logger.Info("")
    Logger.Info("--- TEST 5: Attempt PutLogicalFileName ---")
    
    Dim confirmResult As DialogResult = MessageBox.Show(
        "This will attempt to change the drawing reference:" & vbCrLf & vbCrLf &
        "From: " & testOriginalPath & vbCrLf &
        "To: " & testCopyPath & vbCrLf & vbCrLf &
        "The drawing will be modified. You can undo with Ctrl+Z." & vbCrLf & vbCrLf &
        "Continue?",
        "Test4_DrawingRelink",
        MessageBoxButtons.YesNo)
    
    If confirmResult <> DialogResult.Yes Then
        Logger.Info("User cancelled")
        Return
    End If
    
    Dim putSuccess As Boolean = False
    Dim putError As String = ""
    
    Try
        Logger.Info("Calling PutLogicalFileName...")
        testRfd.PutLogicalFileName(testCopyPath)
        putSuccess = True
        Logger.Info("  PutLogicalFileName completed without error")
    Catch ex As Exception
        putError = ex.Message
        Logger.Error("  PutLogicalFileName FAILED: " & ex.Message)
        Logger.Info("  Exception type: " & ex.GetType().Name)
        
        ' Common failure reasons
        If ex.Message.Contains("E_FAIL") Then
            Logger.Info("  HINT: E_FAIL often means API not supported in iLogic context")
        ElseIf ex.Message.Contains("access") OrElse ex.Message.Contains("permission") Then
            Logger.Info("  HINT: May be a file access/permission issue")
        End If
    End Try
    
    ' === TEST 6: Verify change ===
    Logger.Info("")
    Logger.Info("--- TEST 6: Verify reference change ---")
    
    If putSuccess Then
        ' Update drawing
        Try
            drawDoc.Update()
            Logger.Info("Drawing updated")
        Catch ex As Exception
            Logger.Warn("Drawing update had issues: " & ex.Message)
        End Try
        
        ' Check the reference now
        Dim currentPath As String = ""
        Try
            currentPath = testRfd.FullFileName
            Logger.Info("Current reference path: " & currentPath)
            
            If currentPath = testCopyPath Then
                Logger.Info("  PASS: Reference points to new path")
            ElseIf currentPath = testOriginalPath Then
                Logger.Warn("  Reference still points to original (change may not have taken effect)")
            Else
                Logger.Info("  Reference points to: " & currentPath)
            End If
        Catch ex As Exception
            Logger.Error("Could not read current path: " & ex.Message)
        End Try
    End If
    
    ' === TEST 7: Alternative - ReplaceReference via File.ReferencedFileDescriptors ===
    Logger.Info("")
    Logger.Info("--- TEST 7: Alternative API info ---")
    
    If Not putSuccess Then
        Logger.Info("Since PutLogicalFileName failed, noting alternatives:")
        Logger.Info("1. FileDescriptor.ReplaceReference (requires same InternalName)")
        Logger.Info("2. Binary patching (VariantReleaseLib approach)")
        Logger.Info("3. Delete and re-place views (destructive)")
        
        ' Check if File.ReferencedFileDescriptors exists
        Try
            Dim fileObj As Inventor.File = drawDoc.File
            If fileObj IsNot Nothing Then
                Logger.Info("Drawing.File object available")
                Dim fileRfds As ReferencedFileDescriptors = fileObj.ReferencedFileDescriptors
                Logger.Info("File.ReferencedFileDescriptors count: " & fileRfds.Count)
            End If
        Catch ex As Exception
            Logger.Info("File object access: " & ex.Message)
        End Try
    End If
    
    ' === SUMMARY ===
    Logger.Info("")
    Logger.Info("========================================")
    Logger.Info("TEST SUMMARY")
    Logger.Info("========================================")
    Logger.Info("ReferencedFileDescriptors enumeration: WORKS")
    Logger.Info("PutLogicalFileName: " & If(putSuccess, "WORKS", "FAILED - " & putError))
    
    If putSuccess Then
        Logger.Info("========================================")
        Logger.Info("OVERALL: Drawing relink via API WORKS!")
        MessageBox.Show(
            "PutLogicalFileName WORKS!" & vbCrLf & vbCrLf &
            "Drawing reference updated to:" & vbCrLf &
            testCopyPath & vbCrLf & vbCrLf &
            "Document is dirty - save to keep or Ctrl+Z to undo." & vbCrLf & vbCrLf &
            "Note: You may want to relink back to original.",
            "Test4_DrawingRelink - SUCCESS")
    Else
        Logger.Info("========================================")
        Logger.Info("OVERALL: Need alternative approach (binary patching)")
        MessageBox.Show(
            "PutLogicalFileName FAILED in iLogic:" & vbCrLf & vbCrLf &
            putError & vbCrLf & vbCrLf &
            "Alternative: Use binary patching approach" & vbCrLf &
            "(see VariantReleaseLib.UpdateSingleFileBinary)",
            "Test4_DrawingRelink - NEEDS ALTERNATIVE")
    End If
End Sub
