' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' TestReplaceReference - Test ComponentOccurrence.Replace Functionality
' 
' This test rule verifies that we can:
' 1. Copy a part file to a new location
' 2. Replace the assembly reference to point to the copied file
' 3. Save the assembly with updated references
'
' Usage:
' 1. Open an assembly document
' 2. Select a single part occurrence in the browser or graphics
' 3. Run this rule
' 4. The rule will copy the part and update the reference
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=ComponentOccurrence_Replace
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    ' Validate we have an assembly open
    If doc Is Nothing Then
        MessageBox.Show("No active document.", "Test Replace Reference")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This rule only works in assembly documents (.iam).", "Test Replace Reference")
        Exit Sub
    End If
    
    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    Dim sel As SelectSet = asmDoc.SelectSet
    
    ' Get selected occurrence
    Dim selectedOcc As ComponentOccurrence = Nothing
    
    If sel IsNot Nothing AndAlso sel.Count = 1 Then
        If TypeOf sel.Item(1) Is ComponentOccurrence Then
            selectedOcc = CType(sel.Item(1), ComponentOccurrence)
        End If
    End If
    
    If selectedOcc Is Nothing Then
        ' Prompt user to pick a component
        MessageBox.Show("Please select a single part occurrence, then run this rule again." & vbCrLf & vbCrLf & _
                        "You can select in the browser or in the graphics window.", _
                        "Test Replace Reference")
        Exit Sub
    End If
    
    ' Check it's a part (not a sub-assembly)
    If selectedOcc.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("Please select a part occurrence, not a sub-assembly.", "Test Replace Reference")
        Exit Sub
    End If
    
    ' Get the part document
    Dim partDoc As PartDocument = Nothing
    Try
        partDoc = CType(selectedOcc.Definition.Document, PartDocument)
    Catch ex As Exception
        MessageBox.Show("Could not access part document: " & ex.Message, "Test Replace Reference")
        Exit Sub
    End Try
    
    Dim originalPath As String = partDoc.FullFileName
    Dim originalName As String = System.IO.Path.GetFileNameWithoutExtension(originalPath)
    Dim originalFolder As String = System.IO.Path.GetDirectoryName(originalPath)
    
    ' Create test output folder
    Dim testFolder As String = System.IO.Path.Combine(originalFolder, "_TestCopy")
    If Not System.IO.Directory.Exists(testFolder) Then
        System.IO.Directory.CreateDirectory(testFolder)
    End If
    
    ' Create new file path with "_Copy" suffix
    Dim newFileName As String = originalName & "_Copy.ipt"
    Dim newPath As String = System.IO.Path.Combine(testFolder, newFileName)
    
    ' Show confirmation dialog
    Dim msg As String = "This will test the Replace Reference functionality:" & vbCrLf & vbCrLf & _
                        "Original part: " & originalPath & vbCrLf & vbCrLf & _
                        "Will copy to: " & newPath & vbCrLf & vbCrLf & _
                        "Then replace the assembly reference to use the copy." & vbCrLf & vbCrLf & _
                        "Continue?"
    
    Dim result As MsgBoxResult = MessageBox.Show(msg, "Test Replace Reference", MessageBoxButtons.YesNo)
    If result <> MsgBoxResult.Yes Then
        Exit Sub
    End If
    
    ' Step 1: Copy the file
    Try
        System.IO.File.Copy(originalPath, newPath, True)
        LogMessage("Step 1: File copied to " & newPath)
    Catch ex As Exception
        MessageBox.Show("Failed to copy file: " & ex.Message, "Test Replace Reference")
        Exit Sub
    End Try
    
    ' Step 2: Replace the occurrence reference
    Try
        ' The Replace method takes:
        ' - FileName: Full path to the new file
        ' - ReplaceAll: If True, replaces all occurrences of the same component
        selectedOcc.Replace(newPath, False)
        LogMessage("Step 2: Replaced occurrence reference")
    Catch ex As Exception
        MessageBox.Show("Failed to replace reference: " & ex.Message & vbCrLf & vbCrLf & _
                        "Stack trace: " & ex.StackTrace, "Test Replace Reference")
        Exit Sub
    End Try
    
    ' Step 3: Update the assembly
    Try
        asmDoc.Update()
        LogMessage("Step 3: Assembly updated")
    Catch ex As Exception
        MessageBox.Show("Warning: Assembly update had issues: " & ex.Message, "Test Replace Reference")
    End Try
    
    ' Step 4: Save the assembly (optional - ask user)
    msg = "Reference replacement successful!" & vbCrLf & vbCrLf & _
          "The occurrence now references: " & newPath & vbCrLf & vbCrLf & _
          "Do you want to save the assembly with the updated reference?" & vbCrLf & vbCrLf & _
          "(Click No to leave unsaved - you can undo with Ctrl+Z)"
    
    result = MessageBox.Show(msg, "Test Replace Reference", MessageBoxButtons.YesNo)
    If result = MsgBoxResult.Yes Then
        Try
            asmDoc.Save()
            LogMessage("Step 4: Assembly saved")
            MessageBox.Show("Test complete! Assembly saved with updated reference.", "Test Replace Reference")
        Catch ex As Exception
            MessageBox.Show("Failed to save assembly: " & ex.Message, "Test Replace Reference")
        End Try
    Else
        MessageBox.Show("Test complete! Assembly NOT saved." & vbCrLf & _
                        "Use Ctrl+Z to undo the reference change if needed.", "Test Replace Reference")
    End If
End Sub

Sub LogMessage(msg As String)
    ' Log to Inventor's log for debugging
    ' You can view this in the iLogic Log Browser
    Try
        Logger.Info(msg)
    Catch
        ' Logger may not be available in all contexts
    End Try
End Sub


