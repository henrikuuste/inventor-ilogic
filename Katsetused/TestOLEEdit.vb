' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' TestOLEEdit - Replace derived part reference in binary
' ============================================================================

Imports System.IO
Imports System.Text

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    Dim info As New StringBuilder()
    info.AppendLine("=== REPLACE: Derived Part Reference ===")
    info.AppendLine()
    
    ' Step 1: Select file
    Dim ofd As New System.Windows.Forms.OpenFileDialog()
    ofd.Title = "Select the COPIED derived part file (must be CLOSED)"
    ofd.Filter = "Inventor Part (*.ipt)|*.ipt"
    
    If ofd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
        Exit Sub
    End If
    
    Dim filePath As String = ofd.FileName
    info.AppendLine("File: " & filePath)
    
    ' Make sure not open
    For Each doc As Document In app.Documents
        If doc.FullFileName.Equals(filePath, StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("File is open in Inventor. Close it first!", "Error")
            Exit Sub
        End If
    Next
    
    ' Read file
    Dim fileBytes As Byte() = System.IO.File.ReadAllBytes(filePath)
    info.AppendLine("File size: " & fileBytes.Length & " bytes")
    info.AppendLine()
    
    ' Get old and new paths
    Dim oldPath As String = InputBox( _
        "Enter the OLD path (the original base file path):" & vbCrLf & vbCrLf & _
        "e.g.: C:\Users\henri\Documents\Inventor\ScriptTesting\Põhi\Eskiis.ipt", _
        "Old Path")
    
    If String.IsNullOrEmpty(oldPath) Then Exit Sub
    
    Dim newPath As String = InputBox( _
        "Enter the NEW path (the copied base file path):" & vbCrLf & vbCrLf & _
        "Old: " & oldPath & vbCrLf & _
        "Old length: " & oldPath.Length & " chars", _
        "New Path")
    
    If String.IsNullOrEmpty(newPath) Then Exit Sub
    
    info.AppendLine("Old path: " & oldPath)
    info.AppendLine("Old length: " & oldPath.Length)
    info.AppendLine()
    info.AppendLine("New path: " & newPath)
    info.AppendLine("New length: " & newPath.Length)
    info.AppendLine()
    
    ' Check length difference
    Dim lengthDiff As Integer = newPath.Length - oldPath.Length
    
    If lengthDiff <> 0 Then
        info.AppendLine("*** PATH LENGTH MISMATCH: " & lengthDiff & " characters ***")
        info.AppendLine()
        
        If lengthDiff > 0 Then
            info.AppendLine("New path is LONGER. Options:")
            info.AppendLine("  1. Use a shorter folder name for the variant")
            info.AppendLine("  2. Pad the old path (not recommended)")
            info.AppendLine()
            
            ' Calculate required folder length
            info.AppendLine("To match lengths, your variant folder path should be:")
            Dim neededLength As Integer = oldPath.Length
            info.AppendLine("  Total path: " & neededLength & " characters")
        Else
            info.AppendLine("New path is SHORTER. We can pad with trailing spaces.")
            
            ' Pad new path to match old length
            newPath = newPath.PadRight(oldPath.Length)
            info.AppendLine("Padded new path: '" & newPath & "'")
            info.AppendLine("Padded length: " & newPath.Length)
        End If
        
        Dim proceed As MsgBoxResult = MsgBox( _
            "Path lengths differ by " & lengthDiff & " characters." & vbCrLf & vbCrLf & _
            "Continue anyway? (May corrupt file)", _
            MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, _
            "Length Mismatch")
        
        If proceed <> MsgBoxResult.Yes Then
            info.AppendLine("Aborted.")
            ShowResult(info.ToString())
            Exit Sub
        End If
    End If
    
    ' Create backup
    Dim backupPath As String = filePath & ".backup"
    System.IO.File.Copy(filePath, backupPath, True)
    info.AppendLine("Backup: " & backupPath)
    info.AppendLine()
    
    ' Find and replace (Unicode)
    Dim oldBytes As Byte() = Encoding.Unicode.GetBytes(oldPath)
    Dim newBytes As Byte() = Encoding.Unicode.GetBytes(newPath)
    
    ' If lengths match, do direct replacement
    ' If new is shorter (padded), use padded version
    ' If new is longer, we'll try but it might break
    
    info.AppendLine("=== Searching and replacing ===")
    
    Dim pos As Integer = IndexOfBytes(fileBytes, oldBytes, 0)
    Dim replaceCount As Integer = 0
    
    Do While pos >= 0
        info.AppendLine("  Found at position: " & pos)
        
        If oldBytes.Length = newBytes.Length Then
            ' Same length - safe replacement
            Array.Copy(newBytes, 0, fileBytes, pos, newBytes.Length)
            info.AppendLine("    Replaced (same length)")
            replaceCount += 1
        ElseIf newBytes.Length < oldBytes.Length Then
            ' New is shorter - pad and replace
            Dim paddedNew As Byte() = New Byte(oldBytes.Length - 1) {}
            Array.Copy(newBytes, 0, paddedNew, 0, newBytes.Length)
            ' Fill rest with spaces (Unicode space = 0x20 0x00)
            For i As Integer = newBytes.Length To paddedNew.Length - 1 Step 2
                paddedNew(i) = &H20
                If i + 1 < paddedNew.Length Then paddedNew(i + 1) = &H0
            Next
            Array.Copy(paddedNew, 0, fileBytes, pos, paddedNew.Length)
            info.AppendLine("    Replaced (padded)")
            replaceCount += 1
        Else
            ' New is longer - risky, try anyway
            ' This will overwrite following bytes!
            info.AppendLine("    WARNING: New path longer - overwriting following bytes!")
            Array.Copy(newBytes, 0, fileBytes, pos, Math.Min(newBytes.Length, fileBytes.Length - pos))
            replaceCount += 1
        End If
        
        pos = IndexOfBytes(fileBytes, oldBytes, pos + 1)
    Loop
    
    info.AppendLine()
    info.AppendLine("Total replacements: " & replaceCount)
    
    If replaceCount = 0 Then
        info.AppendLine("No occurrences found to replace!")
        ShowResult(info.ToString())
        Exit Sub
    End If
    
    ' Write file
    info.AppendLine()
    info.AppendLine("=== Writing modified file ===")
    Try
        System.IO.File.WriteAllBytes(filePath, fileBytes)
        info.AppendLine("File written successfully!")
    Catch ex As Exception
        info.AppendLine("Write error: " & ex.Message)
        info.AppendLine("Restoring backup...")
        System.IO.File.Copy(backupPath, filePath, True)
        ShowResult(info.ToString())
        Exit Sub
    End Try
    
    ' Verify
    info.AppendLine()
    info.AppendLine("=== Verifying ===")
    Try
        Dim testDoc As Document = app.Documents.Open(filePath, False)
        info.AppendLine("File opened successfully!")
        
        For Each refDoc As Document In testDoc.ReferencedDocuments
            info.AppendLine("  Referenced: " & refDoc.FullFileName)
            If refDoc.FullFileName.Trim().Equals(newPath.Trim(), StringComparison.OrdinalIgnoreCase) Then
                info.AppendLine("  *** SUCCESS! ***")
            End If
        Next
        
        testDoc.Close(True)
    Catch ex As Exception
        info.AppendLine("ERROR opening file: " & ex.Message)
        info.AppendLine()
        info.AppendLine("File may be corrupted. Restore from backup:")
        info.AppendLine("  " & backupPath)
    End Try
    
    ShowResult(info.ToString())
End Sub

Function IndexOfBytes(source As Byte(), pattern As Byte(), startIndex As Integer) As Integer
    If source Is Nothing OrElse pattern Is Nothing Then Return -1
    If pattern.Length > source.Length - startIndex Then Return -1
    
    For i As Integer = startIndex To source.Length - pattern.Length
        Dim found As Boolean = True
        For j As Integer = 0 To pattern.Length - 1
            If source(i + j) <> pattern(j) Then
                found = False
                Exit For
            End If
        Next
        If found Then Return i
    Next
    
    Return -1
End Function

Sub ShowResult(text As String)
    Dim resultForm As New System.Windows.Forms.Form()
    resultForm.Text = "OLE Edit Results"
    resultForm.Width = 1000
    resultForm.Height = 700
    resultForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    
    Dim txt As New System.Windows.Forms.TextBox()
    txt.Multiline = True
    txt.ScrollBars = System.Windows.Forms.ScrollBars.Both
    txt.Dock = System.Windows.Forms.DockStyle.Fill
    txt.Text = text
    txt.ReadOnly = True
    
    resultForm.Controls.Add(txt)
    resultForm.ShowDialog()
End Sub
