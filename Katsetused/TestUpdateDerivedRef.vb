' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' TestUpdateDerivedRef - Try creating new derived definition approach
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("Please open the COPIED derived part file.", "Test")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim info As New System.Text.StringBuilder()
    
    info.AppendLine("=== TEST: Replace Derived Part via CreateDefinition ===")
    info.AppendLine()
    info.AppendLine("Document: " & partDoc.FullFileName)
    info.AppendLine()
    
    ' Get current base
    Dim originalRef As String = ""
    If partDoc.ReferencedDocuments.Count > 0 Then
        originalRef = partDoc.ReferencedDocuments.Item(1).FullFileName
    End If
    info.AppendLine("Current base reference: " & originalRef)
    info.AppendLine()
    
    ' Step 1: Get existing derived component info
    info.AppendLine("=== Step 1: Get existing derived component ===")
    Dim existingDpc As DerivedPartComponent = Nothing
    Dim existingDef As Object = Nothing
    
    Try
        Dim dpc As DerivedPartComponents = compDef.ReferenceComponents.DerivedPartComponents
        If dpc.Count > 0 Then
            existingDpc = dpc.Item(1)
            existingDef = existingDpc.Definition
            info.AppendLine("  Found: " & existingDpc.Name)
            info.AppendLine("  Definition type: " & TypeName(existingDef))
        Else
            info.AppendLine("  No derived components found")
            ShowResult(info.ToString())
            Exit Sub
        End If
    Catch ex As Exception
        info.AppendLine("  Error: " & ex.Message)
        ShowResult(info.ToString())
        Exit Sub
    End Try
    info.AppendLine()
    
    ' Step 2: Ask for new path
    Dim newPath As String = InputBox( _
        "Enter the new base file path:" & vbCrLf & vbCrLf & _
        "Current: " & originalRef, _
        "New Base Path", originalRef)
    
    If String.IsNullOrEmpty(newPath) OrElse newPath = originalRef Then
        info.AppendLine("Cancelled or no change.")
        ShowResult(info.ToString())
        Exit Sub
    End If
    
    If Not System.IO.File.Exists(newPath) Then
        info.AppendLine("ERROR: File does not exist: " & newPath)
        ShowResult(info.ToString())
        Exit Sub
    End If
    
    info.AppendLine("New path: " & newPath)
    info.AppendLine()
    
    ' Step 3: Try to create a new definition
    info.AppendLine("=== Step 2: Try CreateDefinition ===")
    Try
        Dim dpc As DerivedPartComponents = compDef.ReferenceComponents.DerivedPartComponents
        
        ' Try CreateDefinition with the new path
        info.AppendLine("  Calling CreateDefinition(" & System.IO.Path.GetFileName(newPath) & ")...")
        Dim newDef As Object = dpc.CreateDefinition(newPath)
        info.AppendLine("  SUCCESS! New definition created")
        info.AppendLine("  New definition type: " & TypeName(newDef))
        
        ' Check what properties the new definition has
        info.AppendLine()
        info.AppendLine("  New definition properties:")
        Try
            Dim fullName As String = CStr(CallByName(newDef, "FullDocumentName", CallType.Get))
            info.AppendLine("    .FullDocumentName: " & fullName)
        Catch : End Try
        
        ' Step 4: Try to apply the new definition
        info.AppendLine()
        info.AppendLine("=== Step 3: Apply new definition ===")
        
        ' Option A: Try Add with the new definition
        info.AppendLine("  Trying dpc.Add(newDef)...")
        Try
            Dim newComp As DerivedPartComponent = dpc.Add(newDef)
            info.AppendLine("  SUCCESS! New component added: " & newComp.Name)
            
            ' Delete the old one
            info.AppendLine("  Deleting old component...")
            Try
                existingDpc.Delete()
                info.AppendLine("  Old component deleted")
            Catch ex As Exception
                info.AppendLine("  Delete error: " & ex.Message)
            End Try
        Catch ex As Exception
            info.AppendLine("  Add error: " & ex.Message)
        End Try
        
    Catch ex As Exception
        info.AppendLine("  CreateDefinition error: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' Step 5: Verify
    info.AppendLine("=== Step 4: Verify result ===")
    For Each refDoc As Document In partDoc.ReferencedDocuments
        info.AppendLine("  Current reference: " & refDoc.FullFileName)
        If refDoc.FullFileName.Equals(newPath, StringComparison.OrdinalIgnoreCase) Then
            info.AppendLine("  SUCCESS!")
        ElseIf refDoc.FullFileName.Equals(originalRef, StringComparison.OrdinalIgnoreCase) Then
            info.AppendLine("  STILL ORIGINAL")
        End If
    Next
    
    ShowResult(info.ToString())
End Sub

Sub ShowResult(text As String)
    Dim resultForm As New System.Windows.Forms.Form()
    resultForm.Text = "Replace Derived Part - Test Results"
    resultForm.Width = 900
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
