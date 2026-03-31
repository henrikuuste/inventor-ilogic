' ============================================================================
' TestRunAndSave - Test running Update rule and saving
'
' Run this on the released assembly to verify the Update rule works.
' Output goes to the iLogic Log window in Inventor.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        MessageBox.Show("Please open the released assembly.", "Test")
        Exit Sub
    End If
    
    Logger.Info("=== Test: Run Update Rule and Save ===")
    Logger.Info("Document: " & doc.FullFileName)
    
    Dim iLogicAuto As Object = iLogicVb.Automation
    
    ' Step 1: Find rules matching *Update*
    Logger.Info("--- Step 1: Find matching rules ---")
    Dim rules As Object = iLogicAuto.Rules(doc)
    Dim matchingRules As New List(Of String)
    
    For Each rule As Object In rules
        Dim ruleName As String = CStr(CallByName(rule, "Name", CallType.Get))
        If ruleName.IndexOf("Update", StringComparison.OrdinalIgnoreCase) >= 0 Then
            matchingRules.Add(ruleName)
            Logger.Info("  Found: " & ruleName)
        End If
    Next
    
    If matchingRules.Count = 0 Then
        Logger.Warn("  No rules matching '*Update*' found!")
        Exit Sub
    End If
    
    ' Step 2: Run each matching rule
    Logger.Info("--- Step 2: Run rules ---")
    For Each ruleName As String In matchingRules
        Logger.Info("  Running: " & ruleName)
        Try
            iLogicAuto.RunRule(doc, ruleName)
            Logger.Info("    Completed")
        Catch ex As Exception
            Logger.Error("    ERROR: " & ex.Message)
        End Try
    Next
    
    ' Step 3: Update document
    Logger.Info("--- Step 3: Update document ---")
    Try
        doc.Update()
        Logger.Info("  Update completed")
    Catch ex As Exception
        Logger.Error("  Update error: " & ex.Message)
    End Try
    
    ' Step 4: Update all referenced documents
    Logger.Info("--- Step 4: Update referenced documents ---")
    For Each refDoc As Document In doc.ReferencedDocuments
        Try
            refDoc.Update()
            Logger.Info("  Updated: " & System.IO.Path.GetFileName(refDoc.FullFileName))
        Catch
        End Try
    Next
    
    ' Step 5: Force rebuild
    Logger.Info("--- Step 5: Force rebuild ---")
    Try
        If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
            asmDoc.Update2(True)
            Logger.Info("  Rebuild completed")
        Else
            doc.Update()
            Logger.Info("  Update completed")
        End If
    Catch ex As Exception
        Logger.Error("  Rebuild error: " & ex.Message)
    End Try
    
    ' Step 6: Save all
    Logger.Info("--- Step 6: Save all documents ---")
    For Each refDoc As Document In doc.AllReferencedDocuments
        Try
            refDoc.Save()
            Logger.Info("  Saved: " & System.IO.Path.GetFileName(refDoc.FullFileName))
        Catch ex As Exception
            Logger.Error("  Save error: " & System.IO.Path.GetFileName(refDoc.FullFileName) & " - " & ex.Message)
        End Try
    Next
    
    Try
        doc.Save()
        Logger.Info("  Saved: " & System.IO.Path.GetFileName(doc.FullFileName))
    Catch ex As Exception
        Logger.Error("  Save error: " & ex.Message)
    End Try
    
    Logger.Info("=== Done ===")
    Logger.Info("Close and reopen the document to verify changes persisted.")
End Sub
