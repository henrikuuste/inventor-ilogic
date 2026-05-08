' Test11_DrawingTitleBlockUpdate.vb
' ============================================================================
' Test script to investigate how to update title block after reference changes
'
' The problem: After replacing a drawing's model reference with a new file
' that has a different Part Number, the title block still shows the old number.
'
' ROOT CAUSE IDENTIFIED: The referenced assembly (00014.iam) has Part Number
' still set to 00003 - the model's Part Number was not updated properly.
'
' This script diagnoses and attempts to fix the Part Number mismatch.
' ============================================================================

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    UtilsLib.SetLogger(Logger)
    
    Dim app As Inventor.Application = ThisApplication
    
    If app.ActiveDocument Is Nothing Then
        Logger.Error("No active document. Open a drawing first.")
        Return
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        Logger.Error("Active document is not a drawing.")
        Return
    End If
    
    Dim drawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
    Logger.Info("Testing on drawing: " & drawDoc.DisplayName)
    Logger.Info("Drawing location: " & drawDoc.FullFileName)
    
    ' Test 1: Check current referenced documents and their Part Numbers
    Logger.Info("=== Test 1: Current references ===")
    For Each refDoc As Document In drawDoc.ReferencedDocuments
        Dim partNum As String = GetPartNumber(refDoc)
        Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(refDoc.FullFileName)
        Dim fullPath As String = refDoc.FullFileName
        
        Logger.Info("  Ref: " & fileName & " -> PartNumber: " & partNum)
        Logger.Info("       Path: " & fullPath)
        
        ' Check if Part Number matches filename (expected after release)
        If partNum <> fileName Then
            Logger.Warn("  MISMATCH: PartNumber (" & partNum & ") does not match filename (" & fileName & ")")
        Else
            Logger.Info("  OK: PartNumber matches filename")
        End If
    Next
    
    ' Test 2: Try to fix Part Number on referenced models
    Logger.Info("=== Test 2: Attempt to fix Part Numbers on referenced models ===")
    For Each refDoc As Document In drawDoc.ReferencedDocuments
        Dim partNum As String = GetPartNumber(refDoc)
        Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(refDoc.FullFileName)
        
        If partNum <> fileName Then
            Logger.Info("  Attempting to set Part Number to: " & fileName)
            Try
                Dim designProps = refDoc.PropertySets.Item("Design Tracking Properties")
                designProps.Item("Part Number").Value = fileName
                refDoc.Save()
                Logger.Info("    SUCCESS: Part Number updated and saved")
                
                ' Verify
                Dim newPartNum As String = GetPartNumber(refDoc)
                Logger.Info("    Verification: PartNumber is now: " & newPartNum)
            Catch ex As Exception
                Logger.Error("    FAILED: " & ex.Message)
            End Try
        End If
    Next
    
    ' Test 3: Update drawing after model changes
    Logger.Info("=== Test 3: Refresh drawing ===")
    Try
        drawDoc.Update()
        Logger.Info("  drawDoc.Update() succeeded")
    Catch ex As Exception
        Logger.Error("  drawDoc.Update() failed: " & ex.Message)
    End Try
    
    ' Test 4: Update individual sheets
    Logger.Info("=== Test 4: Update sheets ===")
    For Each sheet As Sheet In drawDoc.Sheets
        Try
            sheet.Update()
            Logger.Info("  Sheet " & sheet.Name & " updated")
        Catch ex As Exception
            Logger.Error("  Sheet " & sheet.Name & " update failed: " & ex.Message)
        End Try
    Next
    
    ' Test 5: Title block info
    Logger.Info("=== Test 5: Title block info ===")
    For Each sheet As Sheet In drawDoc.Sheets
        Logger.Info("Sheet: " & sheet.Name)
        
        If sheet.TitleBlock IsNot Nothing Then
            Logger.Info("  TitleBlock exists")
            
            Try
                Dim tbDef As TitleBlockDefinition = sheet.TitleBlock.Definition
                Logger.Info("  TitleBlock Definition: " & tbDef.Name)
            Catch ex As Exception
                Logger.Info("  Could not get TitleBlock definition: " & ex.Message)
            End Try
        Else
            Logger.Info("  No TitleBlock on this sheet")
        End If
    Next
    
    ' Test 6: Save drawing
    Logger.Info("=== Test 6: Save drawing ===")
    Try
        drawDoc.Save()
        Logger.Info("  Drawing saved")
    Catch ex As Exception
        Logger.Error("  Save failed: " & ex.Message)
    End Try
    
    ' Final verification
    Logger.Info("=== Final: Check references again ===")
    For Each refDoc As Document In drawDoc.ReferencedDocuments
        Dim partNum As String = GetPartNumber(refDoc)
        Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(refDoc.FullFileName)
        Logger.Info("  Ref: " & fileName & " -> PartNumber: " & partNum)
        
        If partNum = fileName Then
            Logger.Info("    OK: PartNumber now matches")
        Else
            Logger.Warn("    Still mismatched - title block may still show old number")
        End If
    Next
    
    Logger.Info("=== Test complete ===")
    Logger.Info("If Part Number was corrected, close and reopen the drawing to verify title block updates.")
End Sub

Function GetPartNumber(doc As Document) As String
    Try
        Return doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
    Catch
        Return "(unknown)"
    End Try
End Function
