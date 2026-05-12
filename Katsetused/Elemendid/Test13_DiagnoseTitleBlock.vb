' Test13_DiagnoseTitleBlock.vb
' Diagnostic script to investigate drawing title block property sources and caching
' Run this on the problematic 00016.idw drawing to understand why it shows wrong Part Number

Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    If app.ActiveDocument Is Nothing OrElse app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        Logger.Error("Please open a drawing file first!")
        Return
    End If
    
    Dim drawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
    
    Logger.Info("=== DRAWING TITLE BLOCK DIAGNOSTIC ===")
    Logger.Info("Drawing: " & drawDoc.FullFileName)
    Logger.Info("")
    
    ' SECTION 1: Drawing's own iProperties
    Logger.Info("--- SECTION 1: Drawing's Own iProperties ---")
    Try
        Dim drawingPN As String = drawDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
        Dim drawingTitle As String = drawDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
        Logger.Info("Drawing Part Number: " & drawingPN)
        Logger.Info("Drawing Title: " & drawingTitle)
    Catch ex As Exception
        Logger.Error("Error reading drawing properties: " & ex.Message)
    End Try
    Logger.Info("")
    
    ' SECTION 2: Referenced Documents via ReferencedDocuments collection
    Logger.Info("--- SECTION 2: ReferencedDocuments Collection ---")
    Logger.Info("Count: " & drawDoc.ReferencedDocuments.Count)
    Dim refIndex As Integer = 1
    For Each refDoc As Document In drawDoc.ReferencedDocuments
        Logger.Info("")
        Logger.Info("  [" & refIndex & "] Path: " & refDoc.FullFileName)
        Try
            Dim refPN As String = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
            Dim refTitle As String = refDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
            Logger.Info("      Part Number: " & refPN)
            Logger.Info("      Title: " & refTitle)
            Logger.Info("      InternalName: " & refDoc.InternalName)
        Catch ex As Exception
            Logger.Error("      Error: " & ex.Message)
        End Try
        refIndex += 1
    Next
    Logger.Info("")
    
    ' SECTION 3: File.ReferencedFileDescriptors
    Logger.Info("--- SECTION 3: File.ReferencedFileDescriptors ---")
    Logger.Info("Count: " & drawDoc.File.ReferencedFileDescriptors.Count)
    For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
        Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
        Logger.Info("")
        Logger.Info("  [" & i & "] FullFileName: " & fd.FullFileName)
        Logger.Info("      ReferenceInternalNameDifferent: " & fd.ReferenceInternalNameDifferent.ToString())
        Logger.Info("      ReferenceLocationDifferent: " & fd.ReferenceLocationDifferent.ToString())
    Next
    Logger.Info("")
    
    ' SECTION 4: Drawing Views and their references
    Logger.Info("--- SECTION 4: Drawing Views ---")
    For Each sheet As Sheet In drawDoc.Sheets
        Logger.Info("")
        Logger.Info("Sheet: " & sheet.Name & " - Views: " & sheet.DrawingViews.Count)
        For Each view As DrawingView In sheet.DrawingViews
            Logger.Info("  View: " & view.Name)
            Try
                If view.ReferencedDocumentDescriptor IsNot Nothing Then
                    Logger.Info("    FullDocumentName: " & view.ReferencedDocumentDescriptor.FullDocumentName)
                    If view.ReferencedDocumentDescriptor.ReferencedDocument IsNot Nothing Then
                        Dim refDoc As Document = view.ReferencedDocumentDescriptor.ReferencedDocument
                        Dim pn As String = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                        Logger.Info("    ReferencedDocument.PartNumber: " & pn)
                    End If
                End If
            Catch ex As Exception
                Logger.Error("    Error: " & ex.Message)
            End Try
        Next
    Next
    Logger.Info("")
    
    ' SECTION 5: Title Block Analysis
    Logger.Info("--- SECTION 5: Title Block Analysis ---")
    For Each sheet As Sheet In drawDoc.Sheets
        Logger.Info("")
        Logger.Info("Sheet: " & sheet.Name)
        If sheet.TitleBlock IsNot Nothing Then
            Dim tb As TitleBlock = sheet.TitleBlock
            Logger.Info("  TitleBlock Definition: " & tb.Definition.Name)
            
            Try
                For Each textBox As Inventor.TextBox In tb.Definition.Sketch.TextBoxes
                    Dim formattedText As String = ""
                    Dim resultText As String = ""
                    Try
                        formattedText = textBox.FormattedText
                    Catch
                    End Try
                    Try
                        resultText = tb.GetResultText(textBox)
                    Catch
                    End Try
                    
                    ' Look for Part Number related fields
                    If formattedText.Contains("Part Number") OrElse formattedText.Contains("PART_NUMBER") OrElse _
                       formattedText.Contains("<Part Number>") OrElse formattedText.ToLower().Contains("partnumber") Then
                        Logger.Info("  FOUND Part Number TextBox:")
                        Logger.Info("    FormattedText: " & formattedText.Replace(vbCr, " ").Replace(vbLf, " "))
                        Logger.Info("    ResultText: " & resultText)
                        
                        If formattedText.Contains("Properties - Model") Then
                            Logger.Info("    SOURCE: Model Properties (live link)")
                        ElseIf formattedText.Contains("Properties - Drawing") Then
                            Logger.Info("    SOURCE: Drawing Properties (copied, NOT live)")
                        Else
                            Logger.Info("    SOURCE: Unknown")
                        End If
                    End If
                Next
            Catch ex As Exception
                Logger.Error("  Error analyzing title block: " & ex.Message)
            End Try
        Else
            Logger.Info("  No TitleBlock on this sheet")
        End If
    Next
    Logger.Info("")
    
    ' SECTION 6: Fresh document comparison
    Logger.Info("--- SECTION 6: Fresh Document Comparison ---")
    If drawDoc.ReferencedDocuments.Count > 0 Then
        Try
            Dim firstRefPath As String = ""
            For Each refDoc As Document In drawDoc.ReferencedDocuments
                firstRefPath = refDoc.FullFileName
                Exit For
            Next
            
            If firstRefPath <> "" Then
                Logger.Info("First referenced file: " & firstRefPath)
                
                ' Open fresh
                Dim freshDoc As Document = app.Documents.Open(firstRefPath, True)
                Dim freshPN As String = freshDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                Dim freshTitle As String = freshDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                Logger.Info("Freshly opened document:")
                Logger.Info("  Part Number: " & freshPN)
                Logger.Info("  Title: " & freshTitle)
                Logger.Info("  InternalName: " & freshDoc.InternalName)
                
                ' Now check what drawing sees for the same file
                For Each refDoc As Document In drawDoc.ReferencedDocuments
                    If refDoc.FullFileName.Equals(firstRefPath, StringComparison.OrdinalIgnoreCase) Then
                        Dim drawRefPN As String = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                        Logger.Info("Drawing's ReferencedDocuments for same file:")
                        Logger.Info("  Part Number: " & drawRefPN)
                        Logger.Info("  InternalName: " & refDoc.InternalName)
                        
                        If freshPN <> drawRefPN Then
                            Logger.Warn("  *** MISMATCH DETECTED ***")
                            Logger.Info("  Fresh doc InternalName same? " & (freshDoc.InternalName = refDoc.InternalName).ToString())
                        Else
                            Logger.Info("  Values match!")
                        End If
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            Logger.Error("Error in fresh comparison: " & ex.Message)
        End Try
    End If
    
    Logger.Info("")
    Logger.Info("=== DIAGNOSTIC COMPLETE ===")
End Sub
