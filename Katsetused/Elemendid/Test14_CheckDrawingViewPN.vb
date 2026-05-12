' Test14_CheckDrawingViewPN.vb
' Simple diagnostic to check Part Number via different access methods
' Run this on a drawing (e.g., 00016.idw) to see what each method returns

Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    If app.ActiveDocument Is Nothing OrElse app.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        Logger.Error("Please open a drawing first!")
        Return
    End If
    
    Dim drawDoc As DrawingDocument = CType(app.ActiveDocument, DrawingDocument)
    
    Logger.Info("=== DRAWING: " & System.IO.Path.GetFileName(drawDoc.FullFileName) & " ===")
    Logger.Info("")
    
    ' 1. Drawing's own iProperties
    Logger.Info("--- Method 1: Drawing's Own iProperties ---")
    Try
        Dim drawingPN As String = drawDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
        Dim drawingTitle As String = drawDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
        Logger.Info("  Drawing Part Number: " & drawingPN)
        Logger.Info("  Drawing Title: " & drawingTitle)
    Catch ex As Exception
        Logger.Error("  Error: " & ex.Message)
    End Try
    Logger.Info("")
    
    ' 2. ReferencedDocuments collection (often stale!)
    Logger.Info("--- Method 2: ReferencedDocuments Collection (may be stale) ---")
    For Each refDoc As Document In drawDoc.ReferencedDocuments
        Try
            Dim refPN As String = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
            Dim refTitle As String = refDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
            Logger.Info("  " & System.IO.Path.GetFileName(refDoc.FullFileName) & " PN=" & refPN & " Title=" & refTitle)
        Catch ex As Exception
            Logger.Error("  " & System.IO.Path.GetFileName(refDoc.FullFileName) & " Error: " & ex.Message)
        End Try
    Next
    Logger.Info("")
    
    ' 3. DrawingView's ReferencedDocumentDescriptor.ReferencedDocument (usually correct!)
    Logger.Info("--- Method 3: DrawingView.ReferencedDocument ---")
    For Each sheet As Sheet In drawDoc.Sheets
        For Each view As DrawingView In sheet.DrawingViews
            Try
                If view.ReferencedDocumentDescriptor IsNot Nothing AndAlso _
                   view.ReferencedDocumentDescriptor.ReferencedDocument IsNot Nothing Then
                    Dim modelDoc As Document = view.ReferencedDocumentDescriptor.ReferencedDocument
                    Dim viewPN As String = modelDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                    Dim viewTitle As String = modelDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                    Logger.Info("  View '" & view.Name & "' -> " & System.IO.Path.GetFileName(modelDoc.FullFileName))
                    Logger.Info("    PN=" & viewPN & " Title=" & viewTitle)
                End If
            Catch ex As Exception
                Logger.Error("  View '" & view.Name & "' Error: " & ex.Message)
            End Try
        Next
    Next
    Logger.Info("")
    
    ' 4. File.ReferencedFileDescriptors
    Logger.Info("--- Method 4: File.ReferencedFileDescriptors ---")
    For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
        Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
        Logger.Info("  " & System.IO.Path.GetFileName(fd.FullFileName))
        Logger.Info("    InternalNameDifferent: " & fd.ReferenceInternalNameDifferent.ToString())
    Next
    Logger.Info("")
    
    ' 5. Fresh document open from disk
    Logger.Info("--- Method 5: Fresh Document Open from Disk ---")
    For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
        Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
        Dim refPath As String = fd.FullFileName
        If refPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) OrElse _
           refPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
            Try
                ' Open fresh (don't close existing - might break drawing)
                Dim freshDoc As Document = app.Documents.Open(refPath, True)
                Dim freshPN As String = freshDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                Dim freshTitle As String = freshDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                Logger.Info("  " & System.IO.Path.GetFileName(refPath) & " (FRESH)")
                Logger.Info("    PN=" & freshPN & " Title=" & freshTitle)
            Catch ex As Exception
                Logger.Error("  " & System.IO.Path.GetFileName(refPath) & " Error: " & ex.Message)
            End Try
        End If
    Next
    Logger.Info("")
    
    ' 6. Title block text analysis
    Logger.Info("--- Method 6: Title Block ---")
    For Each sheet As Sheet In drawDoc.Sheets
        If sheet.TitleBlock IsNot Nothing Then
            Dim tb As TitleBlock = sheet.TitleBlock
            Logger.Info("  Sheet: " & sheet.Name & " - TitleBlock: " & tb.Definition.Name)
            
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
                    
                    ' Look for Part Number fields
                    If formattedText.ToLower().Contains("part") AndAlso formattedText.ToLower().Contains("number") Then
                        Logger.Info("    Part Number field:")
                        Logger.Info("      Result: " & resultText)
                        If formattedText.Contains("Properties - Model") Then
                            Logger.Info("      Source: MODEL Properties")
                        ElseIf formattedText.Contains("Properties - Drawing") Then
                            Logger.Info("      Source: DRAWING Properties")
                        End If
                    End If
                Next
            Catch ex As Exception
                Logger.Error("    Title block error: " & ex.Message)
            End Try
        End If
    Next
    
    Logger.Info("")
    Logger.Info("=== DIAGNOSTIC COMPLETE ===")
End Sub
