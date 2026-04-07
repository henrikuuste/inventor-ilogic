' ============================================================================
' BoundingBoxStockBatch - Batch Material Stock Size Calculator
' 
' Runs from an assembly document. Iterates through all selected parts and
' applies the BoundingBoxStock configuration to each one.
'
' Usage:
' 1. Open an assembly document
' 2. Select one or more part occurrences in the browser or graphics
' 3. Run this rule
' 4. Configure each part's axis settings when prompted
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/BoundingBoxStockLib.vb"

Sub Main()
    UtilsLib.SetLogger(Logger)
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        MessageBox.Show("No active document.", "Bounding Box Stock Batch")
        Exit Sub
    End If

    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This rule only works in assembly documents (.iam)." & vbCrLf & _
                        "Select parts in the assembly, then run this rule.", "Bounding Box Stock Batch")
        Exit Sub
    End If

    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    Dim sel As SelectSet = asmDoc.SelectSet

    If sel Is Nothing OrElse sel.Count = 0 Then
        MessageBox.Show("No parts selected." & vbCrLf & vbCrLf & _
                        "Please select one or more part occurrences in the assembly, then run this rule.", _
                        "Bounding Box Stock Batch")
        Exit Sub
    End If

    ' Collect part occurrences from selection (filter out sub-assemblies and non-occurrences)
    Dim partOccurrences As New System.Collections.Generic.List(Of ComponentOccurrence)
    
    For Each selObj As Object In sel
        If TypeOf selObj Is ComponentOccurrence Then
            Dim occ As ComponentOccurrence = CType(selObj, ComponentOccurrence)
            ' Check if it's a part (not a sub-assembly)
            If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                ' Check if we already have this part (avoid duplicates from same definition)
                Dim alreadyAdded As Boolean = False
                For Each existingOcc As ComponentOccurrence In partOccurrences
                    If existingOcc.Definition Is occ.Definition Then
                        alreadyAdded = True
                        Exit For
                    End If
                Next
                If Not alreadyAdded Then
                    partOccurrences.Add(occ)
                End If
            End If
        End If
    Next

    If partOccurrences.Count = 0 Then
        MessageBox.Show("No part occurrences found in selection." & vbCrLf & vbCrLf & _
                        "Please select part occurrences (not sub-assemblies).", _
                        "Bounding Box Stock Batch")
        Exit Sub
    End If

    ' Process each part
    Dim processedCount As Integer = 0
    Dim skippedCount As Integer = 0
    Dim totalCount As Integer = partOccurrences.Count
    Dim currentIndex As Integer = 0

    For Each occ As ComponentOccurrence In partOccurrences
        currentIndex += 1
        
        ' Get the part document
        Dim partDoc As PartDocument = Nothing
        Try
            partDoc = CType(occ.Definition.Document, PartDocument)
        Catch
            skippedCount += 1
            Continue For
        End Try

        If partDoc Is Nothing Then
            skippedCount += 1
            Continue For
        End If

        ' Build form title with part name and progress
        Dim partName As String = System.IO.Path.GetFileName(partDoc.FullFileName)
        Dim formTitle As String = "Bounding Box Stock - " & partName & " (" & currentIndex & " of " & totalCount & ")"

        ' Process this part using the shared function from BoundingBoxStockLib
        Dim result As String = BoundingBoxStockLib.ProcessPartDocument(app, partDoc, formTitle, True, iLogicVb.Automation)

        If result = "CANCEL" Then
            ' User cancelled - stop processing
            Exit For
        ElseIf result = "SKIP" Then
            skippedCount += 1
        ElseIf result = "OK" Then
            processedCount += 1
        End If
    Next

    ' Show summary
    MessageBox.Show( _
        "Batch processing complete." & vbCrLf & vbCrLf & _
        "Parts processed: " & processedCount & vbCrLf & _
        "Parts skipped: " & skippedCount & vbCrLf & _
        "Total selected: " & totalCount, _
        "Bounding Box Stock Batch")

End Sub
