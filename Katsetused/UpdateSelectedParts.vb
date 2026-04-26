' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' UpdateSelectedParts - Update Physical Properties of Selected Parts
' 
' Run this rule in an assembly to update all selected parts.
' Fixes N/A physical properties on parts created via "Make Components"
' from a multi-body part.
'
' Usage: Select one or more components in the assembly, then run this rule.
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        MessageBox.Show("No active document.", "Update Selected Parts")
        Exit Sub
    End If

    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This rule only works in assembly documents (.iam).", "Update Selected Parts")
        Exit Sub
    End If

    Dim sel As SelectSet = doc.SelectSet
    If sel Is Nothing OrElse sel.Count = 0 Then
        MessageBox.Show("Select one or more components, then run this rule.", "Update Selected Parts")
        Exit Sub
    End If

    Dim updatedCount As Integer = 0
    Dim skippedCount As Integer = 0
    Dim alreadyProcessed As New System.Collections.Generic.HashSet(Of String)()

    For Each obj As Object In sel
        ' Handle both ComponentOccurrence and ComponentOccurrenceProxy
        Dim occ As ComponentOccurrence = Nothing

        If TypeOf obj Is ComponentOccurrenceProxy Then
            Dim occProxy As ComponentOccurrenceProxy = CType(obj, ComponentOccurrenceProxy)
            occ = occProxy.NativeObject
        ElseIf TypeOf obj Is ComponentOccurrence Then
            occ = CType(obj, ComponentOccurrence)
        Else
            skippedCount += 1
            Continue For
        End If

        ' Get the referenced document
        Dim refDoc As Document = Nothing
        Try
            refDoc = occ.Definition.Document
        Catch
            skippedCount += 1
            Continue For
        End Try

        If refDoc Is Nothing Then
            skippedCount += 1
            Continue For
        End If

        ' Only process part documents
        If refDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            skippedCount += 1
            Continue For
        End If

        ' Skip if we already processed this document (multiple instances of same part)
        Dim docPath As String = refDoc.FullFileName
        If alreadyProcessed.Contains(docPath) Then
            Continue For
        End If
        alreadyProcessed.Add(docPath)

        ' Update the part document - use Rebuild to force full recalculation
        Try
            Dim partDoc As PartDocument = CType(refDoc, PartDocument)
            Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
            
            ' Rebuild forces full recalculation including physical properties
            partDoc.Rebuild()
            
            ' Access MassProperties to ensure they are calculated
            Dim massProps As MassProperties = partDef.MassProperties
            Dim dummy As Double = massProps.Mass  ' Force calculation
            
            ' Update the document
            partDoc.Update()
            
            updatedCount += 1
        Catch ex As Exception
            skippedCount += 1
        End Try
    Next

    MessageBox.Show( _
        "Updated: " & updatedCount & " part(s)" & vbCrLf & _
        "Skipped: " & skippedCount, _
        "Update Selected Parts")
End Sub

