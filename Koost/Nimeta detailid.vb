' ============================================================================
' Nimeta detailid - Rename assembly occurrences with descriptive names
' 
' Renames occurrences using pattern: "<Description> (<Part Number>):<instance>"
' This makes parts easier to identify in the assembly browser.
'
' Usage: 
' - Run from an open assembly
' - If occurrences are selected, only those are renamed
' - If nothing is selected, all top-level occurrences are renamed
' ============================================================================

AddVbFile "Lib/OccurrenceNamingLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim logs As New List(Of String)
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        Logger.Error("Nimeta detailid: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Nimeta detailid")
        Exit Sub
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Error("Nimeta detailid: Active document is not an assembly")
        MessageBox.Show("Aktiivseks dokumendiks peab olema koost (.iam).", "Nimeta detailid")
        Exit Sub
    End If
    
    Dim asmDoc As AssemblyDocument = CType(app.ActiveDocument, AssemblyDocument)
    Logger.Info("Nimeta detailid: Starting for " & asmDoc.DisplayName)
    
    ' Check if there are selected occurrences
    Dim selectedCount As Integer = OccurrenceNamingLib.GetSelectedOccurrenceCount(asmDoc)
    Dim renamedCount As Integer = 0
    
    If selectedCount > 0 Then
        ' Rename only selected occurrences
        Logger.Info("Nimeta detailid: Renaming " & selectedCount & " selected occurrence(s)")
        renamedCount = OccurrenceNamingLib.RenameSelectedOccurrences(asmDoc, logs)
    Else
        ' Rename all occurrences
        Dim totalCount As Integer = asmDoc.ComponentDefinition.Occurrences.Count
        Logger.Info("Nimeta detailid: Renaming all " & totalCount & " occurrence(s)")
        renamedCount = OccurrenceNamingLib.RenameAllOccurrences(asmDoc, logs)
    End If
    
    ' Output logs
    For Each log As String In logs
        Logger.Info(log)
    Next
    
    ' Summary
    If selectedCount > 0 Then
        Logger.Info("Nimeta detailid: Renamed " & renamedCount & " of " & selectedCount & " selected occurrence(s)")
    Else
        Logger.Info("Nimeta detailid: Renamed " & renamedCount & " occurrence(s)")
    End If
    
    ' Refresh view
    app.ActiveView.Update()
End Sub
