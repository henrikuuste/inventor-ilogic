' ============================================================================
' Mõõdud - Materjali gabariitmõõtude kalkulaator
' 
' Töötab nii detaili kui koostu dokumentidega:
' - Detailis: töötleb aktiivset detaili
' - Koostus: töötleb valitud detailid
'
' Loob igasse detaili lokaalse reegli "Uuenda mõõdud", mis uuendab
' iProperties väärtusi (Paksus, Laius, Pikkus) gabariitmõõtude alusel.
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/BoundingBoxStockLib.vb"

Sub Main()
    UtilsLib.SetLogger(Logger)
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        MessageBox.Show("Aktiivne dokument puudub.", "Mõõdud")
        Exit Sub
    End If

    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        ' Part document - process directly
        ProcessSinglePart(app, CType(doc, PartDocument))
    ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        ' Assembly document - process selected parts
        ProcessAssemblySelection(app, CType(doc, AssemblyDocument))
    Else
        UtilsLib.LogError("Mõõdud: Unsupported document type.")
        MessageBox.Show("See reegel töötab ainult detaili (.ipt) või koostu (.iam) dokumentidega.", "Mõõdud")
    End If
End Sub

Sub ProcessSinglePart(ByVal app As Inventor.Application, ByVal partDoc As PartDocument)
    ' Build form title with filename and description
    Dim partDesc As String = GetPartDescription(partDoc)
    Dim partName As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName)
    Dim formTitle As String = "Mõõdud - " & partName
    If partDesc <> "" Then formTitle &= " - " & partDesc

    ' Process the part with Estonian UI, no skip button, no batch info
    Dim result As String = BoundingBoxStockLib.ProcessPartDocument(app, partDoc, formTitle, False, iLogicVb.Automation, True)
    ' No success message - exit silently if OK
End Sub

Sub ProcessAssemblySelection(ByVal app As Inventor.Application, ByVal asmDoc As AssemblyDocument)
    Dim sel As SelectSet = asmDoc.SelectSet

    If sel Is Nothing OrElse sel.Count = 0 Then
        MessageBox.Show("Detailid pole valitud." & vbCrLf & vbCrLf & _
                        "Valige koostus üks või mitu detaili, seejärel käivitage see reegel.", _
                        "Mõõdud")
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
        MessageBox.Show("Valikus ei leitud detaile." & vbCrLf & vbCrLf & _
                        "Valige detailid (mitte alamkoostud).", _
                        "Mõõdud")
        Exit Sub
    End If

    ' Process each part
    Dim totalCount As Integer = partOccurrences.Count
    Dim currentIndex As Integer = 0

    For Each occ As ComponentOccurrence In partOccurrences
        currentIndex += 1
        
        ' Get the part document
        Dim partDoc As PartDocument = Nothing
        Try
            partDoc = CType(occ.Definition.Document, PartDocument)
        Catch
            Continue For
        End Try

        If partDoc Is Nothing Then
            Continue For
        End If

        ' Build form title with filename, description, and progress
        Dim partDesc As String = GetPartDescription(partDoc)
        Dim partName As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName)
        Dim formTitle As String = "Mõõdud - " & partName
        If partDesc <> "" Then formTitle &= " - " & partDesc
        formTitle &= " (" & currentIndex & "/" & totalCount & ")"

        ' Highlight the current part in the assembly view
        asmDoc.SelectSet.Clear()
        asmDoc.SelectSet.Select(occ)
        app.ActiveView.Update()

        ' Process this part with Estonian UI and skip button
        Dim result As String = BoundingBoxStockLib.ProcessPartDocument(app, partDoc, formTitle, True, iLogicVb.Automation, True)

        If result = "CANCEL" Then
            ' User cancelled - stop processing
            Exit For
        End If
        ' Continue for OK or SKIP results
    Next

    ' No summary message - exit silently
End Sub

Function GetPartDescription(ByVal partDoc As PartDocument) As String
    Try
        ' Try to get Description from Design Tracking Properties (most common location)
        Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
        Dim desc As String = CStr(designProps.Item("Description").Value)
        If desc IsNot Nothing AndAlso desc.Trim() <> "" Then
            Return desc.Trim()
        End If
    Catch
    End Try
    
    Try
        ' Fallback: try Summary Information "Subject" field
        Dim summaryInfo As PropertySet = partDoc.PropertySets.Item("Inventor Summary Information")
        Dim subj As String = CStr(summaryInfo.Item("Subject").Value)
        If subj IsNot Nothing AndAlso subj.Trim() <> "" Then
            Return subj.Trim()
        End If
    Catch
    End Try
    
    Return ""
End Function
