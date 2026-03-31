' ============================================================================
' UpdateSupportProperties - Batch iProperty Update for Kask24 Supports
' 
' Run this rule in an assembly to update iProperties on all Kask24 support
' parts. This ensures BOM displays correct Width, Height, Thickness values.
'
' Follows BoundingBoxStock conventions:
' - Thickness = Y axis = 24mm (fixed beam thickness)
' - Width = Z axis = Length parameter (longest dimension for BOM)
' - Height = X axis = Width parameter (beam cross-section width)
'
' Usage:
' - Run in an assembly containing Kask24 supports
' - Updates all Kask24 instances automatically
' - Can also select specific components to update only those
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/SupportPlacementLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        MessageBox.Show("No active document.", "Update Support Properties")
        Exit Sub
    End If

    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This rule only works in assembly documents (.iam).", "Update Support Properties")
        Exit Sub
    End If

    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    Dim sel As SelectSet = asmDoc.SelectSet
    
    ' Determine mode: selected components or all
    Dim updateSelected As Boolean = (sel IsNot Nothing AndAlso sel.Count > 0)
    
    Dim updatedCount As Integer = 0
    Dim skippedCount As Integer = 0
    Dim alreadyProcessed As New System.Collections.Generic.HashSet(Of String)()
    
    If updateSelected Then
        ' Process only selected components
        For Each obj As Object In sel
            ProcessObject(obj, alreadyProcessed, updatedCount, skippedCount)
        Next
    Else
        ' Process all occurrences in assembly
        ProcessOccurrences(asmDef.Occurrences, alreadyProcessed, updatedCount, skippedCount)
    End If
    
    Dim modeText As String = If(updateSelected, "selected components", "all components")
    MessageBox.Show( _
        "Updated iProperties on " & modeText & "." & vbCrLf & vbCrLf & _
        "Kask24 supports updated: " & updatedCount & vbCrLf & _
        "Skipped (not Kask24): " & skippedCount, _
        "Update Support Properties")
End Sub

' ============================================================================
' Process a selected object (may be occurrence or proxy)
' ============================================================================
Sub ProcessObject(obj As Object, alreadyProcessed As System.Collections.Generic.HashSet(Of String), _
                  ByRef updatedCount As Integer, ByRef skippedCount As Integer)
    
    Dim occ As ComponentOccurrence = Nothing
    
    If TypeOf obj Is ComponentOccurrenceProxy Then
        Dim proxy As ComponentOccurrenceProxy = CType(obj, ComponentOccurrenceProxy)
        occ = proxy.NativeObject
    ElseIf TypeOf obj Is ComponentOccurrence Then
        occ = CType(obj, ComponentOccurrence)
    End If
    
    If occ Is Nothing Then
        skippedCount += 1
        Return
    End If
    
    ' Only process part documents
    If occ.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        skippedCount += 1
        Return
    End If
    
    Dim partDoc As PartDocument = Nothing
    Try
        partDoc = CType(occ.Definition.Document, PartDocument)
    Catch
        skippedCount += 1
        Return
    End Try
    
    If partDoc Is Nothing Then
        skippedCount += 1
        Return
    End If
    
    ' Skip if already processed (multiple instances of same part)
    Dim docPath As String = partDoc.FullFileName
    If alreadyProcessed.Contains(docPath) Then
        Return
    End If
    alreadyProcessed.Add(docPath)
    
    ' Check if this is a Kask24 member
    If Not SupportPlacementLib.IsKask24Support(partDoc) Then
        skippedCount += 1
        Return
    End If
    
    ' Update iProperties
    Try
        ' Get dimensions from parameters
        Dim params As Parameters = partDoc.ComponentDefinition.Parameters
        Dim widthMm As Integer = CInt(Math.Round(params.Item("Width").Value * 10))
        Dim lengthMm As Integer = CInt(Math.Round(params.Item("Length").Value * 10))
        
        SupportPlacementLib.UpdateSupportiProperties(partDoc, widthMm, lengthMm)
        partDoc.Save()
        updatedCount += 1
    Catch
        skippedCount += 1
    End Try
End Sub

' ============================================================================
' Recursively process all occurrences in the assembly
' ============================================================================
Sub ProcessOccurrences(occurrences As ComponentOccurrences, _
                       alreadyProcessed As System.Collections.Generic.HashSet(Of String), _
                       ByRef updatedCount As Integer, ByRef skippedCount As Integer)
    
    For Each occ As ComponentOccurrence In occurrences
        ' Check if it's a part
        If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As PartDocument = Nothing
            Try
                partDoc = CType(occ.Definition.Document, PartDocument)
            Catch
                Continue For
            End Try
            
            If partDoc Is Nothing Then Continue For
            
            ' Skip if already processed
            Dim docPath As String = partDoc.FullFileName
            If alreadyProcessed.Contains(docPath) Then
                Continue For
            End If
            alreadyProcessed.Add(docPath)
            
            ' Check if Kask24 support
            If SupportPlacementLib.IsKask24Support(partDoc) Then
                Try
                    ' Get dimensions from parameters
                    Dim params As Parameters = partDoc.ComponentDefinition.Parameters
                    Dim widthMm As Integer = CInt(Math.Round(params.Item("Width").Value * 10))
                    Dim lengthMm As Integer = CInt(Math.Round(params.Item("Length").Value * 10))
                    
                    SupportPlacementLib.UpdateSupportiProperties(partDoc, widthMm, lengthMm)
                    partDoc.Save()
                    updatedCount += 1
                Catch
                    skippedCount += 1
                End Try
            Else
                skippedCount += 1
            End If
            
        ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            ' Recurse into sub-assemblies
            Try
                Dim subAsmDef As AssemblyComponentDefinition = _
                    CType(occ.Definition, AssemblyComponentDefinition)
                ProcessOccurrences(subAsmDef.Occurrences, alreadyProcessed, updatedCount, skippedCount)
            Catch
            End Try
        End If
    Next
End Sub
