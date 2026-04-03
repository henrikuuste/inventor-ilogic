' ============================================================================
' UpdateSupportLengths - Recalculate Support Position and Length from References
' 
' Run this rule in an assembly to update all Kask24 support placements based on
' their stored geometry references. This allows supports to update when the
' skeleton or assembly geometry changes.
'
' Updates BOTH position (occurrence transformation) and length (part parameter).
'
' Geometry references are stored per-occurrence (instance), not per-part file.
' This allows the same part file to be reused with different placement geometry.
'
' IMPORTANT: References are work feature names. If work features are renamed
' or deleted, the update will fail for those supports.
'
' Usage:
' - Run manually after skeleton changes
' - Or set up as Event Trigger (on document update/save)
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/SupportPlacementLib.vb"

Sub Main()
    Try
        RunUpdate()
    Catch ex As Exception
        ' Log errors to help debug event trigger issues
        Try
            Dim logPath As String = System.IO.Path.Combine( _
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), _
                "UpdateSupportLengths_Error.log")
            System.IO.File.AppendAllText(logPath, _
                DateTime.Now.ToString() & ": " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf & vbCrLf)
        Catch
        End Try
    End Try
End Sub

Sub RunUpdate()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then Exit Sub
    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then Exit Sub

    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    
    ' Force document update to ensure new parameter values are applied
    ' This is needed when triggered by "Any Parameter Change" event
    Try
        asmDoc.Update2(True) ' True = full update including geometry
    Catch
        ' If update fails, try regular update
        Try
            asmDoc.Update()
        Catch
        End Try
    End Try
    
    Dim updatedCount As Integer = 0
    Dim skippedCount As Integer = 0
    Dim errorCount As Integer = 0
    Dim errorMessages As New System.Text.StringBuilder()
    
    ' Process all occurrences
    ProcessOccurrences(app, asmDoc, asmDef.Occurrences, _
                       updatedCount, skippedCount, errorCount, errorMessages)
    
    ' Only show message if there are errors AND rule was run manually (not from event trigger)
    ' Check if called from event trigger by seeing if there's user interaction expected
    If errorCount > 0 Then
        Dim msg As String = "Update completed with errors." & vbCrLf & vbCrLf & _
            "Updated: " & updatedCount & vbCrLf & _
            "Errors: " & errorCount & vbCrLf & vbCrLf & _
            "Error details:" & vbCrLf & errorMessages.ToString()
        
        MessageBox.Show(msg, "Update Supports - Errors")
    End If
End Sub

' ============================================================================
' Process all occurrences recursively
' ============================================================================
Sub ProcessOccurrences(app As Inventor.Application, asmDoc As AssemblyDocument, _
                       occurrences As ComponentOccurrences, _
                       ByRef updatedCount As Integer, ByRef skippedCount As Integer, _
                       ByRef errorCount As Integer, ByRef errorMessages As System.Text.StringBuilder)
    
    For Each occ As ComponentOccurrence In occurrences
        If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As PartDocument = Nothing
            Try
                partDoc = CType(occ.Definition.Document, PartDocument)
            Catch
                Continue For
            End Try
            
            If partDoc Is Nothing Then Continue For
            
            ' Check if this is a Kask24 support
            If Not SupportPlacementLib.IsKask24Support(partDoc) Then
                Continue For
            End If
            
            ' Check if this occurrence has placement data
            If Not SupportPlacementLib.HasPlacementData(occ) Then
                skippedCount += 1
                Continue For
            End If
            
            ' Try to update this support occurrence
            Dim errorMsg As String = ""
            Dim result As String = UpdateSupportOccurrence(app, asmDoc, partDoc, occ, errorMsg)
            
            Select Case result
                Case "UPDATED", "SYNCED" : updatedCount += 1
                Case "SKIPPED" : skippedCount += 1
                Case "ERROR"
                    errorCount += 1
                    If errorMsg <> "" Then
                        errorMessages.AppendLine("  " & occ.Name & ": " & errorMsg)
                    End If
            End Select
            
        ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            ' Recurse into sub-assemblies
            Try
                Dim subAsmDef As AssemblyComponentDefinition = _
                    CType(occ.Definition, AssemblyComponentDefinition)
                ProcessOccurrences(app, asmDoc, subAsmDef.Occurrences, _
                                   updatedCount, skippedCount, errorCount, errorMessages)
            Catch
            End Try
        End If
    Next
End Sub

' ============================================================================
' Update a single support occurrence from its stored geometry references
' Returns: "UPDATED", "SYNCED", "SKIPPED", or "ERROR"
' ============================================================================
Function UpdateSupportOccurrence(app As Inventor.Application, asmDoc As AssemblyDocument, _
                                  partDoc As PartDocument, occ As ComponentOccurrence, _
                                  ByRef errorMsg As String) As String
    
    errorMsg = ""
    
    ' Read stored references from the occurrence
    Dim refs As Dictionary(Of String, String) = SupportPlacementLib.GetOccurrenceReferences(occ)
    
    Dim mode As String = ""
    Dim ref1Name As String = ""
    Dim ref2Name As String = ""
    Dim ref3Name As String = ""
    Dim alignPoint As String = "Origin"
    Dim orientMode As String = ""
    Dim orientRefName As String = ""
    Dim lengthInput As String = ""
    Dim flipDirectionStr As String = "False"
    Dim offsetXStr As String = ""
    Dim offsetYStr As String = ""
    Dim offsetZStr As String = ""
    
    If refs.ContainsKey("Mode") Then mode = refs("Mode")
    If refs.ContainsKey("Ref1") Then ref1Name = refs("Ref1")
    If refs.ContainsKey("Ref2") Then ref2Name = refs("Ref2")
    If refs.ContainsKey("Ref3") Then ref3Name = refs("Ref3")
    If refs.ContainsKey("AlignPoint") Then alignPoint = refs("AlignPoint")
    If refs.ContainsKey("OrientMode") Then orientMode = refs("OrientMode")
    If refs.ContainsKey("OrientRef") Then orientRefName = refs("OrientRef")
    If refs.ContainsKey("LengthInput") Then lengthInput = refs("LengthInput")
    If refs.ContainsKey("FlipDirection") Then flipDirectionStr = refs("FlipDirection")
    If refs.ContainsKey("OffsetX") Then offsetXStr = refs("OffsetX")
    If refs.ContainsKey("OffsetY") Then offsetYStr = refs("OffsetY")
    If refs.ContainsKey("OffsetZ") Then offsetZStr = refs("OffsetZ")
    Dim customName As String = ""
    If refs.ContainsKey("CustomName") Then customName = refs("CustomName")
    
    ' Skip if no mode stored
    If mode = "" Then
        Return "SKIPPED"
    End If
    
    ' Resolve length input (can be number in mm or parameter name)
    Dim manualLength As Double = SupportPlacementLib.ResolveLengthInput(asmDoc, lengthInput)
    Dim flipDirection As Boolean = False
    Boolean.TryParse(flipDirectionStr, flipDirection)
    
    ' Resolve offsets (can be number in mm or parameter name)
    Dim offsetXCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetXStr)
    Dim offsetYCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetYStr)
    Dim offsetZCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetZStr)
    
    ' Resolve work feature references by name
    Dim ref1 As Object = SupportPlacementLib.ResolveWorkFeatureByName(asmDoc, ref1Name)
    Dim ref2 As Object = SupportPlacementLib.ResolveWorkFeatureByName(asmDoc, ref2Name)
    Dim ref3 As Object = SupportPlacementLib.ResolveWorkFeatureByName(asmDoc, ref3Name)
    
    ' Check if required references were found
    Dim ref1Required As Boolean = (mode = "TWO_POINTS" OrElse mode = "AXIS_TWO_PLANES" OrElse _
                                    mode = "PLANE_AXIS_LENGTH" OrElse mode = "POINT_AXIS_LENGTH" OrElse _
                                    mode = "TWO_PLANES_POINT")
    Dim ref2Required As Boolean = ref1Required
    Dim ref3Required As Boolean = (mode = "AXIS_TWO_PLANES" OrElse mode = "TWO_PLANES_POINT")
    
    If ref1Required AndAlso ref1 Is Nothing Then
        errorMsg = "Work feature '" & ref1Name & "' not found"
        Return "ERROR"
    End If
    If ref2Required AndAlso ref2 Is Nothing Then
        errorMsg = "Work feature '" & ref2Name & "' not found"
        Return "ERROR"
    End If
    If ref3Required AndAlso ref3 Is Nothing Then
        errorMsg = "Work feature '" & ref3Name & "' not found"
        Return "ERROR"
    End If
    
    ' Calculate placement
    Dim startPoint As Point = Nothing
    Dim direction As UnitVector = Nothing
    Dim newLength As Double = 0
    
    Dim success As Boolean = SupportPlacementLib.CalculatePlacement( _
        app, mode, ref1, ref2, ref3, manualLength, flipDirection, _
        startPoint, direction, newLength, errorMsg)
    
    If Not success Then
        If errorMsg = "" Then errorMsg = "Failed to calculate placement"
        Return "ERROR"
    End If
    
    ' Get current values
    Dim currentLength As Double = 0
    Dim currentMatrix As Matrix = Nothing
    Try
        currentLength = partDoc.ComponentDefinition.Parameters.Item("Length").Value
        currentMatrix = occ.Transformation
    Catch ex As Exception
        errorMsg = "Could not read current state: " & ex.Message
        Return "ERROR"
    End Try
    
    ' Get width from part
    Dim widthMm As Integer = 24
    Try
        Dim params As Parameters = partDoc.ComponentDefinition.Parameters
        widthMm = CInt(Math.Round(params.Item("Width").Value * 10))
    Catch
    End Try
    
    ' Resolve orientation reference
    Dim orientRef As Object = Nothing
    If orientRefName <> "" Then
        orientRef = SupportPlacementLib.ResolveWorkFeatureByName(asmDoc, orientRefName)
    End If
    
    ' Calculate new placement matrix (including offsets)
    Dim newMatrix As Matrix = SupportPlacementLib.CreateFullPlacementMatrix( _
        app, startPoint, direction, alignPoint, widthMm, orientMode, orientRef, _
        offsetXCm, offsetYCm, offsetZCm)
    
    ' Check if geometry update is needed
    Dim lengthChanged As Boolean = Math.Abs(newLength - currentLength) > 0.001
    Dim positionChanged As Boolean = Not MatricesEqual(currentMatrix, newMatrix)
    
    ' Update occurrence transformation (position)
    Try
        If positionChanged Then
            occ.Transformation = newMatrix
        End If
    Catch ex As Exception
        errorMsg = "Could not update position: " & ex.Message
        Return "ERROR"
    End Try
    
    ' Update part length parameter if changed
    Try
        If lengthChanged Then
            SupportPlacementLib.UpdateSupportLength(partDoc, newLength)
        End If
    Catch ex As Exception
        errorMsg = "Could not update length: " & ex.Message
        Return "ERROR"
    End Try
    
    ' Always detect renames and sync names (even if dimensions unchanged)
    Try
        ' Detect if user renamed occurrence and adopt that name
        SupportPlacementLib.DetectAndAdoptOccurrenceRename(asmDoc, partDoc)
        
        ' Update iProperties (respects custom Part Number if set)
        Dim lengthMm As Integer = CInt(Math.Round(newLength * 10))
        SupportPlacementLib.UpdateSupportiProperties(partDoc, widthMm, lengthMm, customName)
        
        ' Try to save - may fail for pattern members, which is OK
        Try
            partDoc.Save()
        Catch
            ' Ignore save errors (e.g., pattern members can't be saved directly)
        End Try
        
        ' Sync all occurrence names to match Part Number
        SupportPlacementLib.SyncOccurrenceNames(asmDoc, partDoc)
    Catch ex As Exception
        errorMsg = "Could not update properties: " & ex.Message
        Return "ERROR"
    End Try
    
    If lengthChanged OrElse positionChanged Then
        Return "UPDATED"
    Else
        Return "SYNCED" ' Only names/properties were synced
    End If
End Function

' ============================================================================
' Helper function to compare two matrices
' ============================================================================
Function MatricesEqual(m1 As Matrix, m2 As Matrix) As Boolean
    If m1 Is Nothing OrElse m2 Is Nothing Then Return False
    
    Dim tolerance As Double = 0.0001
    
    For i As Integer = 1 To 4
        For j As Integer = 1 To 4
            If Math.Abs(m1.Cell(i, j) - m2.Cell(i, j)) > tolerance Then
                Return False
            End If
        Next
    Next
    
    Return True
End Function
