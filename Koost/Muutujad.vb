' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Muutujad - Expose child component parameters as assembly-level parameters
'
' This script allows you to select parameters from child parts/subassemblies
' and expose them as editable parameters at the assembly level. Changes to
' the assembly parameters automatically update the child component parameters.
'
' Usage:
' - Run from an open assembly document
' - Select parameters to expose from the tree view
' - Optionally rename parameters to avoid conflicts
' - Click "Rakenda" to create the parameters and update rule
'
' Running multiple times allows you to modify existing exposed parameters.
' ============================================================================

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/DocumentUpdateLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

' ============================================================================
' DATA STRUCTURES
' ============================================================================

Class ExposedParameter
    Public ComponentName As String      ' e.g., "Skeleton.ipt" or "Frame:1"
    Public OriginalParamName As String  ' Parameter name in the child component
    Public ExposedName As String        ' Name to use in assembly (may differ for conflicts)
    Public IsSelected As Boolean        ' Whether to expose this parameter
    Public CurrentValue As Double       ' Current value for display
    Public Units As String              ' Units string for display
    Public HasConflict As Boolean       ' True if name conflicts with another parameter
End Class

Class ComponentInfo
    Public DisplayName As String        ' User-friendly display name
    Public ReferenceName As String      ' Name used in Parameter() function
    Public Parameters As List(Of ExposedParameter)
    
    Public Sub New()
        Parameters = New List(Of ExposedParameter)
    End Sub
End Class

' ============================================================================
' MAIN
' ============================================================================

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Muutujad: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Muutujad")
        Exit Sub
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        UtilsLib.LogError("Muutujad: Active document is not an assembly")
        MessageBox.Show("Aktiivseks dokumendiks peab olema koost (.iam).", "Muutujad")
        Exit Sub
    End If
    
    Dim asmDoc As AssemblyDocument = CType(app.ActiveDocument, AssemblyDocument)
    Dim iLogicAuto As Object = iLogicVb.Automation
    
    UtilsLib.LogInfo("Muutujad: Starting for " & asmDoc.DisplayName)
    
    ' Gather all components and their parameters
    Dim components As List(Of ComponentInfo) = GatherComponentParameters(asmDoc)
    
    If components.Count = 0 Then
        UtilsLib.LogWarn("Muutujad: No components with user parameters found")
        MessageBox.Show("Koostis puuduvad kasutaja parameetritega komponendid.", "Muutujad")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Muutujad: Found " & components.Count & " components with parameters")
    
    ' Load existing exposed parameters (if any)
    LoadExistingExposedParameters(asmDoc, components)
    
    ' Check for name conflicts
    CheckForConflicts(components)
    
    ' Show UI for parameter selection
    Dim result As DialogResult = ShowParameterSelectionForm(components, asmDoc)
    
    If result <> DialogResult.OK Then
        UtilsLib.LogInfo("Muutujad: User cancelled")
        Exit Sub
    End If
    
    ' Apply changes
    ApplyExposedParameters(asmDoc, iLogicAuto, components)
    
    UtilsLib.LogInfo("Muutujad: Completed successfully")
End Sub

' ============================================================================
' GATHER COMPONENT PARAMETERS
' ============================================================================

Function GatherComponentParameters(asmDoc As AssemblyDocument) As List(Of ComponentInfo)
    Dim result As New List(Of ComponentInfo)
    Dim processedDocs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    
    ' Process all occurrences recursively
    For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
        GatherFromOccurrence(occ, result, processedDocs)
    Next
    
    Return result
End Function

Sub GatherFromOccurrence(occ As ComponentOccurrence, result As List(Of ComponentInfo), processedDocs As HashSet(Of String))
    If occ Is Nothing Then Exit Sub
    
    Try
        Dim refDoc As Document = occ.Definition.Document
        If refDoc Is Nothing Then Exit Sub
        
        ' Skip if already processed (same document may appear multiple times)
        Dim docKey As String = refDoc.FullDocumentName
        If processedDocs.Contains(docKey) Then Exit Sub
        processedDocs.Add(docKey)
        
        ' Get the reference name for Parameter() function
        Dim refName As String = System.IO.Path.GetFileName(refDoc.FullDocumentName)
        
        ' Get user parameters
        Dim params As Parameters = Nothing
        If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            params = CType(refDoc, PartDocument).ComponentDefinition.Parameters
        ElseIf refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            params = CType(refDoc, AssemblyDocument).ComponentDefinition.Parameters
        End If
        
        If params IsNot Nothing AndAlso params.UserParameters.Count > 0 Then
            Dim compInfo As New ComponentInfo()
            compInfo.DisplayName = GetDisplayName(refDoc)
            compInfo.ReferenceName = refName
            
            For Each userParam As Parameter In params.UserParameters
                Dim expParam As New ExposedParameter()
                expParam.ComponentName = refName
                expParam.OriginalParamName = userParam.Name
                expParam.ExposedName = SanitizeParamName(userParam.Name)  ' Use simple name by default
                expParam.IsSelected = False
                expParam.CurrentValue = userParam.Value
                expParam.Units = GetUnitsString(userParam)
                expParam.HasConflict = False
                compInfo.Parameters.Add(expParam)
            Next
            
            result.Add(compInfo)
        End If
        
        ' Process sub-occurrences for subassemblies
        If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim subAsm As AssemblyDocument = CType(refDoc, AssemblyDocument)
            For Each subOcc As ComponentOccurrence In subAsm.ComponentDefinition.Occurrences
                GatherFromOccurrence(subOcc, result, processedDocs)
            Next
        End If
    Catch ex As Exception
        ' Skip components that can't be accessed
    End Try
End Sub

Function GetDisplayName(doc As Document) As String
    Try
        Dim desc As String = ""
        Dim pn As String = ""
        
        Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
        Try
            desc = CStr(designProps.Item("Description").Value)
        Catch
        End Try
        Try
            pn = CStr(designProps.Item("Part Number").Value)
        Catch
        End Try
        
        If Not String.IsNullOrEmpty(desc) AndAlso Not String.IsNullOrEmpty(pn) Then
            Return desc & " (" & pn & ")"
        ElseIf Not String.IsNullOrEmpty(desc) Then
            Return desc
        ElseIf Not String.IsNullOrEmpty(pn) Then
            Return pn
        Else
            Return System.IO.Path.GetFileNameWithoutExtension(doc.FullDocumentName)
        End If
    Catch
        Return System.IO.Path.GetFileNameWithoutExtension(doc.FullDocumentName)
    End Try
End Function

Function SanitizeParamName(paramName As String) As String
    ' Use UtilsLib for proper sanitization of parameter names
    ' (handles parentheses, special characters, and digit-start issues)
    Return UtilsLib.SanitizeParameterName(paramName)
End Function

Function SanitizeParamNameWithPrefix(componentName As String, paramName As String) As String
    ' Add component prefix for conflict resolution
    Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(componentName)
    ' Use UtilsLib for proper sanitization (handles parentheses, spaces, etc.)
    Return UtilsLib.SanitizeParameterName(baseName & "_" & paramName)
End Function

Function GetUnitsString(param As Parameter) As String
    Try
        If param.Units Is Nothing OrElse String.IsNullOrEmpty(param.Units.ToString()) Then
            Return ""
        End If
        Return param.Units.ToString()
    Catch
        Return ""
    End Try
End Function

' ============================================================================
' LOAD EXISTING EXPOSED PARAMETERS
' ============================================================================

Sub LoadExistingExposedParameters(asmDoc As AssemblyDocument, components As List(Of ComponentInfo))
    ' Check which parameters are already exposed by looking at assembly user parameters
    ' and checking if they have formulas referencing child components
    
    Dim asmParams As Parameters = asmDoc.ComponentDefinition.Parameters
    
    ' Also parse the existing Uuenda rule to find which parameters are being set
    Dim existingMappings As Dictionary(Of String, String) = ParseExistingMappings(asmDoc)
    
    For Each comp As ComponentInfo In components
        For Each expParam As ExposedParameter In comp.Parameters
            ' Check if there's an assembly parameter that maps to this child parameter
            Dim key As String = comp.ReferenceName & "|" & expParam.OriginalParamName
            If existingMappings.ContainsKey(key) Then
                expParam.IsSelected = True
                expParam.ExposedName = existingMappings(key)
            End If
        Next
    Next
End Sub

Function ParseExistingMappings(asmDoc As AssemblyDocument) As Dictionary(Of String, String)
    Dim mappings As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    
    Try
        Dim iLogicAuto As Object = iLogicVb.Automation
        Dim rule As Object = iLogicAuto.GetRule(asmDoc, DocumentUpdateLib.RULE_NAME)
        If rule Is Nothing Then Return mappings
        
        Dim ruleText As String = CStr(CallByName(rule, "Text", CallType.Get))
        
        ' Parse lines like: Parameter("Skeleton.ipt", "Width") = SomeParamName
        Dim lines() As String = ruleText.Split({vbCrLf, vbLf}, StringSplitOptions.None)
        For Each line As String In lines
            Dim trimmed As String = line.Trim()
            If trimmed.StartsWith("Parameter(") AndAlso trimmed.Contains("=") Then
                ' Extract component name, param name, and assembly param name
                Try
                    Dim match As System.Text.RegularExpressions.Match = _
                        System.Text.RegularExpressions.Regex.Match(trimmed, _
                            "Parameter\s*\(\s*""([^""]+)""\s*,\s*""([^""]+)""\s*\)\s*=\s*(\w+)")
                    If match.Success Then
                        Dim compName As String = match.Groups(1).Value
                        Dim paramName As String = match.Groups(2).Value
                        Dim asmParam As String = match.Groups(3).Value
                        Dim key As String = compName & "|" & paramName
                        mappings(key) = asmParam
                    End If
                Catch
                End Try
            End If
        Next
    Catch
    End Try
    
    Return mappings
End Function

' ============================================================================
' CHECK FOR CONFLICTS
' ============================================================================

Sub CheckForConflicts(components As List(Of ComponentInfo))
    ' Build a map of all exposed names to detect conflicts
    Dim nameCount As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    
    For Each comp As ComponentInfo In components
        For Each expParam As ExposedParameter In comp.Parameters
            Dim name As String = expParam.ExposedName
            If nameCount.ContainsKey(name) Then
                nameCount(name) += 1
            Else
                nameCount(name) = 1
            End If
        Next
    Next
    
    ' For conflicting names, add component prefix to make them unique
    For Each comp As ComponentInfo In components
        For Each expParam As ExposedParameter In comp.Parameters
            If nameCount(expParam.ExposedName) > 1 Then
                expParam.HasConflict = True
                ' Auto-add component prefix for conflicts
                expParam.ExposedName = SanitizeParamNameWithPrefix(expParam.ComponentName, expParam.OriginalParamName)
            Else
                expParam.HasConflict = False
            End If
        Next
    Next
End Sub

' ============================================================================
' UI
' ============================================================================

Function ShowParameterSelectionForm(components As List(Of ComponentInfo), asmDoc As AssemblyDocument) As DialogResult
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Muutujad - Ekspordi parameetrid"
    frm.Width = 800
    frm.Height = 600
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.MinimizeBox = False
    frm.MaximizeBox = True
    
    ' Instructions label
    Dim lblInfo As New System.Windows.Forms.Label()
    lblInfo.Text = "Vali parameetrid, mida soovid koostu tasemel muudetavaks muuta. " & _
                   "Nimekonflikti korral muuda parameetri nime."
    lblInfo.Left = 10
    lblInfo.Top = 10
    lblInfo.Width = 760
    lblInfo.Height = 40
    lblInfo.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
    frm.Controls.Add(lblInfo)
    
    ' DataGridView for parameter selection
    Dim dgv As New DataGridView()
    dgv.Left = 10
    dgv.Top = 55
    dgv.Width = 760
    dgv.Height = 450
    dgv.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    dgv.AllowUserToAddRows = False
    dgv.AllowUserToDeleteRows = False
    dgv.RowHeadersVisible = False
    dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    dgv.MultiSelect = False
    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    
    ' Column: Selected (checkbox)
    Dim colSelected As New DataGridViewCheckBoxColumn()
    colSelected.Name = "colSelected"
    colSelected.HeaderText = "Vali"
    colSelected.Width = 45
    dgv.Columns.Add(colSelected)
    
    ' Column: Component (read-only)
    Dim colComponent As New DataGridViewTextBoxColumn()
    colComponent.Name = "colComponent"
    colComponent.HeaderText = "Komponent"
    colComponent.Width = 200
    colComponent.ReadOnly = True
    dgv.Columns.Add(colComponent)
    
    ' Column: Original Parameter Name (read-only)
    Dim colOriginal As New DataGridViewTextBoxColumn()
    colOriginal.Name = "colOriginal"
    colOriginal.HeaderText = "Algne parameeter"
    colOriginal.Width = 120
    colOriginal.ReadOnly = True
    dgv.Columns.Add(colOriginal)
    
    ' Column: Exposed Name (editable)
    Dim colExposed As New DataGridViewTextBoxColumn()
    colExposed.Name = "colExposed"
    colExposed.HeaderText = "Koostu parameeter"
    colExposed.Width = 180
    dgv.Columns.Add(colExposed)
    
    ' Column: Current Value (read-only)
    Dim colValue As New DataGridViewTextBoxColumn()
    colValue.Name = "colValue"
    colValue.HeaderText = "Väärtus"
    colValue.Width = 100
    colValue.ReadOnly = True
    dgv.Columns.Add(colValue)
    
    ' Column: Units (read-only)
    Dim colUnits As New DataGridViewTextBoxColumn()
    colUnits.Name = "colUnits"
    colUnits.HeaderText = "Ühikud"
    colUnits.Width = 80
    colUnits.ReadOnly = True
    dgv.Columns.Add(colUnits)
    
    ' Populate rows
    Dim rowMap As New Dictionary(Of Integer, ExposedParameter)
    Dim rowIndex As Integer = 0
    
    For Each comp As ComponentInfo In components
        For Each expParam As ExposedParameter In comp.Parameters
            Dim idx As Integer = dgv.Rows.Add()
            dgv.Rows(idx).Cells("colSelected").Value = expParam.IsSelected
            dgv.Rows(idx).Cells("colComponent").Value = comp.DisplayName
            dgv.Rows(idx).Cells("colOriginal").Value = expParam.OriginalParamName
            dgv.Rows(idx).Cells("colExposed").Value = expParam.ExposedName
            dgv.Rows(idx).Cells("colValue").Value = FormatValue(expParam.CurrentValue, expParam.Units)
            dgv.Rows(idx).Cells("colUnits").Value = expParam.Units
            
            ' Mark conflicts with special text
            If expParam.HasConflict Then
                dgv.Rows(idx).Cells("colExposed").Value = "* " & expParam.ExposedName
            End If
            
            dgv.Rows(idx).Tag = expParam
            rowIndex += 1
        Next
    Next
    
    frm.Controls.Add(dgv)
    
    ' Select All button
    Dim btnSelectAll As New System.Windows.Forms.Button()
    btnSelectAll.Text = "Vali kõik"
    btnSelectAll.Left = 10
    btnSelectAll.Top = 515
    btnSelectAll.Width = 100
    btnSelectAll.Height = 30
    btnSelectAll.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    AddHandler btnSelectAll.Click, Sub(s, e)
        For Each row As DataGridViewRow In dgv.Rows
            row.Cells("colSelected").Value = True
        Next
    End Sub
    frm.Controls.Add(btnSelectAll)
    
    ' Deselect All button
    Dim btnDeselectAll As New System.Windows.Forms.Button()
    btnDeselectAll.Text = "Tühista valik"
    btnDeselectAll.Left = 120
    btnDeselectAll.Top = 515
    btnDeselectAll.Width = 100
    btnDeselectAll.Height = 30
    btnDeselectAll.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    AddHandler btnDeselectAll.Click, Sub(s, e)
        For Each row As DataGridViewRow In dgv.Rows
            row.Cells("colSelected").Value = False
        Next
    End Sub
    frm.Controls.Add(btnDeselectAll)
    
    ' OK button
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Rakenda"
    btnOK.Left = 580
    btnOK.Top = 515
    btnOK.Width = 90
    btnOK.Height = 30
    btnOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    AddHandler btnOK.Click, Sub(s, e)
        ' Sync UI to data (strip conflict marker if present)
        For Each row As DataGridViewRow In dgv.Rows
            Dim expParam As ExposedParameter = CType(row.Tag, ExposedParameter)
            expParam.IsSelected = CBool(row.Cells("colSelected").Value)
            Dim rawName As String = CStr(row.Cells("colExposed").Value)
            expParam.ExposedName = If(rawName.StartsWith("* "), rawName.Substring(2), rawName)
        Next
        
        ' Validate names
        Dim errors As String = ValidateParameterNames(components)
        If Not String.IsNullOrEmpty(errors) Then
            MessageBox.Show(errors, "Viga", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        
        frm.DialogResult = DialogResult.OK
        frm.Close()
    End Sub
    frm.Controls.Add(btnOK)
    
    ' Cancel button
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 680
    btnCancel.Top = 515
    btnCancel.Width = 90
    btnCancel.Height = 30
    btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    AddHandler btnCancel.Click, Sub(s, e)
        frm.DialogResult = DialogResult.Cancel
        frm.Close()
    End Sub
    frm.Controls.Add(btnCancel)
    
    ' Handle cell value changes to update conflict highlighting
    AddHandler dgv.CellValueChanged, Sub(s, e)
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = dgv.Columns("colExposed").Index Then
            ' Re-check conflicts when name changes
            UpdateConflictHighlighting(dgv, components)
        End If
    End Sub
    
    AddHandler dgv.CurrentCellDirtyStateChanged, Sub(s, e)
        If dgv.IsCurrentCellDirty Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
    
    Return frm.ShowDialog()
End Function

Function FormatValue(value As Double, units As String) As String
    ' Format value with appropriate precision based on units
    If String.IsNullOrEmpty(units) Then
        Return value.ToString("0.###")
    ElseIf units.Contains("mm") OrElse units.Contains("cm") OrElse units.Contains("in") Then
        ' Length - convert from internal units (cm) to display
        Return (value * 10).ToString("0.##") & " mm"
    ElseIf units.Contains("deg") Then
        ' Angle - convert from radians to degrees
        Return (value * 180 / Math.PI).ToString("0.##") & "°"
    Else
        Return value.ToString("0.###") & " " & units
    End If
End Function

Sub UpdateConflictHighlighting(dgv As DataGridView, components As List(Of ComponentInfo))
    ' Build map of exposed names from current UI state (strip leading * for comparison)
    Dim nameRows As New Dictionary(Of String, List(Of DataGridViewRow))(StringComparer.OrdinalIgnoreCase)
    
    For Each row As DataGridViewRow In dgv.Rows
        Dim rawName As String = CStr(row.Cells("colExposed").Value)
        Dim name As String = If(rawName.StartsWith("* "), rawName.Substring(2), rawName)
        If Not nameRows.ContainsKey(name) Then
            nameRows(name) = New List(Of DataGridViewRow)
        End If
        nameRows(name).Add(row)
    Next
    
    ' Update conflict markers (use * prefix to indicate conflicts)
    For Each row As DataGridViewRow In dgv.Rows
        Dim rawName As String = CStr(row.Cells("colExposed").Value)
        Dim name As String = If(rawName.StartsWith("* "), rawName.Substring(2), rawName)
        
        If nameRows(name).Count > 1 Then
            If Not rawName.StartsWith("* ") Then
                row.Cells("colExposed").Value = "* " & name
            End If
        Else
            If rawName.StartsWith("* ") Then
                row.Cells("colExposed").Value = name
            End If
        End If
    Next
End Sub

Function ValidateParameterNames(components As List(Of ComponentInfo)) As String
    Dim errors As New List(Of String)
    Dim selectedNames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    
    For Each comp As ComponentInfo In components
        For Each expParam As ExposedParameter In comp.Parameters
            If Not expParam.IsSelected Then Continue For
            
            ' Check for empty names
            If String.IsNullOrWhiteSpace(expParam.ExposedName) Then
                errors.Add("Parameeter '" & expParam.OriginalParamName & "' (" & comp.DisplayName & ") nime ei tohi olla tühi.")
                Continue For
            End If
            
            ' Check for invalid characters
            If Not System.Text.RegularExpressions.Regex.IsMatch(expParam.ExposedName, "^[a-zA-Z_][a-zA-Z0-9_]*$") Then
                errors.Add("Parameeter '" & expParam.ExposedName & "' sisaldab keelatud märke. Kasuta ainult tähti, numbreid ja alakriipse.")
                Continue For
            End If
            
            ' Check for duplicates among selected parameters
            If selectedNames.ContainsKey(expParam.ExposedName) Then
                errors.Add("Parameetri nimi '" & expParam.ExposedName & "' on juba kasutusel (" & _
                          selectedNames(expParam.ExposedName) & ").")
                Continue For
            End If
            
            selectedNames(expParam.ExposedName) = comp.DisplayName & " / " & expParam.OriginalParamName
        Next
    Next
    
    If errors.Count > 0 Then
        Return String.Join(vbCrLf, errors)
    End If
    
    Return ""
End Function

' ============================================================================
' APPLY CHANGES
' ============================================================================

Sub ApplyExposedParameters(asmDoc As AssemblyDocument, iLogicAuto As Object, components As List(Of ComponentInfo))
    Dim selectedParams As New List(Of ExposedParameter)
    
    ' Collect all selected parameters
    For Each comp As ComponentInfo In components
        For Each expParam As ExposedParameter In comp.Parameters
            If expParam.IsSelected Then
                selectedParams.Add(expParam)
            End If
        Next
    Next
    
    If selectedParams.Count = 0 Then
        ' Remove handler if no parameters selected
        DocumentUpdateLib.RemoveUpdateHandler(asmDoc, iLogicAuto, "Muutujad")
        UtilsLib.LogInfo("Muutujad: No parameters selected, removed handler")
        Return
    End If
    
    UtilsLib.LogInfo("Muutujad: Creating/updating " & selectedParams.Count & " exposed parameters")
    
    ' Create/update assembly-level parameters
    Dim userParams As UserParameters = asmDoc.ComponentDefinition.Parameters.UserParameters
    
    For Each expParam As ExposedParameter In selectedParams
        EnsureAssemblyParameter(userParams, expParam)
    Next
    
    ' Build update code
    Dim codeLines As New List(Of String)
    
    For Each expParam As ExposedParameter In selectedParams
        ' Generate: Parameter("ComponentName.ipt", "ParamName") = ExposedName
        Dim codeLine As String = "Parameter(""" & expParam.ComponentName & """, """ & _
                                 expParam.OriginalParamName & """) = " & expParam.ExposedName
        codeLines.Add(codeLine)
    Next
    
    ' Add document update call
    codeLines.Add("InventorVb.DocumentUpdate()")
    
    ' Register update handler with parameter change and save triggers
    Dim triggers() As DocumentUpdateLib.UpdateTrigger = { _
        DocumentUpdateLib.UpdateTrigger.UserParameterChange, _
        DocumentUpdateLib.UpdateTrigger.BeforeSave _
    }
    
    DocumentUpdateLib.RegisterUpdateHandler(asmDoc, iLogicAuto, "Muutujad", codeLines.ToArray(), triggers)
    
    ' Remove assembly parameters that are no longer selected
    RemoveUnselectedParameters(asmDoc, components, selectedParams)
    
    UtilsLib.LogInfo("Muutujad: Update handler registered successfully")
End Sub

Sub EnsureAssemblyParameter(userParams As UserParameters, expParam As ExposedParameter)
    ' Create or update assembly-level parameter
    Try
        Dim existingParam As Parameter = Nothing
        
        ' Check if parameter exists
        Try
            existingParam = userParams.Item(expParam.ExposedName)
        Catch
            existingParam = Nothing
        End Try
        
        If existingParam Is Nothing Then
            ' Create new parameter with current value from child
            userParams.AddByValue(expParam.ExposedName, expParam.CurrentValue, GetUnitsType(expParam.Units))
            UtilsLib.LogInfo("Muutujad: Created parameter '" & expParam.ExposedName & "' = " & expParam.CurrentValue)
        Else
            ' Parameter already exists - keep current value
            UtilsLib.LogInfo("Muutujad: Parameter '" & expParam.ExposedName & "' already exists, keeping value")
        End If
    Catch ex As Exception
        UtilsLib.LogWarn("Muutujad: Failed to create parameter '" & expParam.ExposedName & "': " & ex.Message)
    End Try
End Sub

Function GetUnitsType(unitsStr As String) As UnitsTypeEnum
    If String.IsNullOrEmpty(unitsStr) Then
        Return UnitsTypeEnum.kUnitlessUnits
    End If
    
    unitsStr = unitsStr.ToLower()
    
    If unitsStr.Contains("mm") OrElse unitsStr.Contains("millimeter") Then
        Return UnitsTypeEnum.kMillimeterLengthUnits
    ElseIf unitsStr.Contains("cm") OrElse unitsStr.Contains("centimeter") Then
        Return UnitsTypeEnum.kCentimeterLengthUnits
    ElseIf unitsStr.Contains("m") AndAlso Not unitsStr.Contains("mm") Then
        Return UnitsTypeEnum.kMeterLengthUnits
    ElseIf unitsStr.Contains("in") OrElse unitsStr.Contains("inch") Then
        Return UnitsTypeEnum.kInchLengthUnits
    ElseIf unitsStr.Contains("deg") OrElse unitsStr.Contains("°") Then
        Return UnitsTypeEnum.kDegreeAngleUnits
    ElseIf unitsStr.Contains("rad") Then
        Return UnitsTypeEnum.kRadianAngleUnits
    Else
        Return UnitsTypeEnum.kDefaultDisplayLengthUnits
    End If
End Function

Sub RemoveUnselectedParameters(asmDoc As AssemblyDocument, components As List(Of ComponentInfo), selectedParams As List(Of ExposedParameter))
    ' Get all previously exposed parameter names from the Uuenda rule
    Dim previouslyExposed As Dictionary(Of String, String) = ParseExistingMappings(asmDoc)
    
    ' Build set of currently selected exposed names
    Dim currentlySelected As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    For Each expParam As ExposedParameter In selectedParams
        currentlySelected.Add(expParam.ExposedName)
    Next
    
    ' Find parameters to remove (were exposed before, but not now)
    Dim userParams As UserParameters = asmDoc.ComponentDefinition.Parameters.UserParameters
    
    For Each kvp As KeyValuePair(Of String, String) In previouslyExposed
        Dim asmParamName As String = kvp.Value
        If Not currentlySelected.Contains(asmParamName) Then
            ' This parameter was exposed before but is no longer selected
            Try
                Dim param As Parameter = userParams.Item(asmParamName)
                param.Delete()
                UtilsLib.LogInfo("Muutujad: Removed parameter '" & asmParamName & "'")
            Catch
                ' Parameter might not exist or might be in use
            End Try
        End If
    Next
End Sub
