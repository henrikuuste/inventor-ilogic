' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' PlaceSupport - Parametric Birch Support Placement Tool
' 
' Run this rule in an assembly to place Kask24 birch supports.
'
' Features:
' - Multiple placement modes (Two Points, Axis+Planes, etc.)
' - Orientation options via matrix (no constraints)
' - Automatic part file creation/reuse from template
' - Geometry reference storage for parametric updates
' - iProperty updates for BOM compatibility
' - Modeless dialog allowing Inventor interaction
'
' IMPORTANT: Only work features (WorkPoint, WorkAxis, WorkPlane) can be selected.
' Create work features in your assembly/skeleton before placing supports.
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/SupportPlacementLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        MessageBox.Show("No active document.", "Place Support")
        Exit Sub
    End If

    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This rule only works in assembly documents (.iam).", "Place Support")
        Exit Sub
    End If

    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    
    ' Find template
    Dim templatePath As String = SupportPlacementLib.FindTemplatePath(asmDoc)
    If templatePath = "" Then
        MessageBox.Show("Could not find Kask24_Template.ipt." & vbCrLf & vbCrLf & _
                       "Please ensure Kask24_Template.ipt is in:" & vbCrLf & _
                       "- Assembly folder" & vbCrLf & _
                       "- Templates subfolder" & vbCrLf & _
                       "- Parent folder", "Place Support")
        Exit Sub
    End If
    
    ' Load work points from template for align point dropdown
    Dim templateWorkPoints As String() = SupportPlacementLib.GetTemplateWorkPoints(app, templatePath)
    If templateWorkPoints Is Nothing OrElse templateWorkPoints.Length = 0 Then
        templateWorkPoints = New String() {"Origin"}
    End If
    
    ' Run placement loop
    RunPlacementLoop(app, asmDoc, templatePath, templateWorkPoints)
End Sub

' ============================================================================
' Main Placement Loop with Modeless Dialog
' ============================================================================
Sub RunPlacementLoop(app As Inventor.Application, asmDoc As AssemblyDocument, _
                     templatePath As String, templateWorkPoints As String())
    
    ' State variables
    Dim selectedWidth As Integer = 24
    Dim placementMode As String = "AXIS_TWO_PLANES"
    Dim orientMode As String = "ALIGN_BOTTOM"
    Dim alignPoint As String = "Origin"
    Dim lengthInput As String = "500" ' Length input: numeric (mm) or parameter name
    Dim flipDirection As Boolean = False
    Dim offsetX As String = "" ' Offset X: numeric (mm) or parameter name
    Dim offsetY As String = "" ' Offset Y: numeric (mm) or parameter name
    Dim offsetZ As String = "" ' Offset Z: numeric (mm) or parameter name
    Dim customName As String = "" ' Custom name for occurrence and file (optional)
    
    ' Form position/size (persisted across dialog reopens and rule runs)
    Dim formLeft As Integer = -1
    Dim formTop As Integer = -1
    Dim formWidth As Integer = 520
    Dim formHeight As Integer = 720
    
    ' Load saved form position from assembly iProperties
    LoadFormPosition(asmDoc, formLeft, formTop, formWidth, formHeight)
    
    ' Geometry references (work features only)
    Dim ref1 As Object = Nothing
    Dim ref2 As Object = Nothing
    Dim ref3 As Object = Nothing
    Dim orientRef As Object = Nothing
    
    ' Calculated values
    Dim startPoint As Point = Nothing
    Dim direction As UnitVector = Nothing
    Dim calculatedLength As Double = 0
    
    ' Modify mode - tracks occurrence being modified (Nothing = create new)
    Dim modifyOcc As ComponentOccurrence = Nothing
    
    Dim keepPlacing As Boolean = True
    
    Do While keepPlacing
        Dim action As String = ""
        Dim result As DialogResult = ShowPlacementForm(app, asmDoc, templateWorkPoints, _
            selectedWidth, placementMode, orientMode, alignPoint, lengthInput, flipDirection, _
            ref1, ref2, ref3, orientRef, calculatedLength, modifyOcc, action, _
            formLeft, formTop, formWidth, formHeight, offsetX, offsetY, offsetZ, customName)
        
        If result = DialogResult.Cancel Then
            keepPlacing = False
            
        ElseIf action = "PICK_PARAM" Then
            ' Show parameter picker dialog (handle before StartsWith check)
            Dim paramName As String = PickParameter(asmDoc)
            If paramName <> "" Then
                lengthInput = paramName
            End If
            
        ElseIf action.StartsWith("PICK_") OrElse action.StartsWith("CLEAR_") Then
            ' Handle pick and clear actions
            Dim pickResult As Object = Nothing
            
            Select Case action
                Case "PICK_REF1"
                    pickResult = PickForMode(app, placementMode, 1)
                    If pickResult IsNot Nothing Then ref1 = pickResult
                    
                Case "PICK_REF2"
                    pickResult = PickForMode(app, placementMode, 2)
                    If pickResult IsNot Nothing Then ref2 = pickResult
                    
                Case "PICK_REF3"
                    pickResult = PickForMode(app, placementMode, 3)
                    If pickResult IsNot Nothing Then ref3 = pickResult
                    
                Case "PICK_ORIENT"
                    pickResult = PickForOrient(app, orientMode)
                    If pickResult IsNot Nothing Then orientRef = pickResult
                    
                Case "CLEAR_REF1"
                    ref1 = Nothing
                    
                Case "CLEAR_REF2"
                    ref2 = Nothing
                    
                Case "CLEAR_REF3"
                    ref3 = Nothing
                    
                Case "CLEAR_ORIENT"
                    orientRef = Nothing
            End Select
            
            ' Recalculate after valid pick or clear
            If ref1 IsNot Nothing OrElse ref2 IsNot Nothing Then
                Dim errMsg As String = ""
                Dim resolvedLen As Double = SupportPlacementLib.ResolveLengthInput(asmDoc, lengthInput)
                SupportPlacementLib.CalculatePlacement(app, placementMode, ref1, ref2, ref3, _
                    resolvedLen, flipDirection, startPoint, direction, calculatedLength, errMsg)
            End If
            
        ElseIf action = "MODE_CHANGED" Then
            ' Clear references when mode changes
            ref1 = Nothing
            ref2 = Nothing
            ref3 = Nothing
            startPoint = Nothing
            direction = Nothing
            calculatedLength = 0
            
        ElseIf action = "MODIFY" Then
            ' Pick an existing support to modify
            Dim pickedOcc As ComponentOccurrence = PickExistingSupport(app, asmDoc)
            If pickedOcc IsNot Nothing Then
                modifyOcc = pickedOcc
                ' Load settings from the occurrence
                LoadOccurrenceSettings(asmDoc, modifyOcc, _
                    selectedWidth, placementMode, orientMode, alignPoint, _
                    lengthInput, flipDirection, ref1, ref2, ref3, orientRef, _
                    offsetX, offsetY, offsetZ, customName)
                ' Recalculate
                Dim errMsg As String = ""
                Dim resolvedLen As Double = SupportPlacementLib.ResolveLengthInput(asmDoc, lengthInput)
                SupportPlacementLib.CalculatePlacement(app, placementMode, ref1, ref2, ref3, _
                    resolvedLen, flipDirection, startPoint, direction, calculatedLength, errMsg)
            End If

        ElseIf action = "NEW" Then
            ' Switch back to create new mode
            modifyOcc = Nothing
            ref1 = Nothing
            ref2 = Nothing
            ref3 = Nothing
            orientRef = Nothing
            startPoint = Nothing
            direction = Nothing
            calculatedLength = 0
            
        ElseIf action = "PLACE" OrElse action = "PLACE_CLOSE" OrElse action = "UPDATE" OrElse action = "UPDATE_CLOSE" Then
            ' Validate and place/update
            Dim resolvedLen As Double = SupportPlacementLib.ResolveLengthInput(asmDoc, lengthInput)
            If Not ValidatePlacement(placementMode, ref1, ref2, ref3, resolvedLen) Then
                ' Provide detailed error message
                Dim errorDetails As String = "Missing:" & vbCrLf
                If ref1 Is Nothing Then errorDetails &= "  - Reference 1" & vbCrLf
                If ref2 Is Nothing Then errorDetails &= "  - Reference 2" & vbCrLf
                If NeedsRef3(placementMode) AndAlso ref3 Is Nothing Then errorDetails &= "  - Reference 3" & vbCrLf
                If NeedsManualLength(placementMode) AndAlso resolvedLen <= 0 Then
                    errorDetails &= "  - Length ('" & lengthInput & "' resolved to " & resolvedLen & ")" & vbCrLf
                End If
                MessageBox.Show("Please complete all required selections." & vbCrLf & vbCrLf & errorDetails, "Place Support")
            Else
                ' Calculate placement
                Dim errMsg As String = ""
                Dim success As Boolean = SupportPlacementLib.CalculatePlacement(app, placementMode, _
                    ref1, ref2, ref3, resolvedLen, flipDirection, _
                    startPoint, direction, calculatedLength, errMsg)
                
                If Not success Then
                    MessageBox.Show("Could not calculate placement: " & errMsg, "Place Support")
                Else
                    Try
                        If (action = "UPDATE" OrElse action = "UPDATE_CLOSE") AndAlso modifyOcc IsNot Nothing Then
                            ' Update existing support
                            UpdateExistingSupport(app, asmDoc, modifyOcc, _
                                selectedWidth, calculatedLength, lengthInput, _
                                startPoint, direction, _
                                alignPoint, orientMode, orientRef, _
                                placementMode, ref1, ref2, ref3, flipDirection, _
                                offsetX, offsetY, offsetZ, customName)
                            ' Force document update so changes are visible
                            asmDoc.Update()
                            ' Keep modifyOcc so user can update again
                        Else
                            ' Place new support and auto-select it for modification
                            Dim newOcc As ComponentOccurrence = PlaceSupportWithSettings(app, asmDoc, templatePath, _
                                selectedWidth, calculatedLength, lengthInput, _
                                startPoint, direction, _
                                alignPoint, orientMode, orientRef, _
                                placementMode, ref1, ref2, ref3, flipDirection, _
                                offsetX, offsetY, offsetZ, customName)
                            
                            ' Auto-select newly placed support for easy modification
                            If newOcc IsNot Nothing Then
                                modifyOcc = newOcc
                            End If
                        End If
                        
                        If action = "PLACE_CLOSE" OrElse action = "UPDATE_CLOSE" Then
                            keepPlacing = False
                        End If
                    Catch ex As Exception
                        MessageBox.Show("Error: " & ex.Message, "Place Support")
                    End Try
                End If
            End If
            
        ElseIf action = "CLEAR" Then
            ref1 = Nothing
            ref2 = Nothing
            ref3 = Nothing
            orientRef = Nothing
            startPoint = Nothing
            direction = Nothing
            calculatedLength = 0
            modifyOcc = Nothing
        End If
    Loop
    
    ' Save form position to assembly iProperties for next run
    SaveFormPosition(asmDoc, formLeft, formTop, formWidth, formHeight)
End Sub

' ============================================================================
' Pick an existing support occurrence to modify
' ============================================================================
Function PickExistingSupport(app As Inventor.Application, asmDoc As AssemblyDocument) As ComponentOccurrence
    Try
        Dim picked As Object = app.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, _
            "Select an existing Kask24 support to modify:")
        
        If picked Is Nothing Then Return Nothing
        If Not TypeOf picked Is ComponentOccurrence Then Return Nothing
        
        Dim occ As ComponentOccurrence = CType(picked, ComponentOccurrence)
        
        ' Verify it's a Kask24 support
        If occ.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            MessageBox.Show("Please select a part, not an assembly.", "Modify Support")
            Return Nothing
        End If
        
        Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
        If Not SupportPlacementLib.IsKask24Support(partDoc) Then
            MessageBox.Show("Selected component is not a Kask24 support.", "Modify Support")
            Return Nothing
        End If
        
        Return occ
    Catch
        Return Nothing
    End Try
End Function

' ============================================================================
' Pick a parameter from the assembly
' ============================================================================
Function PickParameter(asmDoc As AssemblyDocument) As String
    ' Get list of user parameters
    Dim paramNames As String() = SupportPlacementLib.GetUserParameterNames(asmDoc)
    
    If paramNames Is Nothing OrElse paramNames.Length = 0 Then
        MessageBox.Show("No user parameters found in this assembly." & vbCrLf & _
                       "Create parameters first, then use them for length.", "Pick Parameter")
        Return ""
    End If
    
    ' Create a simple picker dialog
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Select Parameter"
    frm.Width = 300
    frm.Height = 350
    frm.StartPosition = FormStartPosition.CenterParent
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MaximizeBox = False
    frm.MinimizeBox = False
    
    Dim lbl As New System.Windows.Forms.Label()
    lbl.Text = "Select a parameter for length:"
    lbl.Left = 10
    lbl.Top = 10
    lbl.Width = 260
    frm.Controls.Add(lbl)
    
    Dim lst As New System.Windows.Forms.ListBox()
    lst.Left = 10
    lst.Top = 35
    lst.Width = 260
    lst.Height = 220
    For Each pn As String In paramNames
        lst.Items.Add(pn)
    Next
    If lst.Items.Count > 0 Then lst.SelectedIndex = 0
    frm.Controls.Add(lst)
    
    ' Double-click to select and close
    AddHandler lst.DoubleClick, Sub(s, e)
        If lst.SelectedItem IsNot Nothing Then
            frm.DialogResult = DialogResult.OK
        End If
    End Sub
    
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "OK"
    btnOK.Left = 110
    btnOK.Top = 270
    btnOK.Width = 75
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Cancel"
    btnCancel.Left = 195
    btnCancel.Top = 270
    btnCancel.Width = 75
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    
    frm.AcceptButton = btnOK
    frm.CancelButton = btnCancel
    
    If frm.ShowDialog() = DialogResult.OK AndAlso lst.SelectedItem IsNot Nothing Then
        Return lst.SelectedItem.ToString()
    End If
    
    Return ""
End Function

' ============================================================================
' Load settings from an existing occurrence
' ============================================================================
Sub LoadOccurrenceSettings(asmDoc As AssemblyDocument, occ As ComponentOccurrence, _
                            ByRef selectedWidth As Integer, ByRef placementMode As String, _
                            ByRef orientMode As String, ByRef alignPoint As String, _
                            ByRef lengthInput As String, ByRef flipDirection As Boolean, _
                            ByRef ref1 As Object, ByRef ref2 As Object, ByRef ref3 As Object, _
                            ByRef orientRef As Object, _
                            ByRef offsetX As String, ByRef offsetY As String, ByRef offsetZ As String, _
                            ByRef customName As String)
    
    ' Get stored references
    Dim refs As Dictionary(Of String, String) = SupportPlacementLib.GetOccurrenceReferences(occ)
    
    ' Load mode
    If refs.ContainsKey("Mode") AndAlso refs("Mode") <> "" Then
        placementMode = refs("Mode")
    End If
    
    ' Load align point
    If refs.ContainsKey("AlignPoint") AndAlso refs("AlignPoint") <> "" Then
        alignPoint = refs("AlignPoint")
    End If
    
    ' Load orient mode
    If refs.ContainsKey("OrientMode") Then
        orientMode = refs("OrientMode")
    End If
    
    ' Load length input (could be number in mm or parameter name)
    If refs.ContainsKey("LengthInput") AndAlso refs("LengthInput") <> "" Then
        lengthInput = refs("LengthInput")
    End If
    
    ' Load flip direction
    If refs.ContainsKey("FlipDirection") Then
        Boolean.TryParse(refs("FlipDirection"), flipDirection)
    End If
    
    ' Load offsets
    offsetX = ""
    offsetY = ""
    offsetZ = ""
    If refs.ContainsKey("OffsetX") Then offsetX = refs("OffsetX")
    If refs.ContainsKey("OffsetY") Then offsetY = refs("OffsetY")
    If refs.ContainsKey("OffsetZ") Then offsetZ = refs("OffsetZ")
    
    ' Load custom name
    customName = ""
    If refs.ContainsKey("CustomName") Then customName = refs("CustomName")
    
    ' Resolve work feature references
    ref1 = Nothing
    ref2 = Nothing
    ref3 = Nothing
    orientRef = Nothing
    
    If refs.ContainsKey("Ref1") AndAlso refs("Ref1") <> "" Then
        ref1 = SupportPlacementLib.ResolveWorkFeatureByName(asmDoc, refs("Ref1"))
    End If
    If refs.ContainsKey("Ref2") AndAlso refs("Ref2") <> "" Then
        ref2 = SupportPlacementLib.ResolveWorkFeatureByName(asmDoc, refs("Ref2"))
    End If
    If refs.ContainsKey("Ref3") AndAlso refs("Ref3") <> "" Then
        ref3 = SupportPlacementLib.ResolveWorkFeatureByName(asmDoc, refs("Ref3"))
    End If
    If refs.ContainsKey("OrientRef") AndAlso refs("OrientRef") <> "" Then
        orientRef = SupportPlacementLib.ResolveWorkFeatureByName(asmDoc, refs("OrientRef"))
    End If
    
    ' Get width from part parameters
    Try
        Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
        Dim widthVal As Double = partDoc.ComponentDefinition.Parameters.Item("Width").Value * 10 ' cm to mm
        selectedWidth = CInt(Math.Round(widthVal))
    Catch
    End Try
End Sub

' ============================================================================
' Update an existing support occurrence
' ============================================================================
Sub UpdateExistingSupport(app As Inventor.Application, asmDoc As AssemblyDocument, _
                           occ As ComponentOccurrence, _
                           widthMm As Integer, lengthCm As Double, lengthInput As String, _
                           startPoint As Point, direction As UnitVector, _
                           alignPoint As String, orientMode As String, orientRef As Object, _
                           placementMode As String, ref1 As Object, ref2 As Object, ref3 As Object, _
                           flipDirection As Boolean, _
                           offsetXStr As String, offsetYStr As String, offsetZStr As String, _
                           customName As String)
    
    Dim lengthMm As Integer = CInt(Math.Round(lengthCm * 10))
    Dim widthCm As Double = widthMm / 10.0
    
    ' Resolve offsets (mm or parameter name) to cm
    Dim offsetXCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetXStr)
    Dim offsetYCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetYStr)
    Dim offsetZCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetZStr)
    
    ' Update part parameters (width and length)
    Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
    Try
        Dim params As Parameters = partDoc.ComponentDefinition.Parameters
        
        ' Update width parameter
        Dim widthParam As Parameter = params.Item("Width")
        widthParam.Expression = widthCm.ToString() & " cm"
        
        ' Update length parameter
        Dim lengthParam As Parameter = params.Item("Length")
        lengthParam.Expression = lengthCm.ToString() & " cm"
        
        partDoc.Update()
    Catch ex As Exception
        MessageBox.Show("Could not update part parameters: " & ex.Message, "Update Support")
    End Try
    
    ' Calculate new placement matrix (uses updated width for align point offset)
    Dim placementMatrix As Matrix = SupportPlacementLib.CreateFullPlacementMatrix( _
        app, startPoint, direction, alignPoint, widthMm, orientMode, orientRef, _
        offsetXCm, offsetYCm, offsetZCm)
    
    ' Update occurrence transformation
    occ.Transformation = placementMatrix
    
    ' Store updated references (use GetWorkFeatureReference for path-based storage)
    Dim ref1Name As String = UtilsLib.GetWorkFeatureReference(ref1)
    Dim ref2Name As String = UtilsLib.GetWorkFeatureReference(ref2)
    Dim ref3Name As String = UtilsLib.GetWorkFeatureReference(ref3)
    Dim orientRefName As String = UtilsLib.GetWorkFeatureReference(orientRef)
    
    SupportPlacementLib.StoreOccurrenceReferences(occ, _
        placementMode, ref1Name, ref2Name, ref3Name, _
        alignPoint, orientMode, orientRefName, lengthInput, flipDirection, _
        offsetXStr, offsetYStr, offsetZStr, customName)
    
    ' Detect if user renamed occurrence and adopt that name
    SupportPlacementLib.DetectAndAdoptOccurrenceRename(asmDoc, partDoc)
    
    ' Update iProperties (respects custom Part Number if set)
    SupportPlacementLib.UpdateSupportiProperties(partDoc, widthMm, lengthMm, customName)
    partDoc.Save()
    
    ' Sync all occurrence names to match Part Number
    SupportPlacementLib.SyncOccurrenceNames(asmDoc, partDoc)
End Sub

' ============================================================================
' Place Support with All Settings - Returns the placed occurrence
' ============================================================================
Function PlaceSupportWithSettings(app As Inventor.Application, asmDoc As AssemblyDocument, _
                              templatePath As String, _
                              widthMm As Integer, lengthCm As Double, lengthInput As String, _
                              startPoint As Point, direction As UnitVector, _
                              alignPoint As String, orientMode As String, orientRef As Object, _
                              placementMode As String, ref1 As Object, ref2 As Object, ref3 As Object, _
                              flipDirection As Boolean, _
                              offsetXStr As String, offsetYStr As String, offsetZStr As String, _
                              customName As String) As ComponentOccurrence
    
    Dim lengthMm As Integer = CInt(Math.Round(lengthCm * 10))
    
    ' Resolve offsets (mm or parameter name) to cm
    Dim offsetXCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetXStr)
    Dim offsetYCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetYStr)
    Dim offsetZCm As Double = SupportPlacementLib.ResolveOffsetInput(asmDoc, offsetZStr)
    
    Dim partPath As String = ""
    
    If customName <> "" Then
        ' Custom name provided - always create a new file with that name
        partPath = SupportPlacementLib.GenerateSupportFileName(asmDoc, widthMm, customName)
        If Not SupportPlacementLib.CreateSupportPart(app, templatePath, partPath, widthMm, lengthCm) Then
            Throw New Exception("Failed to create support file.")
        End If
    Else
        ' No custom name - check for existing file with same dimensions
        partPath = SupportPlacementLib.FindExistingSupportFile(app, asmDoc, widthMm, lengthMm)
        
        If partPath = "" Then
            ' Create new file from template
            partPath = SupportPlacementLib.GenerateSupportFileName(asmDoc, widthMm)
            If Not SupportPlacementLib.CreateSupportPart(app, templatePath, partPath, widthMm, lengthCm) Then
                Throw New Exception("Failed to create support file.")
            End If
        End If
    End If
    
    ' Calculate full placement matrix including orientation and offsets
    Dim placementMatrix As Matrix = SupportPlacementLib.CreateFullPlacementMatrix( _
        app, startPoint, direction, alignPoint, widthMm, orientMode, orientRef, _
        offsetXCm, offsetYCm, offsetZCm)
    
    ' Place the part
    Dim occ As ComponentOccurrence = asmDoc.ComponentDefinition.Occurrences.Add(partPath, placementMatrix)
    
    If occ Is Nothing Then
        Throw New Exception("Failed to place component.")
    End If
    
    ' Store work feature references for parametric updates (path-based for component work features)
    Try
        Dim ref1Name As String = UtilsLib.GetWorkFeatureReference(ref1)
        Dim ref2Name As String = UtilsLib.GetWorkFeatureReference(ref2)
        Dim ref3Name As String = UtilsLib.GetWorkFeatureReference(ref3)
        Dim orientRefName As String = UtilsLib.GetWorkFeatureReference(orientRef)
        
        SupportPlacementLib.StoreOccurrenceReferences(occ, _
            placementMode, ref1Name, ref2Name, ref3Name, _
            alignPoint, orientMode, orientRefName, lengthInput, flipDirection, _
            offsetXStr, offsetYStr, offsetZStr, customName)
    Catch
    End Try
    
    ' Update iProperties on the part (for BOM)
    Try
        Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
        SupportPlacementLib.UpdateSupportiProperties(partDoc, widthMm, lengthMm, customName)
        partDoc.Save()
        
        ' Sync occurrence name to match Part Number
        SupportPlacementLib.SyncOccurrenceNames(asmDoc, partDoc)
    Catch
    End Try
    
    Return occ
End Function

' ============================================================================
' Pick Functions for Each Mode
' ============================================================================
Function PickForMode(app As Inventor.Application, mode As String, refNum As Integer) As Object
    Select Case mode
        Case "TWO_POINTS"
            If refNum = 1 Then
                Return UtilsLib.PickPoint(app, "Select start point (WorkPoint):")
            ElseIf refNum = 2 Then
                Return UtilsLib.PickPoint(app, "Select end point (WorkPoint):")
            End If
            
        Case "AXIS_TWO_PLANES"
            If refNum = 1 Then
                Return UtilsLib.PickAxis(app, "Select axis (WorkAxis):")
            ElseIf refNum = 2 Then
                Return UtilsLib.PickPlane(app, "Select start plane (WorkPlane):")
            ElseIf refNum = 3 Then
                Return UtilsLib.PickPlane(app, "Select end plane (WorkPlane):")
            End If
            
        Case "PLANE_AXIS_LENGTH"
            If refNum = 1 Then
                Return UtilsLib.PickPlane(app, "Select plane (WorkPlane):")
            ElseIf refNum = 2 Then
                Return UtilsLib.PickAxis(app, "Select axis (WorkAxis):")
            End If
            
        Case "POINT_AXIS_LENGTH"
            If refNum = 1 Then
                Return UtilsLib.PickPoint(app, "Select point (WorkPoint):")
            ElseIf refNum = 2 Then
                Return UtilsLib.PickAxis(app, "Select axis (WorkAxis):")
            End If
            
        Case "TWO_PLANES_POINT"
            If refNum = 1 Then
                Return UtilsLib.PickPlane(app, "Select first plane (WorkPlane):")
            ElseIf refNum = 2 Then
                Return UtilsLib.PickPlane(app, "Select second plane (WorkPlane):")
            ElseIf refNum = 3 Then
                Return UtilsLib.PickPoint(app, "Select position point (WorkPoint):")
            End If
    End Select
    Return Nothing
End Function

Function PickForOrient(app As Inventor.Application, orientMode As String) As Object
    Select Case orientMode
        Case "WIDTH_AXIS", "THICKNESS_AXIS"
            Return UtilsLib.PickAxis(app, "Select axis for orientation (WorkAxis):")
        Case "ALIGN_BOTTOM", "ALIGN_SIDE"
            Return UtilsLib.PickPlane(app, "Select plane for alignment (WorkPlane):")
    End Select
    Return Nothing
End Function

' ============================================================================
' Validation
' ============================================================================
Function ValidatePlacement(mode As String, ref1 As Object, ref2 As Object, ref3 As Object, _
                           manualLen As Double) As Boolean
    
    Select Case mode
        Case "TWO_POINTS"
            Return ref1 IsNot Nothing AndAlso ref2 IsNot Nothing
            
        Case "AXIS_TWO_PLANES"
            Return ref1 IsNot Nothing AndAlso ref2 IsNot Nothing AndAlso ref3 IsNot Nothing
            
        Case "PLANE_AXIS_LENGTH"
            Return ref1 IsNot Nothing AndAlso ref2 IsNot Nothing AndAlso manualLen > 0
            
        Case "POINT_AXIS_LENGTH"
            Return ref1 IsNot Nothing AndAlso ref2 IsNot Nothing AndAlso manualLen > 0
            
        Case "TWO_PLANES_POINT"
            Return ref1 IsNot Nothing AndAlso ref2 IsNot Nothing AndAlso ref3 IsNot Nothing
            
        Case Else
            Return False
    End Select
End Function

' ============================================================================
' Helper Functions
' ============================================================================
Function NeedsRef3(mode As String) As Boolean
    Return mode = "AXIS_TWO_PLANES" OrElse mode = "TWO_PLANES_POINT"
End Function

Function NeedsManualLength(mode As String) As Boolean
    Return mode = "PLANE_AXIS_LENGTH" OrElse mode = "POINT_AXIS_LENGTH"
End Function

Function GetRef1Label(mode As String) As String
    Select Case mode
        Case "TWO_POINTS" : Return "Start Point:"
        Case "AXIS_TWO_PLANES" : Return "Axis:"
        Case "PLANE_AXIS_LENGTH" : Return "Plane:"
        Case "POINT_AXIS_LENGTH" : Return "Point:"
        Case "TWO_PLANES_POINT" : Return "Plane 1:"
        Case Else : Return "Reference 1:"
    End Select
End Function

Function GetRef2Label(mode As String) As String
    Select Case mode
        Case "TWO_POINTS" : Return "End Point:"
        Case "AXIS_TWO_PLANES" : Return "Start Plane:"
        Case "PLANE_AXIS_LENGTH" : Return "Axis:"
        Case "POINT_AXIS_LENGTH" : Return "Axis:"
        Case "TWO_PLANES_POINT" : Return "Plane 2:"
        Case Else : Return "Reference 2:"
    End Select
End Function

Function GetRef3Label(mode As String) As String
    Select Case mode
        Case "AXIS_TWO_PLANES" : Return "End Plane:"
        Case "TWO_PLANES_POINT" : Return "Position Point:"
        Case Else : Return "Reference 3:"
    End Select
End Function

' ============================================================================
' UI Form (Modeless)
' ============================================================================
Function ShowPlacementForm(app As Inventor.Application, asmDoc As AssemblyDocument, _
                           templateWorkPoints As String(), _
                           ByRef selectedWidth As Integer, ByRef placementMode As String, _
                           ByRef orientMode As String, ByRef alignPoint As String, _
                           ByRef lengthInput As String, ByRef flipDirection As Boolean, _
                           ref1 As Object, ref2 As Object, ref3 As Object, orientRef As Object, _
                           calculatedLength As Double, modifyOcc As ComponentOccurrence, _
                           ByRef action As String, _
                           ByRef formLeft As Integer, ByRef formTop As Integer, _
                           ByRef formWidth As Integer, ByRef formHeight As Integer, _
                           ByRef offsetX As String, ByRef offsetY As String, ByRef offsetZ As String, _
                           ByRef customName As String) As DialogResult
    
    action = ""
    Dim uom As UnitsOfMeasure = asmDoc.UnitsOfMeasure
    Dim isModifyMode As Boolean = (modifyOcc IsNot Nothing)
    
    ' Create form
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Place Birch Support"
    frm.Width = formWidth
    frm.Height = formHeight
    frm.FormBorderStyle = FormBorderStyle.Sizable
    frm.MaximizeBox = False
    frm.MinimizeBox = True
    frm.TopMost = False
    
    ' Set position (use saved position if available, otherwise center on screen)
    If formLeft >= 0 AndAlso formTop >= 0 Then
        frm.StartPosition = FormStartPosition.Manual
        frm.Left = formLeft
        frm.Top = formTop
    Else
        frm.StartPosition = FormStartPosition.CenterScreen
    End If
    
    Dim yPos As Integer = 20
    Dim leftCol As Integer = 20
    Dim rightCol As Integer = 130
    Dim labelWidth As Integer = 100
    Dim controlWidth As Integer = 200
    Dim rowHeight As Integer = 32
    
    ' === WIDTH SELECTION ===
    Dim lblWidth As New System.Windows.Forms.Label()
    lblWidth.Text = "Beam Width:"
    lblWidth.Left = leftCol
    lblWidth.Top = yPos
    lblWidth.Width = labelWidth
    frm.Controls.Add(lblWidth)
    
    Dim cboWidth As New System.Windows.Forms.ComboBox()
    cboWidth.Left = rightCol
    cboWidth.Top = yPos - 2
    cboWidth.Width = 100
    cboWidth.DropDownStyle = ComboBoxStyle.DropDownList
    For Each w As Integer In SupportPlacementLib.SUPPORT_WIDTHS
        cboWidth.Items.Add(w.ToString() & " mm")
    Next
    Dim widthIdx As Integer = Array.IndexOf(SupportPlacementLib.SUPPORT_WIDTHS, selectedWidth)
    cboWidth.SelectedIndex = If(widthIdx >= 0, widthIdx, 0)
    frm.Controls.Add(cboWidth)
    
    yPos += rowHeight
    
    ' === CUSTOM NAME ===
    Dim lblName As New System.Windows.Forms.Label()
    lblName.Text = "Name (optional):"
    lblName.Left = leftCol
    lblName.Top = yPos
    lblName.Width = labelWidth
    frm.Controls.Add(lblName)
    
    Dim txtName As New System.Windows.Forms.TextBox()
    txtName.Name = "txtName"
    txtName.Left = rightCol
    txtName.Top = yPos - 2
    txtName.Width = controlWidth
    txtName.Text = customName
    frm.Controls.Add(txtName)
    
    yPos += rowHeight + 10
    
    ' === PLACEMENT MODE ===
    Dim lblMode As New System.Windows.Forms.Label()
    lblMode.Text = "Placement Mode:"
    lblMode.Left = leftCol
    lblMode.Top = yPos
    lblMode.Width = labelWidth
    frm.Controls.Add(lblMode)
    
    Dim cboMode As New System.Windows.Forms.ComboBox()
    cboMode.Left = rightCol
    cboMode.Top = yPos - 2
    cboMode.Width = controlWidth
    cboMode.DropDownStyle = ComboBoxStyle.DropDownList
    cboMode.Items.Add("Two Points")
    cboMode.Items.Add("Axis + Two Planes")
    cboMode.Items.Add("Plane + Axis + Length")
    cboMode.Items.Add("Point + Axis + Length")
    cboMode.Items.Add("Two Planes + Point")
    
    Select Case placementMode
        Case "TWO_POINTS" : cboMode.SelectedIndex = 0
        Case "AXIS_TWO_PLANES" : cboMode.SelectedIndex = 1
        Case "PLANE_AXIS_LENGTH" : cboMode.SelectedIndex = 2
        Case "POINT_AXIS_LENGTH" : cboMode.SelectedIndex = 3
        Case "TWO_PLANES_POINT" : cboMode.SelectedIndex = 4
        Case Else : cboMode.SelectedIndex = 1
    End Select
    frm.Controls.Add(cboMode)
    
    yPos += rowHeight + 10
    
    ' === GEOMETRY REFERENCES ===
    Dim lblRefSection As New System.Windows.Forms.Label()
    lblRefSection.Text = "--- Geometry (Work Features Only) ---"
    lblRefSection.Left = leftCol
    lblRefSection.Top = yPos
    lblRefSection.Width = 350
    frm.Controls.Add(lblRefSection)
    
    yPos += 25
    
    ' Ref1
    Dim lblRef1 As New System.Windows.Forms.Label()
    lblRef1.Text = GetRef1Label(placementMode)
    lblRef1.Left = leftCol
    lblRef1.Top = yPos
    lblRef1.Width = labelWidth
    frm.Controls.Add(lblRef1)
    
    Dim txtRef1 As New System.Windows.Forms.TextBox()
    txtRef1.Left = rightCol
    txtRef1.Top = yPos - 2
    txtRef1.Width = controlWidth
    txtRef1.ReadOnly = True
    txtRef1.Text = UtilsLib.GetObjectDisplayName(ref1)
    frm.Controls.Add(txtRef1)
    
    Dim btnRef1 As New System.Windows.Forms.Button()
    btnRef1.Text = "Pick"
    btnRef1.Left = rightCol + controlWidth + 10
    btnRef1.Top = yPos - 3
    btnRef1.Width = 50
    frm.Controls.Add(btnRef1)
    
    Dim btnClearRef1 As New System.Windows.Forms.Button()
    btnClearRef1.Text = "X"
    btnClearRef1.Left = rightCol + controlWidth + 65
    btnClearRef1.Top = yPos - 3
    btnClearRef1.Width = 25
    frm.Controls.Add(btnClearRef1)
    
    yPos += rowHeight
    
    ' Ref2
    Dim lblRef2 As New System.Windows.Forms.Label()
    lblRef2.Text = GetRef2Label(placementMode)
    lblRef2.Left = leftCol
    lblRef2.Top = yPos
    lblRef2.Width = labelWidth
    frm.Controls.Add(lblRef2)
    
    Dim txtRef2 As New System.Windows.Forms.TextBox()
    txtRef2.Left = rightCol
    txtRef2.Top = yPos - 2
    txtRef2.Width = controlWidth
    txtRef2.ReadOnly = True
    txtRef2.Text = UtilsLib.GetObjectDisplayName(ref2)
    frm.Controls.Add(txtRef2)
    
    Dim btnRef2 As New System.Windows.Forms.Button()
    btnRef2.Text = "Pick"
    btnRef2.Left = rightCol + controlWidth + 10
    btnRef2.Top = yPos - 3
    btnRef2.Width = 50
    frm.Controls.Add(btnRef2)
    
    Dim btnClearRef2 As New System.Windows.Forms.Button()
    btnClearRef2.Text = "X"
    btnClearRef2.Left = rightCol + controlWidth + 65
    btnClearRef2.Top = yPos - 3
    btnClearRef2.Width = 25
    frm.Controls.Add(btnClearRef2)
    
    yPos += rowHeight
    
    ' Ref3 (conditional)
    Dim lblRef3 As New System.Windows.Forms.Label()
    lblRef3.Text = GetRef3Label(placementMode)
    lblRef3.Left = leftCol
    lblRef3.Top = yPos
    lblRef3.Width = labelWidth
    lblRef3.Visible = NeedsRef3(placementMode)
    frm.Controls.Add(lblRef3)
    
    Dim txtRef3 As New System.Windows.Forms.TextBox()
    txtRef3.Left = rightCol
    txtRef3.Top = yPos - 2
    txtRef3.Width = controlWidth
    txtRef3.ReadOnly = True
    txtRef3.Text = UtilsLib.GetObjectDisplayName(ref3)
    txtRef3.Visible = NeedsRef3(placementMode)
    frm.Controls.Add(txtRef3)
    
    Dim btnRef3 As New System.Windows.Forms.Button()
    btnRef3.Text = "Pick"
    btnRef3.Left = rightCol + controlWidth + 10
    btnRef3.Top = yPos - 3
    btnRef3.Width = 50
    btnRef3.Visible = NeedsRef3(placementMode)
    frm.Controls.Add(btnRef3)
    
    Dim btnClearRef3 As New System.Windows.Forms.Button()
    btnClearRef3.Text = "X"
    btnClearRef3.Left = rightCol + controlWidth + 65
    btnClearRef3.Top = yPos - 3
    btnClearRef3.Width = 25
    btnClearRef3.Visible = NeedsRef3(placementMode)
    frm.Controls.Add(btnClearRef3)
    
    yPos += rowHeight
    
    ' Manual Length (conditional)
    Dim lblManualLen As New System.Windows.Forms.Label()
    lblManualLen.Text = "Length:"
    lblManualLen.Left = leftCol
    lblManualLen.Top = yPos
    lblManualLen.Width = labelWidth
    lblManualLen.Visible = NeedsManualLength(placementMode)
    frm.Controls.Add(lblManualLen)
    
    Dim txtManualLen As New System.Windows.Forms.TextBox()
    txtManualLen.Left = rightCol
    txtManualLen.Top = yPos - 2
    txtManualLen.Width = 120
    txtManualLen.Text = lengthInput  ' Can be number (mm) or parameter name
    txtManualLen.Visible = NeedsManualLength(placementMode)
    frm.Controls.Add(txtManualLen)
    
    Dim lblMm As New System.Windows.Forms.Label()
    lblMm.Text = "mm/param"
    lblMm.Left = rightCol + 85
    lblMm.Top = yPos
    lblMm.Width = 55
    lblMm.Visible = NeedsManualLength(placementMode)
    frm.Controls.Add(lblMm)
    
    Dim btnPickParam As New System.Windows.Forms.Button()
    btnPickParam.Text = "..."
    btnPickParam.Left = rightCol + 145
    btnPickParam.Top = yPos - 3
    btnPickParam.Width = 30
    btnPickParam.Visible = NeedsManualLength(placementMode)
    frm.Controls.Add(btnPickParam)
    
    yPos += rowHeight
    
    ' Flip Direction
    Dim chkFlip As New System.Windows.Forms.CheckBox()
    chkFlip.Text = "Flip Direction"
    chkFlip.Left = rightCol
    chkFlip.Top = yPos
    chkFlip.Width = 150
    chkFlip.Checked = flipDirection
    frm.Controls.Add(chkFlip)
    
    yPos += rowHeight + 10
    
    ' === ORIENTATION ===
    Dim lblOrientSection As New System.Windows.Forms.Label()
    lblOrientSection.Text = "--- Orientation ---"
    lblOrientSection.Left = leftCol
    lblOrientSection.Top = yPos
    lblOrientSection.Width = 350
    frm.Controls.Add(lblOrientSection)
    
    yPos += 25
    
    Dim lblOrientMode As New System.Windows.Forms.Label()
    lblOrientMode.Text = "Mode:"
    lblOrientMode.Left = leftCol
    lblOrientMode.Top = yPos
    lblOrientMode.Width = labelWidth
    frm.Controls.Add(lblOrientMode)
    
    Dim cboOrient As New System.Windows.Forms.ComboBox()
    cboOrient.Left = rightCol
    cboOrient.Top = yPos - 2
    cboOrient.Width = controlWidth
    cboOrient.DropDownStyle = ComboBoxStyle.DropDownList
    cboOrient.Items.Add("None")
    cboOrient.Items.Add("Align Bottom to Plane")
    cboOrient.Items.Add("Align Side to Plane")
    cboOrient.Items.Add("Align Width to Axis")
    
    Select Case orientMode
        Case "NONE" : cboOrient.SelectedIndex = 0
        Case "ALIGN_BOTTOM" : cboOrient.SelectedIndex = 1
        Case "ALIGN_SIDE" : cboOrient.SelectedIndex = 2
        Case "WIDTH_AXIS" : cboOrient.SelectedIndex = 3
        Case Else : cboOrient.SelectedIndex = 1
    End Select
    frm.Controls.Add(cboOrient)
    
    yPos += rowHeight
    
    Dim lblOrientRef As New System.Windows.Forms.Label()
    lblOrientRef.Text = "Reference:"
    lblOrientRef.Left = leftCol
    lblOrientRef.Top = yPos
    lblOrientRef.Width = labelWidth
    frm.Controls.Add(lblOrientRef)
    
    Dim txtOrientRef As New System.Windows.Forms.TextBox()
    txtOrientRef.Left = rightCol
    txtOrientRef.Top = yPos - 2
    txtOrientRef.Width = controlWidth
    txtOrientRef.ReadOnly = True
    txtOrientRef.Text = UtilsLib.GetObjectDisplayName(orientRef)
    frm.Controls.Add(txtOrientRef)
    
    Dim btnOrientRef As New System.Windows.Forms.Button()
    btnOrientRef.Text = "Pick"
    btnOrientRef.Left = rightCol + controlWidth + 10
    btnOrientRef.Top = yPos - 3
    btnOrientRef.Width = 50
    frm.Controls.Add(btnOrientRef)
    
    Dim btnClearOrient As New System.Windows.Forms.Button()
    btnClearOrient.Text = "X"
    btnClearOrient.Left = rightCol + controlWidth + 65
    btnClearOrient.Top = yPos - 3
    btnClearOrient.Width = 25
    frm.Controls.Add(btnClearOrient)
    
    yPos += rowHeight + 10
    
    ' === ALIGN POINT ===
    Dim lblAlignSection As New System.Windows.Forms.Label()
    lblAlignSection.Text = "--- Align Point ---"
    lblAlignSection.Left = leftCol
    lblAlignSection.Top = yPos
    lblAlignSection.Width = 350
    frm.Controls.Add(lblAlignSection)
    
    yPos += 25
    
    Dim lblAlignPt As New System.Windows.Forms.Label()
    lblAlignPt.Text = "Align:"
    lblAlignPt.Left = leftCol
    lblAlignPt.Top = yPos
    lblAlignPt.Width = labelWidth
    frm.Controls.Add(lblAlignPt)
    
    Dim cboAlignPt As New System.Windows.Forms.ComboBox()
    cboAlignPt.Left = rightCol
    cboAlignPt.Top = yPos - 2
    cboAlignPt.Width = controlWidth
    cboAlignPt.DropDownStyle = ComboBoxStyle.DropDownList
    For Each wpName As String In templateWorkPoints
        cboAlignPt.Items.Add(wpName)
    Next
    Dim alignIdx As Integer = Array.IndexOf(templateWorkPoints, alignPoint)
    cboAlignPt.SelectedIndex = If(alignIdx >= 0, alignIdx, 0)
    frm.Controls.Add(cboAlignPt)
    
    yPos += rowHeight
    
    ' Offset (same row, compact layout)
    Dim lblOffsetX As New System.Windows.Forms.Label()
    lblOffsetX.Text = "Offset X:"
    lblOffsetX.Left = leftCol
    lblOffsetX.Top = yPos
    lblOffsetX.Width = 55
    frm.Controls.Add(lblOffsetX)
    
    Dim txtOffsetX As New System.Windows.Forms.TextBox()
    txtOffsetX.Left = leftCol + 55
    txtOffsetX.Top = yPos - 2
    txtOffsetX.Width = 50
    txtOffsetX.Text = offsetX
    frm.Controls.Add(txtOffsetX)
    
    Dim lblOffsetY As New System.Windows.Forms.Label()
    lblOffsetY.Text = "Y:"
    lblOffsetY.Left = leftCol + 115
    lblOffsetY.Top = yPos
    lblOffsetY.Width = 20
    frm.Controls.Add(lblOffsetY)
    
    Dim txtOffsetY As New System.Windows.Forms.TextBox()
    txtOffsetY.Left = leftCol + 135
    txtOffsetY.Top = yPos - 2
    txtOffsetY.Width = 50
    txtOffsetY.Text = offsetY
    frm.Controls.Add(txtOffsetY)
    
    Dim lblOffsetZ As New System.Windows.Forms.Label()
    lblOffsetZ.Text = "Z:"
    lblOffsetZ.Left = leftCol + 195
    lblOffsetZ.Top = yPos
    lblOffsetZ.Width = 20
    frm.Controls.Add(lblOffsetZ)
    
    Dim txtOffsetZ As New System.Windows.Forms.TextBox()
    txtOffsetZ.Left = leftCol + 215
    txtOffsetZ.Top = yPos - 2
    txtOffsetZ.Width = 50
    txtOffsetZ.Text = offsetZ
    frm.Controls.Add(txtOffsetZ)
    
    Dim lblOffsetUnit As New System.Windows.Forms.Label()
    lblOffsetUnit.Text = "mm/param"
    lblOffsetUnit.Left = leftCol + 270
    lblOffsetUnit.Top = yPos
    lblOffsetUnit.Width = 60
    frm.Controls.Add(lblOffsetUnit)
    
    yPos += rowHeight + 10
    
    ' === CALCULATED LENGTH ===
    Dim lblCalcLen As New System.Windows.Forms.Label()
    lblCalcLen.Text = "Calculated Length: " & If(calculatedLength > 0, (calculatedLength * 10).ToString("F1") & " mm", "(not yet calculated)")
    lblCalcLen.Left = leftCol
    lblCalcLen.Top = yPos
    lblCalcLen.Width = 350
    frm.Controls.Add(lblCalcLen)
    
    yPos += rowHeight + 10
    
    ' === ACTION BUTTONS (all on one row) ===
    Dim btnPlace As New System.Windows.Forms.Button()
    btnPlace.Text = "Place"
    btnPlace.Left = leftCol
    btnPlace.Top = yPos
    btnPlace.Width = 60
    btnPlace.Height = 28
    frm.Controls.Add(btnPlace)
    
    Dim btnPlaceClose As New System.Windows.Forms.Button()
    btnPlaceClose.Text = "Place+Close"
    btnPlaceClose.Left = leftCol + 65
    btnPlaceClose.Top = yPos
    btnPlaceClose.Width = 85
    btnPlaceClose.Height = 28
    frm.Controls.Add(btnPlaceClose)
    
    Dim btnUpdate As New System.Windows.Forms.Button()
    btnUpdate.Text = "Update"
    btnUpdate.Left = leftCol + 160
    btnUpdate.Top = yPos
    btnUpdate.Width = 65
    btnUpdate.Height = 28
    btnUpdate.Enabled = isModifyMode
    frm.Controls.Add(btnUpdate)
    
    Dim btnUpdateClose As New System.Windows.Forms.Button()
    btnUpdateClose.Text = "Upd+Close"
    btnUpdateClose.Left = leftCol + 230
    btnUpdateClose.Top = yPos
    btnUpdateClose.Width = 80
    btnUpdateClose.Height = 28
    btnUpdateClose.Enabled = isModifyMode
    frm.Controls.Add(btnUpdateClose)
    
    yPos += 35
    
    ' === MODIFY SELECTION ROW ===
    Dim lblModifyLbl As New System.Windows.Forms.Label()
    lblModifyLbl.Text = "Modify:"
    lblModifyLbl.Left = leftCol
    lblModifyLbl.Top = yPos + 5
    lblModifyLbl.Width = 50
    frm.Controls.Add(lblModifyLbl)
    
    Dim lblModifyStatus As New System.Windows.Forms.Label()
    If isModifyMode Then
        lblModifyStatus.Text = modifyOcc.Name
    Else
        lblModifyStatus.Text = "(none)"
    End If
    lblModifyStatus.Left = leftCol + 55
    lblModifyStatus.Top = yPos + 5
    lblModifyStatus.Width = 150
    frm.Controls.Add(lblModifyStatus)
    
    Dim btnModify As New System.Windows.Forms.Button()
    btnModify.Text = "Select..."
    btnModify.Left = leftCol + 210
    btnModify.Top = yPos
    btnModify.Width = 60
    btnModify.Height = 25
    frm.Controls.Add(btnModify)
    
    Dim btnClearModify As New System.Windows.Forms.Button()
    btnClearModify.Text = "New"
    btnClearModify.Left = leftCol + 275
    btnClearModify.Top = yPos
    btnClearModify.Width = 45
    btnClearModify.Height = 25
    btnClearModify.Enabled = isModifyMode
    frm.Controls.Add(btnClearModify)
    
    yPos += 35
    
    ' === BOTTOM BUTTONS ===
    Dim btnClear As New System.Windows.Forms.Button()
    btnClear.Text = "Clear All"
    btnClear.Left = leftCol
    btnClear.Top = yPos
    btnClear.Width = 80
    btnClear.Height = 28
    frm.Controls.Add(btnClear)
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Close"
    btnCancel.Left = leftCol + 250
    btnCancel.Top = yPos
    btnCancel.Width = 70
    btnCancel.Height = 28
    frm.Controls.Add(btnCancel)
    
    ' === Store action in form Tag to avoid ByRef in lambdas ===
    frm.Tag = ""
    
    ' === EVENT HANDLERS (using Tag to avoid ByRef issues) ===
    AddHandler cboMode.SelectedIndexChanged, Sub(s, e)
        frm.Tag = "MODE_CHANGED"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnRef1.Click, Sub(s, e)
        frm.Tag = "PICK_REF1"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnRef2.Click, Sub(s, e)
        frm.Tag = "PICK_REF2"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnRef3.Click, Sub(s, e)
        frm.Tag = "PICK_REF3"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnOrientRef.Click, Sub(s, e)
        frm.Tag = "PICK_ORIENT"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnClearRef1.Click, Sub(s, e)
        frm.Tag = "CLEAR_REF1"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnClearRef2.Click, Sub(s, e)
        frm.Tag = "CLEAR_REF2"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnClearRef3.Click, Sub(s, e)
        frm.Tag = "CLEAR_REF3"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnClearOrient.Click, Sub(s, e)
        frm.Tag = "CLEAR_ORIENT"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnPickParam.Click, Sub(s, e)
        frm.Tag = "PICK_PARAM"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnPlace.Click, Sub(s, e)
        frm.Tag = "PLACE"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnPlaceClose.Click, Sub(s, e)
        frm.Tag = "PLACE_CLOSE"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnUpdate.Click, Sub(s, e)
        frm.Tag = "UPDATE"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnUpdateClose.Click, Sub(s, e)
        frm.Tag = "UPDATE_CLOSE"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnModify.Click, Sub(s, e)
        frm.Tag = "MODIFY"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnClearModify.Click, Sub(s, e)
        frm.Tag = "NEW"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnClear.Click, Sub(s, e)
        frm.Tag = "CLEAR"
        frm.DialogResult = DialogResult.OK
    End Sub
    
    AddHandler btnCancel.Click, Sub(s, e)
        frm.DialogResult = DialogResult.Cancel
    End Sub
    
    ' Handle X button (form close) same as Cancel
    AddHandler frm.FormClosing, Sub(s, e)
        If frm.DialogResult = DialogResult.None Then
            frm.DialogResult = DialogResult.Cancel
        End If
    End Sub
    
    ' Show dialog (modeless with DoEvents loop for Inventor interaction)
    frm.Show()
    
    ' Keep form responsive while allowing Inventor viewport interaction
    Do While frm.Visible AndAlso frm.DialogResult = DialogResult.None
        System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(20)
    Loop
    
    Dim result As DialogResult = frm.DialogResult
    
    ' Save form position/size for next time
    formLeft = frm.Left
    formTop = frm.Top
    formWidth = frm.Width
    formHeight = frm.Height
    
    ' Read values from controls BEFORE closing (controls may be disposed after close)
    If result = DialogResult.OK OrElse result = DialogResult.Cancel Then
        ' Read width
        If cboWidth.SelectedIndex >= 0 Then
            selectedWidth = SupportPlacementLib.SUPPORT_WIDTHS(cboWidth.SelectedIndex)
        End If
        
        ' Read placement mode
        Select Case cboMode.SelectedIndex
            Case 0 : placementMode = "TWO_POINTS"
            Case 1 : placementMode = "AXIS_TWO_PLANES"
            Case 2 : placementMode = "PLANE_AXIS_LENGTH"
            Case 3 : placementMode = "POINT_AXIS_LENGTH"
            Case 4 : placementMode = "TWO_PLANES_POINT"
        End Select
        
        ' Read orientation mode
        Select Case cboOrient.SelectedIndex
            Case 0 : orientMode = "NONE"
            Case 1 : orientMode = "ALIGN_BOTTOM"
            Case 2 : orientMode = "ALIGN_SIDE"
            Case 3 : orientMode = "WIDTH_AXIS"
        End Select
        
        ' Read align point
        If cboAlignPt.SelectedIndex >= 0 AndAlso cboAlignPt.SelectedIndex < templateWorkPoints.Length Then
            alignPoint = templateWorkPoints(cboAlignPt.SelectedIndex)
        End If
        
        ' Read length input (could be number in mm or parameter name)
        lengthInput = txtManualLen.Text.Trim()
        
        ' Read flip direction
        flipDirection = chkFlip.Checked
        
        ' Read offsets
        offsetX = txtOffsetX.Text.Trim()
        offsetY = txtOffsetY.Text.Trim()
        offsetZ = txtOffsetZ.Text.Trim()
        
        ' Read custom name
        customName = txtName.Text.Trim()
        
        ' Read action from Tag
        action = CStr(frm.Tag)
    End If
    
    ' Close the form after reading values
    If frm.Visible Then frm.Close()
    
    Return result
End Function

' ============================================================================
' Form Position Persistence (stored in assembly iProperties)
' ============================================================================
Sub LoadFormPosition(asmDoc As AssemblyDocument, ByRef left As Integer, ByRef top As Integer, _
                     ByRef width As Integer, ByRef height As Integer)
    Try
        Dim customProps As PropertySet = asmDoc.PropertySets.Item("Inventor User Defined Properties")
        
        Try : left = CInt(customProps.Item("PlaceSupportFormLeft").Value) : Catch : End Try
        Try : top = CInt(customProps.Item("PlaceSupportFormTop").Value) : Catch : End Try
        Try : width = CInt(customProps.Item("PlaceSupportFormWidth").Value) : Catch : End Try
        Try : height = CInt(customProps.Item("PlaceSupportFormHeight").Value) : Catch : End Try
        
        ' Validate - ensure form is on screen
        If left < 0 OrElse left > 3000 Then left = -1
        If top < 0 OrElse top > 2000 Then top = -1
        If width < 400 Then width = 520
        If height < 400 Then height = 720
    Catch
    End Try
End Sub

Sub SaveFormPosition(asmDoc As AssemblyDocument, left As Integer, top As Integer, _
                     width As Integer, height As Integer)
    Try
        Dim customProps As PropertySet = asmDoc.PropertySets.Item("Inventor User Defined Properties")
        
        SetFormProp(customProps, "PlaceSupportFormLeft", left)
        SetFormProp(customProps, "PlaceSupportFormTop", top)
        SetFormProp(customProps, "PlaceSupportFormWidth", width)
        SetFormProp(customProps, "PlaceSupportFormHeight", height)
    Catch
    End Try
End Sub

Sub SetFormProp(propSet As PropertySet, name As String, value As Integer)
    Try
        propSet.Item(name).Value = value
    Catch
        Try
            propSet.Add(value, name)
        Catch
        End Try
    End Try
End Sub
