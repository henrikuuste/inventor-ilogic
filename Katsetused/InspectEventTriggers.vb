' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' InspectEventTriggers - Inspect existing iLogic Event Triggers
'
' This diagnostic script reads and displays all event triggers configured
' in the current document. Use this to:
' 1. Verify triggers were set correctly by TestEventTriggerSetup
' 2. Discover PropId values by manually setting triggers in UI first
'
' To discover User Parameter Change PropId:
' 1. Open a part/assembly
' 2. Manually add a rule to "Any User Parameter Change" via Event Triggers dialog
' 3. Run this script to see what PropId Inventor used
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        Logger.Error("InspectEventTriggers: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "InspectEventTriggers")
        Exit Sub
    End If
    
    Logger.Info("InspectEventTriggers: === Event Trigger Inspection ===")
    Logger.Info("InspectEventTriggers: Document: " & doc.DisplayName)
    Logger.Info("InspectEventTriggers: Full Path: " & doc.FullFileName)
    Logger.Info("InspectEventTriggers: Type: " & doc.DocumentType.ToString())
    
    ' List all PropertySets
    Logger.Info("InspectEventTriggers: --- All PropertySets in Document ---")
    For Each ps As PropertySet In doc.PropertySets
        Logger.Info("InspectEventTriggers:   Name: '" & ps.Name & "', InternalName: " & ps.InternalName & ", Count: " & ps.Count)
    Next
    
    ' Look for event triggers PropertySet
    Logger.Info("InspectEventTriggers: --- Event Triggers PropertySet ---")
    Dim eventTriggersPropSet As PropertySet = Nothing
    
    ' Try various known names
    Dim possibleNames() As String = {"iLogicEventsRules", "_iLogicEventsRules", "_!iLogicEventsRules"}
    
    For Each psName As String In possibleNames
        Try
            eventTriggersPropSet = doc.PropertySets.Item(psName)
            Logger.Info("InspectEventTriggers: Found PropertySet by name: '" & psName & "'")
            Exit For
        Catch
        End Try
    Next
    
    ' Try by internal name GUID
    If eventTriggersPropSet Is Nothing Then
        Dim expectedGuid As String = "{2C540830-0723-455E-A8E2-891722EB4C3E}"
        For Each ps As PropertySet In doc.PropertySets
            If ps.InternalName = expectedGuid Then
                eventTriggersPropSet = ps
                Logger.Info("InspectEventTriggers: Found PropertySet by GUID: " & expectedGuid)
                Exit For
            End If
        Next
    End If
    
    If eventTriggersPropSet Is Nothing Then
        Logger.Info("InspectEventTriggers: No event triggers PropertySet found.")
        Logger.Info("InspectEventTriggers: This document has no event triggers configured.")
        Logger.Info("InspectEventTriggers: To discover PropIds:")
        Logger.Info("InspectEventTriggers: 1. Use Manage > Event Triggers in Inventor UI")
        Logger.Info("InspectEventTriggers: 2. Add a rule to an event (e.g., 'Any User Parameter Change')")
        Logger.Info("InspectEventTriggers: 3. Save the document")
        Logger.Info("InspectEventTriggers: 4. Run this script again")
    Else
        Logger.Info("InspectEventTriggers: PropertySet Details:")
        Logger.Info("InspectEventTriggers:   Name: " & eventTriggersPropSet.Name)
        Logger.Info("InspectEventTriggers:   InternalName: " & eventTriggersPropSet.InternalName)
        Logger.Info("InspectEventTriggers:   DisplayName: " & eventTriggersPropSet.DisplayName)
        Logger.Info("InspectEventTriggers:   Count: " & eventTriggersPropSet.Count)
        
        ' List all event trigger properties
        Logger.Info("InspectEventTriggers: --- All Event Trigger Properties ---")
        
        For Each prop As [Property] In eventTriggersPropSet
            Try
                Dim propInfo As String = String.Format("PropId={0}, Name='{1}', Value='{2}'", _
                    prop.PropId, prop.Name, prop.Value.ToString())
                
                ' Add event type hint
                Dim eventType As String = GetEventTypeFromPropId(prop.PropId)
                If eventType <> "" Then
                    propInfo &= " [" & eventType & "]"
                End If
                
                Logger.Info("InspectEventTriggers:   " & propInfo)
            Catch ex As Exception
                Logger.Warn("InspectEventTriggers:   PropId=" & prop.PropId & " - Error: " & ex.Message)
            End Try
        Next
        
        Logger.Info("InspectEventTriggers: --- PropId Range Reference ---")
        Logger.Info("InspectEventTriggers:   400-499:  After Open Document (AfterDocOpen)")
        Logger.Info("InspectEventTriggers:   500-599:  Close Document (DocClose)")
        Logger.Info("InspectEventTriggers:   700-799:  Before Save Document (BeforeDocSave)")
        Logger.Info("InspectEventTriggers:   800-899:  After Save Document (AfterDocSave)")
        Logger.Info("InspectEventTriggers:   1000-1099: After Model Parameter Change (AfterAnyParamChange)")
        Logger.Info("InspectEventTriggers:   3000-3099: After User Parameter Change (AfterAnyUserParamChange)")
        Logger.Info("InspectEventTriggers:   1200-1299: Part Geometry Change (PartBodyChanged)")
        Logger.Info("InspectEventTriggers:   1400-1499: Material Change (AfterMaterialChange)")
        Logger.Info("InspectEventTriggers:   1600-1699: iProperty Change (AfterAnyiPropertyChange)")
        Logger.Info("InspectEventTriggers:   1800-1899: Drawing View Change (AfterDrawingViewsUpdate)")
        Logger.Info("InspectEventTriggers:   2000-2099: Feature Suppression Change")
        Logger.Info("InspectEventTriggers:   2200-2299: Component Suppression Change")
        Logger.Info("InspectEventTriggers:   2400-2499: iPart/iAssembly Change")
        Logger.Info("InspectEventTriggers:   2600-2699: New Document (AfterDocNew)")
        Logger.Info("InspectEventTriggers:   2800-2899: Model State Activated")
    End If
    
    ' Also check rules
    Logger.Info("InspectEventTriggers: --- iLogic Rules in Document ---")
    Dim iLogicAuto As Object = iLogicVb.Automation
    Try
        Dim rules As Object = iLogicAuto.Rules(doc)
        If rules Is Nothing Then
            Logger.Info("InspectEventTriggers:   (no rules)")
        Else
            Dim ruleCount As Integer = 0
            For Each rule As Object In rules
                ruleCount += 1
                Try
                    Dim ruleName As String = CStr(CallByName(rule, "Name", CallType.Get))
                    Logger.Info("InspectEventTriggers:   - " & ruleName)
                Catch
                    Logger.Warn("InspectEventTriggers:   - (error reading rule name)")
                End Try
            Next
            If ruleCount = 0 Then
                Logger.Info("InspectEventTriggers:   (no rules)")
            End If
        End If
    Catch ex As Exception
        Logger.Error("InspectEventTriggers: Error reading rules: " & ex.Message)
    End Try
    
    Logger.Info("InspectEventTriggers: === Inspection Complete ===")
End Sub

Function GetEventTypeFromPropId(propId As Integer) As String
    Select Case True
        Case propId >= 400 AndAlso propId < 500
            Return "AfterDocOpen"
        Case propId >= 500 AndAlso propId < 600
            Return "DocClose"
        Case propId >= 700 AndAlso propId < 800
            Return "BeforeDocSave"
        Case propId >= 800 AndAlso propId < 900
            Return "AfterDocSave"
        Case propId >= 1000 AndAlso propId < 1100
            Return "AfterAnyParamChange"
        Case propId >= 3000 AndAlso propId < 3100
            Return "AfterAnyUserParamChange"
        Case propId >= 1200 AndAlso propId < 1300
            Return "PartBodyChanged"
        Case propId >= 1400 AndAlso propId < 1500
            Return "AfterMaterialChange"
        Case propId >= 1600 AndAlso propId < 1700
            Return "AfterAnyiPropertyChange"
        Case propId >= 1800 AndAlso propId < 1900
            Return "AfterDrawingViewsUpdate"
        Case propId >= 2000 AndAlso propId < 2100
            Return "AfterFeatureSuppressionChange"
        Case propId >= 2200 AndAlso propId < 2300
            Return "AfterComponentSuppressionChange"
        Case propId >= 2400 AndAlso propId < 2500
            Return "AfterComponentReplace"
        Case propId >= 2600 AndAlso propId < 2700
            Return "AfterDocNew"
        Case propId >= 2800 AndAlso propId < 2900
            Return "AfterModelStateActivated"
        Case Else
            Return ""
    End Select
End Function
