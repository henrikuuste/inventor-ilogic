' ============================================================================
' TestEventTriggerSetup - Test iLogic Event Trigger API
'
' This test creates an "Uuenda" rule and sets event triggers programmatically
' for "Any Model Parameter Change" and "Any User Parameter Change".
'
' Event Trigger PropId Ranges (from community documentation):
'   After Open Document:           AfterDocOpen                400-499
'   Close Document:                DocClose                    500-599
'   Before Save Document:          BeforeDocSave               700-799
'   After Save Document:           AfterDocSave                800-899
'   After Model Parameter Change:  AfterAnyParamChange         1000-1099
'   After User Parameter Change:   AfterAnyUserParamChange     3000-3099 (Inventor 2022+)
'   Part Geometry Change:          PartBodyChanged             1200-1299
'   Material Change:               AfterMaterialChange         1400-1499
'   iProperty Change:              AfterAnyiPropertyChange     1600-1699
'   Drawing View Change:           AfterDrawingViewsUpdate     1800-1899
'   Feature Suppression Change:    AfterFeatureSuppressionChange 2000-2099
'   Component Suppression Change:  AfterComponentSuppressionChange 2200-2299
'   iPart/iAssembly Change:        AfterComponentReplace       2400-2499
'   New Document:                  AfterDocNew                 2600-2699
'   Model State Activated:         AfterModelStateActivated    2800-2899 (Inventor 2022+)
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        Logger.Error("TestEventTriggerSetup: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestEventTriggerSetup")
        Exit Sub
    End If
    
    ' Check document type
    If doc.DocumentType = DocumentTypeEnum.kPresentationDocumentObject Then
        Logger.Error("TestEventTriggerSetup: Presentations not supported.")
        MessageBox.Show("Presentation failid ei ole toetatud.", "TestEventTriggerSetup")
        Exit Sub
    End If
    
    Logger.Info("TestEventTriggerSetup: === iLogic Event Trigger Setup Test ===")
    Logger.Info("TestEventTriggerSetup: Document: " & doc.DisplayName)
    Logger.Info("TestEventTriggerSetup: Type: " & doc.DocumentType.ToString())
    
    ' Get iLogic Automation
    Dim iLogicAuto As Object = iLogicVb.Automation
    
    ' Rule name and content
    Dim ruleName As String = "Uuenda"
    Dim ruleText As String = "' Auto-triggered update rule" & vbCrLf & _
        "' Triggers: Model Parameter Change, User Parameter Change" & vbCrLf & _
        "Sub Main()" & vbCrLf & _
        "    Logger.Info(""Uuenda: Rule triggered - "" & Now.ToString())" & vbCrLf & _
        "    ' Add update logic here" & vbCrLf & _
        "    iLogicVb.UpdateWhenDone = True" & vbCrLf & _
        "End Sub"
    
    ' Step 1: Create or update the rule
    Logger.Info("TestEventTriggerSetup: --- Step 1: Create/Update Rule ---")
    Dim existingRule As Object = Nothing
    Try
        existingRule = iLogicAuto.GetRule(doc, ruleName)
        Logger.Info("TestEventTriggerSetup: Rule '" & ruleName & "' exists - updating text...")
        existingRule.Text = ruleText
        Logger.Info("TestEventTriggerSetup: Rule text updated.")
    Catch
        Logger.Info("TestEventTriggerSetup: Rule '" & ruleName & "' not found - creating...")
        Try
            iLogicAuto.AddRule(doc, ruleName, ruleText)
            Logger.Info("TestEventTriggerSetup: Rule created successfully.")
        Catch ex As Exception
            Logger.Error("TestEventTriggerSetup: ERROR creating rule: " & ex.Message)
        End Try
    End Try
    
    ' Step 2: Get or create the Event Triggers PropertySet
    Logger.Info("TestEventTriggerSetup: --- Step 2: Access Event Triggers PropertySet ---")
    Dim eventTriggersPropSet As PropertySet = Nothing
    Dim propSetInternalName As String = "{2C540830-0723-455E-A8E2-891722EB4C3E}"
    Dim propSetDisplayName As String = "iLogicEventsRules"
    
    ' Log all existing PropertySets for debugging
    Logger.Info("TestEventTriggerSetup: Existing PropertySets in document:")
    For Each ps As PropertySet In doc.PropertySets
        Logger.Info("TestEventTriggerSetup:   Name='" & ps.Name & "', InternalName='" & ps.InternalName & "'")
    Next
    
    ' Method 1: Try PropertySetExists (available in some Inventor versions)
    Try
        Dim foundSet As PropertySet = Nothing
        If doc.PropertySets.PropertySetExists(propSetInternalName, foundSet) Then
            eventTriggersPropSet = foundSet
            Logger.Info("TestEventTriggerSetup: Found via PropertySetExists: " & foundSet.Name)
        End If
    Catch
        Logger.Info("TestEventTriggerSetup: PropertySetExists method not available")
    End Try
    
    ' Method 2: Search by GUID
    If eventTriggersPropSet Is Nothing Then
        For Each ps As PropertySet In doc.PropertySets
            If ps.InternalName = propSetInternalName Then
                eventTriggersPropSet = ps
                Logger.Info("TestEventTriggerSetup: Found by GUID search: " & ps.Name)
                Exit For
            End If
        Next
    End If
    
    ' Method 3: Search by various known names
    If eventTriggersPropSet Is Nothing Then
        Dim namesToTry() As String = {"iLogicEventsRules", "_iLogicEventsRules", "_!iLogicEventsRules"}
        For Each tryName As String In namesToTry
            Try
                eventTriggersPropSet = doc.PropertySets.Item(tryName)
                Logger.Info("TestEventTriggerSetup: Found by name: " & tryName)
                Exit For
            Catch
            End Try
        Next
    End If
    
    ' Method 4: Create if not found - try multiple formats
    If eventTriggersPropSet Is Nothing Then
        Logger.Info("TestEventTriggerSetup: PropertySet not found, attempting to create...")
        
        ' Try format 1: DisplayName, GUID
        Try
            eventTriggersPropSet = doc.PropertySets.Add(propSetDisplayName, propSetInternalName)
            Logger.Info("TestEventTriggerSetup: Created with format (DisplayName, GUID)")
        Catch ex1 As Exception
            Logger.Warn("TestEventTriggerSetup: Format 1 failed: " & ex1.Message)
            
            ' Try format 2: Just display name, no GUID
            Try
                eventTriggersPropSet = doc.PropertySets.Add(propSetDisplayName)
                Logger.Info("TestEventTriggerSetup: Created with format (DisplayName only)")
            Catch ex2 As Exception
                Logger.Warn("TestEventTriggerSetup: Format 2 failed: " & ex2.Message)
                
                ' Try format 3: Hidden name format
                Try
                    eventTriggersPropSet = doc.PropertySets.Add("_!" & propSetDisplayName, propSetInternalName)
                    Logger.Info("TestEventTriggerSetup: Created with format (_!DisplayName, GUID)")
                Catch ex3 As Exception
                    Logger.Error("TestEventTriggerSetup: All creation attempts failed")
                    Logger.Error("TestEventTriggerSetup: Last error: " & ex3.Message)
                End Try
            End Try
        End Try
    End If
    
    If eventTriggersPropSet Is Nothing Then
        Logger.Error("TestEventTriggerSetup: Could not access or create Event Triggers PropertySet.")
        Logger.Error("TestEventTriggerSetup: Try manually adding any event trigger via UI first, then run again.")
        Exit Sub
    End If
    
    ' Step 3: Show existing triggers
    Logger.Info("TestEventTriggerSetup: --- Step 3: Existing Triggers ---")
    Logger.Info("TestEventTriggerSetup: PropertySet InternalName: " & eventTriggersPropSet.InternalName)
    Logger.Info("TestEventTriggerSetup: PropertySet Name: " & eventTriggersPropSet.Name)
    Logger.Info("TestEventTriggerSetup: Existing properties in this PropertySet:")
    
    For Each prop As [Property] In eventTriggersPropSet
        Try
            Logger.Info("TestEventTriggerSetup:   PropId=" & prop.PropId & ", Name='" & prop.Name & "', Value='" & prop.Value.ToString() & "'")
        Catch
            Logger.Warn("TestEventTriggerSetup:   PropId=" & prop.PropId & " (error reading)")
        End Try
    Next
    
    ' Step 4: Add event triggers for parameter changes
    Logger.Info("TestEventTriggerSetup: --- Step 4: Add Parameter Change Triggers ---")
    
    ' Model Parameter Change (PropId range 1000-1099)
    Logger.Info("TestEventTriggerSetup: Adding Model Parameter Change trigger (PropId 1000)...")
    AddOrUpdateTrigger(eventTriggersPropSet, ruleName, "AfterAnyParamChange", 1000)
    
    ' User Parameter Change (PropId range 3000-3099) - Inventor 2022+
    Logger.Info("TestEventTriggerSetup: Adding User Parameter Change trigger (PropId 3000)...")
    AddOrUpdateTrigger(eventTriggersPropSet, ruleName, "AfterAnyUserParamChange", 3000)
    
    ' Step 5: Verify triggers
    Logger.Info("TestEventTriggerSetup: --- Step 5: Verify Triggers ---")
    Logger.Info("TestEventTriggerSetup: Properties after adding triggers:")
    For Each prop As [Property] In eventTriggersPropSet
        Try
            Logger.Info("TestEventTriggerSetup:   PropId=" & prop.PropId & ", Name='" & prop.Name & "', Value='" & prop.Value.ToString() & "'")
        Catch
            Logger.Warn("TestEventTriggerSetup:   PropId=" & prop.PropId & " (error reading)")
        End Try
    Next
    
    ' Step 6: Try to read rule's trigger properties
    Logger.Info("TestEventTriggerSetup: --- Step 6: Rule Trigger Properties ---")
    Try
        Dim rule As Object = iLogicAuto.GetRule(doc, ruleName)
        
        ' Try various known properties
        TryReadProperty(rule, "FireOnParameterChange")
        TryReadProperty(rule, "FireOnUserParameterChange")
        TryReadProperty(rule, "ParameterTriggers")
        TryReadProperty(rule, "EventTriggers")
        TryReadProperty(rule, "Triggers")
    Catch ex As Exception
        Logger.Error("TestEventTriggerSetup: Error getting rule: " & ex.Message)
    End Try
    
    Logger.Info("TestEventTriggerSetup: === Test Complete ===")
    Logger.Info("TestEventTriggerSetup: Save the document and test by changing a parameter.")
    Logger.Info("TestEventTriggerSetup: Check for 'Uuenda: Rule triggered' messages.")
End Sub

Sub AddOrUpdateTrigger(ByVal propSet As PropertySet, ByVal ruleName As String, _
                       ByVal triggerName As String, ByVal basePropId As Integer)
    ' Check if this rule is already registered for this event
    Dim alreadyExists As Boolean = False
    Dim existingPropId As Integer = 0
    
    ' Check all properties in the valid range (basePropId to basePropId+99)
    For propId As Integer = basePropId To basePropId + 99
        Try
            Dim existingProp As [Property] = propSet.ItemByPropId(propId)
            If existingProp IsNot Nothing Then
                If CStr(existingProp.Value) = ruleName Then
                    alreadyExists = True
                    existingPropId = propId
                    Logger.Info("TestEventTriggerSetup:   Rule already registered at PropId " & propId)
                    Exit For
                End If
            End If
        Catch
            ' Property doesn't exist at this PropId - this is expected
        End Try
    Next
    
    If alreadyExists Then
        Logger.Info("TestEventTriggerSetup:   Trigger already exists, no action needed.")
        Return
    End If
    
    ' Find next available PropId in range
    Dim nextPropId As Integer = basePropId
    For propId As Integer = basePropId To basePropId + 99
        Try
            Dim existingProp As [Property] = propSet.ItemByPropId(propId)
            ' Property exists, try next
            nextPropId = propId + 1
        Catch
            ' Property doesn't exist - use this one
            nextPropId = propId
            Exit For
        End Try
    Next
    
    ' The property name format is: TriggerName + index (e.g., "AfterAnyParamChange0")
    ' The index is the offset from the base PropId
    Dim propIndex As Integer = nextPropId - basePropId
    Dim propName As String = triggerName & propIndex.ToString()
    
    Logger.Info("TestEventTriggerSetup:   Adding property: Name='" & propName & "', PropId=" & nextPropId & ", Value='" & ruleName & "'")
    
    Try
        propSet.Add(ruleName, propName, nextPropId)
        Logger.Info("TestEventTriggerSetup:   SUCCESS: Trigger added.")
    Catch ex As Exception
        Logger.Error("TestEventTriggerSetup:   ERROR: " & ex.Message)
    End Try
End Sub

Sub TryReadProperty(ByVal obj As Object, ByVal propName As String)
    Try
        Dim value As Object = CallByName(obj, propName, CallType.Get)
        If value Is Nothing Then
            Logger.Info("TestEventTriggerSetup:   ." & propName & ": Nothing")
        ElseIf TypeOf value Is String OrElse TypeOf value Is Boolean OrElse TypeOf value Is Integer Then
            Logger.Info("TestEventTriggerSetup:   ." & propName & ": " & value.ToString())
        Else
            Logger.Info("TestEventTriggerSetup:   ." & propName & ": " & TypeName(value))
            ' Try to get Count if it's a collection
            Try
                Dim count As Integer = CInt(CallByName(value, "Count", CallType.Get))
                Logger.Info("TestEventTriggerSetup:     Count: " & count)
            Catch
            End Try
        End If
    Catch ex As Exception
        Logger.Info("TestEventTriggerSetup:   ." & propName & ": " & ex.Message)
    End Try
End Sub
