' ============================================================================
' DocumentUpdateLib - Manage Local "Uuenda" Update Rule
'
' This library manages a local "Uuenda" iLogic rule that multiple scripts can
' register update handlers with. Each handler is identified by a unique UID
' and the library handles rule creation, section management, and event triggers.
'
' Usage: AddVbFile "Lib/DocumentUpdateLib.vb"
'
' Example:
'   Dim updateCode() As String = {"MyLib.DoUpdate(ThisApplication, ThisDoc.Document)"}
'   Dim triggers() As DocumentUpdateLib.UpdateTrigger = {
'       DocumentUpdateLib.UpdateTrigger.ModelParameterChange,
'       DocumentUpdateLib.UpdateTrigger.UserParameterChange
'   }
'   DocumentUpdateLib.RegisterUpdateHandler(doc, iLogicAuto, "MyFeature", updateCode, triggers)
' ============================================================================

Imports Inventor

Public Module DocumentUpdateLib

    ' ============================================================================
    ' CONSTANTS
    ' ============================================================================
    
    Public Const RULE_NAME As String = "Uuenda"
    Private Const SECTION_BEGIN As String = "' === BEGIN: "
    Private Const SECTION_END As String = "' === END: "
    Private Const SECTION_MARKER_END As String = " ==="
    Private Const EVENT_PROPSET_GUID As String = "{2C540830-0723-455E-A8E2-891722EB4C3E}"
    
    ' ============================================================================
    ' TRIGGER ENUM
    ' ============================================================================
    
    ''' <summary>
    ''' Event triggers for the Uuenda rule. Callers use this enum instead of PropIds.
    ''' </summary>
    Public Enum UpdateTrigger
        ModelParameterChange = 1
        UserParameterChange = 2
        BeforeSave = 3
        AfterSave = 4
        DocumentOpen = 5
        PartGeometryChange = 6
        MaterialChange = 7
        iPropertyChange = 8
    End Enum
    
    ' ============================================================================
    ' LOGGING
    ' ============================================================================
    
    Private m_Logger As Object = Nothing
    
    Public Sub SetLogger(logger As Object)
        m_Logger = logger
    End Sub
    
    Private Sub LogInfo(message As String)
        If m_Logger IsNot Nothing Then
            Try
                m_Logger.Info("DocumentUpdateLib: " & message)
            Catch
            End Try
        End If
    End Sub
    
    Private Sub LogWarn(message As String)
        If m_Logger IsNot Nothing Then
            Try
                m_Logger.Warn("DocumentUpdateLib: " & message)
            Catch
            End Try
        End If
    End Sub
    
    Private Sub LogError(message As String)
        If m_Logger IsNot Nothing Then
            Try
                m_Logger.Error("DocumentUpdateLib: " & message)
            Catch
            End Try
        End If
    End Sub
    
    ' ============================================================================
    ' PUBLIC API
    ' ============================================================================
    
    ''' <summary>
    ''' Registers or updates an update handler section in the Uuenda rule.
    ''' Creates the rule and triggers if they don't exist.
    ''' </summary>
    ''' <param name="doc">The document to add the handler to</param>
    ''' <param name="iLogicAuto">iLogicVb.Automation object</param>
    ''' <param name="uid">Unique identifier for this handler section</param>
    ''' <param name="codeLines">Array of VB code lines to execute</param>
    ''' <param name="triggers">Array of UpdateTrigger enum values</param>
    ''' <returns>True if successful</returns>
    Public Function RegisterUpdateHandler(ByVal doc As Document, ByVal iLogicAuto As Object, _
                                          ByVal uid As String, ByVal codeLines() As String, _
                                          ByVal triggers() As UpdateTrigger) As Boolean
        If doc Is Nothing OrElse iLogicAuto Is Nothing Then
            LogError("RegisterUpdateHandler: doc or iLogicAuto is Nothing")
            Return False
        End If
        
        If String.IsNullOrEmpty(uid) Then
            LogError("RegisterUpdateHandler: uid is empty")
            Return False
        End If
        
        ' Ensure rule exists
        If Not EnsureUpdateRule(doc, iLogicAuto) Then
            LogError("RegisterUpdateHandler: Failed to ensure update rule")
            Return False
        End If
        
        ' Add triggers (skips duplicates internally)
        If triggers IsNot Nothing Then
            For Each trigger As UpdateTrigger In triggers
                AddTrigger(doc, trigger)
            Next
        End If
        
        ' Get current rule text
        Dim rule As Object = Nothing
        Try
            rule = iLogicAuto.GetRule(doc, RULE_NAME)
        Catch
            LogError("RegisterUpdateHandler: Failed to get rule after creation")
            Return False
        End Try
        
        Dim currentText As String = ""
        Try
            currentText = CStr(CallByName(rule, "Text", CallType.Get))
        Catch
            LogError("RegisterUpdateHandler: Failed to read rule text")
            Return False
        End Try
        
        ' Parse, update section, rebuild
        Dim header As String = ""
        Dim sections As System.Collections.Generic.Dictionary(Of String, String) = Nothing
        ParseRuleSections(currentText, header, sections)
        
        ' Build section content
        Dim sectionContent As New System.Text.StringBuilder()
        If codeLines IsNot Nothing Then
            For Each line As String In codeLines
                sectionContent.AppendLine("    " & line)
            Next
        End If
        
        ' Add or replace section
        sections(uid) = sectionContent.ToString().TrimEnd()
        
        ' Rebuild and update rule
        Dim newText As String = BuildRuleText(header, sections)
        Try
            rule.Text = newText
            LogInfo("Registered handler: " & uid)
            Return True
        Catch ex As Exception
            LogError("RegisterUpdateHandler: Failed to update rule text - " & ex.Message)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Removes an update handler section from the Uuenda rule.
    ''' </summary>
    ''' <param name="doc">The document</param>
    ''' <param name="iLogicAuto">iLogicVb.Automation object</param>
    ''' <param name="uid">Unique identifier of the handler to remove</param>
    ''' <returns>True if removed, False if not found or error</returns>
    Public Function RemoveUpdateHandler(ByVal doc As Document, ByVal iLogicAuto As Object, _
                                        ByVal uid As String) As Boolean
        If doc Is Nothing OrElse iLogicAuto Is Nothing Then Return False
        If String.IsNullOrEmpty(uid) Then Return False
        
        ' Get rule
        Dim rule As Object = Nothing
        Try
            rule = iLogicAuto.GetRule(doc, RULE_NAME)
        Catch
            Return False ' Rule doesn't exist, nothing to remove
        End Try
        
        Dim currentText As String = ""
        Try
            currentText = CStr(CallByName(rule, "Text", CallType.Get))
        Catch
            Return False
        End Try
        
        ' Parse sections
        Dim header As String = ""
        Dim sections As System.Collections.Generic.Dictionary(Of String, String) = Nothing
        ParseRuleSections(currentText, header, sections)
        
        ' Remove section if exists
        If Not sections.ContainsKey(uid) Then
            Return False
        End If
        
        sections.Remove(uid)
        
        ' Rebuild and update rule
        Dim newText As String = BuildRuleText(header, sections)
        Try
            rule.Text = newText
            LogInfo("Removed handler: " & uid)
            Return True
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Ensures the Uuenda rule exists in the document. Creates it if missing.
    ''' </summary>
    ''' <returns>True if rule exists or was created</returns>
    Public Function EnsureUpdateRule(ByVal doc As Document, ByVal iLogicAuto As Object) As Boolean
        If doc Is Nothing OrElse iLogicAuto Is Nothing Then Return False
        
        ' Check if rule exists
        Try
            Dim existingRule As Object = iLogicAuto.GetRule(doc, RULE_NAME)
            If existingRule IsNot Nothing Then
                Return True
            End If
        Catch
            ' Rule doesn't exist, create it
        End Try
        
        ' Create empty rule
        Dim ruleText As String = _
            "' Uuenda - Auto-triggered update rule" & vbCrLf & _
            "' Managed by DocumentUpdateLib - do not edit manually" & vbCrLf & _
            vbCrLf & _
            "Sub Main()" & vbCrLf & _
            "End Sub"
        
        Try
            iLogicAuto.AddRule(doc, RULE_NAME, ruleText)
            LogInfo("Created Uuenda rule")
            Return True
        Catch ex As Exception
            LogError("EnsureUpdateRule: Failed to create rule - " & ex.Message)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Adds an event trigger to the Uuenda rule if not already present.
    ''' </summary>
    ''' <param name="doc">The document</param>
    ''' <param name="trigger">The trigger type to add</param>
    ''' <returns>True if added or already exists</returns>
    Public Function AddTrigger(ByVal doc As Document, ByVal trigger As UpdateTrigger) As Boolean
        If doc Is Nothing Then Return False
        
        ' Get trigger info
        Dim propId As Integer = 0
        Dim propName As String = ""
        GetTriggerInfo(trigger, propId, propName)
        
        If propId = 0 Then
            LogWarn("AddTrigger: Unknown trigger type")
            Return False
        End If
        
        ' Get or create PropertySet
        Dim propSet As PropertySet = GetEventTriggersPropertySet(doc)
        If propSet Is Nothing Then
            LogError("AddTrigger: Could not get event triggers PropertySet")
            Return False
        End If
        
        ' Check if trigger already exists for this rule
        If HasTrigger(propSet, RULE_NAME, propId) Then
            Return True ' Already exists, success
        End If
        
        ' Find next available PropId in range
        Dim nextPropId As Integer = propId
        For pid As Integer = propId To propId + 99
            Try
                Dim existingProp As [Property] = propSet.ItemByPropId(pid)
                nextPropId = pid + 1
            Catch
                nextPropId = pid
                Exit For
            End Try
        Next
        
        ' Add the trigger
        Dim propIndex As Integer = nextPropId - propId
        Dim fullPropName As String = propName & propIndex.ToString()
        
        Try
            propSet.Add(RULE_NAME, fullPropName, nextPropId)
            LogInfo("Added trigger: " & trigger.ToString())
            Return True
        Catch ex As Exception
            LogError("AddTrigger: Failed to add - " & ex.Message)
            Return False
        End Try
    End Function
    
    ' ============================================================================
    ' INTERNAL HELPERS
    ' ============================================================================
    
    ''' <summary>
    ''' Gets the event triggers PropertySet, creating if necessary.
    ''' </summary>
    Private Function GetEventTriggersPropertySet(ByVal doc As Document) As PropertySet
        ' Try PropertySetExists method first
        Try
            Dim foundSet As PropertySet = Nothing
            If doc.PropertySets.PropertySetExists(EVENT_PROPSET_GUID, foundSet) Then
                Return foundSet
            End If
        Catch
        End Try
        
        ' Search by GUID
        For Each ps As PropertySet In doc.PropertySets
            If ps.InternalName = EVENT_PROPSET_GUID Then
                Return ps
            End If
        Next
        
        ' Search by known names
        Dim namesToTry() As String = {"iLogicEventsRules", "_iLogicEventsRules", "_!iLogicEventsRules"}
        For Each tryName As String In namesToTry
            Try
                Return doc.PropertySets.Item(tryName)
            Catch
            End Try
        Next
        
        ' Create if not found
        Try
            Return doc.PropertySets.Add("iLogicEventsRules", EVENT_PROPSET_GUID)
        Catch
        End Try
        
        Try
            Return doc.PropertySets.Add("_!iLogicEventsRules", EVENT_PROPSET_GUID)
        Catch
        End Try
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Maps UpdateTrigger enum to PropId and PropertyName.
    ''' </summary>
    Private Sub GetTriggerInfo(ByVal trigger As UpdateTrigger, ByRef propId As Integer, ByRef propName As String)
        Select Case trigger
            Case UpdateTrigger.ModelParameterChange
                propId = 1000
                propName = "AfterAnyParamChange"
            Case UpdateTrigger.UserParameterChange
                propId = 3000
                propName = "AfterAnyUserParamChange"
            Case UpdateTrigger.BeforeSave
                propId = 700
                propName = "BeforeDocSave"
            Case UpdateTrigger.AfterSave
                propId = 800
                propName = "AfterDocSave"
            Case UpdateTrigger.DocumentOpen
                propId = 400
                propName = "AfterDocOpen"
            Case UpdateTrigger.PartGeometryChange
                propId = 1200
                propName = "PartBodyChanged"
            Case UpdateTrigger.MaterialChange
                propId = 1400
                propName = "AfterMaterialChange"
            Case UpdateTrigger.iPropertyChange
                propId = 1600
                propName = "AfterAnyiPropertyChange"
            Case Else
                propId = 0
                propName = ""
        End Select
    End Sub
    
    ''' <summary>
    ''' Checks if a trigger already exists for the given rule.
    ''' </summary>
    Private Function HasTrigger(ByVal propSet As PropertySet, ByVal ruleName As String, ByVal basePropId As Integer) As Boolean
        For pid As Integer = basePropId To basePropId + 99
            Try
                Dim prop As [Property] = propSet.ItemByPropId(pid)
                If prop IsNot Nothing AndAlso CStr(prop.Value) = ruleName Then
                    Return True
                End If
            Catch
            End Try
        Next
        Return False
    End Function
    
    ''' <summary>
    ''' Parses rule text into header (before Sub Main content) and UID-keyed sections.
    ''' </summary>
    Private Sub ParseRuleSections(ByVal ruleText As String, ByRef header As String, _
                                  ByRef sections As System.Collections.Generic.Dictionary(Of String, String))
        sections = New System.Collections.Generic.Dictionary(Of String, String)()
        
        If String.IsNullOrEmpty(ruleText) Then
            header = "' Uuenda - Auto-triggered update rule" & vbCrLf & _
                     "' Managed by DocumentUpdateLib - do not edit manually" & vbCrLf & vbCrLf
            Return
        End If
        
        Dim lines() As String = ruleText.Split({vbCrLf, vbLf}, StringSplitOptions.None)
        Dim headerLines As New System.Text.StringBuilder()
        Dim inSubMain As Boolean = False
        Dim inSection As Boolean = False
        Dim currentUid As String = ""
        Dim currentSectionLines As New System.Text.StringBuilder()
        
        For Each line As String In lines
            Dim trimmedLine As String = line.Trim()
            
            ' Check for Sub Main start
            If trimmedLine.StartsWith("Sub Main", StringComparison.OrdinalIgnoreCase) Then
                inSubMain = True
                Continue For
            End If
            
            ' Check for End Sub
            If trimmedLine.Equals("End Sub", StringComparison.OrdinalIgnoreCase) Then
                ' Save any current section
                If inSection AndAlso currentUid <> "" Then
                    sections(currentUid) = currentSectionLines.ToString().TrimEnd()
                End If
                inSubMain = False
                Continue For
            End If
            
            If Not inSubMain Then
                ' Before Sub Main - add to header
                headerLines.AppendLine(line)
            Else
                ' Inside Sub Main - look for section markers
                If trimmedLine.StartsWith(SECTION_BEGIN) Then
                    ' Save previous section if any
                    If inSection AndAlso currentUid <> "" Then
                        sections(currentUid) = currentSectionLines.ToString().TrimEnd()
                    End If
                    
                    ' Start new section
                    Dim startIdx As Integer = SECTION_BEGIN.Length
                    Dim endIdx As Integer = trimmedLine.IndexOf(SECTION_MARKER_END, startIdx)
                    If endIdx > startIdx Then
                        currentUid = trimmedLine.Substring(startIdx, endIdx - startIdx)
                        inSection = True
                        currentSectionLines.Clear()
                    End If
                ElseIf trimmedLine.StartsWith(SECTION_END) Then
                    ' End current section
                    If inSection AndAlso currentUid <> "" Then
                        sections(currentUid) = currentSectionLines.ToString().TrimEnd()
                    End If
                    inSection = False
                    currentUid = ""
                ElseIf inSection Then
                    ' Content inside a section
                    currentSectionLines.AppendLine(line)
                End If
            End If
        Next
        
        header = headerLines.ToString()
    End Sub
    
    ''' <summary>
    ''' Builds rule text from header and sections dictionary.
    ''' Includes a document update call at the end to ensure formula parameters recalculate.
    ''' </summary>
    Private Function BuildRuleText(ByVal header As String, _
                                   ByVal sections As System.Collections.Generic.Dictionary(Of String, String)) As String
        Dim result As New System.Text.StringBuilder()
        
        ' Add header (comments before Sub Main)
        If Not String.IsNullOrEmpty(header) Then
            result.Append(header.TrimEnd())
            result.AppendLine()
            result.AppendLine()
        Else
            result.AppendLine("' Uuenda - Auto-triggered update rule")
            result.AppendLine("' Managed by DocumentUpdateLib - do not edit manually")
            result.AppendLine()
        End If
        
        ' Start Sub Main
        result.AppendLine("Sub Main()")
        
        ' Add initial update to ensure formula-based parameters are current before handlers run
        If sections IsNot Nothing AndAlso sections.Count > 0 Then
            result.AppendLine("    ' Update document first to recalculate formula dependencies")
            result.AppendLine("    ThisDoc.Document.Update()")
            result.AppendLine()
        End If
        
        ' Add sections
        If sections IsNot Nothing AndAlso sections.Count > 0 Then
            Dim isFirst As Boolean = True
            For Each kvp As System.Collections.Generic.KeyValuePair(Of String, String) In sections
                If Not isFirst Then
                    result.AppendLine()
                End If
                isFirst = False
                
                result.AppendLine("    " & SECTION_BEGIN & kvp.Key & SECTION_MARKER_END)
                If Not String.IsNullOrEmpty(kvp.Value) Then
                    result.AppendLine(kvp.Value)
                End If
                result.AppendLine("    " & SECTION_END & kvp.Key & SECTION_MARKER_END)
            Next
            
            ' Update document to recalculate formula-based parameters
            result.AppendLine()
            result.AppendLine("    ' Update document to recalculate dependent formulas")
            result.AppendLine("    ThisDoc.Document.Update()")
        End If
        
        ' End Sub Main
        result.AppendLine("End Sub")
        
        Return result.ToString()
    End Function

End Module
