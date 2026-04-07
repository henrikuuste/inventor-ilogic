' ============================================================================
' RemoveEventTriggers - Remove event triggers for a specific rule
'
' This script removes all event triggers for the "Uuenda" rule (or specified rule)
' from the current document. Use this to reset and test again.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        Logger.Error("RemoveEventTriggers: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "RemoveEventTriggers")
        Exit Sub
    End If
    
    Dim ruleName As String = InputBox("Sisesta reegli nimi, mille päästikud eemaldada:", _
                                       "Eemalda päästikud", "Uuenda")
    
    If String.IsNullOrEmpty(ruleName) Then
        Logger.Info("RemoveEventTriggers: Cancelled by user.")
        Exit Sub
    End If
    
    Logger.Info("RemoveEventTriggers: === Remove Event Triggers ===")
    Logger.Info("RemoveEventTriggers: Document: " & doc.DisplayName)
    Logger.Info("RemoveEventTriggers: Rule to clear: " & ruleName)
    
    ' Find event triggers PropertySet
    Dim eventTriggersPropSet As PropertySet = Nothing
    
    ' Try various known names
    Dim possibleNames() As String = {"iLogicEventsRules", "_iLogicEventsRules", "_!iLogicEventsRules"}
    
    For Each psName As String In possibleNames
        Try
            eventTriggersPropSet = doc.PropertySets.Item(psName)
            Logger.Info("RemoveEventTriggers: Found PropertySet: '" & psName & "'")
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
                Logger.Info("RemoveEventTriggers: Found PropertySet by GUID")
                Exit For
            End If
        Next
    End If
    
    If eventTriggersPropSet Is Nothing Then
        Logger.Info("RemoveEventTriggers: No event triggers PropertySet found - nothing to remove.")
        Exit Sub
    End If
    
    ' Find and remove all triggers for this rule
    Logger.Info("RemoveEventTriggers: Searching for triggers with rule '" & ruleName & "'...")
    
    Dim propsToDelete As New System.Collections.Generic.List(Of Integer)
    
    For Each prop As [Property] In eventTriggersPropSet
        Try
            If CStr(prop.Value) = ruleName Then
                Logger.Info("RemoveEventTriggers:   Found: PropId=" & prop.PropId & ", Name='" & prop.Name & "'")
                propsToDelete.Add(prop.PropId)
            End If
        Catch
        End Try
    Next
    
    If propsToDelete.Count = 0 Then
        Logger.Info("RemoveEventTriggers: No triggers found for rule '" & ruleName & "'.")
    Else
        Logger.Info("RemoveEventTriggers: Deleting " & propsToDelete.Count & " trigger(s)...")
        
        For Each propId As Integer In propsToDelete
            Try
                Dim prop As [Property] = eventTriggersPropSet.ItemByPropId(propId)
                prop.Delete()
                Logger.Info("RemoveEventTriggers:   Deleted PropId=" & propId)
            Catch ex As Exception
                Logger.Error("RemoveEventTriggers:   Error deleting PropId=" & propId & ": " & ex.Message)
            End Try
        Next
        
        Logger.Info("RemoveEventTriggers: Done. Save the document to persist changes.")
    End If
    
    Logger.Info("RemoveEventTriggers: === Removal Complete ===")
End Sub
