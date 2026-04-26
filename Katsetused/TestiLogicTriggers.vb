' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' TestiLogicTriggers - Diagnose iLogic rule triggers
'
' This test examines all iLogic rules in the document and shows their
' trigger configurations to help debug why rules aren't running.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        MessageBox.Show("Please open a document with iLogic rules.", "Test")
        Exit Sub
    End If
    
    Dim info As New System.Text.StringBuilder()
    info.AppendLine("=== iLogic Rule Trigger Analysis ===")
    info.AppendLine()
    info.AppendLine("Document: " & doc.FullFileName)
    info.AppendLine()
    
    Dim iLogicAuto As Object = Nothing
    Try
        iLogicAuto = iLogicVb.Automation
        info.AppendLine("iLogicVb.Automation: Available")
    Catch ex As Exception
        info.AppendLine("iLogicVb.Automation ERROR: " & ex.Message)
        ShowResult(info.ToString())
        Exit Sub
    End Try
    info.AppendLine()
    
    ' Get all rules
    info.AppendLine("=== Rules in this document ===")
    Dim rules As Object = Nothing
    Try
        rules = iLogicAuto.Rules(doc)
        info.AppendLine("Rules collection type: " & TypeName(rules))
    Catch ex As Exception
        info.AppendLine("Error getting rules: " & ex.Message)
        ShowResult(info.ToString())
        Exit Sub
    End Try
    
    If rules Is Nothing Then
        info.AppendLine("No rules found in this document (Rules returned Nothing).")
        info.AppendLine()
        info.AppendLine("=== Test Complete ===")
        ShowResult(info.ToString())
        Exit Sub
    End If
    
    info.AppendLine()
    
    ' Analyze each rule
    For Each rule As Object In rules
        info.AppendLine("----------------------------------------")
        
        Dim ruleName As String = ""
        Try
            ruleName = CStr(CallByName(rule, "Name", CallType.Get))
            info.AppendLine("Rule: " & ruleName)
        Catch ex As Exception
            info.AppendLine("Rule: (error getting name) - " & ex.Message)
            Continue For
        End Try
        
        ' Check various trigger properties
        info.AppendLine("  Checking trigger properties:")
        
        ' FireOnParameterChange
        Try
            Dim val As Object = CallByName(rule, "FireOnParameterChange", CallType.Get)
            info.AppendLine("    .FireOnParameterChange: " & val.ToString())
        Catch ex As Exception
            info.AppendLine("    .FireOnParameterChange: " & ex.Message)
        End Try
        
        ' FireOnOpen
        Try
            Dim val As Object = CallByName(rule, "FireOnOpen", CallType.Get)
            info.AppendLine("    .FireOnOpen: " & val.ToString())
        Catch ex As Exception
            info.AppendLine("    .FireOnOpen: " & ex.Message)
        End Try
        
        ' FireOnClose
        Try
            Dim val As Object = CallByName(rule, "FireOnClose", CallType.Get)
            info.AppendLine("    .FireOnClose: " & val.ToString())
        Catch ex As Exception
            info.AppendLine("    .FireOnClose: " & ex.Message)
        End Try
        
        ' FireOnSave
        Try
            Dim val As Object = CallByName(rule, "FireOnSave", CallType.Get)
            info.AppendLine("    .FireOnSave: " & val.ToString())
        Catch ex As Exception
            info.AppendLine("    .FireOnSave: " & ex.Message)
        End Try
        
        ' Triggers collection
        Try
            Dim triggers As Object = CallByName(rule, "Triggers", CallType.Get)
            info.AppendLine("    .Triggers: " & TypeName(triggers))
            
            If triggers IsNot Nothing Then
                Try
                    Dim count As Integer = CInt(CallByName(triggers, "Count", CallType.Get))
                    info.AppendLine("      Count: " & count)
                    
                    For i As Integer = 1 To count
                        Try
                            Dim trigger As Object = CallByName(triggers, "Item", CallType.Get, i)
                            info.AppendLine("      Trigger " & i & ":")
                            
                            ' Try to get trigger properties
                            Try
                                Dim trigType As Object = CallByName(trigger, "TriggerType", CallType.Get)
                                info.AppendLine("        .TriggerType: " & trigType.ToString())
                            Catch : End Try
                            
                            Try
                                Dim trigName As Object = CallByName(trigger, "Name", CallType.Get)
                                info.AppendLine("        .Name: " & trigName.ToString())
                            Catch : End Try
                            
                            Try
                                Dim enabled As Object = CallByName(trigger, "Enabled", CallType.Get)
                                info.AppendLine("        .Enabled: " & enabled.ToString())
                            Catch : End Try
                        Catch
                        End Try
                    Next
                Catch ex As Exception
                    info.AppendLine("      Error iterating: " & ex.Message)
                End Try
            End If
        Catch ex As Exception
            info.AppendLine("    .Triggers: " & ex.Message)
        End Try
        
        ' EventTriggers
        Try
            Dim eventTriggers As Object = CallByName(rule, "EventTriggers", CallType.Get)
            info.AppendLine("    .EventTriggers: " & TypeName(eventTriggers))
        Catch ex As Exception
            info.AppendLine("    .EventTriggers: " & ex.Message)
        End Try
        
        ' ParameterTriggers
        Try
            Dim paramTriggers As Object = CallByName(rule, "ParameterTriggers", CallType.Get)
            info.AppendLine("    .ParameterTriggers: " & TypeName(paramTriggers))
            
            If paramTriggers IsNot Nothing Then
                Try
                    Dim count As Integer = CInt(CallByName(paramTriggers, "Count", CallType.Get))
                    info.AppendLine("      Count: " & count)
                    
                    For i As Integer = 1 To count
                        Try
                            Dim pt As Object = CallByName(paramTriggers, "Item", CallType.Get, i)
                            Dim ptName As String = CStr(CallByName(pt, "Name", CallType.Get))
                            info.AppendLine("      - " & ptName)
                        Catch
                        End Try
                    Next
                Catch
                End Try
            End If
        Catch ex As Exception
            info.AppendLine("    .ParameterTriggers: " & ex.Message)
        End Try
        
        info.AppendLine()
    Next
    
    ' Try to manually run a rule
    info.AppendLine("=== Test Manual Rule Execution ===")
    Dim testRuleName As String = InputBox("Enter rule name to test running (or cancel):", "Test Run", "Update")
    
    If Not String.IsNullOrEmpty(testRuleName) Then
        info.AppendLine("Attempting to run rule: " & testRuleName)
        Try
            iLogicAuto.RunRule(doc, testRuleName)
            info.AppendLine("  RunRule completed (no exception)")
        Catch ex As Exception
            info.AppendLine("  RunRule ERROR: " & ex.Message)
        End Try
    End If
    
    ShowResult(info.ToString())
End Sub

Sub ShowResult(text As String)
    Dim resultForm As New System.Windows.Forms.Form()
    resultForm.Text = "iLogic Trigger Analysis"
    resultForm.Width = 900
    resultForm.Height = 700
    resultForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    
    Dim txt As New System.Windows.Forms.TextBox()
    txt.Multiline = True
    txt.ScrollBars = System.Windows.Forms.ScrollBars.Both
    txt.Dock = System.Windows.Forms.DockStyle.Fill
    txt.Text = text
    txt.ReadOnly = True
    
    resultForm.Controls.Add(txt)
    resultForm.ShowDialog()
End Sub

