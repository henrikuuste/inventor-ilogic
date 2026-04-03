' ============================================================================
' TestAlternatives - Final check for any unexplored APIs
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        MessageBox.Show("Please open a derived part document", "Test")
        Exit Sub
    End If
    
    Dim info As New System.Text.StringBuilder()
    info.AppendLine("=== FINAL ALTERNATIVES CHECK ===")
    info.AppendLine()
    info.AppendLine("Document: " & doc.FullFileName)
    info.AppendLine()
    
    ' 1. Check Document.File deeply
    info.AppendLine("=== 1. Document.File exploration ===")
    Try
        Dim fileObj As Inventor.File = doc.File
        info.AppendLine("  Document.File type: " & TypeName(fileObj))
        
        ' ReferencedFiles
        info.AppendLine()
        info.AppendLine("  ReferencedFiles:")
        Dim refFiles As FilesEnumerator = fileObj.ReferencedFiles
        For i As Integer = 1 To refFiles.Count
            Dim rf As Inventor.File = refFiles.Item(i)
            info.AppendLine("    File " & i & ": " & rf.FullFileName)
            
            ' Check for any methods on File object
            Try
                Dim dn As String = rf.DisplayName
                info.AppendLine("      .DisplayName: " & dn)
            Catch : End Try
            
            ' Check for ReferencingFiles (reverse lookup)
            Try
                Dim rfs As Object = rf.ReferencingFiles
                info.AppendLine("      .ReferencingFiles: " & TypeName(rfs))
            Catch ex As Exception
                info.AppendLine("      .ReferencingFiles: " & ex.Message)
            End Try
        Next
        
        ' Check for FileDescriptors
        info.AppendLine()
        info.AppendLine("  Checking for Descriptors:")
        Try
            Dim descriptors As Object = CallByName(fileObj, "Descriptors", CallType.Get)
            info.AppendLine("    .Descriptors: " & TypeName(descriptors))
        Catch ex As Exception
            info.AppendLine("    .Descriptors: " & ex.Message)
        End Try
        
        Try
            Dim refs As Object = CallByName(fileObj, "References", CallType.Get)
            info.AppendLine("    .References: " & TypeName(refs))
        Catch ex As Exception
            info.AppendLine("    .References: " & ex.Message)
        End Try
        
    Catch ex As Exception
        info.AppendLine("  Error: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' 2. Check DesignProjectManager
    info.AppendLine("=== 2. DesignProjectManager ===")
    Try
        Dim dpm As DesignProjectManager = app.DesignProjectManager
        info.AppendLine("  DesignProjectManager available")
        info.AppendLine("  ActiveDesignProject: " & dpm.ActiveDesignProject.FullFileName)
        
        ' Check for any path resolution methods
        Try
            Dim resolve As Object = CallByName(dpm, "ResolveFile", CallType.Method, "test.ipt")
            info.AppendLine("    .ResolveFile: " & TypeName(resolve))
        Catch ex As Exception
            info.AppendLine("    .ResolveFile: " & ex.Message)
        End Try
    Catch ex As Exception
        info.AppendLine("  Error: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' 3. Check FileManager more thoroughly
    info.AppendLine("=== 3. FileManager deep check ===")
    Try
        Dim fm As FileManager = app.FileManager
        info.AppendLine("  FileManager available")
        
        ' List all methods/properties via late binding attempts
        Dim methodsToTry As String() = { _
            "CopyFile", "MoveFile", "ReplaceReferences", _
            "GetReferenceFiles", "UpdateReferences", _
            "MigrateReferences", "RedirectReferences", _
            "ChangeFilePath", "RemapReferences" _
        }
        
        For Each methodName As String In methodsToTry
            Try
                Dim result As Object = CallByName(fm, methodName, CallType.Get)
                info.AppendLine("    ." & methodName & ": " & TypeName(result))
            Catch
                Try
                    Dim result As Object = CallByName(fm, methodName, CallType.Method)
                    info.AppendLine("    ." & methodName & "(): " & TypeName(result))
                Catch ex As Exception
                    ' Only report if it's not "not found"
                    If Not ex.Message.Contains("not found") Then
                        info.AppendLine("    ." & methodName & ": " & ex.Message)
                    End If
                End Try
            End Try
        Next
    Catch ex As Exception
        info.AppendLine("  Error: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' 4. Check if there's a TransientObjects related to file management
    info.AppendLine("=== 4. TransientObjects check ===")
    Try
        Dim to_ As TransientObjects = app.TransientObjects
        info.AppendLine("  TransientObjects available")
        
        ' Check for any file-related creators
        Try
            Dim fd As Object = CallByName(to_, "CreateFileDescriptor", CallType.Method, "test.ipt")
            info.AppendLine("    .CreateFileDescriptor: " & TypeName(fd))
        Catch ex As Exception
            If Not ex.Message.Contains("not found") Then
                info.AppendLine("    .CreateFileDescriptor: " & ex.Message)
            End If
        End Try
    Catch ex As Exception
        info.AppendLine("  Error: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' 5. Check iLogicVb.Automation for any file utilities
    info.AppendLine("=== 5. iLogicVb.Automation check ===")
    Try
        Dim iLogicAuto As Object = iLogicVb.Automation
        info.AppendLine("  iLogicVb.Automation available")
        
        Dim methodsToTry2 As String() = { _
            "FileManager", "FileUtils", "UpdateReferences", _
            "ReplaceReference", "MigrateFile" _
        }
        
        For Each methodName As String In methodsToTry2
            Try
                Dim result As Object = CallByName(iLogicAuto, methodName, CallType.Get)
                info.AppendLine("    ." & methodName & ": " & TypeName(result))
            Catch
            End Try
        Next
    Catch ex As Exception
        info.AppendLine("  Error: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' 6. Check CommandManager for any file migration commands
    info.AppendLine("=== 6. Check for UI Commands we could invoke ===")
    Try
        Dim cmdMgr As CommandManager = app.CommandManager
        info.AppendLine("  CommandManager available")
        
        ' Try to find "Replace Base Component" command
        Try
            Dim controlDefs As ControlDefinitions = cmdMgr.ControlDefinitions
            info.AppendLine("  ControlDefinitions count: " & controlDefs.Count)
            
            ' Search for relevant commands
            For Each ctrlDef As ControlDefinition In controlDefs
                Dim name As String = ""
                Try
                    name = ctrlDef.DisplayName
                Catch
                    Continue For
                End Try
                
                If name.ToLower().Contains("replace") OrElse _
                   name.ToLower().Contains("base") OrElse _
                   name.ToLower().Contains("derive") OrElse _
                   name.ToLower().Contains("reference") Then
                    info.AppendLine("    Found: " & name)
                    Try
                        info.AppendLine("      InternalName: " & ctrlDef.InternalName)
                    Catch : End Try
                End If
            Next
        Catch ex As Exception
            info.AppendLine("  ControlDefinitions error: " & ex.Message)
        End Try
    Catch ex As Exception
        info.AppendLine("  Error: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' 7. Check if we can run a UI command programmatically
    info.AppendLine("=== 7. Summary ===")
    info.AppendLine("Binary editing works with same-length paths.")
    info.AppendLine("Consider designing folder structure for matching lengths.")
    
    ShowResult(info.ToString())
End Sub

Sub ShowResult(text As String)
    Dim resultForm As New System.Windows.Forms.Form()
    resultForm.Text = "Alternatives Check"
    resultForm.Width = 1000
    resultForm.Height = 800
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

