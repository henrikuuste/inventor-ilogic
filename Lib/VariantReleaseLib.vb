' ============================================================================
' VariantReleaseLib - Core Variant Release Functionality
' 
' Provides functions for copying assemblies with all dependencies to variant
' folders and updating all references to point to the copied files.
'
' Usage: AddVbFile "Lib/VariantReleaseLib.vb"
'
' Key Functions:
' - GetAllReferencedFiles: Collect full dependency tree
' - FindAllDrawings: Find drawings that reference the assembly/parts
' - BuildCopyMap: Create mapping from source to target paths
' - UpdateAllReferences: Update assembly, sub-assembly, and derived part refs
' - ApplyParameters: Set parameter values on the copied assembly
' - RunParameterChangeRules: Execute iLogic rules with param change triggers
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=ComponentOccurrence_Replace
' ============================================================================

Imports Inventor
Imports System.Collections.Generic

Public Module VariantReleaseLib

    ' Track which files have been processed to avoid infinite loops
    Private ProcessedFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    ' ============================================================================
    ' SECTION 1: File Discovery
    ' ============================================================================

    ''' <summary>
    ''' Get all files referenced by an assembly, recursively.
    ''' Returns a list of full file paths.
    ''' </summary>
    Public Function GetAllReferencedFiles(asmDoc As AssemblyDocument) As List(Of String)
        Dim files As New List(Of String)
        
        ' Add the assembly itself
        files.Add(asmDoc.FullFileName)
        
        ' Get all referenced documents recursively
        For Each refDoc As Document In asmDoc.AllReferencedDocuments
            If Not files.Contains(refDoc.FullFileName) Then
                files.Add(refDoc.FullFileName)
            End If
        Next
        
        Return files
    End Function

    ''' <summary>
    ''' Find all drawing files (IDW/DWG) that reference any of the given files.
    ''' Searches in the specified folder and subfolders.
    ''' </summary>
    Public Function FindAllDrawings(app As Inventor.Application, searchFolder As String, _
                                    referencedFiles As List(Of String), _
                                    Optional ByRef logMessages As List(Of String) = Nothing) As List(Of String)
        Dim drawings As New List(Of String)
        
        If logMessages Is Nothing Then logMessages = New List(Of String)
        
        If Not System.IO.Directory.Exists(searchFolder) Then
            logMessages.Add("  Drawing search folder does not exist: " & searchFolder)
            Return drawings
        End If
        
        logMessages.Add("  Searching for drawings in: " & searchFolder)
        
        ' Build a set of referenced file paths for fast lookup (using normalized paths)
        Dim refFileSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each f As String In referencedFiles
            Dim normalized As String = System.IO.Path.GetFullPath(f)
            refFileSet.Add(normalized)
        Next
        
        logMessages.Add("  Looking for drawings that reference " & refFileSet.Count & " files")
        
        ' Find all IDW and DWG files
        Dim drawingFiles As New List(Of String)
        Try
            drawingFiles.AddRange(System.IO.Directory.GetFiles(searchFolder, "*.idw", System.IO.SearchOption.AllDirectories))
        Catch ex As Exception
            logMessages.Add("  Error searching for IDW files: " & ex.Message)
        End Try
        
        Try
            drawingFiles.AddRange(System.IO.Directory.GetFiles(searchFolder, "*.dwg", System.IO.SearchOption.AllDirectories))
        Catch ex As Exception
            logMessages.Add("  Error searching for DWG files: " & ex.Message)
        End Try
        
        logMessages.Add("  Found " & drawingFiles.Count & " drawing files to check")
        
        ' For first few drawings, show what they reference for debugging
        Dim debugCount As Integer = 0
        
        ' Check each drawing to see if it references any of our files
        For Each drawingPath As String In drawingFiles
            Try
                ' Open drawing silently
                Dim wasSilent As Boolean = app.SilentOperation
                app.SilentOperation = True
                
                Dim drawDoc As Document = app.Documents.Open(drawingPath, False)
                
                app.SilentOperation = wasSilent
                
                Try
                    Dim referencesOurFiles As Boolean = False
                    Dim drawingRefs As New List(Of String)
                    
                    For Each refDoc As Document In drawDoc.ReferencedDocuments
                        Dim refPath As String = System.IO.Path.GetFullPath(refDoc.FullFileName)
                        drawingRefs.Add(refPath)
                        If refFileSet.Contains(refPath) Then
                            referencesOurFiles = True
                        End If
                    Next
                    
                    ' Show debug info for first few drawings
                    If debugCount < 3 Then
                        logMessages.Add("    Checking: " & System.IO.Path.GetFileName(drawingPath))
                        For Each r As String In drawingRefs
                            Dim matched As String = If(refFileSet.Contains(r), " [MATCH]", "")
                            logMessages.Add("      -> " & System.IO.Path.GetFileName(r) & matched)
                        Next
                        debugCount += 1
                    End If
                    
                    If referencesOurFiles Then
                        drawings.Add(drawingPath)
                        logMessages.Add("  Found drawing: " & System.IO.Path.GetFileName(drawingPath))
                    End If
                Finally
                    drawDoc.Close(True)
                End Try
            Catch ex As Exception
                logMessages.Add("  Error checking drawing " & System.IO.Path.GetFileName(drawingPath) & ": " & ex.Message)
            End Try
        Next
        
        Return drawings
    End Function

    ''' <summary>
    ''' Check if a file path is inside the master folder.
    ''' </summary>
    Public Function IsInsideMasterFolder(filePath As String, masterFolder As String) As Boolean
        Dim normalizedFile As String = System.IO.Path.GetFullPath(filePath).ToLowerInvariant()
        Dim normalizedMaster As String = System.IO.Path.GetFullPath(masterFolder).ToLowerInvariant()
        
        If Not normalizedMaster.EndsWith(System.IO.Path.DirectorySeparatorChar) Then
            normalizedMaster &= System.IO.Path.DirectorySeparatorChar
        End If
        
        Return normalizedFile.StartsWith(normalizedMaster)
    End Function

    ''' <summary>
    ''' Find the master folder - the first subfolder from project root that contains the file.
    ''' Example: projectRoot="C:\Project", filePath="C:\Project\Master\Parts\file.ipt"
    '''          Returns "C:\Project\Master"
    ''' </summary>
    Public Function FindMasterFolder(projectRoot As String, filePath As String) As String
        projectRoot = System.IO.Path.GetFullPath(projectRoot).TrimEnd(System.IO.Path.DirectorySeparatorChar)
        filePath = System.IO.Path.GetFullPath(filePath)
        
        ' Check if file is under project root
        If Not filePath.StartsWith(projectRoot, StringComparison.OrdinalIgnoreCase) Then
            ' File not under project root - return immediate parent as fallback
            Return System.IO.Path.GetDirectoryName(filePath)
        End If
        
        ' Get the relative path from project root
        Dim relativePath As String = filePath.Substring(projectRoot.Length).TrimStart(System.IO.Path.DirectorySeparatorChar)
        
        ' Split into parts and get the first folder
        Dim parts() As String = relativePath.Split(System.IO.Path.DirectorySeparatorChar)
        
        If parts.Length > 1 Then
            ' Return project_root\first_folder
            Return System.IO.Path.Combine(projectRoot, parts(0))
        Else
            ' File is directly in project root - return project root
            Return projectRoot
        End If
    End Function

    ''' <summary>
    ''' Get the project root from an Inventor application.
    ''' Falls back to parent of master folder if project can't be determined.
    ''' </summary>
    Public Function GetProjectRoot(app As Inventor.Application, fallbackPath As String) As String
        Try
            Dim projectPath As String = app.DesignProjectManager.ActiveDesignProject.FullFileName
            Return System.IO.Path.GetDirectoryName(projectPath)
        Catch
            ' Fallback: walk up one level
            Return System.IO.Path.GetDirectoryName(fallbackPath)
        End Try
    End Function

    ' ============================================================================
    ' SECTION 2: Path Mapping
    ' ============================================================================

    ''' <summary>
    ''' Build a mapping from source file paths to target file paths.
    ''' When keepOriginalNames is True, all filenames stay the same (required for binary ref editing).
    ''' When False, the main assembly and files with matching names are renamed to variantName.
    ''' </summary>
    Public Function BuildCopyMap(sourceFiles As List(Of String), masterRoot As String, _
                                  targetRoot As String, variantName As String, _
                                  mainAsmPath As String, _
                                  Optional keepOriginalNames As Boolean = True) As Dictionary(Of String, String)
        Dim copyMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        
        Dim normalizedMaster As String = System.IO.Path.GetFullPath(masterRoot)
        If Not normalizedMaster.EndsWith(System.IO.Path.DirectorySeparatorChar) Then
            normalizedMaster &= System.IO.Path.DirectorySeparatorChar
        End If
        
        Dim mainAsmName As String = System.IO.Path.GetFileNameWithoutExtension(mainAsmPath)
        
        For Each sourcePath As String In sourceFiles
            If Not IsInsideMasterFolder(sourcePath, masterRoot) Then
                Continue For
            End If
            
            Dim normalizedSource As String = System.IO.Path.GetFullPath(sourcePath)
            Dim relativePath As String = normalizedSource.Substring(normalizedMaster.Length)
            
            Dim targetPath As String
            
            If keepOriginalNames Then
                ' Keep all filenames the same - only folder changes
                ' This is required for binary reference editing to work
                targetPath = System.IO.Path.Combine(targetRoot, relativePath)
            Else
                ' Rename main assembly and matching files to variant name
                Dim fileNameNoExt As String = System.IO.Path.GetFileNameWithoutExtension(sourcePath)
                Dim ext As String = System.IO.Path.GetExtension(sourcePath)
                
                If sourcePath.Equals(mainAsmPath, StringComparison.OrdinalIgnoreCase) Then
                    Dim relativeDir As String = System.IO.Path.GetDirectoryName(relativePath)
                    targetPath = System.IO.Path.Combine(targetRoot, relativeDir, variantName & ext)
                ElseIf fileNameNoExt.Equals(mainAsmName, StringComparison.OrdinalIgnoreCase) Then
                    Dim relativeDir As String = System.IO.Path.GetDirectoryName(relativePath)
                    targetPath = System.IO.Path.Combine(targetRoot, relativeDir, variantName & ext)
                Else
                    targetPath = System.IO.Path.Combine(targetRoot, relativePath)
                End If
            End If
            
            copyMap(sourcePath) = targetPath
        Next
        
        Return copyMap
    End Function

    ''' <summary>
    ''' Create target folders for all files in the copy map.
    ''' </summary>
    Public Sub CreateTargetFolders(copyMap As Dictionary(Of String, String))
        Dim folders As New HashSet(Of String)
        For Each targetPath As String In copyMap.Values
            Dim folder As String = System.IO.Path.GetDirectoryName(targetPath)
            If Not String.IsNullOrEmpty(folder) Then
                folders.Add(folder)
            End If
        Next
        For Each folder As String In folders
            If Not System.IO.Directory.Exists(folder) Then
                System.IO.Directory.CreateDirectory(folder)
            End If
        Next
    End Sub

    ' ============================================================================
    ' SECTION 3: File Copy Operations
    ' ============================================================================

    ''' <summary>
    ''' Copy all files according to the copy map (pure file system copy).
    ''' </summary>
    Public Sub CopyFiles(copyMap As Dictionary(Of String, String))
        For Each kvp As KeyValuePair(Of String, String) In copyMap
            System.IO.File.Copy(kvp.Key, kvp.Value, True)
        Next
    End Sub

    ' ============================================================================
    ' SECTION 4: Reference Replacement - Assemblies
    ' ============================================================================

    ''' <summary>
    ''' Update all references in all copied files (assemblies and parts).
    ''' This handles nested assemblies and derived parts.
    ''' </summary>
    Public Function UpdateAllReferences(app As Inventor.Application, _
                                         copyMap As Dictionary(Of String, String), _
                                         ByRef logMessages As List(Of String)) As Boolean
        Dim success As Boolean = True
        ProcessedFiles.Clear()
        
        Dim wasSilent As Boolean = app.SilentOperation
        app.SilentOperation = True
        
        Try
            ' Process all copied files
            For Each kvp As KeyValuePair(Of String, String) In copyMap
                Dim originalPath As String = kvp.Key
                Dim copiedPath As String = kvp.Value
                
                ' Skip if already processed
                If ProcessedFiles.Contains(copiedPath) Then Continue For
                
                Dim ext As String = System.IO.Path.GetExtension(copiedPath).ToLowerInvariant()
                
                If ext = ".iam" Then
                    ' Update assembly references
                    logMessages.Add("Updating assembly: " & System.IO.Path.GetFileName(copiedPath))
                    Dim result As Boolean = UpdateAssemblyReferencesRecursive(app, copiedPath, copyMap, logMessages)
                    success = success AndAlso result
                ElseIf ext = ".ipt" Then
                    ' Update derived part references
                    logMessages.Add("Updating part: " & System.IO.Path.GetFileName(copiedPath))
                    Dim result As Boolean = UpdatePartDerivedReferences(app, copiedPath, copyMap, logMessages)
                    success = success AndAlso result
                End If
            Next
        Finally
            app.SilentOperation = wasSilent
        End Try
        
        Return success
    End Function

    ''' <summary>
    ''' Update references in an assembly, including nested sub-assemblies.
    ''' </summary>
    Private Function UpdateAssemblyReferencesRecursive(app As Inventor.Application, _
                                                        copiedAsmPath As String, _
                                                        copyMap As Dictionary(Of String, String), _
                                                        ByRef logMessages As List(Of String)) As Boolean
        If ProcessedFiles.Contains(copiedAsmPath) Then Return True
        ProcessedFiles.Add(copiedAsmPath)
        
        Dim success As Boolean = True
        
        Try
            Dim asmDoc As AssemblyDocument = CType(app.Documents.Open(copiedAsmPath, False), AssemblyDocument)
            
            Try
                ' Get list of occurrences to process
                Dim occsToProcess As New List(Of ComponentOccurrence)
                CollectAllOccurrences(asmDoc.ComponentDefinition.Occurrences, occsToProcess)
                
                ' Process each occurrence
                For Each occ As ComponentOccurrence In occsToProcess
                    Try
                        Dim currentPath As String = ""
                        Try
                            currentPath = occ.Definition.Document.FullFileName
                        Catch
                            Continue For
                        End Try
                        
                        ' Check if this file was copied and needs replacement
                        If copyMap.ContainsKey(currentPath) Then
                            Dim newPath As String = copyMap(currentPath)
                            occ.Replace(newPath, False)
                        End If
                    Catch ex As Exception
                        logMessages.Add("  Warning: Could not replace occurrence - " & ex.Message)
                    End Try
                Next
                
                asmDoc.Save()
                
                ' Now recursively process any sub-assemblies that were copied
                For Each occ As ComponentOccurrence In occsToProcess
                    If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Try
                            Dim subAsmPath As String = occ.Definition.Document.FullFileName
                            ' Check if this sub-assembly is one of our copied files
                            If copyMap.Values.Contains(subAsmPath) Then
                                If Not ProcessedFiles.Contains(subAsmPath) Then
                                    UpdateAssemblyReferencesRecursive(app, subAsmPath, copyMap, logMessages)
                                End If
                            End If
                        Catch
                        End Try
                    End If
                Next
                
            Finally
                asmDoc.Close()
            End Try
        Catch ex As Exception
            logMessages.Add("  Error updating assembly: " & ex.Message)
            success = False
        End Try
        
        Return success
    End Function

    ''' <summary>
    ''' Collect all occurrences from an assembly, recursively through sub-occurrences.
    ''' </summary>
    Private Sub CollectAllOccurrences(occs As ComponentOccurrences, ByRef result As List(Of ComponentOccurrence))
        For Each occ As ComponentOccurrence In occs
            result.Add(occ)
            ' Also collect sub-occurrences for nested assemblies
            If occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Try
                    If occ.SubOccurrences IsNot Nothing Then
                        CollectAllOccurrences(occ.SubOccurrences, result)
                    End If
                Catch
                End Try
            End If
        Next
    End Sub

    ' ============================================================================
    ' SECTION 5: Reference Replacement - Derived Parts
    ' ============================================================================

    ''' <summary>
    ''' Update derived part references in a copied part file.
    ''' Uses Document.ReferencedFileDescriptors to access and update references.
    ''' </summary>
    Private Function UpdatePartDerivedReferences(app As Inventor.Application, _
                                                  copiedPartPath As String, _
                                                  copyMap As Dictionary(Of String, String), _
                                                  ByRef logMessages As List(Of String)) As Boolean
        If ProcessedFiles.Contains(copiedPartPath) Then Return True
        ProcessedFiles.Add(copiedPartPath)
        
        Dim success As Boolean = True
        Dim wasModified As Boolean = False
        
        Try
            Dim partDoc As PartDocument = CType(app.Documents.Open(copiedPartPath, False), PartDocument)
            
            Try
                ' Use ReferencedFileDescriptors (correct API for part references)
                Dim refDescriptors As ReferencedFileDescriptors = partDoc.ReferencedFileDescriptors
                
                For i As Integer = 1 To refDescriptors.Count
                    Try
                        Dim rfd As ReferencedFileDescriptor = refDescriptors.Item(i)
                        Dim refPath As String = rfd.FullFileName
                        
                        ' Check if this referenced file was copied
                        If copyMap.ContainsKey(refPath) Then
                            Dim newPath As String = copyMap(refPath)
                            
                            ' Update the reference using PutLogicalFileName
                            rfd.PutLogicalFileName(newPath)
                            wasModified = True
                            logMessages.Add("  Updated derived ref: " & System.IO.Path.GetFileName(refPath) & _
                                          " -> " & System.IO.Path.GetFileName(newPath))
                            
                            ' Recursively process the base file if it's a part
                            If newPath.ToLowerInvariant().EndsWith(".ipt") Then
                                If Not ProcessedFiles.Contains(newPath) Then
                                    UpdatePartDerivedReferences(app, newPath, copyMap, logMessages)
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        logMessages.Add("  Warning: Could not update reference - " & ex.Message)
                    End Try
                Next
                
                If wasModified Then
                    partDoc.Update()
                    partDoc.Save()
                End If
                
            Finally
                partDoc.Close()
            End Try
        Catch ex As Exception
            logMessages.Add("  Error updating part: " & ex.Message)
            success = False
        End Try
        
        Return success
    End Function

    ' ============================================================================
    ' SECTION 6: Drawing Reference Replacement
    ' ============================================================================

    ''' <summary>
    ''' Update all references in a copied drawing.
    ''' Uses Document.ReferencedFileDescriptors to access and update references.
    ''' </summary>
    Public Function UpdateDrawingReferences(app As Inventor.Application, _
                                             copiedDrawingPath As String, _
                                             copyMap As Dictionary(Of String, String)) As Boolean
        Dim success As Boolean = True
        Dim wasSilent As Boolean = app.SilentOperation
        app.SilentOperation = True
        
        Try
            Dim drawDoc As DrawingDocument = CType(app.Documents.Open(copiedDrawingPath, False), DrawingDocument)
            
            Try
                ' Use ReferencedFileDescriptors (correct API)
                Dim refDescriptors As ReferencedFileDescriptors = drawDoc.ReferencedFileDescriptors
                
                For i As Integer = 1 To refDescriptors.Count
                    Try
                        Dim rfd As ReferencedFileDescriptor = refDescriptors.Item(i)
                        Dim oldPath As String = rfd.FullFileName
                        If copyMap.ContainsKey(oldPath) Then
                            rfd.PutLogicalFileName(copyMap(oldPath))
                        End If
                    Catch
                    End Try
                Next
                drawDoc.Save()
            Finally
                drawDoc.Close()
            End Try
        Catch
            success = False
        Finally
            app.SilentOperation = wasSilent
        End Try
        
        Return success
    End Function

    ' ============================================================================
    ' SECTION 7: Parameter Application
    ' ============================================================================

    ''' <summary>
    ''' Apply parameter values to a document.
    ''' </summary>
    Public Function ApplyParameters(doc As Document, parameters As Dictionary(Of String, String)) As Boolean
        Dim success As Boolean = True
        Dim docParams As Parameters = Nothing
        
        If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            docParams = CType(doc, AssemblyDocument).ComponentDefinition.Parameters
        ElseIf doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            docParams = CType(doc, PartDocument).ComponentDefinition.Parameters
        Else
            Return False
        End If
        
        For Each kvp As KeyValuePair(Of String, String) In parameters
            If kvp.Key.StartsWith("_") Then Continue For
            Try
                Dim param As Parameter = docParams.Item(kvp.Key)
                param.Expression = kvp.Value
            Catch
                success = False
            End Try
        Next
        
        Return success
    End Function

    ' ============================================================================
    ' SECTION 8: iLogic Rule Execution
    ' ============================================================================

    ''' <summary>
    ''' Run iLogic rules matching a pattern, or all rules if pattern is "*".
    ''' Since trigger properties are not accessible via API, we run rules by name pattern.
    ''' Default pattern "Update*" runs rules starting with "Update".
    ''' </summary>
    ''' <param name="doc">Document to run rules on</param>
    ''' <param name="iLogicAuto">iLogic automation object (pass iLogicVb.Automation from calling rule)</param>
    ''' <param name="logMessages">Log messages output</param>
    ''' <param name="rulePattern">Pattern to match rule names. Use "*" for all rules, "Update*" for rules starting with Update, etc.</param>
    Public Sub RunMatchingRules(doc As Document, iLogicAuto As Object, ByRef logMessages As List(Of String), _
                                 Optional rulePattern As String = "*")
        If iLogicAuto Is Nothing Then
            logMessages.Add("Warning: iLogic automation not available")
            Exit Sub
        End If
        
        Try
            ' Get all rules in the document
            Dim rules As Object = Nothing
            Try
                rules = iLogicAuto.Rules(doc)
            Catch
                logMessages.Add("No iLogic rules found in document")
                Exit Sub
            End Try
            
            If rules Is Nothing Then Exit Sub
            
            ' Determine matching mode
            ' Patterns: "*" = all, "Update*" = starts with, "*Update" = ends with, "*Update*" = contains
            Dim matchAll As Boolean = (rulePattern = "*")
            Dim matchContains As String = ""
            Dim matchPrefix As String = ""
            Dim matchSuffix As String = ""
            
            If Not matchAll Then
                If rulePattern.StartsWith("*") AndAlso rulePattern.EndsWith("*") AndAlso rulePattern.Length > 2 Then
                    ' *text* = contains
                    matchContains = rulePattern.Substring(1, rulePattern.Length - 2)
                ElseIf rulePattern.EndsWith("*") Then
                    ' text* = starts with
                    matchPrefix = rulePattern.Substring(0, rulePattern.Length - 1)
                ElseIf rulePattern.StartsWith("*") Then
                    ' *text = ends with
                    matchSuffix = rulePattern.Substring(1)
                End If
            End If
            
            ' Run matching rules
            Dim rulesRun As Integer = 0
            
            For Each rule As Object In rules
                Try
                    Dim ruleName As String = CStr(CallByName(rule, "Name", CallType.Get))
                    
                    ' Check if rule matches pattern
                    Dim shouldRun As Boolean = False
                    
                    If matchAll Then
                        shouldRun = True
                    ElseIf Not String.IsNullOrEmpty(matchContains) Then
                        shouldRun = ruleName.IndexOf(matchContains, StringComparison.OrdinalIgnoreCase) >= 0
                    ElseIf Not String.IsNullOrEmpty(matchPrefix) Then
                        shouldRun = ruleName.StartsWith(matchPrefix, StringComparison.OrdinalIgnoreCase)
                    ElseIf Not String.IsNullOrEmpty(matchSuffix) Then
                        shouldRun = ruleName.EndsWith(matchSuffix, StringComparison.OrdinalIgnoreCase)
                    Else
                        shouldRun = ruleName.Equals(rulePattern, StringComparison.OrdinalIgnoreCase)
                    End If
                    
                    If shouldRun Then
                        logMessages.Add("Running rule: " & ruleName)
                        Try
                            iLogicAuto.RunRule(doc, ruleName)
                            rulesRun += 1
                        Catch ex As Exception
                            logMessages.Add("  Warning: Rule execution error - " & ex.Message)
                        End Try
                    End If
                Catch
                End Try
            Next
            
            logMessages.Add("  Ran " & rulesRun & " rule(s)")
            
        Catch ex As Exception
            logMessages.Add("Error enumerating rules: " & ex.Message)
        End Try
    End Sub
    
    ''' <summary>
    ''' Run all iLogic rules that have parameter change triggers.
    ''' NOTE: Trigger detection via API is not possible, so this runs ALL rules.
    ''' Consider using RunMatchingRules with a pattern instead.
    ''' </summary>
    Public Sub RunParameterChangeRules(doc As Document, iLogicAuto As Object, ByRef logMessages As List(Of String))
        ' Since we can't detect triggers via API, run all rules
        RunMatchingRules(doc, iLogicAuto, logMessages, "*")
    End Sub

    ''' <summary>
    ''' Run iLogic rules in all referenced documents recursively.
    ''' Processes BOTTOM-UP: parts first, then sub-assemblies, then main assembly.
    ''' This ensures part rules run before assembly rules that may depend on them.
    ''' </summary>
    ''' <param name="rulePattern">Pattern to match rule names. Default "*" runs all rules.</param>
    Public Sub RunRulesRecursive(app As Inventor.Application, doc As Document, _
                                  iLogicAuto As Object, ByRef logMessages As List(Of String), _
                                  ByRef processedDocs As HashSet(Of String), _
                                  Optional rulePattern As String = "*")
        If processedDocs Is Nothing Then processedDocs = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        
        Dim docPath As String = doc.FullFileName
        If processedDocs.Contains(docPath) Then Exit Sub
        processedDocs.Add(docPath)
        
        ' FIRST: Process referenced documents (bottom-up approach)
        Try
            For Each refDoc As Document In doc.ReferencedDocuments
                If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject OrElse _
                   refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    RunRulesRecursive(app, refDoc, iLogicAuto, logMessages, processedDocs, rulePattern)
                End If
            Next
        Catch
        End Try
        
        ' THEN: Run rules on this document (after all children are processed)
        logMessages.Add("  Checking rules in: " & System.IO.Path.GetFileName(docPath))
        RunMatchingRules(doc, iLogicAuto, logMessages, rulePattern)
    End Sub
    
    ''' <summary>
    ''' Run all iLogic rules in all referenced documents recursively.
    ''' </summary>
    Public Sub RunParameterChangeRulesRecursive(app As Inventor.Application, doc As Document, _
                                                 iLogicAuto As Object, ByRef logMessages As List(Of String), _
                                                 ByRef processedDocs As HashSet(Of String))
        RunRulesRecursive(app, doc, iLogicAuto, logMessages, processedDocs, "*")
    End Sub

    ' ============================================================================
    ' SECTION 9: iProperty Updates
    ' ============================================================================

    ''' <summary>
    ''' Update iProperties on a document.
    ''' </summary>
    Public Sub UpdateiProperties(doc As Document, partNumber As String, variantName As String, _
                                  Optional description As String = "")
        Try
            Dim propSets As PropertySets = doc.PropertySets
            Dim designProps As PropertySet = propSets.Item("Design Tracking Properties")
            
            If Not String.IsNullOrEmpty(partNumber) Then
                Try
                    designProps.Item("Part Number").Value = partNumber
                Catch
                End Try
            End If
            
            If Not String.IsNullOrEmpty(description) Then
                Try
                    designProps.Item("Description").Value = description
                Catch
                End Try
            End If
            
            Try
                Dim customProps As PropertySet = propSets.Item("Inventor User Defined Properties")
                Try
                    customProps.Item("VariantName").Value = variantName
                Catch
                    customProps.Add(variantName, "VariantName")
                End Try
            Catch
            End Try
        Catch
        End Try
    End Sub

    ' ============================================================================
    ' SECTION 10: Main Release Function
    ' ============================================================================

    ''' <summary>
    ''' Release a single variant of an assembly.
    ''' </summary>
    Public Function ReleaseVariant(app As Inventor.Application, _
                                    masterAsmPath As String, _
                                    variantName As String, _
                                    partNumber As String, _
                                    parameters As Dictionary(Of String, String), _
                                    releaseRoot As String, _
                                    includeDrawings As Boolean, _
                                    ByRef logMessages As List(Of String), _
                                    Optional iLogicAuto As Object = Nothing) As String
        
        If logMessages Is Nothing Then logMessages = New List(Of String)
        
        Try
            logMessages.Add("Starting release of variant: " & variantName)
            
            ' Open the master assembly to get dependencies
            logMessages.Add("Opening master assembly...")
            Dim masterDoc As AssemblyDocument = CType(app.Documents.Open(masterAsmPath, False), AssemblyDocument)
            
            ' Get project root and find master folder (first subfolder from project root)
            Dim projectRoot As String = GetProjectRoot(app, masterAsmPath)
            Dim masterFolder As String = FindMasterFolder(projectRoot, masterAsmPath)
            
            logMessages.Add("Project root: " & projectRoot)
            logMessages.Add("Master folder: " & masterFolder)
            
            ' Get all referenced files
            logMessages.Add("Collecting dependency tree...")
            Dim allFiles As List(Of String) = GetAllReferencedFiles(masterDoc)
            logMessages.Add("Found " & allFiles.Count & " files in dependency tree")
            
            ' Find drawings if requested
            Dim drawingFiles As New List(Of String)
            If includeDrawings Then
                logMessages.Add("Searching for drawings...")
                
                ' Search in master folder (includes all subfolders)
                drawingFiles = FindAllDrawings(app, masterFolder, allFiles, logMessages)
                
                logMessages.Add("Total drawings found: " & drawingFiles.Count)
                
                For Each dwg As String In drawingFiles
                    If Not allFiles.Contains(dwg) Then
                        allFiles.Add(dwg)
                    End If
                Next
            End If
            
            ' Close master document
            masterDoc.Close(True)
            
            ' Calculate length-matched variant folder for binary reference updating
            ' The variant folder path length must match the master folder path length
            logMessages.Add("Calculating length-matched variant folder...")
            Dim variantFolder As String = CalculateLengthMatchedVariantFolder( _
                masterFolder, releaseRoot, variantName, logMessages)
            
            ' Check if folder exists and increment number suffix
            ' Initial folder ends with _1, so we start at 2 if it exists
            Dim folderIndex As Integer = 1
            Dim baseFolderName As String = variantFolder
            While System.IO.Directory.Exists(variantFolder)
                folderIndex += 1
                ' Adjust the folder name while maintaining length
                variantFolder = AdjustFolderNameWithIndex(baseFolderName, folderIndex, masterFolder.Length)
            End While
            
            logMessages.Add("Creating release folder: " & variantFolder)
            logMessages.Add("  Master folder length: " & masterFolder.Length)
            logMessages.Add("  Variant folder length: " & variantFolder.Length)
            
            ' Build copy map
            Dim copyMap As Dictionary(Of String, String) = BuildCopyMap(allFiles, masterFolder, variantFolder, variantName, masterAsmPath)
            logMessages.Add("Built copy map with " & copyMap.Count & " files")
            
            ' Create target folders
            CreateTargetFolders(copyMap)
            
            ' PHASE 1: Copy all files
            logMessages.Add("Phase 1: Copying files...")
            CopyFiles(copyMap)
            
            ' Get the path to the copied assembly
            Dim copiedAsmPath As String = copyMap(masterAsmPath)
            
            ' PHASE 2: Update derived part references using BINARY editing
            ' This must be done BEFORE opening files, while they are closed
            logMessages.Add("Phase 2: Updating derived part references (binary)...")
            Dim binarySuccess As Boolean = UpdateDerivedPartReferencesBinary(copyMap, logMessages)
            If Not binarySuccess Then
                logMessages.Add("Warning: Some derived part references could not be updated")
            End If
            
            ' PHASE 2b: Update assembly references using API (opens files)
            logMessages.Add("Phase 2b: Updating assembly references...")
            Dim refSuccess As Boolean = UpdateAllReferences(app, copyMap, logMessages)
            If Not refSuccess Then
                logMessages.Add("Warning: Some assembly references could not be updated")
            End If
            
            ' PHASE 3: Update drawing references using binary editing
            ' (PutLogicalFileName API doesn't work in iLogic)
            If includeDrawings Then
                logMessages.Add("Phase 3: Updating drawing references (binary)...")
                For Each drawingPath As String In drawingFiles
                    If copyMap.ContainsKey(drawingPath) Then
                        Dim copiedDrawingPath As String = copyMap(drawingPath)
                        logMessages.Add("  Updating: " & System.IO.Path.GetFileName(copiedDrawingPath))
                        UpdateSingleFileBinary(copiedDrawingPath, copyMap, logMessages)
                    End If
                Next
            End If
            
            ' PHASE 4: Apply parameters and run iLogic rules
            logMessages.Add("Phase 4: Applying parameters and running rules...")
            Dim wasSilent As Boolean = app.SilentOperation
            app.SilentOperation = True
            
            Try
                Dim releasedDoc As AssemblyDocument = CType(app.Documents.Open(copiedAsmPath, False), AssemblyDocument)
                Try
                    ' Apply parameters
                    If parameters IsNot Nothing AndAlso parameters.Count > 0 Then
                        logMessages.Add("  Applying parameters...")
                        ApplyParameters(releasedDoc, parameters)
                    End If
                    
                    ' Update iProperties
                    UpdateiProperties(releasedDoc, partNumber, variantName)
                    
                    ' Update the document to recalculate
                    logMessages.Add("  Updating document...")
                    releasedDoc.Update()
                    
                    ' Run iLogic rules - BOTTOM-UP order (parts first, then assembly)
                    If iLogicAuto IsNot Nothing Then
                        ' PASS 1: Run rules on all parts first (bottom-up)
                        logMessages.Add("  Pass 1: Running iLogic rules (parts first)...")
                        Dim processedDocs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                        RunRulesRecursive(app, releasedDoc, iLogicAuto, logMessages, processedDocs, "*Update*")
                        
                        ' Update and SAVE all parts before assembly uses their values
                        logMessages.Add("  Pass 1: Updating and saving all parts...")
                        For Each refDoc As Document In releasedDoc.AllReferencedDocuments
                            Try
                                refDoc.Update()
                                refDoc.Save()
                            Catch
                            End Try
                        Next
                        
                        ' Now update the main assembly (which reads saved part values)
                        logMessages.Add("  Pass 1: Updating main assembly...")
                        releasedDoc.Update()
                        
                        ' PASS 2: Run rules again to catch any cascading changes
                        logMessages.Add("  Pass 2: Running iLogic rules again (parts first)...")
                        processedDocs.Clear()
                        RunRulesRecursive(app, releasedDoc, iLogicAuto, logMessages, processedDocs, "*Update*")

                        
                        ' Update and save all parts again
                        logMessages.Add("  Pass 2: Updating and saving all parts...")
                        For Each refDoc As Document In releasedDoc.AllReferencedDocuments
                            Try
                                refDoc.Update()
                                refDoc.Save()
                            Catch
                            End Try
                        Next
                        
                        ' Update main assembly again
                        releasedDoc.Update()
                    End If
                    
                    ' Final rebuild to ensure all updates are applied
                    logMessages.Add("  Final rebuild...")
                    Try
                        releasedDoc.Update2(True) ' Force full rebuild
                    Catch
                        releasedDoc.Update()
                    End Try
                    
                    ' Save all referenced documents first
                    logMessages.Add("  Saving all documents...")
                    Dim savedCount As Integer = 0
                    For Each refDoc As Document In releasedDoc.AllReferencedDocuments
                        Try
                            refDoc.Save()
                            savedCount += 1
                        Catch
                        End Try
                    Next
                    logMessages.Add("    Saved " & savedCount & " referenced documents")
                    
                    ' Save the main assembly
                    releasedDoc.Save()
                    logMessages.Add("    Saved main assembly")
 
                    logMessages.Add("Parameters applied and saved")
                Finally
                    releasedDoc.Close()
                End Try
            Finally
                app.SilentOperation = wasSilent
            End Try
            
            logMessages.Add("Release complete: " & copiedAsmPath)
            Return copiedAsmPath
            
        Catch ex As Exception
            logMessages.Add("ERROR: " & ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Write log messages to a file in the release folder.
    ''' </summary>
    Public Sub WriteLogFile(releaseFolder As String, variantName As String, logMessages As List(Of String))
        Try
            Dim logPath As String = System.IO.Path.Combine(releaseFolder, variantName & "_release.log")
            Dim logContent As String = String.Join(vbCrLf, logMessages.ToArray())
            System.IO.File.WriteAllText(logPath, logContent)
        Catch
        End Try
    End Sub

    ' ============================================================================
    ' SECTION 11: Length-Matched Folder Calculation
    ' ============================================================================

    ''' <summary>
    ''' Calculate a variant folder path that EXACTLY matches the master folder length.
    ''' This is required for binary reference updating to work.
    ''' 
    ''' The path is auto-generated to ensure exact length match:
    ''' - Uses releaseRoot (e.g., "r") subfolder
    ''' - Folder name is built from variant name + number to match length
    ''' </summary>
    Private Function CalculateLengthMatchedVariantFolder(masterFolder As String, _
                                                          releaseRoot As String, _
                                                          variantName As String, _
                                                          ByRef logMessages As List(Of String)) As String
        ' Normalize master path
        masterFolder = System.IO.Path.GetFullPath(masterFolder).TrimEnd(System.IO.Path.DirectorySeparatorChar)
        Dim masterLength As Integer = masterFolder.Length
        
        logMessages.Add("  Master folder: " & masterFolder)
        logMessages.Add("  Master length: " & masterLength)
        
        ' Normalize release root (e.g., "C:\...\ScriptTesting\r")
        releaseRoot = System.IO.Path.GetFullPath(releaseRoot).TrimEnd(System.IO.Path.DirectorySeparatorChar)
        
        ' Create release root if it doesn't exist
        If Not System.IO.Directory.Exists(releaseRoot) Then
            System.IO.Directory.CreateDirectory(releaseRoot)
        End If
        
        ' We need: releaseRoot + "\" + generatedFolderName = masterLength
        ' So: generatedFolderName.Length = masterLength - releaseRoot.Length - 1
        Dim requiredFolderNameLength As Integer = masterLength - releaseRoot.Length - 1
        
        logMessages.Add("  Release root: " & releaseRoot)
        logMessages.Add("  Required folder name length: " & requiredFolderNameLength)
        
        If requiredFolderNameLength <= 0 Then
            logMessages.Add("  WARNING: Release path too long for length matching!")
            logMessages.Add("  Using variant name directly (binary update may fail)")
            Return System.IO.Path.Combine(releaseRoot, variantName & "_1")
        End If
        
        ' Build the folder name to exactly match the required length
        ' Format: variantName + number (always include number)
        Dim folderName As String = GenerateLengthMatchedFolderName(variantName, requiredFolderNameLength)
        
        logMessages.Add("  Generated folder name: " & folderName & " (length: " & folderName.Length & ")")
        
        Dim result As String = System.IO.Path.Combine(releaseRoot, folderName)
        logMessages.Add("  Variant folder: " & result)
        logMessages.Add("  Variant length: " & result.Length)
        
        If result.Length <> masterLength Then
            logMessages.Add("  WARNING: Length mismatch! Binary update may fail.")
        Else
            logMessages.Add("  Length match: OK")
        End If
        
        Return result
    End Function

    ''' <summary>
    ''' Generate a folder name of exactly the specified length.
    ''' Always ends with _1 (first release number).
    ''' Uses variant name as base, pads with underscores or truncates as needed.
    ''' </summary>
    Private Function GenerateLengthMatchedFolderName(variantName As String, targetLength As Integer) As String
        If targetLength <= 0 Then
            Return "1"
        End If
        
        ' Clean variant name of invalid characters
        Dim cleanName As String = variantName
        For Each c As Char In System.IO.Path.GetInvalidFileNameChars()
            cleanName = cleanName.Replace(c, "_"c)
        Next
        
        ' Always end with _1 (release number)
        Dim suffix As String = "_1"
        Dim availableForName As Integer = targetLength - suffix.Length
        
        If availableForName <= 0 Then
            ' Very short target - just use number
            Return "1".PadLeft(targetLength, "0"c)
        End If
        
        Dim baseName As String
        If cleanName.Length <= availableForName Then
            ' Name fits - pad with underscores if needed
            Dim padding As Integer = availableForName - cleanName.Length
            baseName = cleanName & New String("_"c, padding)
        Else
            ' Truncate name to fit
            baseName = cleanName.Substring(0, availableForName)
        End If
        
        Return baseName & suffix
    End Function

    ''' <summary>
    ''' Adjust folder name with index while maintaining target length.
    ''' </summary>
    Private Function AdjustFolderNameWithIndex(baseFolderPath As String, index As Integer, targetLength As Integer) As String
        Dim parentDir As String = System.IO.Path.GetDirectoryName(baseFolderPath)
        Dim baseName As String = System.IO.Path.GetFileName(baseFolderPath)
        
        Dim indexStr As String = index.ToString()
        Dim requiredBaseLength As Integer = targetLength - parentDir.Length - 1
        
        If requiredBaseLength <= 0 Then
            ' Can't fit in target length - use minimal versioning
            Return baseFolderPath & indexStr
        End If
        
        If requiredBaseLength <= indexStr.Length Then
            ' Just use the index
            Return System.IO.Path.Combine(parentDir, indexStr.PadLeft(requiredBaseLength, "0"c))
        End If
        
        ' Replace last characters with index
        Dim newName As String = baseName.Substring(0, requiredBaseLength - indexStr.Length) & indexStr
        Return System.IO.Path.Combine(parentDir, newName)
    End Function

    ' ============================================================================
    ' SECTION 12: Binary Reference Updating
    ' ============================================================================

    ''' <summary>
    ''' Update derived part references using binary file editing.
    ''' Files must be CLOSED before calling this function.
    ''' </summary>
    Private Function UpdateDerivedPartReferencesBinary(copyMap As Dictionary(Of String, String), _
                                                        ByRef logMessages As List(Of String)) As Boolean
        Dim success As Boolean = True
        
        ' Process all copied IPT files (parts may have derived references)
        For Each kvp As KeyValuePair(Of String, String) In copyMap
            Dim originalPath As String = kvp.Key
            Dim copiedPath As String = kvp.Value
            
            ' Only process part files
            If Not copiedPath.ToLowerInvariant().EndsWith(".ipt") Then
                Continue For
            End If
            
            ' Check if this part file references any files that were also copied
            Dim refsUpdated As Boolean = UpdateSingleFileBinary(copiedPath, copyMap, logMessages)
            If Not refsUpdated Then
                ' Not necessarily an error - file might not have derived refs
            End If
        Next
        
        Return success
    End Function

    ''' <summary>
    ''' Update references in a single file using binary editing.
    ''' </summary>
    Private Function UpdateSingleFileBinary(filePath As String, _
                                             copyMap As Dictionary(Of String, String), _
                                             ByRef logMessages As List(Of String)) As Boolean
        If Not System.IO.File.Exists(filePath) Then
            Return False
        End If
        
        ' Read file bytes
        Dim fileBytes As Byte() = Nothing
        Try
            fileBytes = System.IO.File.ReadAllBytes(filePath)
        Catch ex As Exception
            logMessages.Add("  ERROR reading: " & System.IO.Path.GetFileName(filePath) & " - " & ex.Message)
            Return False
        End Try
        
        Dim modified As Boolean = False
        Dim fileName As String = System.IO.Path.GetFileName(filePath)
        
        ' Try to replace each original path with its copy
        For Each kvp As KeyValuePair(Of String, String) In copyMap
            Dim oldPath As String = kvp.Key
            Dim newPath As String = kvp.Value
            
            ' Skip if same or if new is longer
            If oldPath.Equals(newPath, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If
            
            If newPath.Length > oldPath.Length Then
                ' Can't replace - new path is longer
                Continue For
            End If
            
            ' Search for old path in file (Unicode)
            Dim oldBytes As Byte() = System.Text.Encoding.Unicode.GetBytes(oldPath)
            Dim pos As Integer = IndexOfBytesInArray(fileBytes, oldBytes, 0)
            
            If pos < 0 Then
                ' Path not found in this file
                Continue For
            End If
            
            ' Prepare new bytes (padded with nulls if shorter)
            Dim newBytes As Byte() = Nothing
            If newPath.Length = oldPath.Length Then
                newBytes = System.Text.Encoding.Unicode.GetBytes(newPath)
            Else
                ' Pad with null characters
                Dim paddedNew As String = newPath & New String(ChrW(0), oldPath.Length - newPath.Length)
                newBytes = System.Text.Encoding.Unicode.GetBytes(paddedNew)
            End If
            
            ' Replace all occurrences
            Dim replaceCount As Integer = 0
            Do While pos >= 0
                Array.Copy(newBytes, 0, fileBytes, pos, newBytes.Length)
                replaceCount += 1
                pos = IndexOfBytesInArray(fileBytes, oldBytes, pos + oldBytes.Length)
            Loop
            
            If replaceCount > 0 Then
                modified = True
                logMessages.Add("  " & fileName & ": " & System.IO.Path.GetFileName(oldPath) & _
                              " -> " & System.IO.Path.GetFileName(newPath) & " (" & replaceCount & "x)")
            End If
        Next
        
        ' Write modified file
        If modified Then
            Try
                ' Create backup
                System.IO.File.Copy(filePath, filePath & ".backup", True)
                ' Write
                System.IO.File.WriteAllBytes(filePath, fileBytes)
            Catch ex As Exception
                logMessages.Add("  ERROR writing: " & fileName & " - " & ex.Message)
                Return False
            End Try
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' Find byte sequence in array.
    ''' </summary>
    Private Function IndexOfBytesInArray(source As Byte(), pattern As Byte(), startIndex As Integer) As Integer
        If source Is Nothing OrElse pattern Is Nothing Then Return -1
        If pattern.Length > source.Length - startIndex Then Return -1
        
        For i As Integer = startIndex To source.Length - pattern.Length
            Dim found As Boolean = True
            For j As Integer = 0 To pattern.Length - 1
                If source(i + j) <> pattern(j) Then
                    found = False
                    Exit For
                End If
            Next
            If found Then Return i
        Next
        
        Return -1
    End Function

End Module
