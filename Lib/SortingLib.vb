' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' SortingLib - Assembly Component Sorting by Material
' 
' Categorizes assembly components into browser folders based on material name
' patterns, then creates model states (suppression) and design views (visibility)
' for different folder combinations.
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/SortingLib.vb"
'
' Dependencies: UtilsLib (for logging and browser/occurrence utilities)
' ============================================================================

Imports Inventor
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Public Module SortingLib

    ' ============================================================================
    ' SECTION 1: Public Entry Point
    ' ============================================================================

    ''' <summary>
    ''' Main entry point - sorts occurrences into folders and creates model states.
    ''' </summary>
    ''' <param name="asmDoc">The assembly document to process</param>
    ''' <param name="folderPatterns">Dictionary mapping folder name to list of regex patterns for material matching</param>
    ''' <param name="stateDefinitions">Dictionary mapping state name to list of included folder names (Nothing = all enabled)</param>
    Public Sub Run( _
        asmDoc As AssemblyDocument, _
        folderPatterns As Dictionary(Of String, List(Of String)), _
        stateDefinitions As Dictionary(Of String, List(Of String)))
        
        Run(asmDoc, folderPatterns, stateDefinitions, Nothing)
    End Sub
    
    ''' <summary>
    ''' Main entry point - sorts occurrences into folders and creates model states and design views.
    ''' </summary>
    ''' <param name="asmDoc">The assembly document to process</param>
    ''' <param name="folderPatterns">Dictionary mapping folder name to list of regex patterns for material matching</param>
    ''' <param name="stateDefinitions">Dictionary mapping state name to list of included folder names (Nothing = all enabled)</param>
    ''' <param name="viewDefinitions">Dictionary mapping view name to list of VISIBLE folder names (Nothing = all visible)</param>
    Public Sub Run( _
        asmDoc As AssemblyDocument, _
        folderPatterns As Dictionary(Of String, List(Of String)), _
        stateDefinitions As Dictionary(Of String, List(Of String)), _
        viewDefinitions As Dictionary(Of String, List(Of String)))
        
        UtilsLib.LogInfo("SortingLib: Starting for " & asmDoc.DisplayName)
        
        Dim oPane As BrowserPane = asmDoc.BrowserPanes.Item("Model")
        
        ' Step 1: Assign occurrences to folders based on material (returns created folder names)
        Dim createdFolders As List(Of String) = ApplyFolders(asmDoc, oPane, folderPatterns)
        
        ' Step 2: Create/update model states (only for folders that have parts)
        ApplyModelStates(asmDoc, oPane, createdFolders, stateDefinitions)
        
        ' Step 3: Create/update design views (visibility)
        If viewDefinitions IsNot Nothing AndAlso viewDefinitions.Count > 0 Then
            ApplyDesignViews(asmDoc, oPane, createdFolders, viewDefinitions)
        End If
        
        ' Step 4: Restore defaults
        RestoreDefaults(asmDoc)
        
        UtilsLib.LogInfo("SortingLib: Completed")
    End Sub

    ' ============================================================================
    ' SECTION 2: Folder Assignment
    ' ============================================================================

    ''' <summary>
    ''' Creates folders and assigns occurrences based on material patterns.
    ''' Only creates folders that will have parts assigned to them.
    ''' Uses AddBrowserFolder with ObjectCollection to support patterns.
    ''' Returns the list of folder names that were created (have parts).
    ''' </summary>
    Private Function ApplyFolders( _
        asmDoc As AssemblyDocument, _
        oPane As BrowserPane, _
        folderPatterns As Dictionary(Of String, List(Of String))) As List(Of String)
        
        UtilsLib.LogInfo("SortingLib: Assigning occurrences to folders...")
        
        Dim app As Inventor.Application = asmDoc.Parent
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Collect browser nodes by target folder
        ' Key: folder name, Value: list of browser nodes to add
        Dim folderNodes As New Dictionary(Of String, List(Of BrowserNode))
        Dim processedNodes As New HashSet(Of String)
        Dim processedPatterns As New HashSet(Of String)
        
        Dim assignedCount As Integer = 0
        Dim skippedCount As Integer = 0
        Dim patternCount As Integer = 0
        Dim errorCount As Integer = 0
        
        UtilsLib.LogInfo("SortingLib: Processing " & asmDef.Occurrences.Count & " occurrences...")
        
        ' First pass: collect standalone occurrences (skip pattern elements)
        For Each occ As ComponentOccurrence In asmDef.Occurrences
            Dim materialName As String = UtilsLib.GetOccurrenceMaterial(occ)
            
            If String.IsNullOrEmpty(materialName) Then
                skippedCount += 1
                Continue For
            End If
            
            Dim targetFolder As String = GetTargetFolder(materialName, folderPatterns)
            
            If String.IsNullOrEmpty(targetFolder) Then
                UtilsLib.LogInfo("SortingLib: No folder match for '" & occ.Name & "' (material: " & materialName & ")")
                skippedCount += 1
                Continue For
            End If
            
            Try
                ' Check if this occurrence is part of a pattern
                Dim isPatternElement As Boolean = False
                Try
                    isPatternElement = occ.IsPatternElement
                Catch
                End Try
                
                ' Handle pattern elements - find parent pattern via browser tree
                If isPatternElement Then
                    Try
                        ' Find the pattern's browser node by walking up the tree
                        Dim occNode As BrowserNode = oPane.GetBrowserNodeFromObject(occ)
                        Dim patternNode As BrowserNode = FindParentPatternNode(occNode)
                        
                        If patternNode IsNot Nothing Then
                            Dim patternKey As String = patternNode.FullPath
                            
                            If Not processedPatterns.Contains(patternKey) Then
                                processedPatterns.Add(patternKey)
                                
                                ' Check if pattern already in target folder
                                Dim patternCurrentFolder As String = UtilsLib.GetNodeFolder(patternNode)
                                Dim patternAlreadyInPlace As Boolean = (patternCurrentFolder = targetFolder)
                                
                                ' Add pattern to collection (even if already there, for folder recreation)
                                If Not folderNodes.ContainsKey(targetFolder) Then
                                    folderNodes.Add(targetFolder, New List(Of BrowserNode))
                                End If
                                folderNodes(targetFolder).Add(patternNode)
                                
                                If Not patternAlreadyInPlace Then
                                    UtilsLib.LogInfo("SortingLib: Queuing pattern '" & patternNode.BrowserNodeDefinition.Label & "' for '" & targetFolder & "'")
                                    patternCount += 1
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        UtilsLib.LogWarn("SortingLib: Error processing pattern element '" & occ.Name & "': " & ex.Message)
                    End Try
                    Continue For
                End If
                
                ' Standalone occurrence - get its movable node
                Dim oNode As BrowserNode = oPane.GetBrowserNodeFromObject(occ)
                Dim movableNode As BrowserNode = UtilsLib.GetMovableParentNode(oPane, oNode)
                
                Dim nodeKey As String = movableNode.FullPath
                
                If processedNodes.Contains(nodeKey) Then Continue For
                processedNodes.Add(nodeKey)
                
                ' Check if already in target folder
                Dim currentFolder As String = UtilsLib.GetNodeFolder(movableNode)
                Dim alreadyInPlace As Boolean = (currentFolder = targetFolder)
                
                ' Add to collection for this folder (even if already there, for folder recreation)
                If Not folderNodes.ContainsKey(targetFolder) Then
                    folderNodes.Add(targetFolder, New List(Of BrowserNode))
                End If
                folderNodes(targetFolder).Add(movableNode)
                
                If Not alreadyInPlace Then
                    UtilsLib.LogInfo("SortingLib: Queuing '" & movableNode.BrowserNodeDefinition.Label & "' for '" & targetFolder & "'")
                    assignedCount += 1
                End If
                
            Catch ex As Exception
                UtilsLib.LogWarn("SortingLib: Error processing '" & occ.Name & "': " & ex.Message)
                errorCount += 1
            End Try
        Next
        
        ' Second pass: collect any remaining patterns via OccurrencePatterns collection
        ' (Rectangular/Circular patterns that weren't detected via occurrence iteration)
        Try
            Dim occPatterns As OccurrencePatterns = asmDef.OccurrencePatterns
            If occPatterns.Count > 0 Then
                UtilsLib.LogInfo("SortingLib: Checking " & occPatterns.Count & " OccurrencePatterns...")
            End If
            
            For Each pattern As OccurrencePattern In occPatterns
                Try
                    ' Get material from first occurrence in the pattern
                    Dim firstOcc As ComponentOccurrence = Nothing
                    Try
                        firstOcc = pattern.OccurrencePatternElements.Item(1).Occurrences.Item(1)
                    Catch
                        Continue For
                    End Try
                    
                    If firstOcc Is Nothing Then Continue For
                    
                    Dim materialName As String = UtilsLib.GetOccurrenceMaterial(firstOcc)
                    If String.IsNullOrEmpty(materialName) Then Continue For
                    
                    Dim targetFolder As String = GetTargetFolder(materialName, folderPatterns)
                    If String.IsNullOrEmpty(targetFolder) Then
                        UtilsLib.LogInfo("SortingLib: No folder match for pattern '" & pattern.Name & "' (material: " & materialName & ")")
                        Continue For
                    End If
                    
                    ' Check if pattern already processed
                    If processedPatterns.Contains(pattern.Name) Then Continue For
                    processedPatterns.Add(pattern.Name)
                    
                    ' Get pattern browser node
                    Dim patternNode As BrowserNode = oPane.GetBrowserNodeFromObject(pattern)
                    If patternNode Is Nothing Then
                        UtilsLib.LogWarn("SortingLib: Could not get browser node for pattern '" & pattern.Name & "'")
                        Continue For
                    End If
                    
                    ' Check if already in target folder
                    Dim currentFolder As String = UtilsLib.GetNodeFolder(patternNode)
                    If currentFolder = targetFolder Then
                        UtilsLib.LogInfo("SortingLib: Pattern '" & pattern.Name & "' already in '" & targetFolder & "'")
                        Continue For
                    End If
                    
                    ' Add to collection for this folder
                    If Not folderNodes.ContainsKey(targetFolder) Then
                        folderNodes.Add(targetFolder, New List(Of BrowserNode))
                    End If
                    folderNodes(targetFolder).Add(patternNode)
                    
                    UtilsLib.LogInfo("SortingLib: Queuing pattern '" & pattern.Name & "' for '" & targetFolder & "'")
                    patternCount += 1
                    
                Catch ex As Exception
                    UtilsLib.LogWarn("SortingLib: Error processing pattern: " & ex.Message)
                    errorCount += 1
                End Try
            Next
        Catch ex As Exception
            UtilsLib.LogWarn("SortingLib: Error accessing OccurrencePatterns: " & ex.Message)
        End Try
        
        ' Third pass: create folders or add items to existing folders
        Dim createdFolders As New List(Of String)
        
        For Each kvp As KeyValuePair(Of String, List(Of BrowserNode)) In folderNodes
            Dim folderName As String = kvp.Key
            Dim nodesToAdd As List(Of BrowserNode) = kvp.Value
            
            If nodesToAdd.Count = 0 Then Continue For
            
            Try
                ' Check if folder already exists
                Dim existingFolder As BrowserFolder = UtilsLib.FindFolder(oPane, folderName)
                
                ' Separate items into: already in place vs need to be moved
                Dim itemsToMove As New List(Of BrowserNode)
                For Each node As BrowserNode In nodesToAdd
                    Dim nodeFolder As String = UtilsLib.GetNodeFolder(node)
                    If nodeFolder <> folderName Then
                        itemsToMove.Add(node)
                    End If
                Next
                
                ' If all items are already in place, nothing to do
                If itemsToMove.Count = 0 Then
                    UtilsLib.LogInfo("SortingLib: Folder '" & folderName & "' already has all " & nodesToAdd.Count & " items")
                    createdFolders.Add(folderName)
                    Continue For
                End If
                
                If existingFolder IsNot Nothing Then
                    ' Folder exists - try to add new items one by one
                    Dim failedNodes As New List(Of BrowserNode)
                    
                    For Each node As BrowserNode In itemsToMove
                        Try
                            existingFolder.Add(node)
                        Catch
                            ' BrowserFolder.Add fails for patterns - collect for retry
                            failedNodes.Add(node)
                        End Try
                    Next
                    
                    ' If some nodes failed (likely patterns), recreate folder with fresh references
                    If failedNodes.Count > 0 Then
                        UtilsLib.LogInfo("SortingLib: Recreating folder '" & folderName & "' to include " & failedNodes.Count & " pattern(s)...")
                        
                        Try
                            ' Remember all items that should be in the folder
                            Dim allOccNames As New List(Of String)
                            For Each node As BrowserNode In nodesToAdd
                                Try
                                    Dim nativeObj As Object = node.NativeObject
                                    If TypeOf nativeObj Is ComponentOccurrence Then
                                        allOccNames.Add(CType(nativeObj, ComponentOccurrence).Name)
                                    End If
                                Catch
                                End Try
                            Next
                            
                            ' Delete the folder
                            existingFolder.Delete()
                            
                            ' Get fresh browser node references from occurrences and patterns
                            Dim freshCollection As ObjectCollection = app.TransientObjects.CreateObjectCollection()
                            
                            ' Add occurrences by name lookup
                            For Each occName As String In allOccNames
                                Try
                                    Dim occ As ComponentOccurrence = asmDef.Occurrences.ItemByName(occName)
                                    Dim freshNode As BrowserNode = oPane.GetBrowserNodeFromObject(occ)
                                    Dim movableNode As BrowserNode = UtilsLib.GetMovableParentNode(oPane, freshNode)
                                    freshCollection.Add(movableNode)
                                Catch
                                End Try
                            Next
                            
                            ' Add patterns via FindParentPatternNode for pattern elements
                            For Each node As BrowserNode In failedNodes
                                Try
                                    ' For patterns, get fresh reference via the occurrence
                                    Dim nativeObj As Object = Nothing
                                    Try : nativeObj = node.NativeObject : Catch : End Try
                                    
                                    If nativeObj Is Nothing Then
                                        ' Mirror pattern - find via browser tree
                                        ' The pattern node itself should still be valid after folder delete
                                        freshCollection.Add(node)
                                    ElseIf TypeOf nativeObj Is OccurrencePattern Then
                                        Dim freshNode As BrowserNode = oPane.GetBrowserNodeFromObject(nativeObj)
                                        freshCollection.Add(freshNode)
                                    End If
                                Catch
                                End Try
                            Next
                            
                            ' Create new folder with all items
                            Dim newFolder As BrowserFolder = oPane.AddBrowserFolder(folderName, freshCollection)
                            newFolder.AllowReorder = True
                            newFolder.AllowDelete = True
                            UtilsLib.LogInfo("SortingLib: Recreated folder '" & folderName & "'")
                        Catch ex As Exception
                            UtilsLib.LogWarn("SortingLib: Could not recreate folder '" & folderName & "': " & ex.Message)
                            errorCount += 1
                        End Try
                    Else
                        UtilsLib.LogInfo("SortingLib: Added " & itemsToMove.Count & " items to folder '" & folderName & "'")
                    End If
                    
                    createdFolders.Add(folderName)
                Else
                    ' New folder - create with all items using AddBrowserFolder (supports patterns)
                    UtilsLib.LogInfo("SortingLib: Creating new folder '" & folderName & "' with " & itemsToMove.Count & " items...")
                    
                    Dim nodeCollection As ObjectCollection = app.TransientObjects.CreateObjectCollection()
                    For Each node As BrowserNode In itemsToMove
                        nodeCollection.Add(node)
                    Next
                    
                    Dim newFolder As BrowserFolder = oPane.AddBrowserFolder(folderName, nodeCollection)
                    newFolder.AllowReorder = True
                    newFolder.AllowDelete = True
                    
                    createdFolders.Add(folderName)
                    UtilsLib.LogInfo("SortingLib: Created folder '" & folderName & "'")
                End If
                
            Catch ex As Exception
                UtilsLib.LogWarn("SortingLib: Error with folder '" & folderName & "': " & ex.Message)
                errorCount += 1
            End Try
        Next
        
        ' Also track folders that already exist and have correct items
        For Each folderName As String In folderPatterns.Keys
            If Not createdFolders.Contains(folderName) Then
                Dim existingFolder As BrowserFolder = UtilsLib.FindFolder(oPane, folderName)
                If existingFolder IsNot Nothing Then
                    createdFolders.Add(folderName)
                End If
            End If
        Next
        
        UtilsLib.LogInfo("SortingLib: Moved " & assignedCount & " occurrence(s), " & patternCount & " pattern(s), skipped " & skippedCount & " (no material), " & errorCount & " errors")
        
        Return createdFolders
    End Function
    
    ''' <summary>
    ''' Finds the parent pattern browser node for a pattern element occurrence.
    ''' Works for Mirror, Rectangular, and Circular component patterns.
    ''' </summary>
    Private Function FindParentPatternNode(occNode As BrowserNode) As BrowserNode
        If occNode Is Nothing Then Return Nothing
        
        Try
            ' Walk up the browser tree looking for a pattern node
            Dim currentNode As BrowserNode = occNode.Parent
            Dim depth As Integer = 0
            
            While currentNode IsNot Nothing AndAlso depth < 10
                depth += 1
                Try
                    Dim label As String = currentNode.BrowserNodeDefinition.Label
                    
                    ' Check NativeObject type
                    Dim nativeObj As Object = Nothing
                    Dim nativeObjError As Boolean = False
                    Try
                        nativeObj = currentNode.NativeObject
                    Catch
                        nativeObjError = True
                    End Try
                    
                    ' If NativeObject threw error (Mirror patterns do this) and parent is root/folder
                    If nativeObjError Then
                        Try
                            If currentNode.Parent IsNot Nothing AndAlso IsAssemblyRootOrFolder(currentNode.Parent) Then
                                Return currentNode
                            End If
                        Catch
                        End Try
                    End If
                    
                    ' If NativeObject is an OccurrencePattern, we found it
                    If nativeObj IsNot Nothing AndAlso TypeOf nativeObj Is OccurrencePattern Then
                        Return currentNode
                    End If
                    
                    ' If this node is the assembly root, stop searching
                    If IsAssemblyRootOrFolder(currentNode) Then
                        Exit While
                    End If
                    
                Catch
                End Try
                
                Try
                    currentNode = currentNode.Parent
                Catch
                    Exit While
                End Try
            End While
        Catch
        End Try
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Checks if a browser node is the assembly root or a browser folder.
    ''' </summary>
    Private Function IsAssemblyRootOrFolder(node As BrowserNode) As Boolean
        If node Is Nothing Then Return False
        
        Dim label As String = ""
        Try
            label = node.BrowserNodeDefinition.Label
        Catch
        End Try
        
        Try
            ' Check if it's a folder
            Dim nativeObj As Object = node.NativeObject
            If nativeObj IsNot Nothing AndAlso TypeOf nativeObj Is BrowserFolder Then
                Return True
            End If
        Catch
        End Try
        
        Try
            ' Check if it's the top/root node (no parent)
            If node.Parent Is Nothing Then 
                Return True
            End If
        Catch
        End Try
        
        ' Check if label indicates assembly root (ends with .iam or contains .iam followed by space/bracket)
        If label.EndsWith(".iam") OrElse label.Contains(".iam ") OrElse label.Contains(".iam[") Then
            Return True
        End If
        
        Return False
    End Function

    ''' <summary>
    ''' Matches material name against folder patterns, returns folder name or empty string.
    ''' </summary>
    Private Function GetTargetFolder( _
        materialName As String, _
        folderPatterns As Dictionary(Of String, List(Of String))) As String
        
        For Each kvp As KeyValuePair(Of String, List(Of String)) In folderPatterns
            If UtilsLib.MaterialMatchesPatterns(materialName, kvp.Value) Then
                Return kvp.Key
            End If
        Next
        
        Return ""
    End Function

    ' ============================================================================
    ' SECTION 3: Model States
    ' ============================================================================

    ''' <summary>
    ''' Creates/updates model states based on state definitions.
    ''' </summary>
    Private Sub ApplyModelStates( _
        asmDoc As AssemblyDocument, _
        oPane As BrowserPane, _
        knownFolders As List(Of String), _
        stateDefinitions As Dictionary(Of String, List(Of String)))
        
        UtilsLib.LogInfo("SortingLib: Creating/updating model states...")
        
        Dim modelStates As Object = asmDoc.ComponentDefinition.ModelStates
        
        ' Process states: "all-enabled" first (where includedFolders is Nothing), then others
        Dim orderedStates As New List(Of String)
        For Each stateName As String In stateDefinitions.Keys
            If stateDefinitions(stateName) Is Nothing Then
                orderedStates.Insert(0, stateName)
            Else
                orderedStates.Add(stateName)
            End If
        Next
        
        For Each stateName As String In orderedStates
            Dim includedFolders As List(Of String) = stateDefinitions(stateName)
            
            UtilsLib.LogInfo("SortingLib: Processing model state '" & stateName & "'")
            
            ' Activate Primary first so new state copies from clean base
            ActivatePrimaryModelState(modelStates)
            
            ' Find or create the model state
            Dim ms As Object = FindModelState(modelStates, stateName)
            If ms Is Nothing Then
                ms = modelStates.Add(stateName)
                UtilsLib.LogInfo("SortingLib: Created model state '" & stateName & "'")
            End If
            
            ' Activate and configure
            ms.Activate()
            UtilsLib.LogInfo("SortingLib: Activated model state '" & stateName & "'")
            
            ' Apply suppression to folder contents
            For Each folderName As String In knownFolders
                Dim shouldSuppress As Boolean = (includedFolders IsNot Nothing) AndAlso (Not includedFolders.Contains(folderName))
                UtilsLib.LogInfo("SortingLib:   Folder '" & folderName & "': " & If(shouldSuppress, "suppress", "unsuppress"))
                
                Dim folder As BrowserFolder = UtilsLib.FindFolder(oPane, folderName)
                If folder Is Nothing Then 
                    UtilsLib.LogWarn("SortingLib:   Folder not found!")
                    Continue For
                End If
                
                ' Use folder-level suppression (handles patterns properly)
                SetFolderSuppression(asmDoc, folder, shouldSuppress)
            Next
            
            UtilsLib.LogInfo("SortingLib: Configured model state '" & stateName & "'")
        Next
    End Sub
    
    ''' <summary>
    ''' Suppresses or unsuppresses an entire folder using the UI command.
    ''' This properly handles all contents including Mirror Component Patterns.
    ''' </summary>
    Private Sub SetFolderSuppression(asmDoc As AssemblyDocument, folder As BrowserFolder, shouldSuppress As Boolean)
        ' Check current suppression state by looking at first item in folder
        Dim currentlySuppressed As Boolean = IsFolderSuppressed(folder)
        
        ' Only toggle if current state doesn't match desired state
        If currentlySuppressed = shouldSuppress Then
            UtilsLib.LogInfo("SortingLib:   Folder already in desired state")
            Return
        End If
        
        Try
            ' Select the folder and execute the suppress toggle command
            Dim app As Inventor.Application = asmDoc.Parent
            app.ActiveDocument.SelectSet.Clear()
            app.ActiveDocument.SelectSet.Select(folder)
            
            ' Execute the suppress/unsuppress context command (toggles state)
            app.CommandManager.ControlDefinitions.Item("AssemblyCompSuppressionCtxCmd").Execute()
            
            UtilsLib.LogInfo("SortingLib:   Folder suppression toggled via command")
        Catch ex As Exception
            UtilsLib.LogWarn("SortingLib:   Failed to toggle folder suppression: " & ex.Message)
        End Try
    End Sub
    
    ''' <summary>
    ''' Checks if a folder is currently suppressed by examining its contents.
    ''' Returns True if all items are suppressed (or folder is empty).
    ''' </summary>
    Private Function IsFolderSuppressed(folder As BrowserFolder) As Boolean
        For Each node As BrowserNode In folder.BrowserNode.BrowserNodes
            Try
                Dim nativeObj As Object = Nothing
                Try
                    nativeObj = node.NativeObject
                Catch
                    ' Mirror patterns throw exception - check via occurrences
                    Continue For
                End Try
                
                If nativeObj IsNot Nothing AndAlso TypeOf nativeObj Is ComponentOccurrence Then
                    Dim occ As ComponentOccurrence = CType(nativeObj, ComponentOccurrence)
                    If Not occ.Suppressed Then
                        Return False
                    End If
                End If
            Catch
            End Try
        Next
        
        Return True
    End Function
    
    ''' <summary>
    ''' Finds a model state by name, returns Nothing if not found.
    ''' </summary>
    Private Function FindModelState(modelStates As Object, stateName As String) As Object
        For Each ms As Object In modelStates
            If ms.Name = stateName Then
                Return ms
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Activates the [Primary] model state.
    ''' </summary>
    Private Sub ActivatePrimaryModelState(modelStates As Object)
        For Each ms As Object In modelStates
            If ms.Name = "[Primary]" Then
                ms.Activate()
                Exit For
            End If
        Next
    End Sub

    ' ============================================================================
    ' SECTION 4: Design Views (Visibility)
    ' ============================================================================

    ''' <summary>
    ''' Creates/updates design views based on view definitions.
    ''' Design views control visibility (what you see), not suppression (what's in BOM).
    ''' </summary>
    Private Sub ApplyDesignViews( _
        asmDoc As AssemblyDocument, _
        oPane As BrowserPane, _
        knownFolders As List(Of String), _
        viewDefinitions As Dictionary(Of String, List(Of String)))
        
        UtilsLib.LogInfo("SortingLib: Creating/updating design views...")
        
        ' First, restore Primary model state so all items are unsuppressed
        ' This ensures visibility checks work correctly
        Dim modelStates As Object = asmDoc.ComponentDefinition.ModelStates
        ActivatePrimaryModelState(modelStates)
        
        Dim repManager As RepresentationsManager = asmDoc.ComponentDefinition.RepresentationsManager
        Dim designViews As DesignViewRepresentations = repManager.DesignViewRepresentations
        
        For Each viewName As String In viewDefinitions.Keys
            Dim visibleFolders As List(Of String) = viewDefinitions(viewName)
            
            UtilsLib.LogInfo("SortingLib: Processing design view '" & viewName & "'")
            
            ' Activate Default view first so new view copies from clean base
            ActivateDefaultDesignView(designViews)
            
            ' Find or create the design view
            Dim dv As DesignViewRepresentation = FindDesignView(designViews, viewName)
            If dv Is Nothing Then
                dv = designViews.Add(viewName)
                UtilsLib.LogInfo("SortingLib: Created design view '" & viewName & "'")
            End If
            
            ' Disable camera auto-save so view only controls visibility
            Try
                dv.AutoSaveCamera = False
            Catch
            End Try
            
            ' Activate and configure
            dv.Activate()
            UtilsLib.LogInfo("SortingLib: Activated design view '" & viewName & "'")
            
            ' Apply visibility to folder contents
            For Each folderName As String In knownFolders
                Dim shouldBeVisible As Boolean = (visibleFolders Is Nothing) OrElse visibleFolders.Contains(folderName)
                UtilsLib.LogInfo("SortingLib:   Folder '" & folderName & "': " & If(shouldBeVisible, "visible", "hidden"))
                
                Dim folder As BrowserFolder = UtilsLib.FindFolder(oPane, folderName)
                If folder Is Nothing Then 
                    UtilsLib.LogWarn("SortingLib:   Folder not found!")
                    Continue For
                End If
                
                SetFolderVisibility(asmDoc, folder, shouldBeVisible)
            Next
            
            UtilsLib.LogInfo("SortingLib: Configured design view '" & viewName & "'")
        Next
    End Sub
    
    ''' <summary>
    ''' Sets visibility on an entire folder using the UI command.
    ''' This properly handles all contents including Mirror Component Patterns.
    ''' </summary>
    Private Sub SetFolderVisibility(asmDoc As AssemblyDocument, folder As BrowserFolder, shouldBeVisible As Boolean)
        ' Check current visibility state by looking at first item in folder
        Dim currentlyVisible As Boolean = IsFolderVisible(folder)
        
        ' Only toggle if current state doesn't match desired state
        If currentlyVisible = shouldBeVisible Then
            UtilsLib.LogInfo("SortingLib:   Folder already in desired visibility state")
            Return
        End If
        
        Try
            ' Select the folder and execute the visibility toggle command
            Dim app As Inventor.Application = asmDoc.Parent
            app.ActiveDocument.SelectSet.Clear()
            app.ActiveDocument.SelectSet.Select(folder)
            
            ' Execute the visibility context command (toggles state)
            app.CommandManager.ControlDefinitions.Item("AssemblyVisibilityCtxCmd").Execute()
            
            UtilsLib.LogInfo("SortingLib:   Folder visibility toggled via command")
        Catch ex As Exception
            UtilsLib.LogWarn("SortingLib:   Failed to toggle folder visibility: " & ex.Message)
        End Try
    End Sub
    
    ''' <summary>
    ''' Checks if a folder is currently visible by examining its contents.
    ''' Returns True if any item is visible.
    ''' </summary>
    Private Function IsFolderVisible(folder As BrowserFolder) As Boolean
        For Each node As BrowserNode In folder.BrowserNode.BrowserNodes
            Try
                Dim nativeObj As Object = Nothing
                Try
                    nativeObj = node.NativeObject
                Catch
                    ' Mirror patterns throw exception - skip
                    Continue For
                End Try
                
                If nativeObj IsNot Nothing AndAlso TypeOf nativeObj Is ComponentOccurrence Then
                    Dim occ As ComponentOccurrence = CType(nativeObj, ComponentOccurrence)
                    If occ.Visible Then
                        Return True
                    End If
                End If
            Catch
            End Try
        Next
        
        Return False
    End Function
    
    ''' <summary>
    ''' Finds a design view by name, returns Nothing if not found.
    ''' </summary>
    Private Function FindDesignView(designViews As DesignViewRepresentations, viewName As String) As DesignViewRepresentation
        For Each dv As DesignViewRepresentation In designViews
            If dv.Name = viewName Then
                Return dv
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Activates the Default design view.
    ''' </summary>
    Private Sub ActivateDefaultDesignView(designViews As DesignViewRepresentations)
        For Each dv As DesignViewRepresentation In designViews
            If dv.Name = "Default" OrElse dv.Name = "[Default]" Then
                dv.Activate()
                Exit For
            End If
        Next
    End Sub

    ''' <summary>
    ''' Restores the assembly to default state (Primary model state, Default design view).
    ''' </summary>
    Private Sub RestoreDefaults(asmDoc As AssemblyDocument)
        UtilsLib.LogInfo("SortingLib: Restoring defaults...")
        
        ' Restore Primary model state
        Dim modelStates As Object = asmDoc.ComponentDefinition.ModelStates
        For Each ms As Object In modelStates
            If ms.Name = "[Primary]" Then
                ms.Activate()
                Exit For
            End If
        Next
        
        ' Restore Default design view
        Dim designViews As DesignViewRepresentations = asmDoc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations
        ActivateDefaultDesignView(designViews)
        
        UtilsLib.LogInfo("SortingLib: Restored to Primary model state and Default design view")
    End Sub

End Module
