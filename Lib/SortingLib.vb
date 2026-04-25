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
    ''' Component patterns cannot be moved via API and are logged for manual action.
    ''' Returns the list of folder names that were created (have parts).
    ''' </summary>
    Private Function ApplyFolders( _
        asmDoc As AssemblyDocument, _
        oPane As BrowserPane, _
        folderPatterns As Dictionary(Of String, List(Of String))) As List(Of String)
        
        UtilsLib.LogInfo("SortingLib: Assigning occurrences to folders...")
        
        ' First pass: determine which folders will have parts
        Dim foldersWithParts As New HashSet(Of String)
        
        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            Dim materialName As String = UtilsLib.GetOccurrenceMaterial(occ)
            If String.IsNullOrEmpty(materialName) Then Continue For
            
            Dim targetFolder As String = GetTargetFolder(materialName, folderPatterns)
            If Not String.IsNullOrEmpty(targetFolder) Then
                foldersWithParts.Add(targetFolder)
            End If
        Next
        
        ' Create only folders that will have parts
        For Each folderName As String In foldersWithParts
            UtilsLib.LogInfo("SortingLib: Creating/verifying folder '" & folderName & "'")
            UtilsLib.GetOrCreateFolder(oPane, folderName)
        Next
        
        ' Log folders that won't be created
        For Each folderName As String In folderPatterns.Keys
            If Not foldersWithParts.Contains(folderName) Then
                UtilsLib.LogInfo("SortingLib: Skipping folder '" & folderName & "' (no matching parts)")
            End If
        Next
        
        ' Track which patterns and nodes we've already processed
        Dim processedPatterns As New HashSet(Of String)
        Dim processedNodes As New HashSet(Of String)
        
        Dim assignedCount As Integer = 0
        Dim skippedCount As Integer = 0
        Dim patternCount As Integer = 0
        Dim errorCount As Integer = 0
        
        UtilsLib.LogInfo("SortingLib: Processing " & asmDoc.ComponentDefinition.Occurrences.Count & " occurrences...")
        
        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
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
                
                If isPatternElement Then
                    ' Component patterns cannot be moved to folders via API.
                    ' The pattern itself must be moved manually, but folder-level 
                    ' suppression will handle it during model state configuration.
                    Dim patternName As String = GetPatternName(occ, oPane)
                    
                    If Not String.IsNullOrEmpty(patternName) AndAlso Not processedPatterns.Contains(patternName) Then
                        processedPatterns.Add(patternName)
                        UtilsLib.LogInfo("SortingLib: Pattern '" & patternName & "' -> '" & targetFolder & "' (move pattern to folder manually)")
                        patternCount += 1
                    End If
                    Continue For
                Else
                    ' Standalone occurrence - move it directly
                    Dim oNode As BrowserNode = oPane.GetBrowserNodeFromObject(occ)
                    Dim movableNode As BrowserNode = UtilsLib.GetMovableParentNode(oPane, oNode)
                    
                    Dim nodeKey As String = movableNode.FullPath
                    
                    If processedNodes.Contains(nodeKey) Then Continue For
                    processedNodes.Add(nodeKey)
                    
                    Dim currentFolder As String = UtilsLib.GetNodeFolder(movableNode)
                    If currentFolder = targetFolder Then Continue For
                    
                    UtilsLib.LogInfo("SortingLib: Moving '" & movableNode.BrowserNodeDefinition.Label & "' to '" & targetFolder & "'")
                    Dim folder As BrowserFolder = UtilsLib.GetOrCreateFolder(oPane, targetFolder)
                    folder.Add(movableNode)
                    assignedCount += 1
                End If
                
            Catch ex As Exception
                UtilsLib.LogWarn("SortingLib: Error moving '" & occ.Name & "': " & ex.Message)
                errorCount += 1
            End Try
        Next
        
        UtilsLib.LogInfo("SortingLib: Moved " & assignedCount & " occurrence(s), " & patternCount & " pattern(s) need manual move, skipped " & skippedCount & " (no material), " & errorCount & " errors")
        
        Return New List(Of String)(foldersWithParts)
    End Function
    
    ''' <summary>
    ''' Gets the pattern name for a pattern element occurrence.
    ''' </summary>
    Private Function GetPatternName(occ As ComponentOccurrence, oPane As BrowserPane) As String
        ' Try via PatternElement.Parent
        Try
            Dim patternElement As Object = occ.PatternElement
            If patternElement IsNot Nothing AndAlso patternElement.Parent IsNot Nothing Then
                Return patternElement.Parent.Name
            End If
        Catch
        End Try
        
        ' Fallback: extract from FullPath
        Try
            Dim oNode As BrowserNode = oPane.GetBrowserNodeFromObject(occ)
            Return UtilsLib.ExtractPatternNameFromPath(oNode.FullPath)
        Catch
        End Try
        
        Return ""
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
