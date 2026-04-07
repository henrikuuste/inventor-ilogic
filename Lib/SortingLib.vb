' ============================================================================
' SortingLib - Assembly Component Sorting by Material
' 
' Categorizes assembly components into browser folders based on material name
' patterns, then creates model states with suppression for different folder 
' combinations.
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
        
        UtilsLib.LogInfo("SortingLib: Starting for " & asmDoc.DisplayName)
        
        Dim oPane As BrowserPane = asmDoc.BrowserPanes.Item("Model")
        Dim knownFolders As New List(Of String)(folderPatterns.Keys)
        
        ' Step 1: Assign occurrences to folders based on material
        ApplyFolders(asmDoc, oPane, folderPatterns)
        
        ' Step 2: Create/update model states
        ApplyModelStates(asmDoc, oPane, knownFolders, stateDefinitions)
        
        ' Step 3: Restore defaults
        RestoreDefaults(asmDoc)
        
        UtilsLib.LogInfo("SortingLib: Completed")
    End Sub

    ' ============================================================================
    ' SECTION 2: Folder Assignment
    ' ============================================================================

    ''' <summary>
    ''' Creates folders and assigns occurrences based on material patterns.
    ''' Component patterns cannot be moved via API and are logged for manual action.
    ''' </summary>
    Private Sub ApplyFolders( _
        asmDoc As AssemblyDocument, _
        oPane As BrowserPane, _
        folderPatterns As Dictionary(Of String, List(Of String)))
        
        UtilsLib.LogInfo("SortingLib: Assigning occurrences to folders...")
        
        ' First, ensure all folders exist
        For Each folderName As String In folderPatterns.Keys
            UtilsLib.LogInfo("SortingLib: Creating/verifying folder '" & folderName & "'")
            UtilsLib.GetOrCreateFolder(oPane, folderName)
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
                    ' Log warning for manual action.
                    Dim patternName As String = GetPatternName(occ, oPane)
                    
                    If Not String.IsNullOrEmpty(patternName) AndAlso Not processedPatterns.Contains(patternName) Then
                        processedPatterns.Add(patternName)
                        UtilsLib.LogWarn("SortingLib: Pattern '" & patternName & "' -> '" & targetFolder & "' - must be moved manually (API limitation)")
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
    End Sub
    
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
                
                SetFolderContentsSuppression(folder, shouldSuppress)
            Next
            
            UtilsLib.LogInfo("SortingLib: Configured model state '" & stateName & "'")
        Next
    End Sub
    
    ''' <summary>
    ''' Sets suppression on all items in a folder.
    ''' Mirror Component Patterns are SKIPPED to avoid breaking their associative relationship.
    ''' </summary>
    Private Sub SetFolderContentsSuppression(folder As BrowserFolder, shouldSuppress As Boolean)
        For Each node As BrowserNode In folder.BrowserNode.BrowserNodes
            Try
                Dim nativeObj As Object = Nothing
                Dim isMirrorPattern As Boolean = False
                
                Try
                    nativeObj = node.NativeObject
                Catch
                    ' NativeObject throws E_NOTIMPL for Mirror Component Patterns
                    isMirrorPattern = True
                End Try
                
                If isMirrorPattern Then
                    ' Mirror Component Pattern - SKIP to avoid breaking associativity
                    Dim label As String = ""
                    Try : label = node.BrowserNodeDefinition.Label : Catch : End Try
                    UtilsLib.LogWarn("SortingLib: Mirror pattern '" & label & "' - cannot suppress via API (would break pattern)")
                    Continue For
                End If
                
                If nativeObj IsNot Nothing AndAlso TypeOf nativeObj Is ComponentOccurrence Then
                    ' Direct occurrence
                    UtilsLib.SuppressOccurrence(CType(nativeObj, ComponentOccurrence), shouldSuppress)
                ElseIf nativeObj IsNot Nothing AndAlso TypeName(nativeObj).Contains("Pattern") Then
                    ' Regular pattern (Rectangular, Circular) - suppress via OccurrencePatternElements
                    Try
                        Dim elements As Object = nativeObj.OccurrencePatternElements
                        If elements IsNot Nothing Then
                            For Each elem As Object In elements
                                Try
                                    Dim occs As Object = elem.Occurrences
                                    If occs IsNot Nothing Then
                                        For Each patternOcc As ComponentOccurrence In occs
                                            UtilsLib.SuppressOccurrence(patternOcc, shouldSuppress)
                                        Next
                                    End If
                                Catch
                                End Try
                            Next
                        End If
                    Catch
                    End Try
                End If
            Catch
            End Try
        Next
    End Sub
    
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

    ''' <summary>
    ''' Restores the assembly to default state (Primary model state).
    ''' </summary>
    Private Sub RestoreDefaults(asmDoc As AssemblyDocument)
        UtilsLib.LogInfo("SortingLib: Restoring defaults...")
        
        Dim modelStates As Object = asmDoc.ComponentDefinition.ModelStates
        For Each ms As Object In modelStates
            If ms.Name = "[Primary]" Then
                ms.Activate()
                Exit For
            End If
        Next
        
        UtilsLib.LogInfo("SortingLib: Restored to Primary model state")
    End Sub

End Module
