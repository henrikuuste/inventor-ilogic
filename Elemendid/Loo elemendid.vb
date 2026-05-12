' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Loo elemendid - Element Release System
' 
' Releases parametric Inventor elements with optimal file sharing.
' Analyzes element parameters, computes geometry fingerprints, and creates
' standalone copies only where geometry differs. Shared parts are consolidated
' in a common folder (Ühine), reducing Vault file numbers and simplifying
' manufacturing.
'
' Terminology updated 2026-05-12 per docs/UBIQUITOUS_LANGUAGE.md:
'   - "Alusmoodul" (old) → "Aluselement" (base element)
'   - "Moodul" (old) → "Väljastatud element" (released element)
'
' Usage: 
' 1. Open the main assembly of a base element (Aluselemendid/{ElementName}/*.iam)
' 2. Ensure elemendid.xlsx exists in the element folder
' 3. Run this rule
' 4. UI shows: left panel = elements with checkboxes, right panel = file tree preview
' 5. Select which elements to release (default: all)
' 6. Click "Väljasta" to start - progress window shows status
' 7. Files are created in Elemendid/{ElementName}/ and Elemendid/Ühine/
'
' Ref: docs/plans/2026-04-26-module-release-cycle.md
' ============================================================================

' References must come FIRST, before any AddVbFile
AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Connectivity.InventorAddin.EdmAddin"

' Libraries
AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/UILib.vb"
AddVbFile "Lib/ExcelReaderLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"
AddVbFile "Lib/OccurrenceNamingLib.vb"
AddVbFile "Lib/ElementReleaseLib.vb"
AddVbFile "Lib/ElementReleaseUILib.vb"

Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    Logger.Info("Loo elemendid: Starting...")
    
    ' Validate active document
    Dim activeDoc As Document = app.ActiveDocument
    If activeDoc Is Nothing Then
        Logger.Error("Loo elemendid: No active document")
        MessageBox.Show("Aktiivne dokument puudub. Ava esmalt koost.", "Loo elemendid")
        Return
    End If
    
    If activeDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Error("Loo elemendid: Active document is not an assembly")
        MessageBox.Show("Aktiivne dokument pole koost. Ava aluselemendi põhikoost.", "Loo elemendid")
        Return
    End If
    
    Logger.Info("Loo elemendid: Active assembly: " & activeDoc.DisplayName)
    
    ' Step 1: Discover context (always use FullElement mode - UI will allow element selection)
    Logger.Info("Loo elemendid: Discovering context...")
    Dim context As ElementReleaseLib.ElementReleaseContext = ElementReleaseLib.DiscoverContext(app, ElementReleaseLib.ReleaseMode.FullElement)
    If context Is Nothing Then
        Logger.Info("Loo elemendid: Context discovery failed")
        MessageBox.Show("Konteksti tuvastamine ebaõnnestus. Kontrolli logi.", "Loo elemendid")
        Return
    End If
    
    ' Step 2: Discover assembly tree
    Logger.Info("Loo elemendid: Discovering assembly tree...")
    context.AssemblyTree = ElementReleaseLib.DiscoverAssemblyTree(app, _
        CType(activeDoc, AssemblyDocument).FullFileName, _
        context.SourceRoot)
    
    If context.AssemblyTree.Parts.Count = 0 Then
        Logger.Info("Loo elemendid: No parts found in assembly tree")
        MessageBox.Show("Koostus pole detaile. Kontrolli koost.", "Loo elemendid")
        Return
    End If
    
    ' Step 3: Discover drawings
    Logger.Info("Loo elemendid: Discovering drawings...")
    Dim searchFolders As New List(Of String)
    searchFolders.Add(context.SourceRoot)
    context.AssemblyTree.Drawings = ElementReleaseLib.DiscoverDrawings(app, context.AssemblyTree, searchFolders)
    
    ' Step 4: Get master paths
    context.MasterPaths = ElementReleaseLib.GetMasterPaths(context.AssemblyTree)
    Logger.Info("Loo elemendid: Found " & context.MasterPaths.Count & " master documents")
    
    ' Step 5: Build element matrix (fingerprint analysis)
    Logger.Info("Loo elemendid: Building element matrix...")
    context.ElementMatrix = ElementReleaseLib.BuildElementMatrix(app, _
        context.AssemblyTree, _
        context.Elements, _
        context.MasterPaths)
    
    ' Step 6: Classify part groups
    Logger.Info("Loo elemendid: Classifying part groups...")
    context.PartGroups = ElementReleaseLib.ClassifyPartGroups(context.ElementMatrix, context.AssemblyTree)
    
    Dim sharedCount As Integer = 0
    Dim uniqueCount As Integer = 0
    For Each group As ElementReleaseLib.PartGroup In context.PartGroups
        If group.UniqueFingerprints.Count = 1 Then
            Dim elementCount As Integer = GetFirstValueCount(group.UniqueFingerprints)
            If elementCount > 1 Then
                sharedCount += 1
            Else
                uniqueCount += 1
            End If
        Else
            uniqueCount += group.UniqueFingerprints.Count
        End If
    Next
    Logger.Info("Loo elemendid: Shared parts: " & sharedCount & ", Unique parts: " & uniqueCount)
    
    ' Step 7: Calculate required file numbers
    Dim requiredNumbers As Integer = CalculateRequiredNumbers(context)
    Logger.Info("Loo elemendid: Required file numbers: " & requiredNumbers)
    
    ' Step 8: Generate PLACEHOLDER numbers for preview (not real Vault numbers yet)
    ' Real numbers are allocated only AFTER user confirms release
    Logger.Info("Loo elemendid: Generating placeholder numbers for preview...")
    Dim placeholderNumbers As List(Of String) = ElementReleaseLib.GeneratePlaceholderNumbers(requiredNumbers)
    
    ' Step 9: Compute release plan with placeholders (reads iProperties from source files)
    Logger.Info("Loo elemendid: Computing release plan preview...")
    context.ReleasePlan = ElementReleaseLib.ComputeReleasePlan(
        app, _
        context.AssemblyTree, _
        context.PartGroups, _
        context.Elements, _
        context.TargetRoot, _
        placeholderNumbers)
    
    ' Step 10: Show comprehensive UI for element selection and confirmation
    ' UI shows PLACEHOLDER numbers and PROJECTED properties (what files will become)
    Logger.Info("Loo elemendid: Showing release UI...")
    Dim uiResult As ElementReleaseUILib.ReleaseUIResult = ElementReleaseUILib.ShowReleaseUI( _
        app, context.Elements, context.ReleasePlan, context)
    
    If uiResult.Cancelled Then
        Logger.Info("Loo elemendid: Cancelled by user")
        Return
    End If
    
    ' Step 11: Filter release plan based on selected elements
    Logger.Info("Loo elemendid: Filtering plan for selected elements...")
    Logger.Info("Loo elemendid: Selected elements: " & uiResult.SelectedElements.Count)
    
    ' Update context with selected elements
    context.Elements = uiResult.SelectedElements
    
    ' Filter the release plan
    context.ReleasePlan = ElementReleaseLib.FilterReleasePlan(context.ReleasePlan, uiResult.SelectedElements)
    Logger.Info("Loo elemendid: Filtered plan has " & context.ReleasePlan.Files.Count & " files")
    
    If context.ReleasePlan.Files.Count = 0 Then
        Logger.Info("Loo elemendid: No files to release after filtering")
        MessageBox.Show("Valitud elementide jaoks pole faile väljastamiseks.", "Loo elemendid")
        Return
    End If
    
    ' Step 12: NOW allocate real file numbers (only after user confirms)
    ' This is when we actually consume Vault numbers
    Logger.Info("Loo elemendid: Allocating real file numbers...")
    If Not ElementReleaseLib.AllocateRealNumbers(context.ReleasePlan, context.TargetRoot) Then
        Logger.Error("Loo elemendid: Failed to allocate file numbers")
        MessageBox.Show("Failinumbrite hankimine ebaõnnestus.", "Loo elemendid")
        Return
    End If
    Logger.Info("Loo elemendid: Real numbers allocated for " & context.ReleasePlan.Files.Count & " files")
    
    ' Step 13: Show execution form and execute release with progress tracking
    Logger.Info("Loo elemendid: Executing release...")
    
    ' Show execution form with file tree
    Dim execForm As Form = ElementReleaseUILib.ShowExecutionForm(context, context.ReleasePlan)
    
    ' Log start
    ElementReleaseUILib.LogMessage("Väljastamine alustatud: " & DateTime.Now.ToString("HH:mm:ss"))
    ElementReleaseUILib.LogMessage("Elemente: " & context.Elements.Count)
    ElementReleaseUILib.LogMessage("Faile: " & context.ReleasePlan.Files.Count)
    ElementReleaseUILib.LogMessage("")
    
    ' Execute with progress tracking
    Dim success As Boolean = ExecuteReleaseWithProgress(app, context)
    
    ' Mark execution complete
    ElementReleaseUILib.MarkExecutionComplete(success)
    
    If success Then
        Logger.Info("Loo elemendid: Release completed successfully")
    Else
        Logger.Error("Loo elemendid: Release failed")
    End If
    
    ' Wait for user to close execution form
    ElementReleaseUILib.WaitForExecutionFormClose()
End Sub

''' <summary>
''' Executes release with progress tracking via UI callbacks
''' </summary>
Function ExecuteReleaseWithProgress(app As Inventor.Application, _
                                    context As ElementReleaseLib.ElementReleaseContext) As Boolean
    Try
        ' Execute the release - it will use the callbacks we set up
        Dim success As Boolean = ElementReleaseLib.ExecuteRelease(app, context)
        Return success
    Catch ex As Exception
        Logger.Error("ExecuteReleaseWithProgress: " & ex.Message)
        ElementReleaseUILib.LogMessage("VIGA: " & ex.Message)
        Return False
    End Try
End Function

''' <summary>
''' Calculate required file numbers based on context
''' </summary>
Function CalculateRequiredNumbers(context As ElementReleaseLib.ElementReleaseContext) As Integer
    Dim requiredNumbers As Integer = 0
    
    For Each group2 As ElementReleaseLib.PartGroup In context.PartGroups
        If group2.UniqueFingerprints.Count = 1 Then
            requiredNumbers += 1
        Else
            requiredNumbers += group2.UniqueFingerprints.Count
        End If
    Next
    
    requiredNumbers += context.AssemblyTree.Assemblies.Count * context.Elements.Count
    
    ' Drawings - only count those that need unique numbers
    Dim canShareDrawings As Boolean = (context.Elements.Count >= 2)
    For Each dwgInfo As ElementReleaseLib.DrawingInfo In context.AssemblyTree.Drawings
        Dim dwgFileName As String = System.IO.Path.GetFileNameWithoutExtension(dwgInfo.DrawingPath)
        Dim primaryModelPath As String = If(dwgInfo.ReferencedModelPaths.Count > 0, dwgInfo.ReferencedModelPaths(0), "")
        Dim modelNumber As String = System.IO.Path.GetFileNameWithoutExtension(primaryModelPath)
        Dim shareNumberWithModel As Boolean = Not String.IsNullOrEmpty(modelNumber) AndAlso _
            dwgFileName.StartsWith(modelNumber, StringComparison.OrdinalIgnoreCase)
        
        If shareNumberWithModel Then
            Continue For
        End If
        
        Dim allRefsShared As Boolean = canShareDrawings
        If canShareDrawings Then
            For Each refPath In dwgInfo.ReferencedModelPaths
                Dim grp As ElementReleaseLib.PartGroup = FindPartGroupByPath(context.PartGroups, refPath)
                If grp Is Nothing OrElse grp.UniqueFingerprints.Count > 1 Then
                    allRefsShared = False
                    Exit For
                End If
            Next
        End If
        If allRefsShared Then
            requiredNumbers += 1
        Else
            requiredNumbers += context.Elements.Count
        End If
    Next
    
    Return requiredNumbers
End Function

Function FindPartGroupByPath(partGroups As List(Of ElementReleaseLib.PartGroup), refPath As String) As ElementReleaseLib.PartGroup
    For Each g As ElementReleaseLib.PartGroup In partGroups
        If g.PartPath.Equals(refPath, StringComparison.OrdinalIgnoreCase) Then
            Return g
        End If
    Next
    Return Nothing
End Function

Function GetFirstValueCount(dict As Dictionary(Of String, List(Of String))) As Integer
    For Each v As List(Of String) In dict.Values
        Return v.Count
    Next
    Return 0
End Function
