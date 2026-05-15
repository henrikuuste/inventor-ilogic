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
AddReference "System.Drawing"  ' Required for UI colors (ForeColor, Font)

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
Imports System.Collections.Generic

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
    
    ' Step 8: Placeholder numbers for plan construction (replaced by manifest reuse where applicable)
    Logger.Info("Loo elemendid: Generating placeholder numbers for plan...")
    Dim placeholderNumbers As List(Of String) = ElementReleaseLib.GeneratePlaceholderNumbers(requiredNumbers)
    
    ' Step 9: Compute release plan, then apply manifest reuse so preview shows real targets vs new numbers
    Logger.Info("Loo elemendid: Computing release plan...")
    context.ReleasePlan = ElementReleaseLib.ComputeReleasePlan(
        app, _
        context.AssemblyTree, _
        context.PartGroups, _
        context.Elements, _
        context.TargetRoot, _
        placeholderNumbers)
    
    Dim releaseManifestPath As String = ElementReleaseLib.GetReleaseManifestPath(context.SourceRoot)
    Dim releaseManifest As ElementReleaseLib.ReleaseManifest = ElementReleaseLib.ReadManifest(releaseManifestPath)
    ElementReleaseLib.ApplyReleaseManifestReuse(app, context.ReleasePlan, context.TargetRoot, releaseManifest, context.AssemblyTree)
    
    Dim reusePreview As Integer = 0
    Dim newPreview As Integer = 0
    For Each pf As ElementReleaseLib.PlannedFile In context.ReleasePlan.Files
        If pf.IsReuse Then reusePreview += 1
        If pf.IsPlaceholder Then newPreview += 1
    Next
    Logger.Info("Loo elemendid: Preview — taaskasutus: " & reusePreview & ", uued numbrid (eelvaade): " & newPreview)
    
    Dim allVariantNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    For Each el As ExcelReaderLib.ElementConfig In context.Elements
        allVariantNames.Add(el.ElementName)
    Next
    context.RemovedFiles = ElementReleaseLib.FindRemovedFilesFromManifest(releaseManifest, context.ReleasePlan, allVariantNames)
    If context.RemovedFiles.Count > 0 Then
        UtilsLib.LogWarn("Loo elemendid: " & context.RemovedFiles.Count & " faili eelmisest väljastusest puudub praeguses plaanis (orbaanid)")
    End If
    
    ' Step 10: Release UI — tree rows: 🔄 reuse · 📄 new · 🔗 shared
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
    
    ' Refresh orphan list for actual selection (plan rows already have reuse applied)
    Dim selectedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    For Each el As ExcelReaderLib.ElementConfig In context.Elements
        selectedNames.Add(el.ElementName)
    Next
    context.RemovedFiles = ElementReleaseLib.FindRemovedFilesFromManifest(releaseManifest, context.ReleasePlan, selectedNames)
    If context.RemovedFiles.Count > 0 Then
        UtilsLib.LogWarn("Loo elemendid: Orbaanid (ei uuendata): " & context.RemovedFiles.Count)
        For Each rm As ElementReleaseLib.FileMappingEntry In context.RemovedFiles
            UtilsLib.LogWarn("  - " & rm.TargetName & " (allikas: " & rm.SourceName & ", " & rm.FileType & ")")
        Next
    End If
    
    ' Step 12: Allocate new Vault numbers for any remaining placeholders
    Logger.Info("Loo elemendid: Allocating real file numbers...")
    If Not ElementReleaseLib.AllocateRealNumbers(context.ReleasePlan, context.TargetRoot, context.SourceRoot) Then
        Logger.Error("Loo elemendid: Failed to allocate file numbers")
        MessageBox.Show("Failinumbrite hankimine ebaõnnestus.", "Loo elemendid")
        Return
    End If
    Logger.Info("Loo elemendid: Real numbers allocated for " & context.ReleasePlan.Files.Count & " files")
    
    Dim readOnlyTargets As List(Of String) = ElementReleaseLib.ValidateTargetFilesWritable(context.ReleasePlan)
    If readOnlyTargets IsNot Nothing AndAlso readOnlyTargets.Count > 0 Then
        Logger.Error("Loo elemendid: Read-only sihtfailid (Vault checkout?): " & readOnlyTargets.Count)
        Dim listText As String = String.Join(vbCrLf, readOnlyTargets.ToArray())
        MessageBox.Show("Järgmised failid on kirjutuskaitstud. Tee Vaultis checkout ja proovi uuesti:" & vbCrLf & vbCrLf & listText, _
            "Loo elemendid", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return
    End If
    
    ' Step 13: Show execution form and execute release with progress tracking
    Logger.Info("Loo elemendid: Executing release...")
    
    ' Show execution form with file tree (same structure as preview)
    Dim execForm As Form = ElementReleaseUILib.ShowExecutionForm(context, context.ReleasePlan, context.Elements)
    
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
''' Calculate required file numbers based on context.
''' NOTE: Drawings do NOT get unique numbers - they derive their number from
''' the referenced part/assembly (with optional suffix preserved).
''' </summary>
Function CalculateRequiredNumbers(context As ElementReleaseLib.ElementReleaseContext) As Integer
    Dim requiredNumbers As Integer = 0

    ' Parts: count unique fingerprints
    For Each group2 As ElementReleaseLib.PartGroup In context.PartGroups
        If group2.UniqueFingerprints.Count = 1 Then
            requiredNumbers += 1
        Else
            requiredNumbers += group2.UniqueFingerprints.Count
        End If
    Next

    ' Assemblies: one per variant
    requiredNumbers += context.AssemblyTree.Assemblies.Count * context.Elements.Count
    
    ' External masters: one per variant (each element gets its own copy)
    requiredNumbers += context.AssemblyTree.ExternalMasters.Count * context.Elements.Count
    
    ' Intermediate assemblies: one per variant (each element gets its own copy)
    ' Only count those not already in tree.Assemblies
    For Each intAsm In context.AssemblyTree.IntermediateAssemblies
        If Not context.AssemblyTree.Assemblies.ContainsKey(intAsm) Then
            requiredNumbers += context.Elements.Count
        End If
    Next

    ' Drawings don't need unique numbers - they use the referenced model's number

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
