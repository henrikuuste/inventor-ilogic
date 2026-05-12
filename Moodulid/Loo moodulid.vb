' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Loo moodulid - Module Release System
' 
' Releases parametric Inventor modules with optimal file sharing.
' Analyzes variant parameters, computes geometry fingerprints, and creates
' standalone copies only where geometry differs. Shared parts are consolidated
' in a common folder (Ühine), reducing Vault file numbers and simplifying
' manufacturing.
'
' Usage: 
' 1. Open the main assembly of a base module (Alusmoodulid/{ModuleName}/*.iam)
' 2. Ensure moodulid.xlsx exists in the module folder
' 3. Run this rule
' 4. Select release mode
' 5. Review the plan and confirm
' 6. Files are created in Moodulid/{MooduliNimi}/ and Moodulid/Ühine/
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
AddVbFile "Lib/ExcelReaderLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"
AddVbFile "Lib/ModuleReleaseLib.vb"

Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    Logger.Info("Loo moodulid: Starting...")
    
    ' Validate active document
    Dim activeDoc As Document = app.ActiveDocument
    If activeDoc Is Nothing Then
        Logger.Error("Loo moodulid: No active document")
        MessageBox.Show("Aktiivne dokument puudub. Ava esmalt koost.", "Loo moodulid")
        Return
    End If
    
    If activeDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Error("Loo moodulid: Active document is not an assembly")
        MessageBox.Show("Aktiivne dokument pole koost. Ava alusmooduli põhikoost.", "Loo moodulid")
        Return
    End If
    
    Logger.Info("Loo moodulid: Active assembly: " & activeDoc.DisplayName)
    
    ' Step 1: Show mode selection dialog
    Dim mode As ModuleReleaseLib.ReleaseMode = ModuleReleaseLib.ShowModeSelectionDialog(app)
    If mode = ModuleReleaseLib.ReleaseMode.Cancelled Then
        Logger.Info("Loo moodulid: Cancelled by user")
        Return
    End If
    
    Logger.Info("Loo moodulid: Mode selected: " & mode.ToString())
    
    ' Step 2: Discover context
    Dim context As ModuleReleaseLib.ReleaseContext = ModuleReleaseLib.DiscoverContext(app, mode)
    If context Is Nothing Then
        Logger.Info("Loo moodulid: Context discovery failed")
        MessageBox.Show("Konteksti tuvastamine ebaõnnestus. Kontrolli logi.", "Loo moodulid")
        Return
    End If
    
    ' Step 3: Discover assembly tree
    Logger.Info("Loo moodulid: Discovering assembly tree...")
    context.AssemblyTree = ModuleReleaseLib.DiscoverAssemblyTree(app, _
        CType(activeDoc, AssemblyDocument).FullFileName, _
        context.SourceRoot)
    
    If context.AssemblyTree.Parts.Count = 0 Then
        Logger.Info("Loo moodulid: No parts found in assembly tree")
        MessageBox.Show("Koostus pole detaile. Kontrolli koost.", "Loo moodulid")
        Return
    End If
    
    ' Step 4: Discover drawings
    Logger.Info("Loo moodulid: Discovering drawings...")
    Dim searchFolders As New List(Of String)
    searchFolders.Add(context.SourceRoot)
    context.AssemblyTree.Drawings = ModuleReleaseLib.DiscoverDrawings(app, context.AssemblyTree, searchFolders)
    
    ' Step 5: Get master paths
    context.MasterPaths = ModuleReleaseLib.GetMasterPaths(context.AssemblyTree)
    Logger.Info("Loo moodulid: Found " & context.MasterPaths.Count & " master documents")
    
    ' Step 6: Build variant matrix (fingerprint analysis)
    Logger.Info("Loo moodulid: Building moodul matrix...")
    context.VariantMatrix = ModuleReleaseLib.BuildVariantMatrix(app, _
        context.AssemblyTree, _
        context.Variants, _
        context.MasterPaths)
    
    ' Step 7: Classify part groups
    Logger.Info("Loo moodulid: Classifying part groups...")
    context.PartGroups = ModuleReleaseLib.ClassifyPartGroups(context.VariantMatrix, context.AssemblyTree)
    
    Dim sharedCount As Integer = 0
    Dim uniqueCount As Integer = 0
    For Each group As ModuleReleaseLib.PartGroup In context.PartGroups
        If group.UniqueFingerprints.Count = 1 Then
            Dim moodulCount As Integer = GetFirstValueCount(group.UniqueFingerprints)
            If moodulCount > 1 Then
                sharedCount += 1
            Else
                uniqueCount += 1
            End If
        Else
            uniqueCount += group.UniqueFingerprints.Count
        End If
    Next
    Logger.Info("Loo moodulid: Shared parts: " & sharedCount & ", Unique parts: " & uniqueCount)
    
    ' Step 8: Calculate required file numbers
    Dim requiredNumbers As Integer = 0
    
    For Each group2 As ModuleReleaseLib.PartGroup In context.PartGroups
        If group2.UniqueFingerprints.Count = 1 Then
            requiredNumbers += 1
        Else
            requiredNumbers += group2.UniqueFingerprints.Count
        End If
    Next
    
    requiredNumbers += context.AssemblyTree.Assemblies.Count * context.Variants.Count
    
    ' Drawings - only count those that need unique numbers
    ' (drawings that start with their model's number reuse the model's number)
    Dim canShareDrawings As Boolean = (context.Variants.Count >= 2)
    For Each dwgInfo As ModuleReleaseLib.DrawingInfo In context.AssemblyTree.Drawings
        ' Check if drawing filename starts with its primary model's number
        Dim dwgFileName As String = System.IO.Path.GetFileNameWithoutExtension(dwgInfo.DrawingPath)
        Dim primaryModelPath As String = If(dwgInfo.ReferencedModelPaths.Count > 0, dwgInfo.ReferencedModelPaths(0), "")
        Dim modelNumber As String = System.IO.Path.GetFileNameWithoutExtension(primaryModelPath)
        Dim shareNumberWithModel As Boolean = Not String.IsNullOrEmpty(modelNumber) AndAlso _
            dwgFileName.StartsWith(modelNumber, StringComparison.OrdinalIgnoreCase)
        
        ' Skip number allocation if drawing reuses model's number
        If shareNumberWithModel Then
            Continue For
        End If
        
        Dim allRefsShared As Boolean = canShareDrawings
        If canShareDrawings Then
            For Each refPath In dwgInfo.ReferencedModelPaths
                Dim grp As ModuleReleaseLib.PartGroup = FindPartGroupByPath(context.PartGroups, refPath)
                If grp Is Nothing OrElse grp.UniqueFingerprints.Count > 1 Then
                    allRefsShared = False
                    Exit For
                End If
            Next
        End If
        If allRefsShared Then
            requiredNumbers += 1
        Else
            requiredNumbers += context.Variants.Count
        End If
    Next
    
    Logger.Info("Loo moodulid: Required file numbers: " & requiredNumbers)
    
    ' Step 9: Get file numbers
    Logger.Info("Loo moodulid: Getting file numbers...")
    Dim fileNumbers As List(Of String) = ModuleReleaseLib.GetFileNumbers(context.TargetRoot, requiredNumbers)
    
    If fileNumbers Is Nothing OrElse fileNumbers.Count < requiredNumbers Then
        Logger.Info("Loo moodulid: Failed to get enough file numbers")
        MessageBox.Show("Failinumbrite hankimine ebaõnnestus.", "Loo moodulid")
        Return
    End If
    
    ' Step 10: Compute release plan
    Logger.Info("Loo moodulid: Computing release plan...")
    context.ReleasePlan = ModuleReleaseLib.ComputeReleasePlan(
        context.AssemblyTree, _
        context.PartGroups, _
        context.Variants, _
        context.TargetRoot, _
        fileNumbers)
    
    ' Step 11: Show confirmation dialog
    If Not ModuleReleaseLib.ShowPlanConfirmationDialog(context.ReleasePlan) Then
        Logger.Info("Loo moodulid: Cancelled by user at confirmation")
        Return
    End If
    
    ' Step 12: Execute release
    Logger.Info("Loo moodulid: Executing release...")
    Dim success As Boolean = ModuleReleaseLib.ExecuteRelease(app, context)
    
    ' Step 13: Show completion summary
    If success Then
        ModuleReleaseLib.ShowCompletionSummary(context.ReleasePlan)
    Else
        MessageBox.Show("Väljastamine ebaõnnestus. Kontrolli logi.", "Loo moodulid", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End If
End Sub

Function FindPartGroupByPath(partGroups As List(Of ModuleReleaseLib.PartGroup), refPath As String) As ModuleReleaseLib.PartGroup
    For Each g As ModuleReleaseLib.PartGroup In partGroups
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
