' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' ElementReleaseLib - Element Release System Core Library
' 
' Provides functions for releasing parametric Inventor elements with optimal
' file sharing. Analyzes variant parameters, computes geometry fingerprints,
' and creates standalone copies only where geometry differs.
'
' Terminology updated 2026-05-12: module -> element per UBIQUITOUS_LANGUAGE.md
'
' Dependencies:
'   UtilsLib - logging via UtilsLib.LogInfo / UtilsLib.LogWarn
'   ExcelReaderLib - variant table reading (ElementConfig)
'   VaultNumberingLib - Vault operations (production mode only)
'
' Usage: 
'   In calling script (BEFORE AddVbFile):
'     AddReference "Autodesk.Connectivity.WebServices"
'     AddReference "Autodesk.DataManagement.Client.Framework.Vault"
'     AddReference "Connectivity.InventorAddin.EdmAddin"
'     AddVbFile "Lib/UtilsLib.vb"
'     AddVbFile "Lib/ExcelReaderLib.vb"
'     AddVbFile "Lib/VaultNumberingLib.vb"
'     AddVbFile "Lib/ElementReleaseLib.vb"
'
' ============================================================================

Imports Inventor
Imports System.Collections.Generic
Imports System.Windows.Forms

Public Module ElementReleaseLib

    ' ============================================================================
    ' Configuration Constants
    ' ============================================================================
    
    Public Const DEVELOPMENT_MODE As Boolean = True
    Public Const NUMBERING_SCHEME As String = "Test numbriskeem"
    
    ' TEMPORARY: Phased feature flags to isolate crash
    ' Enable one at a time to find the culprit
    Public Const ENABLE_MULTI_MASTER_DISCOVERY As Boolean = True  ' Master flag
    Public Const ENABLE_DISCOVER_ALL_MASTERS As Boolean = True    ' Phase 1: DiscoverAllMasters
    Public Const ENABLE_DETECT_PROJECTED_GEOM As Boolean = True  ' Phase 2: DetectProjectedGeometry
    Public Const ENABLE_SEARCH_PRODUCT_FAMILY As Boolean = True  ' Phase 3: SearchProductFamilyForIntermediates
    
    ' ============================================================================
    ' Enums and Data Structures
    ' ============================================================================
    
    Public Enum ReleaseMode
        Cancelled = 0
        FullElement = 1
        CurrentAssembly = 2
    End Enum
    
    Public Enum PartRole
        Derived
        Manual
    End Enum
    
    Public Enum FileType
        Part
        Assembly
        Drawing
    End Enum
    
    ''' <summary>
    ''' Main context object carrying all release information.
    ''' </summary>
    Public Class ElementReleaseContext
        Public Mode As ReleaseMode
        Public ExcelPath As String
        Public Elements As List(Of ExcelReaderLib.ElementConfig)
        Public SourceRoot As String
        Public TargetRoot As String
        Public ElementName As String
        Public AssemblyTree As AssemblyTree
        Public ElementMatrix As ElementMatrix
        Public PartGroups As List(Of PartGroup)
        Public ReleasePlan As ReleasePlan
        Public MasterPaths As List(Of String)
    End Class
    
    ''' <summary>
    ''' Assembly tree structure with all discovered files.
    ''' </summary>
    Public Class AssemblyTree
        Public RootAssemblyPath As String
        Public SourceRoot As String
        Public Parts As New Dictionary(Of String, PartInfo)(StringComparer.OrdinalIgnoreCase)
        Public Assemblies As New Dictionary(Of String, AssemblyInfo)(StringComparer.OrdinalIgnoreCase)
        Public Drawings As New List(Of DrawingInfo)
        
        ' Multi-master support (Phase 1)
        Public ExternalMasters As New Dictionary(Of String, MasterInfo)(StringComparer.OrdinalIgnoreCase)
        Public MasterDependencies As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
        Public IntermediateAssemblies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Public ProjectedGeometryChains As New List(Of Tuple(Of String, String, String)) ' (SourcePart, ViaAssembly, TargetPart)
    End Class
    
    ''' <summary>
    ''' Information about a master file (skeleton/eskiis part) for multi-master support.
    ''' </summary>
    Public Class MasterInfo
        Public FilePath As String
        Public RelativePath As String         ' Relative to its own Aluselemendid folder
        Public SourceElement As String        ' Which element it comes from
        Public DependsOn As New List(Of String) ' Other masters this one references
        Public IsIntermediate As Boolean      ' True if this is an assembly used for projection
    End Class
    
    Public Class PartInfo
        Public FilePath As String
        Public RelativePath As String
        Public Role As PartRole
        Public DerivedFromMaster As String
        Public BodyName As String
        Public PartNumber As String
    End Class
    
    Public Class AssemblyInfo
        Public FilePath As String
        Public RelativePath As String
    End Class
    
    Public Class DrawingInfo
        Public DrawingPath As String
        Public RelativePath As String
        Public ReferencedModelPaths As New List(Of String)
    End Class
    
    ''' <summary>
    ''' Element matrix with fingerprints per part per released element.
    ''' </summary>
    Public Class ElementMatrix
        Public PartPaths As New List(Of String)
        Public ElementNames As New List(Of String)
        Public Fingerprints As New Dictionary(Of String, Dictionary(Of String, String))
    End Class
    
    ''' <summary>
    ''' Part group classification for sharing detection.
    ''' </summary>
    Public Class PartGroup
        Public PartPath As String
        Public RelativePath As String
        Public PartNumber As String
        Public UniqueFingerprints As New Dictionary(Of String, List(Of String))
    End Class
    
    ''' <summary>
    ''' Release plan with all planned files.
    ''' </summary>
    Public Class ReleasePlan
        Public Files As New List(Of PlannedFile)
        Public SharedFolder As String
        Public VariantFolders As New Dictionary(Of String, String)
    End Class
    
    Public Class PlannedFile
        Public SourcePath As String
        Public TargetVaultPath As String
        Public TargetLocalPath As String
        Public VaultNumber As String              ' Placeholder during preview, real number after allocation
        Public FileType As FileType
        Public IsShared As Boolean
        Public IsExisting As Boolean
        Public ForVariants As New List(Of String)
        Public ForModules As New List(Of String)
        Public Fingerprint As String
        ' iProperties from source file
        Public SourceDescription As String
        Public SourceProject As String
        Public SourcePartNumber As String
        ' Projected properties (what they WILL BE after release)
        Public ProjectedDescription As String     ' What Description will be set to
        Public ProjectedProject As String         ' What Project will be set to  
        Public IsPlaceholder As Boolean = True    ' True until real number is assigned
    End Class
    
    ' ============================================================================
    ' Phase 1: UI and Discovery
    ' ============================================================================
    
    ''' <summary>
    ''' Show mode selection dialog.
    ''' Returns the selected release mode or Cancelled.
    ''' </summary>
    Public Function ShowModeSelectionDialog(app As Inventor.Application) As ReleaseMode
        Dim frm As New Form()
        frm.Text = StringsLib.TITLE_ELEMENT_RELEASE
        frm.Width = 350
        frm.Height = 180
        frm.StartPosition = FormStartPosition.Manual
        frm.Left = 100
        frm.Top = 100
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False
        frm.MinimizeBox = False
        frm.TopMost = True
        
        frm.Tag = ReleaseMode.Cancelled
        
        Dim lblPrompt As New Label()
        lblPrompt.Text = "Vali väljastamise režiim (elemendid Excelist):"
        lblPrompt.Left = 20
        lblPrompt.Top = 20
        lblPrompt.Width = 300
        frm.Controls.Add(lblPrompt)
        
        Dim btnFull As New Button()
        btnFull.Text = StringsLib.BTN_ALL_ELEMENTS
        btnFull.Left = 20
        btnFull.Top = 50
        btnFull.Width = 140
        btnFull.Height = 35
        AddHandler btnFull.Click, Sub(s, e)
            frm.Tag = ReleaseMode.FullElement
            frm.DialogResult = DialogResult.OK
        End Sub
        frm.Controls.Add(btnFull)
        
        Dim btnCurrent As New Button()
        btnCurrent.Text = StringsLib.BTN_FIRST_ELEMENT
        btnCurrent.Left = 175
        btnCurrent.Top = 50
        btnCurrent.Width = 140
        btnCurrent.Height = 35
        AddHandler btnCurrent.Click, Sub(s, e)
            frm.Tag = ReleaseMode.CurrentAssembly
            frm.DialogResult = DialogResult.OK
        End Sub
        frm.Controls.Add(btnCurrent)
        
        Dim btnCancel As New Button()
        btnCancel.Text = "Tühista"
        btnCancel.Left = 125
        btnCancel.Top = 100
        btnCancel.Width = 100
        btnCancel.Height = 28
        btnCancel.DialogResult = DialogResult.Cancel
        frm.Controls.Add(btnCancel)
        frm.CancelButton = btnCancel
        
        Dim result As DialogResult = frm.ShowDialog()
        
        If result = DialogResult.OK Then
            Return CType(frm.Tag, ReleaseMode)
        End If
        
        Return ReleaseMode.Cancelled
    End Function
    
    ''' <summary>
    ''' Discover the Excel configuration file for the element.
    ''' Looks for elemendid.xlsx (or moodulid.xlsx for backward compatibility) in the source folder.
    ''' </summary>
    Public Function DiscoverExcel(sourceFolder As String) As String
        If String.IsNullOrEmpty(sourceFolder) OrElse Not System.IO.Directory.Exists(sourceFolder) Then
            UtilsLib.LogInfo("DiscoverExcel: Source folder not found: " & sourceFolder)
            Return Nothing
        End If
        
        UtilsLib.LogInfo("DiscoverExcel: Searching in " & sourceFolder)
        
        ' Look for elemendid.xlsx first (new naming)
        Dim excelPath As String = System.IO.Path.Combine(sourceFolder, "elemendid.xlsx")
        If System.IO.File.Exists(excelPath) Then
            UtilsLib.LogInfo("DiscoverExcel: Found elemendid.xlsx")
            Return excelPath
        End If
        
        ' Backward compatibility: look for moodulid.xlsx
        excelPath = System.IO.Path.Combine(sourceFolder, "moodulid.xlsx")
        If System.IO.File.Exists(excelPath) Then
            UtilsLib.LogInfo("DiscoverExcel: Found moodulid.xlsx (legacy name)")
            Return excelPath
        End If
        
        ' Fallback to .xls
        excelPath = System.IO.Path.Combine(sourceFolder, "elemendid.xls")
        If System.IO.File.Exists(excelPath) Then
            UtilsLib.LogInfo("DiscoverExcel: Found elemendid.xls")
            Return excelPath
        End If
        
        UtilsLib.LogInfo("DiscoverExcel: No elemendid.xlsx found")
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Discover release context based on mode selection.
    ''' </summary>
    Public Function DiscoverContext(app As Inventor.Application, mode As ReleaseMode) As ElementReleaseContext
        If mode = ReleaseMode.Cancelled Then Return Nothing
        
        Dim context As New ElementReleaseContext()
        context.Mode = mode
        
        Dim activeDoc As Document = app.ActiveDocument
        If activeDoc Is Nothing Then
            UtilsLib.LogInfo("DiscoverContext: No active document")
            Return Nothing
        End If
        
        If activeDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            UtilsLib.LogInfo("DiscoverContext: Active document is not an assembly")
            Return Nothing
        End If
        
        Dim asmDoc As AssemblyDocument = CType(activeDoc, AssemblyDocument)
        Dim asmPath As String = asmDoc.FullFileName
        
        UtilsLib.LogInfo("DiscoverContext: Active assembly: " & asmPath)
        
        context.SourceRoot = FindElementSourceRoot(asmPath)
        If String.IsNullOrEmpty(context.SourceRoot) Then
            UtilsLib.LogInfo("DiscoverContext: Could not determine source root")
            Return Nothing
        End If
        
        context.ElementName = System.IO.Path.GetFileName(context.SourceRoot)
        UtilsLib.LogInfo("DiscoverContext: Element name: " & context.ElementName)
        
        context.TargetRoot = ComputeTargetRoot(context.SourceRoot)
        UtilsLib.LogInfo("DiscoverContext: Target root: " & context.TargetRoot)
        
        ' Always require Excel file
        context.ExcelPath = DiscoverExcel(context.SourceRoot)
        If String.IsNullOrEmpty(context.ExcelPath) Then
            UtilsLib.LogInfo("DiscoverContext: Excel file required")
            MessageBox.Show("Excel faili (elemendid.xlsx) ei leitud kaustast:" & vbCrLf & _
                           context.SourceRoot & vbCrLf & vbCrLf & _
                           "Loo Excel fail elementide kirjeldusega." & vbCrLf & _
                           "Vaata malli: Elemendid\_elemendid_template.xlsx", "Loo elemendid")
            Return Nothing
        End If
        
        Dim allVariants As List(Of ExcelReaderLib.ElementConfig) = ExcelReaderLib.ReadVariantTable(context.ExcelPath)
        UtilsLib.LogInfo("DiscoverContext: Loaded " & allVariants.Count & " elements from Excel")
        
        If allVariants.Count = 0 Then
            UtilsLib.LogInfo("DiscoverContext: No elements found in Excel")
            MessageBox.Show("Excel fail on tühi või vigane:" & vbCrLf & _
                           context.ExcelPath, "Loo elemendid")
            Return Nothing
        End If
        
        ' Filter based on mode
        If mode = ReleaseMode.FullElement Then
            context.Elements = allVariants
            UtilsLib.LogInfo("DiscoverContext: Full element mode - using all " & context.Elements.Count & " elements")
        Else
            ' CurrentAssembly mode - use only first element
            context.Elements = New List(Of ExcelReaderLib.ElementConfig)
            context.Elements.Add(allVariants(0))
            UtilsLib.LogInfo("DiscoverContext: Single element mode - using: " & allVariants(0).ElementName)
        End If
        
        Return context
    End Function
    
    ''' <summary>
    ''' Find the element source root folder (Aluselemendid/{ElementName}).
    ''' Also supports legacy Alusmoodulid for backward compatibility.
    ''' </summary>
    Private Function FindElementSourceRoot(asmPath As String) As String
        Dim folder As String = System.IO.Path.GetDirectoryName(asmPath)
        
        Do While Not String.IsNullOrEmpty(folder)
            Dim parentName As String = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(folder))
            ' Support both new (Aluselemendid) and legacy (Alusmoodulid) folder names
            If parentName.Equals("Aluselemendid", StringComparison.OrdinalIgnoreCase) OrElse _
               parentName.Equals("Alusmoodulid", StringComparison.OrdinalIgnoreCase) Then
                UtilsLib.LogInfo("FindElementSourceRoot: Found element root: " & folder)
                Return folder
            End If
            folder = System.IO.Path.GetDirectoryName(folder)
        Loop
        
        UtilsLib.LogInfo("FindElementSourceRoot: Aluselemendid not found in path, using assembly folder")
        Return System.IO.Path.GetDirectoryName(asmPath)
    End Function
    
    ''' <summary>
    ''' Compute target root folder (Elemendid/ at same level as Aluselemendid/).
    ''' Also supports legacy Alusmoodulid → Moodulid for backward compatibility.
    ''' </summary>
    Private Function ComputeTargetRoot(sourceRoot As String) As String
        Dim parent As String = System.IO.Path.GetDirectoryName(sourceRoot)
        Dim grandparent As String = System.IO.Path.GetDirectoryName(parent)
        
        ' New naming: Aluselemendid → Elemendid
        If System.IO.Path.GetFileName(parent).Equals("Aluselemendid", StringComparison.OrdinalIgnoreCase) Then
            Return System.IO.Path.Combine(grandparent, "Elemendid")
        End If
        
        ' Legacy naming: Alusmoodulid → Moodulid (backward compatibility)
        If System.IO.Path.GetFileName(parent).Equals("Alusmoodulid", StringComparison.OrdinalIgnoreCase) Then
            Return System.IO.Path.Combine(grandparent, "Moodulid")
        End If
        
        Return System.IO.Path.Combine(parent, "Elemendid")
    End Function
    
    ''' <summary>
    ''' Check if a path contains OldVersions folder (case-insensitive).
    ''' OldVersions is a special Vault folder that stores previous versions and should not be released.
    ''' </summary>
    Private Function IsOldVersionsPath(path As String) As Boolean
        If String.IsNullOrEmpty(path) Then Return False
        Dim lowerPath As String = path.ToLower()
        Return lowerPath.Contains("\oldversions\") OrElse lowerPath.Contains("/oldversions/")
    End Function
    
    ''' <summary>
    ''' Check if a file path represents an old Vault version.
    ''' Detects both OldVersions folder and versioned filenames (e.g., 000114.0025.iam).
    ''' </summary>
    Private Function IsVaultOldVersion(filePath As String) As Boolean
        If String.IsNullOrEmpty(filePath) Then Return False
        
        ' Check for OldVersions folder
        If IsOldVersionsPath(filePath) Then Return True
        
        ' Check for versioned filename pattern: 000114.0025.iam (has numeric suffix before extension)
        Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
        If fileName.Contains(".") Then
            Dim parts() As String = fileName.Split("."c)
            If parts.Length >= 2 Then
                Dim lastPart As String = parts(parts.Length - 1)
                ' If the last part (before extension) is all digits, it's a version number
                If lastPart.All(Function(c) Char.IsDigit(c)) Then
                    Return True
                End If
            End If
        End If
        
        Return False
    End Function
    
    ''' <summary>
    ''' Get the product family root folder by walking up from sourceRoot to find "Aluselemendid" parent.
    ''' Falls back to 3 levels up if Aluselemendid not found.
    ''' </summary>
    Private Function GetProductFamilyRoot(sourceRoot As String) As String
        Dim current As String = sourceRoot
        
        ' Walk up looking for Aluselemendid folder
        For i As Integer = 1 To 5
            If String.IsNullOrEmpty(current) Then Exit For
            Dim parent As String = System.IO.Path.GetDirectoryName(current)
            If String.IsNullOrEmpty(parent) Then Exit For
            
            Dim folderName As String = System.IO.Path.GetFileName(current)
            If folderName.Equals("Aluselemendid", StringComparison.OrdinalIgnoreCase) Then
                Return parent ' Return the folder containing Aluselemendid
            End If
            current = parent
        Next
        
        ' Fallback: go up 3 levels from sourceRoot
        current = sourceRoot
        For i As Integer = 1 To 3
            Dim parent As String = System.IO.Path.GetDirectoryName(current)
            If Not String.IsNullOrEmpty(parent) Then
                current = parent
            End If
        Next
        
        Return current
    End Function
    
    ' ============================================================================
    ' Phase 2: File Discovery and Classification
    ' ============================================================================
    
    ''' <summary>
    ''' Discover the complete assembly tree with all parts and sub-assemblies.
    ''' </summary>
    Public Function DiscoverAssemblyTree(app As Inventor.Application, _
                                          rootAsmPath As String, _
                                          sourceRoot As String) As AssemblyTree
        Dim tree As New AssemblyTree()
        tree.RootAssemblyPath = rootAsmPath
        tree.SourceRoot = sourceRoot
        
        UtilsLib.LogInfo("DiscoverAssemblyTree: Starting from " & System.IO.Path.GetFileName(rootAsmPath))
        
        Dim asmDoc As AssemblyDocument = Nothing
        Dim wasOpen As Boolean = False
        
        Try
            For Each doc As Document In app.Documents
                If doc.FullFileName.Equals(rootAsmPath, StringComparison.OrdinalIgnoreCase) Then
                    asmDoc = CType(doc, AssemblyDocument)
                    wasOpen = True
                    Exit For
                End If
            Next
            
            If asmDoc Is Nothing Then
                asmDoc = CType(app.Documents.Open(rootAsmPath, False), AssemblyDocument)
            End If
            
            tree.Assemblies.Add(rootAsmPath, New AssemblyInfo With {
                .FilePath = rootAsmPath,
                .RelativePath = GetRelativePath(sourceRoot, rootAsmPath)
            })
            
            ' Force update to ensure all references are resolved
            Try
                asmDoc.Update2(True) ' True = full update including suppressed
            Catch
                ' Update may fail for various reasons, continue anyway
            End Try
            
            ' First pass: AllReferencedDocuments (catches most files)
            For Each refDoc As Document In asmDoc.AllReferencedDocuments
                Dim refPath As String = refDoc.FullFileName
                
                ' Skip files outside source root
                If Not IsInsideSourceRoot(refPath, sourceRoot) Then
                    Continue For
                End If
                
                ' Skip OldVersions folder (special Vault folder)
                If IsOldVersionsPath(refPath) Then
                    UtilsLib.LogInfo("DiscoverAssemblyTree: Skipping OldVersions file: " & System.IO.Path.GetFileName(refPath))
                    Continue For
                End If
                
                Dim ext As String = System.IO.Path.GetExtension(refPath).ToLower()
                
                If ext = ".ipt" Then
                    If Not tree.Parts.ContainsKey(refPath) Then
                        Dim info As PartInfo = ClassifyPart(CType(refDoc, PartDocument), sourceRoot)
                        tree.Parts.Add(refPath, info)
                    End If
                ElseIf ext = ".iam" Then
                    If Not tree.Assemblies.ContainsKey(refPath) Then
                        tree.Assemblies.Add(refPath, New AssemblyInfo With {
                            .FilePath = refPath,
                            .RelativePath = GetRelativePath(sourceRoot, refPath)
                        })
                    End If
                End If
            Next
            
            ' Second pass: Iterate through ReferencedFileDescriptors
            ' This includes ALL referenced files regardless of suppression state
            For i As Integer = 1 To asmDoc.File.ReferencedFileDescriptors.Count
                Try
                    Dim fd As FileDescriptor = asmDoc.File.ReferencedFileDescriptors.Item(i)
                    Dim refPath As String = fd.FullFileName
                    
                    If Not IsInsideSourceRoot(refPath, sourceRoot) Then Continue For
                    If IsOldVersionsPath(refPath) Then Continue For
                    
                    Dim ext As String = System.IO.Path.GetExtension(refPath).ToLower()
                    
                    If ext = ".ipt" AndAlso Not tree.Parts.ContainsKey(refPath) Then
                        Dim partDoc As PartDocument = CType(app.Documents.Open(refPath, False), PartDocument)
                        UtilsLib.LogInfo("DiscoverAssemblyTree: Found via FileDescriptor: " & System.IO.Path.GetFileName(refPath))
                        Dim info As PartInfo = ClassifyPart(partDoc, sourceRoot)
                        tree.Parts.Add(refPath, info)
                    ElseIf ext = ".iam" AndAlso Not tree.Assemblies.ContainsKey(refPath) Then
                        UtilsLib.LogInfo("DiscoverAssemblyTree: Found subassembly via FileDescriptor: " & System.IO.Path.GetFileName(refPath))
                        tree.Assemblies.Add(refPath, New AssemblyInfo With {
                            .FilePath = refPath,
                            .RelativePath = GetRelativePath(sourceRoot, refPath)
                        })
                    End If
                Catch
                End Try
            Next
            
            ' Third pass: Iterate through component occurrences to catch any remaining parts
            ' This handles lazy-loaded parts and parts in nested subassemblies
            DiscoverFromOccurrences(app, asmDoc.ComponentDefinition.Occurrences, tree, sourceRoot)
            
            UtilsLib.LogInfo("DiscoverAssemblyTree: Found " & tree.Parts.Count & " parts, " & tree.Assemblies.Count & " assemblies")
            
            ' ================================================================
            ' Multi-Master Discovery (Phase 1 of multi-master plan)
            ' ================================================================
            If ENABLE_MULTI_MASTER_DISCOVERY Then
                ' Now discover external masters and build dependency graph
                
                ' Get initial master paths from discovered parts
                Dim initialMasters As New List(Of String)
                For Each kvp In tree.Parts
                    If kvp.Value.Role = PartRole.Derived AndAlso Not String.IsNullOrEmpty(kvp.Value.DerivedFromMaster) Then
                        If Not initialMasters.Contains(kvp.Value.DerivedFromMaster) Then
                            initialMasters.Add(kvp.Value.DerivedFromMaster)
                        End If
                    End If
                Next
                
                UtilsLib.LogInfo("DiscoverAssemblyTree: Starting multi-master discovery with " & initialMasters.Count & " initial masters")
                
                ' Phase 1: Recursively discover all masters (including external)
                If ENABLE_DISCOVER_ALL_MASTERS Then
                    DiscoverAllMasters(app, tree, initialMasters)
                    UtilsLib.LogInfo("DiscoverAssemblyTree: DiscoverAllMasters complete")
                Else
                    UtilsLib.LogInfo("DiscoverAssemblyTree: DiscoverAllMasters SKIPPED")
                End If
                
                ' Build set of all discovered masters for analysis
                Dim discoveredMasters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each m In initialMasters
                    discoveredMasters.Add(m)
                Next
                For Each kvp In tree.ExternalMasters
                    discoveredMasters.Add(kvp.Key)
                Next
                For Each kvp In tree.MasterDependencies
                    discoveredMasters.Add(kvp.Key)
                    For Each dep In kvp.Value
                        If System.IO.Path.GetExtension(dep).ToLower() = ".ipt" Then
                            discoveredMasters.Add(dep)
                        End If
                    Next
                Next
                
                Dim allMasters As New List(Of String)(discoveredMasters)
                UtilsLib.LogInfo("DiscoverAssemblyTree: Total masters to check: " & allMasters.Count)
                
                ' Detect projected geometry links - collect unresolved references
                Dim allUnresolved As New List(Of Tuple(Of String, String, String))
                
                ' Phase 2: Check intermediate assemblies for projected geometry
                If ENABLE_DETECT_PROJECTED_GEOM Then
                    For Each intAsm In tree.IntermediateAssemblies.ToList()
                        If Not tree.Assemblies.ContainsKey(intAsm) Then
                            Dim unresolved = DetectProjectedGeometry(app, tree, intAsm, allMasters)
                            allUnresolved.AddRange(unresolved)
                        End If
                    Next
                    UtilsLib.LogInfo("DiscoverAssemblyTree: DetectProjectedGeometry complete")
                Else
                    UtilsLib.LogInfo("DiscoverAssemblyTree: DetectProjectedGeometry SKIPPED")
                End If
                
                ' Phase 3: Search product family for additional intermediate assemblies
                If ENABLE_SEARCH_PRODUCT_FAMILY Then
                    If allUnresolved.Count > 0 OrElse tree.MasterDependencies.Count > 0 Then
                        SearchProductFamilyForIntermediates(app, tree, allUnresolved, discoveredMasters)
                    End If
                    UtilsLib.LogInfo("DiscoverAssemblyTree: SearchProductFamilyForIntermediates complete")
                Else
                    UtilsLib.LogInfo("DiscoverAssemblyTree: SearchProductFamilyForIntermediates SKIPPED")
                End If
                
                ' Classify which intermediate assemblies are actually needed
                ClassifyIntermediateAssemblies(tree)
                
                UtilsLib.LogInfo("DiscoverAssemblyTree: Multi-master discovery complete")
                UtilsLib.LogInfo("  External masters: " & tree.ExternalMasters.Count)
                UtilsLib.LogInfo("  Intermediate assemblies: " & tree.IntermediateAssemblies.Count)
                UtilsLib.LogInfo("  Projected geometry chains: " & tree.ProjectedGeometryChains.Count)
                
                ' Cleanup: Close documents opened during discovery, keep only root assembly
                CleanupDiscoveryDocuments(app, tree.RootAssemblyPath)
            Else
                UtilsLib.LogInfo("DiscoverAssemblyTree: Multi-master discovery DISABLED")
            End If

        Catch ex As Exception
            UtilsLib.LogError("DiscoverAssemblyTree: ERROR - " & ex.Message)
        End Try

        Return tree
    End Function
    
    ''' <summary>
    ''' Close documents opened during discovery, keeping only the root assembly open.
    ''' </summary>
    Private Sub CleanupDiscoveryDocuments(app As Inventor.Application, rootAssemblyPath As String)
        Dim closedCount As Integer = 0
        Dim docsToClose As New List(Of Document)
        
        ' Collect documents to close (can't modify collection while iterating)
        For Each doc As Document In app.Documents
            Try
                Dim docPath As String = doc.FullFileName
                ' Keep the root assembly open
                If Not docPath.Equals(rootAssemblyPath, StringComparison.OrdinalIgnoreCase) Then
                    ' Close assemblies that were opened during discovery
                    If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        docsToClose.Add(doc)
                    End If
                End If
            Catch : End Try
        Next
        
        ' Close collected documents
        For Each doc In docsToClose
            Try
                doc.Close(True) ' True = skip save
                closedCount += 1
            Catch : End Try
        Next
        
        ' Re-activate the root assembly
        Try
            For Each doc As Document In app.Documents
                If doc.FullFileName.Equals(rootAssemblyPath, StringComparison.OrdinalIgnoreCase) Then
                    doc.Activate()
                    Exit For
                End If
            Next
        Catch : End Try
        
        UtilsLib.LogInfo("CleanupDiscoveryDocuments: Closed " & closedCount & " assemblies, keeping root assembly active")
    End Sub
    
    ''' <summary>
    ''' Recursively discover parts from component occurrences.
    ''' This catches lazy-loaded parts that may not appear in AllReferencedDocuments.
    ''' IMPORTANT: We discover ALL referenced files regardless of suppression, visibility,
    ''' or model state - the release must include everything the assembly references.
    ''' </summary>
    Private Sub DiscoverFromOccurrences(app As Inventor.Application, _
                                         occurrences As ComponentOccurrences, _
                                         tree As AssemblyTree, _
                                         sourceRoot As String)
        For Each occ As ComponentOccurrence In occurrences
            Try
                ' DO NOT skip suppressed, invisible, or any other state
                ' We need to discover ALL referenced files for a complete release
                
                ' Get the definition document - this works even for suppressed occurrences
                Dim defDoc As Document = Nothing
                Try
                    defDoc = occ.Definition.Document
                Catch
                    ' For suppressed occurrences, try ReferencedFileDescriptor
                    Try
                        If occ.ReferencedFileDescriptor IsNot Nothing Then
                            Dim refPath2 As String = occ.ReferencedFileDescriptor.FullFileName
                            If IsInsideSourceRoot(refPath2, sourceRoot) AndAlso Not IsOldVersionsPath(refPath2) Then
                                Dim ext2 As String = System.IO.Path.GetExtension(refPath2).ToLower()
                                If ext2 = ".ipt" AndAlso Not tree.Parts.ContainsKey(refPath2) Then
                                    ' Open the document to classify it
                                    Dim partDoc As PartDocument = CType(app.Documents.Open(refPath2, False), PartDocument)
                                    UtilsLib.LogInfo("DiscoverAssemblyTree: Found suppressed part: " & System.IO.Path.GetFileName(refPath2))
                                    Dim info As PartInfo = ClassifyPart(partDoc, sourceRoot)
                                    tree.Parts.Add(refPath2, info)
                                ElseIf ext2 = ".iam" AndAlso Not tree.Assemblies.ContainsKey(refPath2) Then
                                    UtilsLib.LogInfo("DiscoverAssemblyTree: Found suppressed subassembly: " & System.IO.Path.GetFileName(refPath2))
                                    tree.Assemblies.Add(refPath2, New AssemblyInfo With {
                                        .FilePath = refPath2,
                                        .RelativePath = GetRelativePath(sourceRoot, refPath2)
                                    })
                                    ' Also discover its contents
                                    Dim subAsmDoc As AssemblyDocument = CType(app.Documents.Open(refPath2, False), AssemblyDocument)
                                    DiscoverFromOccurrences(app, subAsmDoc.ComponentDefinition.Occurrences, tree, sourceRoot)
                                End If
                            End If
                        End If
                    Catch
                    End Try
                    Continue For
                End Try
                
                If defDoc Is Nothing Then Continue For
                
                Dim refPath As String = defDoc.FullFileName
                
                ' Skip files outside source root
                If Not IsInsideSourceRoot(refPath, sourceRoot) Then
                    Continue For
                End If
                
                ' Skip OldVersions folder
                If IsOldVersionsPath(refPath) Then
                    Continue For
                End If
                
                Dim ext As String = System.IO.Path.GetExtension(refPath).ToLower()
                
                If ext = ".ipt" Then
                    If Not tree.Parts.ContainsKey(refPath) Then
                        UtilsLib.LogInfo("DiscoverAssemblyTree: Found via occurrence: " & System.IO.Path.GetFileName(refPath))
                        Dim info As PartInfo = ClassifyPart(CType(defDoc, PartDocument), sourceRoot)
                        tree.Parts.Add(refPath, info)
                    End If
                ElseIf ext = ".iam" Then
                    If Not tree.Assemblies.ContainsKey(refPath) Then
                        UtilsLib.LogInfo("DiscoverAssemblyTree: Found subassembly via occurrence: " & System.IO.Path.GetFileName(refPath))
                        tree.Assemblies.Add(refPath, New AssemblyInfo With {
                            .FilePath = refPath,
                            .RelativePath = GetRelativePath(sourceRoot, refPath)
                        })
                    End If
                    
                    ' Recursively discover from subassembly's occurrences
                    Dim subAsmDoc As AssemblyDocument = CType(defDoc, AssemblyDocument)
                    DiscoverFromOccurrences(app, subAsmDoc.ComponentDefinition.Occurrences, tree, sourceRoot)
                End If
            Catch ex As Exception
                ' Log but don't skip - try to continue with other occurrences
                UtilsLib.LogWarn("DiscoverAssemblyTree: Error processing occurrence - " & ex.Message)
            End Try
        Next
    End Sub
    
    ''' <summary>
    ''' Classify a part as derived or manual.
    ''' </summary>
    Public Function ClassifyPart(partDoc As PartDocument, sourceRoot As String) As PartInfo
        Dim info As New PartInfo()
        info.FilePath = partDoc.FullFileName
        info.RelativePath = GetRelativePath(sourceRoot, partDoc.FullFileName)
        
        Try
            info.PartNumber = partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
        Catch
            info.PartNumber = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName)
        End Try
        
        Dim dpcs = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
        If dpcs.Count > 0 AndAlso partDoc.ReferencedDocuments.Count > 0 Then
            info.Role = PartRole.Derived
            Try
                info.DerivedFromMaster = partDoc.ReferencedDocuments.Item(1).FullFileName
            Catch
            End Try
            Try
                info.BodyName = GetDerivedBodyName(dpcs.Item(1))
            Catch
            End Try
        Else
            info.Role = PartRole.Manual
        End If
        
        Return info
    End Function
    
    ''' <summary>
    ''' Get the body name from a derived part component.
    ''' </summary>
    Private Function GetDerivedBodyName(dpc As DerivedPartComponent) As String
        Try
            Dim dpDef As Object = dpc.Definition
            For Each dpe As DerivedPartEntity In dpDef.Solids
                If dpe.IncludeEntity Then
                    Dim refEntity As Object = dpe.ReferencedEntity
                    If TypeOf refEntity Is SurfaceBody Then
                        Return CType(refEntity, SurfaceBody).Name
                    End If
                End If
            Next
        Catch
        End Try
        Return ""
    End Function
    
    ''' <summary>
    ''' Discover all drawings that reference the assembly tree.
    ''' </summary>
    Public Function DiscoverDrawings(app As Inventor.Application, _
                                      tree As AssemblyTree, _
                                      searchFolders As List(Of String)) As List(Of DrawingInfo)
        Dim drawings As New List(Of DrawingInfo)
        Dim treeFiles As New HashSet(Of String)(tree.Parts.Keys, StringComparer.OrdinalIgnoreCase)
        treeFiles.UnionWith(tree.Assemblies.Keys)
        
        UtilsLib.LogInfo("DiscoverDrawings: Looking for drawings referencing " & treeFiles.Count & " files")
        
        For Each folder In searchFolders
            If Not System.IO.Directory.Exists(folder) Then Continue For
            
            Try
                For Each idwPath In System.IO.Directory.GetFiles(folder, "*.idw", System.IO.SearchOption.AllDirectories)
                    ' Skip OldVersions folder (special Vault folder)
                    If IsOldVersionsPath(idwPath) Then
                        UtilsLib.LogInfo("DiscoverDrawings: Skipping OldVersions file: " & System.IO.Path.GetFileName(idwPath))
                        Continue For
                    End If
                    
                    Try
                        Dim drawDoc As DrawingDocument = CType(app.Documents.Open(idwPath, False), DrawingDocument)
                        Try
                            Dim refs As New List(Of String)
                            
                            For Each refDoc As Document In drawDoc.ReferencedDocuments
                                If treeFiles.Contains(refDoc.FullFileName) Then
                                    refs.Add(refDoc.FullFileName)
                                End If
                            Next
                            
                            If refs.Count > 0 Then
                                drawings.Add(New DrawingInfo With {
                                    .DrawingPath = idwPath,
                                    .RelativePath = GetRelativePath(tree.SourceRoot, idwPath),
                                    .ReferencedModelPaths = refs
                                })
                                UtilsLib.LogInfo("DiscoverDrawings: Found " & System.IO.Path.GetFileName(idwPath))
                            End If
                        Finally
                            drawDoc.Close(True)
                        End Try
                    Catch ex As Exception
                        UtilsLib.LogWarn("DiscoverDrawings: Error checking " & System.IO.Path.GetFileName(idwPath) & ": " & ex.Message)
                    End Try
                Next
            Catch ex As Exception
                UtilsLib.LogWarn("DiscoverDrawings: Error searching " & folder & ": " & ex.Message)
            End Try
        Next
        
        UtilsLib.LogInfo("DiscoverDrawings: Found " & drawings.Count & " drawings total")
        Return drawings
    End Function
    
    ' ============================================================================
    ' Multi-Master Discovery (Phase 1 of multi-master plan)
    ' ============================================================================
    
    ''' <summary>
    ''' Discover all masters (internal and external) and build dependency graph.
    ''' This is called after basic assembly tree discovery to find external masters
    ''' that are referenced via derivation or projected geometry.
    ''' </summary>
    Public Sub DiscoverAllMasters(app As Inventor.Application, _
                                  tree As AssemblyTree, _
                                  initialMasters As List(Of String))
        UtilsLib.LogInfo("DiscoverAllMasters: Starting with " & initialMasters.Count & " initial masters")
        
        Dim toProcess As New Queue(Of String)
        Dim discoveredMasters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim candidateAssemblies As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase) ' asm -> masters it contains
        
        ' Start with known masters
        For Each m In initialMasters
            toProcess.Enqueue(m)
        Next
        
        ' Process masters recursively
        Do While toProcess.Count > 0
            Dim masterPath As String = toProcess.Dequeue()
            If discoveredMasters.Contains(masterPath) Then Continue Do
            discoveredMasters.Add(masterPath)
            
            Dim isExternal As Boolean = Not IsInsideSourceRoot(masterPath, tree.SourceRoot)
            If isExternal Then
                UtilsLib.LogInfo("DiscoverAllMasters: Found EXTERNAL master: " & masterPath)
            End If
            
            ' Get dependencies of this master
            Dim deps As New List(Of String)
            Dim masterDoc As Document = Nothing
            Dim weOpenedIt As Boolean = False
            Try
                ' Check if document is already open
                For Each doc As Document In app.Documents
                    If doc.FullFileName.Equals(masterPath, StringComparison.OrdinalIgnoreCase) Then
                        masterDoc = doc
                        Exit For
                    End If
                Next
                
                ' Only open if not already open
                If masterDoc Is Nothing Then
                    masterDoc = app.Documents.Open(masterPath, False)
                    weOpenedIt = True
                End If
                
                If masterDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim partMaster As PartDocument = CType(masterDoc, PartDocument)
                    Dim refComps = partMaster.ComponentDefinition.ReferenceComponents
                    
                    ' DerivedPartComponents - derives from another part (or assembly)
                    For Each dpc As DerivedPartComponent In refComps.DerivedPartComponents
                        Try
                            Dim refFile As String = dpc.ReferencedFile.FullFileName
                            deps.Add(refFile)
                            
                            ' Check if derivation is via an assembly
                            Try
                                Dim dpDef As Object = dpc.Definition
                                Dim fullDocName As String = CStr(CallByName(dpDef, "FullDocumentName", CallType.Get))
                                If fullDocName.ToLower().Contains(".iam") Then
                                    Dim asmPath As String = ExtractAssemblyPathFromFullDocName(fullDocName)
                                    If Not String.IsNullOrEmpty(asmPath) Then
                                        tree.IntermediateAssemblies.Add(asmPath)
                                        If Not deps.Contains(asmPath) Then deps.Add(asmPath)
                                        
                                        ' Track which masters this assembly contains
                                        If Not candidateAssemblies.ContainsKey(asmPath) Then
                                            candidateAssemblies(asmPath) = New List(Of String)
                                        End If
                                        If Not candidateAssemblies(asmPath).Contains(masterPath) Then
                                            candidateAssemblies(asmPath).Add(masterPath)
                                        End If
                                    End If
                                End If
                            Catch
                            End Try
                            
                            ' Queue for processing if not already discovered
                            If Not discoveredMasters.Contains(refFile) AndAlso Not toProcess.Contains(refFile) Then
                                toProcess.Enqueue(refFile)
                            End If
                        Catch
                        End Try
                    Next
                    
                    ' DerivedAssemblyComponents - derives from an assembly (e.g., for projected geometry)
                    For Each dac As DerivedAssemblyComponent In refComps.DerivedAssemblyComponents
                        Try
                            Dim refFile As String = dac.ReferencedFile.FullFileName
                            deps.Add(refFile)
                            tree.IntermediateAssemblies.Add(refFile)
                            
                            If Not discoveredMasters.Contains(refFile) AndAlso Not toProcess.Contains(refFile) Then
                                toProcess.Enqueue(refFile)
                            End If
                        Catch
                        End Try
                    Next
                    
                    ' CRITICAL: Check ReferencedDocuments for assemblies
                    ' Masters with projected geometry will reference the assembly used for projection
                    For Each refDoc As Document In partMaster.ReferencedDocuments
                        Try
                            Dim refPath As String = refDoc.FullFileName
                            If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                                tree.IntermediateAssemblies.Add(refPath)
                                UtilsLib.LogInfo("DiscoverAllMasters: Found intermediate assembly via ReferencedDocuments: " & System.IO.Path.GetFileName(refPath))
                                
                                If Not deps.Contains(refPath) Then deps.Add(refPath)
                                If Not discoveredMasters.Contains(refPath) AndAlso Not toProcess.Contains(refPath) Then
                                    toProcess.Enqueue(refPath)
                                End If
                            ElseIf refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                ' Could be another master
                                If Not deps.Contains(refPath) Then deps.Add(refPath)
                                If Not discoveredMasters.Contains(refPath) AndAlso Not toProcess.Contains(refPath) Then
                                    toProcess.Enqueue(refPath)
                                End If
                            End If
                        Catch
                        End Try
                    Next
                    
                    ' Also check ReferencedFileDescriptors for context info (FullDocumentName)
                    Try
                        For i As Integer = 1 To partMaster.File.ReferencedFileDescriptors.Count
                            Try
                                Dim fd As FileDescriptor = partMaster.File.ReferencedFileDescriptors.Item(i)
                                Dim refPath As String = fd.FullFileName
                                
                                ' Check FullDocumentName for assembly context
                                Dim fullDocName As String = ""
                                Try : fullDocName = fd.FullDocumentName : Catch : End Try
                                
                                If Not String.IsNullOrEmpty(fullDocName) AndAlso fullDocName.ToLower().Contains(".iam") Then
                                    Dim asmPath As String = ExtractAssemblyPathFromFullDocName(fullDocName)
                                    If Not String.IsNullOrEmpty(asmPath) Then
                                        tree.IntermediateAssemblies.Add(asmPath)
                                        UtilsLib.LogInfo("DiscoverAllMasters: Found intermediate assembly via FullDocumentName: " & System.IO.Path.GetFileName(asmPath))
                                        
                                        If Not deps.Contains(asmPath) Then deps.Add(asmPath)
                                        If Not discoveredMasters.Contains(asmPath) AndAlso Not toProcess.Contains(asmPath) Then
                                            toProcess.Enqueue(asmPath)
                                        End If
                                    End If
                                End If
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                    
                    ' CRITICAL: Check ReferencingFiles - assemblies that CONTAIN this master
                    ' This finds intermediate assemblies used for projected geometry
                    Try
                        Dim referencingFiles As Object = partMaster.File.ReferencingFiles
                        If referencingFiles IsNot Nothing Then
                            For Each rf As Object In referencingFiles
                                Try
                                    Dim rfPath As String = CStr(CallByName(rf, "FullFileName", CallType.Get))
                                    
                                    ' Skip old Vault versions (OldVersions folder or versioned filenames)
                                    If IsVaultOldVersion(rfPath) Then Continue For
                                    
                                    If rfPath.ToLower().EndsWith(".iam") Then
                                        tree.IntermediateAssemblies.Add(rfPath)
                                        UtilsLib.LogInfo("DiscoverAllMasters: Found containing assembly via ReferencingFiles: " & System.IO.Path.GetFileName(rfPath))
                                        
                                        If Not candidateAssemblies.ContainsKey(rfPath) Then
                                            candidateAssemblies(rfPath) = New List(Of String)
                                        End If
                                        candidateAssemblies(rfPath).Add(masterPath)
                                        
                                        If Not discoveredMasters.Contains(rfPath) AndAlso Not toProcess.Contains(rfPath) Then
                                            toProcess.Enqueue(rfPath)
                                        End If
                                    End If
                                Catch
                                End Try
                            Next
                        End If
                    Catch ex As Exception
                        UtilsLib.LogWarn("DiscoverAllMasters: ReferencingFiles check failed for " & System.IO.Path.GetFileName(masterPath) & ": " & ex.Message)
                    End Try
                ElseIf masterDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim asmMaster As AssemblyDocument = CType(masterDoc, AssemblyDocument)
                    
                    ' Collect parts in this assembly
                    For Each occ As ComponentOccurrence In asmMaster.ComponentDefinition.Occurrences.AllLeafOccurrences
                        Try
                            Dim occPath As String = occ.Definition.Document.FullFileName
                            If System.IO.Path.GetExtension(occPath).ToLower() = ".ipt" Then
                                If Not candidateAssemblies.ContainsKey(masterPath) Then
                                    candidateAssemblies(masterPath) = New List(Of String)
                                End If
                                candidateAssemblies(masterPath).Add(occPath)
                            End If
                        Catch
                        End Try
                    Next
                End If
            Catch ex As Exception
                UtilsLib.LogWarn("DiscoverAllMasters: Error processing " & System.IO.Path.GetFileName(masterPath) & ": " & ex.Message)
            Finally
                ' Only close if WE opened it (don't close documents that were already open)
                If weOpenedIt AndAlso masterDoc IsNot Nothing Then
                    Try : masterDoc.Close(True) : Catch : End Try
                End If
            End Try
            
            ' Store dependencies
            If deps.Count > 0 Then
                tree.MasterDependencies(masterPath) = deps
            End If
            
            ' Store external master info
            If isExternal Then
                Dim info As New MasterInfo()
                info.FilePath = masterPath
                info.RelativePath = System.IO.Path.GetFileName(masterPath)
                info.SourceElement = ExtractElementName(masterPath)
                info.DependsOn = deps
                info.IsIntermediate = System.IO.Path.GetExtension(masterPath).ToLower() = ".iam"
                tree.ExternalMasters(masterPath) = info
            End If
        Loop
        
        UtilsLib.LogInfo("DiscoverAllMasters: Discovered " & discoveredMasters.Count & " total masters")
        UtilsLib.LogInfo("DiscoverAllMasters: " & tree.ExternalMasters.Count & " external masters")
        UtilsLib.LogInfo("DiscoverAllMasters: " & tree.IntermediateAssemblies.Count & " intermediate assemblies")
    End Sub
    
    ''' <summary>
    ''' Detect projected geometry links by parsing browser node labels.
    ''' Must be called with the assembly open and active.
    ''' Returns list of unresolved references (OccName, PartNum, DependentMasterPath) for further search.
    ''' 
    ''' IMPORTANT: Following test script pattern - documents are NOT closed during analysis.
    ''' This prevents crashes from closing documents that may still be needed.
    ''' </summary>
    Public Function DetectProjectedGeometry(app As Inventor.Application, _
                                            tree As AssemblyTree, _
                                            assemblyPath As String, _
                                            mastersToCheck As List(Of String)) As List(Of Tuple(Of String, String, String))
        Dim unresolvedOccurrences As New List(Of Tuple(Of String, String, String))
        Dim asmDoc As AssemblyDocument = Nothing
        
        Try
            ' Check if document is already open
            For Each doc As Document In app.Documents
                Try
                    If doc.FullFileName.Equals(assemblyPath, StringComparison.OrdinalIgnoreCase) Then
                        asmDoc = CType(doc, AssemblyDocument)
                        Exit For
                    End If
                Catch : End Try
            Next
            
            ' Open visibly if not already open (following test script pattern)
            If asmDoc Is Nothing Then
                asmDoc = CType(app.Documents.Open(assemblyPath, True), AssemblyDocument)
            End If
            
            ' Activate for browser node access
            asmDoc.Activate()
            Try : app.ActiveView.Fit() : Catch : End Try
            
            UtilsLib.LogInfo("DetectProjectedGeometry: Analyzing " & System.IO.Path.GetFileName(assemblyPath))
            
            ' Build occurrence maps
            Dim occByName As New Dictionary(Of String, ComponentOccurrence)(StringComparer.OrdinalIgnoreCase)
            Dim occByPath As New Dictionary(Of String, List(Of ComponentOccurrence))(StringComparer.OrdinalIgnoreCase)
            
            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences.AllLeafOccurrences
                Try
                    Dim occPath As String = occ.Definition.Document.FullFileName
                    occByName(occ.Name) = occ
                    
                    If Not occByPath.ContainsKey(occPath) Then
                        occByPath(occPath) = New List(Of ComponentOccurrence)
                    End If
                    occByPath(occPath).Add(occ)
                Catch
                End Try
            Next
            
            ' Check each master that's present in this assembly
            For Each masterPath In mastersToCheck
                If Not occByPath.ContainsKey(masterPath) Then Continue For
                
                Dim masterOccs As List(Of ComponentOccurrence) = occByPath(masterPath)
                For Each masterOcc In masterOccs
                    Try
                        Dim partDef As PartComponentDefinition = CType(masterOcc.Definition, PartComponentDefinition)
                        
                        ' Find the browser node for this occurrence's sketches
                        For Each sk As PlanarSketch In partDef.Sketches
                            Try
                                ' Get sketch browser node via the active document's browser pane
                                Dim sketchNode As BrowserNode = Nothing
                                Try
                                    For Each pane As BrowserPane In app.ActiveDocument.BrowserPanes
                                        Try
                                            sketchNode = pane.GetBrowserNodeFromObject(sk)
                                            If sketchNode IsNot Nothing Then Exit For
                                        Catch : End Try
                                    Next
                                Catch : End Try
                                
                                If sketchNode Is Nothing Then Continue For
                                
                                ' Check child nodes for projected geometry references
                                For Each childNode As BrowserNode In sketchNode.BrowserNodes
                                    Try
                                        Dim label As String = ""
                                        Try : label = childNode.BrowserNodeDefinition.Label : Catch : End Try
                                        If String.IsNullOrEmpty(label) Then Continue For
                                        
                                        ' Parse "ReferenceXX (OccurrenceName)" format
                                        If Not label.StartsWith("Reference") OrElse Not label.Contains("(") Then Continue For
                                        
                                        Dim parenStart As Integer = label.IndexOf("(")
                                        Dim parenEnd As Integer = label.LastIndexOf(")")
                                        If parenEnd <= parenStart Then Continue For
                                        
                                        Dim occNameFromLabel As String = label.Substring(parenStart + 1, parenEnd - parenStart - 1)
                                        
                                        ' Try to find this occurrence in the assembly
                                        If occByName.ContainsKey(occNameFromLabel) Then
                                            ' RESOLVED - occurrence exists in this assembly
                                            Dim refOcc As ComponentOccurrence = occByName(occNameFromLabel)
                                            Dim refPath As String = refOcc.Definition.Document.FullFileName
                                            
                                            ' Check if it's a cross-master reference
                                            If mastersToCheck.Any(Function(m) m.Equals(refPath, StringComparison.OrdinalIgnoreCase)) AndAlso _
                                               Not refPath.Equals(masterPath, StringComparison.OrdinalIgnoreCase) Then
                                                ' Found a cross-master projected geometry chain!
                                                Dim chain As New Tuple(Of String, String, String)(refPath, assemblyPath, masterPath)
                                                
                                                Dim isDuplicate As Boolean = tree.ProjectedGeometryChains.Any(Function(c) _
                                                    c.Item1.Equals(chain.Item1, StringComparison.OrdinalIgnoreCase) AndAlso _
                                                    c.Item2.Equals(chain.Item2, StringComparison.OrdinalIgnoreCase) AndAlso _
                                                    c.Item3.Equals(chain.Item3, StringComparison.OrdinalIgnoreCase))
                                                
                                                If Not isDuplicate Then
                                                    tree.ProjectedGeometryChains.Add(chain)
                                                    tree.IntermediateAssemblies.Add(assemblyPath)
                                                    UtilsLib.LogInfo("DetectProjectedGeometry: Found chain: " & _
                                                        System.IO.Path.GetFileName(refPath) & " --(" & _
                                                        System.IO.Path.GetFileName(assemblyPath) & ")--> " & _
                                                        System.IO.Path.GetFileName(masterPath))
                                                End If
                                            End If
                                        Else
                                            ' UNRESOLVED - occurrence not in this assembly
                                            ' Extract part number for later search
                                            Dim partNum As String = ""
                                            If occNameFromLabel.Contains("(") Then
                                                Dim innerStart As Integer = occNameFromLabel.LastIndexOf("(")
                                                Dim innerEnd As Integer = occNameFromLabel.IndexOf(")", innerStart)
                                                If innerEnd > innerStart Then
                                                    partNum = occNameFromLabel.Substring(innerStart + 1, innerEnd - innerStart - 1)
                                                End If
                                            ElseIf occNameFromLabel.Contains(":") Then
                                                partNum = occNameFromLabel.Split(":"c)(0)
                                            Else
                                                partNum = occNameFromLabel
                                            End If
                                            
                                            ' Record for later search (if not already recorded)
                                            Dim unresolvedEntry As New Tuple(Of String, String, String)(occNameFromLabel, partNum, masterPath)
                                            If Not unresolvedOccurrences.Any(Function(u) u.Item1 = occNameFromLabel AndAlso u.Item3.Equals(masterPath, StringComparison.OrdinalIgnoreCase)) Then
                                                unresolvedOccurrences.Add(unresolvedEntry)
                                            End If
                                        End If
                                    Catch
                                    End Try
                                Next
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                Next
            Next
            
        Catch ex As Exception
            UtilsLib.LogWarn("DetectProjectedGeometry: Error analyzing " & System.IO.Path.GetFileName(assemblyPath) & ": " & ex.Message)
        End Try
        ' NOTE: Following test script pattern - we do NOT close the document here
        ' The test script leaves documents open during analysis and never crashes
        
        Return unresolvedOccurrences
    End Function
    
    ''' <summary>
    ''' Classify which assemblies are actually needed for the release.
    ''' An assembly is needed if it has actual projected geometry (from ProjectedGeometryChains).
    ''' </summary>
    Public Sub ClassifyIntermediateAssemblies(tree As AssemblyTree)
        UtilsLib.LogInfo("ClassifyIntermediateAssemblies: Analyzing " & tree.IntermediateAssemblies.Count & " candidates")
        
        ' Assemblies that have actual projected geometry are definitely needed
        Dim neededAssemblies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each chain In tree.ProjectedGeometryChains
            neededAssemblies.Add(chain.Item2) ' ViaAssembly
        Next
        
        ' Filter IntermediateAssemblies to only those that are actually needed
        Dim toRemove As New List(Of String)
        For Each asmPath In tree.IntermediateAssemblies
            If Not neededAssemblies.Contains(asmPath) Then
                toRemove.Add(asmPath)
                UtilsLib.LogInfo("ClassifyIntermediateAssemblies: Removing (not needed): " & System.IO.Path.GetFileName(asmPath))
            End If
        Next
        
        For Each asmPath In toRemove
            tree.IntermediateAssemblies.Remove(asmPath)
        Next
        
        UtilsLib.LogInfo("ClassifyIntermediateAssemblies: " & tree.IntermediateAssemblies.Count & " assemblies remain needed")
    End Sub
    
    ''' <summary>
    ''' Search product family folder for assemblies that resolve unresolved projected geometry references.
    ''' This finds intermediate assemblies that weren't directly referenced but contain the necessary masters.
    ''' 
    ''' IMPORTANT: This function mirrors the logic from TestMasterDependencies.vb EXACTLY to avoid crashes.
    ''' Key patterns from the test script that MUST be preserved:
    ''' 1. Open documents invisibly first for containment check
    ''' 2. Activate (not close-reopen) when browser access is needed
    ''' 3. Do NOT close documents during the loop - the test script leaves them open
    ''' 4. Use occurrence browser node approach for reference detection
    ''' </summary>
    Public Sub SearchProductFamilyForIntermediates(app As Inventor.Application, _
                                                   tree As AssemblyTree, _
                                                   unresolvedOccurrences As List(Of Tuple(Of String, String, String)), _
                                                   discoveredMasters As HashSet(Of String))
        If unresolvedOccurrences.Count = 0 AndAlso tree.MasterDependencies.Count = 0 Then Return
        
        ' Build dependency pairs from unresolved occurrences AND derivation chains
        Dim dependencyPairs As New List(Of Tuple(Of String, String)) ' (DependentMaster, SourceMaster)
        
        ' Add from unresolved occurrences
        For Each unresolved In unresolvedOccurrences
            Dim partNum As String = unresolved.Item2
            Dim dependentMaster As String = unresolved.Item3
            
            ' Find the source master by part number
            For Each masterPath In discoveredMasters
                If System.IO.Path.GetFileNameWithoutExtension(masterPath) = partNum Then
                    Dim pair As New Tuple(Of String, String)(dependentMaster, masterPath)
                    If Not dependencyPairs.Any(Function(p) p.Item1.Equals(pair.Item1, StringComparison.OrdinalIgnoreCase) AndAlso p.Item2.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase)) Then
                        dependencyPairs.Add(pair)
                    End If
                    Exit For
                End If
            Next
        Next
        
        ' Add from derivation chains (as fallback when browser node analysis doesn't resolve)
        For Each kvp In tree.MasterDependencies
            Dim dependentMaster As String = kvp.Key
            For Each sourceMaster In kvp.Value
                If System.IO.Path.GetExtension(sourceMaster).ToLower() = ".ipt" AndAlso _
                   discoveredMasters.Contains(sourceMaster) Then
                    Dim pair As New Tuple(Of String, String)(dependentMaster, sourceMaster)
                    If Not dependencyPairs.Any(Function(p) p.Item1.Equals(pair.Item1, StringComparison.OrdinalIgnoreCase) AndAlso p.Item2.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase)) Then
                        dependencyPairs.Add(pair)
                    End If
                End If
            Next
        Next
        
        If dependencyPairs.Count = 0 Then Return
        
        UtilsLib.LogInfo("SearchProductFamilyForIntermediates: Searching for " & dependencyPairs.Count & " dependency pairs")
        
        ' Get product family root folder
        Dim productFamilyRoot As String = GetProductFamilyRoot(tree.SourceRoot)
        UtilsLib.LogInfo("SearchProductFamilyForIntermediates: Product family root: " & productFamilyRoot)
        
        ' Find all assemblies in the product family
        Dim allAssemblies As String() = Nothing
        Try
            allAssemblies = System.IO.Directory.GetFiles(productFamilyRoot, "*.iam", System.IO.SearchOption.AllDirectories)
            UtilsLib.LogInfo("SearchProductFamilyForIntermediates: Scanning " & allAssemblies.Length & " assemblies...")
        Catch ex As Exception
            UtilsLib.LogWarn("SearchProductFamilyForIntermediates: Error scanning folder: " & ex.Message)
            Return
        End Try
        
        Dim currentAsmPath As String = tree.RootAssemblyPath
        Dim testedCount As Integer = 0
        
        ' NOTE: Following test script pattern - we do NOT close documents during the loop
        ' This is critical for stability
        
        For Each iamFile In allAssemblies
            ' Skip already processed or invalid (matches test script exactly)
            If iamFile.Equals(currentAsmPath, StringComparison.OrdinalIgnoreCase) Then Continue For
            If tree.IntermediateAssemblies.Contains(iamFile) Then Continue For
            If IsVaultOldVersion(iamFile) Then Continue For
            
            Try
                ' First, quick check: does this assembly contain BOTH masters of any dependency pair?
                ' IMPORTANT: Must open VISIBLY - invisible open + activate does not work for browser access
                Dim extAsm As AssemblyDocument = CType(app.Documents.Open(iamFile, True), AssemblyDocument)
                Dim containedFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each refDoc As Document In extAsm.AllReferencedDocuments
                    containedFiles.Add(refDoc.FullFileName)
                Next
                
                ' Check each dependency pair
                For Each pair In dependencyPairs
                    ' Skip if this pair is already resolved
                    If tree.ProjectedGeometryChains.Any(Function(c) c.Item1.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase) AndAlso c.Item3.Equals(pair.Item1, StringComparison.OrdinalIgnoreCase)) Then
                        Continue For
                    End If
                    
                    ' Does assembly contain BOTH the dependent and source master?
                    Dim hasDep As Boolean = containedFiles.Contains(pair.Item1)
                    Dim hasSrc As Boolean = containedFiles.Contains(pair.Item2)
                    If Not hasDep OrElse Not hasSrc Then
                        ' Log why we're skipping (only for first few to avoid spam)
                        If testedCount < 3 Then
                            UtilsLib.LogInfo("  Skipping " & System.IO.Path.GetFileName(iamFile) & ": hasDep=" & hasDep & " hasSrc=" & hasSrc)
                        End If
                        Continue For
                    End If
                    
                    testedCount += 1
                    UtilsLib.LogInfo("  Testing: " & System.IO.Path.GetFileName(iamFile))
                    UtilsLib.LogInfo("    Dependent: " & pair.Item1)
                    UtilsLib.LogInfo("    Source: " & pair.Item2)
                    
                    Try
                        ' Activate assembly for browser node access (matches test script)
                        ' Must activate to access browser panes
                        extAsm.Activate()
                        app.ActiveView.Fit()
                    Catch actEx As Exception
                        UtilsLib.LogWarn("    Failed to activate: " & actEx.Message)
                        Continue For
                    End Try
                    
                    ' Build occurrence map (matches test script)
                    Dim occByName As New Dictionary(Of String, ComponentOccurrence)(StringComparer.OrdinalIgnoreCase)
                    Try
                        For Each occ As ComponentOccurrence In extAsm.ComponentDefinition.Occurrences.AllLeafOccurrences
                            Try
                                occByName(occ.Name) = occ
                            Catch : End Try
                        Next
                    Catch occEx As Exception
                        UtilsLib.LogWarn("    Failed to get occurrences: " & occEx.Message)
                        Continue For
                    End Try
                    UtilsLib.LogInfo("    Occurrences in assembly: " & occByName.Count)
                    
                    ' Find the dependent master occurrence (matches test script - full path match)
                    Dim dependentOcc As ComponentOccurrence = Nothing
                    For Each kvp2 In occByName
                        Try
                            If kvp2.Value.Definition.Document.FullFileName.Equals(pair.Item1, StringComparison.OrdinalIgnoreCase) Then
                                dependentOcc = kvp2.Value
                                UtilsLib.LogInfo("    Found dependent occurrence: " & kvp2.Key)
                                Exit For
                            End If
                        Catch : End Try
                    Next
                    
                    If dependentOcc Is Nothing Then
                        UtilsLib.LogInfo("    Dependent master not found as occurrence")
                        Continue For
                    End If
                    
                    ' Check if projected geometry references RESOLVE in this assembly
                    Dim partDef As PartComponentDefinition = Nothing
                    Try
                        partDef = CType(dependentOcc.Definition, PartComponentDefinition)
                    Catch castEx As Exception
                        UtilsLib.LogWarn("    Failed to get PartComponentDefinition: " & castEx.Message)
                        Continue For
                    End Try
                    
                    Dim referencesResolved As Boolean = False
                    Dim sourceOccName As String = ""
                    
                    ' Find the browser node for this occurrence (matches test script approach)
                    ' In assembly context, the part's sketches are under the occurrence node
                    ' Match by occurrence name since COM object comparison with Is may not work
                    Dim occBrowserNode As BrowserNode = Nothing
                    Dim dependentOccName As String = dependentOcc.Name
                    UtilsLib.LogInfo("    Looking for browser node of: " & dependentOccName)
                    Try
                        For Each pane As BrowserPane In app.ActiveDocument.BrowserPanes
                            Try
                                For Each topNode As BrowserNode In pane.TopNode.BrowserNodes
                                    Try
                                        Dim nodeLabel As String = topNode.BrowserNodeDefinition.Label
                                        ' Match occurrence name (e.g., "Nurk - Eskiis (000131):1")
                                        If nodeLabel.Equals(dependentOccName, StringComparison.OrdinalIgnoreCase) Then
                                            occBrowserNode = topNode
                                            Exit For
                                        End If
                                    Catch : End Try
                                Next
                            Catch : End Try
                            If occBrowserNode IsNot Nothing Then Exit For
                        Next
                    Catch : End Try
                    
                    If occBrowserNode Is Nothing Then
                        UtilsLib.LogInfo("    Occurrence browser node not found - trying sketch approach")
                    Else
                        UtilsLib.LogInfo("    Found occurrence browser node: " & occBrowserNode.BrowserNodeDefinition.Label)
                    End If
                    
                    ' Check sketches for browser nodes with cross-part references
                    ' Method 1: Look through the occurrence's browser nodes for sketches (matches test script)
                    Dim sketchesChecked As Integer = 0
                    
                    If occBrowserNode IsNot Nothing Then
                        ' Recursively find sketch nodes under the occurrence
                        For Each childNode As BrowserNode In occBrowserNode.BrowserNodes
                            Try
                                Dim childLabel As String = childNode.BrowserNodeDefinition.Label
                                
                                ' Check if this node has reference children
                                For Each refNode As BrowserNode In childNode.BrowserNodes
                                    Try
                                        Dim refLabel As String = refNode.BrowserNodeDefinition.Label
                                        If refLabel.StartsWith("Reference") AndAlso refLabel.Contains("(") Then
                                            sketchesChecked += 1
                                            
                                            ' Extract occurrence name from label
                                            Dim parenStart As Integer = refLabel.IndexOf("(")
                                            Dim parenEnd As Integer = refLabel.LastIndexOf(")")
                                            If parenEnd > parenStart Then
                                                Dim occNameFromLabel As String = refLabel.Substring(parenStart + 1, parenEnd - parenStart - 1)
                                                
                                                ' Does this occurrence exist in the assembly?
                                                If occByName.ContainsKey(occNameFromLabel) Then
                                                    Dim refOcc As ComponentOccurrence = occByName(occNameFromLabel)
                                                    Dim refPath As String = refOcc.Definition.Document.FullFileName
                                                    
                                                    ' Does it point to our source master?
                                                    If refPath.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase) Then
                                                        referencesResolved = True
                                                        sourceOccName = occNameFromLabel
                                                        UtilsLib.LogInfo("    ** REFERENCE RESOLVED: " & refLabel & " -> " & occNameFromLabel)
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Catch : End Try
                                Next
                                
                                If referencesResolved Then Exit For
                            Catch : End Try
                        Next
                    End If
                    
                    ' Method 2: If no occurrence browser node, try the old approach with sketches (matches test script)
                    If Not referencesResolved AndAlso occBrowserNode Is Nothing Then
                        For Each sk As PlanarSketch In partDef.Sketches
                            Try
                                ' Try to find this sketch's browser node by matching name
                                Dim sketchName As String = sk.Name
                                Dim sketchNode As BrowserNode = Nothing
                                
                                Try
                                    For Each pane As BrowserPane In app.ActiveDocument.BrowserPanes
                                        Try
                                            For Each topNode As BrowserNode In pane.TopNode.BrowserNodes
                                                Try
                                                    Dim nodeLabel As String = topNode.BrowserNodeDefinition.Label
                                                    If nodeLabel.Contains(sketchName) Then
                                                        sketchNode = topNode
                                                        Exit For
                                                    End If
                                                Catch : End Try
                                            Next
                                        Catch : End Try
                                        If sketchNode IsNot Nothing Then Exit For
                                    Next
                                Catch : End Try
                                
                                If sketchNode Is Nothing Then Continue For
                                sketchesChecked += 1
                                
                                ' Check child nodes for reference labels
                                For Each childNode As BrowserNode In sketchNode.BrowserNodes
                                    Try
                                        Dim label As String = childNode.BrowserNodeDefinition.Label
                                        If label.StartsWith("Reference") AndAlso label.Contains("(") Then
                                            Dim parenStart As Integer = label.IndexOf("(")
                                            Dim parenEnd As Integer = label.LastIndexOf(")")
                                            If parenEnd > parenStart Then
                                                Dim occNameFromLabel As String = label.Substring(parenStart + 1, parenEnd - parenStart - 1)
                                                
                                                If occByName.ContainsKey(occNameFromLabel) Then
                                                    Dim refOcc As ComponentOccurrence = occByName(occNameFromLabel)
                                                    Dim refPath As String = refOcc.Definition.Document.FullFileName
                                                    
                                                    If refPath.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase) Then
                                                        referencesResolved = True
                                                        sourceOccName = occNameFromLabel
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Catch : End Try
                                Next
                                
                                If referencesResolved Then Exit For
                            Catch : End Try
                        Next
                    End If
                    
                    UtilsLib.LogInfo("    Sketches/nodes checked: " & sketchesChecked)
                    
                    If referencesResolved Then
                        UtilsLib.LogInfo("    ** VERIFIED: References resolve in this assembly!")
                        UtilsLib.LogInfo("       Source occurrence: " & sourceOccName)
                        tree.IntermediateAssemblies.Add(iamFile)
                        
                        Dim chain As New Tuple(Of String, String, String)(pair.Item2, iamFile, pair.Item1)
                        If Not tree.ProjectedGeometryChains.Any(Function(c) c.Item1.Equals(chain.Item1, StringComparison.OrdinalIgnoreCase) AndAlso c.Item2.Equals(chain.Item2, StringComparison.OrdinalIgnoreCase) AndAlso c.Item3.Equals(chain.Item3, StringComparison.OrdinalIgnoreCase)) Then
                            tree.ProjectedGeometryChains.Add(chain)
                        End If
                        UtilsLib.LogInfo("       *** CHAIN: " & System.IO.Path.GetFileName(pair.Item2) & " --(via " & System.IO.Path.GetFileName(iamFile) & ")--> " & System.IO.Path.GetFileName(pair.Item1))
                    Else
                        UtilsLib.LogInfo("    (references do not resolve)")
                    End If
                Next
            Catch openEx As Exception
                UtilsLib.LogWarn("  Error testing " & System.IO.Path.GetFileName(iamFile) & ": " & openEx.Message)
            End Try
            ' NOTE: Following test script pattern - we do NOT close extAsm here
            ' The test script leaves documents open and never crashes
        Next
        
        UtilsLib.LogInfo("SearchProductFamilyForIntermediates: Tested " & testedCount & " candidate assemblies")
    End Sub
    
    ''' <summary>
    ''' Extract assembly path from FullDocumentName format.
    ''' Format: "C:\path\to\file.ipt|C:\path\to\assembly.iam"
    ''' </summary>
    Private Function ExtractAssemblyPathFromFullDocName(fullDocName As String) As String
        If fullDocName.Contains("|") Then
            Dim parts() As String = fullDocName.Split("|"c)
            For Each part In parts
                If part.ToLower().EndsWith(".iam") Then
                    Return part.Trim()
                End If
            Next
        End If
        Return ""
    End Function
    
    ''' <summary>
    ''' Extract element name from a file path (assumes Aluselemendid folder structure).
    ''' </summary>
    Private Function ExtractElementName(filePath As String) As String
        Try
            Dim dir As String = System.IO.Path.GetDirectoryName(filePath)
            Dim parts() As String = dir.Split(System.IO.Path.DirectorySeparatorChar)
            
            For i As Integer = parts.Length - 1 To 0 Step -1
                If parts(i).Equals("Aluselemendid", StringComparison.OrdinalIgnoreCase) Then
                    If i + 1 < parts.Length Then
                        Return parts(i + 1)
                    End If
                End If
            Next
        Catch
        End Try
        Return ""
    End Function
    
    ' ============================================================================
    ' Phase 3: Variant Analysis
    ' ============================================================================
    
    ''' <summary>
    ''' Compute geometry fingerprint for a part document.
    ''' Uses geometry-only hash (no source path) for intra-module comparison.
    ''' </summary>
    Public Function ComputeGeometryFingerprint(partDoc As PartDocument) As String
        Dim bodies = partDoc.ComponentDefinition.SurfaceBodies
        If bodies.Count = 0 Then Return "NO_BODIES"
        
        Dim fps As New List(Of String)
        For Each body As SurfaceBody In bodies
            If body.IsSolid Then
                fps.Add(ComputeBodyFingerprint(body))
            End If
        Next
        
        fps.Sort()
        Return String.Join("|", fps.ToArray())
    End Function
    
    ''' <summary>
    ''' Compute fingerprint for a single solid body.
    ''' </summary>
    Public Function ComputeBodyFingerprint(body As SurfaceBody) As String
        Try
            Dim tol As Double = 0.001
            
            Dim vol As Double = 0
            Try : vol = Math.Round(body.Volume(tol), 4) : Catch : End Try
            
            Dim area As Double = 0
            Try
                For Each face As Face In body.Faces
                    area += face.Evaluator.Area
                Next
                area = Math.Round(area, 4)
            Catch : End Try
            
            Dim bb As Box = body.RangeBox
            Dim dims() As Double = {
                Math.Round(bb.MaxPoint.X - bb.MinPoint.X, 3),
                Math.Round(bb.MaxPoint.Y - bb.MinPoint.Y, 3),
                Math.Round(bb.MaxPoint.Z - bb.MinPoint.Z, 3)
            }
            Array.Sort(dims)
            
            Return String.Format("V:{0}|A:{1}|BB:{2}x{3}x{4}",
                vol.ToString("F4", System.Globalization.CultureInfo.InvariantCulture),
                area.ToString("F4", System.Globalization.CultureInfo.InvariantCulture),
                dims(0).ToString("F3", System.Globalization.CultureInfo.InvariantCulture),
                dims(1).ToString("F3", System.Globalization.CultureInfo.InvariantCulture),
                dims(2).ToString("F3", System.Globalization.CultureInfo.InvariantCulture))
        Catch ex As Exception
            Return "ERROR:" & ex.Message
        End Try
    End Function
    
    ''' <summary>
    ''' Full fingerprint including source Part Number for cross-module sharing.
    ''' </summary>
    Public Function ComputeFullFingerprint(partDoc As PartDocument) As String
        Dim geometryFp = ComputeGeometryFingerprint(partDoc)
        Dim partNumber As String = ""
        Try
            partNumber = partDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
        Catch
            partNumber = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName)
        End Try
        Return "PN:" & partNumber & "|GEO:" & geometryFp
    End Function
    
    ''' <summary>
    ''' Snapshot all model parameters from master documents.
    ''' </summary>
    Public Function SnapshotMasterParameters(app As Inventor.Application, _
                                              masterPaths As List(Of String)) As Dictionary(Of String, Dictionary(Of String, String))
        Dim snapshot As New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)
        
        For Each masterPath In masterPaths
            Dim doc As Document = Nothing
            Try
                For Each d As Document In app.Documents
                    If d.FullFileName.Equals(masterPath, StringComparison.OrdinalIgnoreCase) Then
                        doc = d
                        Exit For
                    End If
                Next
            Catch
            End Try
            
            If doc Is Nothing Then Continue For
            
            Dim paramSnapshot As New Dictionary(Of String, String)
            Try
                Dim params As Parameters = Nothing
                If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    params = CType(doc, PartDocument).ComponentDefinition.Parameters
                ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    params = CType(doc, AssemblyDocument).ComponentDefinition.Parameters
                End If
                
                If params IsNot Nothing Then
                    For Each param As Parameter In params.ModelParameters
                        Try
                            paramSnapshot.Add(param.Name, param.Expression)
                        Catch
                        End Try
                    Next
                    For Each param As Parameter In params.UserParameters
                        Try
                            If Not paramSnapshot.ContainsKey(param.Name) Then
                                paramSnapshot.Add(param.Name, param.Expression)
                            End If
                        Catch
                        End Try
                    Next
                End If
            Catch
            End Try
            
            snapshot.Add(masterPath, paramSnapshot)
        Next
        
        Return snapshot
    End Function
    
    ''' <summary>
    ''' Restore master parameters from snapshot.
    ''' </summary>
    Public Sub RestoreMasterParameters(app As Inventor.Application, _
                                        snapshot As Dictionary(Of String, Dictionary(Of String, String)))
        Dim restoredCount As Integer = 0
        For Each kvp In snapshot
            Dim doc As Document = Nothing
            Try
                For Each d As Document In app.Documents
                    If d.FullFileName.Equals(kvp.Key, StringComparison.OrdinalIgnoreCase) Then
                        doc = d
                        Exit For
                    End If
                Next
            Catch
            End Try
            
            If doc Is Nothing Then
                UtilsLib.LogInfo("    Master not open, skipping: " & System.IO.Path.GetFileName(kvp.Key))
                Continue For
            End If
            
            Dim params As Parameters = Nothing
            Try
                If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    params = CType(doc, PartDocument).ComponentDefinition.Parameters
                ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    params = CType(doc, AssemblyDocument).ComponentDefinition.Parameters
                End If
            Catch
            End Try
            
            If params Is Nothing Then Continue For
            
            Dim paramCount As Integer = 0
            For Each paramKvp In kvp.Value
                Try
                    params.Item(paramKvp.Key).Expression = paramKvp.Value
                    paramCount += 1
                Catch
                End Try
            Next
            
            If paramCount > 0 Then
                UtilsLib.LogInfo("    Restored " & paramCount & " params in: " & System.IO.Path.GetFileName(kvp.Key))
                restoredCount += 1
            End If
        Next
        UtilsLib.LogInfo("    Total masters restored: " & restoredCount & "/" & snapshot.Count)
    End Sub
    
    ''' <summary>
    ''' Set document properties (Part Number, Description, Title).
    ''' </summary>
    Public Sub SetDocumentProperties(doc As Document, partNumber As String, Optional description As String = Nothing)
        Try
            Dim designProps = doc.PropertySets.Item("Design Tracking Properties")
            designProps.Item("Part Number").Value = partNumber
            
            If Not String.IsNullOrEmpty(description) Then
                designProps.Item("Description").Value = description
            End If
        Catch
        End Try
        
        Try
            Dim summaryProps = doc.PropertySets.Item("Inventor Summary Information")
            Dim title As String = If(Not String.IsNullOrEmpty(description), description, partNumber)
            summaryProps.Item("Title").Value = title
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Apply parameters to a document from a dictionary.
    ''' </summary>
    Public Sub ApplyParameters(doc As Document, parameters As Dictionary(Of String, String))
        If doc Is Nothing OrElse parameters Is Nothing Then Return

        Dim params As Parameters = Nothing
        Try
            If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                params = CType(doc, PartDocument).ComponentDefinition.Parameters
            ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                params = CType(doc, AssemblyDocument).ComponentDefinition.Parameters
            End If
        Catch ex As Exception
            UtilsLib.LogWarn("ApplyParameters: Failed to get parameters from " & System.IO.Path.GetFileName(doc.FullFileName) & ": " & ex.Message)
            Return
        End Try

        If params Is Nothing Then
            UtilsLib.LogWarn("ApplyParameters: No parameters collection in " & System.IO.Path.GetFileName(doc.FullFileName))
            Return
        End If

        For Each kvp In parameters
            If kvp.Key.StartsWith("_") Then Continue For
            Try
                ' First check if parameter exists
                Dim param As Parameter = Nothing
                Try
                    param = params.Item(kvp.Key)
                Catch
                    ' Parameter not found by name, try user parameters
                    Try
                        If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                            param = CType(doc, PartDocument).ComponentDefinition.Parameters.UserParameters.Item(kvp.Key)
                        ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            param = CType(doc, AssemblyDocument).ComponentDefinition.Parameters.UserParameters.Item(kvp.Key)
                        End If
                    Catch
                    End Try
                End Try
                
                If param Is Nothing Then
                    ' List available parameters for debugging
                    Dim availableParams As New List(Of String)
                    Try
                        For Each p As Parameter In params
                            If Not p.Name.StartsWith("d") Then ' Skip model params like d0, d1
                                availableParams.Add(p.Name)
                            End If
                        Next
                    Catch
                    End Try
                    UtilsLib.LogWarn("ApplyParameters: Parameter '" & kvp.Key & "' not found in " & System.IO.Path.GetFileName(doc.FullFileName))
                    If availableParams.Count > 0 Then
                        UtilsLib.LogWarn("  Available user parameters: " & String.Join(", ", availableParams.ToArray()))
                    End If
                    Continue For
                End If
                
                Dim oldValue As String = param.Expression
                param.Expression = kvp.Value
                UtilsLib.LogInfo("      Set " & kvp.Key & ": " & oldValue & " -> " & kvp.Value)
            Catch ex As Exception
                UtilsLib.LogWarn("ApplyParameters: Failed to set " & kvp.Key & " = " & kvp.Value & ": " & ex.Message)
            End Try
        Next
    End Sub
    
    ''' <summary>
    ''' Build variant matrix with fingerprints for all parts across all variants.
    ''' This actually applies parameters and measures geometry changes to determine
    ''' which parts are shared (same geometry) vs element-specific (geometry changes).
    ''' </summary>
    Public Function BuildElementMatrix(app As Inventor.Application, _
                                        tree As AssemblyTree, _
                                        variants As List(Of ExcelReaderLib.ElementConfig), _
                                        masterPaths As List(Of String)) As ElementMatrix
        Dim matrix As New ElementMatrix()
        matrix.PartPaths = New List(Of String)(tree.Parts.Keys)
        matrix.ElementNames = New List(Of String)
        For Each vc As ExcelReaderLib.ElementConfig In variants
            matrix.ElementNames.Add(vc.ElementName)
        Next
        
        UtilsLib.LogInfo("BuildElementMatrix: Starting fingerprint analysis")
        UtilsLib.LogInfo("  Parts to analyze: " & matrix.PartPaths.Count)
        UtilsLib.LogInfo("  Elements to test: " & variants.Count)
        UtilsLib.LogInfo("  Master documents: " & masterPaths.Count)
        For Each mp In masterPaths
            UtilsLib.LogInfo("    - " & System.IO.Path.GetFileName(mp))
        Next
        
        ' Step 1: Ensure all parts are loaded (open if not)
        UtilsLib.LogInfo("  Step 1: Ensuring all parts are loaded...")
        Dim loadedParts As New Dictionary(Of String, PartDocument)(StringComparer.OrdinalIgnoreCase)
        Dim partsWeOpened As New List(Of String)
        
        For Each partPath In matrix.PartPaths
            Dim partDoc As PartDocument = Nothing
            
            ' Check if already open
            For Each d As Document In app.Documents
                If d.FullFileName.Equals(partPath, StringComparison.OrdinalIgnoreCase) Then
                    If d.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        partDoc = CType(d, PartDocument)
                    End If
                    Exit For
                End If
            Next
            
            ' Open if not loaded
            If partDoc Is Nothing Then
                Try
                    UtilsLib.LogInfo("    Opening: " & System.IO.Path.GetFileName(partPath))
                    partDoc = CType(app.Documents.Open(partPath, True), PartDocument)
                    partsWeOpened.Add(partPath)
                Catch ex As Exception
                    UtilsLib.LogWarn("    Failed to open: " & System.IO.Path.GetFileName(partPath) & " - " & ex.Message)
                End Try
            End If
            
            If partDoc IsNot Nothing Then
                loadedParts(partPath) = partDoc
            End If
        Next
        UtilsLib.LogInfo("  Loaded " & loadedParts.Count & " of " & matrix.PartPaths.Count & " parts")
        
        ' Step 2: Snapshot current master parameters
        UtilsLib.LogInfo("  Step 2: Snapshotting master parameters...")
        Dim snapshot = SnapshotMasterParameters(app, masterPaths)
        
        Try
            ' Step 3: For each element, apply parameters and compute fingerprints
            For Each variantCfg As ExcelReaderLib.ElementConfig In variants
                UtilsLib.LogInfo("  Step 3: Analyzing element '" & variantCfg.ElementName & "'")
                
                ' Log the parameters being applied
                For Each kvp In variantCfg.Parameters
                    UtilsLib.LogInfo("    Parameter: " & kvp.Key & " = " & kvp.Value)
                Next
                
                ' Apply parameters to all masters
                For Each masterPath In masterPaths
                    Dim doc As Document = Nothing
                    For Each d As Document In app.Documents
                        If d.FullFileName.Equals(masterPath, StringComparison.OrdinalIgnoreCase) Then
                            doc = d
                            Exit For
                        End If
                    Next
                    
                    If doc IsNot Nothing Then
                        UtilsLib.LogInfo("    Applying parameters to: " & System.IO.Path.GetFileName(masterPath))
                        ApplyParameters(doc, variantCfg.Parameters)
                        ' Update master immediately after parameter change
                        Try
                            doc.Update()
                            UtilsLib.LogInfo("    Updated master: " & System.IO.Path.GetFileName(masterPath))
                        Catch ex As Exception
                            UtilsLib.LogWarn("    Failed to update master: " & ex.Message)
                        End Try
                    Else
                        UtilsLib.LogWarn("    Master not loaded: " & System.IO.Path.GetFileName(masterPath))
                    End If
                Next
                
                ' Update the entire assembly tree to propagate changes
                UtilsLib.LogInfo("    Updating assembly to propagate parameter changes...")
                Try
                    ' Update the active document (main assembly)
                    app.ActiveDocument.Update()
                    
                    ' Also update each part individually to ensure derived geometry updates
                    For Each kvp In loadedParts
                        Try
                            kvp.Value.Update()
                        Catch
                        End Try
                    Next
                Catch ex As Exception
                    UtilsLib.LogWarn("    Update warning: " & ex.Message)
                End Try
                
                ' Compute fingerprints for all parts with this parameter configuration
                Dim variantFps As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                UtilsLib.LogInfo("    Computing fingerprints...")
                
                For Each partPath In matrix.PartPaths
                    Dim fp As String
                    
                    If loadedParts.ContainsKey(partPath) Then
                        Dim partDoc As PartDocument = loadedParts(partPath)
                        fp = ComputeGeometryFingerprint(partDoc)
                        UtilsLib.LogInfo("      " & System.IO.Path.GetFileName(partPath) & ": " & fp)
                    Else
                        fp = "NOT_LOADED"
                        UtilsLib.LogWarn("      " & System.IO.Path.GetFileName(partPath) & ": NOT_LOADED")
                    End If
                    
                    variantFps(partPath) = fp
                Next
                
                matrix.Fingerprints.Add(variantCfg.ElementName, variantFps)
            Next
        Finally
            ' Step 4: Restore original master parameters
            UtilsLib.LogInfo("  Step 4: Restoring original master parameters...")
            RestoreMasterParameters(app, snapshot)
            Try
                app.ActiveDocument.Update()
                For Each kvp In loadedParts
                    Try
                        kvp.Value.Update()
                    Catch
                    End Try
                Next
            Catch
            End Try
            
            ' Close parts we opened (only those we opened, not pre-existing)
            UtilsLib.LogInfo("  Step 5: Closing parts we opened...")
            For Each partPath In partsWeOpened
                Try
                    If loadedParts.ContainsKey(partPath) Then
                        loadedParts(partPath).Close(True) ' Close without saving
                    End If
                Catch
                End Try
            Next
        End Try
        
        ' Step 5: Log fingerprint comparison summary
        UtilsLib.LogInfo("BuildElementMatrix: Fingerprint comparison summary:")
        For Each partPath In matrix.PartPaths
            Dim fps As New HashSet(Of String)
            For Each variantName In matrix.ElementNames
                fps.Add(matrix.Fingerprints(variantName)(partPath))
            Next
            Dim shareStatus As String = If(fps.Count = 1, "SHARED", "DIFFERS (" & fps.Count & " variants)")
            UtilsLib.LogInfo("  " & System.IO.Path.GetFileName(partPath) & ": " & shareStatus)
            If fps.Count > 1 Then
                For Each variantName In matrix.ElementNames
                    UtilsLib.LogInfo("    " & variantName & ": " & matrix.Fingerprints(variantName)(partPath))
                Next
            End If
        Next
        
        UtilsLib.LogInfo("BuildElementMatrix: Complete - " & matrix.PartPaths.Count & " parts x " & matrix.ElementNames.Count & " elements")
        Return matrix
    End Function
    
    ''' <summary>
    ''' Classify parts into groups based on fingerprint sharing.
    ''' </summary>
    Public Function ClassifyPartGroups(matrix As ElementMatrix, _
                                        tree As AssemblyTree) As List(Of PartGroup)
        Dim groups As New List(Of PartGroup)
        
        For Each partPath In matrix.PartPaths
            Dim group As New PartGroup()
            group.PartPath = partPath
            group.RelativePath = tree.Parts(partPath).RelativePath
            group.PartNumber = tree.Parts(partPath).PartNumber
            
            For Each variantName In matrix.ElementNames
                Dim fp = matrix.Fingerprints(variantName)(partPath)
                If Not group.UniqueFingerprints.ContainsKey(fp) Then
                    group.UniqueFingerprints.Add(fp, New List(Of String))
                End If
                group.UniqueFingerprints(fp).Add(variantName)
            Next
            
            groups.Add(group)
        Next
        
        Return groups
    End Function
    
    ''' <summary>
    ''' Get master document paths from the assembly tree.
    ''' Masters are identified as the source of derivations.
    ''' Now includes external masters and returns them in dependency order (roots first).
    ''' </summary>
    Public Function GetMasterPaths(tree As AssemblyTree) As List(Of String)
        Dim masters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        
        ' Collect internal masters (from parts in tree)
        For Each kvp In tree.Parts
            If kvp.Value.Role = PartRole.Derived AndAlso Not String.IsNullOrEmpty(kvp.Value.DerivedFromMaster) Then
                masters.Add(kvp.Value.DerivedFromMaster)
            End If
        Next
        
        ' Add external masters discovered during multi-master analysis
        For Each kvp In tree.ExternalMasters
            masters.Add(kvp.Key)
        Next
        
        ' Return in dependency order (roots first)
        Return SortMastersByDependency(masters.ToList(), tree.MasterDependencies)
    End Function
    
    ''' <summary>
    ''' Sort masters by dependency order using topological sort (Kahn's algorithm).
    ''' Returns list with roots first, then dependents, then intermediate assemblies last.
    ''' </summary>
    Public Function SortMastersByDependency(masters As List(Of String), _
                                            dependencies As Dictionary(Of String, List(Of String))) As List(Of String)
        ' Build a set of all masters for quick lookup
        Dim masterSet As New HashSet(Of String)(masters, StringComparer.OrdinalIgnoreCase)
        
        ' Separate parts from assemblies (assemblies go last)
        Dim parts As New List(Of String)
        Dim assemblies As New List(Of String)
        
        For Each m In masters
            If System.IO.Path.GetExtension(m).ToLower() = ".iam" Then
                assemblies.Add(m)
            Else
                parts.Add(m)
            End If
        Next
        
        ' Build in-degree map (count of dependencies for each part)
        Dim inDegree As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For Each m In parts
            inDegree(m) = 0
        Next
        
        ' Count incoming edges (dependencies)
        For Each m In parts
            If dependencies.ContainsKey(m) Then
                For Each dep In dependencies(m)
                    ' Only count if the dependency is also a master we're sorting
                    If masterSet.Contains(dep) AndAlso inDegree.ContainsKey(dep) Then
                        ' This master (m) depends on dep, so m has an incoming edge
                        inDegree(m) += 1
                    End If
                Next
            End If
        Next
        
        ' Start with parts that have no dependencies (in-degree = 0)
        Dim queue As New Queue(Of String)
        For Each kvp In inDegree
            If kvp.Value = 0 Then
                queue.Enqueue(kvp.Key)
            End If
        Next
        
        ' Process in topological order
        Dim sorted As New List(Of String)
        Dim processed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        
        Do While queue.Count > 0
            Dim current As String = queue.Dequeue()
            sorted.Add(current)
            processed.Add(current)
            
            ' Find all parts that depend on 'current' and decrease their in-degree
            For Each kvp In inDegree
                If processed.Contains(kvp.Key) Then Continue For
                
                If dependencies.ContainsKey(kvp.Key) Then
                    If dependencies(kvp.Key).Any(Function(d) d.Equals(current, StringComparison.OrdinalIgnoreCase)) Then
                        inDegree(kvp.Key) -= 1
                        If inDegree(kvp.Key) = 0 Then
                            queue.Enqueue(kvp.Key)
                        End If
                    End If
                End If
            Next
        Loop
        
        ' Add any remaining parts (in case of cycles or unprocessed items)
        For Each m In parts
            If Not processed.Contains(m) Then
                sorted.Add(m)
            End If
        Next
        
        ' Add assemblies at the end
        sorted.AddRange(assemblies)
        
        If sorted.Count > 0 Then
            UtilsLib.LogInfo("SortMastersByDependency: Order:")
            For i As Integer = 0 To sorted.Count - 1
                UtilsLib.LogInfo("  " & (i + 1) & ". " & System.IO.Path.GetFileName(sorted(i)))
            Next
        End If
        
        Return sorted
    End Function
    
    ' ============================================================================
    ' Phase 4: Release Planning
    ' ============================================================================
    
    ''' <summary>
    ''' Get file numbers (local dev mode or Vault production mode).
    ''' </summary>
    Public Function GetFileNumbers(targetRoot As String, count As Integer) As List(Of String)
        If DEVELOPMENT_MODE Then
            Return GenerateLocalNumbers(count, targetRoot)
        Else
            Dim conn = VaultNumberingLib.GetVaultConnection()
            If conn Is Nothing Then
                UtilsLib.LogError("GetFileNumbers: ERROR - No Vault connection available!")
                Return Nothing
            End If
            Return ReserveVaultNumbers(conn, count)
        End If
    End Function
    
    ''' <summary>
    ''' Generate local sequential numbers for development mode.
    ''' In dev mode, we start from 1000 to avoid conflicts with existing files
    ''' (which are typically numbered below 1000 in test projects).
    ''' Vault would never give duplicate numbers, so this is only for local testing.
    ''' </summary>
    Public Function GenerateLocalNumbers(count As Integer, outputRoot As String) As List(Of String)
        UtilsLib.LogInfo("GenerateLocalNumbers: Generating " & count & " local numbers (development mode)")
        
        ' Start from 1000 to avoid conflicts with existing source files (typically < 1000)
        Dim startNum As Integer = 1000
        Dim manifestPath = outputRoot & "\_manifest.json"
        If System.IO.File.Exists(manifestPath) Then
            Try
                Dim manifest = ReadManifest(manifestPath)
                If manifest IsNot Nothing AndAlso manifest.SharedParts.Count > 0 Then
                    For Each sp In manifest.SharedParts
                        Dim num As Integer = 0
                        If Integer.TryParse(sp.VaultNumber, num) Then
                            If num >= startNum Then startNum = num + 1
                        End If
                    Next
                End If
            Catch
            End Try
        End If
        
        Dim numbers As New List(Of String)
        For i As Integer = 0 To count - 1
            numbers.Add((startNum + i).ToString("D5"))
        Next
        
        If numbers.Count > 0 Then
            UtilsLib.LogInfo("GenerateLocalNumbers: Generated " & numbers(0) & " to " & numbers(numbers.Count - 1))
        End If
        Return numbers
    End Function
    
    ''' <summary>
    ''' Reserve Vault numbers (production mode).
    ''' </summary>
    Public Function ReserveVaultNumbers(conn As Object, count As Integer) As List(Of String)
        UtilsLib.LogInfo("ReserveVaultNumbers: Reserving " & count & " Vault numbers (scheme: " & NUMBERING_SCHEME & ")")
        
        Dim scheme = VaultNumberingLib.FindSchemeByName(conn, NUMBERING_SCHEME)
        If scheme Is Nothing Then
            UtilsLib.LogError("ReserveVaultNumbers: ERROR - Numbering scheme '" & NUMBERING_SCHEME & "' not found!")
            Return Nothing
        End If
        
        Return VaultNumberingLib.GenerateFileNumbers(conn, scheme, count)
    End Function
    
    ''' <summary>
    ''' Cached file properties for release plan computation
    ''' </summary>
    Private m_FilePropertiesCache As New Dictionary(Of String, FileProperties)(StringComparer.OrdinalIgnoreCase)
    
    ''' <summary>
    ''' File properties structure
    ''' </summary>
    Private Class FileProperties
        Public Description As String
        Public Project As String
        Public PartNumber As String
    End Class
    
    ''' <summary>
    ''' Gets iProperties from a source file (cached)
    ''' </summary>
    Private Function GetFileProperties(app As Inventor.Application, filePath As String) As FileProperties
        If m_FilePropertiesCache.ContainsKey(filePath) Then
            Return m_FilePropertiesCache(filePath)
        End If
        
        Dim props As New FileProperties()
        props.Description = System.IO.Path.GetFileNameWithoutExtension(filePath)
        props.Project = ""
        props.PartNumber = System.IO.Path.GetFileNameWithoutExtension(filePath)
        
        Try
            ' Check if file is already open
            Dim doc As Document = Nothing
            For Each openDoc As Document In app.Documents
                If String.Equals(openDoc.FullFileName, filePath, StringComparison.OrdinalIgnoreCase) Then
                    doc = openDoc
                    Exit For
                End If
            Next
            
            If doc IsNot Nothing Then
                ' File is open - read properties directly
                Try
                    Dim propSet As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                    Dim descProp As Inventor.Property = propSet.Item("Description")
                    If descProp.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(descProp.Value.ToString()) Then
                        props.Description = descProp.Value.ToString()
                    End If
                    Dim pnProp As Inventor.Property = propSet.Item("Part Number")
                    If pnProp.Value IsNot Nothing Then
                        props.PartNumber = pnProp.Value.ToString()
                    End If
                Catch : End Try
                Try
                    Dim summarySet As PropertySet = doc.PropertySets.Item("Inventor Summary Information")
                    Dim projProp As Inventor.Property = summarySet.Item("Project")
                    If projProp.Value IsNot Nothing Then
                        props.Project = projProp.Value.ToString()
                    End If
                Catch : End Try
            Else
                ' File not open - try to open briefly to read properties
                Try
                    If System.IO.File.Exists(filePath) Then
                        doc = app.Documents.Open(filePath, True) ' Open silently
                        Try
                            Dim propSet As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                            Dim descProp As Inventor.Property = propSet.Item("Description")
                            If descProp.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(descProp.Value.ToString()) Then
                                props.Description = descProp.Value.ToString()
                            End If
                            Dim pnProp As Inventor.Property = propSet.Item("Part Number")
                            If pnProp.Value IsNot Nothing Then
                                props.PartNumber = pnProp.Value.ToString()
                            End If
                        Catch : End Try
                        Try
                            Dim summarySet As PropertySet = doc.PropertySets.Item("Inventor Summary Information")
                            Dim projProp As Inventor.Property = summarySet.Item("Project")
                            If projProp.Value IsNot Nothing Then
                                props.Project = projProp.Value.ToString()
                            End If
                        Catch : End Try
                        doc.Close(True) ' Close without saving
                    End If
                Catch : End Try
            End If
        Catch : End Try
        
        m_FilePropertiesCache(filePath) = props
        Return props
    End Function
    
    ''' <summary>
    ''' Clears the file properties cache
    ''' </summary>
    Public Sub ClearFilePropertiesCache()
        m_FilePropertiesCache.Clear()
    End Sub
    
    ''' <summary>
    ''' Generates placeholder numbers for preview (not real Vault numbers)
    ''' Format: "UUS-001", "UUS-002", etc.
    ''' </summary>
    Public Function GeneratePlaceholderNumbers(count As Integer) As List(Of String)
        Dim placeholders As New List(Of String)
        For i As Integer = 1 To count
            placeholders.Add("UUS-" & i.ToString("D3"))
        Next
        Return placeholders
    End Function
    
    ''' <summary>
    ''' Allocates real file numbers and updates the release plan.
    ''' Call this AFTER user confirms release, not before UI.
    ''' </summary>
    Public Function AllocateRealNumbers(plan As ReleasePlan, targetRoot As String) As Boolean
        ' Count how many real numbers we need (parts and assemblies only, not drawings)
        Dim count As Integer = 0
        For Each f As PlannedFile In plan.Files
            If f.IsPlaceholder AndAlso Not f.IsExisting Then
                count += 1
            End If
        Next
        
        If count = 0 Then Return True
        
        ' Get real numbers
        Dim realNumbers As List(Of String) = GetFileNumbers(targetRoot, count)
        If realNumbers Is Nothing OrElse realNumbers.Count < count Then
            UtilsLib.LogError("AllocateRealNumbers: Failed to get " & count & " numbers")
            Return False
        End If
        
        ' Map of placeholder -> real number (for updating drawings)
        Dim numberMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        
        ' Replace placeholders with real numbers for parts and assemblies
        Dim numberIndex As Integer = 0
        For Each f As PlannedFile In plan.Files
            If f.IsPlaceholder AndAlso Not f.IsExisting Then
                Dim oldNumber As String = f.VaultNumber
                Dim newNumber As String = realNumbers(numberIndex)
                
                ' Track mapping for drawings
                If Not numberMap.ContainsKey(oldNumber) Then
                    numberMap.Add(oldNumber, newNumber)
                End If
                
                ' Update VaultNumber
                f.VaultNumber = newNumber
                f.IsPlaceholder = False
                
                ' Update target path (replace placeholder in filename)
                If f.TargetLocalPath.Contains(oldNumber) Then
                    f.TargetLocalPath = f.TargetLocalPath.Replace(oldNumber, newNumber)
                End If
                
                UtilsLib.LogInfo("  Allocated: " & oldNumber & " -> " & newNumber)
                numberIndex += 1
            End If
        Next
        
        ' Now update drawings to use the new model numbers
        ' Drawings don't get their own numbers - they use their referenced model's number
        Dim drawingsUpdated As Integer = 0
        For Each f As PlannedFile In plan.Files
            If f.FileType = FileType.Drawing AndAlso Not f.IsExisting Then
                Dim oldNumber As String = f.VaultNumber
                If numberMap.ContainsKey(oldNumber) Then
                    Dim newNumber As String = numberMap(oldNumber)
                    f.VaultNumber = newNumber
                    
                    ' Update target path (replace old number with new in filename)
                    If f.TargetLocalPath.Contains(oldNumber) Then
                        f.TargetLocalPath = f.TargetLocalPath.Replace(oldNumber, newNumber)
                    End If
                    
                    UtilsLib.LogInfo("  Drawing: " & oldNumber & " -> " & newNumber)
                    drawingsUpdated += 1
                End If
            End If
        Next
        
        UtilsLib.LogInfo("AllocateRealNumbers: Allocated " & numberIndex & " numbers, updated " & drawingsUpdated & " drawings")
        Return True
    End Function
    
    ''' <summary>
    ''' Computes projected description for a file based on what it will be after release.
    ''' </summary>
    Private Function ComputeProjectedDescription(f As PlannedFile, elementName As String) As String
        Select Case f.FileType
            Case FileType.Assembly
                ' Assemblies get element name as their description
                Return elementName
            Case FileType.Part
                ' Parts keep their source description
                Return f.SourceDescription
            Case FileType.Drawing
                ' Drawings keep their source description (or derive from model)
                Return f.SourceDescription
            Case Else
                Return f.SourceDescription
        End Select
    End Function
    
    ''' <summary>
    ''' Computes projected project name for a file.
    ''' </summary>
    Private Function ComputeProjectedProject(f As PlannedFile, elementName As String, productFamily As String) As String
        ' Project could be the product family or element name
        If Not String.IsNullOrEmpty(productFamily) Then
            Return productFamily
        End If
        Return elementName
    End Function
    
    ''' <summary>
    ''' Compute the complete release plan.
    ''' </summary>
    Public Function ComputeReleasePlan(app As Inventor.Application, _
                                        tree As AssemblyTree, _
                                        partGroups As List(Of PartGroup), _
                                        variants As List(Of ExcelReaderLib.ElementConfig), _
                                        targetRoot As String, _
                                        fileNumbers As List(Of String)) As ReleasePlan
        Dim plan As New ReleasePlan()
        plan.SharedFolder = System.IO.Path.Combine(targetRoot, "Ühine")
        
        ' Clear properties cache for fresh computation
        ClearFilePropertiesCache()

        Dim numberIndex As Integer = 0

        For Each variantCfg As ExcelReaderLib.ElementConfig In variants
            plan.VariantFolders.Add(variantCfg.ElementName, System.IO.Path.Combine(targetRoot, variantCfg.ElementName))
        Next

        ' Sharing only makes sense with 2+ moodulid
        Dim canShare As Boolean = (variants.Count >= 2)
        
        For Each group In partGroups
            ' A part is shared if:
            ' 1. We have 2+ moodulid to share between
            ' 2. The part has exactly 1 unique fingerprint (same geometry in all moodulid)
            ' 3. That fingerprint is used by 2+ moodulid
            Dim isShared As Boolean = False
            If canShare AndAlso group.UniqueFingerprints.Count = 1 Then
                For Each fpKvp In group.UniqueFingerprints
                    If fpKvp.Value.Count >= 2 Then
                        isShared = True
                        Exit For
                    End If
                Next
            End If
            
            If isShared Then
                Dim fp As String = GetFirstKey(group.UniqueFingerprints)
                Dim allVariantNames As New List(Of String)
                For Each vc As ExcelReaderLib.ElementConfig In variants
                    allVariantNames.Add(vc.ElementName)
                Next
                Dim partProps As FileProperties = GetFileProperties(app, group.PartPath)
                ' Construct path with new filename (not original filename)
                Dim newFileNameShared As String = fileNumbers(numberIndex) & System.IO.Path.GetExtension(group.PartPath)
                Dim relDirShared As String = System.IO.Path.GetDirectoryName(group.RelativePath)
                Dim pf As New PlannedFile With {
                    .SourcePath = group.PartPath,
                    .TargetLocalPath = System.IO.Path.Combine(plan.SharedFolder, relDirShared, newFileNameShared),
                    .VaultNumber = fileNumbers(numberIndex),
                    .FileType = FileType.Part,
                    .IsShared = True,
                    .IsExisting = False,
                    .ForVariants = allVariantNames,
                    .Fingerprint = fp,
                    .SourceDescription = partProps.Description,
                    .SourceProject = partProps.Project,
                    .SourcePartNumber = partProps.PartNumber,
                    .IsPlaceholder = True
                }
                ' Parts keep their source description
                pf.ProjectedDescription = partProps.Description
                pf.ProjectedProject = partProps.Project
                plan.Files.Add(pf)
                numberIndex += 1
            Else
                Dim partPropsNonShared As FileProperties = GetFileProperties(app, group.PartPath)
                For Each fpKvp In group.UniqueFingerprints
                    Dim firstVariant = fpKvp.Value(0)
                    Dim newFileName As String = fileNumbers(numberIndex) & System.IO.Path.GetExtension(group.PartPath)
                    Dim relDir As String = System.IO.Path.GetDirectoryName(group.RelativePath)
                    
                    Dim pf As New PlannedFile With {
                        .SourcePath = group.PartPath,
                        .TargetLocalPath = System.IO.Path.Combine(plan.VariantFolders(firstVariant), relDir, newFileName),
                        .VaultNumber = fileNumbers(numberIndex),
                        .FileType = FileType.Part,
                        .IsShared = False,
                        .IsExisting = False,
                        .ForVariants = fpKvp.Value,
                        .Fingerprint = fpKvp.Key,
                        .SourceDescription = partPropsNonShared.Description,
                        .SourceProject = partPropsNonShared.Project,
                        .SourcePartNumber = partPropsNonShared.PartNumber,
                        .IsPlaceholder = True
                    }
                    ' Parts keep their source description
                    pf.ProjectedDescription = partPropsNonShared.Description
                    pf.ProjectedProject = partPropsNonShared.Project
                    plan.Files.Add(pf)
                    numberIndex += 1
                Next
            End If
        Next
        
        For Each variantCfg As ExcelReaderLib.ElementConfig In variants
            For Each asmKvp In tree.Assemblies
                Dim relativePath = asmKvp.Value.RelativePath
                Dim newFileName As String = fileNumbers(numberIndex) & ".iam"
                Dim relDir As String = System.IO.Path.GetDirectoryName(relativePath)
                Dim asmProps As FileProperties = GetFileProperties(app, asmKvp.Key)
                
                ' Build projected description from element name and parameters
                Dim projDesc As String = variantCfg.ElementName
                If variantCfg.Parameters.Count > 0 Then
                    Dim paramParts As New List(Of String)
                    For Each kvp In variantCfg.Parameters
                        paramParts.Add(kvp.Key & "=" & kvp.Value)
                    Next
                    projDesc &= " (" & String.Join(", ", paramParts.ToArray()) & ")"
                End If
                
                Dim pf As New PlannedFile With {
                    .SourcePath = asmKvp.Key,
                    .TargetLocalPath = System.IO.Path.Combine(plan.VariantFolders(variantCfg.ElementName), relDir, newFileName),
                    .VaultNumber = fileNumbers(numberIndex),
                    .FileType = FileType.Assembly,
                    .IsShared = False,
                    .ForVariants = New List(Of String) From {variantCfg.ElementName},
                    .SourceDescription = asmProps.Description,
                    .SourceProject = asmProps.Project,
                    .SourcePartNumber = asmProps.PartNumber,
                    .IsPlaceholder = True
                }
                ' Assemblies get element name as their projected description
                pf.ProjectedDescription = projDesc
                pf.ProjectedProject = asmProps.Project ' Keep source project or could use family name
                plan.Files.Add(pf)
                numberIndex += 1
            Next
        Next
        
        For Each dwgInfo As DrawingInfo In tree.Drawings
            ' Get drawing filename and primary referenced model
            Dim dwgFileName As String = System.IO.Path.GetFileNameWithoutExtension(dwgInfo.DrawingPath)
            Dim primaryModelPath As String = If(dwgInfo.ReferencedModelPaths.Count > 0, dwgInfo.ReferencedModelPaths(0), "")
            Dim sourceModelNumber As String = System.IO.Path.GetFileNameWithoutExtension(primaryModelPath)
            
            ' Extract suffix/prefix from drawing filename
            ' The drawing number should be: [new model number] + [suffix]
            ' e.g., "00005_sheet2.idw" referencing "00005.ipt" -> "01234_sheet2.idw" when model becomes "01234.ipt"
            ' e.g., "00005.idw" referencing "00005.iam" -> "01234.idw" when model becomes "01234.iam"
            Dim dwgSuffix As String = ""
            If Not String.IsNullOrEmpty(sourceModelNumber) AndAlso dwgFileName.StartsWith(sourceModelNumber, StringComparison.OrdinalIgnoreCase) Then
                dwgSuffix = dwgFileName.Substring(sourceModelNumber.Length)
            End If
            
            ' A drawing can only be shared if:
            ' 1. We have 2+ variants to share between
            ' 2. All its referenced parts have the same geometry across all variants
            Dim allRefsShared As Boolean = canShare
            If canShare Then
                For Each refPath In dwgInfo.ReferencedModelPaths
                    Dim group = FindPartGroupByPath(partGroups, refPath)
                    If group Is Nothing OrElse group.UniqueFingerprints.Count > 1 Then
                        allRefsShared = False
                        Exit For
                    End If
                Next
            End If
            
            If allRefsShared Then
                ' SHARED DRAWING: Find the referenced model's vault number
                ' Drawings don't get their own numbers - they derive from the model they reference
                Dim modelFile As PlannedFile = FindPlannedFileBySource(plan.Files, primaryModelPath)
                If modelFile Is Nothing Then
                    UtilsLib.LogWarn("Drawing " & dwgFileName & ".idw: Could not find referenced model " & System.IO.Path.GetFileName(primaryModelPath))
                    Continue For
                End If
                
                Dim vaultNum As String = modelFile.VaultNumber
                Dim newFileName As String = vaultNum & dwgSuffix & ".idw"
                UtilsLib.LogInfo("Drawing " & dwgFileName & ".idw -> " & newFileName & " (from model " & modelFile.VaultNumber & ")")
                
                Dim relDir As String = System.IO.Path.GetDirectoryName(dwgInfo.RelativePath)
                Dim allVariantNames2 As New List(Of String)
                For Each vc2 As ExcelReaderLib.ElementConfig In variants
                    allVariantNames2.Add(vc2.ElementName)
                Next
                Dim dwgPropsShared As FileProperties = GetFileProperties(app, dwgInfo.DrawingPath)
                
                Dim pf As New PlannedFile With {
                    .SourcePath = dwgInfo.DrawingPath,
                    .TargetLocalPath = System.IO.Path.Combine(plan.SharedFolder, relDir, newFileName),
                    .VaultNumber = vaultNum,
                    .FileType = FileType.Drawing,
                    .IsShared = True,
                    .ForVariants = allVariantNames2,
                    .SourceDescription = dwgPropsShared.Description,
                    .SourceProject = dwgPropsShared.Project,
                    .SourcePartNumber = dwgPropsShared.PartNumber,
                    .IsPlaceholder = False  ' Drawing number comes from model, not placeholder
                }
                pf.ProjectedDescription = dwgPropsShared.Description
                pf.ProjectedProject = dwgPropsShared.Project
                plan.Files.Add(pf)
            Else
                ' VARIANT-SPECIFIC DRAWING: Each variant gets its own copy with that variant's model number
                For Each variantCfg2 As ExcelReaderLib.ElementConfig In variants
                    ' Find the referenced model's vault number for this variant
                    Dim modelFile As PlannedFile = FindPlannedFileBySourceAndVariant(plan.Files, primaryModelPath, variantCfg2.ElementName)
                    If modelFile Is Nothing Then
                        UtilsLib.LogWarn("Drawing " & dwgFileName & ".idw (" & variantCfg2.ElementName & "): Could not find referenced model")
                        Continue For
                    End If
                    
                    Dim vaultNum As String = modelFile.VaultNumber
                    Dim newFileName As String = vaultNum & dwgSuffix & ".idw"
                    UtilsLib.LogInfo("Drawing " & dwgFileName & ".idw (" & variantCfg2.ElementName & ") -> " & newFileName & " (from model " & modelFile.VaultNumber & ")")
                    
                    Dim relDir As String = System.IO.Path.GetDirectoryName(dwgInfo.RelativePath)
                    Dim dwgPropsNonShared As FileProperties = GetFileProperties(app, dwgInfo.DrawingPath)
                    
                    Dim pf As New PlannedFile With {
                        .SourcePath = dwgInfo.DrawingPath,
                        .TargetLocalPath = System.IO.Path.Combine(plan.VariantFolders(variantCfg2.ElementName), relDir, newFileName),
                        .VaultNumber = vaultNum,
                        .FileType = FileType.Drawing,
                        .IsShared = False,
                        .ForVariants = New List(Of String) From {variantCfg2.ElementName},
                        .SourceDescription = dwgPropsNonShared.Description,
                        .SourceProject = dwgPropsNonShared.Project,
                        .SourcePartNumber = dwgPropsNonShared.PartNumber,
                        .IsPlaceholder = False  ' Drawing number comes from model, not placeholder
                    }
                    pf.ProjectedDescription = dwgPropsNonShared.Description
                    pf.ProjectedProject = dwgPropsNonShared.Project
                    plan.Files.Add(pf)
                Next
            End If
        Next
        
        ' ================================================================
        ' Multi-Master Support: Add external masters and intermediate assemblies
        ' ================================================================
        ' External masters are copied to each variant's Eskiis folder
        ' They're per-variant because each element gets its own copy
        
        Dim externalMasterCount As Integer = 0
        For Each variantCfg As ExcelReaderLib.ElementConfig In variants
            ' Add external masters
            For Each kvp In tree.ExternalMasters
                Dim masterPath As String = kvp.Key
                Dim masterInfo As MasterInfo = kvp.Value
                
                ' Skip if already added for this variant
                If plan.Files.Any(Function(f) f.SourcePath.Equals(masterPath, StringComparison.OrdinalIgnoreCase) AndAlso f.ForVariants.Contains(variantCfg.ElementName)) Then
                    Continue For
                End If
                
                Dim masterFileName As String = System.IO.Path.GetFileName(masterPath)
                Dim newFileName As String = fileNumbers(numberIndex) & System.IO.Path.GetExtension(masterPath)
                Dim eskiisFolder As String = System.IO.Path.Combine(plan.VariantFolders(variantCfg.ElementName), "Eskiis")
                Dim masterProps As FileProperties = GetFileProperties(app, masterPath)
                
                Dim fileType As FileType = If(System.IO.Path.GetExtension(masterPath).ToLower() = ".iam", FileType.Assembly, FileType.Part)
                
                Dim pf As New PlannedFile With {
                    .SourcePath = masterPath,
                    .TargetLocalPath = System.IO.Path.Combine(eskiisFolder, newFileName),
                    .VaultNumber = fileNumbers(numberIndex),
                    .FileType = fileType,
                    .IsShared = False,
                    .IsExisting = False,
                    .ForVariants = New List(Of String) From {variantCfg.ElementName},
                    .SourceDescription = masterProps.Description,
                    .SourceProject = masterProps.Project,
                    .SourcePartNumber = masterProps.PartNumber,
                    .IsPlaceholder = True
                }
                pf.ProjectedDescription = masterProps.Description
                pf.ProjectedProject = masterProps.Project
                plan.Files.Add(pf)
                numberIndex += 1
                externalMasterCount += 1
                
                UtilsLib.LogInfo("ComputeReleasePlan: Added external master: " & masterFileName & " -> " & variantCfg.ElementName & "/Eskiis/")
            Next
            
            ' Add intermediate assemblies (needed for projected geometry)
            For Each intAsmPath In tree.IntermediateAssemblies
                ' Skip if it's also in tree.Assemblies (already handled)
                If tree.Assemblies.ContainsKey(intAsmPath) Then Continue For
                
                ' Skip if already added for this variant
                If plan.Files.Any(Function(f) f.SourcePath.Equals(intAsmPath, StringComparison.OrdinalIgnoreCase) AndAlso f.ForVariants.Contains(variantCfg.ElementName)) Then
                    Continue For
                End If
                
                Dim asmFileName As String = System.IO.Path.GetFileName(intAsmPath)
                Dim newFileName As String = fileNumbers(numberIndex) & ".iam"
                Dim eskiisFolder As String = System.IO.Path.Combine(plan.VariantFolders(variantCfg.ElementName), "Eskiis")
                Dim asmProps As FileProperties = GetFileProperties(app, intAsmPath)
                
                Dim pf As New PlannedFile With {
                    .SourcePath = intAsmPath,
                    .TargetLocalPath = System.IO.Path.Combine(eskiisFolder, newFileName),
                    .VaultNumber = fileNumbers(numberIndex),
                    .FileType = FileType.Assembly,
                    .IsShared = False,
                    .IsExisting = False,
                    .ForVariants = New List(Of String) From {variantCfg.ElementName},
                    .SourceDescription = asmProps.Description,
                    .SourceProject = asmProps.Project,
                    .SourcePartNumber = asmProps.PartNumber,
                    .IsPlaceholder = True
                }
                pf.ProjectedDescription = asmProps.Description
                pf.ProjectedProject = asmProps.Project
                plan.Files.Add(pf)
                numberIndex += 1
                
                UtilsLib.LogInfo("ComputeReleasePlan: Added intermediate assembly: " & asmFileName & " -> " & variantCfg.ElementName & "/Eskiis/")
            Next
        Next
        
        If externalMasterCount > 0 Then
            UtilsLib.LogInfo("ComputeReleasePlan: Added " & externalMasterCount & " external master files")
        End If
        
        Dim sharedCount As Integer = 0
        Dim variantSpecificCount As Integer = 0
        Dim partsCount As Integer = 0
        Dim assembliesCount As Integer = 0
        Dim drawingsCount As Integer = 0
        For Each f As PlannedFile In plan.Files
            If f.IsShared Then sharedCount += 1 Else variantSpecificCount += 1
            If f.FileType = FileType.Part Then partsCount += 1
            If f.FileType = FileType.Assembly Then assembliesCount += 1
            If f.FileType = FileType.Drawing Then drawingsCount += 1
        Next
        
        UtilsLib.LogInfo("ComputeReleasePlan: Total files: " & plan.Files.Count)
        UtilsLib.LogInfo("  - Shared: " & sharedCount)
        UtilsLib.LogInfo("  - Element-specific: " & variantSpecificCount)
        UtilsLib.LogInfo("  - Parts: " & partsCount)
        UtilsLib.LogInfo("  - Assemblies: " & assembliesCount)
        UtilsLib.LogInfo("  - Drawings: " & drawingsCount)
        
        Return plan
    End Function
    
    ''' <summary>
    ''' Find a part group by path.
    ''' </summary>
    Private Function FindPartGroupByPath(partGroups As List(Of PartGroup), refPath As String) As PartGroup
        For Each g As PartGroup In partGroups
            If g.PartPath.Equals(refPath, StringComparison.OrdinalIgnoreCase) Then
                Return g
            End If
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Get first key from a dictionary.
    ''' </summary>
    Private Function GetFirstKey(dict As Dictionary(Of String, List(Of String))) As String
        For Each k As String In dict.Keys
            Return k
        Next
        Return ""
    End Function
    
    ''' <summary>
    ''' Find a planned file by source path and variant name.
    ''' </summary>
    Private Function FindPlannedFile(files As List(Of PlannedFile), sourcePath As String, variantName As String) As PlannedFile
        For Each f As PlannedFile In files
            If f.SourcePath.Equals(sourcePath, StringComparison.OrdinalIgnoreCase) AndAlso f.ForVariants.Contains(variantName) Then
                Return f
            End If
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Find the first planned file by source path (for shared files).
    ''' </summary>
    Private Function FindPlannedFileBySource(files As List(Of PlannedFile), sourcePath As String) As PlannedFile
        For Each f As PlannedFile In files
            If f.SourcePath.Equals(sourcePath, StringComparison.OrdinalIgnoreCase) Then
                Return f
            End If
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Find a planned file by source path and variant name.
    ''' </summary>
    Private Function FindPlannedFileBySourceAndVariant(files As List(Of PlannedFile), sourcePath As String, variantName As String) As PlannedFile
        For Each f As PlannedFile In files
            If f.SourcePath.Equals(sourcePath, StringComparison.OrdinalIgnoreCase) AndAlso f.ForVariants.Contains(variantName) Then
                Return f
            End If
        Next
        ' Also check shared files that apply to all variants
        For Each f As PlannedFile In files
            If f.SourcePath.Equals(sourcePath, StringComparison.OrdinalIgnoreCase) AndAlso f.IsShared Then
                Return f
            End If
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Filters a release plan to only include files for selected elements.
    ''' Shared files are included if ANY selected element needs them.
    ''' </summary>
    Public Function FilterReleasePlan(plan As ReleasePlan, _
                                       selectedElements As List(Of ExcelReaderLib.ElementConfig)) As ReleasePlan
        Dim filteredPlan As New ReleasePlan()
        filteredPlan.SharedFolder = plan.SharedFolder
        For Each kvp In plan.VariantFolders
            filteredPlan.VariantFolders(kvp.Key) = kvp.Value
        Next
        
        ' Build set of selected element names
        Dim selectedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each elem As ExcelReaderLib.ElementConfig In selectedElements
            selectedNames.Add(elem.ElementName)
        Next
        
        ' Filter files
        For Each f As PlannedFile In plan.Files
            Dim includeFile As Boolean = False
            
            If f.IsShared Then
                ' Shared files are always included if any element is selected
                includeFile = (selectedNames.Count > 0)
            Else
                ' Element-specific files - check if any ForVariants match selected
                For Each variantName In f.ForVariants
                    If selectedNames.Contains(variantName) Then
                        includeFile = True
                        Exit For
                    End If
                Next
            End If
            
            If includeFile Then
                filteredPlan.Files.Add(f)
            End If
        Next
        
        Return filteredPlan
    End Function
    
    ''' <summary>
    ''' Show plan confirmation dialog.
    ''' </summary>
    Public Function ShowPlanConfirmationDialog(plan As ReleasePlan) As Boolean
        Dim firstNum As String = If(plan.Files.Count > 0, plan.Files(0).VaultNumber, "N/A")
        Dim lastNum As String = If(plan.Files.Count > 0, plan.Files(plan.Files.Count - 1).VaultNumber, "N/A")
        
        Dim sharedPartsCount As Integer = 0
        Dim variantPartsCount As Integer = 0
        Dim asmCount As Integer = 0
        Dim dwgCount As Integer = 0
        For Each f As PlannedFile In plan.Files
            If f.FileType = FileType.Part Then
                If f.IsShared Then sharedPartsCount += 1 Else variantPartsCount += 1
            ElseIf f.FileType = FileType.Assembly Then
                asmCount += 1
            ElseIf f.FileType = FileType.Drawing Then
                dwgCount += 1
            End If
        Next
        
        Dim message As String = "Väljastamise plaan:" & vbCrLf & vbCrLf &
            "Faile kokku: " & plan.Files.Count & vbCrLf &
            "  - Jagatud detailid: " & sharedPartsCount & vbCrLf &
            "  - Elemendi detailid: " & variantPartsCount & vbCrLf &
            "  - Koostud: " & asmCount & vbCrLf &
            "  - Joonised: " & dwgCount & vbCrLf & vbCrLf &
            "Numbrid: " & firstNum & " kuni " & lastNum & vbCrLf & vbCrLf &
            "Kas jätkata väljastamisega?"
        
        Return MessageBox.Show(message, "Kinnita väljastamine", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function
    
    ' ============================================================================
    ' Phase 5: File Release Execution
    ' ============================================================================
    
    ''' <summary>
    ''' Create a standalone part by using SaveAs (new GUID) and breaking derivation links.
    ''' </summary>
    Public Function CreateStandalonePart(app As Inventor.Application, _
                                          sourcePartPath As String, _
                                          targetPath As String, _
                                          newPartNumber As String, _
                                          Optional projectedDescription As String = Nothing) As Boolean
        Try
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(targetPath))

            ' Open source part (might already be open as part of the assembly)
            Dim partDoc As PartDocument = Nothing
            Dim wasAlreadyOpen As Boolean = False

            For Each doc As Document In app.Documents
                If doc.FullFileName.Equals(sourcePartPath, StringComparison.OrdinalIgnoreCase) Then
                    partDoc = CType(doc, PartDocument)
                    wasAlreadyOpen = True
                    Exit For
                End If
            Next

            If partDoc Is Nothing Then
                partDoc = CType(app.Documents.Open(sourcePartPath, True), PartDocument)
            End If

            Try
                ' Use SaveAs FIRST to create new file with NEW GUID
                ' This avoids GUID conflicts when both source and target are open
                partDoc.SaveAs(targetPath, False)
                UtilsLib.LogInfo("  SaveAs with new GUID: " & System.IO.Path.GetFileName(targetPath))

                ' Document is now the target file - break ALL external references
                ' This is critical to ensure released parts don't depend on source/master files
                ' TEMPORARILY DISABLED FOR TESTING - see docs/research/2026-05-13-element-release-failures.md
                ' BreakAllExternalLinks(partDoc)
                UtilsLib.LogInfo("  [TEST] Skipping BreakAllExternalLinks - keeping derived links")

                ' Set Part Number and Description in Design Tracking Properties
                Try
                    Dim designProps = partDoc.PropertySets.Item("Design Tracking Properties")
                    designProps.Item("Part Number").Value = newPartNumber
                    UtilsLib.LogInfo("  Set Part Number: " & newPartNumber)
                    
                    ' Set Description if projected description provided
                    If Not String.IsNullOrEmpty(projectedDescription) Then
                        designProps.Item("Description").Value = projectedDescription
                        UtilsLib.LogInfo("  Set Description: " & projectedDescription)
                    End If
                Catch ex As Exception
                    UtilsLib.LogWarn("  WARNING: Failed to set Part Number: " & ex.Message)
                End Try

                ' Also set Title in Summary Information (title blocks often use this)
                ' Title should be the description, not the part number
                Try
                    Dim summaryProps = partDoc.PropertySets.Item("Inventor Summary Information")
                    Dim newTitle As String = If(Not String.IsNullOrEmpty(projectedDescription), projectedDescription, newPartNumber)
                    summaryProps.Item("Title").Value = newTitle
                Catch ex As Exception
                End Try

                partDoc.Save()
                UtilsLib.LogInfo("Created standalone: " & System.IO.Path.GetFileName(targetPath))
                
                ' Close the target document (source will be closed at end of ExecuteRelease)
                partDoc.Close(True)
                
                Return True
                
            Catch ex As Exception
                Throw
            End Try
            
        Catch ex As Exception
            UtilsLib.LogError("ERROR creating standalone: " & ex.Message)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Create a standalone part from an already-open document with updated geometry.
    ''' This variant is used when we have the part already loaded with correct parameter values.
    ''' </summary>
    Public Function CreateStandalonePartFromDocument(app As Inventor.Application, _
                                                      partDoc As PartDocument, _
                                                      targetPath As String, _
                                                      newPartNumber As String, _
                                                      Optional projectedDescription As String = Nothing) As Boolean
        Try
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(targetPath))
            
            ' Use SaveAs FIRST to create new file with NEW GUID
            ' This avoids GUID conflicts when both source and target are open
            partDoc.SaveAs(targetPath, False)
            UtilsLib.LogInfo("  SaveAs with new GUID: " & System.IO.Path.GetFileName(targetPath))
            
            ' Document is now the target file - break ALL external references
            ' This is critical to ensure released parts don't depend on source/master files
            ' TEMPORARILY DISABLED FOR TESTING - see docs/research/2026-05-13-element-release-failures.md
            ' BreakAllExternalLinks(partDoc)
            UtilsLib.LogInfo("  [TEST] Skipping BreakAllExternalLinks - keeping derived links")
            
            ' Set Part Number and Description in Design Tracking Properties
            Try
                Dim designProps = partDoc.PropertySets.Item("Design Tracking Properties")
                designProps.Item("Part Number").Value = newPartNumber
                UtilsLib.LogInfo("  Set Part Number: " & newPartNumber)
                
                ' Set Description if projected description provided
                If Not String.IsNullOrEmpty(projectedDescription) Then
                    designProps.Item("Description").Value = projectedDescription
                    UtilsLib.LogInfo("  Set Description: " & projectedDescription)
                End If
            Catch ex As Exception
                UtilsLib.LogWarn("  WARNING: Failed to set Part Number: " & ex.Message)
            End Try
            
            ' Also set Title in Summary Information (title blocks often use this)
            ' Title should be the description, not the part number
            Try
                Dim summaryProps = partDoc.PropertySets.Item("Inventor Summary Information")
                Dim newTitle As String = If(Not String.IsNullOrEmpty(projectedDescription), projectedDescription, newPartNumber)
                summaryProps.Item("Title").Value = newTitle
            Catch ex As Exception
            End Try
            
            partDoc.Save()
            UtilsLib.LogInfo("Created standalone: " & System.IO.Path.GetFileName(targetPath))
            
            ' Close the target document (source will be closed at end of ExecuteRelease)
            partDoc.Close(True)
            
            Return True
            
        Catch ex As Exception
            UtilsLib.LogError("ERROR creating standalone from document: " & ex.Message)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Break ALL external links in a part document.
    ''' This is critical to ensure released parts are completely standalone and
    ''' don't depend on source/master/skeleton files.
    ''' </summary>
    Private Sub BreakAllExternalLinks(partDoc As PartDocument)
        Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
        Dim refComps = compDef.ReferenceComponents
        
        ' 1. Break DerivedPartComponents (parts derived from multibody masters)
        Dim dpcList As New List(Of DerivedPartComponent)
        For Each dpc As DerivedPartComponent In refComps.DerivedPartComponents
            dpcList.Add(dpc)
        Next
        For Each dpc As DerivedPartComponent In dpcList
            Try
                Dim masterPath As String = ""
                Try : masterPath = dpc.ReferencedFile.FullFileName : Catch : End Try
                dpc.BreakLinkToFile()
                UtilsLib.LogInfo("  Broke DerivedPartComponent link: " & If(masterPath <> "", System.IO.Path.GetFileName(masterPath), dpc.Name))
            Catch ex As Exception
                UtilsLib.LogWarn("  WARNING: Could not break DerivedPartComponent - " & ex.Message)
            End Try
        Next
        
        ' 2. Break DerivedAssemblyComponents (parts derived from assemblies)
        Dim dacList As New List(Of DerivedAssemblyComponent)
        For Each dac As DerivedAssemblyComponent In refComps.DerivedAssemblyComponents
            dacList.Add(dac)
        Next
        For Each dac As DerivedAssemblyComponent In dacList
            Try
                Dim masterPath As String = ""
                Try : masterPath = dac.ReferencedFile.FullFileName : Catch : End Try
                dac.BreakLinkToFile()
                UtilsLib.LogInfo("  Broke DerivedAssemblyComponent link: " & If(masterPath <> "", System.IO.Path.GetFileName(masterPath), dac.Name))
            Catch ex As Exception
                UtilsLib.LogWarn("  WARNING: Could not break DerivedAssemblyComponent - " & ex.Message)
            End Try
        Next
        
        ' 3. Break ImportedComponents
        Dim icList As New List(Of ImportedComponent)
        For Each ic As ImportedComponent In refComps.ImportedComponents
            icList.Add(ic)
        Next
        For Each ic As ImportedComponent In icList
            Try
                ic.BreakLink()
                UtilsLib.LogInfo("  Broke ImportedComponent link")
            Catch ex As Exception
                UtilsLib.LogWarn("  WARNING: Could not break ImportedComponent - " & ex.Message)
            End Try
        Next
        
        ' 4. Check for embedded ReferenceFeatures that might link to external files
        ' These are features that reference geometry from other documents
        Try
            For Each feature As Object In compDef.Features
                Try
                    ' Try to access as ReferenceFeature
                    Dim refFeature As ReferenceFeature = TryCast(feature, ReferenceFeature)
                    If refFeature IsNot Nothing Then
                        Try
                            refFeature.BreakLink()
                            UtilsLib.LogInfo("  Broke ReferenceFeature link")
                        Catch
                        End Try
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
        
        ' 5. Check for sketches with projected geometry from external documents
        ' This is particularly important for sheet metal parts
        For Each sketch As PlanarSketch In compDef.Sketches
            Try
                ' Projected cuts are often linked to external geometry
                ' Delete projected geometry that references external files
                Dim projectedGeoToDelete As New List(Of Object)
                For Each projCut As ProjectedCut In sketch.ProjectedCuts
                    Try
                        projectedGeoToDelete.Add(projCut)
                    Catch
                    End Try
                Next
                For Each pg In projectedGeoToDelete
                    Try
                        CType(pg, ProjectedCut).Delete()
                        UtilsLib.LogInfo("  Deleted external ProjectedCut in sketch: " & sketch.Name)
                    Catch
                    End Try
                Next
            Catch
            End Try
        Next
        
        ' 6. Verify all external references are broken
        Dim remainingRefs As Integer = partDoc.ReferencedDocuments.Count
        If remainingRefs > 0 Then
            UtilsLib.LogWarn("  WARNING: " & remainingRefs & " external references remain after link breaking:")
            For Each refDoc As Document In partDoc.ReferencedDocuments
                UtilsLib.LogWarn("    -> " & System.IO.Path.GetFileName(refDoc.FullFileName))
            Next
        Else
            UtilsLib.LogInfo("  All external links broken successfully")
        End If
    End Sub
    
    ''' <summary>
    ''' Create assembly snapshot with updated references using SaveAs for new GUID.
    ''' </summary>
    Public Function CreateAssemblySnapshot(app As Inventor.Application, _
                                            sourceAsmPath As String, _
                                            targetPath As String, _
                                            referenceMap As Dictionary(Of String, String), _
                                            variantParams As Dictionary(Of String, String), _
                                            newPartNumber As String, _
                                            Optional projectedDescription As String = Nothing, _
                                            Optional savedTransforms As Dictionary(Of String, Matrix) = Nothing) As Boolean
        Try
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(targetPath))
            
            ' Open source assembly (might already be open)
            Dim asmDoc As AssemblyDocument = Nothing
            Dim wasAlreadyOpen As Boolean = False
            
            For Each doc As Document In app.Documents
                If doc.FullFileName.Equals(sourceAsmPath, StringComparison.OrdinalIgnoreCase) Then
                    asmDoc = CType(doc, AssemblyDocument)
                    wasAlreadyOpen = True
                    Exit For
                End If
            Next
            
            If asmDoc Is Nothing Then
                asmDoc = CType(app.Documents.Open(sourceAsmPath, True), AssemblyDocument)
            End If
            
            Try
                ' Log the reference map being used
                UtilsLib.LogInfo("  Reference map has " & referenceMap.Count & " entries")
                
                ' Replace component references to point to new parts
                Dim occsToProcess As New List(Of ComponentOccurrence)
                CollectAllOccurrences(asmDoc.ComponentDefinition.Occurrences, occsToProcess)
                
                UtilsLib.LogInfo("  Assembly has " & occsToProcess.Count & " component occurrences")
                
                ' Build set of target paths to detect already-replaced references
                Dim targetPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each kvp In referenceMap
                    targetPaths.Add(kvp.Value)
                Next
                
                ' Build reverse map: target path -> source path (for transform lookup)
                Dim reverseMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For Each kvp In referenceMap
                    reverseMap(kvp.Value) = kvp.Key
                Next
                
                For Each occ As ComponentOccurrence In occsToProcess
                    Try
                        Dim currentPath As String = occ.Definition.Document.FullFileName
                        If referenceMap.ContainsKey(currentPath) Then
                            occ.Replace(referenceMap(currentPath), True)
                            UtilsLib.LogInfo("    Replaced: " & System.IO.Path.GetFileName(currentPath) & " -> " & System.IO.Path.GetFileName(referenceMap(currentPath)))
                        ElseIf targetPaths.Contains(currentPath) Then
                            ' Already points to a target path (was auto-updated by Inventor after a previous Replace)
                            UtilsLib.LogInfo("    Already target: " & System.IO.Path.GetFileName(currentPath))
                        Else
                            UtilsLib.LogWarn("    NOT IN MAP: " & currentPath)
                        End If
                    Catch ex As Exception
                        UtilsLib.LogWarn("    Replace failed: " & ex.Message)
                    End Try
                Next
                
                ' Apply saved transforms if provided
                ' This restores the correct positions that were captured when variant params were applied
                If savedTransforms IsNot Nothing AndAlso savedTransforms.Count > 0 Then
                    UtilsLib.LogInfo("  Applying saved occurrence transforms...")
                    Dim appliedCount As Integer = 0
                    For Each occ As ComponentOccurrence In occsToProcess
                        Try
                            Dim currentPath As String = occ.Definition.Document.FullFileName
                            
                            ' Try to find the transform using current path (if target) or original path
                            Dim origPath As String = currentPath
                            If reverseMap.ContainsKey(currentPath) Then
                                origPath = reverseMap(currentPath)
                            End If
                            
                            ' Key format: "occName|partPath"
                            Dim key As String = occ.Name & "|" & origPath
                            If savedTransforms.ContainsKey(key) Then
                                occ.Transformation = savedTransforms(key)
                                appliedCount += 1
                            End If
                        Catch
                        End Try
                    Next
                    UtilsLib.LogInfo("    Applied " & appliedCount & "/" & occsToProcess.Count & " transforms")
                End If
                
                ' CRITICAL: Force full assembly update to recalculate constraint positions
                ' After reference replacements (or SaveAs-induced reference changes), the assembly's
                ' constraint solver needs to recalculate all positions based on new part geometries.
                ' Update2(True) forces a recursive update of all components and constraints.
                Try
                    UtilsLib.LogInfo("  Updating assembly constraints...")
                    asmDoc.Update2(True)  ' True = rebuild all
                    
                    ' Also explicitly rebuild if needed
                    Try
                        asmDoc.ComponentDefinition.RepairComponents(True)  ' Attempt to repair/resolve any constraint issues
                    Catch
                        ' RepairComponents may not be available or applicable
                    End Try
                Catch ex As Exception
                    UtilsLib.LogWarn("  Update warning: " & ex.Message)
                End Try
                
                ' Apply variant parameters
                If variantParams IsNot Nothing Then
                    ApplyParameters(asmDoc, variantParams)
                End If
                
                ' DEBUG: Log current document state before setting properties
                UtilsLib.LogInfo("  DEBUG: Before property set - Doc=" & asmDoc.FullFileName)
                
                ' Set Part Number in Design Tracking Properties
                Try
                    Dim designProps = asmDoc.PropertySets.Item("Design Tracking Properties")
                    Dim oldPN As String = designProps.Item("Part Number").Value.ToString()
                    UtilsLib.LogInfo("  DEBUG: Old Part Number=" & oldPN)
                    designProps.Item("Part Number").Value = newPartNumber
                    UtilsLib.LogInfo("  Set Part Number: " & newPartNumber)
                    
                    ' Set Description if projected description provided
                    If Not String.IsNullOrEmpty(projectedDescription) Then
                        Dim oldDesc As String = ""
                        Try : oldDesc = designProps.Item("Description").Value.ToString() : Catch : End Try
                        UtilsLib.LogInfo("  DEBUG: Old Description=" & oldDesc)
                        designProps.Item("Description").Value = projectedDescription
                        UtilsLib.LogInfo("  Set Description: " & projectedDescription)
                    End If
                    
                    ' Verify immediately
                    Dim verifyPN As String = designProps.Item("Part Number").Value.ToString()
                    UtilsLib.LogInfo("  DEBUG: Verify Part Number=" & verifyPN)
                Catch ex As Exception
                    UtilsLib.LogWarn("  WARNING: Failed to set Part Number: " & ex.Message)
                End Try
                
                ' Also set Title in Summary Information (title blocks often use this)
                ' Title should be the description, not the part number
                Try
                    Dim summaryProps = asmDoc.PropertySets.Item("Inventor Summary Information")
                    Dim oldTitle As String = summaryProps.Item("Title").Value.ToString()
                    UtilsLib.LogInfo("  DEBUG: Old Title=" & oldTitle)
                    Dim newTitle As String = If(Not String.IsNullOrEmpty(projectedDescription), projectedDescription, newPartNumber)
                    summaryProps.Item("Title").Value = newTitle
                    UtilsLib.LogInfo("  Set Title: " & newTitle)
                    
                    ' Verify immediately
                    Dim verifyTitle As String = summaryProps.Item("Title").Value.ToString()
                    UtilsLib.LogInfo("  DEBUG: Verify Title=" & verifyTitle)
                Catch ex As Exception
                    UtilsLib.LogWarn("  WARNING: Failed to set Title: " & ex.Message)
                End Try
                
                asmDoc.Update()
                
                ' Use SaveAs to create NEW file with NEW GUID (not File.Copy which preserves GUID)
                ' This allows both source and target to be open simultaneously without conflicts
                UtilsLib.LogInfo("  DEBUG: Before SaveAs - Doc=" & asmDoc.FullFileName)
                asmDoc.SaveAs(targetPath, False)
                UtilsLib.LogInfo("  SaveAs with new GUID: " & System.IO.Path.GetFileName(targetPath))
                UtilsLib.LogInfo("  DEBUG: After SaveAs - Doc=" & asmDoc.FullFileName)
                
                ' Verify properties after SaveAs
                Try
                    Dim pnAfter As String = asmDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                    Dim titleAfter As String = asmDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                    UtilsLib.LogInfo("  DEBUG: After SaveAs - PN=" & pnAfter & ", Title=" & titleAfter)
                Catch
                End Try
                
                ' Document is now the target file - save to ensure all changes are committed
                asmDoc.Save()
                
                ' Verify properties after final save
                Try
                    Dim pnFinal As String = asmDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                    Dim titleFinal As String = asmDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                    UtilsLib.LogInfo("  DEBUG: After Save - PN=" & pnFinal & ", Title=" & titleFinal)
                Catch
                End Try
                
                ' Rename occurrences with descriptive names: "Description (PartNumber):instance"
                ' This removes old part numbers from occurrence names after reference replacement
                Try
                    Dim renamedCount As Integer = OccurrenceNamingLib.RenameAllOccurrences(asmDoc)
                    If renamedCount > 0 Then
                        UtilsLib.LogInfo("  Renamed " & renamedCount & " occurrences with descriptive names")
                        asmDoc.Save()
                    End If
                Catch ex As Exception
                    UtilsLib.LogWarn("  WARNING: Could not rename occurrences - " & ex.Message)
                End Try

                UtilsLib.LogInfo("Created assembly: " & System.IO.Path.GetFileName(targetPath))

                ' Close the target document (source will be closed at end of ExecuteRelease)
                asmDoc.Close(True)
                
                Return True
                
            Catch ex As Exception
                Throw
            End Try
            
        Catch ex As Exception
            UtilsLib.LogError("ERROR creating assembly: " & ex.Message)
            Return False
        End Try
    End Function
    
    Private Sub CollectAllOccurrences(occs As ComponentOccurrences, ByRef result As List(Of ComponentOccurrence))
        For Each occ As ComponentOccurrence In occs
            result.Add(occ)
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
    
    ''' <summary>
    ''' Create drawing copy with updated references using SaveAs for new GUID.
    ''' 
    ''' IMPORTANT: Due to Inventor API limitations, FileDescriptor.ReplaceReference requires
    ''' that source and target files share the same InternalName (GUID). Since we use SaveAs
    ''' to create assemblies with new GUIDs (to avoid document conflicts), the ancestry chain
    ''' is broken. The ONLY reliable workaround is to:
    ''' 1. Replace references in source drawing
    ''' 2. Save and CLOSE the drawing
    ''' 3. REOPEN the drawing - this forces Inventor to re-resolve all references from disk
    ''' 4. Now ReferencedDocuments will properly reflect the new model properties
    ''' 5. SaveAs to create final drawing with new GUID
    ''' 
    ''' See: https://forums.autodesk.com/t5/inventor-programming-forum/filedescriptor-replacereference-method-quot-must-share-ancestry/td-p/13723228
    ''' </summary>
    Public Function CreateDrawingCopy(app As Inventor.Application, _
                                       sourceDrawingPath As String, _
                                       targetPath As String, _
                                       referenceMap As Dictionary(Of String, String), _
                                       newPartNumber As String, _
                                       Optional projectedDescription As String = Nothing) As Boolean
        Try
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(targetPath))
            
            ' Use a temporary path for the intermediate save (before final SaveAs)
            Dim tempPath As String = System.IO.Path.Combine( _
                System.IO.Path.GetDirectoryName(targetPath), _
                "_TEMP_" & System.IO.Path.GetFileName(targetPath))
            
            ' ============================================================
            ' PHASE 1: Open source, replace references, save to temp file
            ' ============================================================
            
            ' Open source drawing
            Dim drawDoc As DrawingDocument = CType(app.Documents.Open(sourceDrawingPath, True), DrawingDocument)
            
            ' Log reference map contents for debugging
            UtilsLib.LogInfo("  Reference map entries (" & referenceMap.Count & " total):")
            For Each kvp In referenceMap
                UtilsLib.LogInfo("    " & System.IO.Path.GetFileName(kvp.Key) & " -> " & System.IO.Path.GetFileName(kvp.Value))
            Next
            
            ' Log current drawing references before replacement
            UtilsLib.LogInfo("  Drawing refs before replacement:")
            For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
                Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
                Dim inMap As String = If(referenceMap.ContainsKey(fd.FullFileName), "-> " & System.IO.Path.GetFileName(referenceMap(fd.FullFileName)), "(NOT IN MAP - this is a problem!)")
                UtilsLib.LogInfo("    " & System.IO.Path.GetFileName(fd.FullFileName) & " " & inMap)
            Next
            
            ' Replace references to released files
            For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
                Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
                Dim oldPath As String = fd.FullFileName
                If referenceMap.ContainsKey(oldPath) Then
                    Dim newPath As String = referenceMap(oldPath)
                    Try
                        fd.ReplaceReference(newPath)
                        UtilsLib.LogInfo("  Replaced: " & System.IO.Path.GetFileName(oldPath) & " -> " & System.IO.Path.GetFileName(newPath))
                    Catch ex As Exception
                        UtilsLib.LogWarn("  WARNING: ReplaceReference failed for " & System.IO.Path.GetFileName(oldPath) & " -> " & System.IO.Path.GetFileName(newPath) & ": " & ex.Message)
                    End Try
                End If
            Next
            
            ' Verify references after replacement
            UtilsLib.LogInfo("  Drawing refs AFTER replacement:")
            For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
                Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
                UtilsLib.LogInfo("    " & System.IO.Path.GetFileName(fd.FullFileName))
            Next
            
            ' Save to temp file (this commits the reference changes)
            drawDoc.SaveAs(tempPath, False)
            UtilsLib.LogInfo("  Saved temp file: " & System.IO.Path.GetFileName(tempPath))
            
            ' CRITICAL: Close the drawing - this clears all internal caches
            drawDoc.Close(True)
            UtilsLib.LogInfo("  Closed drawing to force reference re-resolution")
            
            ' ============================================================
            ' PHASE 2: Reopen temp file - Inventor now re-resolves references from disk
            ' ============================================================
            
            ' Make sure the new referenced models are not already open with stale state
            ' Close ALL documents to ensure clean slate
            For i As Integer = app.Documents.Count To 1 Step -1
                Try
                    Dim doc As Document = app.Documents.Item(i)
                    doc.Close(True)
                Catch
                End Try
            Next
            
            ' Reopen the temp drawing - this forces Inventor to load fresh references
            drawDoc = CType(app.Documents.Open(tempPath, True), DrawingDocument)
            UtilsLib.LogInfo("  Reopened drawing - references now resolved from disk")
            
            ' Verify the drawing now sees correct properties
            For Each refDoc As Document In drawDoc.ReferencedDocuments
                Dim refPN As String = ""
                Dim refTitle As String = ""
                Try
                    refPN = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                    refTitle = refDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                Catch
                End Try
                UtilsLib.LogInfo("  After reopen ref check: " & System.IO.Path.GetFileName(refDoc.FullFileName) & " PN=" & refPN & " Title=" & refTitle)
            Next
            
            ' ============================================================
            ' PHASE 3: Set drawing properties using KNOWN correct values
            ' ============================================================
            
            ' CRITICAL: Do NOT try to read Part Number from the drawing's cached references!
            ' The ReferencedDocuments, ReferencedDocumentDescriptors, and DrawingView.ReferencedDocument
            ' all return STALE cached property values even after ReplaceReference and close/reopen.
            ' This is a documented Inventor caching issue with no reliable workaround.
            '
            ' Instead, we use the newPartNumber parameter which is the KNOWN correct value
            ' from the release plan. This ensures:
            ' 1. Released drawings are self-contained and don't depend on cached state
            ' 2. The title block shows the correct Part Number even if linked to "model" properties
            ' 3. The release is reproducible and not affected by Inventor's internal caches
            
            UtilsLib.LogInfo("  Using known Part Number from release plan: " & newPartNumber)
            
            ' Set drawing's own Part Number and Description (for title blocks)
            Try
                Dim designProps = drawDoc.PropertySets.Item("Design Tracking Properties")
                designProps.Item("Part Number").Value = newPartNumber
                UtilsLib.LogInfo("  Set Drawing Part Number: " & newPartNumber)
                
                ' Set Description if projected description provided
                If Not String.IsNullOrEmpty(projectedDescription) Then
                    designProps.Item("Description").Value = projectedDescription
                    UtilsLib.LogInfo("  Set Drawing Description: " & projectedDescription)
                End If
            Catch ex As Exception
                UtilsLib.LogWarn("  WARNING: Failed to set Drawing Part Number: " & ex.Message)
            End Try
            
            ' Also set Title in Summary Information (use description, not part number)
            Try
                Dim summaryProps = drawDoc.PropertySets.Item("Inventor Summary Information")
                Dim newTitle As String = If(Not String.IsNullOrEmpty(projectedDescription), projectedDescription, newPartNumber)
                summaryProps.Item("Title").Value = newTitle
                UtilsLib.LogInfo("  Set Drawing Title: " & newTitle)
            Catch ex As Exception
                UtilsLib.LogWarn("  WARNING: Failed to set Drawing Title: " & ex.Message)
            End Try
            
            ' Try to force update of copied model iProperties (ribbon command)
            ' This should sync model properties to drawing properties
            Try
                Dim oControlDef As ControlDefinition = app.CommandManager.ControlDefinitions.Item("UpdateCopiedModeliPropertiesCmd")
                oControlDef.Execute2(True)
                UtilsLib.LogInfo("  Executed UpdateCopiedModeliPropertiesCmd")
            Catch ex As Exception
                ' Command may not exist or may fail silently - this is OK
            End Try
            
            ' Update document and all sheets to refresh views and title blocks
            drawDoc.Update()
            For Each sheet As Sheet In drawDoc.Sheets
                Try
                    sheet.Update()
                Catch
                End Try
            Next
            
            ' SaveAs to final path with NEW GUID
            drawDoc.SaveAs(targetPath, False)
            UtilsLib.LogInfo("  SaveAs with new GUID: " & System.IO.Path.GetFileName(targetPath))
            
            ' Save to ensure all changes are committed
            drawDoc.Save()
            
            UtilsLib.LogInfo("Created drawing: " & System.IO.Path.GetFileName(targetPath))
            
            ' Close the drawing
            drawDoc.Close(True)
            
            ' Delete the temp file
            Try
                System.IO.File.Delete(tempPath)
            Catch
            End Try
            
            Return True
            
        Catch ex As Exception
            UtilsLib.LogError("ERROR creating drawing: " & ex.Message)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Execute the complete release plan.
    ''' </summary>
    Public Function ExecuteRelease(app As Inventor.Application, context As ElementReleaseContext) As Boolean
        UtilsLib.LogInfo("ExecuteRelease: Starting...")
        
        Try
            If Not DEVELOPMENT_MODE Then
                UtilsLib.LogInfo("ExecuteRelease: Production mode - would disconnect from Vault here")
            End If
            
            ' CRITICAL: Close ALL source documents BEFORE creating parts
            ' This prevents the SaveAs GUID issue where parts become "detached" from the assembly
            ' When the assembly is open, its parts are also open. If we SaveAs a part while
            ' it's open as part of the assembly, the assembly's reference switches to the target.
            UtilsLib.LogInfo("ExecuteRelease: Closing source documents before processing...")
            Dim sourceFolder As String = context.SourceRoot
            For i As Integer = app.Documents.Count To 1 Step -1
                Try
                    Dim doc As Document = app.Documents.Item(i)
                    If doc.FullFileName.StartsWith(sourceFolder, StringComparison.OrdinalIgnoreCase) Then
                        doc.Close(True) ' Discard changes
                        UtilsLib.LogInfo("  Pre-closed: " & System.IO.Path.GetFileName(doc.FullFileName))
                    End If
                Catch
                End Try
            Next
            
            ' Track which parts have been created (shared parts only created once)
            Dim processedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            
            ' ============================================================
            ' NEW ARCHITECTURE: Copy-First, Modify-Copies-Only
            ' 
            ' See docs/research/2026-05-13-element-release-failures.md
            ' 
            ' STAGE 1: Create file copies with correct references
            '   - Copy masters, parts, assemblies to target locations
            '   - Update references in copies: source files → copied files
            '   - NO parameter changes yet
            ' 
            ' STAGE 2: Apply parameters to COPIES only
            '   - Open copied master
            '   - Apply variant parameters
            '   - Update copied master and assembly
            '   - Save all
            ' 
            ' Key principle: Source files are NEVER modified
            ' ============================================================
            For Each variantCfg As ExcelReaderLib.ElementConfig In context.Elements
                UtilsLib.LogInfo("ExecuteRelease: Processing variant '" & variantCfg.ElementName & "'...")
                
                ' Build reference map for this variant
                Dim refMap = BuildReferenceMapForVariant(context, variantCfg.ElementName)
                
                ' ================================================================
                ' STAGE 1: Create file copies (NO parameter changes)
                ' ================================================================
                UtilsLib.LogInfo("  STAGE 1: Creating file copies...")
                
                ' Step 1a: Copy masters first in DEPENDENCY ORDER (roots first, then dependents)
                ' This ensures when we copy a master that depends on another master,
                ' the dependency is already copied and we can update the reference.
                UtilsLib.LogInfo("    Copying masters in dependency order...")
                Dim copiedMasters As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                
                ' Get masters for this variant in sorted order (already sorted by GetMasterPaths)
                Dim mastersForVariant As New List(Of PlannedFile)
                For Each masterPath In context.MasterPaths
                    Dim file As PlannedFile = context.ReleasePlan.Files.FirstOrDefault(Function(f) _
                        f.SourcePath.Equals(masterPath, StringComparison.OrdinalIgnoreCase) AndAlso _
                        f.ForVariants.Contains(variantCfg.ElementName) AndAlso _
                        Not f.IsExisting)
                    If file IsNot Nothing AndAlso Not processedParts.Contains(file.TargetLocalPath) Then
                        mastersForVariant.Add(file)
                    End If
                Next
                
                UtilsLib.LogInfo("      Masters to copy: " & mastersForVariant.Count)
                
                For Each file As PlannedFile In mastersForVariant
                    Try
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(file.TargetLocalPath))
                        
                        ' Open source, SaveAs to target with new GUID
                        Dim masterDoc As Document = app.Documents.Open(file.SourcePath, False)
                        masterDoc.SaveAs(file.TargetLocalPath, False)
                        UtilsLib.LogInfo("      Copied master: " & System.IO.Path.GetFileName(file.TargetLocalPath))
                        
                        ' CRITICAL: Update references to previously copied masters
                        ' This handles master-to-master dependencies (e.g., 000131 derives from 000130)
                        Dim updatedRefs As Integer = 0
                        For Each refDesc As DocumentDescriptor In masterDoc.ReferencedDocumentDescriptors
                            Try
                                Dim refPath As String = refDesc.ReferencedFileDescriptor.FullFileName
                                If copiedMasters.ContainsKey(refPath) Then
                                    refDesc.ReferencedFileDescriptor.ReplaceReference(copiedMasters(refPath))
                                    updatedRefs += 1
                                    UtilsLib.LogInfo("        Updated cross-master ref: " & System.IO.Path.GetFileName(refPath) & " -> " & System.IO.Path.GetFileName(copiedMasters(refPath)))
                                End If
                            Catch
                            End Try
                        Next
                        If updatedRefs > 0 Then
                            UtilsLib.LogInfo("        Updated " & updatedRefs & " cross-master reference(s)")
                        End If
                        
                        ' Also update occurrence references if this is an assembly (intermediate assembly)
                        If masterDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            Dim asmMaster As AssemblyDocument = CType(masterDoc, AssemblyDocument)
                            Dim allOccs As New List(Of ComponentOccurrence)
                            CollectAllOccurrences(asmMaster.ComponentDefinition.Occurrences, allOccs)
                            
                            For Each occ As ComponentOccurrence In allOccs
                                Try
                                    Dim currentPath As String = occ.Definition.Document.FullFileName
                                    If copiedMasters.ContainsKey(currentPath) Then
                                        occ.Replace(copiedMasters(currentPath), True)
                                        UtilsLib.LogInfo("        Replaced occurrence: " & System.IO.Path.GetFileName(currentPath) & " -> " & System.IO.Path.GetFileName(copiedMasters(currentPath)))
                                    End If
                                Catch
                                End Try
                            Next
                        End If
                        
                        ' Set properties on the copy
                        SetDocumentProperties(masterDoc, file.VaultNumber, file.ProjectedDescription)
                        masterDoc.Save()
                        
                        copiedMasters(file.SourcePath) = file.TargetLocalPath
                        processedParts.Add(file.TargetLocalPath)
                        
                        masterDoc.Close(True)
                        
                        Try
                            ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.Completed)
                        Catch : End Try
                    Catch ex As Exception
                        UtilsLib.LogError("      Failed to copy master: " & ex.Message)
                    End Try
                Next
                
                ' Step 1b: Copy all parts and update their master references
                UtilsLib.LogInfo("    Copying parts and updating master references...")
                For Each file As PlannedFile In context.ReleasePlan.Files
                    If file.FileType <> FileType.Part Then Continue For
                    If context.MasterPaths.Contains(file.SourcePath) Then Continue For ' Skip masters
                    If Not file.ForVariants.Contains(variantCfg.ElementName) Then Continue For
                    If file.IsExisting Then Continue For
                    If processedParts.Contains(file.TargetLocalPath) Then
                        UtilsLib.LogInfo("      Already created: " & System.IO.Path.GetFileName(file.TargetLocalPath))
                        Continue For
                    End If
                    
                    Try
                        ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.InProgress)
                    Catch : End Try
                    
                    Try
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(file.TargetLocalPath))
                        
                        ' Open source part
                        Dim partDoc As PartDocument = CType(app.Documents.Open(file.SourcePath, False), PartDocument)
                        
                        ' SaveAs to target with new GUID
                        partDoc.SaveAs(file.TargetLocalPath, False)
                        UtilsLib.LogInfo("      Copied: " & System.IO.Path.GetFileName(file.TargetLocalPath))
                        
                        ' Update references: source master → copied master
                        For Each refDesc As DocumentDescriptor In partDoc.ReferencedDocumentDescriptors
                            Try
                                Dim refPath As String = refDesc.ReferencedFileDescriptor.FullFileName
                                If copiedMasters.ContainsKey(refPath) Then
                                    refDesc.ReferencedFileDescriptor.ReplaceReference(copiedMasters(refPath))
                                    UtilsLib.LogInfo("        Updated ref: " & System.IO.Path.GetFileName(refPath) & " -> " & System.IO.Path.GetFileName(copiedMasters(refPath)))
                                ElseIf refMap.ContainsKey(refPath) Then
                                    refDesc.ReferencedFileDescriptor.ReplaceReference(refMap(refPath))
                                    UtilsLib.LogInfo("        Updated ref: " & System.IO.Path.GetFileName(refPath) & " -> " & System.IO.Path.GetFileName(refMap(refPath)))
                                End If
                            Catch ex As Exception
                                UtilsLib.LogWarn("        Ref update failed: " & ex.Message)
                            End Try
                        Next
                        
                        ' Set properties on the copy
                        SetDocumentProperties(partDoc, file.VaultNumber, file.ProjectedDescription)
                        partDoc.Save()
                        
                        processedParts.Add(file.TargetLocalPath)
                        partDoc.Close(True)
                        
                        Try
                            ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.Completed)
                        Catch : End Try
                    Catch ex As Exception
                        UtilsLib.LogError("      Failed to copy part: " & ex.Message)
                        Try
                            ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.Failed)
                        Catch : End Try
                    End Try
                Next
                
                ' Step 1c: Copy assembly and update all part references
                UtilsLib.LogInfo("    Copying assembly and updating part references...")
                Dim copiedAsmPath As String = Nothing
                For Each file As PlannedFile In context.ReleasePlan.Files
                    If file.FileType <> FileType.Assembly Then Continue For
                    If Not file.ForVariants.Contains(variantCfg.ElementName) Then Continue For
                    If file.IsExisting Then Continue For
                    
                    Try
                        ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.InProgress)
                    Catch : End Try
                    
                    Try
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(file.TargetLocalPath))
                        
                        ' Open source assembly
                        Dim asmDoc As AssemblyDocument = CType(app.Documents.Open(file.SourcePath, False), AssemblyDocument)
                        
                        ' SaveAs to target with new GUID
                        asmDoc.SaveAs(file.TargetLocalPath, False)
                        copiedAsmPath = file.TargetLocalPath
                        UtilsLib.LogInfo("      Copied: " & System.IO.Path.GetFileName(file.TargetLocalPath))
                        
                        ' Replace all component references
                        Dim allOccs As New List(Of ComponentOccurrence)
                        CollectAllOccurrences(asmDoc.ComponentDefinition.Occurrences, allOccs)
                        
                        For Each occ As ComponentOccurrence In allOccs
                            Try
                                Dim currentPath As String = occ.Definition.Document.FullFileName
                                If refMap.ContainsKey(currentPath) Then
                                    occ.Replace(refMap(currentPath), True)
                                    UtilsLib.LogInfo("        Replaced: " & System.IO.Path.GetFileName(currentPath) & " -> " & System.IO.Path.GetFileName(refMap(currentPath)))
                                End If
                            Catch ex As Exception
                                UtilsLib.LogWarn("        Replace failed: " & ex.Message)
                            End Try
                        Next
                        
                        ' Set properties on the copy
                        SetDocumentProperties(asmDoc, file.VaultNumber, file.ProjectedDescription)
                        
                        ' Rename occurrences
                        Try
                            OccurrenceNamingLib.RenameAllOccurrences(asmDoc)
                        Catch : End Try
                        
                        asmDoc.Save()
                        asmDoc.Close(True)
                        
                        Try
                            ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.Completed)
                        Catch : End Try
                    Catch ex As Exception
                        UtilsLib.LogError("      Failed to copy assembly: " & ex.Message)
                        Try
                            ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.Failed)
                        Catch : End Try
                    End Try
                Next
                
                ' ================================================================
                ' STAGE 2: Apply parameters to COPIES only
                ' ================================================================
                UtilsLib.LogInfo("  STAGE 2: Applying parameters to copies...")
                
                ' Open and update copied masters with variant parameters
                For Each kvp In copiedMasters
                    Try
                        Dim masterDoc As Document = app.Documents.Open(kvp.Value, False)
                        ApplyParameters(masterDoc, variantCfg.Parameters)
                        masterDoc.Update()
                        masterDoc.Save()
                        UtilsLib.LogInfo("    Applied params and saved: " & System.IO.Path.GetFileName(kvp.Value))
                        masterDoc.Close(True)
                    Catch ex As Exception
                        UtilsLib.LogError("    Failed to apply params to master: " & ex.Message)
                    End Try
                Next
                
                ' Open copied assembly and update (propagates through references)
                If copiedAsmPath IsNot Nothing Then
                    Try
                        Dim asmDoc As AssemblyDocument = CType(app.Documents.Open(copiedAsmPath, True), AssemblyDocument)
                        asmDoc.Update2(True)
                        asmDoc.Save()
                        UtilsLib.LogInfo("    Updated and saved assembly: " & System.IO.Path.GetFileName(copiedAsmPath))
                        asmDoc.Close(True)
                    Catch ex As Exception
                        UtilsLib.LogError("    Failed to update assembly: " & ex.Message)
                    End Try
                End If
                
                ' Close all documents before next variant
                UtilsLib.LogInfo("  Closing all documents...")
                For i As Integer = app.Documents.Count To 1 Step -1
                    Try
                        app.Documents.Item(i).Close(True)
                    Catch : End Try
                Next
            Next
            
            ' ============================================================
            ' STAGE 3: Create drawings
            ' ============================================================
            UtilsLib.LogInfo("ExecuteRelease: Creating drawings...")
            For Each file As PlannedFile In context.ReleasePlan.Files
                If file.FileType = FileType.Drawing Then
                    If IsOldVersionsPath(file.SourcePath) OrElse IsOldVersionsPath(file.TargetLocalPath) Then
                        UtilsLib.LogWarn("  Skipping OldVersions file: " & System.IO.Path.GetFileName(file.SourcePath))
                        Continue For
                    End If
                    Dim variantName = file.ForVariants(0)
                    Dim refMap = BuildReferenceMapForVariant(context, variantName)
                    
                    Try
                        ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.InProgress)
                    Catch : End Try
                    
                    Dim dwgSuccess As Boolean = CreateDrawingCopy(app, file.SourcePath, file.TargetLocalPath, refMap, file.VaultNumber, file.ProjectedDescription)
                    
                    Try
                        If dwgSuccess Then
                            ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.Completed)
                        Else
                            ElementReleaseUILib.UpdateFileStatus(file.TargetLocalPath, ElementReleaseUILib.FileStatus.Failed)
                        End If
                    Catch : End Try
                End If
            Next
            
            ' Old code removed - using new copy-first architecture above
            ' Close ALL documents to completely clear Inventor's caches
            UtilsLib.LogInfo("ExecuteRelease: Closing all documents to clear caches...")
            For i As Integer = app.Documents.Count To 1 Step -1
                Try
                    Dim doc As Document = app.Documents.Item(i)
                    doc.Close(True)
                Catch
                End Try
            Next
            
            ' ============================================================
            ' CRITICAL FIX-UP PASS: Reopen drawings fresh and fix properties
            ' ============================================================
            ' After all files are created and all documents are closed, Inventor's
            ' internal caches are cleared. Now we can reopen each drawing fresh,
            ' which forces Inventor to load references from disk with correct properties.
            ' We then re-set the drawing's own properties to ensure title blocks
            ' show correct values regardless of their link source (model vs drawing).
            
            ' ============================================================
            ' FIX-UP PASS: Force correct properties in drawings
            ' ============================================================
            ' STRATEGY: The title block links to "model properties" but Inventor's
            ' internal cache shows wrong values even after reference replacement.
            ' 
            ' The model files on disk have correct properties (verified by fresh open).
            ' We need to make the drawing save with correct embedded property values.
            '
            ' Approach:
            ' 1. First, open all models (parts/assemblies) fresh - they'll have correct properties
            ' 2. Then open each drawing - it will use the already-open models
            ' 3. Update and save the drawing
            
            UtilsLib.LogInfo("ExecuteRelease: Fix-up pass...")
            Try
                ElementReleaseUILib.LogMessage("")
                ElementReleaseUILib.LogMessage("Fix-up pass: Verifying drawing references...")
            Catch : End Try
            
            ' Step 1: Open ALL models first (so they're in memory with correct properties)
            UtilsLib.LogInfo("  Opening models fresh from disk...")
            Dim openedModels As New List(Of Document)
            For Each file As PlannedFile In context.ReleasePlan.Files
                If file.FileType = FileType.Drawing Then Continue For
                If file.IsExisting Then Continue For
                
                Try
                    Dim doc As Document = app.Documents.Open(file.TargetLocalPath, True)
                    Dim pn As String = doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                    UtilsLib.LogInfo("    " & System.IO.Path.GetFileName(file.TargetLocalPath) & " PN=" & pn)
                    openedModels.Add(doc)
                Catch ex As Exception
                    UtilsLib.LogWarn("    Failed to open " & System.IO.Path.GetFileName(file.TargetLocalPath) & ": " & ex.Message)
                End Try
            Next
            
            ' Step 2: Now open each drawing (models are already in memory with correct properties)
            UtilsLib.LogInfo("  Processing drawings...")
            For Each file As PlannedFile In context.ReleasePlan.Files
                If file.FileType <> FileType.Drawing Then Continue For
                If file.IsExisting Then Continue For
                
                Dim targetPath As String = file.TargetLocalPath
                Dim expectedPN As String = file.VaultNumber
                
                Try
                    Dim drawDoc As DrawingDocument = CType(app.Documents.Open(targetPath, True), DrawingDocument)
                    
                    ' Check and fix drawing's reference cache
                    ' The ReferencedDocuments collection may have stale cached properties
                    ' We can update them directly through this reference
                    For Each refDoc As Document In drawDoc.ReferencedDocuments
                        Try
                            Dim refPN As String = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                            Dim refFileName As String = System.IO.Path.GetFileNameWithoutExtension(refDoc.FullFileName)
                            UtilsLib.LogInfo("    " & System.IO.Path.GetFileName(targetPath) & " ref=" & refFileName & " PN=" & refPN)
                            
                            ' If the cached PN doesn't match the filename, fix it
                            If refPN <> refFileName Then
                                refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = refFileName
                                refDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value = refFileName
                                UtilsLib.LogInfo("      -> Fixed ref cache to: " & refFileName)
                            End If
                        Catch ex As Exception
                            UtilsLib.LogWarn("      -> Error: " & ex.Message)
                        End Try
                    Next
                    
                    ' Set drawing's own properties
                    Try
                        drawDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = expectedPN
                        drawDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value = expectedPN
                    Catch
                    End Try
                    
                    ' Execute UpdateCopiedModeliPropertiesCmd
                    Try
                        Dim oControlDef As ControlDefinition = app.CommandManager.ControlDefinitions.Item("UpdateCopiedModeliPropertiesCmd")
                        oControlDef.Execute2(True)
                    Catch
                    End Try
                    
                    ' Update drawing
                    drawDoc.Update()
                    For Each sheet As Sheet In drawDoc.Sheets
                        Try : sheet.Update() : Catch : End Try
                    Next
                    
                    ' Save
                    drawDoc.Save()
                    UtilsLib.LogInfo("    Saved: " & System.IO.Path.GetFileName(targetPath) & " PN=" & expectedPN)
                    
                    drawDoc.Close(False)
                    
                Catch ex As Exception
                    UtilsLib.LogWarn("    Fix-up error: " & System.IO.Path.GetFileName(targetPath) & ": " & ex.Message)
                End Try
            Next
            
            ' Close all the models we opened
            For Each doc As Document In openedModels
                Try : doc.Close(False) : Catch : End Try
            Next
            
            ' Close everything again before verification
            For i As Integer = app.Documents.Count To 1 Step -1
                Try
                    app.Documents.Item(i).Close(True)
                Catch
                End Try
            Next
            
            ' Verify created files
            UtilsLib.LogInfo("ExecuteRelease: Verifying created files...")
            Try
                ElementReleaseUILib.LogMessage("")
                ElementReleaseUILib.LogMessage("Verifying created files...")
            Catch : End Try
            Dim verificationPassed As Boolean = VerifyReleasedFiles(app, context)

            ' NOTE: In the new copy-first architecture, we never modify source masters
            ' so there's nothing to restore. The masterSnapshot variable is unused.
            
            ' Reopen the original source assembly so user is back where they started
            UtilsLib.LogInfo("ExecuteRelease: Reopening documents...")
            Dim originalAssembly As String = context.AssemblyTree.RootAssemblyPath
            Try
                app.Documents.Open(originalAssembly, True)
                UtilsLib.LogInfo("  Reopened source: " & System.IO.Path.GetFileName(originalAssembly))
            Catch ex As Exception
                UtilsLib.LogWarn("  Failed to reopen " & originalAssembly & ": " & ex.Message)
            End Try
            
            ' Also open the released assemblies for each element so user can see the results
            For Each file As PlannedFile In context.ReleasePlan.Files
                If file.FileType = FileType.Assembly AndAlso Not file.IsExisting Then
                    Try
                        app.Documents.Open(file.TargetLocalPath, True)
                        UtilsLib.LogInfo("  Opened released: " & System.IO.Path.GetFileName(file.TargetLocalPath))
                    Catch ex As Exception
                        UtilsLib.LogWarn("  Failed to open " & file.TargetLocalPath & ": " & ex.Message)
                    End Try
                End If
            Next
            
            If verificationPassed Then
                UtilsLib.LogInfo("ExecuteRelease: Complete - all verifications passed!")
            Else
                UtilsLib.LogWarn("ExecuteRelease: Complete - VERIFICATION WARNINGS (see above)")
            End If
            
            Return True
            
        Finally
            ' Cleanup handled above
        End Try
    End Function
    
    ''' <summary>
    ''' Verify that all released files have correct Part Number and Title.
    ''' Opens each file fresh to avoid caching issues.
    ''' </summary>
    Private Function VerifyReleasedFiles(app As Inventor.Application, context As ElementReleaseContext) As Boolean
        Dim allPassed As Boolean = True
        
        For Each file As PlannedFile In context.ReleasePlan.Files
            If file.IsExisting Then Continue For
            
            Dim targetPath As String = file.TargetLocalPath
            Dim expectedPN As String = file.VaultNumber
            
            If Not System.IO.File.Exists(targetPath) Then
                UtilsLib.LogError("  VERIFY FAIL: File not created - " & System.IO.Path.GetFileName(targetPath))
                allPassed = False
                Continue For
            End If
            
            Try
                ' Open file fresh (no caching)
                Dim doc As Document = app.Documents.Open(targetPath, True)
                
                ' Check Part Number
                Dim actualPN As String = ""
                Try
                    actualPN = doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                Catch
                    actualPN = "(error)"
                End Try
                
                ' Check Title
                Dim actualTitle As String = ""
                Try
                    actualTitle = doc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                Catch
                    actualTitle = "(error)"
                End Try
                
                ' For drawings, check both drawing's own PN and referenced documents
                If file.FileType = FileType.Drawing Then
                    Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
                    UtilsLib.LogInfo("  VERIFY Drawing: " & System.IO.Path.GetFileName(targetPath))
                    UtilsLib.LogInfo("    PartNumber=" & actualPN & ", Title=" & actualTitle)
                    
                    ' Verify drawing's own Part Number matches expected
                    If actualPN <> expectedPN Then
                        UtilsLib.LogWarn("    -> MISMATCH: Drawing PN=" & actualPN & " but expected=" & expectedPN)
                        allPassed = False
                    End If
                    
                    For Each refDoc As Document In drawDoc.ReferencedDocuments
                        Dim refPN As String = ""
                        Dim refTitle As String = ""
                        Try
                            refPN = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
                            refTitle = refDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                        Catch
                        End Try
                        
                        Dim refFileName As String = System.IO.Path.GetFileNameWithoutExtension(refDoc.FullFileName)
                        UtilsLib.LogInfo("    -> Ref: " & refDoc.FullFileName)
                        UtilsLib.LogInfo("       PartNumber=" & refPN & ", Title=" & refTitle)
                        
                        ' Check if referenced model has matching Part Number
                        If refPN <> refFileName Then
                            UtilsLib.LogWarn("    -> MISMATCH: File=" & refFileName & " but PartNumber=" & refPN)
                            allPassed = False
                        End If
                    Next
                Else
                    ' Part or Assembly - just check PN matches expected
                    If actualPN = expectedPN Then
                        UtilsLib.LogInfo("  VERIFY OK: " & System.IO.Path.GetFileName(targetPath) & " PN=" & actualPN)
                    Else
                        UtilsLib.LogWarn("  VERIFY FAIL: " & System.IO.Path.GetFileName(targetPath) & " Expected=" & expectedPN & " Actual=" & actualPN)
                        allPassed = False
                    End If
                End If
                
                doc.Close(True)
                
            Catch ex As Exception
                UtilsLib.LogError("  VERIFY ERROR: " & System.IO.Path.GetFileName(targetPath) & " - " & ex.Message)
                allPassed = False
            End Try
        Next
        
        Return allPassed
    End Function
    
    ''' <summary>
    ''' Find a variant configuration by name.
    ''' </summary>
    Private Function FindVariantByName(variants As List(Of ExcelReaderLib.ElementConfig), name As String) As ExcelReaderLib.ElementConfig
        For Each v As ExcelReaderLib.ElementConfig In variants
            If v.ElementName = name Then Return v
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Build reference map for a specific variant.
    ''' </summary>
    Private Function BuildReferenceMapForVariant(context As ElementReleaseContext, variantName As String) As Dictionary(Of String, String)
        Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        
        For Each file As PlannedFile In context.ReleasePlan.Files
            If file.ForVariants.Contains(variantName) Then
                map(file.SourcePath) = file.TargetLocalPath
            End If
        Next
        
        For Each file As PlannedFile In context.ReleasePlan.Files
            If file.IsShared AndAlso Not map.ContainsKey(file.SourcePath) Then
                map(file.SourcePath) = file.TargetLocalPath
            End If
        Next
        
        Return map
    End Function
    
    ' ============================================================================
    ' Manifest Operations
    ' ============================================================================
    
    Public Class ReleaseManifest
        Public LastUpdated As DateTime
        Public Elements As New List(Of ElementEntry)
        Public SharedParts As New List(Of SharedPartEntry)
    End Class
    
    Public Class ElementEntry
        Public ElementName As String
        Public Variants As New List(Of VariantEntry)
        Public ReleaseDate As DateTime
    End Class
    
    Public Class VariantEntry
        Public ReleasedElementName As String
        Public VaultFolder As String
        Public Parts As New List(Of String)
        Public Assemblies As New List(Of String)
        Public Drawings As New List(Of String)
    End Class
    
    Public Class SharedPartEntry
        Public VaultPath As String
        Public VaultNumber As String
        Public SourcePartNumber As String
        Public GeometryFingerprint As String
        Public UsedByElements As New List(Of String)
        Public UsedByVariants As New List(Of String)
        Public ReleaseDate As DateTime
    End Class
    
    Public Function ReadManifest(manifestPath As String) As ReleaseManifest
        If Not System.IO.File.Exists(manifestPath) Then Return Nothing
        
        Try
            Dim json As String = System.IO.File.ReadAllText(manifestPath)
            Return DeserializeManifest(json)
        Catch
            Return Nothing
        End Try
    End Function
    
    Public Sub WriteManifest(manifestPath As String, manifest As ReleaseManifest)
        Try
            manifest.LastUpdated = DateTime.Now
            Dim json As String = SerializeManifest(manifest)
            System.IO.File.WriteAllText(manifestPath, json)
        Catch
        End Try
    End Sub
    
    Private Function SerializeManifest(manifest As ReleaseManifest) As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("{")
        sb.AppendLine("  ""LastUpdated"": """ & manifest.LastUpdated.ToString("o") & """,")
        sb.AppendLine("  ""Modules"": [],")
        sb.AppendLine("  ""SharedParts"": [")
        
        For i As Integer = 0 To manifest.SharedParts.Count - 1
            Dim sp = manifest.SharedParts(i)
            sb.AppendLine("    {")
            sb.AppendLine("      ""VaultPath"": """ & EscapeJson(sp.VaultPath) & """,")
            sb.AppendLine("      ""VaultNumber"": """ & sp.VaultNumber & """,")
            sb.AppendLine("      ""SourcePartNumber"": """ & EscapeJson(sp.SourcePartNumber) & """,")
            sb.AppendLine("      ""GeometryFingerprint"": """ & EscapeJson(sp.GeometryFingerprint) & """,")
            sb.AppendLine("      ""ReleaseDate"": """ & sp.ReleaseDate.ToString("o") & """")
            sb.Append("    }")
            If i < manifest.SharedParts.Count - 1 Then sb.Append(",")
            sb.AppendLine()
        Next
        
        sb.AppendLine("  ]")
        sb.AppendLine("}")
        Return sb.ToString()
    End Function
    
    Private Function DeserializeManifest(json As String) As ReleaseManifest
        Dim manifest As New ReleaseManifest()
        Return manifest
    End Function
    
    Private Function EscapeJson(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""")
    End Function
    
    ' ============================================================================
    ' Utility Functions
    ' ============================================================================
    
    Public Function GetRelativePath(sourceRoot As String, filePath As String) As String
        If String.IsNullOrEmpty(sourceRoot) OrElse String.IsNullOrEmpty(filePath) Then Return filePath
        
        sourceRoot = sourceRoot.TrimEnd("\"c)
        If Not sourceRoot.EndsWith("\") Then sourceRoot &= "\"
        
        If filePath.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase) Then
            Return filePath.Substring(sourceRoot.Length)
        End If
        
        Return filePath
    End Function
    
    Public Function IsInsideSourceRoot(filePath As String, sourceRoot As String) As Boolean
        If String.IsNullOrEmpty(sourceRoot) OrElse String.IsNullOrEmpty(filePath) Then Return False
        
        sourceRoot = sourceRoot.TrimEnd("\"c) & "\"
        Return filePath.StartsWith(sourceRoot, StringComparison.OrdinalIgnoreCase)
    End Function
    
    Public Sub ShowCompletionSummary(plan As ReleasePlan)
        Dim firstNum As String = If(plan.Files.Count > 0, plan.Files(0).VaultNumber, "N/A")
        Dim lastNum As String = If(plan.Files.Count > 0, plan.Files(plan.Files.Count - 1).VaultNumber, "N/A")
        
        Dim sharedCnt As Integer = 0
        Dim variantCnt As Integer = 0
        For Each f As PlannedFile In plan.Files
            If f.IsShared Then sharedCnt += 1 Else variantCnt += 1
        Next
        
        Dim summary As String = "Elementide väljastamine lõpetatud!" & vbCrLf & vbCrLf &
            "Faile loodud: " & plan.Files.Count & vbCrLf &
            "  - Jagatud: " & sharedCnt & vbCrLf &
            "  - Elemendispetsiifilised: " & variantCnt & vbCrLf & vbCrLf &
            "Numbrid: " & firstNum & " kuni " & lastNum
        
        MessageBox.Show(summary, StringsLib.MSG_RELEASE_COMPLETE, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Module
