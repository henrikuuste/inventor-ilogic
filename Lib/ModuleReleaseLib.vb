' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' ModuleReleaseLib - Module Release System Core Library
' 
' Provides functions for releasing parametric Inventor modules with optimal
' file sharing. Analyzes variant parameters, computes geometry fingerprints,
' and creates standalone copies only where geometry differs.
'
' Dependencies:
'   UtilsLib - logging via UtilsLib.LogInfo / UtilsLib.LogWarn
'   ExcelReaderLib - variant table reading (ReleaseConfig)
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
'     AddVbFile "Lib/ModuleReleaseLib.vb"
'
' ============================================================================

Imports Inventor
Imports System.Collections.Generic
Imports System.Windows.Forms

Public Module ModuleReleaseLib

    ' ============================================================================
    ' Configuration Constants
    ' ============================================================================
    
    Public Const DEVELOPMENT_MODE As Boolean = True
    Public Const NUMBERING_SCHEME As String = "Test numbriskeem"
    
    ' ============================================================================
    ' Enums and Data Structures
    ' ============================================================================
    
    Public Enum ReleaseMode
        Cancelled = 0
        FullModule = 1
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
    Public Class ReleaseContext
        Public Mode As ReleaseMode
        Public ExcelPath As String
        Public Variants As List(Of ExcelReaderLib.ReleaseConfig)
        Public SourceRoot As String
        Public TargetRoot As String
        Public ModuleName As String
        Public AssemblyTree As AssemblyTree
        Public VariantMatrix As VariantMatrix
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
    ''' Variant matrix with fingerprints per part per variant.
    ''' </summary>
    Public Class VariantMatrix
        Public PartPaths As New List(Of String)
        Public VariantNames As New List(Of String)
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
        Public VaultNumber As String
        Public FileType As FileType
        Public IsShared As Boolean
        Public IsExisting As Boolean
        Public ForVariants As New List(Of String)
        Public ForModules As New List(Of String)
        Public Fingerprint As String
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
        frm.Text = "Moodulite väljastamine"
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
        lblPrompt.Text = "Vali väljastamise režiim (moodulid Excelist):"
        lblPrompt.Left = 20
        lblPrompt.Top = 20
        lblPrompt.Width = 300
        frm.Controls.Add(lblPrompt)
        
        Dim btnFull As New Button()
        btnFull.Text = "Kõik moodulid"
        btnFull.Left = 20
        btnFull.Top = 50
        btnFull.Width = 140
        btnFull.Height = 35
        AddHandler btnFull.Click, Sub(s, e)
            frm.Tag = ReleaseMode.FullModule
            frm.DialogResult = DialogResult.OK
        End Sub
        frm.Controls.Add(btnFull)
        
        Dim btnCurrent As New Button()
        btnCurrent.Text = "Esimene moodul"
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
    ''' Discover the Excel configuration file for the module.
    ''' Looks for moodulid.xlsx in the module source folder.
    ''' </summary>
    Public Function DiscoverExcel(sourceFolder As String) As String
        If String.IsNullOrEmpty(sourceFolder) OrElse Not System.IO.Directory.Exists(sourceFolder) Then
            UtilsLib.LogInfo("DiscoverExcel: Source folder not found: " & sourceFolder)
            Return Nothing
        End If
        
        UtilsLib.LogInfo("DiscoverExcel: Searching in " & sourceFolder)
        
        ' Look for moodulid.xlsx first
        Dim excelPath As String = System.IO.Path.Combine(sourceFolder, "moodulid.xlsx")
        If System.IO.File.Exists(excelPath) Then
            UtilsLib.LogInfo("DiscoverExcel: Found moodulid.xlsx")
            Return excelPath
        End If
        
        ' Fallback to .xls
        excelPath = System.IO.Path.Combine(sourceFolder, "moodulid.xls")
        If System.IO.File.Exists(excelPath) Then
            UtilsLib.LogInfo("DiscoverExcel: Found moodulid.xls")
            Return excelPath
        End If
        
        UtilsLib.LogInfo("DiscoverExcel: No moodulid.xlsx found")
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Discover release context based on mode selection.
    ''' </summary>
    Public Function DiscoverContext(app As Inventor.Application, mode As ReleaseMode) As ReleaseContext
        If mode = ReleaseMode.Cancelled Then Return Nothing
        
        Dim context As New ReleaseContext()
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
        
        context.SourceRoot = FindModuleSourceRoot(asmPath)
        If String.IsNullOrEmpty(context.SourceRoot) Then
            UtilsLib.LogInfo("DiscoverContext: Could not determine source root")
            Return Nothing
        End If
        
        context.ModuleName = System.IO.Path.GetFileName(context.SourceRoot)
        UtilsLib.LogInfo("DiscoverContext: Module name: " & context.ModuleName)
        
        context.TargetRoot = ComputeTargetRoot(context.SourceRoot)
        UtilsLib.LogInfo("DiscoverContext: Target root: " & context.TargetRoot)
        
        ' Always require Excel file
        context.ExcelPath = DiscoverExcel(context.SourceRoot)
        If String.IsNullOrEmpty(context.ExcelPath) Then
            UtilsLib.LogInfo("DiscoverContext: Excel file required")
            MessageBox.Show("Excel faili (moodulid.xlsx) ei leitud kaustast:" & vbCrLf & _
                           context.SourceRoot & vbCrLf & vbCrLf & _
                           "Loo Excel fail moodulite kirjeldusega." & vbCrLf & _
                           "Vaata malli: Moodulid\_moodulid_template.xlsx", "Loo moodulid")
            Return Nothing
        End If
        
        Dim allVariants As List(Of ExcelReaderLib.ReleaseConfig) = ExcelReaderLib.ReadVariantTable(context.ExcelPath)
        UtilsLib.LogInfo("DiscoverContext: Loaded " & allVariants.Count & " moodulid from Excel")
        
        If allVariants.Count = 0 Then
            UtilsLib.LogInfo("DiscoverContext: No moodulid found in Excel")
            MessageBox.Show("Excel fail on tühi või vigane:" & vbCrLf & _
                           context.ExcelPath, "Loo moodulid")
            Return Nothing
        End If
        
        ' Filter based on mode
        If mode = ReleaseMode.FullModule Then
            context.Variants = allVariants
            UtilsLib.LogInfo("DiscoverContext: Full module mode - using all " & context.Variants.Count & " moodulid")
        Else
            ' CurrentAssembly mode - use only first moodul
            context.Variants = New List(Of ExcelReaderLib.ReleaseConfig)
            context.Variants.Add(allVariants(0))
            UtilsLib.LogInfo("DiscoverContext: Single moodul mode - using: " & allVariants(0).ConfigName)
        End If
        
        Return context
    End Function
    
    ''' <summary>
    ''' Find the module source root folder (Alusmoodulid/{ModuleName}).
    ''' </summary>
    Private Function FindModuleSourceRoot(asmPath As String) As String
        Dim folder As String = System.IO.Path.GetDirectoryName(asmPath)
        
        Do While Not String.IsNullOrEmpty(folder)
            Dim parentName As String = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(folder))
            If parentName.Equals("Alusmoodulid", StringComparison.OrdinalIgnoreCase) Then
                UtilsLib.LogInfo("FindModuleSourceRoot: Found module root: " & folder)
                Return folder
            End If
            folder = System.IO.Path.GetDirectoryName(folder)
        Loop
        
        UtilsLib.LogInfo("FindModuleSourceRoot: Alusmoodulid not found in path, using assembly folder")
        Return System.IO.Path.GetDirectoryName(asmPath)
    End Function
    
    ''' <summary>
    ''' Compute target root folder (Moodulid/ at same level as Alusmoodulid/).
    ''' </summary>
    Private Function ComputeTargetRoot(sourceRoot As String) As String
        Dim parent As String = System.IO.Path.GetDirectoryName(sourceRoot)
        Dim grandparent As String = System.IO.Path.GetDirectoryName(parent)
        
        If System.IO.Path.GetFileName(parent).Equals("Alusmoodulid", StringComparison.OrdinalIgnoreCase) Then
            Return System.IO.Path.Combine(grandparent, "Moodulid")
        End If
        
        Return System.IO.Path.Combine(parent, "Moodulid")
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
            
            For Each refDoc As Document In asmDoc.AllReferencedDocuments
                Dim refPath As String = refDoc.FullFileName
                
                ' Skip files outside source root
                If Not IsInsideSourceRoot(refPath, sourceRoot) Then
                    Continue For
                End If
                
                ' Skip OldVersions folder (special Vault folder)
                If refPath.Contains("\OldVersions\") Then
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
            
            UtilsLib.LogInfo("DiscoverAssemblyTree: Found " & tree.Parts.Count & " parts, " & tree.Assemblies.Count & " assemblies")
            
        Catch ex As Exception
            UtilsLib.LogError("DiscoverAssemblyTree: ERROR - " & ex.Message)
        End Try
        
        Return tree
    End Function
    
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
                    If idwPath.Contains("\OldVersions\") Then Continue For
                    
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
            
            If doc Is Nothing Then Continue For
            
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
            
            For Each paramKvp In kvp.Value
                Try
                    params.Item(paramKvp.Key).Expression = paramKvp.Value
                Catch
                End Try
            Next
        Next
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
        Catch
            Return
        End Try
        
        If params Is Nothing Then Return
        
        For Each kvp In parameters
            If kvp.Key.StartsWith("_") Then Continue For
            Try
                params.Item(kvp.Key).Expression = kvp.Value
            Catch
            End Try
        Next
    End Sub
    
    ''' <summary>
    ''' Build variant matrix with fingerprints for all parts across all variants.
    ''' </summary>
    Public Function BuildVariantMatrix(app As Inventor.Application, _
                                        tree As AssemblyTree, _
                                        variants As List(Of ExcelReaderLib.ReleaseConfig), _
                                        masterPaths As List(Of String)) As VariantMatrix
        Dim matrix As New VariantMatrix()
        matrix.PartPaths = New List(Of String)(tree.Parts.Keys)
        matrix.VariantNames = New List(Of String)
        For Each vc As ExcelReaderLib.ReleaseConfig In variants
            matrix.VariantNames.Add(vc.ConfigName)
        Next
        
        Dim snapshot = SnapshotMasterParameters(app, masterPaths)
        
        Try
            For Each variantCfg As ExcelReaderLib.ReleaseConfig In variants
                UtilsLib.LogInfo("BuildVariantMatrix: Analyzing moodul " & variantCfg.ConfigName)
                
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
                    
                    If doc IsNot Nothing Then
                        ApplyParameters(doc, variantCfg.Parameters)
                    End If
                Next
                
                Try
                    app.ActiveDocument.Update()
                Catch
                End Try
                
                Dim variantFps As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For Each partPath In matrix.PartPaths
                    Dim partDoc As PartDocument = Nothing
                    Try
                        For Each d As Document In app.Documents
                            If d.FullFileName.Equals(partPath, StringComparison.OrdinalIgnoreCase) Then
                                partDoc = CType(d, PartDocument)
                                Exit For
                            End If
                        Next
                    Catch
                    End Try
                    
                    If partDoc IsNot Nothing Then
                        variantFps.Add(partPath, ComputeGeometryFingerprint(partDoc))
                    Else
                        variantFps.Add(partPath, "NOT_LOADED")
                    End If
                Next
                
                matrix.Fingerprints.Add(variantCfg.ConfigName, variantFps)
            Next
        Finally
            RestoreMasterParameters(app, snapshot)
            Try
                app.ActiveDocument.Update()
            Catch
            End Try
        End Try
        
        UtilsLib.LogInfo("BuildVariantMatrix: Complete - " & matrix.PartPaths.Count & " parts x " & matrix.VariantNames.Count & " moodulid")
        Return matrix
    End Function
    
    ''' <summary>
    ''' Classify parts into groups based on fingerprint sharing.
    ''' </summary>
    Public Function ClassifyPartGroups(matrix As VariantMatrix, _
                                        tree As AssemblyTree) As List(Of PartGroup)
        Dim groups As New List(Of PartGroup)
        
        For Each partPath In matrix.PartPaths
            Dim group As New PartGroup()
            group.PartPath = partPath
            group.RelativePath = tree.Parts(partPath).RelativePath
            group.PartNumber = tree.Parts(partPath).PartNumber
            
            For Each variantName In matrix.VariantNames
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
    ''' </summary>
    Public Function GetMasterPaths(tree As AssemblyTree) As List(Of String)
        Dim masters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        
        For Each kvp In tree.Parts
            If kvp.Value.Role = PartRole.Derived AndAlso Not String.IsNullOrEmpty(kvp.Value.DerivedFromMaster) Then
                masters.Add(kvp.Value.DerivedFromMaster)
            End If
        Next
        
        Return masters.ToList()
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
    ''' </summary>
    Public Function GenerateLocalNumbers(count As Integer, outputRoot As String) As List(Of String)
        UtilsLib.LogInfo("GenerateLocalNumbers: Generating " & count & " local numbers (development mode)")
        
        Dim startNum As Integer = 1
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
    ''' Compute the complete release plan.
    ''' </summary>
    Public Function ComputeReleasePlan(tree As AssemblyTree, _
                                        partGroups As List(Of PartGroup), _
                                        variants As List(Of ExcelReaderLib.ReleaseConfig), _
                                        targetRoot As String, _
                                        fileNumbers As List(Of String)) As ReleasePlan
        Dim plan As New ReleasePlan()
        plan.SharedFolder = System.IO.Path.Combine(targetRoot, "Ühine")
        
        Dim numberIndex As Integer = 0
        
        For Each variantCfg As ExcelReaderLib.ReleaseConfig In variants
            plan.VariantFolders.Add(variantCfg.ConfigName, System.IO.Path.Combine(targetRoot, variantCfg.ConfigName))
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
                For Each vc As ExcelReaderLib.ReleaseConfig In variants
                    allVariantNames.Add(vc.ConfigName)
                Next
                plan.Files.Add(New PlannedFile With {
                    .SourcePath = group.PartPath,
                    .TargetLocalPath = System.IO.Path.Combine(plan.SharedFolder, group.RelativePath),
                    .VaultNumber = fileNumbers(numberIndex),
                    .FileType = FileType.Part,
                    .IsShared = True,
                    .IsExisting = False,
                    .ForVariants = allVariantNames,
                    .Fingerprint = fp
                })
                numberIndex += 1
            Else
                For Each fpKvp In group.UniqueFingerprints
                    Dim firstVariant = fpKvp.Value(0)
                    Dim newFileName As String = fileNumbers(numberIndex) & System.IO.Path.GetExtension(group.PartPath)
                    Dim relDir As String = System.IO.Path.GetDirectoryName(group.RelativePath)
                    
                    plan.Files.Add(New PlannedFile With {
                        .SourcePath = group.PartPath,
                        .TargetLocalPath = System.IO.Path.Combine(plan.VariantFolders(firstVariant), relDir, newFileName),
                        .VaultNumber = fileNumbers(numberIndex),
                        .FileType = FileType.Part,
                        .IsShared = False,
                        .IsExisting = False,
                        .ForVariants = fpKvp.Value,
                        .Fingerprint = fpKvp.Key
                    })
                    numberIndex += 1
                Next
            End If
        Next
        
        For Each variantCfg As ExcelReaderLib.ReleaseConfig In variants
            For Each asmKvp In tree.Assemblies
                Dim relativePath = asmKvp.Value.RelativePath
                Dim newFileName As String = fileNumbers(numberIndex) & ".iam"
                Dim relDir As String = System.IO.Path.GetDirectoryName(relativePath)
                
                plan.Files.Add(New PlannedFile With {
                    .SourcePath = asmKvp.Key,
                    .TargetLocalPath = System.IO.Path.Combine(plan.VariantFolders(variantCfg.ConfigName), relDir, newFileName),
                    .VaultNumber = fileNumbers(numberIndex),
                    .FileType = FileType.Assembly,
                    .IsShared = False,
                    .ForVariants = New List(Of String) From {variantCfg.ConfigName}
                })
                numberIndex += 1
            Next
        Next
        
        For Each dwgInfo As DrawingInfo In tree.Drawings
            ' Get drawing filename and check if it starts with the model number
            Dim dwgFileName As String = System.IO.Path.GetFileNameWithoutExtension(dwgInfo.DrawingPath)
            Dim primaryModelPath As String = If(dwgInfo.ReferencedModelPaths.Count > 0, dwgInfo.ReferencedModelPaths(0), "")
            Dim modelNumber As String = System.IO.Path.GetFileNameWithoutExtension(primaryModelPath)
            
            ' Extract suffix from drawing filename (anything after the model number)
            ' e.g., "00005_sheet2" with model "00005" -> suffix "_sheet2"
            Dim dwgSuffix As String = ""
            Dim shareNumberWithModel As Boolean = False
            If Not String.IsNullOrEmpty(modelNumber) AndAlso dwgFileName.StartsWith(modelNumber, StringComparison.OrdinalIgnoreCase) Then
                shareNumberWithModel = True
                dwgSuffix = dwgFileName.Substring(modelNumber.Length)
            End If
            
            ' A drawing can only be shared if:
            ' 1. We have 2+ moodulid to share between
            ' 2. All its referenced parts have the same geometry across all moodulid
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
                ' Find the released model's number if drawing shares number with model
                Dim vaultNum As String
                If shareNumberWithModel AndAlso Not String.IsNullOrEmpty(primaryModelPath) Then
                    Dim modelFile As PlannedFile = FindPlannedFileBySource(plan.Files, primaryModelPath)
                    If modelFile IsNot Nothing Then
                        vaultNum = modelFile.VaultNumber
                        UtilsLib.LogInfo("Drawing " & dwgFileName & ".idw reuses model number " & vaultNum & " with suffix '" & dwgSuffix & "'")
                    Else
                        vaultNum = fileNumbers(numberIndex)
                        numberIndex += 1
                    End If
                Else
                    vaultNum = fileNumbers(numberIndex)
                    numberIndex += 1
                End If
                
                ' Preserve any suffix from the original drawing filename
                Dim newFileName As String = vaultNum & dwgSuffix & ".idw"
                Dim relDir As String = System.IO.Path.GetDirectoryName(dwgInfo.RelativePath)
                Dim allVariantNames2 As New List(Of String)
                For Each vc2 As ExcelReaderLib.ReleaseConfig In variants
                    allVariantNames2.Add(vc2.ConfigName)
                Next
                
                plan.Files.Add(New PlannedFile With {
                    .SourcePath = dwgInfo.DrawingPath,
                    .TargetLocalPath = System.IO.Path.Combine(plan.SharedFolder, relDir, newFileName),
                    .VaultNumber = vaultNum,
                    .FileType = FileType.Drawing,
                    .IsShared = True,
                    .ForVariants = allVariantNames2
                })
            Else
                For Each variantCfg2 As ExcelReaderLib.ReleaseConfig In variants
                    ' Find the released model's number for this variant if drawing shares number
                    Dim vaultNum As String
                    If shareNumberWithModel AndAlso Not String.IsNullOrEmpty(primaryModelPath) Then
                        Dim modelFile As PlannedFile = FindPlannedFileBySourceAndVariant(plan.Files, primaryModelPath, variantCfg2.ConfigName)
                        If modelFile IsNot Nothing Then
                            vaultNum = modelFile.VaultNumber
                            UtilsLib.LogInfo("Drawing " & dwgFileName & ".idw (" & variantCfg2.ConfigName & ") reuses model number " & vaultNum & " with suffix '" & dwgSuffix & "'")
                        Else
                            vaultNum = fileNumbers(numberIndex)
                            numberIndex += 1
                        End If
                    Else
                        vaultNum = fileNumbers(numberIndex)
                        numberIndex += 1
                    End If
                    
                    ' Preserve any suffix from the original drawing filename
                    Dim newFileName As String = vaultNum & dwgSuffix & ".idw"
                    Dim relDir As String = System.IO.Path.GetDirectoryName(dwgInfo.RelativePath)
                    
                    plan.Files.Add(New PlannedFile With {
                        .SourcePath = dwgInfo.DrawingPath,
                        .TargetLocalPath = System.IO.Path.Combine(plan.VariantFolders(variantCfg2.ConfigName), relDir, newFileName),
                        .VaultNumber = vaultNum,
                        .FileType = FileType.Drawing,
                        .IsShared = False,
                        .ForVariants = New List(Of String) From {variantCfg2.ConfigName}
                    })
                Next
            End If
        Next
        
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
        UtilsLib.LogInfo("  - Moodul-specific: " & variantSpecificCount)
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
            "  - Mooduli detailid: " & variantPartsCount & vbCrLf &
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
                                          newPartNumber As String) As Boolean
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
                
                ' Document is now the target file - break derivation links
                Dim dpcs = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
                Dim dpcList As New List(Of DerivedPartComponent)
                For Each dpc As DerivedPartComponent In dpcs
                    dpcList.Add(dpc)
                Next
                
                For Each dpc As DerivedPartComponent In dpcList
                    Try
                        dpc.BreakLinkToFile()
                        UtilsLib.LogInfo("  Broke link: " & dpc.Name)
                    Catch ex As Exception
                        UtilsLib.LogWarn("  WARNING: Could not break link - " & ex.Message)
                    End Try
                Next
                
                ' Set Part Number in Design Tracking Properties
                Try
                    Dim designProps = partDoc.PropertySets.Item("Design Tracking Properties")
                    designProps.Item("Part Number").Value = newPartNumber
                    UtilsLib.LogInfo("  Set Part Number: " & newPartNumber)
                Catch ex As Exception
                    UtilsLib.LogWarn("  WARNING: Failed to set Part Number: " & ex.Message)
                End Try
                
                ' Also set Title in Summary Information (title blocks often use this)
                Try
                    Dim summaryProps = partDoc.PropertySets.Item("Inventor Summary Information")
                    summaryProps.Item("Title").Value = newPartNumber
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
    ''' Create assembly snapshot with updated references using SaveAs for new GUID.
    ''' </summary>
    Public Function CreateAssemblySnapshot(app As Inventor.Application, _
                                            sourceAsmPath As String, _
                                            targetPath As String, _
                                            referenceMap As Dictionary(Of String, String), _
                                            variantParams As Dictionary(Of String, String), _
                                            newPartNumber As String) As Boolean
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
                For Each occ As ComponentOccurrence In occsToProcess
                    Try
                        Dim currentPath As String = occ.Definition.Document.FullFileName
                        If referenceMap.ContainsKey(currentPath) Then
                            occ.Replace(referenceMap(currentPath), True)
                            UtilsLib.LogInfo("    Replaced: " & System.IO.Path.GetFileName(currentPath) & " -> " & System.IO.Path.GetFileName(referenceMap(currentPath)))
                        Else
                            UtilsLib.LogWarn("    NOT IN MAP: " & currentPath)
                        End If
                    Catch ex As Exception
                        UtilsLib.LogWarn("    Replace failed: " & ex.Message)
                    End Try
                Next
                
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
                    
                    ' Verify immediately
                    Dim verifyPN As String = designProps.Item("Part Number").Value.ToString()
                    UtilsLib.LogInfo("  DEBUG: Verify Part Number=" & verifyPN)
                Catch ex As Exception
                    UtilsLib.LogWarn("  WARNING: Failed to set Part Number: " & ex.Message)
                End Try
                
                ' Also set Title in Summary Information (title blocks often use this)
                Try
                    Dim summaryProps = asmDoc.PropertySets.Item("Inventor Summary Information")
                    Dim oldTitle As String = summaryProps.Item("Title").Value.ToString()
                    UtilsLib.LogInfo("  DEBUG: Old Title=" & oldTitle)
                    summaryProps.Item("Title").Value = newPartNumber
                    UtilsLib.LogInfo("  Set Title: " & newPartNumber)
                    
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
                                       newPartNumber As String) As Boolean
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
            
            ' Log current drawing references before replacement
            UtilsLib.LogInfo("  Drawing refs before replacement:")
            For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
                Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
                Dim inMap As String = If(referenceMap.ContainsKey(fd.FullFileName), "-> " & System.IO.Path.GetFileName(referenceMap(fd.FullFileName)), "(not in map)")
                UtilsLib.LogInfo("    " & System.IO.Path.GetFileName(fd.FullFileName) & " " & inMap)
            Next
            
            ' Replace references to released files
            For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
                Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
                Dim oldPath As String = fd.FullFileName
                If referenceMap.ContainsKey(oldPath) Then
                    Try
                        fd.ReplaceReference(referenceMap(oldPath))
                        UtilsLib.LogInfo("  Replaced: " & System.IO.Path.GetFileName(oldPath))
                    Catch ex As Exception
                        UtilsLib.LogWarn("  WARNING: ReplaceReference failed - " & ex.Message)
                    End Try
                End If
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
            
            ' Set drawing's own Part Number (for title blocks that use "Drawing Properties")
            Try
                Dim designProps = drawDoc.PropertySets.Item("Design Tracking Properties")
                designProps.Item("Part Number").Value = newPartNumber
                UtilsLib.LogInfo("  Set Drawing Part Number: " & newPartNumber)
            Catch ex As Exception
                UtilsLib.LogWarn("  WARNING: Failed to set Drawing Part Number: " & ex.Message)
            End Try
            
            ' Also set Title in Summary Information
            Try
                Dim summaryProps = drawDoc.PropertySets.Item("Inventor Summary Information")
                summaryProps.Item("Title").Value = newPartNumber
                UtilsLib.LogInfo("  Set Drawing Title: " & newPartNumber)
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
    Public Function ExecuteRelease(app As Inventor.Application, context As ReleaseContext) As Boolean
        UtilsLib.LogInfo("ExecuteRelease: Starting...")
        
        Dim masterSnapshot = SnapshotMasterParameters(app, context.MasterPaths)
        
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
            
            UtilsLib.LogInfo("ExecuteRelease: Creating parts...")
            Dim processedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            
            For Each file As PlannedFile In context.ReleasePlan.Files
                If file.FileType <> FileType.Part Then Continue For
                If file.IsExisting Then Continue For
                
                ' Skip if already processed (e.g., shared parts used by multiple moodulid)
                If processedParts.Contains(file.TargetLocalPath) Then Continue For
                processedParts.Add(file.TargetLocalPath)
                
                ' Note: Parameters are applied during assembly creation, not part creation
                ' Parts get their geometry from derivation, not from assembly parameters
                CreateStandalonePart(app, file.SourcePath, file.TargetLocalPath, file.VaultNumber)
            Next
            
            UtilsLib.LogInfo("ExecuteRelease: Creating assemblies...")
            For Each variantCfg As ExcelReaderLib.ReleaseConfig In context.Variants
                Dim refMap = BuildReferenceMapForVariant(context, variantCfg.ConfigName)
                
                For Each file As PlannedFile In context.ReleasePlan.Files
                    If file.FileType = FileType.Assembly AndAlso file.ForVariants.Contains(variantCfg.ConfigName) Then
                        CreateAssemblySnapshot(app, file.SourcePath, file.TargetLocalPath, refMap, variantCfg.Parameters, file.VaultNumber)
                    End If
                Next
            Next
            
            UtilsLib.LogInfo("ExecuteRelease: Creating drawings...")
            For Each file As PlannedFile In context.ReleasePlan.Files
                If file.FileType = FileType.Drawing Then
                    Dim variantName = file.ForVariants(0)
                    Dim refMap = BuildReferenceMapForVariant(context, variantName)
                    CreateDrawingCopy(app, file.SourcePath, file.TargetLocalPath, refMap, file.VaultNumber)
                End If
            Next
            
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
            Dim verificationPassed As Boolean = VerifyReleasedFiles(app, context)
            
            ' Reopen the original source assembly so user is back where they started
            UtilsLib.LogInfo("ExecuteRelease: Reopening source assembly...")
            Dim originalAssembly As String = context.AssemblyTree.RootAssemblyPath
            Try
                app.Documents.Open(originalAssembly, True)
                UtilsLib.LogInfo("  Reopened: " & System.IO.Path.GetFileName(originalAssembly))
            Catch ex As Exception
                UtilsLib.LogWarn("  Failed to reopen " & originalAssembly & ": " & ex.Message)
            End Try
            
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
    Private Function VerifyReleasedFiles(app As Inventor.Application, context As ReleaseContext) As Boolean
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
    Private Function FindVariantByName(variants As List(Of ExcelReaderLib.ReleaseConfig), name As String) As ExcelReaderLib.ReleaseConfig
        For Each v As ExcelReaderLib.ReleaseConfig In variants
            If v.ConfigName = name Then Return v
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Build reference map for a specific variant.
    ''' </summary>
    Private Function BuildReferenceMapForVariant(context As ReleaseContext, variantName As String) As Dictionary(Of String, String)
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
        Public Modules As New List(Of ModuleEntry)
        Public SharedParts As New List(Of SharedPartEntry)
    End Class
    
    Public Class ModuleEntry
        Public ModuleName As String
        Public Variants As New List(Of VariantEntry)
        Public ReleaseDate As DateTime
    End Class
    
    Public Class VariantEntry
        Public ConfigName As String
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
        Public UsedByModules As New List(Of String)
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
        
        Dim summary As String = "Moodulite väljastamine lõpetatud!" & vbCrLf & vbCrLf &
            "Faile loodud: " & plan.Files.Count & vbCrLf &
            "  - Jagatud: " & sharedCnt & vbCrLf &
            "  - Moodulispetsiifilised: " & variantCnt & vbCrLf & vbCrLf &
            "Numbrid: " & firstNum & " kuni " & lastNum
        
        MessageBox.Show(summary, "Väljastamine lõpetatud", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Module
