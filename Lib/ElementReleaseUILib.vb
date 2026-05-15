' ElementReleaseUILib.vb - UI for Element Release System
' Provides interactive UI for selecting elements, previewing files, and tracking progress
' NOTE: Requires AddReference "System.Drawing" in main rule for color support

Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Public Module ElementReleaseUILib

    ' ============================================================================
    ' Module-level font for regular (non-bold) TreeView nodes
    ' Set when TreeView is created with bold base font
    ' ============================================================================
    Private m_RegularFont As Font = Nothing
    
    ' ============================================================================
    ' UI State
    ' ============================================================================
    
    ''' <summary>
    ''' Result from the release UI dialog
    ''' </summary>
    Public Class ReleaseUIResult
        Public Property Cancelled As Boolean = False
        Public Property SelectedElements As New List(Of ExcelReaderLib.ElementConfig)
        Public Property ExecutionCompleted As Boolean = False
        Public Property SuccessCount As Integer = 0
        Public Property FailureCount As Integer = 0
        Public Property FailedFiles As New List(Of String)
    End Class
    
    ''' <summary>
    ''' State for tracking file progress in the tree
    ''' </summary>
    Public Enum FileStatus
        Pending
        InProgress
        Completed
        Failed
        Skipped
    End Enum
    
    ''' <summary>
    ''' Node tag for tree items
    ''' </summary>
    Public Class TreeNodeData
        Public Property FilePath As String
        Public Property FileType As String
        Public Property ElementName As String
        Public Property Status As FileStatus = FileStatus.Pending
        Public Property VaultNumber As String
        Public Property Description As String
        Public Property IsShared As Boolean = False
        Public Property BaseText As String  ' Original formatted text without status prefix
    End Class
    
    ''' <summary>
    ''' Selection data for element nodes (used instead of checkboxes)
    ''' </summary>
    Public Class ElementSelectionData
        Public Property Element As ExcelReaderLib.ElementConfig
        Public Property IsSelected As Boolean = True
        Public Property RelativePath As String  ' For display in node text
        Public Property FileCount As Integer     ' Total file count (own + shared)
        Public Property OwnFileCount As Integer  ' Element-specific file count
        Public Property SharedFileCount As Integer ' Shared file count
    End Class
    
    ' ============================================================================
    ' Main UI Entry Point
    ' ============================================================================
    
    ''' <summary>
    ''' Shows the element release UI.
    ''' Returns selected elements and whether to proceed.
    ''' </summary>
    Public Function ShowReleaseUI(app As Inventor.Application, _
                                  elements As List(Of ExcelReaderLib.ElementConfig), _
                                  plan As ElementReleaseLib.ReleasePlan, _
                                  context As ElementReleaseLib.ElementReleaseContext) As ReleaseUIResult
        
        Dim result As New ReleaseUIResult()
        
        ' Create the main form
        Dim frm As Form = CreateReleaseForm(elements, plan, context, result)
        
        ' Show modal - the form handles everything
        frm.ShowDialog()
        
        Return result
    End Function
    
    ''' <summary>
    ''' Creates the main release form with unified tree view
    ''' </summary>
    Private Function CreateReleaseForm(elements As List(Of ExcelReaderLib.ElementConfig), _
                                       plan As ElementReleaseLib.ReleasePlan, _
                                       context As ElementReleaseLib.ElementReleaseContext, _
                                       result As ReleaseUIResult) As Form
        
        Dim frm As New Form()
        frm.Text = StringsLib.TITLE_ELEMENT_RELEASE & " - " & context.ElementName
        frm.Width = 1100
        frm.Height = 750
        frm.StartPosition = FormStartPosition.CenterScreen
        frm.FormBorderStyle = FormBorderStyle.Sizable
        frm.MaximizeBox = True
        frm.MinimizeBox = True
        UILib.SetMinimumSize(frm, 900, 600)
        
        ' Main panel for tree and status
        Dim mainPanel As New Panel()
        mainPanel.Dock = DockStyle.Fill
        mainPanel.Padding = New Padding(10)
        frm.Controls.Add(mainPanel)
        
        ' ========================================
        ' TOP: Header with stats and buttons
        ' ========================================
        Dim headerPanel As New Panel()
        headerPanel.Dock = DockStyle.Top
        headerPanel.Height = 60
        mainPanel.Controls.Add(headerPanel)
        
        ' Title and stats
        Dim lblTitle As New Label()
        lblTitle.Text = "VÄLJASTAMISE PLAAN — " & elements.Count & " elementi, " & plan.Files.Count & " faili " & GetPlanTitleCategorySuffix(plan)
        lblTitle.Dock = DockStyle.Top
        lblTitle.Height = 25
        headerPanel.Controls.Add(lblTitle)
        
        ' Button panel for select all/none
        Dim btnPanel As New FlowLayoutPanel()
        btnPanel.Dock = DockStyle.Top
        btnPanel.Height = 30
        btnPanel.FlowDirection = FlowDirection.LeftToRight
        headerPanel.Controls.Add(btnPanel)
        
        Dim btnSelectAll As New Button()
        btnSelectAll.Text = "Vali kõik elemendid"
        btnSelectAll.AutoSize = True
        btnPanel.Controls.Add(btnSelectAll)
        
        Dim btnSelectNone As New Button()
        btnSelectNone.Text = "Tühista valik"
        btnSelectNone.AutoSize = True
        btnPanel.Controls.Add(btnSelectNone)
        
        Dim btnExpandAll As New Button()
        btnExpandAll.Text = "Laienda kõik"
        btnExpandAll.AutoSize = True
        btnPanel.Controls.Add(btnExpandAll)
        
        Dim btnCollapseAll As New Button()
        btnCollapseAll.Text = "Ahenda kõik"
        btnCollapseAll.AutoSize = True
        btnPanel.Controls.Add(btnCollapseAll)
        
        ' Stats label
        Dim lblStats As New Label()
        lblStats.Name = "lblStats"
        lblStats.Dock = DockStyle.Top
        lblStats.Height = 20
        lblStats.Text = GetPlanStats(plan) & "  (topeltklikk elemendil valiku muutmiseks)"
        headerPanel.Controls.Add(lblStats)
        
        ' Set header control order
        headerPanel.Controls.SetChildIndex(lblStats, 0)
        headerPanel.Controls.SetChildIndex(btnPanel, 1)
        headerPanel.Controls.SetChildIndex(lblTitle, 2)
        
        ' ========================================
        ' MAIN: Unified TreeView (NO checkboxes - use text markers for selection)
        ' ========================================
        Dim treeView As New TreeView()
        treeView.Dock = DockStyle.Fill
        treeView.CheckBoxes = False  ' No checkboxes - use text markers instead
        treeView.ShowLines = True
        treeView.ShowPlusMinus = True
        treeView.ShowRootLines = True
        treeView.FullRowSelect = True
        treeView.HideSelection = False
        treeView.Indent = 20
        treeView.ShowNodeToolTips = True  ' Enable tooltips for file details
        treeView.Scrollable = True  ' Enable scrollbars (horizontal when text is wider than view)
        ' Set TreeView base font to bold - this fixes width calculation for bold nodes
        ' Store regular font for non-bold nodes (file nodes, folder nodes)
        m_RegularFont = New Font(treeView.Font, FontStyle.Regular)
        treeView.Font = New Font(treeView.Font, FontStyle.Bold)
        mainPanel.Controls.Add(treeView)
        
        ' Populate unified tree (reads iProperties from source files)
        PopulateUnifiedTree(treeView, elements, plan, context)
        
        ' Handle double-click to toggle element selection
        AddHandler treeView.DoubleClick, Sub(s, e)
            Dim node As TreeNode = treeView.SelectedNode
            If node IsNot Nothing AndAlso node.Tag IsNot Nothing Then
                If TypeOf node.Tag Is ElementSelectionData Then
                    Dim selData As ElementSelectionData = CType(node.Tag, ElementSelectionData)
                    selData.IsSelected = Not selData.IsSelected
                    UpdateElementNodeText(node, selData)
                    UpdateStats(frm)
                End If
            End If
        End Sub
        
        ' Set main panel control order
        mainPanel.Controls.SetChildIndex(treeView, 0)
        mainPanel.Controls.SetChildIndex(headerPanel, 1)
        
        ' ========================================
        ' BOTTOM: Status and Buttons
        ' ========================================
        Dim bottomPanel As New Panel()
        bottomPanel.Dock = DockStyle.Bottom
        bottomPanel.Height = 80
        bottomPanel.Padding = New Padding(10, 5, 10, 5)
        frm.Controls.Add(bottomPanel)
        
        ' Progress bar
        Dim progressBar As New System.Windows.Forms.ProgressBar()
        progressBar.Name = "progressBar"
        progressBar.Dock = DockStyle.Top
        progressBar.Height = 20
        progressBar.Minimum = 0
        progressBar.Maximum = 100
        progressBar.Value = 0
        bottomPanel.Controls.Add(progressBar)
        
        ' Progress label
        Dim lblProgress As New Label()
        lblProgress.Name = "lblProgress"
        lblProgress.Dock = DockStyle.Top
        lblProgress.Height = 20
        lblProgress.Text = "Valmis alustamiseks"
        bottomPanel.Controls.Add(lblProgress)
        
        ' Button panel
        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.Dock = DockStyle.Bottom
        buttonPanel.Height = 35
        buttonPanel.FlowDirection = FlowDirection.RightToLeft
        bottomPanel.Controls.Add(buttonPanel)
        
        Dim btnCancel As New Button()
        btnCancel.Name = "btnCancel"
        btnCancel.Text = "Tühista"
        btnCancel.Width = 100
        btnCancel.Height = 30
        buttonPanel.Controls.Add(btnCancel)
        
        Dim btnExecute As New Button()
        btnExecute.Name = "btnExecute"
        btnExecute.Text = ">>> Väljasta <<<"
        btnExecute.Width = 140
        btnExecute.Height = 30
        buttonPanel.Controls.Add(btnExecute)
        
        ' Set bottom panel control order
        bottomPanel.Controls.SetChildIndex(buttonPanel, 0)
        bottomPanel.Controls.SetChildIndex(lblProgress, 1)
        bottomPanel.Controls.SetChildIndex(progressBar, 2)
        
        ' Wire up select all/none buttons
        AddHandler btnSelectAll.Click, Sub(s, e)
            SetAllElementsSelected(treeView, True)
            UpdateStats(frm)
        End Sub
        
        AddHandler btnSelectNone.Click, Sub(s, e)
            SetAllElementsSelected(treeView, False)
            UpdateStats(frm)
        End Sub
        
        AddHandler btnExpandAll.Click, Sub(s, e)
            treeView.ExpandAll()
        End Sub
        
        AddHandler btnCollapseAll.Click, Sub(s, e)
            treeView.CollapseAll()
        End Sub
        
        ' Store references for event handlers
        frm.Tag = New Dictionary(Of String, Object) From {
            {"result", result},
            {"elements", elements},
            {"plan", plan},
            {"context", context},
            {"treeView", treeView},
            {"progressBar", progressBar},
            {"lblProgress", lblProgress},
            {"lblStats", lblStats},
            {"btnExecute", btnExecute},
            {"btnCancel", btnCancel}
        }
        
        ' Initial stats update (sets correct shared node appearance)
        UpdateStats(frm)

        ' Wire up button events
        AddHandler btnCancel.Click, Sub(s, e)
            result.Cancelled = True
            frm.Close()
        End Sub
        
        ' Handle form close (X button) same as Cancel
        AddHandler frm.FormClosing, Sub(s, e)
            ' Only treat as cancel if user closed via X button (not via Execute or Cancel buttons)
            ' Execute sets DialogResult = OK, Cancel button sets Cancelled = True directly
            If frm.DialogResult <> DialogResult.OK AndAlso Not result.Cancelled Then
                result.Cancelled = True
            End If
        End Sub
        
        AddHandler btnExecute.Click, AddressOf OnExecuteClick
        
        Return frm
    End Function
    
    ''' <summary>
    ''' Gets statistics string for the plan
    ''' </summary>
    Private Function GetPlanStats(plan As ElementReleaseLib.ReleasePlan) As String
        Dim sharedParts As Integer = 0
        Dim uniqueParts As Integer = 0
        Dim assemblies As Integer = 0
        Dim drawings As Integer = 0
        Dim reuseCount As Integer = 0
        Dim newNumberCount As Integer = 0
        
        For Each f As ElementReleaseLib.PlannedFile In plan.Files
            If f.IsReuse Then reuseCount += 1
            If f.IsPlaceholder Then newNumberCount += 1
            Select Case f.FileType
                Case ElementReleaseLib.FileType.Part
                    If f.IsShared Then sharedParts += 1 Else uniqueParts += 1
                Case ElementReleaseLib.FileType.Assembly
                    assemblies += 1
                Case ElementReleaseLib.FileType.Drawing
                    drawings += 1
            End Select
        Next
        
        Return String.Format("Taaskasutus: {0} | Uued numbrid: {1} | Jagatud: {2} | Unikaalsed: {3} | Koostud: {4} | Joonised: {5}", _
            reuseCount, newNumberCount, sharedParts, uniqueParts, assemblies, drawings)
    End Function
    
    ''' <summary>
    ''' Category counts for the "VÄLJASTAMISE PLAAN" header line only (same priority as row markers: jagatud, then taaskasutus, then uus).
    ''' Uses digits only so WinForms default font always renders.
    ''' </summary>
    Private Function GetPlanTitleCategorySuffix(plan As ElementReleaseLib.ReleasePlan) As String
        Dim nShared As Integer = 0
        Dim nReuse As Integer = 0
        Dim nNew As Integer = 0
        Dim nOther As Integer = 0
        For Each f As ElementReleaseLib.PlannedFile In plan.Files
            If f.IsShared Then
                nShared += 1
            ElseIf f.IsReuse Then
                nReuse += 1
            ElseIf f.IsPlaceholder Then
                nNew += 1
            Else
                nOther += 1
            End If
        Next
        Dim core As String = nReuse.ToString() & " taaskasutus | " & nNew.ToString() & " uus | " & nShared.ToString() & " jagatud"
        If nOther > 0 Then
            Return "(" & core & " | " & nOther.ToString() & " muu)"
        End If
        Return "(" & core & ")"
    End Function
    
    ' Status prefixes for tree nodes (no ImageList due to iLogic constraints)
    Private Const STATUS_PENDING As String = "[ ] "
    Private Const STATUS_INPROGRESS As String = "[...] "
    Private Const STATUS_COMPLETED As String = "[OK] "
    Private Const STATUS_FAILED As String = "[X] "
    Private Const STATUS_SKIPPED As String = "[-] "
    
    ''' <summary>
    ''' Populates a unified tree with Ühine (shared) section at top, then elements with their own files
    ''' </summary>
    ''' <summary>
    ''' Populates the unified tree view with elements and files
    ''' </summary>
    ''' <param name="isExecutionMode">If True, show status prefixes and hide selection markers</param>
    Private Sub PopulateUnifiedTree(tree As TreeView, _
                                    elements As List(Of ExcelReaderLib.ElementConfig), _
                                    plan As ElementReleaseLib.ReleasePlan, _
                                    context As ElementReleaseLib.ElementReleaseContext, _
                                    Optional isExecutionMode As Boolean = False)
        tree.BeginUpdate()
        tree.Nodes.Clear()
        
        ' Group files by shared/element
        Dim sharedFiles As New List(Of ElementReleaseLib.PlannedFile)
        Dim elementFiles As New Dictionary(Of String, List(Of ElementReleaseLib.PlannedFile))
        
        For Each elem As ExcelReaderLib.ElementConfig In elements
            elementFiles(elem.ElementName) = New List(Of ElementReleaseLib.PlannedFile)
        Next
        
        For Each f As ElementReleaseLib.PlannedFile In plan.Files
            If f.IsShared Then
                sharedFiles.Add(f)
            Else
                For Each varName In f.ForVariants
                    If elementFiles.ContainsKey(varName) Then
                        elementFiles(varName).Add(f)
                    End If
                Next
            End If
        Next
        
        ' Get product family name for relative paths (last folder in TargetRoot)
        Dim productFamily As String = System.IO.Path.GetFileName(context.TargetRoot.TrimEnd("\"c))
        
        ' ========================================
        ' SECTION 1: Ühine (Shared files) at TOP
        ' ========================================
        If sharedFiles.Count > 0 Then
            Dim sharedFolder As String = If(plan.SharedFolder, context.TargetRoot & "\Ühine")
            Dim sharedRelativePath As String = productFamily & "\Ühine"
            Dim sharedNodeText As String = "ÜHINE — " & sharedFiles.Count & " faili → " & sharedRelativePath
            Dim sharedNode As New TreeNode(sharedNodeText)
            sharedNode.Tag = "SHARED_ROOT"  ' Not an element, not checkable
            
            ' Shared root inherits bold from TreeView, just set color
            sharedNode.ForeColor = Drawing.Color.DarkOrange
            sharedNode.ToolTipText = "Jagatud failid luuakse ühiskausta ja on kasutatavad kõigi valitud elementide poolt." & vbCrLf & _
                                     "Need failid luuakse ainult siis, kui vähemalt üks element on valitud." & vbCrLf & _
                                     "Täielik tee: " & sharedFolder
            
            ' Add shared files in folder structure - use type colors (not gray) for Ühine
            AddFilesToNodeByFolderMerged(sharedNode, sharedFiles, context, forceTypeColors:=True, showStatusPrefix:=isExecutionMode)
            
            tree.Nodes.Add(sharedNode)
        End If
        
        ' ========================================
        ' SECTION 2: Each element with its OWN files only
        ' ========================================
        For Each elem As ExcelReaderLib.ElementConfig In elements
            ' Build parameters string
            Dim paramStr As String = ""
            For Each kvp In elem.Parameters
                If paramStr <> "" Then paramStr &= ", "
                paramStr &= kvp.Key & "=" & kvp.Value
            Next
            
            Dim elemFileCount As Integer = elementFiles(elem.ElementName).Count
            Dim elemFolder As String = ""
            If plan.VariantFolders.ContainsKey(elem.ElementName) Then
                elemFolder = plan.VariantFolders(elem.ElementName)
            Else
                elemFolder = context.TargetRoot & "\" & elem.ElementName
            End If
            Dim elemRelativePath As String = productFamily & "\" & elem.ElementName
            
            ' Create element node with selection marker
            Dim selData As New ElementSelectionData()
            selData.Element = elem
            selData.IsSelected = True
            selData.RelativePath = elemRelativePath
            selData.FileCount = elemFileCount + sharedFiles.Count  ' Total including shared
            selData.OwnFileCount = elemFileCount
            selData.SharedFileCount = sharedFiles.Count
            
            ' Show element file count + output folder
            Dim totalFiles As Integer = elemFileCount + sharedFiles.Count
            Dim elemText As String = ""
            If Not isExecutionMode Then
                elemText = "[✓] "  ' Selection marker only in preview mode
            End If
            elemText &= elem.ElementName
            If paramStr <> "" Then elemText &= " (" & paramStr & ")"
            elemText &= " — " & totalFiles & " faili (" & elemFileCount & " oma + " & sharedFiles.Count & " jagatud) → " & elemRelativePath
            
            Dim elemNode As New TreeNode(elemText)
            If Not isExecutionMode Then
                elemNode.Tag = selData  ' Selection data only in preview mode
            End If
            elemNode.ToolTipText = "Elemendi failid luuakse kausta: " & elemFolder & vbCrLf & _
                                   "Kasutab ka " & sharedFiles.Count & " jagatud faili kaustast Ühine"
            
            ' Element nodes inherit bold from TreeView (base font is bold)
            
            ' Combine element-specific files and shared files into one list
            Dim allFilesForElement As New List(Of ElementReleaseLib.PlannedFile)
            allFilesForElement.AddRange(elementFiles(elem.ElementName))
            allFilesForElement.AddRange(sharedFiles)
            
            ' Add all files in unified folder structure
            ' Shared files will be gray with 🔗 row marker, own files get type colors
            AddFilesToNodeByFolderMerged(elemNode, allFilesForElement, context, showStatusPrefix:=isExecutionMode)
            
            tree.Nodes.Add(elemNode)
        Next
        
        ' Orphaned files from previous release (not in current plan; not deleted automatically)
        If context IsNot Nothing AndAlso context.RemovedFiles IsNot Nothing AndAlso context.RemovedFiles.Count > 0 Then
            Dim orphanNode As New TreeNode("EEMALDATUD (eelnev väljastus, " & context.RemovedFiles.Count & " faili)")
            orphanNode.ForeColor = Drawing.Color.DimGray
            orphanNode.Tag = "ORPHAN_ROOT"
            orphanNode.ToolTipText = "Need failid olid eelmises väljastuses, kuid pole enam praeguses plaanis." & vbCrLf &
                "Neid ei kustutata automaatselt — eemalda käsitsi, kui pole enam vaja."
            For Each rm As ElementReleaseLib.FileMappingEntry In context.RemovedFiles
                Dim ext As String = ""
                Select Case rm.FileType.ToLowerInvariant()
                    Case "part" : ext = ".ipt"
                    Case "assembly" : ext = ".iam"
                    Case "drawing" : ext = ".idw"
                    Case Else : ext = ""
                End Select
                Dim child As New TreeNode(rm.TargetName & ext & "  ←  " & rm.SourceName & "  [" & rm.FileType & If(String.IsNullOrEmpty(rm.ElementVariant), "", ", " & rm.ElementVariant) & "]")
                child.ForeColor = Drawing.Color.Gray
                child.Tag = "ORPHAN"
                orphanNode.Nodes.Add(child)
            Next
            tree.Nodes.Add(orphanNode)
        End If
        
        ' Expand all nodes
        tree.ExpandAll()
        
        tree.EndUpdate()
    End Sub
    
    
    ''' <summary>
    ''' Updates element node text based on selection state
    ''' </summary>
    Private Sub UpdateElementNodeText(node As TreeNode, selData As ElementSelectionData)
        Dim marker As String = If(selData.IsSelected, "[✓] ", "[ ] ")
        Dim elem As ExcelReaderLib.ElementConfig = selData.Element

        Dim paramStr As String = ""
        For Each kvp In elem.Parameters
            If paramStr <> "" Then paramStr &= ", "
            paramStr &= kvp.Key & "=" & kvp.Value
        Next

        Dim elemText As String = marker & elem.ElementName
        If paramStr <> "" Then elemText &= " (" & paramStr & ")"
        elemText &= " — " & selData.FileCount & " faili (" & selData.OwnFileCount & " oma + " & selData.SharedFileCount & " jagatud) → " & selData.RelativePath

        node.Text = elemText
    End Sub
    
    ''' <summary>
    ''' Counts file nodes under a tree node
    ''' </summary>
    Private Function CountFilesInNode(node As TreeNode) As Integer
        Dim count As Integer = 0
        For Each child As TreeNode In node.Nodes
            If child.Tag IsNot Nothing AndAlso TypeOf child.Tag Is TreeNodeData Then
                count += 1
            Else
                count += CountFilesInNode(child)
            End If
        Next
        Return count
    End Function
    
    ''' <summary>
    ''' Sets all elements to selected or deselected
    ''' </summary>
    Private Sub SetAllElementsSelected(tree As TreeView, selected As Boolean)
        For Each node As TreeNode In tree.Nodes
            If node.Tag IsNot Nothing AndAlso TypeOf node.Tag Is ElementSelectionData Then
                Dim selData As ElementSelectionData = CType(node.Tag, ElementSelectionData)
                selData.IsSelected = selected
                UpdateElementNodeText(node, selData)
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' Adds files to a tree node, organized by folder path, merging shared and own files
    ''' </summary>
    ''' <param name="forceTypeColors">If True, use type-based colors even for shared files (for Ühine section)</param>
    ''' <param name="showStatusPrefix">If True, add status prefix for execution mode</param>
    Private Sub AddFilesToNodeByFolderMerged(parentNode As TreeNode, _
                                              files As List(Of ElementReleaseLib.PlannedFile), _
                                              context As ElementReleaseLib.ElementReleaseContext, _
                                              Optional forceTypeColors As Boolean = False, _
                                              Optional showStatusPrefix As Boolean = False)
        ' Group files by their relative folder path
        Dim folderGroups As New Dictionary(Of String, List(Of ElementReleaseLib.PlannedFile))

        For Each f As ElementReleaseLib.PlannedFile In files
            Dim targetDir As String = System.IO.Path.GetDirectoryName(f.TargetLocalPath)
            Dim relativePath As String = GetRelativeFolderPath(targetDir, context.TargetRoot)

            If Not folderGroups.ContainsKey(relativePath) Then
                folderGroups(relativePath) = New List(Of ElementReleaseLib.PlannedFile)
            End If
            folderGroups(relativePath).Add(f)
        Next

        ' Sort folders
        Dim sortedFolders As New List(Of String)(folderGroups.Keys)
        sortedFolders.Sort()

        ' Add folder nodes with their files
        For Each folderPath As String In sortedFolders
            Dim folderFiles As List(Of ElementReleaseLib.PlannedFile) = folderGroups(folderPath)

            If String.IsNullOrEmpty(folderPath) Then
                ' Root level files - add directly
                For Each f As ElementReleaseLib.PlannedFile In folderFiles
                    Dim fileNode As TreeNode = CreateFileNodeWithDescription(f, context, forceTypeColors, showStatusPrefix)
                    parentNode.Nodes.Add(fileNode)
                Next
            Else
                ' Create nested folder structure
                Dim folderParts() As String = folderPath.Split("\"c)
                Dim currentNode As TreeNode = parentNode

                For i As Integer = 0 To folderParts.Length - 1
                    Dim folderName As String = folderParts(i)
                    If String.IsNullOrEmpty(folderName) Then Continue For

                    Dim existingFolder As TreeNode = Nothing
                    For Each child As TreeNode In currentNode.Nodes
                        If child.Tag IsNot Nothing AndAlso TypeOf child.Tag Is String AndAlso CStr(child.Tag) = "FOLDER" Then
                            If child.Text.StartsWith(folderName & " (") OrElse child.Text = folderName Then
                                existingFolder = child
                                Exit For
                            End If
                        End If
                    Next

                    If existingFolder Is Nothing Then
                        Dim isLeaf As Boolean = (i = folderParts.Length - 1)
                        Dim folderText As String = folderName
                        If isLeaf Then
                            folderText &= " (" & folderFiles.Count & ")"
                        End If
                        existingFolder = New TreeNode(folderText)
                        existingFolder.Tag = "FOLDER"
                        ' Set regular font (TreeView base is bold)
                        If m_RegularFont IsNot Nothing Then existingFolder.NodeFont = m_RegularFont
                        currentNode.Nodes.Add(existingFolder)
                    End If

                    currentNode = existingFolder
                Next

                ' Add files to deepest folder
                For Each f As ElementReleaseLib.PlannedFile In folderFiles
                    Dim fileNode As TreeNode = CreateFileNodeWithDescription(f, context, forceTypeColors, showStatusPrefix)
                    currentNode.Nodes.Add(fileNode)
                Next
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' Creates a file node with full property display
    ''' Format: [STATUS] 🔄/📄/🔗 Number | Description | Type
    ''' </summary>
    ''' <param name="forceTypeColors">If True, use type-based colors even for shared files (for Ühine section)</param>
    ''' <param name="showStatusPrefix">If True, add status prefix for execution mode</param>
    Private Function CreateFileNodeWithDescription(f As ElementReleaseLib.PlannedFile, _
                                                    context As ElementReleaseLib.ElementReleaseContext, _
                                                    Optional forceTypeColors As Boolean = False, _
                                                    Optional showStatusPrefix As Boolean = False) As TreeNode
        Dim nodeData As New TreeNodeData()
        nodeData.FilePath = f.TargetLocalPath
        nodeData.VaultNumber = f.VaultNumber
        nodeData.Status = FileStatus.Pending
        nodeData.IsShared = f.IsShared
        
        Dim fileTypeStr As String = GetFileTypeString(f.FileType)
        Select Case f.FileType
            Case ElementReleaseLib.FileType.Part
                nodeData.FileType = "Part"
            Case ElementReleaseLib.FileType.Assembly
                nodeData.FileType = "Assembly"
            Case ElementReleaseLib.FileType.Drawing
                nodeData.FileType = "Drawing"
        End Select
        
        ' Use PROJECTED Description (what file will become), fall back to source, then filename
        Dim description As String = f.ProjectedDescription
        If String.IsNullOrEmpty(description) Then description = f.SourceDescription
        If String.IsNullOrEmpty(description) Then
            description = System.IO.Path.GetFileNameWithoutExtension(f.SourcePath)
        End If
        nodeData.Description = description
        
        ' Use PROJECTED Project
        Dim project As String = f.ProjectedProject
        If String.IsNullOrEmpty(project) Then project = f.SourceProject
        If String.IsNullOrEmpty(project) Then project = ""
        
        ' Build display text: [STATUS] + row emoji (🔗/🔄/📄) + Number | Description | Type
        Dim numberDisplay As String = f.VaultNumber
        If f.IsPlaceholder Then
            numberDisplay = "[" & f.VaultNumber & "]"  ' Brackets = new number not allocated yet
        End If
        
        Dim rowEmoji As String = GetPlannedFileRowEmoji(f)
        
        ' Build base text (without execution status prefix) - used for status updates
        Dim baseText As String = rowEmoji & numberDisplay & " | " & description & " | " & fileTypeStr
        If Not String.IsNullOrEmpty(f.SourcePartNumber) AndAlso f.SourcePartNumber <> f.VaultNumber Then
            baseText &= " (alg: " & f.SourcePartNumber & ")"
        End If
        
        ' Store base text for later status updates
        nodeData.BaseText = baseText
        
        ' Build full node text with optional status prefix
        Dim nodeText As String = ""
        If showStatusPrefix Then
            nodeText = STATUS_PENDING  ' Add status prefix for execution mode
        End If
        nodeText &= baseText

        Dim fileNode As New TreeNode(nodeText)
        fileNode.Tag = nodeData
        
        ' Set regular font (TreeView base is bold for root nodes)
        If m_RegularFont IsNot Nothing Then fileNode.NodeFont = m_RegularFont
        
        ' Set styling based on file type and shared status
        ' forceTypeColors = True means use type colors even for shared files (for Ühine section)
        If f.IsShared AndAlso Not forceTypeColors Then
            ' Shared files shown as gray references (under element nodes)
            fileNode.ForeColor = Drawing.Color.Gray
        Else
            ' Color code by file type - for non-shared files or for Ühine section
            Select Case f.FileType
                Case ElementReleaseLib.FileType.Part
                    fileNode.ForeColor = Drawing.Color.DarkBlue
                Case ElementReleaseLib.FileType.Assembly
                    fileNode.ForeColor = Drawing.Color.DarkGreen
                Case ElementReleaseLib.FileType.Drawing
                    fileNode.ForeColor = Drawing.Color.DarkRed
            End Select
        End If
        
        ' Build tooltip with all properties
        Dim tooltip As String = ""
        If f.IsShared Then
            tooltip = "🔗 Jagatud fail (Ühine)." & vbCrLf & vbCrLf
        End If
        If f.IsReuse Then
            tooltip &= "🔄 Taaskasutus: kasutatakse olemasolevat Vault numbrit; sihtfail kirjutatakse üle (kui fail juba olemas)." & vbCrLf
            If System.IO.File.Exists(f.TargetLocalPath) Then
                tooltip &= "Sihtfail leitud — ülekirjutus väljastamisel." & vbCrLf
            Else
                tooltip &= "Sihtfaili pole veel — luuakse väljastamisel." & vbCrLf
            End If
            tooltip &= vbCrLf
        ElseIf f.IsPlaceholder Then
            tooltip &= "📄 Uus Vault number — määratakse pärast kinnitust (eelvaade: " & f.VaultNumber & ")." & vbCrLf & vbCrLf
        End If
        tooltip &= "Number: " & f.VaultNumber & vbCrLf & _
                   "Kirjeldus: " & description & vbCrLf & _
                   "Tüüp: " & fileTypeStr
        If Not String.IsNullOrEmpty(project) Then
            tooltip &= vbCrLf & "Projekt: " & project
        End If
        If Not String.IsNullOrEmpty(f.SourcePartNumber) Then
            tooltip &= vbCrLf & "Algne number: " & f.SourcePartNumber
        End If
        If f.SourceDescription <> description Then
            tooltip &= vbCrLf & "Algne kirjeldus: " & f.SourceDescription
        End If
        tooltip &= vbCrLf & "Allikas: " & f.SourcePath & vbCrLf & _
                   "Sihtkoht: " & f.TargetLocalPath & vbCrLf & _
                   "Märge: 🔄 taaskasutus · 📄 uus number · 🔗 jagatud"
        
        fileNode.ToolTipText = tooltip
        
        Return fileNode
    End Function
    
    ''' <summary>
    ''' Adds files to a tree node, organized by relative folder path (legacy)
    ''' </summary>
    Private Sub AddFilesToNodeByFolder(parentNode As TreeNode, _
                                       files As List(Of ElementReleaseLib.PlannedFile), _
                                       context As ElementReleaseLib.ElementReleaseContext, _
                                       isSharedRef As Boolean)
        ' Group files by their relative folder path within the target
        Dim folderGroups As New Dictionary(Of String, List(Of ElementReleaseLib.PlannedFile))
        
        For Each f As ElementReleaseLib.PlannedFile In files
            ' Get the folder path relative to the element folder
            Dim targetDir As String = System.IO.Path.GetDirectoryName(f.TargetLocalPath)
            Dim relativePath As String = GetRelativeFolderPath(targetDir, context.TargetRoot)
            
            If Not folderGroups.ContainsKey(relativePath) Then
                folderGroups(relativePath) = New List(Of ElementReleaseLib.PlannedFile)
            End If
            folderGroups(relativePath).Add(f)
        Next
        
        ' Sort folders to get a consistent order
        Dim sortedFolders As New List(Of String)(folderGroups.Keys)
        sortedFolders.Sort()
        
        ' Add folder nodes with their files
        For Each folderPath As String In sortedFolders
            Dim folderFiles As List(Of ElementReleaseLib.PlannedFile) = folderGroups(folderPath)
            
            ' If there's only one root folder or no subfolder, add files directly
            If String.IsNullOrEmpty(folderPath) OrElse Not folderPath.Contains("\") Then
                ' Get or create folder node
                Dim folderName As String = If(String.IsNullOrEmpty(folderPath), "", folderPath)
                Dim targetNode As TreeNode = parentNode
                
                If Not String.IsNullOrEmpty(folderName) Then
                    ' Create folder node with file type summary
                    Dim folderNode As TreeNode = FindOrCreateFolderNode(parentNode, folderName, folderFiles)
                    targetNode = folderNode
                End If
                
                ' Add files to folder
                For Each f As ElementReleaseLib.PlannedFile In folderFiles
                    Dim fileNode As TreeNode = CreateFileNode(f, context, isSharedRef)
                    targetNode.Nodes.Add(fileNode)
                Next
            Else
                ' Create nested folder structure
                Dim folderParts() As String = folderPath.Split("\"c)
                Dim currentNode As TreeNode = parentNode
                
                For i As Integer = 0 To folderParts.Length - 1
                    Dim folderName As String = folderParts(i)
                    If String.IsNullOrEmpty(folderName) Then Continue For
                    
                    ' Find or create this folder node
                    Dim existingFolder As TreeNode = Nothing
                    For Each child As TreeNode In currentNode.Nodes
                        If child.Text.StartsWith(folderName & " (") OrElse child.Text = folderName Then
                            existingFolder = child
                            Exit For
                        End If
                    Next
                    
                    If existingFolder Is Nothing Then
                        ' Create new folder node - only show count for leaf folders
                        Dim isLeaf As Boolean = (i = folderParts.Length - 1)
                        Dim folderText As String = folderName
                        If isLeaf Then
                            folderText &= " (" & folderFiles.Count & ")"
                        End If
                        existingFolder = New TreeNode(folderText)
                        existingFolder.Tag = "FOLDER"
                        currentNode.Nodes.Add(existingFolder)
                    End If
                    
                    currentNode = existingFolder
                Next
                
                ' Add files to the deepest folder
                For Each f As ElementReleaseLib.PlannedFile In folderFiles
                    Dim fileNode As TreeNode = CreateFileNode(f, context, isSharedRef)
                    currentNode.Nodes.Add(fileNode)
                Next
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' Gets the relative path from a base folder
    ''' </summary>
    Private Function GetRelativeFolderPath(fullPath As String, basePath As String) As String
        If String.IsNullOrEmpty(fullPath) OrElse String.IsNullOrEmpty(basePath) Then
            Return ""
        End If
        
        ' Normalize paths
        fullPath = fullPath.TrimEnd("\"c)
        basePath = basePath.TrimEnd("\"c)
        
        If fullPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase) Then
            Dim relative As String = fullPath.Substring(basePath.Length).TrimStart("\"c)
            ' Skip the first folder (element name or Ühine) to get internal structure
            Dim firstSlash As Integer = relative.IndexOf("\"c)
            If firstSlash > 0 Then
                Return relative.Substring(firstSlash + 1)
            End If
            Return ""
        End If
        
        Return System.IO.Path.GetFileName(fullPath)
    End Function
    
    ''' <summary>
    ''' Finds or creates a folder node with file count
    ''' </summary>
    Private Function FindOrCreateFolderNode(parentNode As TreeNode, folderName As String, _
                                            files As List(Of ElementReleaseLib.PlannedFile)) As TreeNode
        ' Check if folder already exists
        For Each child As TreeNode In parentNode.Nodes
            If child.Text.StartsWith(folderName & " (") Then
                Return child
            End If
        Next
        
        ' Create new folder node
        Dim folderText As String = folderName & " (" & files.Count & ")"
        Dim folderNode As New TreeNode(folderText)
        folderNode.Tag = "FOLDER"
        parentNode.Nodes.Add(folderNode)
        
        Return folderNode
    End Function
    
    ''' <summary>
    ''' Leading row marker: 🔗 shared, 🔄 reuse overwrite, 📄 new number (preview).
    ''' </summary>
    Private Function GetPlannedFileRowEmoji(f As ElementReleaseLib.PlannedFile) As String
        If f.IsShared Then Return "🔗 "
        If f.IsReuse Then Return "🔄 "
        If f.IsPlaceholder Then Return "📄 "
        Return ""
    End Function
    
    ''' <summary>
    ''' Gets file type string in Estonian
    ''' </summary>
    Private Function GetFileTypeString(fileType As ElementReleaseLib.FileType) As String
        Select Case fileType
            Case ElementReleaseLib.FileType.Part
                Return "Detail"
            Case ElementReleaseLib.FileType.Assembly
                Return "Koost"
            Case ElementReleaseLib.FileType.Drawing
                Return "Joonis"
            Case Else
                Return "Fail"
        End Select
    End Function
    
    ''' <summary>
    ''' Creates a tree node for a planned file with full details
    ''' Format: 🔄/📄/🔗 + PartNumber | SourceName | Type | → TargetFileName
    ''' </summary>
    Private Function CreateFileNode(f As ElementReleaseLib.PlannedFile, _
                                    context As ElementReleaseLib.ElementReleaseContext, _
                                    Optional isSharedRef As Boolean = False) As TreeNode
        Dim nodeData As New TreeNodeData()
        nodeData.FilePath = f.TargetLocalPath
        nodeData.VaultNumber = f.VaultNumber
        nodeData.Status = FileStatus.Pending
        
        ' Determine file type
        Dim fileTypeStr As String = GetFileTypeString(f.FileType)
        Select Case f.FileType
            Case ElementReleaseLib.FileType.Part
                nodeData.FileType = "Part"
            Case ElementReleaseLib.FileType.Assembly
                nodeData.FileType = "Assembly"
            Case ElementReleaseLib.FileType.Drawing
                nodeData.FileType = "Drawing"
        End Select
        
        ' Get source file name for description context
        Dim sourceFileName As String = System.IO.Path.GetFileNameWithoutExtension(f.SourcePath)
        Dim targetFileName As String = System.IO.Path.GetFileName(f.TargetLocalPath)
        
        Dim rowEmoji As String = GetPlannedFileRowEmoji(f)
        If isSharedRef AndAlso rowEmoji = "" Then
            rowEmoji = "🔗 "
        End If
        
        ' Build descriptive text with all relevant info
        Dim nodeText As String
        If isSharedRef Then
            nodeText = "    " & rowEmoji & f.VaultNumber & " | " & sourceFileName & " | " & fileTypeStr
        Else
            nodeText = rowEmoji & f.VaultNumber & " | " & sourceFileName & " | " & fileTypeStr & " → " & targetFileName
        End If
        
        Dim fileNode As New TreeNode(nodeText)
        fileNode.Tag = nodeData
        Dim tip As String = ""
        If f.IsShared Then tip &= "🔗 Jagatud fail (Ühine)." & vbCrLf
        If f.IsReuse Then tip &= "🔄 Taaskasutus." & vbCrLf
        If f.IsPlaceholder Then tip &= "📄 Uus number (eelvaade)." & vbCrLf
        If tip <> "" Then tip &= vbCrLf
        tip &= "Allikas: " & f.SourcePath & vbCrLf & _
               "Sihtkoht: " & f.TargetLocalPath & vbCrLf & _
               "Part Number: " & f.VaultNumber & vbCrLf & _
               "Tüüp: " & fileTypeStr & vbCrLf & _
               "Märge: 🔄 taaskasutus · 📄 uus number · 🔗 jagatud"
        fileNode.ToolTipText = tip
        
        Return fileNode
    End Function
    
    ''' <summary>
    ''' Sets all element nodes (top-level with ElementConfig tag) to checked/unchecked
    ''' </summary>
    ''' <summary>
    ''' Updates the stats label based on selected elements in tree (using ElementSelectionData)
    ''' </summary>
    Private Sub UpdateStats(frm As Form)
        Dim data As Dictionary(Of String, Object) = CType(frm.Tag, Dictionary(Of String, Object))
        Dim treeView As TreeView = CType(data("treeView"), TreeView)
        Dim lblStats As Label = CType(data("lblStats"), Label)
        Dim plan As ElementReleaseLib.ReleasePlan = CType(data("plan"), ElementReleaseLib.ReleasePlan)
        
        ' Count selected elements using ElementSelectionData
        Dim selectedElements As Integer = 0
        Dim selectedNames As New List(Of String)
        
        For Each node As TreeNode In treeView.Nodes
            If node.Tag IsNot Nothing AndAlso TypeOf node.Tag Is ElementSelectionData Then
                Dim selData As ElementSelectionData = CType(node.Tag, ElementSelectionData)
                If selData.IsSelected Then
                    selectedElements += 1
                    selectedNames.Add(selData.Element.ElementName)
                End If
            End If
        Next
        
        ' Count files for selected elements
        Dim sharedFileCount As Integer = 0
        Dim elementFileCount As Integer = 0
        
        For Each f As ElementReleaseLib.PlannedFile In plan.Files
            If f.IsShared Then
                ' Shared files included if ANY element is selected
                If selectedElements > 0 Then sharedFileCount += 1
            Else
                ' Element-specific files
                For Each varName In f.ForVariants
                    If selectedNames.Contains(varName) Then
                        elementFileCount += 1
                        Exit For
                    End If
                Next
            End If
        Next
        
        Dim reuseSelected As Integer = 0
        Dim newSelected As Integer = 0
        For Each f As ElementReleaseLib.PlannedFile In plan.Files
            Dim fileApplies As Boolean = False
            If f.IsShared Then
                fileApplies = (selectedElements > 0)
            Else
                For Each varName As String In f.ForVariants
                    If selectedNames.Contains(varName) Then
                        fileApplies = True
                        Exit For
                    End If
                Next
            End If
            If Not fileApplies Then Continue For
            If f.IsReuse Then reuseSelected += 1
            If f.IsPlaceholder Then newSelected += 1
        Next
        
        Dim totalFiles As Integer = sharedFileCount + elementFileCount
        lblStats.Text = String.Format("Valitud: {0} elementi | Faile kokku: {1} (jagatud: {2}, elemendispetsiifilised: {3}) | Taaskasutus: {4} | Uued nr: {5}", _
            selectedElements, totalFiles, sharedFileCount, elementFileCount, reuseSelected, newSelected)
        
        ' Update shared node appearance based on selection
        UpdateSharedNodeAppearance(treeView, selectedElements > 0)
    End Sub
    
    ''' <summary>
    ''' Updates the shared (Ühine) node appearance based on whether any elements are selected
    ''' </summary>
    Private Sub UpdateSharedNodeAppearance(tree As TreeView, hasSelectedElements As Boolean)
        For Each node As TreeNode In tree.Nodes
            If node.Tag IsNot Nothing AndAlso TypeOf node.Tag Is String AndAlso CStr(node.Tag) = "SHARED_ROOT" Then
                If hasSelectedElements Then
                    ' Shared files will be created - show normally with type-based colors
                    node.ForeColor = Drawing.Color.DarkOrange
                    ' Restore type-based colors for file nodes
                    RestoreTypeBasedColors(node.Nodes)
                Else
                    ' No elements selected - shared files won't be created
                    node.ForeColor = Drawing.Color.LightGray
                    ' Gray out all child nodes
                    SetNodeTreeColor(node.Nodes, Drawing.Color.LightGray)
                End If
                Exit For
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' Recursively sets color for all nodes in a collection
    ''' </summary>
    Private Sub SetNodeTreeColor(nodes As TreeNodeCollection, color As Drawing.Color)
        For Each node As TreeNode In nodes
            node.ForeColor = color
            If node.Nodes.Count > 0 Then
                SetNodeTreeColor(node.Nodes, color)
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' Recursively restores type-based colors for file nodes
    ''' </summary>
    Private Sub RestoreTypeBasedColors(nodes As TreeNodeCollection)
        For Each node As TreeNode In nodes
            If node.Tag IsNot Nothing AndAlso TypeOf node.Tag Is TreeNodeData Then
                Dim data As TreeNodeData = DirectCast(node.Tag, TreeNodeData)
                Select Case data.FileType
                    Case "Part"
                        node.ForeColor = Drawing.Color.DarkBlue
                    Case "Assembly"
                        node.ForeColor = Drawing.Color.DarkGreen
                    Case "Drawing"
                        node.ForeColor = Drawing.Color.DarkRed
                    Case Else
                        node.ForeColor = Drawing.Color.Black
                End Select
            End If
            If node.Nodes.Count > 0 Then
                RestoreTypeBasedColors(node.Nodes)
            End If
        Next
    End Sub
    
    ''' <summary>
    ''' Populates the files tree view with the release plan (legacy - for execution form)
    ''' </summary>
    Private Sub PopulateFilesTree(tree As TreeView, plan As ElementReleaseLib.ReleasePlan, _
                                  context As ElementReleaseLib.ElementReleaseContext)
        tree.BeginUpdate()
        tree.Nodes.Clear()
        
        Dim rootNode As New TreeNode("Elemendid")
        tree.Nodes.Add(rootNode)
        
        Dim sharedNode As New TreeNode("Ühine (Jagatud)")
        Dim elementNodes As New Dictionary(Of String, TreeNode)
        For Each elem As ExcelReaderLib.ElementConfig In context.Elements
            Dim elemNode As New TreeNode(elem.ElementName)
            elemNode.Tag = elem.ElementName
            elementNodes(elem.ElementName) = elemNode
        Next
        
        For Each f As ElementReleaseLib.PlannedFile In plan.Files
            Dim nodeData As New TreeNodeData()
            nodeData.FilePath = f.TargetLocalPath
            nodeData.VaultNumber = f.VaultNumber
            nodeData.Status = FileStatus.Pending
            
            Dim fileName As String = System.IO.Path.GetFileName(f.TargetLocalPath)
            Dim nodeText As String = STATUS_PENDING & f.VaultNumber & " - " & fileName
            
            Select Case f.FileType
                Case ElementReleaseLib.FileType.Part
                    nodeData.FileType = "Part"
                Case ElementReleaseLib.FileType.Assembly
                    nodeData.FileType = "Assembly"
                Case ElementReleaseLib.FileType.Drawing
                    nodeData.FileType = "Drawing"
            End Select
            
            Dim fileNode As New TreeNode(nodeText)
            fileNode.Tag = nodeData
            fileNode.ToolTipText = f.TargetLocalPath
            
            If f.IsShared Then
                nodeData.ElementName = "Ühine"
                sharedNode.Nodes.Add(fileNode)
            Else
                For Each elemName In f.ForVariants
                    If elementNodes.ContainsKey(elemName) Then
                        nodeData.ElementName = elemName
                        elementNodes(elemName).Nodes.Add(fileNode)
                        Exit For
                    End If
                Next
            End If
        Next
        
        If sharedNode.Nodes.Count > 0 Then
            rootNode.Nodes.Add(sharedNode)
        End If
        
        For Each elemName As String In elementNodes.Keys
            If elementNodes(elemName).Nodes.Count > 0 Then
                rootNode.Nodes.Add(elementNodes(elemName))
            End If
        Next
        
        rootNode.ExpandAll()
        tree.EndUpdate()
    End Sub
    
    ''' <summary>
    ''' Updates stats label based on selected elements (legacy)
    ''' </summary>
    Private Sub UpdateStatsForSelection(frm As Form)
        Dim data As Dictionary(Of String, Object) = CType(frm.Tag, Dictionary(Of String, Object))
        Dim elementsListView As ListView = CType(data("elementsListView"), ListView)
        Dim lblStats As Label = CType(data("lblStats"), Label)
        Dim plan As ElementReleaseLib.ReleasePlan = CType(data("plan"), ElementReleaseLib.ReleasePlan)
        
        ' Count selected elements
        Dim selectedCount As Integer = 0
        Dim selectedNames As New HashSet(Of String)
        For Each item As ListViewItem In elementsListView.Items
            If item.Checked Then
                selectedCount += 1
                selectedNames.Add(CType(item.Tag, ExcelReaderLib.ElementConfig).ElementName)
            End If
        Next
        
        ' Count files for selected elements
        Dim fileCount As Integer = 0
        For Each f As ElementReleaseLib.PlannedFile In plan.Files
            If f.IsShared Then
                ' Shared files are always included if any element is selected
                If selectedCount > 0 Then fileCount += 1
            Else
                ' Element-specific files
                For Each elemName In f.ForVariants
                    If selectedNames.Contains(elemName) Then
                        fileCount += 1
                        Exit For
                    End If
                Next
            End If
        Next
        
        lblStats.Text = String.Format("Valitud: {0} elementi | Faile: {1}", selectedCount, fileCount)
    End Sub
    
    ''' <summary>
    ''' Handles Execute button click
    ''' </summary>
    Private Sub OnExecuteClick(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim frm As Form = btn.FindForm()
        Dim data As Dictionary(Of String, Object) = CType(frm.Tag, Dictionary(Of String, Object))
        
        Dim result As ReleaseUIResult = CType(data("result"), ReleaseUIResult)
        Dim treeView As TreeView = CType(data("treeView"), TreeView)
        Dim btnExecute As Button = CType(data("btnExecute"), Button)
        Dim btnCancel As Button = CType(data("btnCancel"), Button)
        
        ' Get selected elements from ElementSelectionData
        result.SelectedElements.Clear()
        For Each node As TreeNode In treeView.Nodes
            If node.Tag IsNot Nothing AndAlso TypeOf node.Tag Is ElementSelectionData Then
                Dim selData As ElementSelectionData = CType(node.Tag, ElementSelectionData)
                If selData.IsSelected Then
                    result.SelectedElements.Add(selData.Element)
                End If
            End If
        Next
        
        If result.SelectedElements.Count = 0 Then
            MessageBox.Show("Vali vähemalt üks element.", "Loo elemendid", _
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        
        ' Disable UI during execution
        btnExecute.Enabled = False
        btnCancel.Text = "Sulge"
        treeView.Enabled = False
        
        ' Mark execution started
        result.ExecutionCompleted = False
        result.Cancelled = False
        
        ' Store ready-to-execute flag for caller to check
        data("readyToExecute") = True
        
        ' Close the dialog - caller will proceed with execution
        frm.DialogResult = DialogResult.OK
        frm.Close()
    End Sub
    
    ' ============================================================================
    ' Execution Form (stays open during release with progress)
    ' ============================================================================
    
    Private m_ExecutionForm As Form = Nothing
    Private m_TreeView As TreeView = Nothing
    Private m_ProgressBar As System.Windows.Forms.ProgressBar = Nothing
    Private m_ProgressLabel As Label = Nothing
    Private m_LogTextBox As RichTextBox = Nothing
    
    ' UI update throttling - prevents excessive DoEvents calls during batch processing
    Private m_LastUIUpdate As DateTime = DateTime.MinValue
    Private Const UI_THROTTLE_MS As Integer = 2000  ' Update UI max every 2 seconds
    Private m_TotalFiles As Integer = 0
    Private m_CompletedFiles As Integer = 0
    Private m_FailedFiles As Integer = 0
    
    ''' <summary>
    ''' Creates and shows the execution form that stays open during release
    ''' </summary>
    Public Function ShowExecutionForm(context As ElementReleaseLib.ElementReleaseContext, _
                                       filteredPlan As ElementReleaseLib.ReleasePlan, _
                                       selectedElements As List(Of ExcelReaderLib.ElementConfig)) As Form
        
        Dim frm As New Form()
        frm.Text = StringsLib.TITLE_ELEMENT_RELEASE & " - " & context.ElementName
        frm.Width = 800
        frm.Height = 600
        frm.StartPosition = FormStartPosition.CenterScreen
        frm.FormBorderStyle = FormBorderStyle.Sizable
        frm.TopMost = True
        UILib.SetMinimumSize(frm, 600, 400)
        
        ' Main split container: tree and log
        Dim mainSplit As New SplitContainer()
        mainSplit.Dock = DockStyle.Fill
        mainSplit.Orientation = Orientation.Horizontal
        mainSplit.SplitterDistance = 350
        mainSplit.SplitterWidth = 6
        frm.Controls.Add(mainSplit)
        
        ' ========================================
        ' TOP PANEL: Files Tree
        ' ========================================
        Dim treePanel As New Panel()
        treePanel.Dock = DockStyle.Fill
        treePanel.Padding = New Padding(10)
        mainSplit.Panel1.Controls.Add(treePanel)
        
        ' Header
        Dim lblHeader As New Label()
        lblHeader.Text = "FAILIDE VÄLJASTAMINE"
        lblHeader.Dock = DockStyle.Top
        lblHeader.Height = 25
        treePanel.Controls.Add(lblHeader)
        
        ' Progress bar
        Dim progressBar As New System.Windows.Forms.ProgressBar()
        progressBar.Dock = DockStyle.Top
        progressBar.Height = 25
        progressBar.Minimum = 0
        progressBar.Maximum = filteredPlan.Files.Count
        progressBar.Value = 0
        treePanel.Controls.Add(progressBar)
        
        ' Progress label
        Dim lblProgress As New Label()
        lblProgress.Dock = DockStyle.Top
        lblProgress.Height = 25
        lblProgress.Text = String.Format("0 / {0} faili", filteredPlan.Files.Count)
        treePanel.Controls.Add(lblProgress)
        
        ' Tree view (no ImageList due to iLogic constraints)
        Dim filesTree As New TreeView()
        filesTree.Dock = DockStyle.Fill
        filesTree.ShowLines = True
        filesTree.ShowPlusMinus = True
        filesTree.ShowRootLines = True
        filesTree.FullRowSelect = True
        filesTree.HideSelection = False
        filesTree.ShowNodeToolTips = True
        filesTree.Scrollable = True
        ' Set TreeView base font to bold (same as preview)
        m_RegularFont = New Font(filesTree.Font, FontStyle.Regular)
        filesTree.Font = New Font(filesTree.Font, FontStyle.Bold)
        treePanel.Controls.Add(filesTree)
        
        ' Populate the tree using unified structure (same as preview, with status indicators)
        PopulateUnifiedTree(filesTree, selectedElements, filteredPlan, context, isExecutionMode:=True)
        
        ' Set control order
        treePanel.Controls.SetChildIndex(filesTree, 0)
        treePanel.Controls.SetChildIndex(lblProgress, 1)
        treePanel.Controls.SetChildIndex(progressBar, 2)
        treePanel.Controls.SetChildIndex(lblHeader, 3)
        
        ' ========================================
        ' BOTTOM PANEL: Log
        ' ========================================
        Dim logPanel As New Panel()
        logPanel.Dock = DockStyle.Fill
        logPanel.Padding = New Padding(10)
        mainSplit.Panel2.Controls.Add(logPanel)
        
        Dim lblLog As New Label()
        lblLog.Text = "LOGI"
        lblLog.Dock = DockStyle.Top
        lblLog.Height = 25
        logPanel.Controls.Add(lblLog)
        
        Dim txtLog As New RichTextBox()
        txtLog.Dock = DockStyle.Fill
        txtLog.ReadOnly = True
        logPanel.Controls.Add(txtLog)
        
        logPanel.Controls.SetChildIndex(txtLog, 0)
        logPanel.Controls.SetChildIndex(lblLog, 1)
        
        ' ========================================
        ' BOTTOM: Close Button
        ' ========================================
        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.Dock = DockStyle.Bottom
        buttonPanel.Height = 45
        buttonPanel.FlowDirection = FlowDirection.RightToLeft
        buttonPanel.Padding = New Padding(10, 5, 10, 5)
        frm.Controls.Add(buttonPanel)
        
        Dim btnClose As New Button()
        btnClose.Text = "Sulge"
        btnClose.Width = 100
        btnClose.Height = 30
        btnClose.Enabled = False ' Disabled until execution completes
        buttonPanel.Controls.Add(btnClose)
        
        AddHandler btnClose.Click, Sub(s, e)
            frm.Close()
        End Sub
        
        ' Store references
        m_ExecutionForm = frm
        m_TreeView = filesTree
        m_ProgressBar = progressBar
        m_ProgressLabel = lblProgress
        m_LogTextBox = txtLog
        m_TotalFiles = filteredPlan.Files.Count
        m_CompletedFiles = 0
        m_FailedFiles = 0
        
        frm.Tag = btnClose ' Store close button reference
        
        ' Show non-modal
        frm.Show()
        System.Windows.Forms.Application.DoEvents()
        
        Return frm
    End Function
    
    ''' <summary>
    ''' Updates a file's status in the execution tree
    ''' </summary>
    Public Sub UpdateFileStatus(targetPath As String, status As FileStatus, Optional message As String = Nothing)
        If m_ExecutionForm Is Nothing OrElse m_ExecutionForm.IsDisposed Then Return
        If m_TreeView Is Nothing Then Return
        
        ' Find the node by path
        Dim node As TreeNode = FindNodeByPath(m_TreeView.Nodes, targetPath)
        If node IsNot Nothing Then
            Dim nodeData As TreeNodeData = CType(node.Tag, TreeNodeData)
            nodeData.Status = status
            
            ' Update node text with status prefix and set color
            ' Use stored BaseText to preserve original formatting (description, type, etc.)
            Dim baseText As String = nodeData.BaseText
            If String.IsNullOrEmpty(baseText) Then
                ' Fallback if BaseText wasn't set
                baseText = nodeData.VaultNumber & " - " & System.IO.Path.GetFileName(nodeData.FilePath)
            End If
            
            Select Case status
                Case FileStatus.InProgress
                    node.Text = STATUS_INPROGRESS & baseText
                    node.ForeColor = Drawing.Color.Blue
                Case FileStatus.Completed
                    node.Text = STATUS_COMPLETED & baseText
                    node.ForeColor = Drawing.Color.Green
                    m_CompletedFiles += 1
                Case FileStatus.Failed
                    node.Text = STATUS_FAILED & baseText
                    node.ForeColor = Drawing.Color.Red
                    m_FailedFiles += 1
                Case FileStatus.Skipped
                    node.Text = STATUS_SKIPPED & baseText
                    node.ForeColor = Drawing.Color.Gray
            End Select
            
            ' Ensure node is visible
            node.EnsureVisible()
        End If
        
        ' Update progress
        If m_ProgressBar IsNot Nothing Then
            m_ProgressBar.Value = Math.Min(m_CompletedFiles + m_FailedFiles, m_ProgressBar.Maximum)
        End If
        
        If m_ProgressLabel IsNot Nothing Then
            m_ProgressLabel.Text = String.Format("{0} / {1} faili ({2} ebaõnnestus)", _
                m_CompletedFiles + m_FailedFiles, m_TotalFiles, m_FailedFiles)
        End If
        
        ' Add to log if message provided
        If message IsNot Nothing AndAlso m_LogTextBox IsNot Nothing Then
            m_LogTextBox.AppendText(message & vbCrLf)
            m_LogTextBox.ScrollToCaret()
        End If
        
        ' Throttled UI update - only call DoEvents every few seconds to improve performance
        If (DateTime.Now - m_LastUIUpdate).TotalMilliseconds > UI_THROTTLE_MS Then
            System.Windows.Forms.Application.DoEvents()
            m_LastUIUpdate = DateTime.Now
        End If
    End Sub
    
    Private Function FindNodeByPath(nodes As TreeNodeCollection, targetPath As String) As TreeNode
        For Each node As TreeNode In nodes
            If node.Tag IsNot Nothing AndAlso TypeOf node.Tag Is TreeNodeData Then
                Dim data As TreeNodeData = CType(node.Tag, TreeNodeData)
                If data.FilePath IsNot Nothing AndAlso _
                   data.FilePath.Equals(targetPath, StringComparison.OrdinalIgnoreCase) Then
                    Return node
                End If
            End If
            
            If node.Nodes.Count > 0 Then
                Dim found As TreeNode = FindNodeByPath(node.Nodes, targetPath)
                If found IsNot Nothing Then Return found
            End If
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Adds a log message to the execution form
    ''' </summary>
    Public Sub LogMessage(message As String)
        If m_LogTextBox Is Nothing OrElse m_LogTextBox.IsDisposed Then Return
        m_LogTextBox.AppendText(message & vbCrLf)
        m_LogTextBox.ScrollToCaret()
        
        ' Throttled UI update
        If (DateTime.Now - m_LastUIUpdate).TotalMilliseconds > UI_THROTTLE_MS Then
            System.Windows.Forms.Application.DoEvents()
            m_LastUIUpdate = DateTime.Now
        End If
    End Sub
    
    ''' <summary>
    ''' Marks execution as complete and enables the close button
    ''' </summary>
    Public Sub MarkExecutionComplete(success As Boolean)
        If m_ExecutionForm Is Nothing OrElse m_ExecutionForm.IsDisposed Then Return
        
        ' Enable close button
        Dim btnClose As Button = CType(m_ExecutionForm.Tag, Button)
        If btnClose IsNot Nothing Then
            btnClose.Enabled = True
        End If
        
        ' Update progress label (no color change due to iLogic System.Drawing constraints)
        If m_ProgressLabel IsNot Nothing Then
            If success AndAlso m_FailedFiles = 0 Then
                m_ProgressLabel.Text = "OK: " & String.Format("Valmis! {0} faili edukalt loodud.", m_CompletedFiles)
            Else
                m_ProgressLabel.Text = "VEAD: " & String.Format("Lõpetatud. Õnnestus: {0}, Ebaõnnestus: {1}", _
                    m_CompletedFiles, m_FailedFiles)
            End If
        End If
        
        ' Log summary
        LogMessage("")
        LogMessage("=== KOKKUVÕTE ===")
        LogMessage(String.Format("Edukalt loodud: {0}", m_CompletedFiles))
        If m_FailedFiles > 0 Then
            LogMessage(String.Format("Ebaõnnestunud: {0}", m_FailedFiles))
        End If
        
        ' Allow window to be non-topmost now
        m_ExecutionForm.TopMost = False
        
        System.Windows.Forms.Application.DoEvents()
    End Sub
    
    ''' <summary>
    ''' Waits for user to close the execution form
    ''' </summary>
    Public Sub WaitForExecutionFormClose()
        If m_ExecutionForm Is Nothing Then Return
        
        Do While m_ExecutionForm.Visible
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(50)
        Loop
        
        ' Clean up references
        m_ExecutionForm = Nothing
        m_TreeView = Nothing
        m_ProgressBar = Nothing
        m_ProgressLabel = Nothing
        m_LogTextBox = Nothing
    End Sub
    
    ' ============================================================================
    ' Progress Tracking UI (for use during execution)
    ' ============================================================================
    
    ''' <summary>
    ''' Creates a progress window that stays open during execution
    ''' </summary>
    Public Function CreateProgressWindow(elementName As String, totalFiles As Integer) As Form
        Dim frm As New Form()
        frm.Text = StringsLib.TITLE_ELEMENT_RELEASE & " - Progress"
        frm.Width = 600
        frm.Height = 400
        frm.StartPosition = FormStartPosition.CenterScreen
        frm.FormBorderStyle = FormBorderStyle.Sizable
        frm.TopMost = True
        UILib.SetMinimumSize(frm, 400, 300)
        
        Dim panel As New Panel()
        panel.Dock = DockStyle.Fill
        panel.Padding = New Padding(10)
        frm.Controls.Add(panel)
        
        ' Progress bar
        Dim progressBar As New System.Windows.Forms.ProgressBar()
        progressBar.Name = "progressBar"
        progressBar.Dock = DockStyle.Top
        progressBar.Height = 25
        progressBar.Minimum = 0
        progressBar.Maximum = totalFiles
        progressBar.Value = 0
        panel.Controls.Add(progressBar)
        
        ' Progress label
        Dim lblProgress As New Label()
        lblProgress.Name = "lblProgress"
        lblProgress.Dock = DockStyle.Top
        lblProgress.Height = 25
        lblProgress.Text = String.Format("0 / {0} faili", totalFiles)
        panel.Controls.Add(lblProgress)
        
        ' Status label
        Dim lblStatus As New Label()
        lblStatus.Name = "lblStatus"
        lblStatus.Dock = DockStyle.Top
        lblStatus.Height = 25
        lblStatus.Text = "Alustamine..."
        panel.Controls.Add(lblStatus)
        
        ' Log text box
        Dim txtLog As New RichTextBox()
        txtLog.Name = "txtLog"
        txtLog.Dock = DockStyle.Fill
        txtLog.ReadOnly = True
        panel.Controls.Add(txtLog)
        
        ' Ensure proper control order
        panel.Controls.SetChildIndex(txtLog, 0)
        panel.Controls.SetChildIndex(lblStatus, 1)
        panel.Controls.SetChildIndex(lblProgress, 2)
        panel.Controls.SetChildIndex(progressBar, 3)
        
        ' Close button (initially hidden, shown when done)
        Dim btnClose As New Button()
        btnClose.Name = "btnClose"
        btnClose.Text = "Sulge"
        btnClose.Width = 100
        btnClose.Height = 30
        btnClose.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        btnClose.Left = frm.Width - btnClose.Width - 40
        btnClose.Top = frm.Height - btnClose.Height - 60
        btnClose.Visible = False
        frm.Controls.Add(btnClose)
        btnClose.BringToFront()
        
        AddHandler btnClose.Click, Sub(s, e)
            frm.Close()
        End Sub
        
        Return frm
    End Function
    
    ''' <summary>
    ''' Updates progress window with current status
    ''' </summary>
    Public Sub UpdateProgress(frm As Form, currentFile As Integer, totalFiles As Integer, _
                              status As String, Optional logMessage As String = Nothing)
        If frm Is Nothing OrElse frm.IsDisposed Then Return
        
        ' Update progress bar
        Dim progressBar As System.Windows.Forms.ProgressBar = CType(frm.Controls.Find("progressBar", True)(0), System.Windows.Forms.ProgressBar)
        If currentFile <= progressBar.Maximum Then
            progressBar.Value = currentFile
        End If
        
        ' Update progress label
        Dim lblProgress As Label = CType(frm.Controls.Find("lblProgress", True)(0), Label)
        lblProgress.Text = String.Format("{0} / {1} faili", currentFile, totalFiles)
        
        ' Update status label
        Dim lblStatus As Label = CType(frm.Controls.Find("lblStatus", True)(0), Label)
        lblStatus.Text = status
        
        ' Add to log if message provided
        If logMessage IsNot Nothing Then
            Dim txtLog As RichTextBox = CType(frm.Controls.Find("txtLog", True)(0), RichTextBox)
            txtLog.AppendText(logMessage & vbCrLf)
            txtLog.ScrollToCaret()
        End If
        
        ' Process events to keep UI responsive
        System.Windows.Forms.Application.DoEvents()
    End Sub
    
    ''' <summary>
    ''' Marks progress window as completed
    ''' </summary>
    Public Sub CompleteProgress(frm As Form, successCount As Integer, failureCount As Integer, _
                                Optional failedFiles As List(Of String) = Nothing)
        If frm Is Nothing OrElse frm.IsDisposed Then Return
        
        ' Update status (no color change due to iLogic System.Drawing constraints)
        Dim lblStatus As Label = CType(frm.Controls.Find("lblStatus", True)(0), Label)
        If failureCount = 0 Then
            lblStatus.Text = "OK: " & String.Format("Valmis! {0} faili edukalt loodud.", successCount)
        Else
            lblStatus.Text = "VEAD: " & String.Format("Lõpetatud. Õnnestus: {0}, Ebaõnnestus: {1}", _
                successCount, failureCount)
        End If
        
        ' Log summary
        Dim txtLog As RichTextBox = CType(frm.Controls.Find("txtLog", True)(0), RichTextBox)
        txtLog.AppendText(vbCrLf)
        txtLog.AppendText("=== KOKKUVÕTE ===" & vbCrLf)
        txtLog.AppendText(String.Format("Edukalt loodud: {0}" & vbCrLf, successCount))
        If failureCount > 0 Then
            txtLog.AppendText(String.Format("Ebaõnnestunud: {0}" & vbCrLf, failureCount))
            If failedFiles IsNot Nothing Then
                For Each f As String In failedFiles
                    txtLog.AppendText("  - " & f & vbCrLf)
                Next
            End If
        End If
        txtLog.ScrollToCaret()
        
        ' Show close button
        Dim btnClose As Button = CType(frm.Controls.Find("btnClose", True)(0), Button)
        btnClose.Visible = True
        
        ' Allow form to be closed
        frm.TopMost = False
    End Sub
    
    ''' <summary>
    ''' Shows progress window non-modally and returns it
    ''' </summary>
    Public Function ShowProgressWindow(elementName As String, totalFiles As Integer) As Form
        Dim frm As Form = CreateProgressWindow(elementName, totalFiles)
        frm.Show()
        System.Windows.Forms.Application.DoEvents()
        Return frm
    End Function

End Module
