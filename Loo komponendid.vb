' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Loo komponendid - Create components from multi-body master part
' 
' Features:
' - Detect and display dimensions (T/W/L) for each solid body
' - Generate Vault file numbers
' - Create derived parts with iProperties
' - Optional sheet metal conversion with auto A-side detection
' - Material assignment (from SoftcomMaterials library)
' - Assembly placement
' - Axis override via face picking
'
' Usage: Run from an open multi-body part document
'        Optionally select specific solid bodies before running
' ============================================================================

' References must come FIRST, before any AddVbFile
AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
AddReference "Connectivity.InventorAddin.EdmAddin"

' Libraries come after references (UtilsLib before VaultNumberingLib for Vault logging)
AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/DocumentUpdateLib.vb"
AddVbFile "Lib/DimensionUpdateLib.vb"
AddVbFile "Lib/CustomPropertiesLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"
AddVbFile "Lib/FileSearchLib.vb"
AddVbFile "Lib/MakeComponentsLib.vb"
AddVbFile "Lib/SheetMetalLib.vb"
AddVbFile "Lib/BoundingBoxStockLib.vb"
AddVbFile "Lib/OccurrenceNamingLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Loo komponendid: No active document")
        MessageBox.Show("Ava esmalt multi-body detail.", "Loo komponendid")
        Exit Sub
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        UtilsLib.LogError("Loo komponendid: Active document is not a part")
        MessageBox.Show("Aktiivseks dokumendiks peab olema detail (.ipt).", "Loo komponendid")
        Exit Sub
    End If
    
    Dim masterDoc As PartDocument = CType(app.ActiveDocument, PartDocument)
    
    If masterDoc.ComponentDefinition.SurfaceBodies.Count < 1 Then
        UtilsLib.LogError("Loo komponendid: No solid bodies in part")
        MessageBox.Show("Detailis puuduvad tahked kehad.", "Loo komponendid")
        Exit Sub
    End If
    
    ' Ensure master is saved
    If String.IsNullOrEmpty(masterDoc.FullDocumentName) Then
        UtilsLib.LogError("Loo komponendid: Master document not saved")
        MessageBox.Show("Salvesta esmalt master-detail.", "Loo komponendid")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Loo komponendid: Starting for " & masterDoc.DisplayName)
    
    ' Get selected bodies from SelectSet (if any)
    Dim selectedBodyNames As New List(Of String)
    For Each obj As Object In masterDoc.SelectSet
        If TypeOf obj Is SurfaceBody Then
            selectedBodyNames.Add(CType(obj, SurfaceBody).Name)
        End If
    Next
    
    If selectedBodyNames.Count > 0 Then
        UtilsLib.LogInfo("Loo komponendid: User selected " & selectedBodyNames.Count & " body(ies)")
    Else
        UtilsLib.LogInfo("Loo komponendid: No selection, using all " & masterDoc.ComponentDefinition.SurfaceBodies.Count & " body(ies)")
    End If
    
    ' Get Vault connection
    Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
    Dim vaultConnected As Boolean = (vaultConn IsNot Nothing)
    
    If vaultConnected Then
        UtilsLib.LogInfo("Loo komponendid: Vault connected - " & VaultNumberingLib.GetConnectionInfo(vaultConn))
    Else
        UtilsLib.LogWarn("Loo komponendid: Vault not connected - manual filenames required")
    End If
    
    ' Get workspace root for Vault path conversion
    ' This detects the actual Vault root by testing path prefixes against Vault
    Dim workspaceRoot As String = ""
    If vaultConnected Then
        Dim docFolder As String = System.IO.Path.GetDirectoryName(masterDoc.FullDocumentName)
        workspaceRoot = VaultNumberingLib.DetectWorkspaceRoot(vaultConn, docFolder)
    End If
    
    If String.IsNullOrEmpty(workspaceRoot) Then
        UtilsLib.LogWarn("Loo komponendid: Could not detect Vault workspace root")
    Else
        UtilsLib.LogInfo("Loo komponendid: Workspace root: " & workspaceRoot)
    End If
    
    ' Get all bodies with detected axes
    Dim allBodies As List(Of MakeComponentsLib.BodyInfo) = MakeComponentsLib.GetBodiesWithAxes(masterDoc)
    
    ' Filter to selected bodies if any were selected
    Dim bodies As List(Of MakeComponentsLib.BodyInfo)
    If selectedBodyNames.Count > 0 Then
        bodies = New List(Of MakeComponentsLib.BodyInfo)
        For Each bi As MakeComponentsLib.BodyInfo In allBodies
            If selectedBodyNames.Contains(bi.Name) Then
                bodies.Add(bi)
            End If
        Next
    Else
        bodies = allBodies
    End If
    
    ' Load stored settings from master document
    Dim storedData As List(Of MakeComponentsLib.StoredBodyData) = _
        MakeComponentsLib.LoadBodyDataFromMaster(masterDoc)
    
    ' Apply stored settings to bodies (matches by name, then by geometry signature)
    ' Use depth-first search: start from master folder, limit to workspace root
    Dim masterFolder As String = System.IO.Path.GetDirectoryName(masterDoc.FullDocumentName)
    Dim vaultRoot As String = If(Not String.IsNullOrEmpty(workspaceRoot), workspaceRoot, masterFolder)
    
    ' Get project root for relative path resolution (e.g., "C:\_SoftcomVault\Tooted\Lume")
    Dim projectRoot As String = UtilsLib.GetProjectPath(masterDoc.FullDocumentName)
    If String.IsNullOrEmpty(projectRoot) Then
        UtilsLib.LogWarn("Loo komponendid: Could not detect project root, using master folder")
        projectRoot = masterFolder
    Else
        UtilsLib.LogInfo("Loo komponendid: Project root: " & projectRoot)
    End If
    
    If storedData.Count > 0 Then
        MakeComponentsLib.ApplyStoredDataToBodies(bodies, storedData, masterFolder, vaultRoot, projectRoot)
    End If
    
    ' Load general settings (template, subfolder, project, assembly)
    ' Paths are stored relative to project root and converted to absolute on load
    Dim generalSettings As MakeComponentsLib.GeneralSettings = _
        MakeComponentsLib.LoadGeneralSettings(masterDoc, projectRoot)
    
    ' Get available materials from document's Materials collection
    Dim materials As List(Of String) = MakeComponentsLib.GetAvailableMaterials(masterDoc)
    UtilsLib.LogInfo("Loo komponendid: Materials available for selection: " & materials.Count)
    
    ' Note: Vault numbering scheme fetching removed - Vault handles numbering on save
    
    ' Get templates
    Dim templates As New List(Of String)
    templates.Add("Part.ipt")
    templates.Add("Sheet Metal.ipt")
    templates.Add("SheetMetal Part.ipt")
    
    ' Extract default project name and master path
    Dim masterPath As String = System.IO.Path.GetDirectoryName(masterDoc.FullDocumentName)
    Dim defaultProject As String = MakeComponentsLib.ExtractProjectName(masterDoc.FullDocumentName)
    If String.IsNullOrEmpty(defaultProject) Then
        defaultProject = System.IO.Path.GetFileName(masterPath)
    End If
    
    ' Dialog loop for face picking
    Dim dialogResult As DialogResult
    
    ' Apply loaded settings as defaults, or use computed defaults
    Dim selectedTemplate As String = If(Not String.IsNullOrEmpty(generalSettings.Template), _
                                        generalSettings.Template, "Part.ipt")
    
    ' Handle subfolder - LoadGeneralSettings already converts relative paths to absolute
    Dim selectedSubfolder As String
    If String.IsNullOrEmpty(generalSettings.Subfolder) Then
        ' Default to Detailid subfolder under project root
        selectedSubfolder = System.IO.Path.Combine(projectRoot, "Detailid")
    Else
        ' Use the loaded subfolder (already converted to absolute path)
        selectedSubfolder = generalSettings.Subfolder
    End If
    
    Dim projectName As String = If(Not String.IsNullOrEmpty(generalSettings.ProjectName), _
                                   generalSettings.ProjectName, defaultProject)
    
    ' If assembly was previously created and still exists, default to UPDATE
    Dim assemblyAction As String = generalSettings.AssemblyAction
    Dim assemblyPath As String = generalSettings.AssemblyPath
    If Not String.IsNullOrEmpty(assemblyPath) AndAlso System.IO.File.Exists(assemblyPath) Then
        If assemblyAction = "CREATE" OrElse assemblyAction = "NONE" Then
            assemblyAction = "UPDATE"
            UtilsLib.LogInfo("Loo komponendid: Found existing assembly, defaulting to UPDATE")
        End If
    End If
    
    Dim pickBodyIndex As Integer = -1
    
    Do
        dialogResult = ShowMainDialog(app, masterDoc, bodies, materials, vaultConnected, templates, masterPath, _
                                      workspaceRoot, selectedTemplate, selectedSubfolder, _
                                      projectName, assemblyAction, assemblyPath, pickBodyIndex)
        
        If dialogResult = DialogResult.Retry AndAlso pickBodyIndex >= 0 AndAlso pickBodyIndex < bodies.Count Then
            ' User clicked "Vali pind" - do face pick
            Dim bi As MakeComponentsLib.BodyInfo = bodies(pickBodyIndex)
            UtilsLib.LogInfo("Loo komponendid: Picking face for '" & bi.Name & "'")
            
            Try
                Dim pickedFace As Object = app.CommandManager.Pick( _
                    SelectionFilterEnum.kPartFacePlanarFilter, _
                    "Vali paksuse pind kehale '" & bi.Name & "' - ESC tühistamiseks")
                
                If pickedFace IsNot Nothing AndAlso TypeOf pickedFace Is Face Then
                    Dim face As Face = CType(pickedFace, Face)
                    RecalculateAxesFromFace(bi, face)
                End If
            Catch
                ' User cancelled pick
            End Try
            
            pickBodyIndex = -1
            Continue Do
        End If
        
        Exit Do
    Loop
    
    If dialogResult <> DialogResult.OK Then
        UtilsLib.LogInfo("Loo komponendid: Cancelled by user")
        Exit Sub
    End If
    
    ' Filter selected bodies
    Dim selectedBodies As New List(Of MakeComponentsLib.BodyInfo)
    For Each bi As MakeComponentsLib.BodyInfo In bodies
        If bi.Selected Then selectedBodies.Add(bi)
    Next
    
    ' Count linked bodies (existing parts that user linked to bodies)
    Dim linkedCount As Integer = 0
    For Each bi As MakeComponentsLib.BodyInfo In bodies
        If bi.PartExists Then linkedCount += 1
    Next
    
    If selectedBodies.Count = 0 Then
        ' No bodies selected for creation, but may have linked files - save and exit
        If linkedCount > 0 Then
            ' Save body links to master document (paths stored relative to project root)
            MakeComponentsLib.SaveBodyDataToMaster(masterDoc, bodies, projectRoot)
            MakeComponentsLib.SaveGeneralSettings(masterDoc, New MakeComponentsLib.GeneralSettings() With {
                .ProjectName = projectName,
                .Template = selectedTemplate,
                .Subfolder = selectedSubfolder,
                .AssemblyAction = assemblyAction,
                .AssemblyPath = assemblyPath
            }, projectRoot)
            Try
                masterDoc.Save()
                UtilsLib.LogInfo("Loo komponendid: Saved " & linkedCount & " body link(s) to master")
            Catch ex As Exception
                UtilsLib.LogWarn("Loo komponendid: Could not save master: " & ex.Message)
            End Try
        Else
            UtilsLib.LogWarn("Loo komponendid: No bodies selected")
        End If
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Loo komponendid: Processing " & selectedBodies.Count & " body(ies)")
    
    ' Use body names as filenames - Vault will assign part numbers on save
    ' Note: Vault numbering scheme selection is shown in dialog but actual numbering
    ' happens when Vault's "Add to Vault" dialog appears during save
    Dim fileNumbers As New List(Of String)
    For Each bi As MakeComponentsLib.BodyInfo In selectedBodies
        ' Sanitize body name for use as filename
        Dim safeName As String = SanitizeFileName(bi.Name)
        fileNumbers.Add(safeName)
    Next
    
    ' Prepare output folder - selectedSubfolder is now a full path
    ' User creates folder on disk, we ensure it exists in Vault
    Dim outputFolder As String = selectedSubfolder
    
    If Not System.IO.Directory.Exists(outputFolder) Then
        UtilsLib.LogError("Loo komponendid: Output folder does not exist: " & outputFolder)
        MessageBox.Show("Väljundkausta ei leitud: " & vbCrLf & outputFolder & vbCrLf & vbCrLf & _
                        "Loo kaust enne jätkamist.", "Loo komponendid")
        Exit Sub
    End If
    
    ' Ensure folder exists in Vault (if connected)
    MakeComponentsLib.EnsureFolderInVault(outputFolder, vaultConn, workspaceRoot)
    UtilsLib.LogInfo("Loo komponendid: Output folder: " & outputFolder)
    
    ' Find template path
    Dim templatePath As String = MakeComponentsLib.FindTemplate(app, selectedTemplate)
    
    ' Create assembly if needed
    Dim asmDoc As AssemblyDocument = Nothing
    If assemblyAction = "CREATE" Then
        asmDoc = MakeComponentsLib.CreateAssembly(app, "")
    ElseIf assemblyAction = "UPDATE" AndAlso Not String.IsNullOrEmpty(assemblyPath) Then
        Try
            asmDoc = CType(app.Documents.Open(assemblyPath, True), AssemblyDocument)
            UtilsLib.LogInfo("Loo komponendid: Opened assembly for update: " & assemblyPath)
        Catch ex As Exception
            UtilsLib.LogError("Loo komponendid: Could not open assembly: " & ex.Message)
        End Try
    End If
    
    ' Process each body
    Dim createdParts As New List(Of String)
    
    For i As Integer = 0 To selectedBodies.Count - 1
        Dim bi As MakeComponentsLib.BodyInfo = selectedBodies(i)
        Dim filePath As String
        Dim newPart As PartDocument = Nothing
        Dim isRecreate As Boolean = False
        
        If bi.PartExists AndAlso System.IO.File.Exists(bi.CreatedPartPath) Then
            ' RECREATE existing part - user has opted to overwrite
            isRecreate = True
            filePath = bi.CreatedPartPath
            UtilsLib.LogInfo("Loo komponendid: Recreating '" & bi.Name & "' - " & System.IO.Path.GetFileName(filePath))
            
            ' Close the file if it's open and delete it
            Try
                For Each doc As Document In app.Documents
                    If doc.FullDocumentName.Equals(filePath, StringComparison.OrdinalIgnoreCase) Then
                        doc.Close(True)
                        Exit For
                    End If
                Next
                System.IO.File.Delete(filePath)
                UtilsLib.LogInfo("Loo komponendid: Deleted old file")
            Catch ex As Exception
                UtilsLib.LogError("Loo komponendid: Could not delete old file: " & ex.Message)
                Continue For
            End Try
        Else
            ' CREATE new part
            Dim fileNumber As String = fileNumbers(i)
            Dim fileName As String = fileNumber & ".ipt"
            filePath = System.IO.Path.Combine(outputFolder, fileName)
            
            UtilsLib.LogInfo("Loo komponendid: Creating '" & bi.Name & "' as " & fileName)
        End If
        
        ' Create new part from template (both for new and recreate)
        newPart = MakeComponentsLib.CreatePartFromTemplate(app, templatePath)
        
        If newPart Is Nothing Then
            UtilsLib.LogError("Loo komponendid: Failed to create part for '" & bi.Name & "'")
            Continue For
        End If
        
        ' Derive body
        If Not MakeComponentsLib.DeriveBodyAsNewPart(masterDoc, bi.Name, newPart) Then
            UtilsLib.LogError("Loo komponendid: Failed to derive body '" & bi.Name & "'")
            newPart.Close(True)
            Continue For
        End If
        
        ' Set iProperties
        MakeComponentsLib.SetPartProperties(newPart, projectName, bi.Name, "")
        
        ' Assign material
        If Not String.IsNullOrEmpty(bi.MaterialName) Then
            MakeComponentsLib.AssignMaterial(newPart, bi.MaterialName)
        End If
        
        ' Convert to sheet metal or set dimensions and register update handler
        If bi.ConvertToSheetMetal Then
            If SheetMetalLib.ConvertToSheetMetal(newPart, bi.ThicknessVector, bi.ThicknessValue) Then
                UtilsLib.LogInfo("Loo komponendid: Converted '" & bi.Name & "' to sheet metal")
                ' Register dimension handler for sheet metal (empty axes = use flat pattern formulas)
                DimensionUpdateLib.RegisterDimensionHandler(newPart, iLogicVb.Automation, "", "", "")
            Else
                ' Sheet metal conversion failed - set dimension properties and register update handler
                MakeComponentsLib.SetDimensionProperties(newPart, bi.ThicknessValue, bi.WidthValue, bi.LengthValue)
                DimensionUpdateLib.RegisterDimensionHandler(newPart, iLogicVb.Automation, bi.ThicknessVector, bi.WidthVector, bi.LengthVector)
                UtilsLib.LogInfo("Loo komponendid: Registered dimension handler for '" & bi.Name & "'")
            End If
        Else
            MakeComponentsLib.SetDimensionProperties(newPart, bi.ThicknessValue, bi.WidthValue, bi.LengthValue)
            DimensionUpdateLib.RegisterDimensionHandler(newPart, iLogicVb.Automation, bi.ThicknessVector, bi.WidthVector, bi.LengthVector)
            UtilsLib.LogInfo("Loo komponendid: Registered dimension handler for '" & bi.Name & "'")
        End If
        
        ' Save part (always SaveAs - either new file or recreated file)
        Try
            newPart.SaveAs(filePath, False)
            
            ' Read actual path after save (Vault may have renamed the file)
            Dim actualPath As String = newPart.FullDocumentName
            UtilsLib.LogInfo("Loo komponendid: Saved " & actualPath)
            
            ' Only add to createdParts if this is a new part (not recreate)
            ' Recreated parts are already in the assembly
            If Not isRecreate Then
                createdParts.Add(actualPath)
            End If
            
            ' Update body info with actual created part path
            bi.CreatedPartPath = actualPath
            bi.PartExists = True
        Catch ex As Exception
            UtilsLib.LogError("Loo komponendid: Failed to save: " & ex.Message)
        End Try
        
        ' Close part
        newPart.Close(False)
    Next
    
    ' Save body data to master document (settings and part references)
    ' Save ALL bodies, not just selected ones, to preserve unprocessed body data
    ' Paths are stored relative to project root for portability
    MakeComponentsLib.SaveBodyDataToMaster(masterDoc, bodies, projectRoot)
    
    ' Save master document to persist the stored body data
    If bodies.Count > 0 Then
        Try
            masterDoc.Save()
            UtilsLib.LogInfo("Loo komponendid: Master document saved with body references")
        Catch ex As Exception
            UtilsLib.LogWarn("Loo komponendid: Could not save master document: " & ex.Message)
        End Try
    End If
    
    ' Place in assembly and save
    Dim actualAssemblyPath As String = ""
    If asmDoc IsNot Nothing AndAlso createdParts.Count > 0 Then
        For Each partPath As String In createdParts
            Dim occ As ComponentOccurrence = MakeComponentsLib.PlaceComponentGrounded(asmDoc, partPath)
            If occ IsNot Nothing Then
                OccurrenceNamingLib.RenameOccurrence(occ)
            End If
        Next
        
        ' Set assembly properties BEFORE save (Vault reads VDS properties at SaveAs time)
        MakeComponentsLib.SetAssemblyProperties(asmDoc, projectName)
        
        ' Save assembly
        If assemblyAction = "CREATE" AndAlso Not String.IsNullOrEmpty(assemblyPath) Then
            Try
                asmDoc.SaveAs(assemblyPath, False)
                actualAssemblyPath = asmDoc.FullDocumentName
                UtilsLib.LogInfo("Loo komponendid: Saved assembly: " & actualAssemblyPath)
            Catch ex As Exception
                UtilsLib.LogError("Loo komponendid: Failed to save assembly: " & ex.Message)
            End Try
        ElseIf assemblyAction = "CREATE" Then
            Dim asmFileName As String = defaultProject & "_asm.iam"
            Dim asmFilePath As String = System.IO.Path.Combine(outputFolder, asmFileName)
            Try
                asmDoc.SaveAs(asmFilePath, False)
                actualAssemblyPath = asmDoc.FullDocumentName
                UtilsLib.LogInfo("Loo komponendid: Saved assembly: " & actualAssemblyPath)
            Catch ex As Exception
                UtilsLib.LogError("Loo komponendid: Failed to save assembly: " & ex.Message)
            End Try
        Else
            asmDoc.Save()
            actualAssemblyPath = asmDoc.FullDocumentName
        End If
    End If
    
    ' Save general settings to master document
    Dim settingsToSave As New MakeComponentsLib.GeneralSettings()
    settingsToSave.ProjectName = projectName
    settingsToSave.Template = selectedTemplate
    settingsToSave.Subfolder = selectedSubfolder
    settingsToSave.AssemblyAction = assemblyAction
    settingsToSave.AssemblyPath = If(Not String.IsNullOrEmpty(actualAssemblyPath), actualAssemblyPath, assemblyPath)
    
    MakeComponentsLib.SaveGeneralSettings(masterDoc, settingsToSave, projectRoot)
    
    ' Save master document again to persist general settings
    Try
        masterDoc.Save()
        UtilsLib.LogInfo("Loo komponendid: Master document saved with general settings")
    Catch ex As Exception
        UtilsLib.LogWarn("Loo komponendid: Could not save master document: " & ex.Message)
    End Try
    
    Dim recreatedCount As Integer = selectedBodies.Count - createdParts.Count
    If recreatedCount > 0 Then
        UtilsLib.LogInfo("Loo komponendid: Completed - " & createdParts.Count & " new part(s), " & recreatedCount & " recreated")
    Else
        UtilsLib.LogInfo("Loo komponendid: Completed - created " & createdParts.Count & " part(s)")
    End If
End Sub

' Recalculate axes based on a picked face
Sub RecalculateAxesFromFace(bi As MakeComponentsLib.BodyInfo, face As Face)
    Try
        Dim geom As Object = face.Geometry
        If TypeOf geom Is Plane Then
            Dim plane As Plane = CType(geom, Plane)
            Dim normal As UnitVector = plane.Normal
            Dim nx As Double = normal.X
            Dim ny As Double = normal.Y
            Dim nz As Double = normal.Z
            
            ' Recalculate thickness along this normal
            bi.ThicknessValue = MakeComponentsLib.GetOrientedExtentForBody(bi.Body, nx, ny, nz)
            bi.ThicknessVector = MakeComponentsLib.VectorToString(nx, ny, nz)
            
            ' Compute perpendicular axes for width/length
            Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
            Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
            
            ' Simple perpendicular calculation
            Dim refX As Double = 1, refY As Double = 0, refZ As Double = 0
            Dim dot As Double = nx * refX + ny * refY + nz * refZ
            If Math.Abs(dot) > 0.9 Then
                refX = 0 : refY = 1 : refZ = 0
            End If
            
            wx = ny * refZ - nz * refY
            wy = nz * refX - nx * refZ
            wz = nx * refY - ny * refX
            Dim wLen As Double = Math.Sqrt(wx * wx + wy * wy + wz * wz)
            If wLen > 0.0001 Then
                wx /= wLen : wy /= wLen : wz /= wLen
            End If
            
            lx = ny * wz - nz * wy
            ly = nz * wx - nx * wz
            lz = nx * wy - ny * wx
            Dim lLen As Double = Math.Sqrt(lx * lx + ly * ly + lz * lz)
            If lLen > 0.0001 Then
                lx /= lLen : ly /= lLen : lz /= lLen
            End If
            
            Dim widthExtent As Double = MakeComponentsLib.GetOrientedExtentForBody(bi.Body, wx, wy, wz)
            Dim lengthExtent As Double = MakeComponentsLib.GetOrientedExtentForBody(bi.Body, lx, ly, lz)
            
            If lengthExtent >= widthExtent Then
                bi.WidthValue = widthExtent
                bi.LengthValue = lengthExtent
                bi.WidthVector = MakeComponentsLib.VectorToString(wx, wy, wz)
                bi.LengthVector = MakeComponentsLib.VectorToString(lx, ly, lz)
            Else
                bi.WidthValue = lengthExtent
                bi.LengthValue = widthExtent
                bi.WidthVector = MakeComponentsLib.VectorToString(lx, ly, lz)
                bi.LengthVector = MakeComponentsLib.VectorToString(wx, wy, wz)
            End If
            
            UtilsLib.LogInfo("Loo komponendid: Recalculated axes for '" & bi.Name & "' - T:" & _
                     FormatNumber(bi.ThicknessValue * 10, 2) & " W:" & FormatNumber(bi.WidthValue * 10, 2) & _
                     " L:" & FormatNumber(bi.LengthValue * 10, 2))
        End If
    Catch ex As Exception
        UtilsLib.LogError("Loo komponendid: Error recalculating axes: " & ex.Message)
    End Try
End Sub

' ============================================================================
' Main Dialog with DataGridView
' ============================================================================

Function ShowMainDialog(app As Inventor.Application, _
                        masterDoc As PartDocument, _
                        bodies As List(Of MakeComponentsLib.BodyInfo), _
                        materials As List(Of String), _
                        vaultConnected As Boolean, _
                        templates As List(Of String), _
                        masterPath As String, _
                        workspaceRoot As String, _
                        ByRef selectedTemplate As String, _
                        ByRef selectedSubfolder As String, _
                        ByRef projectName As String, _
                        ByRef assemblyAction As String, _
                        ByRef assemblyPath As String, _
                        ByRef pickBodyIndex As Integer) As DialogResult
    
    pickBodyIndex = -1
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Loo komponendid"
    frm.Width = 950
    frm.Height = 650
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.Sizable
    frm.MinimizeBox = True
    frm.MaximizeBox = True
    
    Dim currentY As Integer = 10
    
    ' Project name
    Dim lblProject As New System.Windows.Forms.Label()
    lblProject.Text = "Projekt:"
    lblProject.Left = 10
    lblProject.Top = currentY + 3
    lblProject.Width = 80
    frm.Controls.Add(lblProject)
    
    Dim txtProject As New System.Windows.Forms.TextBox()
    txtProject.Name = "txtProject"
    txtProject.Text = projectName
    txtProject.Left = 90
    txtProject.Top = currentY
    txtProject.Width = 200
    frm.Controls.Add(txtProject)
    
    currentY += 30
    
    ' Template
    Dim lblTemplate As New System.Windows.Forms.Label()
    lblTemplate.Text = "Šabloon:"
    lblTemplate.Left = 10
    lblTemplate.Top = currentY + 3
    lblTemplate.Width = 80
    frm.Controls.Add(lblTemplate)
    
    Dim cboTemplate As New System.Windows.Forms.ComboBox()
    cboTemplate.Name = "cboTemplate"
    cboTemplate.Left = 90
    cboTemplate.Top = currentY
    cboTemplate.Width = 200
    cboTemplate.DropDownStyle = ComboBoxStyle.DropDownList
    For Each t As String In templates
        cboTemplate.Items.Add(t)
    Next
    cboTemplate.SelectedIndex = 0
    frm.Controls.Add(cboTemplate)
    
    ' Subfolder (output folder)
    Dim lblSubfolder As New System.Windows.Forms.Label()
    lblSubfolder.Text = "Väljundkaust:"
    lblSubfolder.Left = 310
    lblSubfolder.Top = currentY + 3
    lblSubfolder.Width = 80
    frm.Controls.Add(lblSubfolder)
    
    Dim txtSubfolder As New System.Windows.Forms.TextBox()
    txtSubfolder.Name = "txtSubfolder"
    txtSubfolder.Text = selectedSubfolder
    txtSubfolder.Left = 395
    txtSubfolder.Top = currentY
    txtSubfolder.Width = 470
    txtSubfolder.ReadOnly = True
    frm.Controls.Add(txtSubfolder)
    
    Dim btnBrowseFolder As New System.Windows.Forms.Button()
    btnBrowseFolder.Name = "btnBrowseFolder"
    btnBrowseFolder.Text = "..."
    btnBrowseFolder.Left = 870
    btnBrowseFolder.Top = currentY
    btnBrowseFolder.Width = 30
    btnBrowseFolder.Height = 23
    frm.Controls.Add(btnBrowseFolder)
    
    AddHandler btnBrowseFolder.Click, Sub(s, e)
        Dim fbd As New FolderBrowserDialog()
        fbd.Description = "Vali väljundkaust komponentidele"
        fbd.ShowNewFolderButton = True
        If Not String.IsNullOrEmpty(txtSubfolder.Text) AndAlso System.IO.Directory.Exists(txtSubfolder.Text) Then
            fbd.SelectedPath = txtSubfolder.Text
        Else
            fbd.SelectedPath = masterPath
        End If
        If fbd.ShowDialog() = DialogResult.OK Then
            txtSubfolder.Text = fbd.SelectedPath
        End If
    End Sub
    
    currentY += 30
    
    ' Assembly options
    Dim lblAssembly As New System.Windows.Forms.Label()
    lblAssembly.Text = "Koost:"
    lblAssembly.Left = 10
    lblAssembly.Top = currentY + 3
    lblAssembly.Width = 80
    frm.Controls.Add(lblAssembly)
    
    Dim cboAssembly As New System.Windows.Forms.ComboBox()
    cboAssembly.Name = "cboAssembly"
    cboAssembly.Left = 90
    cboAssembly.Top = currentY
    cboAssembly.Width = 200
    cboAssembly.DropDownStyle = ComboBoxStyle.DropDownList
    cboAssembly.Items.Add("Ära loo koostu")
    cboAssembly.Items.Add("Loo uus koost")
    cboAssembly.Items.Add("Uuenda olemasolevat")
    ' Set default based on assemblyAction
    Select Case assemblyAction
        Case "UPDATE" : cboAssembly.SelectedIndex = 2
        Case "CREATE" : cboAssembly.SelectedIndex = 1
        Case Else : cboAssembly.SelectedIndex = 0
    End Select
    frm.Controls.Add(cboAssembly)
    
    Dim btnBrowseAsm As New System.Windows.Forms.Button()
    btnBrowseAsm.Name = "btnBrowseAsm"
    btnBrowseAsm.Text = "..."
    btnBrowseAsm.Left = 295
    btnBrowseAsm.Top = currentY
    btnBrowseAsm.Width = 30
    btnBrowseAsm.Height = 23
    btnBrowseAsm.Enabled = (assemblyAction = "CREATE" OrElse assemblyAction = "UPDATE")
    frm.Controls.Add(btnBrowseAsm)
    
    Dim txtAsmPath As New System.Windows.Forms.TextBox()
    txtAsmPath.Name = "txtAsmPath"
    txtAsmPath.Text = assemblyPath
    txtAsmPath.Left = 330
    txtAsmPath.Top = currentY
    txtAsmPath.Width = 265
    txtAsmPath.ReadOnly = True
    txtAsmPath.Enabled = (assemblyAction = "CREATE" OrElse assemblyAction = "UPDATE")
    frm.Controls.Add(txtAsmPath)
    
    AddHandler cboAssembly.SelectedIndexChanged, Sub(s, e)
        Dim enableBrowse As Boolean = (cboAssembly.SelectedIndex = 1 OrElse cboAssembly.SelectedIndex = 2)
        btnBrowseAsm.Enabled = enableBrowse
        txtAsmPath.Enabled = enableBrowse
    End Sub
    
    AddHandler btnBrowseAsm.Click, Sub(s, e)
        If cboAssembly.SelectedIndex = 1 Then
            ' CREATE - use FolderBrowserDialog
            Dim fbd As New FolderBrowserDialog()
            fbd.Description = "Vali kaust uuele koostule"
            If String.IsNullOrEmpty(txtAsmPath.Text) Then
                fbd.SelectedPath = masterPath
            Else
                fbd.SelectedPath = System.IO.Path.GetDirectoryName(txtAsmPath.Text)
            End If
            If fbd.ShowDialog() = DialogResult.OK Then
                ' Use txtProject.Text instead of ByRef projectName
                Dim asmName As String = txtProject.Text & "_asm.iam"
                txtAsmPath.Text = System.IO.Path.Combine(fbd.SelectedPath, asmName)
            End If
        Else
            ' UPDATE - use OpenFileDialog
            Dim ofd As New OpenFileDialog()
            ofd.Filter = "Inventor Assembly|*.iam"
            ofd.Title = "Vali koost"
            If String.IsNullOrEmpty(txtAsmPath.Text) Then
                ofd.InitialDirectory = masterPath
            Else
                ofd.InitialDirectory = System.IO.Path.GetDirectoryName(txtAsmPath.Text)
            End If
            If ofd.ShowDialog() = DialogResult.OK Then
                txtAsmPath.Text = ofd.FileName
            End If
        End If
    End Sub
    
    currentY += 35
    
    ' Bodies table header
    Dim lblBodies As New System.Windows.Forms.Label()
    lblBodies.Text = "Kehad:"
    lblBodies.Left = 10
    lblBodies.Top = currentY
    lblBodies.Width = 80
    frm.Controls.Add(lblBodies)
    
    currentY += 20
    
    ' DataGridView for bodies
    Dim dgv As New System.Windows.Forms.DataGridView()
    dgv.Name = "dgvBodies"
    dgv.Left = 10
    dgv.Top = currentY
    dgv.Width = 910
    dgv.Height = 350
    dgv.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    dgv.AllowUserToAddRows = False
    dgv.AllowUserToDeleteRows = False
    dgv.RowHeadersVisible = False
    dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    dgv.MultiSelect = False
    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    
    ' Column: Selected (checkbox)
    Dim colSelected As New DataGridViewCheckBoxColumn()
    colSelected.Name = "colSelected"
    colSelected.HeaderText = "Vali"
    colSelected.Width = 40
    dgv.Columns.Add(colSelected)
    
    ' Column: Name (read-only)
    Dim colName As New DataGridViewTextBoxColumn()
    colName.Name = "colName"
    colName.HeaderText = "Nimi"
    colName.Width = 150
    colName.ReadOnly = True
    dgv.Columns.Add(colName)
    
    ' Column: Status (shows if part exists)
    Dim colStatus As New DataGridViewTextBoxColumn()
    colStatus.Name = "colStatus"
    colStatus.HeaderText = "Olek"
    colStatus.Width = 100
    colStatus.ReadOnly = True
    dgv.Columns.Add(colStatus)
    
    ' Column: Link file button (for linking/unlinking existing parts)
    Dim colLink As New DataGridViewButtonColumn()
    colLink.Name = "colLink"
    colLink.HeaderText = ""
    colLink.Width = 90
    dgv.Columns.Add(colLink)
    
    ' Column: Thickness (read-only)
    Dim colT As New DataGridViewTextBoxColumn()
    colT.Name = "colT"
    colT.HeaderText = "T (mm)"
    colT.Width = 60
    colT.ReadOnly = True
    dgv.Columns.Add(colT)
    
    ' Column: Width (read-only)
    Dim colW As New DataGridViewTextBoxColumn()
    colW.Name = "colW"
    colW.HeaderText = "W (mm)"
    colW.Width = 60
    colW.ReadOnly = True
    dgv.Columns.Add(colW)
    
    ' Column: Length (read-only)
    Dim colL As New DataGridViewTextBoxColumn()
    colL.Name = "colL"
    colL.HeaderText = "L (mm)"
    colL.Width = 60
    colL.ReadOnly = True
    dgv.Columns.Add(colL)
    
    ' Column: Lehtmetall (checkbox)
    Dim colSM As New DataGridViewCheckBoxColumn()
    colSM.Name = "colSM"
    colSM.HeaderText = "Lehtmetall"
    colSM.Width = 80
    dgv.Columns.Add(colSM)
    
    ' Column: Material (combobox)
    Dim colMat As New DataGridViewComboBoxColumn()
    colMat.Name = "colMat"
    colMat.HeaderText = "Materjal"
    colMat.Width = 180
    colMat.Items.Add("")
    For Each mat As String In materials
        colMat.Items.Add(mat)
    Next
    dgv.Columns.Add(colMat)
    
    ' Column: Pick face button
    Dim colPick As New DataGridViewButtonColumn()
    colPick.Name = "colPick"
    colPick.HeaderText = "Teljed"
    colPick.Text = "Vali pind"
    colPick.UseColumnTextForButtonValue = True
    colPick.Width = 80
    dgv.Columns.Add(colPick)
    
    ' Populate rows
    For i As Integer = 0 To bodies.Count - 1
        Dim bi As MakeComponentsLib.BodyInfo = bodies(i)
        Dim rowIndex As Integer = dgv.Rows.Add()
        dgv.Rows(rowIndex).Tag = i
        dgv.Rows(rowIndex).Cells("colSelected").Value = bi.Selected
        dgv.Rows(rowIndex).Cells("colName").Value = bi.Name
        
        ' Show status and link button - new or existing part
        If bi.PartExists Then
            Dim partName As String = System.IO.Path.GetFileName(bi.CreatedPartPath)
            dgv.Rows(rowIndex).Cells("colStatus").Value = "* " & partName
            dgv.Rows(rowIndex).Cells("colLink").Value = "Eemalda seos"
        Else
            dgv.Rows(rowIndex).Cells("colStatus").Value = "(uus)"
            dgv.Rows(rowIndex).Cells("colLink").Value = "Seo fail..."
        End If
        
        dgv.Rows(rowIndex).Cells("colT").Value = FormatNumber(bi.ThicknessValue * 10, 2)
        dgv.Rows(rowIndex).Cells("colW").Value = FormatNumber(bi.WidthValue * 10, 2)
        dgv.Rows(rowIndex).Cells("colL").Value = FormatNumber(bi.LengthValue * 10, 2)
        dgv.Rows(rowIndex).Cells("colSM").Value = bi.ConvertToSheetMetal
        dgv.Rows(rowIndex).Cells("colMat").Value = GetValidatedMaterial(bi.MaterialName, materials)
    Next
    
    ' Store pick index in form Tag (can't use ByRef in lambda)
    frm.Tag = -1
    
    ' Handle button clicks for face picking and file linking
    AddHandler dgv.CellContentClick, Sub(s, e)
        If e.RowIndex < 0 Then Exit Sub
        
        Dim idx As Integer = CInt(dgv.Rows(e.RowIndex).Tag)
        
        ' Handle "Vali pind" (pick face) button
        If e.ColumnIndex = dgv.Columns("colPick").Index Then
            ' Save current state before closing
            SyncGridToBodyInfo(dgv, bodies)
            ' Store index in form.Tag instead of ByRef parameter
            frm.Tag = idx
            frm.DialogResult = DialogResult.Retry
            frm.Close()
        End If
        
        ' Handle "Seo fail..." / "Eemalda seos" (link/unlink) button
        If e.ColumnIndex = dgv.Columns("colLink").Index Then
            Dim bi As MakeComponentsLib.BodyInfo = bodies(idx)
            
            If bi.PartExists Then
                ' Unlink: Clear the association
                bi.CreatedPartPath = ""
                bi.PartExists = False
                bi.Selected = True  ' Now available for creation
                dgv.Rows(e.RowIndex).Cells("colStatus").Value = "(uus)"
                dgv.Rows(e.RowIndex).Cells("colLink").Value = "Seo fail..."
                dgv.Rows(e.RowIndex).Cells("colSelected").Value = True
            Else
                ' Link: Try automatic search first to pre-fill file dialog
                Dim autoFoundPath As String = MakeComponentsLib.FindPartByDescription( _
                    app, bi.Name, masterPath, workspaceRoot)
                
                Dim ofd As New OpenFileDialog()
                ofd.Filter = "Inventor Part|*.ipt"
                ofd.Title = "Vali olemasolev detail"
                
                If Not String.IsNullOrEmpty(autoFoundPath) Then
                    ' Found a match - open dialog at that location with file pre-selected
                    ofd.InitialDirectory = System.IO.Path.GetDirectoryName(autoFoundPath)
                    ofd.FileName = System.IO.Path.GetFileName(autoFoundPath)
                Else
                    ' No match found - open at master document location
                    ofd.InitialDirectory = masterPath
                End If
                
                If ofd.ShowDialog() = DialogResult.OK Then
                    ' Read properties from the selected file
                    MakeComponentsLib.ReadPropertiesFromPart(app, ofd.FileName, bi)
                    
                    bi.CreatedPartPath = ofd.FileName
                    bi.PartExists = True
                    bi.Selected = False  ' Don't recreate linked parts
                    
                    ' Update grid cells with imported properties
                    Dim partName As String = System.IO.Path.GetFileName(ofd.FileName)
                    dgv.Rows(e.RowIndex).Cells("colStatus").Value = "* " & partName
                    dgv.Rows(e.RowIndex).Cells("colLink").Value = "Eemalda seos"
                    dgv.Rows(e.RowIndex).Cells("colSelected").Value = False
                    dgv.Rows(e.RowIndex).Cells("colSM").Value = bi.ConvertToSheetMetal
                    dgv.Rows(e.RowIndex).Cells("colMat").Value = GetValidatedMaterial(bi.MaterialName, materials)
                    dgv.Rows(e.RowIndex).Cells("colT").Value = FormatNumber(bi.ThicknessValue * 10, 2)
                    dgv.Rows(e.RowIndex).Cells("colW").Value = FormatNumber(bi.WidthValue * 10, 2)
                    dgv.Rows(e.RowIndex).Cells("colL").Value = FormatNumber(bi.LengthValue * 10, 2)
                End If
            End If
        End If
    End Sub
    
    frm.Controls.Add(dgv)
    
    ' Material apply to all controls
    Dim lblApplyMat As New System.Windows.Forms.Label()
    lblApplyMat.Text = "Materjal kõigile:"
    lblApplyMat.Left = 10
    lblApplyMat.Top = 573
    lblApplyMat.Width = 100
    lblApplyMat.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(lblApplyMat)
    
    Dim cboApplyMat As New System.Windows.Forms.ComboBox()
    cboApplyMat.Left = 115
    cboApplyMat.Top = 570
    cboApplyMat.Width = 200
    cboApplyMat.DropDownStyle = ComboBoxStyle.DropDownList
    cboApplyMat.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    cboApplyMat.Items.Add("(vali materjal)")
    For Each mat As String In materials
        cboApplyMat.Items.Add(mat)
    Next
    cboApplyMat.SelectedIndex = 0
    frm.Controls.Add(cboApplyMat)
    
    Dim btnApplyMatAll As New System.Windows.Forms.Button()
    btnApplyMatAll.Text = "Rakenda kõigile"
    btnApplyMatAll.Left = 320
    btnApplyMatAll.Top = 570
    btnApplyMatAll.Width = 110
    btnApplyMatAll.Height = 25
    btnApplyMatAll.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(btnApplyMatAll)
    
    AddHandler btnApplyMatAll.Click, Sub(s, e)
        If cboApplyMat.SelectedIndex > 0 Then
            Dim matName As String = cboApplyMat.SelectedItem.ToString()
            For Each row As DataGridViewRow In dgv.Rows
                row.Cells("colMat").Value = matName
            Next
        End If
    End Sub
    
    ' Lehtmetall apply to all
    Dim btnSMAll As New System.Windows.Forms.Button()
    btnSMAll.Text = "Lehtmetall kõigile"
    btnSMAll.Left = 440
    btnSMAll.Top = 570
    btnSMAll.Width = 120
    btnSMAll.Height = 25
    btnSMAll.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(btnSMAll)
    
    AddHandler btnSMAll.Click, Sub(s, e)
        For Each row As DataGridViewRow In dgv.Rows
            row.Cells("colSM").Value = True
        Next
    End Sub
    
    ' OK/Cancel buttons (wider OK button)
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Loo komponendid"
    btnOK.Left = 700
    btnOK.Top = 570
    btnOK.Width = 130
    btnOK.Height = 28
    btnOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 835
    btnCancel.Top = 570
    btnCancel.Width = 85
    btnCancel.Height = 28
    btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Extract values
    If result = DialogResult.OK OrElse result = DialogResult.Retry Then
        projectName = txtProject.Text
        ' selectedScheme not used - Vault handles numbering on save
        selectedTemplate = If(cboTemplate.SelectedItem IsNot Nothing, cboTemplate.SelectedItem.ToString(), "Part.ipt")
        selectedSubfolder = txtSubfolder.Text
        
        Select Case cboAssembly.SelectedIndex
            Case 0 : assemblyAction = "NONE"
            Case 1 : assemblyAction = "CREATE"
            Case 2 : assemblyAction = "UPDATE"
        End Select
        assemblyPath = txtAsmPath.Text
        
        ' Read pick index from form.Tag (stored by lambda handler)
        pickBodyIndex = CInt(frm.Tag)
        
        ' Sync grid to body info
        SyncGridToBodyInfo(dgv, bodies)
    End If
    
    frm.Dispose()
    Return result
End Function

' Sync DataGridView state to BodyInfo list
Sub SyncGridToBodyInfo(dgv As DataGridView, bodies As List(Of MakeComponentsLib.BodyInfo))
    For Each row As DataGridViewRow In dgv.Rows
        Dim idx As Integer = CInt(row.Tag)
        If idx >= 0 AndAlso idx < bodies.Count Then
            bodies(idx).Selected = CBool(row.Cells("colSelected").Value)
            bodies(idx).ConvertToSheetMetal = CBool(row.Cells("colSM").Value)
            Dim matVal As Object = row.Cells("colMat").Value
            bodies(idx).MaterialName = If(matVal IsNot Nothing, matVal.ToString(), "")
        End If
    Next
End Sub

' Get validated material value for ComboBox (returns empty if not in items)
Function GetValidatedMaterial(materialName As String, materials As List(Of String)) As String
    If String.IsNullOrEmpty(materialName) Then Return ""
    If materials.Contains(materialName) Then Return materialName
    Return ""
End Function

' Sanitize a string for use as a filename
Function SanitizeFileName(name As String) As String
    Dim invalid() As Char = System.IO.Path.GetInvalidFileNameChars()
    Dim result As String = name
    For Each c As Char In invalid
        result = result.Replace(c, "_"c)
    Next
    ' Also replace some additional problematic characters
    result = result.Replace(" ", "_")
    Return result
End Function
