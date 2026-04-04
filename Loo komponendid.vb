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

' Libraries come after references
AddVbFile "Lib/VaultNumberingLib.vb"
AddVbFile "Lib/MakeComponentsLib.vb"
AddVbFile "Lib/SheetMetalLib.vb"
AddVbFile "Lib/BoundingBoxStockLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim logs As New List(Of String)
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        Logger.Error("Loo komponendid: No active document")
        MessageBox.Show("Ava esmalt multi-body detail.", "Loo komponendid")
        Exit Sub
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("Loo komponendid: Active document is not a part")
        MessageBox.Show("Aktiivseks dokumendiks peab olema detail (.ipt).", "Loo komponendid")
        Exit Sub
    End If
    
    Dim masterDoc As PartDocument = CType(app.ActiveDocument, PartDocument)
    
    If masterDoc.ComponentDefinition.SurfaceBodies.Count < 1 Then
        Logger.Error("Loo komponendid: No solid bodies in part")
        MessageBox.Show("Detailis puuduvad tahked kehad.", "Loo komponendid")
        Exit Sub
    End If
    
    ' Ensure master is saved
    If String.IsNullOrEmpty(masterDoc.FullDocumentName) Then
        Logger.Error("Loo komponendid: Master document not saved")
        MessageBox.Show("Salvesta esmalt master-detail.", "Loo komponendid")
        Exit Sub
    End If
    
    Logger.Info("Loo komponendid: Starting for " & masterDoc.DisplayName)
    
    ' Get selected bodies from SelectSet (if any)
    Dim selectedBodyNames As New List(Of String)
    For Each obj As Object In masterDoc.SelectSet
        If TypeOf obj Is SurfaceBody Then
            selectedBodyNames.Add(CType(obj, SurfaceBody).Name)
        End If
    Next
    
    If selectedBodyNames.Count > 0 Then
        Logger.Info("Loo komponendid: User selected " & selectedBodyNames.Count & " body(ies)")
    Else
        Logger.Info("Loo komponendid: No selection, using all " & masterDoc.ComponentDefinition.SurfaceBodies.Count & " body(ies)")
    End If
    
    ' Get Vault connection
    Dim vaultConn As Object = VaultNumberingLib.GetVaultConnection()
    Dim vaultConnected As Boolean = (vaultConn IsNot Nothing)
    
    If vaultConnected Then
        Logger.Info("Loo komponendid: Vault connected - " & VaultNumberingLib.GetConnectionInfo(vaultConn))
    Else
        Logger.Warn("Loo komponendid: Vault not connected - manual filenames required")
    End If
    
    ' Get all bodies with detected axes
    Dim allBodies As List(Of MakeComponentsLib.BodyInfo) = MakeComponentsLib.GetBodiesWithAxes(masterDoc, logs)
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    
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
        MakeComponentsLib.LoadBodyDataFromMaster(masterDoc, logs)
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    
    ' Apply stored settings to bodies (matches by name, then by geometry signature)
    If storedData.Count > 0 Then
        MakeComponentsLib.ApplyStoredDataToBodies(bodies, storedData, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
    End If
    
    ' Load general settings (template, subfolder, project, assembly)
    Dim generalSettings As MakeComponentsLib.GeneralSettings = _
        MakeComponentsLib.LoadGeneralSettings(masterDoc, logs)
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    
    ' Get available materials from document's Materials collection
    Dim materials As List(Of String) = MakeComponentsLib.GetAvailableMaterials(masterDoc, logs)
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    Logger.Info("Loo komponendid: Materials available for selection: " & materials.Count)
    
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
    Dim selectedSubfolder As String = If(Not String.IsNullOrEmpty(generalSettings.Subfolder), _
                                         generalSettings.Subfolder, "Detailid")
    Dim projectName As String = If(Not String.IsNullOrEmpty(generalSettings.ProjectName), _
                                   generalSettings.ProjectName, defaultProject)
    
    ' If assembly was previously created and still exists, default to UPDATE
    Dim assemblyAction As String = generalSettings.AssemblyAction
    Dim assemblyPath As String = generalSettings.AssemblyPath
    If Not String.IsNullOrEmpty(assemblyPath) AndAlso System.IO.File.Exists(assemblyPath) Then
        If assemblyAction = "CREATE" OrElse assemblyAction = "NONE" Then
            assemblyAction = "UPDATE"
            Logger.Info("Loo komponendid: Found existing assembly, defaulting to UPDATE")
        End If
    End If
    
    Dim pickBodyIndex As Integer = -1
    
    Do
        dialogResult = ShowMainDialog(app, masterDoc, bodies, materials, vaultConnected, templates, masterPath, _
                                      selectedTemplate, selectedSubfolder, _
                                      projectName, assemblyAction, assemblyPath, pickBodyIndex)
        
        If dialogResult = DialogResult.Retry AndAlso pickBodyIndex >= 0 AndAlso pickBodyIndex < bodies.Count Then
            ' User clicked "Vali pind" - do face pick
            Dim bi As MakeComponentsLib.BodyInfo = bodies(pickBodyIndex)
            Logger.Info("Loo komponendid: Picking face for '" & bi.Name & "'")
            
            Try
                Dim pickedFace As Object = app.CommandManager.Pick( _
                    SelectionFilterEnum.kPartFacePlanarFilter, _
                    "Vali paksuse pind kehale '" & bi.Name & "' - ESC tühistamiseks")
                
                If pickedFace IsNot Nothing AndAlso TypeOf pickedFace Is Face Then
                    Dim face As Face = CType(pickedFace, Face)
                    RecalculateAxesFromFace(bi, face, logs)
                    For Each log As String In logs : Logger.Info(log) : Next
                    logs.Clear()
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
        Logger.Info("Loo komponendid: Cancelled by user")
        Exit Sub
    End If
    
    ' Filter selected bodies
    Dim selectedBodies As New List(Of MakeComponentsLib.BodyInfo)
    For Each bi As MakeComponentsLib.BodyInfo In bodies
        If bi.Selected Then selectedBodies.Add(bi)
    Next
    
    If selectedBodies.Count = 0 Then
        Logger.Warn("Loo komponendid: No bodies selected")
        Exit Sub
    End If
    
    Logger.Info("Loo komponendid: Processing " & selectedBodies.Count & " body(ies)")
    
    ' Use body names as filenames - Vault will assign part numbers on save
    ' Note: Vault numbering scheme selection is shown in dialog but actual numbering
    ' happens when Vault's "Add to Vault" dialog appears during save
    Dim fileNumbers As New List(Of String)
    For Each bi As MakeComponentsLib.BodyInfo In selectedBodies
        ' Sanitize body name for use as filename
        Dim safeName As String = SanitizeFileName(bi.Name)
        fileNumbers.Add(safeName)
    Next
    
    ' Prepare output folder
    Dim outputFolder As String = MakeComponentsLib.EnsureSubfolder(masterPath, selectedSubfolder)
    Logger.Info("Loo komponendid: Output folder: " & outputFolder)
    
    ' Find template path
    Dim templatePath As String = MakeComponentsLib.FindTemplate(app, selectedTemplate, logs)
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    
    ' Create assembly if needed
    Dim asmDoc As AssemblyDocument = Nothing
    If assemblyAction = "CREATE" Then
        asmDoc = MakeComponentsLib.CreateAssembly(app, "", logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
    ElseIf assemblyAction = "UPDATE" AndAlso Not String.IsNullOrEmpty(assemblyPath) Then
        Try
            asmDoc = CType(app.Documents.Open(assemblyPath, True), AssemblyDocument)
            Logger.Info("Loo komponendid: Opened assembly for update: " & assemblyPath)
        Catch ex As Exception
            Logger.Error("Loo komponendid: Could not open assembly: " & ex.Message)
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
            Logger.Info("Loo komponendid: Recreating '" & bi.Name & "' - " & System.IO.Path.GetFileName(filePath))
            
            ' Close the file if it's open and delete it
            Try
                For Each doc As Document In app.Documents
                    If doc.FullDocumentName.Equals(filePath, StringComparison.OrdinalIgnoreCase) Then
                        doc.Close(True)
                        Exit For
                    End If
                Next
                System.IO.File.Delete(filePath)
                Logger.Info("Loo komponendid: Deleted old file")
            Catch ex As Exception
                Logger.Error("Loo komponendid: Could not delete old file: " & ex.Message)
                Continue For
            End Try
        Else
            ' CREATE new part
            Dim fileNumber As String = fileNumbers(i)
            Dim fileName As String = fileNumber & ".ipt"
            filePath = System.IO.Path.Combine(outputFolder, fileName)
            
            Logger.Info("Loo komponendid: Creating '" & bi.Name & "' as " & fileName)
        End If
        
        ' Create new part from template (both for new and recreate)
        newPart = MakeComponentsLib.CreatePartFromTemplate(app, templatePath, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        If newPart Is Nothing Then
            Logger.Error("Loo komponendid: Failed to create part for '" & bi.Name & "'")
            Continue For
        End If
        
        ' Derive body
        If Not MakeComponentsLib.DeriveBodyAsNewPart(masterDoc, bi.Name, newPart, logs) Then
            Logger.Error("Loo komponendid: Failed to derive body '" & bi.Name & "'")
            newPart.Close(True)
            For Each log As String In logs : Logger.Info(log) : Next
            logs.Clear()
            Continue For
        End If
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Set iProperties
        MakeComponentsLib.SetPartProperties(newPart, projectName, bi.Name, "", logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Assign material
        If Not String.IsNullOrEmpty(bi.MaterialName) Then
            MakeComponentsLib.AssignMaterial(newPart, bi.MaterialName, logs)
            For Each log As String In logs : Logger.Info(log) : Next
            logs.Clear()
        End If
        
        ' Convert to sheet metal or set dimensions and create local rule
        If bi.ConvertToSheetMetal Then
            If SheetMetalLib.ConvertToSheetMetal(newPart, bi.ThicknessVector, bi.ThicknessValue, logs) Then
                Logger.Info("Loo komponendid: Converted '" & bi.Name & "' to sheet metal")
            Else
                ' Sheet metal conversion failed - set dimension properties and create update rule
                MakeComponentsLib.SetDimensionProperties(newPart, bi.ThicknessValue, bi.WidthValue, bi.LengthValue, logs)
                BoundingBoxStockLib.CreateOrUpdateRule(newPart, bi.ThicknessVector, bi.WidthVector, bi.LengthVector, iLogicVb.Automation)
                Logger.Info("Loo komponendid: Created 'Uuenda mõõdud' rule for '" & bi.Name & "'")
            End If
        Else
            MakeComponentsLib.SetDimensionProperties(newPart, bi.ThicknessValue, bi.WidthValue, bi.LengthValue, logs)
            BoundingBoxStockLib.CreateOrUpdateRule(newPart, bi.ThicknessVector, bi.WidthVector, bi.LengthVector, iLogicVb.Automation)
            Logger.Info("Loo komponendid: Created 'Uuenda mõõdud' rule for '" & bi.Name & "'")
        End If
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Save part (always SaveAs - either new file or recreated file)
        Try
            newPart.SaveAs(filePath, False)
            
            ' Read actual path after save (Vault may have renamed the file)
            Dim actualPath As String = newPart.FullDocumentName
            Logger.Info("Loo komponendid: Saved " & actualPath)
            
            ' Only add to createdParts if this is a new part (not recreate)
            ' Recreated parts are already in the assembly
            If Not isRecreate Then
                createdParts.Add(actualPath)
            End If
            
            ' Update body info with actual created part path
            bi.CreatedPartPath = actualPath
            bi.PartExists = True
        Catch ex As Exception
            Logger.Error("Loo komponendid: Failed to save: " & ex.Message)
        End Try
        
        ' Close part
        newPart.Close(False)
    Next
    
    ' Save body data to master document (settings and part references)
    ' Save ALL bodies, not just selected ones, to preserve unprocessed body data
    MakeComponentsLib.SaveBodyDataToMaster(masterDoc, bodies, logs)
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    
    ' Save master document to persist the stored body data
    If bodies.Count > 0 Then
        Try
            masterDoc.Save()
            Logger.Info("Loo komponendid: Master document saved with body references")
        Catch ex As Exception
            Logger.Warn("Loo komponendid: Could not save master document: " & ex.Message)
        End Try
    End If
    
    ' Place in assembly and save
    Dim actualAssemblyPath As String = ""
    If asmDoc IsNot Nothing AndAlso createdParts.Count > 0 Then
        For Each partPath As String In createdParts
            MakeComponentsLib.PlaceComponentGrounded(asmDoc, partPath, logs)
        Next
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Set assembly properties BEFORE save (Vault reads VDS properties at SaveAs time)
        MakeComponentsLib.SetAssemblyProperties(asmDoc, projectName, logs)
        For Each log As String In logs : Logger.Info(log) : Next
        logs.Clear()
        
        ' Save assembly
        If assemblyAction = "CREATE" AndAlso Not String.IsNullOrEmpty(assemblyPath) Then
            Try
                asmDoc.SaveAs(assemblyPath, False)
                actualAssemblyPath = asmDoc.FullDocumentName
                Logger.Info("Loo komponendid: Saved assembly: " & actualAssemblyPath)
            Catch ex As Exception
                Logger.Error("Loo komponendid: Failed to save assembly: " & ex.Message)
            End Try
        ElseIf assemblyAction = "CREATE" Then
            Dim asmFileName As String = defaultProject & "_asm.iam"
            Dim asmFilePath As String = System.IO.Path.Combine(outputFolder, asmFileName)
            Try
                asmDoc.SaveAs(asmFilePath, False)
                actualAssemblyPath = asmDoc.FullDocumentName
                Logger.Info("Loo komponendid: Saved assembly: " & actualAssemblyPath)
            Catch ex As Exception
                Logger.Error("Loo komponendid: Failed to save assembly: " & ex.Message)
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
    
    MakeComponentsLib.SaveGeneralSettings(masterDoc, settingsToSave, logs)
    For Each log As String In logs : Logger.Info(log) : Next
    logs.Clear()
    
    ' Save master document again to persist general settings
    Try
        masterDoc.Save()
        Logger.Info("Loo komponendid: Master document saved with general settings")
    Catch ex As Exception
        Logger.Warn("Loo komponendid: Could not save master document: " & ex.Message)
    End Try
    
    Dim recreatedCount As Integer = selectedBodies.Count - createdParts.Count
    If recreatedCount > 0 Then
        Logger.Info("Loo komponendid: Completed - " & createdParts.Count & " new part(s), " & recreatedCount & " recreated")
    Else
        Logger.Info("Loo komponendid: Completed - created " & createdParts.Count & " part(s)")
    End If
End Sub

' Recalculate axes based on a picked face
Sub RecalculateAxesFromFace(bi As MakeComponentsLib.BodyInfo, face As Face, logs As List(Of String))
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
            
            logs.Add("Loo komponendid: Recalculated axes for '" & bi.Name & "' - T:" & _
                     FormatNumber(bi.ThicknessValue * 10, 2) & " W:" & FormatNumber(bi.WidthValue * 10, 2) & _
                     " L:" & FormatNumber(bi.LengthValue * 10, 2))
        End If
    Catch ex As Exception
        logs.Add("Loo komponendid: Error recalculating axes: " & ex.Message)
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
    
    ' Subfolder
    Dim lblSubfolder As New System.Windows.Forms.Label()
    lblSubfolder.Text = "Alamkaust:"
    lblSubfolder.Left = 310
    lblSubfolder.Top = currentY + 3
    lblSubfolder.Width = 100
    frm.Controls.Add(lblSubfolder)
    
    Dim cboSubfolder As New System.Windows.Forms.ComboBox()
    cboSubfolder.Name = "cboSubfolder"
    cboSubfolder.Left = 415
    cboSubfolder.Top = currentY
    cboSubfolder.Width = 180
    cboSubfolder.DropDownStyle = ComboBoxStyle.DropDown
    cboSubfolder.Items.Add("Detailid")
    cboSubfolder.Items.Add("Poroloon")
    cboSubfolder.Items.Add("Lehtmetall")
    cboSubfolder.SelectedIndex = 0
    frm.Controls.Add(cboSubfolder)
    
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
    cboAssembly.Items.Add("Ära loo koosti")
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
    colStatus.Width = 120
    colStatus.ReadOnly = True
    dgv.Columns.Add(colStatus)
    
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
        
        ' Show status - new or existing part
        If bi.PartExists Then
            Dim partName As String = System.IO.Path.GetFileName(bi.CreatedPartPath)
            dgv.Rows(rowIndex).Cells("colStatus").Value = "* " & partName
        Else
            dgv.Rows(rowIndex).Cells("colStatus").Value = "(uus)"
        End If
        
        dgv.Rows(rowIndex).Cells("colT").Value = FormatNumber(bi.ThicknessValue * 10, 2)
        dgv.Rows(rowIndex).Cells("colW").Value = FormatNumber(bi.WidthValue * 10, 2)
        dgv.Rows(rowIndex).Cells("colL").Value = FormatNumber(bi.LengthValue * 10, 2)
        dgv.Rows(rowIndex).Cells("colSM").Value = bi.ConvertToSheetMetal
        dgv.Rows(rowIndex).Cells("colMat").Value = If(String.IsNullOrEmpty(bi.MaterialName), "", bi.MaterialName)
    Next
    
    ' Store pick index in form Tag (can't use ByRef in lambda)
    frm.Tag = -1
    
    ' Handle button click for face picking
    AddHandler dgv.CellContentClick, Sub(s, e)
        If e.ColumnIndex = dgv.Columns("colPick").Index AndAlso e.RowIndex >= 0 Then
            ' Save current state before closing
            SyncGridToBodyInfo(dgv, bodies)
            ' Store index in form.Tag instead of ByRef parameter
            frm.Tag = CInt(dgv.Rows(e.RowIndex).Tag)
            frm.DialogResult = DialogResult.Retry
            frm.Close()
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
        selectedSubfolder = If(String.IsNullOrEmpty(cboSubfolder.Text), "Detailid", cboSubfolder.Text)
        
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
