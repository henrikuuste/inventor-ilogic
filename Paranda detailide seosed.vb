' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Paranda detailide seosed - Repair body-to-part links in master document
' 
' Runs discovery (same as Loo detailid) and rebuilds the part GUID cache.
' Properties are read from actual derived part files, not stored on the master.
' ============================================================================

AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
AddReference "Connectivity.InventorAddin.EdmAddin"

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/FileSearchLib.vb"
AddVbFile "Lib/CustomPropertiesLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"
AddVbFile "Lib/MakeComponentsLib.vb"
AddVbFile "Lib/BaseElementLayoutLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Paranda seosed: No active document")
        MessageBox.Show("Ava esmalt multi-body master detail.", "Paranda detailide seosed")
        Exit Sub
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        UtilsLib.LogError("Paranda seosed: Active document is not a part")
        MessageBox.Show("Aktiivseks dokumendiks peab olema detail (.ipt).", "Paranda detailide seosed")
        Exit Sub
    End If
    
    Dim masterDoc As PartDocument = CType(app.ActiveDocument, PartDocument)
    
    If masterDoc.ComponentDefinition.SurfaceBodies.Count < 1 Then
        UtilsLib.LogError("Paranda seosed: No solid bodies in part")
        MessageBox.Show("Detailis puuduvad tahked kehad.", "Paranda detailide seosed")
        Exit Sub
    End If
    
    If String.IsNullOrEmpty(masterDoc.FullDocumentName) Then
        UtilsLib.LogError("Paranda seosed: Master document not saved")
        MessageBox.Show("Salvesta esmalt master-detail.", "Paranda detailide seosed")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Paranda seosed: Starting discovery for " & masterDoc.DisplayName)
    
    Dim projectRoot As String = UtilsLib.GetProjectPath(masterDoc.FullDocumentName)
    If String.IsNullOrEmpty(projectRoot) Then
        projectRoot = System.IO.Path.GetDirectoryName(masterDoc.FullDocumentName)
    End If
    
    Dim elementRoot As String = BaseElementLayoutLib.DetectElementRootFromMasterPath( _
        masterDoc.FullDocumentName, projectRoot, UtilsLib.ExtractProjectName(masterDoc.FullDocumentName))
    
    Dim bodies As List(Of MakeComponentsLib.BodyInfo) = MakeComponentsLib.GetBodiesWithAxes(masterDoc)
    
    Dim orphans As List(Of MakeComponentsLib.OrphanPartInfo) = _
        MakeComponentsLib.LinkBodiesFromDiscovery(app, masterDoc, bodies, projectRoot, elementRoot)
    
    Dim linkedCount As Integer = 0
    For Each bi As MakeComponentsLib.BodyInfo In bodies
        If bi.PartExists Then linkedCount += 1
    Next
    
    If linkedCount = 0 AndAlso orphans.Count = 0 Then
        MessageBox.Show("Ei leitud ühtegi sellest masterist tuletatud detaili.", "Paranda detailide seosed")
        Exit Sub
    End If
    
    Dim confirmMsg As String = "Leiti " & linkedCount & " keha seost:" & vbCrLf & vbCrLf
    For Each bi As MakeComponentsLib.BodyInfo In bodies
        If bi.PartExists Then
            Dim props As String = ""
            If Not String.IsNullOrEmpty(bi.MaterialName) Then props &= " [" & bi.MaterialName & "]"
            If bi.ConvertToSheetMetal Then props &= " [Lehtmetall]"
            confirmMsg &= "• " & bi.Name & " -> " & System.IO.Path.GetFileName(bi.CreatedPartPath) & props & vbCrLf
        End If
    Next
    
    If orphans.Count > 0 Then
        confirmMsg &= vbCrLf & "Hoiatus - " & orphans.Count & " detaili ilma kehata:" & vbCrLf
        For Each orphan As MakeComponentsLib.OrphanPartInfo In orphans
            confirmMsg &= "• " & System.IO.Path.GetFileName(orphan.PartPath) & vbCrLf
        Next
    End If
    
    confirmMsg &= vbCrLf & "Kas salvestada GUID vahemälu masterisse?"
    
    If MessageBox.Show(confirmMsg, "Paranda detailide seosed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
        Exit Sub
    End If
    
    MakeComponentsLib.SaveBodyDataToMaster(masterDoc, bodies, projectRoot)
    
    Try
        masterDoc.Save()
        UtilsLib.LogInfo("Paranda seosed: Saved cache for " & linkedCount & " link(s)")
        MessageBox.Show("Salvestatud " & linkedCount & " seose GUID vahemälu.", "Paranda detailide seosed")
    Catch ex As Exception
        UtilsLib.LogError("Paranda seosed: Could not save: " & ex.Message)
        MessageBox.Show("Ei saanud salvestada: " & ex.Message, "Paranda detailide seosed")
    End Try
End Sub
