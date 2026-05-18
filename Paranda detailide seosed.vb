' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Paranda detailide seosed - Repair body-to-part links in master document
' 
' This script scans for derived parts that reference the current master document
' and re-populates the body link properties that may have been lost.
'
' How it works:
' 1. Opens the current master part document
' 2. For each solid body, searches for derived parts in nearby folders
' 3. Checks if the derived part references this master AND the specific body
' 4. Re-links found parts to the corresponding body properties
'
' Search locations (relative to master):
' - Sibling folders: ../Karkass/Detailid, ../Poroloon/Detailid
' - Element structure: ../../<element>/Karkass/Detailid, etc.
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
    
    ' Validate document
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
    
    UtilsLib.LogInfo("Paranda seosed: Starting repair for " & masterDoc.DisplayName)
    UtilsLib.LogInfo("Paranda seosed: Master path: " & masterDoc.FullDocumentName)
    
    ' Get project root for relative path storage
    Dim projectRoot As String = UtilsLib.GetProjectPath(masterDoc.FullDocumentName)
    If String.IsNullOrEmpty(projectRoot) Then
        projectRoot = System.IO.Path.GetDirectoryName(masterDoc.FullDocumentName)
    End If
    UtilsLib.LogInfo("Paranda seosed: Project root: " & projectRoot)
    
    ' Build list of folders to search for derived parts
    Dim searchFolders As List(Of String) = GetSearchFolders(masterDoc.FullDocumentName, projectRoot)
    UtilsLib.LogInfo("Paranda seosed: Search folders: " & searchFolders.Count)
    For Each folder As String In searchFolders
        UtilsLib.LogInfo("  - " & folder)
    Next
    
    ' Get current body data (with any existing links)
    Dim bodies As List(Of MakeComponentsLib.BodyInfo) = MakeComponentsLib.GetBodiesWithAxes(masterDoc)
    Dim storedData As List(Of MakeComponentsLib.StoredBodyData) = MakeComponentsLib.LoadBodyDataFromMaster(masterDoc)
    
    ' Apply any existing stored data
    If storedData.Count > 0 Then
        Dim masterFolder As String = System.IO.Path.GetDirectoryName(masterDoc.FullDocumentName)
        MakeComponentsLib.ApplyStoredDataToBodies(bodies, storedData, masterFolder, projectRoot, projectRoot)
    End If
    
    ' Count bodies with and without links
    Dim linkedCount As Integer = 0
    Dim unlinkedCount As Integer = 0
    Dim missingMetadataCount As Integer = 0
    For Each bi As MakeComponentsLib.BodyInfo In bodies
        If bi.PartExists AndAlso Not String.IsNullOrEmpty(bi.CreatedPartPath) Then
            linkedCount += 1
            ' Check if linked body is missing metadata (material)
            If String.IsNullOrEmpty(bi.MaterialName) Then
                missingMetadataCount += 1
            End If
        Else
            unlinkedCount += 1
        End If
    Next
    
    UtilsLib.LogInfo("Paranda seosed: Bodies with links: " & linkedCount & ", without links: " & unlinkedCount & ", missing metadata: " & missingMetadataCount)
    
    ' If all have links but some are missing metadata, offer to refresh
    If unlinkedCount = 0 AndAlso missingMetadataCount = 0 Then
        UtilsLib.LogInfo("Paranda seosed: All bodies already have links and metadata - nothing to repair")
        MessageBox.Show("Kõigil kehadel on juba seosed ja omadused olemas.", "Paranda detailide seosed")
        Exit Sub
    End If
    
    ' If all have links but missing metadata, refresh from existing parts
    If unlinkedCount = 0 AndAlso missingMetadataCount > 0 Then
        UtilsLib.LogInfo("Paranda seosed: " & missingMetadataCount & " body(ies) missing metadata - refreshing from parts")
        
        Dim refreshedCount As Integer = 0
        For Each bi As MakeComponentsLib.BodyInfo In bodies
            If bi.PartExists AndAlso Not String.IsNullOrEmpty(bi.CreatedPartPath) Then
                If System.IO.File.Exists(bi.CreatedPartPath) Then
                    Dim oldMaterial As String = bi.MaterialName
                    If MakeComponentsLib.ReadPropertiesFromPart(app, bi.CreatedPartPath, bi) Then
                        If bi.MaterialName <> oldMaterial Then
                            refreshedCount += 1
                            UtilsLib.LogInfo("Paranda seosed: Refreshed '" & bi.Name & "' - Material: " & bi.MaterialName)
                        End If
                    End If
                End If
            End If
        Next
        
        If refreshedCount = 0 Then
            MessageBox.Show("Ei leidnud uusi omadusi taastamiseks.", "Paranda detailide seosed")
            Exit Sub
        End If
        
        ' Ask to save refreshed metadata
        Dim refreshMsg As String = "Taastati " & refreshedCount & " keha omadused:" & vbCrLf & vbCrLf
        For Each bi As MakeComponentsLib.BodyInfo In bodies
            If Not String.IsNullOrEmpty(bi.MaterialName) OrElse bi.ConvertToSheetMetal Then
                Dim props As String = ""
                If Not String.IsNullOrEmpty(bi.MaterialName) Then props &= " [" & bi.MaterialName & "]"
                If bi.ConvertToSheetMetal Then props &= " [Lehtmetall]"
                refreshMsg &= "• " & bi.Name & props & vbCrLf
            End If
        Next
        refreshMsg &= vbCrLf & "Kas salvestada omadused?"
        
        If MessageBox.Show(refreshMsg, "Paranda detailide seosed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Exit Sub
        End If
        
        MakeComponentsLib.SaveBodyDataToMaster(masterDoc, bodies, projectRoot)
        Try
            masterDoc.Save()
            MessageBox.Show("Taastatud " & refreshedCount & " keha omadused.", "Paranda detailide seosed")
        Catch ex As Exception
            MessageBox.Show("Ei saanud salvestada: " & ex.Message, "Paranda detailide seosed")
        End Try
        Exit Sub
    End If
    
    ' Scan for derived parts
    Dim foundLinks As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)  ' bodyName -> partPath
    Dim masterFullPath As String = masterDoc.FullDocumentName
    
    UtilsLib.LogInfo("Paranda seosed: Scanning for derived parts...")
    
    For Each folder As String In searchFolders
        If Not System.IO.Directory.Exists(folder) Then Continue For
        
        Try
            Dim iptFiles() As String = System.IO.Directory.GetFiles(folder, "*.ipt", System.IO.SearchOption.TopDirectoryOnly)
            For Each iptPath As String In iptFiles
                ' Skip OldVersions
                If iptPath.IndexOf("\OldVersions\", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
                
                ' Check if this part derives from our master
                Dim derivedBodyName As String = GetDerivedBodyName(app, iptPath, masterFullPath)
                If Not String.IsNullOrEmpty(derivedBodyName) Then
                    If Not foundLinks.ContainsKey(derivedBodyName) Then
                        foundLinks.Add(derivedBodyName, iptPath)
                        UtilsLib.LogInfo("Paranda seosed: Found '" & derivedBodyName & "' -> " & System.IO.Path.GetFileName(iptPath))
                    End If
                End If
            Next
        Catch ex As Exception
            UtilsLib.LogWarn("Paranda seosed: Error scanning " & folder & ": " & ex.Message)
        End Try
    Next
    
    UtilsLib.LogInfo("Paranda seosed: Found " & foundLinks.Count & " derived part(s)")
    
    ' Apply found links to bodies and restore properties from part files
    Dim repairedCount As Integer = 0
    For Each bi As MakeComponentsLib.BodyInfo In bodies
        ' Skip if already has a valid link
        If bi.PartExists AndAlso Not String.IsNullOrEmpty(bi.CreatedPartPath) Then Continue For
        
        ' Check if we found a part for this body
        If foundLinks.ContainsKey(bi.Name) Then
            Dim partPath As String = foundLinks(bi.Name)
            bi.CreatedPartPath = partPath
            bi.PartExists = True
            bi.Selected = False
            
            ' Read properties from the part file (material, sheet metal, dimensions, etc.)
            If MakeComponentsLib.ReadPropertiesFromPart(app, partPath, bi) Then
                UtilsLib.LogInfo("Paranda seosed: Restored properties for '" & bi.Name & "' - Material: " & bi.MaterialName)
            End If
            
            repairedCount += 1
            UtilsLib.LogInfo("Paranda seosed: Repaired link for '" & bi.Name & "'")
        End If
    Next
    
    If repairedCount = 0 Then
        UtilsLib.LogInfo("Paranda seosed: No new links found to repair")
        MessageBox.Show("Ei leidnud uusi seoseid parandamiseks." & vbCrLf & vbCrLf & _
                       "Kontrollitud kaustad:" & vbCrLf & String.Join(vbCrLf, searchFolders.ToArray()), _
                       "Paranda detailide seosed")
        Exit Sub
    End If
    
    ' Ask user to confirm
    Dim confirmMsg As String = "Leiti " & repairedCount & " uut seost:" & vbCrLf & vbCrLf
    For Each bi As MakeComponentsLib.BodyInfo In bodies
        If foundLinks.ContainsKey(bi.Name) Then
            Dim props As String = ""
            If Not String.IsNullOrEmpty(bi.MaterialName) Then
                props &= " [" & bi.MaterialName & "]"
            End If
            If bi.ConvertToSheetMetal Then
                props &= " [Lehtmetall]"
            End If
            confirmMsg &= "• " & bi.Name & " -> " & System.IO.Path.GetFileName(bi.CreatedPartPath) & props & vbCrLf
        End If
    Next
    confirmMsg &= vbCrLf & "Kas salvestada seosed ja omadused?"
    
    If MessageBox.Show(confirmMsg, "Paranda detailide seosed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
        UtilsLib.LogInfo("Paranda seosed: Cancelled by user")
        Exit Sub
    End If
    
    ' Save repaired data to master document
    MakeComponentsLib.SaveBodyDataToMaster(masterDoc, bodies, projectRoot)
    
    Try
        masterDoc.Save()
        UtilsLib.LogInfo("Paranda seosed: Saved " & repairedCount & " repaired link(s)")
        MessageBox.Show("Parandatud " & repairedCount & " seost.", "Paranda detailide seosed")
    Catch ex As Exception
        UtilsLib.LogError("Paranda seosed: Could not save: " & ex.Message)
        MessageBox.Show("Ei saanud salvestada: " & ex.Message, "Paranda detailide seosed")
    End Try
End Sub

' Build list of folders to search for derived parts
Function GetSearchFolders(masterPath As String, projectRoot As String) As List(Of String)
    Dim folders As New List(Of String)
    Dim masterFolder As String = System.IO.Path.GetDirectoryName(masterPath)
    
    ' Common detail folder patterns relative to master location
    ' Master is typically in: .../Aluselemendid/<element>/Eskiis/
    ' Parts are in: .../Aluselemendid/<element>/Karkass/Detailid/ or Poroloon/Detailid/
    
    ' Try to detect element root from master path
    Dim elementRoot As String = ""
    Dim masterParent As String = System.IO.Path.GetDirectoryName(masterFolder)
    If Not String.IsNullOrEmpty(masterParent) Then
        ' Check if parent folder name suggests we're in Eskiis
        Dim folderName As String = System.IO.Path.GetFileName(masterFolder)
        If String.Equals(folderName, "Eskiis", StringComparison.OrdinalIgnoreCase) Then
            elementRoot = masterParent
        Else
            ' Maybe master is directly in element root
            elementRoot = masterFolder
        End If
    End If
    
    ' Add standard part folders under element root
    If Not String.IsNullOrEmpty(elementRoot) Then
        AddFolderIfExists(folders, System.IO.Path.Combine(elementRoot, "Karkass", "Detailid"))
        AddFolderIfExists(folders, System.IO.Path.Combine(elementRoot, "Poroloon", "Detailid"))
        AddFolderIfExists(folders, System.IO.Path.Combine(elementRoot, "Karkass"))
        AddFolderIfExists(folders, System.IO.Path.Combine(elementRoot, "Poroloon"))
        AddFolderIfExists(folders, System.IO.Path.Combine(elementRoot, "Detailid"))
    End If
    
    ' Also check siblings of master folder
    If Not String.IsNullOrEmpty(masterParent) Then
        AddFolderIfExists(folders, System.IO.Path.Combine(masterParent, "Karkass", "Detailid"))
        AddFolderIfExists(folders, System.IO.Path.Combine(masterParent, "Poroloon", "Detailid"))
        AddFolderIfExists(folders, System.IO.Path.Combine(masterParent, "Detailid"))
    End If
    
    ' Check master folder itself (in case parts are in same folder)
    AddFolderIfExists(folders, masterFolder)
    
    ' If we have project root, also check standard project structure
    If Not String.IsNullOrEmpty(projectRoot) Then
        ' Try to find element name from master path
        Dim projectName As String = UtilsLib.ExtractProjectName(masterPath)
        If Not String.IsNullOrEmpty(projectName) Then
            Dim baseElementsPath As String = System.IO.Path.Combine(projectRoot, BaseElementLayoutLib.SEG_BASE_ELEMENTS)
            If System.IO.Directory.Exists(baseElementsPath) Then
                ' Search all elements in the project
                Try
                    For Each elementDir As String In System.IO.Directory.GetDirectories(baseElementsPath)
                        AddFolderIfExists(folders, System.IO.Path.Combine(elementDir, "Karkass", "Detailid"))
                        AddFolderIfExists(folders, System.IO.Path.Combine(elementDir, "Poroloon", "Detailid"))
                    Next
                Catch
                End Try
            End If
        End If
    End If
    
    Return folders
End Function

Sub AddFolderIfExists(folders As List(Of String), path As String)
    If String.IsNullOrEmpty(path) Then Exit Sub
    If folders.Contains(path) Then Exit Sub
    If System.IO.Directory.Exists(path) Then
        folders.Add(path)
    End If
End Sub

' Check if a part file derives from the specified master and return the derived body name
' Returns empty string if not a derived part from this master
Function GetDerivedBodyName(app As Inventor.Application, partPath As String, masterPath As String) As String
    Dim partDoc As PartDocument = Nothing
    Dim wasAlreadyOpen As Boolean = False
    
    Try
        ' Check if document is already open
        For Each doc As Document In app.Documents
            If doc.FullDocumentName.Equals(partPath, StringComparison.OrdinalIgnoreCase) Then
                partDoc = CType(doc, PartDocument)
                wasAlreadyOpen = True
                Exit For
            End If
        Next
        
        ' Open document if not already open (invisible)
        If partDoc Is Nothing Then
            partDoc = CType(app.Documents.Open(partPath, False), PartDocument)
        End If
        
        ' Check for derived part components
        Dim dpcs As DerivedPartComponents = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
        If dpcs.Count = 0 Then Return ""
        
        For Each dpc As DerivedPartComponent In dpcs
            Try
                ' Get the referenced file path
                Dim refFile As String = dpc.ReferencedFile.FullFileName
                
                ' Check if it references our master
                If refFile.Equals(masterPath, StringComparison.OrdinalIgnoreCase) Then
                    ' Found a match - now get the body name
                    ' The body name can be read from Description property or by checking included solids
                    Dim bodyName As String = GetIncludedBodyName(dpc)
                    If Not String.IsNullOrEmpty(bodyName) Then
                        Return bodyName
                    End If
                    
                    ' Fallback: try Description property
                    Try
                        Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
                        Dim desc As String = CStr(designProps.Item("Description").Value)
                        If Not String.IsNullOrEmpty(desc) Then
                            Return desc.Trim()
                        End If
                    Catch
                    End Try
                End If
            Catch
            End Try
        Next
        
    Catch ex As Exception
        ' Silently handle - return empty
    Finally
        ' Close document if we opened it
        If partDoc IsNot Nothing AndAlso Not wasAlreadyOpen Then
            Try
                partDoc.Close(True)
            Catch
            End Try
        End If
    End Try
    
    Return ""
End Function

' Get the name of the included solid body from a derived part component
Function GetIncludedBodyName(dpc As DerivedPartComponent) As String
    Try
        ' Access the definition to get included solids
        Dim dpDef As Object = dpc.Definition
        If dpDef Is Nothing Then Return ""
        
        ' Try to access Solids collection (may not work for all derivation types)
        Try
            Dim solids As Object = CallByName(dpDef, "Solids", Microsoft.VisualBasic.CallType.Get)
            If solids IsNot Nothing Then
                For Each dpe As Object In solids
                    Try
                        Dim includeEntity As Boolean = CBool(CallByName(dpe, "IncludeEntity", Microsoft.VisualBasic.CallType.Get))
                        If includeEntity Then
                            Dim refEntity As Object = CallByName(dpe, "ReferencedEntity", Microsoft.VisualBasic.CallType.Get)
                            If refEntity IsNot Nothing AndAlso TypeOf refEntity Is SurfaceBody Then
                                Return CType(refEntity, SurfaceBody).Name
                            End If
                        End If
                    Catch
                    End Try
                Next
            End If
        Catch
        End Try
        
    Catch
    End Try
    
    Return ""
End Function
