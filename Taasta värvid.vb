AddVbFile "Lib/UtilsLib.vb"

Imports Inventor

' Taasta värvid - Restore colors by clearing custom appearances
'
' Works at assembly or part level:
' - Assembly: Clears appearance overrides on all component occurrences
'             and resets each part to use its material appearance
' - Part: Clears appearance overrides on the current part
'
' The appearance will be restored to what is defined by the material.

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        UtilsLib.LogError("Taasta värvid: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "Taasta värvid")
        Exit Sub
    End If
    
    ' Wrap in transaction for single undo
    Dim trans As Transaction = app.TransactionManager.StartTransaction(doc, "Taasta värvid")
    
    Try
        If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
            ProcessAssembly(asmDoc)
            
        ElseIf doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As PartDocument = CType(doc, PartDocument)
            ProcessPart(partDoc)
            
        Else
            UtilsLib.LogError("Taasta värvid: Unsupported document type.")
            MessageBox.Show("Toetatud ainult koostes (.iam) ja detailis (.ipt).", "Taasta värvid")
            trans.Abort()
            Exit Sub
        End If
        
        trans.End()
        UtilsLib.LogInfo("Taasta värvid: Completed successfully")
        
    Catch ex As Exception
        trans.Abort()
        UtilsLib.LogError("Taasta värvid: Error - " & ex.Message)
        MessageBox.Show("Viga: " & ex.Message, "Taasta värvid")
    End Try
End Sub

' Process assembly - clear appearance overrides on all occurrences
Private Sub ProcessAssembly(asmDoc As AssemblyDocument)
    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    Dim clearedCount As Integer = 0
    Dim partCount As Integer = 0
    
    ' Clear assembly-level appearance overrides on occurrences
    For Each obj As Object In asmDef.AppearanceOverridesObjects
        If TypeOf obj Is ComponentOccurrence Then
            Dim occ As ComponentOccurrence = CType(obj, ComponentOccurrence)
            occ.AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
            clearedCount += 1
        End If
    Next
    
    ' Process all leaf occurrences (actual parts, not sub-assemblies)
    For Each occ As ComponentOccurrence In asmDef.Occurrences.AllLeafOccurrences
        Try
            ' Get the part definition
            Dim partDef As PartComponentDefinition = Nothing
            
            If TypeOf occ.Definition Is PartComponentDefinition Then
                partDef = CType(occ.Definition, PartComponentDefinition)
            Else
                Continue For
            End If
            
            ' Get the part document
            Dim partDoc As PartDocument = CType(partDef.Document, PartDocument)
            
            ' Check if we can modify (needs checkout in Vault)
            If Not partDoc.IsModifiable Then
                UtilsLib.LogWarn("Taasta värvid: Skipping read-only part: " & partDoc.DisplayName)
                Continue For
            End If
            
            Dim changed As Boolean = False
            
            ' Update material from global library
            Try
                partDef.Material.UpdateFromGlobal()
            Catch
            End Try
            
            ' Set part appearance to use material appearance
            If partDoc.AppearanceSourceType <> AppearanceSourceTypeEnum.kMaterialAppearance Then
                partDoc.AppearanceSourceType = AppearanceSourceTypeEnum.kMaterialAppearance
                changed = True
            End If
            
            ' Clear feature/face/body level appearance overrides
            For Each overrideObj As Object In partDef.AppearanceOverridesObjects
                If TypeOf overrideObj Is SurfaceBody Then
                    CType(overrideObj, SurfaceBody).AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
                    changed = True
                ElseIf TypeOf overrideObj Is PartFeature Then
                    CType(overrideObj, PartFeature).AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
                    changed = True
                ElseIf TypeOf overrideObj Is Face Then
                    CType(overrideObj, Face).AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
                    changed = True
                End If
            Next
            
            If changed Then
                partCount += 1
                UtilsLib.LogInfo("Taasta värvid: Cleared overrides in " & partDoc.DisplayName)
            End If
            
        Catch ex As Exception
            UtilsLib.LogWarn("Taasta värvid: Error processing " & occ.Name & ": " & ex.Message)
        End Try
    Next
    
    UtilsLib.LogInfo("Taasta värvid: Processed " & partCount & " part(s), cleared " & clearedCount & " assembly override(s)")
End Sub

' Process a single part - clear all appearance overrides
Private Sub ProcessPart(partDoc As PartDocument)
    Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim clearedCount As Integer = 0
    
    ' Update material from global library
    Try
        partDef.Material.UpdateFromGlobal()
    Catch
    End Try
    
    ' Set part appearance to use material appearance
    If partDoc.AppearanceSourceType <> AppearanceSourceTypeEnum.kMaterialAppearance Then
        partDoc.AppearanceSourceType = AppearanceSourceTypeEnum.kMaterialAppearance
        clearedCount += 1
    End If
    
    ' Clear feature/face/body level appearance overrides
    For Each obj As Object In partDef.AppearanceOverridesObjects
        If TypeOf obj Is SurfaceBody Then
            CType(obj, SurfaceBody).AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
            clearedCount += 1
        ElseIf TypeOf obj Is PartFeature Then
            CType(obj, PartFeature).AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
            clearedCount += 1
        ElseIf TypeOf obj Is Face Then
            CType(obj, Face).AppearanceSourceType = AppearanceSourceTypeEnum.kPartAppearance
            clearedCount += 1
        End If
    Next
    
    UtilsLib.LogInfo("Taasta värvid: Cleared " & clearedCount & " override(s) in " & partDoc.DisplayName)
End Sub
