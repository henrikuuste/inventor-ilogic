' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Mõõdud - Materjali gabariitmõõtude kalkulaator
' 
' Töötab nii detaili kui koostu dokumentidega:
' - Detailis: töötleb aktiivset detaili
' - Koostus: töötleb valitud detailid
' - Töötab nii tavaliste, lehtmetalli kui pinnalaotusega detailidega
'
' Shows all parts in a single DataGridView dialog where user can:
' - See T/W/L measurements for each part
' - Change thickness axis (X/Y/Z/Custom) - for normal parts
' - Flip width/length
' - Pick a face for custom axis orientation
'
' Pinnalaotus (Unwrap) is detected from UnwrapFeatures, not from part SubType.
' If Unwrap exists, Telg defaults to Pinnalaotus but you can switch to Normal (gabariit)
' or Lehtmetall (flat pattern) when the part is sheet-metal subtype; choice is stored in BB_DimensionSource.
' Pinnalaotus thickness axis defaults to the unwrap flat surface plane normal; W/L come from that basis on
' the measurement body. If no planar unwrap face is found, falls back to smallest-extent heuristic on the body.
' Use "Vali pind" if the automatic thickness direction is wrong.
'
' Sheet metal without Unwrap uses flat pattern for dimensions (Telg locked to Lehtmetall).
' Pure Pinnalaotus uses unwrap+thicken for dimensions.
' Normal parts use bounding box with configurable axis orientation.
'
' Registers "Uuenda" rule handler that auto-updates dimension properties
' (Thickness, Width, Length) on geometry/parameter changes.
' ============================================================================

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/DocumentUpdateLib.vb"
AddVbFile "Lib/DimensionUpdateLib.vb"
AddVbFile "Lib/BoundingBoxStockLib.vb"
AddVbFile "Lib/UnwrapLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    UtilsLib.SetLogger(Logger)
    
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        Logger.Error("Mõõdud: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Mõõdud")
        Exit Sub
    End If

    ' Collect parts to process
    Dim partDocs As New List(Of PartDocument)
    Dim partNames As New List(Of String)
    Dim thicknessAxes As New List(Of String)
    Dim widthAxes As New List(Of String)
    Dim lengthAxes As New List(Of String)
    Dim customAxisDescs As New List(Of String)
    Dim selectedFlags As New List(Of Boolean)

    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        ' Single part document (works with both normal and sheet metal parts)
        Dim partDoc As PartDocument = CType(doc, PartDocument)
        CollectPartData(partDoc, partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
        Logger.Info("Mõõdud: Processing single part - " & partDoc.DisplayName)
        
    ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        ' Assembly document - process selected parts or all parts if none selected
        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        Dim sel As SelectSet = asmDoc.SelectSet
        Dim useAllParts As Boolean = (sel Is Nothing OrElse sel.Count = 0)

        ' Collect unique part occurrences
        Dim processedDefs As New HashSet(Of Object)
        
        If useAllParts Then
            ' No selection - collect all parts from assembly
            CollectAllPartsFromAssembly(asmDoc.ComponentDefinition.Occurrences, processedDefs, _
                                        partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
            Logger.Info("Mõõdud: No selection - using all " & partDocs.Count & " part(s) from assembly")
        Else
            ' Use selection
            For Each selObj As Object In sel
                If TypeOf selObj Is ComponentOccurrence Then
                    Dim occ As ComponentOccurrence = CType(selObj, ComponentOccurrence)
                    If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        ' Avoid duplicates from same definition
                        If Not processedDefs.Contains(occ.Definition) Then
                            processedDefs.Add(occ.Definition)
                            Try
                                Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
                                CollectPartData(partDoc, partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
                            Catch
                            End Try
                        End If
                    End If
                End If
            Next
            Logger.Info("Mõõdud: Processing " & partDocs.Count & " part(s) from assembly selection")
        End If

        If partDocs.Count = 0 Then
            MessageBox.Show("Koostus ei leitud sobivaid detaile.", "Mõõdud")
            Exit Sub
        End If
    Else
        MessageBox.Show("See reegel töötab ainult detaili (.ipt) või koostu (.iam) dokumentidega.", "Mõõdud")
        Exit Sub
    End If

    ' Dialog loop for face picking
    Dim dlgResult As DialogResult
    Dim pickRowIndex As Integer = -1

    Do
        dlgResult = ShowBatchDialog(app, partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, _
                                    customAxisDescs, selectedFlags, pickRowIndex)

        If dlgResult = DialogResult.Retry AndAlso pickRowIndex >= 0 AndAlso pickRowIndex < partDocs.Count Then
            ' Sheet metal flat pattern: no face pick
            If IsLehtmetallMode(customAxisDescs(pickRowIndex)) Then
                pickRowIndex = -1
                Continue Do
            End If
            
            ' User clicked "Vali pind" or Kohandatud — do face pick
            Dim partDoc As PartDocument = partDocs(pickRowIndex)
            
            ' In Pinnalaotus mode, first ensure we have a valid measurement body
            If customAxisDescs(pickRowIndex) = "Pinnalaotus" Then
                Dim pinnBody As SurfaceBody = TryAutoDetectPinnalaotusBody(partDoc)
                If pinnBody Is Nothing Then
                    Logger.Info("Mõõdud: Pinnalaotus body not found for '" & partNames(pickRowIndex) & "', prompting user")
                    pinnBody = PromptForPinnalaotusBody(partDoc)
                    If pinnBody Is Nothing Then
                        ' User cancelled body selection
                        pickRowIndex = -1
                        Continue Do
                    End If
                End If
            End If
            
            Logger.Info("Mõõdud: Picking face for '" & partNames(pickRowIndex) & "'")

            Try
                Dim planeDesc As String = ""
                Dim pickedVector As String = BoundingBoxStockLib.PickPlaneForThickness(app, planeDesc, True)
                
                If pickedVector <> "" Then
                    thicknessAxes(pickRowIndex) = pickedVector
                    If customAxisDescs(pickRowIndex) <> "Pinnalaotus" Then
                        customAxisDescs(pickRowIndex) = planeDesc
                    End If
                    
                    ' Compute perpendicular vectors for width/length
                    Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
                    Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
                    Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
                    BoundingBoxStockLib.ParseVectorComponents(pickedVector, tx, ty, tz)
                    BoundingBoxStockLib.ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
                    
                    Dim widthExtent As Double
                    Dim lengthExtent As Double
                    If customAxisDescs(pickRowIndex) = "Pinnalaotus" Then
                        ' Use autodetect for Pinnalaotus body - never fall back to all-body measurement
                        Dim measBody As SurfaceBody = TryAutoDetectPinnalaotusBody(partDoc)
                        If measBody IsNot Nothing Then
                            widthExtent = UnwrapLib.GetOrientedExtentForBody(measBody, wx, wy, wz)
                            lengthExtent = UnwrapLib.GetOrientedExtentForBody(measBody, lx, ly, lz)
                        Else
                            ' Fall back to unwrap surface if available
                            Dim unwrapFeat As UnwrapFeature = UnwrapLib.GetUnwrapFeature(partDoc)
                            Dim unwrapSurf As SurfaceBody = If(unwrapFeat IsNot Nothing, UnwrapLib.GetUnwrappedSurfaceBody(unwrapFeat), Nothing)
                            If unwrapSurf IsNot Nothing Then
                                widthExtent = UnwrapLib.GetOrientedExtentForBody(unwrapSurf, wx, wy, wz)
                                lengthExtent = UnwrapLib.GetOrientedExtentForBody(unwrapSurf, lx, ly, lz)
                            Else
                                widthExtent = 0
                                lengthExtent = 0
                            End If
                        End If
                    Else
                        widthExtent = BoundingBoxStockLib.GetOrientedExtent(partDoc, wx, wy, wz)
                        lengthExtent = BoundingBoxStockLib.GetOrientedExtent(partDoc, lx, ly, lz)
                    End If
                    
                    If lengthExtent >= widthExtent Then
                        lengthAxes(pickRowIndex) = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                        widthAxes(pickRowIndex) = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                    Else
                        lengthAxes(pickRowIndex) = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                        widthAxes(pickRowIndex) = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                    End If
                    
                    Logger.Info("Mõõdud: Applied custom axis for '" & partNames(pickRowIndex) & "' - " & planeDesc)
                End If
            Catch
                ' User cancelled pick
            End Try

            pickRowIndex = -1
            Continue Do
        End If

        Exit Do
    Loop

    If dlgResult <> DialogResult.OK Then
        Logger.Info("Mõõdud: Cancelled by user")
        Exit Sub
    End If

    ' Apply rules to selected parts
    Dim processedCount As Integer = 0
    DocumentUpdateLib.SetLogger(Logger)
    DimensionUpdateLib.SetLogger(Logger)
    For i As Integer = 0 To partDocs.Count - 1
        If selectedFlags(i) Then
            Dim partDoc As PartDocument = partDocs(i)
            Dim dimSrc As String = UnwrapLib.DIMENSION_SOURCE_NORMAL
            If customAxisDescs(i) = "Lehtmetall" Then
                dimSrc = UnwrapLib.DIMENSION_SOURCE_LEHTMETALL
            ElseIf customAxisDescs(i) = "Pinnalaotus" Then
                dimSrc = UnwrapLib.DIMENSION_SOURCE_PINNALAOTUS
            End If
            If dimSrc = UnwrapLib.DIMENSION_SOURCE_PINNALAOTUS Then
                UnwrapLib.StorePinnalaotusMeasurementBodyProperty(partDoc)
            End If
            DimensionUpdateLib.RegisterDimensionHandler(partDoc, iLogicVb.Automation, thicknessAxes(i), widthAxes(i), lengthAxes(i), dimSrc)
            processedCount += 1
            Logger.Info("Mõõdud: Updated '" & partNames(i) & "' - T:" & thicknessAxes(i) & " W:" & widthAxes(i) & " L:" & lengthAxes(i) & " [" & dimSrc & "]")
        End If
    Next

    Logger.Info("Mõõdud: Completed - processed " & processedCount & " part(s)")
End Sub

' ============================================================================
' Normal-part axis lists (gabariit X/Y/Z või Kohandatud; works when Unwrap exists if user chose Normal)
' ============================================================================
Sub AppendNormalPartAxisLists(ByVal partDoc As PartDocument, _
                              ByVal thicknessAxes As List(Of String), _
                              ByVal widthAxes As List(Of String), _
                              ByVal lengthAxes As List(Of String), _
                              ByVal customAxisDescs As List(Of String))
    Dim thicknessAxis As String = BoundingBoxStockLib.GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
    Dim widthAxis As String = BoundingBoxStockLib.GetCustomPropertyValue(partDoc, "BB_WidthAxis", "")
    Dim lengthAxis As String = ""
    Dim customAxisDesc As String = ""
    
    Dim xSize As Double = 0, ySize As Double = 0, zSize As Double = 0
    BoundingBoxStockLib.GetBoundingBoxSizes(partDoc, xSize, ySize, zSize)
    
    If thicknessAxis = "" Then
        If Not BoundingBoxStockLib.AutoDetectAxesFromGeometry(partDoc, thicknessAxis, widthAxis, lengthAxis) Then
            BoundingBoxStockLib.AutoDetectAxes(xSize, ySize, zSize, thicknessAxis, widthAxis, lengthAxis)
        End If
        
        If BoundingBoxStockLib.IsVectorFormat(thicknessAxis) Then
            Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
            BoundingBoxStockLib.ParseVectorComponents(thicknessAxis, tx, ty, tz)
            customAxisDesc = "Auto (" & BoundingBoxStockLib.FormatVectorDesc(tx, ty, tz) & ")"
        End If
    ElseIf BoundingBoxStockLib.IsVectorFormat(thicknessAxis) Then
        Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
        If BoundingBoxStockLib.ParseVectorComponents(thicknessAxis, tx, ty, tz) Then
            customAxisDesc = "Custom (" & BoundingBoxStockLib.FormatVectorDesc(tx, ty, tz) & ")"
            
            If widthAxis = "" OrElse Not BoundingBoxStockLib.IsVectorFormat(widthAxis) Then
                Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
                Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
                BoundingBoxStockLib.ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
                
                Dim widthExtent As Double = BoundingBoxStockLib.GetOrientedExtent(partDoc, wx, wy, wz)
                Dim lengthExtent As Double = BoundingBoxStockLib.GetOrientedExtent(partDoc, lx, ly, lz)
                
                If lengthExtent >= widthExtent Then
                    lengthAxis = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                    widthAxis = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                Else
                    lengthAxis = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                    widthAxis = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                End If
            Else
                Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
                BoundingBoxStockLib.ParseVectorComponents(widthAxis, wx, wy, wz)
                Dim lx As Double = ty * wz - tz * wy
                Dim ly As Double = tz * wx - tx * wz
                Dim lz As Double = tx * wy - ty * wx
                lengthAxis = BoundingBoxStockLib.VectorToString(lx, ly, lz)
            End If
        Else
            thicknessAxis = "Z"
            BoundingBoxStockLib.AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
        End If
    Else
        If widthAxis = "" Then
            BoundingBoxStockLib.AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
        Else
            lengthAxis = BoundingBoxStockLib.GetRemainingAxis(thicknessAxis, widthAxis)
        End If
    End If
    
    thicknessAxes.Add(thicknessAxis)
    widthAxes.Add(widthAxis)
    lengthAxes.Add(lengthAxis)
    customAxisDescs.Add(customAxisDesc)
End Sub

' ============================================================================
' Collect part data and auto-detect axes
' ============================================================================
Sub CollectPartData(ByVal partDoc As PartDocument, _
                    ByVal partDocs As List(Of PartDocument), _
                    ByVal partNames As List(Of String), _
                    ByVal thicknessAxes As List(Of String), _
                    ByVal widthAxes As List(Of String), _
                    ByVal lengthAxes As List(Of String), _
                    ByVal customAxisDescs As List(Of String), _
                    ByVal selectedFlags As List(Of Boolean))
    
    partDocs.Add(partDoc)
    
    ' Build display name: filename + description
    Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName)
    Dim desc As String = GetPartDescription(partDoc)
    Dim displayName As String = fileName
    If desc <> "" Then displayName &= " - " & desc
    partNames.Add(displayName)
    
    If UnwrapLib.HasUnwrapFeature(partDoc) Then
        Dim src As String = BoundingBoxStockLib.GetCustomPropertyValue(partDoc, UnwrapLib.PROP_DIMENSION_SOURCE, "")
        If src = UnwrapLib.DIMENSION_SOURCE_LEHTMETALL AndAlso IsSheetMetalPart(partDoc) Then
            thicknessAxes.Add("")
            widthAxes.Add("")
            lengthAxes.Add("")
            customAxisDescs.Add("Lehtmetall")
            selectedFlags.Add(True)
            Exit Sub
        End If
        If src = UnwrapLib.DIMENSION_SOURCE_NORMAL Then
            AppendNormalPartAxisLists(partDoc, thicknessAxes, widthAxes, lengthAxes, customAxisDescs)
            selectedFlags.Add(True)
            Exit Sub
        End If
        
        ' Default: Pinnalaotus (saved source empty, Pinnalaotus, or unknown)
        Dim tAx As String = ""
        Dim wAx As String = ""
        Dim lAx As String = ""
        tAx = BoundingBoxStockLib.GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
        wAx = BoundingBoxStockLib.GetCustomPropertyValue(partDoc, "BB_WidthAxis", "")
        lAx = BoundingBoxStockLib.GetCustomPropertyValue(partDoc, "BB_LengthAxis", "")
        If BoundingBoxStockLib.IsVectorFormat(tAx) AndAlso Not BoundingBoxStockLib.IsVectorFormat(wAx) Then
            Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
            Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
            Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
            BoundingBoxStockLib.ParseVectorComponents(tAx, tx, ty, tz)
            BoundingBoxStockLib.ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
            ' Use autodetect for correct Pinnalaotus body - never fall back to all-body measurement
            Dim measBody As SurfaceBody = TryAutoDetectPinnalaotusBody(partDoc)
            Dim wExt As Double = 0, lExt As Double = 0
            If measBody IsNot Nothing Then
                wExt = UnwrapLib.GetOrientedExtentForBody(measBody, wx, wy, wz)
                lExt = UnwrapLib.GetOrientedExtentForBody(measBody, lx, ly, lz)
            Else
                ' Fall back to unwrap surface if available
                Dim unwrapFeatFB As UnwrapFeature = UnwrapLib.GetUnwrapFeature(partDoc)
                Dim unwrapSurfFB As SurfaceBody = If(unwrapFeatFB IsNot Nothing, UnwrapLib.GetUnwrappedSurfaceBody(unwrapFeatFB), Nothing)
                If unwrapSurfFB IsNot Nothing Then
                    wExt = UnwrapLib.GetOrientedExtentForBody(unwrapSurfFB, wx, wy, wz)
                    lExt = UnwrapLib.GetOrientedExtentForBody(unwrapSurfFB, lx, ly, lz)
                End If
            End If
            If lExt >= wExt Then
                wAx = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                lAx = BoundingBoxStockLib.VectorToString(lx, ly, lz)
            Else
                wAx = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                lAx = BoundingBoxStockLib.VectorToString(wx, wy, wz)
            End If
        ElseIf Not (BoundingBoxStockLib.IsVectorFormat(tAx) AndAlso BoundingBoxStockLib.IsVectorFormat(wAx)) Then
            ' Prefer unwrap flat surface plane normal as thickness; fallback = smallest-extent heuristic on measurement body
            ' Use autodetect for correct body
            Dim measBodyAD As SurfaceBody = TryAutoDetectPinnalaotusBody(partDoc)
            Dim unwrapF As UnwrapFeature = UnwrapLib.GetUnwrapFeature(partDoc)
            Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
            Dim gotUnwrapNormal As Boolean = (unwrapF IsNot Nothing AndAlso UnwrapLib.TryGetUnwrapFlatSurfaceNormal(unwrapF, nx, ny, nz))
            If gotUnwrapNormal AndAlso measBodyAD IsNot Nothing Then
                Dim tx As Double = nx, ty As Double = ny, tz As Double = nz
                Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
                Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
                BoundingBoxStockLib.ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
                Dim wExt As Double = UnwrapLib.GetOrientedExtentForBody(measBodyAD, wx, wy, wz)
                Dim lExt As Double = UnwrapLib.GetOrientedExtentForBody(measBodyAD, lx, ly, lz)
                tAx = BoundingBoxStockLib.VectorToString(tx, ty, tz)
                If lExt >= wExt Then
                    wAx = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                    lAx = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                Else
                    wAx = BoundingBoxStockLib.VectorToString(lx, ly, lz)
                    lAx = BoundingBoxStockLib.VectorToString(wx, wy, wz)
                End If
            ElseIf measBodyAD IsNot Nothing Then
                Dim tg As String = "", wg As String = "", lg As String = ""
                If BoundingBoxStockLib.AutoDetectAxesFromSurfaceBody(measBodyAD, tg, wg, lg) Then
                    tAx = BoundingBoxStockLib.PrincipalAxisToVectorString(tg)
                    wAx = BoundingBoxStockLib.PrincipalAxisToVectorString(wg)
                    lAx = BoundingBoxStockLib.PrincipalAxisToVectorString(lg)
                End If
            End If
        End If
        thicknessAxes.Add(tAx)
        widthAxes.Add(wAx)
        lengthAxes.Add(lAx)
        customAxisDescs.Add("Pinnalaotus")
        selectedFlags.Add(True)
        Exit Sub
    End If
    
    If IsSheetMetalPart(partDoc) Then
        thicknessAxes.Add("")
        widthAxes.Add("")
        lengthAxes.Add("")
        customAxisDescs.Add("Lehtmetall")
        selectedFlags.Add(True)
        Exit Sub
    End If
    
    AppendNormalPartAxisLists(partDoc, thicknessAxes, widthAxes, lengthAxes, customAxisDescs)
    selectedFlags.Add(True)
End Sub

' ============================================================================
' Recursively collect all parts from assembly occurrences
' ============================================================================
Sub CollectAllPartsFromAssembly(ByVal occurrences As ComponentOccurrences, _
                                 ByVal processedDefs As HashSet(Of Object), _
                                 ByVal partDocs As List(Of PartDocument), _
                                 ByVal partNames As List(Of String), _
                                 ByVal thicknessAxes As List(Of String), _
                                 ByVal widthAxes As List(Of String), _
                                 ByVal lengthAxes As List(Of String), _
                                 ByVal customAxisDescs As List(Of String), _
                                 ByVal selectedFlags As List(Of Boolean))
    
    For Each occ As ComponentOccurrence In occurrences
        Try
            If occ.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                ' Part occurrence - collect if not already processed (both normal and sheet metal)
                If Not processedDefs.Contains(occ.Definition) Then
                    processedDefs.Add(occ.Definition)
                    Dim partDoc As PartDocument = CType(occ.Definition.Document, PartDocument)
                    CollectPartData(partDoc, partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
                End If
            ElseIf occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ' Sub-assembly - recurse into it
                Dim subAsmDef As AssemblyComponentDefinition = CType(occ.Definition, AssemblyComponentDefinition)
                CollectAllPartsFromAssembly(subAsmDef.Occurrences, processedDefs, _
                                           partDocs, partNames, thicknessAxes, widthAxes, lengthAxes, customAxisDescs, selectedFlags)
            End If
        Catch
        End Try
    Next
End Sub

' ============================================================================
' Get part description from iProperties
' ============================================================================
Function GetPartDescription(ByVal partDoc As PartDocument) As String
    Try
        Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
        Dim desc As String = CStr(designProps.Item("Description").Value)
        If desc IsNot Nothing AndAlso desc.Trim() <> "" Then
            Return desc.Trim()
        End If
    Catch
    End Try
    
    Try
        Dim summaryInfo As PropertySet = partDoc.PropertySets.Item("Inventor Summary Information")
        Dim subj As String = CStr(summaryInfo.Item("Subject").Value)
        If subj IsNot Nothing AndAlso subj.Trim() <> "" Then
            Return subj.Trim()
        End If
    Catch
    End Try
    
    Return ""
End Function

' ============================================================================
' Check if part is a sheet metal part
' ============================================================================
Function IsSheetMetalPart(ByVal partDoc As PartDocument) As Boolean
    Try
        ' Sheet metal SubType GUID
        Const SHEET_METAL_SUBTYPE As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
        Return partDoc.SubType = SHEET_METAL_SUBTYPE
    Catch
        Return False
    End Try
End Function

' Telg readonly only for pure sheet metal (no Unwrap): flat pattern is the only mode.
Function IsTelgDropdownLocked(ByVal partDoc As PartDocument, ByVal customAxisDesc As String) As Boolean
    If customAxisDesc <> "Lehtmetall" Then Return False
    Return Not UnwrapLib.HasUnwrapFeature(partDoc)
End Function

' Only sheet metal flat-pattern mode blocks face pick and thickness-plane workflow
Function IsLehtmetallMode(ByVal customAxisDesc As String) As Boolean
    Return customAxisDesc = "Lehtmetall"
End Function

' ============================================================================
' Auto-detect or prompt for Pinnalaotus measurement body
' Similar logic to Pinnalaotuse vaated.vb - finds the correct body to measure
' ============================================================================
Function TryAutoDetectPinnalaotusBody(ByVal partDoc As PartDocument) As SurfaceBody
    ' 1. Try stored property first
    Dim resolved As SurfaceBody = UnwrapLib.ResolveManufacturedSolidBody(partDoc)
    If resolved IsNot Nothing Then Return resolved
    
    ' 2. Try Thicken output
    Dim thickBody As SurfaceBody = UnwrapLib.TryGetThickenManufacturedSolidBody(partDoc)
    If thickBody IsNot Nothing Then Return thickBody
    
    ' 3. Find first non-unwrap-surface solid body (like Pinnalaotuse vaated.vb)
    Dim unwrapFeat As UnwrapFeature = UnwrapLib.GetUnwrapFeature(partDoc)
    If unwrapFeat Is Nothing Then Return Nothing
    
    Dim unwrapSurf As SurfaceBody = UnwrapLib.GetUnwrappedSurfaceBody(unwrapFeat)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    For Each body As SurfaceBody In compDef.SurfaceBodies
        ' Skip the unwrap surface
        If unwrapSurf IsNot Nothing Then
            Try
                If ReferenceEquals(body, unwrapSurf) OrElse _
                   String.Equals(body.Name, unwrapSurf.Name, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If
            Catch
            End Try
        End If
        
        ' Check if it's a solid (has volume) - precision 1% is sufficient for this check
        Try
            If body.Volume(0.01) > 0 Then Return body
        Catch
            ' Surface bodies may not have volume, skip them
            Continue For
        End Try
    Next
    
    Return Nothing
End Function

''' <summary>
''' Show body selection dialog for Pinnalaotus when autodetect fails.
''' Returns the selected body or Nothing if cancelled.
''' </summary>
Function PromptForPinnalaotusBody(ByVal partDoc As PartDocument) As SurfaceBody
    Dim bodies As New List(Of SurfaceBody)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    Dim unwrapFeat As UnwrapFeature = UnwrapLib.GetUnwrapFeature(partDoc)
    Dim unwrapSurf As SurfaceBody = Nothing
    If unwrapFeat IsNot Nothing Then unwrapSurf = UnwrapLib.GetUnwrappedSurfaceBody(unwrapFeat)
    
    For Each b As SurfaceBody In compDef.SurfaceBodies
        bodies.Add(b)
    Next
    
    If bodies.Count = 0 Then
        MessageBox.Show("Detailis pole ühtegi keha.", "Mõõdud")
        Return Nothing
    End If
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Mõõdud — vali toodetud keha"
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.MinimizeBox = False
    frm.MaximizeBox = False
    frm.Width = 440
    frm.Height = 200
    
    Dim lbl As New System.Windows.Forms.Label()
    lbl.Left = 12
    lbl.Top = 12
    lbl.Width = 400
    lbl.Height = 48
    lbl.Text = "Pinnalaotuse keha automaatne tuvastamine ebaõnnestus." & vbCrLf &
               "Vali toodetud lame keha (Thicken või Extrude)."
    frm.Controls.Add(lbl)
    
    Dim cb As New System.Windows.Forms.ComboBox()
    cb.Left = 12
    cb.Top = 64
    cb.Width = 400
    cb.DropDownStyle = ComboBoxStyle.DropDownList
    For Each b As SurfaceBody In bodies
        cb.Items.Add(b.Name)
    Next
    frm.Controls.Add(cb)
    
    ' Default to first non-unwrap body
    Dim defaultIdx As Integer = 0
    For i As Integer = 0 To bodies.Count - 1
        If unwrapSurf Is Nothing Then Exit For
        Try
            If ReferenceEquals(bodies(i), unwrapSurf) OrElse _
               String.Equals(bodies(i).Name, unwrapSurf.Name, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If
        Catch
        End Try
        defaultIdx = i
        Exit For
    Next
    cb.SelectedIndex = Math.Min(Math.Max(0, defaultIdx), cb.Items.Count - 1)
    
    Dim btnOk As New System.Windows.Forms.Button()
    btnOk.Text = "OK"
    btnOk.DialogResult = DialogResult.OK
    btnOk.Left = 240
    btnOk.Top = 110
    btnOk.Width = 80
    frm.Controls.Add(btnOk)
    frm.AcceptButton = btnOk
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Loobu"
    btnCancel.DialogResult = DialogResult.Cancel
    btnCancel.Left = 332
    btnCancel.Top = 110
    btnCancel.Width = 80
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    If frm.ShowDialog() <> DialogResult.OK Then
        frm.Dispose()
        Return Nothing
    End If
    
    Dim selIdx As Integer = cb.SelectedIndex
    frm.Dispose()
    
    If selIdx < 0 OrElse selIdx >= bodies.Count Then Return Nothing
    
    Dim chosen As SurfaceBody = bodies(selIdx)
    
    ' Warn if user selected the unwrap surface
    If unwrapSurf IsNot Nothing Then
        Try
            If ReferenceEquals(chosen, unwrapSurf) OrElse _
               String.Equals(chosen.Name, unwrapSurf.Name, StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("Unwrap väljundi pinda ei saa toodetud kehana valida. Vali Extrude/Thicken keha.",
                                "Mõõdud")
                Return Nothing
            End If
        Catch
        End Try
    End If
    
    ' Store the selected body name
    UnwrapLib.SetManufacturedSolidBodyNameProperty(partDoc, chosen.Name)
    Logger.Info("Mõõdud: Stored Pinnalaotus body: " & chosen.Name)
    
    Return chosen
End Function

''' <summary>
''' Get Pinnalaotus measurement body with autodetect and optional user prompt.
''' Returns the body to use for measurements, or Nothing if failed/cancelled.
''' </summary>
Function GetPinnalaotusBodyForMeasurement(ByVal partDoc As PartDocument, ByVal promptIfMissing As Boolean) As SurfaceBody
    ' First try autodetect
    Dim body As SurfaceBody = TryAutoDetectPinnalaotusBody(partDoc)
    If body IsNot Nothing Then Return body
    
    ' Autodetect failed - prompt if allowed
    If promptIfMissing Then
        body = PromptForPinnalaotusBody(partDoc)
    End If
    
    Return body
End Function

' ============================================================================
' Show batch dialog with DataGridView
' ============================================================================
Function ShowBatchDialog(ByVal app As Inventor.Application, _
                         ByVal partDocs As List(Of PartDocument), _
                         ByVal partNames As List(Of String), _
                         ByVal thicknessAxes As List(Of String), _
                         ByVal widthAxes As List(Of String), _
                         ByVal lengthAxes As List(Of String), _
                         ByVal customAxisDescs As List(Of String), _
                         ByVal selectedFlags As List(Of Boolean), _
                         ByRef pickRowIndex As Integer) As DialogResult
    
    pickRowIndex = -1
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Mõõdud - Gabariitmõõtude seadistamine"
    frm.Width = 900
    frm.Height = If(partDocs.Count = 1, 220, 450)
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.Sizable
    frm.MinimizeBox = True
    frm.MaximizeBox = True
    
    Dim currentY As Integer = 10
    
    ' Header label
    Dim lblHeader As New System.Windows.Forms.Label()
    lblHeader.Text = "Detailid (" & partDocs.Count & "):"
    lblHeader.Left = 10
    lblHeader.Top = currentY
    lblHeader.Width = 200
    frm.Controls.Add(lblHeader)
    
    currentY += 20
    
    ' DataGridView
    Dim dgv As New System.Windows.Forms.DataGridView()
    dgv.Name = "dgvParts"
    dgv.Left = 10
    dgv.Top = currentY
    dgv.Width = 860
    dgv.Height = If(partDocs.Count = 1, 60, 280)
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
    
    ' Column: Part Name
    Dim colName As New DataGridViewTextBoxColumn()
    colName.Name = "colName"
    colName.HeaderText = "Detail"
    colName.Width = 250
    colName.ReadOnly = True
    dgv.Columns.Add(colName)
    
    ' Column: Thickness
    Dim colT As New DataGridViewTextBoxColumn()
    colT.Name = "colT"
    colT.HeaderText = "T (mm)"
    colT.Width = 70
    colT.ReadOnly = True
    dgv.Columns.Add(colT)
    
    ' Column: Width
    Dim colW As New DataGridViewTextBoxColumn()
    colW.Name = "colW"
    colW.HeaderText = "W (mm)"
    colW.Width = 70
    colW.ReadOnly = True
    dgv.Columns.Add(colW)
    
    ' Column: Length
    Dim colL As New DataGridViewTextBoxColumn()
    colL.Name = "colL"
    colL.HeaderText = "L (mm)"
    colL.Width = 70
    colL.ReadOnly = True
    dgv.Columns.Add(colL)
    
    ' Column: Axis (ComboBox)
    Dim colAxis As New DataGridViewComboBoxColumn()
    colAxis.Name = "colAxis"
    colAxis.HeaderText = "Telg"
    colAxis.Width = 100
    colAxis.FlatStyle = FlatStyle.Flat
    colAxis.Items.Add("X")
    colAxis.Items.Add("Y")
    colAxis.Items.Add("Z")
    colAxis.Items.Add("Kohandatud")
    colAxis.Items.Add("Lehtmetall")
    colAxis.Items.Add("Pinnalaotus")
    dgv.Columns.Add(colAxis)
    
    ' Column: Flip button
    Dim colFlip As New DataGridViewButtonColumn()
    colFlip.Name = "colFlip"
    colFlip.HeaderText = "W/L"
    colFlip.Text = "Vaheta"
    colFlip.UseColumnTextForButtonValue = True
    colFlip.Width = 70
    dgv.Columns.Add(colFlip)
    
    ' Column: Pick face button
    Dim colPick As New DataGridViewButtonColumn()
    colPick.Name = "colPick"
    colPick.HeaderText = "Pind"
    colPick.Text = "Vali pind"
    colPick.UseColumnTextForButtonValue = True
    colPick.Width = 80
    dgv.Columns.Add(colPick)
    
    ' Populate rows
    For i As Integer = 0 To partDocs.Count - 1
        Dim rowIndex As Integer = dgv.Rows.Add()
        dgv.Rows(rowIndex).Tag = i
        
        dgv.Rows(rowIndex).Cells("colSelected").Value = selectedFlags(i)
        dgv.Rows(rowIndex).Cells("colName").Value = partNames(i)
        
        ' Calculate display values
        UpdateRowDisplayValues(dgv.Rows(rowIndex), partDocs(i), thicknessAxes(i), widthAxes(i), lengthAxes(i), customAxisDescs(i))
        
        ' Set axis combo value
        If customAxisDescs(i) = "Lehtmetall" Then
            dgv.Rows(rowIndex).Cells("colAxis").Value = "Lehtmetall"
        ElseIf customAxisDescs(i) = "Pinnalaotus" Then
            dgv.Rows(rowIndex).Cells("colAxis").Value = "Pinnalaotus"
        ElseIf BoundingBoxStockLib.IsVectorFormat(thicknessAxes(i)) Then
            dgv.Rows(rowIndex).Cells("colAxis").Value = "Kohandatud"
        Else
            dgv.Rows(rowIndex).Cells("colAxis").Value = thicknessAxes(i)
        End If
        If IsTelgDropdownLocked(partDocs(i), customAxisDescs(i)) Then
            dgv.Rows(rowIndex).Cells("colAxis").ReadOnly = True
        End If
    Next
    
    ' Store form tag as -1 (no pick requested yet)
    frm.Tag = -1
    
    ' Handle button clicks
    AddHandler dgv.CellContentClick, Sub(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        
        Dim idx As Integer = CInt(dgv.Rows(e.RowIndex).Tag)
        
        If e.ColumnIndex = dgv.Columns("colFlip").Index Then
            If IsLehtmetallMode(customAxisDescs(idx)) Then Exit Sub
            If customAxisDescs(idx) = "Pinnalaotus" AndAlso Not BoundingBoxStockLib.IsVectorFormat(thicknessAxes(idx)) Then Exit Sub
            ' Flip width/length
            Dim tempAxis As String = widthAxes(idx)
            widthAxes(idx) = lengthAxes(idx)
            lengthAxes(idx) = tempAxis
            
            ' Update display
            UpdateRowDisplayValues(dgv.Rows(e.RowIndex), partDocs(idx), thicknessAxes(idx), widthAxes(idx), lengthAxes(idx), customAxisDescs(idx))
            
        ElseIf e.ColumnIndex = dgv.Columns("colPick").Index Then
            If IsLehtmetallMode(customAxisDescs(idx)) Then Exit Sub
            ' Pick face - sync state and close form
            SyncGridToLists(dgv, selectedFlags)
            frm.Tag = idx
            frm.DialogResult = DialogResult.Retry
            frm.Close()
        End If
    End Sub
    
    ' Handle axis combo change
    AddHandler dgv.CellValueChanged, Sub(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex <> dgv.Columns("colAxis").Index Then Exit Sub
        
        Dim idx As Integer = CInt(dgv.Rows(e.RowIndex).Tag)
        
        If IsTelgDropdownLocked(partDocs(idx), customAxisDescs(idx)) Then Exit Sub
        
        Dim newAxisValue As Object = dgv.Rows(e.RowIndex).Cells("colAxis").Value
        If newAxisValue Is Nothing Then Exit Sub
        
        Dim newAxis As String = newAxisValue.ToString()
        
        If newAxis = "Kohandatud" Then
            ' Trigger face pick
            SyncGridToLists(dgv, selectedFlags)
            frm.Tag = idx
            frm.DialogResult = DialogResult.Retry
            frm.Close()
        ElseIf newAxis = "Lehtmetall" Then
            If Not IsSheetMetalPart(partDocs(idx)) Then
                MessageBox.Show("Lehtmetalli mõõdud on võimalikud ainult lehtmetalli alamtüüpi detailil.", "Mõõdud")
                dgv.Rows(e.RowIndex).Cells("colAxis").Value = "Pinnalaotus"
                thicknessAxes(idx) = ""
                widthAxes(idx) = ""
                lengthAxes(idx) = ""
                customAxisDescs(idx) = "Pinnalaotus"
                UpdateRowDisplayValues(dgv.Rows(e.RowIndex), partDocs(idx), thicknessAxes(idx), widthAxes(idx), lengthAxes(idx), customAxisDescs(idx))
                Exit Sub
            End If
            thicknessAxes(idx) = ""
            widthAxes(idx) = ""
            lengthAxes(idx) = ""
            customAxisDescs(idx) = "Lehtmetall"
            UpdateRowDisplayValues(dgv.Rows(e.RowIndex), partDocs(idx), thicknessAxes(idx), widthAxes(idx), lengthAxes(idx), customAxisDescs(idx))
        ElseIf newAxis = "Pinnalaotus" Then
            thicknessAxes(idx) = ""
            widthAxes(idx) = ""
            lengthAxes(idx) = ""
            customAxisDescs(idx) = "Pinnalaotus"
            UpdateRowDisplayValues(dgv.Rows(e.RowIndex), partDocs(idx), thicknessAxes(idx), widthAxes(idx), lengthAxes(idx), customAxisDescs(idx))
        ElseIf newAxis = "X" OrElse newAxis = "Y" OrElse newAxis = "Z" Then
            ' Recalculate axes for standard axis
            thicknessAxes(idx) = newAxis
            customAxisDescs(idx) = ""
            
            Dim xSize As Double = 0, ySize As Double = 0, zSize As Double = 0
            BoundingBoxStockLib.GetBoundingBoxSizes(partDocs(idx), xSize, ySize, zSize)
            
            Dim newWidth As String = ""
            Dim newLength As String = ""
            BoundingBoxStockLib.AssignWidthLength(newAxis, xSize, ySize, zSize, newWidth, newLength)
            widthAxes(idx) = newWidth
            lengthAxes(idx) = newLength
            
            ' Update display
            UpdateRowDisplayValues(dgv.Rows(e.RowIndex), partDocs(idx), thicknessAxes(idx), widthAxes(idx), lengthAxes(idx), customAxisDescs(idx))
        End If
    End Sub
    
    ' Commit edit when cell loses focus (needed for combo box)
    AddHandler dgv.CurrentCellDirtyStateChanged, Sub(sender As Object, e As EventArgs)
        If dgv.IsCurrentCellDirty Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub
    
    frm.Controls.Add(dgv)
    
    currentY += dgv.Height + 10
    
    ' Select all / none buttons
    Dim btnSelectAll As New System.Windows.Forms.Button()
    btnSelectAll.Text = "Vali kõik"
    btnSelectAll.Left = 10
    btnSelectAll.Top = currentY
    btnSelectAll.Width = 80
    btnSelectAll.Height = 25
    btnSelectAll.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(btnSelectAll)
    
    AddHandler btnSelectAll.Click, Sub(s, e)
        For Each row As DataGridViewRow In dgv.Rows
            row.Cells("colSelected").Value = True
        Next
    End Sub
    
    Dim btnSelectNone As New System.Windows.Forms.Button()
    btnSelectNone.Text = "Tühjenda"
    btnSelectNone.Left = 95
    btnSelectNone.Top = currentY
    btnSelectNone.Width = 80
    btnSelectNone.Height = 25
    btnSelectNone.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
    frm.Controls.Add(btnSelectNone)
    
    AddHandler btnSelectNone.Click, Sub(s, e)
        For Each row As DataGridViewRow In dgv.Rows
            row.Cells("colSelected").Value = False
        Next
    End Sub
    
    ' OK/Cancel buttons
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Käivita"
    btnOK.Left = 700
    btnOK.Top = currentY
    btnOK.Width = 90
    btnOK.Height = 28
    btnOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 795
    btnCancel.Top = currentY
    btnCancel.Width = 75
    btnCancel.Height = 28
    btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Sync final state back to lists
    If result = DialogResult.OK Then
        SyncGridToLists(dgv, selectedFlags)
    End If
    
    ' Read pick index from form.Tag
    pickRowIndex = CInt(frm.Tag)
    
    frm.Dispose()
    Return result
End Function

' ============================================================================
' Update row display values (T/W/L)
' ============================================================================
Sub UpdateRowDisplayValues(ByVal row As DataGridViewRow, ByVal partDoc As PartDocument, _
                           ByVal thicknessAxis As String, ByVal widthAxis As String, ByVal lengthAxis As String, _
                           Optional ByVal customAxisDesc As String = "")
    
    Dim thicknessValue As Double = 0
    Dim widthValue As Double = 0
    Dim lengthValue As Double = 0
    
    ' Pinnalaotus only when selected (Unwrap may exist but user chose Normal or Lehtmetall)
    If customAxisDesc = "Pinnalaotus" Then
        Dim pt As Double = 0, pw As Double = 0, pl As Double = 0
        If UnwrapLib.GetPinnalaotusDimensions(partDoc, pt, pw, pl, thicknessAxis, widthAxis, lengthAxis) Then
            thicknessValue = pt
            widthValue = pw
            lengthValue = pl
        ElseIf UnwrapLib.TryGetUnwrapSurfacePreviewExtents(partDoc, pw, pl) Then
            thicknessValue = 0
            widthValue = pw
            lengthValue = pl
        End If
    ElseIf String.IsNullOrEmpty(thicknessAxis) AndAlso IsSheetMetalPart(partDoc) Then
        ' Sheet metal: get dimensions from flat pattern
        Try
            Dim smCompDef As SheetMetalComponentDefinition = CType(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
            thicknessValue = smCompDef.Thickness.Value
            
            If smCompDef.HasFlatPattern Then
                Dim fpBox As Box = smCompDef.FlatPattern.RangeBox
                Dim fpX As Double = Math.Abs(fpBox.MaxPoint.X - fpBox.MinPoint.X)
                Dim fpY As Double = Math.Abs(fpBox.MaxPoint.Y - fpBox.MinPoint.Y)
                ' Width is smaller, Length is larger
                If fpX <= fpY Then
                    widthValue = fpX
                    lengthValue = fpY
                Else
                    widthValue = fpY
                    lengthValue = fpX
                End If
            End If
        Catch
        End Try
    ElseIf BoundingBoxStockLib.IsVectorFormat(thicknessAxis) Then
        BoundingBoxStockLib.GetOrientedSizes(partDoc, thicknessAxis, widthAxis, lengthAxis, thicknessValue, widthValue, lengthValue)
    Else
        Dim xSize As Double = 0, ySize As Double = 0, zSize As Double = 0
        BoundingBoxStockLib.GetBoundingBoxSizes(partDoc, xSize, ySize, zSize)
        thicknessValue = BoundingBoxStockLib.GetAxisSize(thicknessAxis, xSize, ySize, zSize)
        widthValue = BoundingBoxStockLib.GetAxisSize(widthAxis, xSize, ySize, zSize)
        lengthValue = BoundingBoxStockLib.GetAxisSize(lengthAxis, xSize, ySize, zSize)
    End If
    
    ' Convert from cm to mm and format
    row.Cells("colT").Value = FormatMm(thicknessValue * 10)
    row.Cells("colW").Value = FormatMm(widthValue * 10)
    row.Cells("colL").Value = FormatMm(lengthValue * 10)
End Sub

' ============================================================================
' Format value in mm
' ============================================================================
Function FormatMm(ByVal valueMm As Double) As String
    Return valueMm.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)
End Function

' ============================================================================
' Sync grid state to lists
' ============================================================================
Sub SyncGridToLists(ByVal dgv As DataGridView, ByVal selectedFlags As List(Of Boolean))
    For Each row As DataGridViewRow In dgv.Rows
        Dim idx As Integer = CInt(row.Tag)
        If idx >= 0 AndAlso idx < selectedFlags.Count Then
            selectedFlags(idx) = CBool(row.Cells("colSelected").Value)
        End If
    Next
End Sub
