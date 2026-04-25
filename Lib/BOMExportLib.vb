' ============================================================================
' BOMExportLib - Export assembly BOM to Excel using template placeholders
' See project docs for template layout (auto-detected last mapping row with {{,
' assembly area above, Parts Only or Structured BOM).
' Requires: AddReference "System.Data" before AddVbFile, AddVbFile "Lib/UtilsLib.vb"
'           AddVbFile "Lib/BOMExportLib.vb"
' ============================================================================

Imports System.Collections.Generic
Imports System.Data
Imports System.Text.RegularExpressions
Imports Inventor

' Column metadata read from the mapping row (per-file — Friend by default in VB)
Friend Class BOMExportColSpec
    Public Col As Integer
    Public IsExcelFormula As Boolean
    Public R1C1 As String
    Public TextTemplate As String
End Class

Friend Class BOMDrawingInfo
    Public Found As Boolean = False
    Public FullPath As String = ""
    Public FileName As String = ""
    Public PartNumber As String = ""
    Public Description As String = ""
End Class

Public Class ExportConfig
    Public Name As String = ""
    Public ModelState As String = ""
    Public TemplatePath As String = ""
    Public OutputPath As String = ""
End Class

Public Module BOMExportLib

    Public Const MAPPING_ROW_MIN_PLACEHOLDERS As Integer = 2
    Public Const kDesignProps As String = "Design Tracking Properties"
    Public Const kUserProps As String = "Inventor User Defined Properties"
    Public Const kSummaryProps As String = "Inventor Summary Information"
    Public Const kDocSummaryProps As String = "Inventor Document Summary Information"
    Private Const kDrawAssocPartNumber As String = "BB_SourcePartNumber"
    Private Const kDrawAssocType As String = "BB_DrawingType"
    Private Const kDrawAssocType1to1 As String = "1:1"
    Private m_DrawingCache As New Dictionary(Of String, BOMDrawingInfo)(StringComparer.OrdinalIgnoreCase)
    Public Const VERBOSE_LOGGING As Boolean = True

#Region "Public API"

    Public Sub ExportWithDialog(asmDoc As AssemblyDocument, Optional usePartsOnlyBomView As Boolean = False)
        Dim ofd As New System.Windows.Forms.OpenFileDialog()
        ofd.Title = "Vali Excel mall"
        ofd.Filter = "Excel|*.xlsx;*.xls;*.xlsm|Kõik failid|*.*"
        ofd.InitialDirectory = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
        If ofd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return

        Dim sfd As New System.Windows.Forms.SaveFileDialog()
        sfd.Title = "Salvesta tekitatud BOM"
        sfd.Filter = "Excel|*.xlsx;*.xls;*.xlsm"
        sfd.InitialDirectory = ofd.InitialDirectory
        sfd.FileName = System.IO.Path.GetFileNameWithoutExtension(asmDoc.DisplayName) & "_BOM.xlsx"
        If sfd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return

        InternalExport(asmDoc, ofd.FileName, sfd.FileName, usePartsOnlyBomView, Nothing, Nothing, Nothing)
    End Sub

    Public Sub ExportBOM(asmDoc As AssemblyDocument, templatePath As String, outputPath As String, Optional usePartsOnlyBomView As Boolean = False)
        InternalExport(asmDoc, templatePath, outputPath, usePartsOnlyBomView, Nothing, Nothing, Nothing)
    End Sub

    Public Sub ExportBatch(asmDoc As AssemblyDocument, items As List(Of ExportConfig), Optional usePartsOnlyBomView As Boolean = False)
        If items Is Nothing Then Return
        For Each ex As ExportConfig In items
            If ex Is Nothing OrElse String.IsNullOrEmpty(ex.TemplatePath) Then Continue For
            Dim outP As String = ex.OutputPath
            If String.IsNullOrEmpty(outP) Then
                Dim d As String = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
                outP = System.IO.Path.Combine(d, System.IO.Path.GetFileNameWithoutExtension(asmDoc.DisplayName) & If(String.IsNullOrEmpty(ex.Name), "", "_" & ex.Name) & ".xlsx")
            End If
            InternalExport(asmDoc, ex.TemplatePath, outP, usePartsOnlyBomView, ex.ModelState, Nothing, Nothing)
        Next
    End Sub

#End Region

#Region "Internal: main"

    Private Sub AddLog(m As String, logLines As List(Of String), logger As Object)
        UtilsLib.LogInfo(m)
    End Sub

    Private Sub VLog(m As String, logLines As List(Of String), logger As Object)
        If VERBOSE_LOGGING Then
            AddLog(m, logLines, logger)
        End If
    End Sub

    Private Function LogValue(v As Object) As String
        If v Is Nothing Then Return "<Nothing>"
        Dim s As String = Convert.ToString(v)
        If s Is Nothing Then s = ""
        s = s.Replace(vbCr, "\r").Replace(vbLf, "\n")
        If s.Length > 140 Then
            s = s.Substring(0, 140) & "..."
        End If
        Return s
    End Function

    Private Function BuildContextLabel(isAssembly As Boolean, bom As BOMRow, oDoc As Document) As String
        If isAssembly Then Return "AssemblyHeader"
        Dim item As String = ""
        Try : item = CStr(bom.ItemNumber) : Catch : item = "?" : End Try
        Dim docName As String = ""
        Try
            If oDoc IsNot Nothing Then docName = oDoc.DisplayName
        Catch
        End Try
        Return "BOMItem=" & item & If(String.IsNullOrEmpty(docName), "", ", Doc=" & docName)
    End Function

    Private Sub InternalExport(
        asmDoc As AssemblyDocument, templatePath As String, outputPath As String, usePartsOnlyBomView As Boolean,
        primaryModelState As String, logLines As List(Of String), logger As Object
    )
        If asmDoc Is Nothing OrElse String.IsNullOrEmpty(templatePath) OrElse String.IsNullOrEmpty(outputPath) Then Return
        VLog("BOMExport: Start export for '" & asmDoc.DisplayName & "'", logLines, logger)
        VLog("BOMExport: Template='" & templatePath & "'", logLines, logger)
        VLog("BOMExport: Output='" & outputPath & "'", logLines, logger)
        VLog("BOMExport: ViewMode=" & If(usePartsOnlyBomView, "PartsOnly", "Structured"), logLines, logger)
        If Not System.IO.File.Exists(templatePath) Then
            AddLog("BOMExport: Template not found: " & templatePath, logLines, logger)
            Return
        End If
        If Not String.IsNullOrEmpty(primaryModelState) Then
            ActivateModelState(asmDoc, primaryModelState, logLines, logger)
        End If
        m_DrawingCache.Clear()

        Dim oView As BOMView = Nothing
        If Not GetBomView(asmDoc, usePartsOnlyBomView, oView) Then
            AddLog("BOMExport: Could not obtain BOM view", logLines, logger)
        End If

        Dim dataRows As New List(Of BOMRow)()
        If oView IsNot Nothing Then
            If usePartsOnlyBomView Then
                CollectBomRowsFlat(oView, dataRows)
            Else
                If oView.BOMRows IsNot Nothing Then
                    For Each oRow As BOMRow In oView.BOMRows
                        CollectBomRowsRecursive(oRow, dataRows)
                    Next
                End If
            End If
        End If
        VLog("BOMExport: Collected BOM rows = " & dataRows.Count, logLines, logger)

        Dim xls As Object = Nothing
        Dim wb As Object = Nothing
        Try
            xls = CreateObject("Excel.Application")
            CType(xls, Object).Visible = False
            CType(xls, Object).DisplayAlerts = False
            wb = CType(xls, Object).Workbooks.Open(templatePath, 0, False)
            Dim sh As Object = Nothing
            Try
                sh = CType(wb, Object).Sheets(1)
                ProcessWorksheet(asmDoc, CType(sh, Object), dataRows, logLines, logger)
            Finally
                If sh IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sh)
                End If
            End Try

            If System.IO.File.Exists(outputPath) Then
                System.IO.File.Delete(outputPath)
            End If
            SaveWorkbookAs(CType(wb, Object), outputPath, logLines, logger)
            AddLog("BOMExport: Wrote " & dataRows.Count & " row(s) to " & outputPath, logLines, logger)
        Catch ex As Exception
            AddLog("BOMExport: " & ex.Message, logLines, logger)
            Throw
        Finally
            If wb IsNot Nothing Then
                Try
                    CType(wb, Object).Close(SaveChanges:=False)
                Catch
                End Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb)
            End If
            If xls IsNot Nothing Then
                Try
                    CType(xls, Object).Quit()
                Catch
                End Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xls)
            End If
        End Try
    End Sub

    Private Sub SaveWorkbookAs(wb As Object, outputPath As String, logLines As List(Of String), logger As Object)
        Dim ext As String = System.IO.Path.GetExtension(outputPath).ToLowerInvariant()
        Dim fmt As Integer = 51
        If ext = ".xlsm" Then fmt = 52
        If ext = ".xls" Then fmt = 56
        Try
            CType(wb, Object).SaveAs(Filename := outputPath, FileFormat := fmt)
        Catch
            Try
                CType(wb, Object).SaveAs(Filename := outputPath)
            Catch ex As Exception
                AddLog("BOMExport: SaveAs failed: " & ex.Message, logLines, logger)
                Throw
            End Try
        End Try
    End Sub

#End Region

#Region "Model state & BOM view"

    Private Sub ActivateModelState(asmDoc As AssemblyDocument, stateName As String, logLines As List(Of String), logger As Object)
        Try
            Dim ms As Object = asmDoc.ComponentDefinition.ModelStates
            Dim s As Object = FindModelStateByName(ms, stateName)
            If s IsNot Nothing Then
                s.Activate()
                AddLog("BOMExport: Activated model state: " & stateName, logLines, logger)
            Else
                AddLog("BOMExport: Model state not found: " & stateName, logLines, logger)
            End If
        Catch ex As Exception
            AddLog("BOMExport: Model state error: " & ex.Message, logLines, logger)
        End Try
    End Sub

    Private Function FindModelStateByName(ms As Object, name As String) As Object
        If ms Is Nothing OrElse String.IsNullOrEmpty(name) Then Return Nothing
        For Each s As Object In ms
            If s IsNot Nothing AndAlso CStr(s.Name) = name Then Return s
        Next
        Return Nothing
    End Function

    Private Function GetBomView(asmDoc As AssemblyDocument, usePartsOnly As Boolean, ByRef oView As BOMView) As Boolean
        oView = Nothing
        Try
            Dim bom As BOM = asmDoc.ComponentDefinition.BOM
            If Not usePartsOnly Then
                ' Preferred path: Structured BOM
                Try
                    bom.StructuredViewEnabled = True
                    bom.StructuredViewFirstLevelOnly = False
                    oView = bom.BOMViews.Item("Structured")
                    If oView IsNot Nothing Then
                        VLog("BOMExport: Using BOM view 'Structured' (all levels)", Nothing, Nothing)
                        Return True
                    End If
                Catch exStructured As Exception
                    VLog("BOMExport: Structured BOM unavailable, falling back to Parts Only. Reason: " & exStructured.Message, Nothing, Nothing)
                End Try
            End If

            ' Fallback path: Parts Only BOM
            Try : bom.PartsOnlyViewEnabled = True : Catch : End Try
            oView = bom.BOMViews.Item("Parts Only")
            If oView IsNot Nothing Then
                VLog("BOMExport: Using BOM view 'Parts Only' (fallback)", Nothing, Nothing)
                Return True
            End If
            Return oView IsNot Nothing
        Catch
            Return False
        End Try
    End Function

#End Region

#Region "BOM rows"

    Private Sub CollectBomRowsFlat(view As BOMView, list As List(Of BOMRow))
        If view Is Nothing OrElse view.BOMRows Is Nothing Then Return
        For Each oRow As BOMRow In view.BOMRows
            list.Add(oRow)
        Next
    End Sub

    Private Sub CollectBomRowsRecursive(oRow As BOMRow, list As List(Of BOMRow))
        If oRow Is Nothing Then Return
        list.Add(oRow)
        If oRow.ChildRows Is Nothing Then Return
        For Each c As BOMRow In oRow.ChildRows
            CollectBomRowsRecursive(c, list)
        Next
    End Sub

    Private Function GetBomRowDocument(b As BOMRow, logLines As List(Of String), logger As Object) As Document
        Try
            If b Is Nothing OrElse b.ComponentDefinitions Is Nothing OrElse b.ComponentDefinitions.Count = 0 Then
                Return Nothing
            End If
            Dim d As Document = b.ComponentDefinitions.Item(1).Document
            VLog("BOMExport: Row doc resolved: Item=" & LogValue(b.ItemNumber) & ", Doc=" & LogValue(d.DisplayName), logLines, logger)
            Return d
        Catch
        End Try
        Return Nothing
    End Function

#End Region

#Region "Excel worksheet"

    Private Sub ProcessWorksheet(asmDoc As AssemblyDocument, sheet As Object, bomList As List(Of BOMRow), logLines As List(Of String), logger As Object)
        Dim used As Object = sheet.UsedRange
        If used Is Nothing Then
            AddLog("BOMExport: Empty sheet", logLines, logger)
            Return
        End If

        Dim uRow As Integer = 1
        Dim uCol As Integer = 1
        Try
            uRow = CInt(used.Row) : uCol = CInt(used.Column)
        Catch
        End Try
        Dim nRows As Integer = 1
        Dim nCols As Integer = 1
        Try
            nRows = CInt(used.Rows.Count) : nCols = CInt(used.Columns.Count)
        Catch
        End Try
        Dim lastR As Integer = uRow + nRows - 1
        Dim lastC As Integer = uCol + nCols - 1
        If lastC < 1 Then lastC = 1
        If lastR < 1 Then lastR = 1
        VLog("BOMExport: Worksheet used range rows " & uRow & "-" & lastR & ", cols " & uCol & "-" & lastC, logLines, logger)

        Dim mapRow As Integer = FindLastMappingRow(sheet, uRow, lastR, uCol, lastC, MAPPING_ROW_MIN_PLACEHOLDERS)
        If mapRow < 0 Then
            AddLog("BOMExport: No mapping row (2+ {{ cells) found", logLines, logger)
            Return
        End If
        If mapRow <= 1 Then
            AddLog("BOMExport: Invalid mapping on row 1", logLines, logger)
        End If
        VLog("BOMExport: Mapping row detected at row " & mapRow, logLines, logger)
        VLog("BOMExport: Header row assumed at row " & (mapRow - 1), logLines, logger)

        Dim colSpecs As New List(Of BOMExportColSpec)()
        For c As Integer = uCol To lastC
            Dim cs As New BOMExportColSpec() With { .Col = c }
            Dim cell As Object = sheet.Cells(mapRow, c)
            Try
                Dim a1 As String = ""
                Try
                    a1 = CStr(cell.Formula)
                Catch
                End Try
                Dim t As String = a1.Trim()
                If t.StartsWith("=") AndAlso Not t.Contains("{{") Then
                    cs.IsExcelFormula = True
                    Try
                        cs.R1C1 = CStr(cell.FormulaR1C1)
                    Catch
                    End Try
                    VLog("BOMExport: Col " & c & " = Excel formula mapping, R1C1='" & LogValue(cs.R1C1) & "'", logLines, logger)
                Else
                    Try
                        cs.TextTemplate = CStr(cell.Value2)
                    Catch
                        cs.TextTemplate = t
                    End Try
                    VLog("BOMExport: Col " & c & " = Template mapping '" & LogValue(cs.TextTemplate) & "'", logLines, logger)
                End If
            Catch
            End Try
            colSpecs.Add(cs)
        Next

        If mapRow > 1 Then
            ReplaceInRectangle(asmDoc, sheet, 1, mapRow - 1, uCol, lastC, isAssembly := True, Nothing, Nothing, logLines, logger)
        End If

        Dim n As Integer = 0
        If bomList IsNot Nothing Then n = bomList.Count
        If n = 0 Then
            For Each cs As BOMExportColSpec In colSpecs
                If Not cs.IsExcelFormula Then
                    Try
                        CType(sheet.Cells(mapRow, cs.Col), Object).Value2 = ""
                    Catch
                    End Try
                End If
            Next
            AddLog("BOMExport: No BOM rows; cleared data mapping row", logLines, logger)
            Return
        End If

        If n > 1 Then
            InsertEntireRows(sheet, mapRow + 1, n - 1)
            VLog("BOMExport: Inserted " & (n - 1) & " row(s) below mapping row", logLines, logger)
        End If

        For r As Integer = 0 To n - 1
            Dim bRow As BOMRow = bomList(r)
            Dim exRow As Integer = mapRow + r
            Dim oDoc As Document = GetBomRowDocument(bRow, logLines, logger)
            VLog("BOMExport: Writing Excel row " & exRow & " for " & BuildContextLabel(False, bRow, oDoc), logLines, logger)
            For Each cs As BOMExportColSpec In colSpecs
                If cs.IsExcelFormula Then
                    If Not String.IsNullOrEmpty(cs.R1C1) Then
                        Try
                            CType(sheet.Cells(exRow, cs.Col), Object).FormulaR1C1 = cs.R1C1
                            VLog("BOMExport:   Col " & cs.Col & " formula applied", logLines, logger)
                        Catch
                        End Try
                    End If
                Else
                    Dim sOut As String = ProcessTemplateText(cs.TextTemplate, asmDoc, bRow, oDoc, isAssembly := False, logLines, logger)
                    Try
                        SetCellStringValue(CType(sheet.Cells(exRow, cs.Col), Object), sOut, logLines, logger, "Data r" & exRow & " c" & cs.Col)
                        VLog("BOMExport:   Col " & cs.Col & " value='" & LogValue(sOut) & "'", logLines, logger)
                    Catch
                    End Try
                End If
            Next
        Next
    End Sub

    Private Sub InsertEntireRows(sheet As Object, atRow As Integer, count As Integer)
        If count < 1 Then Return
        Try
            CType(sheet, Object).Rows(CStr(atRow) & ":" & CStr(atRow + count - 1)).EntireRow.Insert()
        Catch
            For i As Integer = 0 To count - 1
                Try
                    CType(sheet, Object).Rows(CStr(atRow + i)).EntireRow.Insert()
                Catch
                End Try
            Next
        End Try
    End Sub

    Private Function FindLastMappingRow(sheet As Object, r0 As Integer, r1 As Integer, c0 As Integer, c1 As Integer, minPh As Integer) As Integer
        Dim lastMap As Integer = -1
        For r As Integer = r0 To r1
            Dim n As Integer = CountPlaceholderCellsInRow(sheet, r, c0, c1)
            If n >= minPh Then
                lastMap = r
            End If
        Next
        Return lastMap
    End Function

    Private Function CountPlaceholderCellsInRow(sheet As Object, r As Integer, c0 As Integer, c1 As Integer) As Integer
        Dim n As Integer = 0
        For c As Integer = c0 To c1
            Try
                Dim s As String = CStr(CType(sheet, Object).Cells(r, c).Value2)
                If s IsNot Nothing AndAlso s.Contains("{{") Then n += 1
            Catch
            End Try
        Next
        Return n
    End Function

#End Region

#Region "Assembly area replace"

    Private Sub ReplaceInRectangle(
        asmDoc As AssemblyDocument, sheet As Object, r0 As Integer, r1 As Integer, c0 As Integer, c1 As Integer,
        isAssembly As Boolean, bom As BOMRow, oDoc As Document, logLines As List(Of String), logger As Object
    )
        VLog("BOMExport: Replacing assembly placeholders in rectangle r" & r0 & "-r" & r1 & ", c" & c0 & "-c" & c1, logLines, logger)
        For r As Integer = r0 To r1
            For c As Integer = c0 To c1
                Try
                    Dim cell As Object = CType(sheet, Object).Cells(r, c)
                    Dim raw As String = CStr(cell.Value2)
                    If String.IsNullOrEmpty(raw) OrElse Not raw.Contains("{{") Then Continue For
                    Dim replaced As String = ProcessTemplateText(raw, asmDoc, bom, oDoc, isAssembly, logLines, logger)
                    SetCellStringValue(cell, replaced, logLines, logger, "Header r" & r & " c" & c)
                    VLog("BOMExport: Header cell (" & r & "," & c & ") '" & LogValue(raw) & "' -> '" & LogValue(replaced) & "'", logLines, logger)
                Catch
                End Try
            Next
        Next
    End Sub

#End Region

#Region "Template resolution"

    Private Sub SetCellStringValue(cell As Object, value As String, logLines As List(Of String), logger As Object, context As String)
        If value Is Nothing Then value = ""
        If ShouldForceText(value) Then
            Try
                cell.NumberFormat = "@"
            Catch
            End Try
            cell.Value2 = value
            VLog("BOMExport: " & context & " forced TEXT format for leading-zero value '" & LogValue(value) & "'", logLines, logger)
        Else
            cell.Value2 = value
        End If
    End Sub

    Private Function ShouldForceText(value As String) As Boolean
        If String.IsNullOrEmpty(value) Then Return False
        ' Preserve identifiers like 0000123 as literal strings in Excel.
        Return Regex.IsMatch(value, "^0[0-9]+$")
    End Function

    Private Function PlaceholderRegex() As Regex
        Static m As Regex = Nothing
        If m Is Nothing Then m = New Regex("\{\{([^}]+)\}\}", RegexOptions.Compiled)
        Return m
    End Function

    Private Function ProcessTemplateText(
        text As String, asmDoc As AssemblyDocument, bom As BOMRow, oDoc As Document, isAssembly As Boolean,
        logLines As List(Of String), logger As Object) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Dim s As String = text
        Dim ctx As String = BuildContextLabel(isAssembly, bom, oDoc)
        Dim m As Match = PlaceholderRegex().Match(s)
        Do While m.Success
            Dim inner0 As String = m.Groups(1).Value.Trim()
            Dim obj As Object = ResolveToken(inner0, asmDoc, bom, oDoc, isAssembly, logLines, logger)
            Dim rep As String = ObjectToString(obj)
            VLog("BOMExport: Resolve token '{{" & inner0 & "}}' -> '" & LogValue(rep) & "' [" & ctx & "]", logLines, logger)
            s = s.Substring(0, m.Index) & rep & s.Substring(m.Index + m.Length)
            m = PlaceholderRegex().Match(s)
        Loop
        Return s
    End Function

    Private Function ObjectToString(obj As Object) As String
        If obj Is Nothing Then Return ""
        If TypeOf obj Is System.DateTime Then
            Return CType(obj, System.DateTime).ToString("yyyy-MM-dd")
        End If
        Return obj.ToString()
    End Function

    Private Function ResolveToken(
        inner As String, asmDoc As AssemblyDocument, bom As BOMRow, oDoc As Document, isAssembly As Boolean,
        logLines As List(Of String), logger As Object) As Object
        Dim ctxDoc As Document = If(oDoc IsNot Nothing, oDoc, CType(asmDoc, Document))

        If inner.Trim().StartsWith("=") Then
            Dim e As String = inner.Trim().Substring(1).Trim()
            Return EvaluateExpressionString(e, asmDoc, bom, oDoc, isAssembly, logLines, logger)
        End If

        If inner.Equals("File", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("File.", StringComparison.OrdinalIgnoreCase) Then
            Return GetFileValue(inner, ctxDoc, logLines, logger)
        End If

        If inner.Equals("Drawing", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("Drawing.", StringComparison.OrdinalIgnoreCase) Then
            Return GetDrawingValue(inner, asmDoc, ctxDoc, logLines, logger)
        End If

        If inner.Equals("Summary", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("Summary.", StringComparison.OrdinalIgnoreCase) Then
            Dim p As String = inner
            If p.StartsWith("Summary.", StringComparison.OrdinalIgnoreCase) Then p = p.Substring(8).Trim()
            Return GetPropertyString(ctxDoc, p, kSummaryProps)
        End If

        If inner.Equals("DocSummary", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("DocSummary.", StringComparison.OrdinalIgnoreCase) Then
            Dim p As String = inner
            If p.StartsWith("DocSummary.", StringComparison.OrdinalIgnoreCase) Then p = p.Substring(11).Trim()
            Return GetPropertyString(ctxDoc, p, kDocSummaryProps)
        End If

        If inner.Equals("Vault", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("Vault.", StringComparison.OrdinalIgnoreCase) Then
            Dim p As String = inner
            If p.StartsWith("Vault.", StringComparison.OrdinalIgnoreCase) Then p = p.Substring(6).Trim()
            Return GetVaultLikeValue(p, ctxDoc, logLines, logger)
        End If

        If isAssembly AndAlso (inner.Equals("BOM", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("BOM.", StringComparison.OrdinalIgnoreCase)) Then
            Return ""
        End If
        If inner.Equals("BOM", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("BOM.", StringComparison.OrdinalIgnoreCase) Then
            Return GetBomValue(inner, bom, logLines, logger)
        End If

        If inner.Equals("Custom", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("Custom.", StringComparison.OrdinalIgnoreCase) Then
            Dim p As String = inner
            If p.Length > 7 AndAlso p.StartsWith("Custom.", StringComparison.OrdinalIgnoreCase) Then
                p = p.Substring(7).Trim()
            End If
            If isAssembly OrElse oDoc Is Nothing Then
                Return GetPropertyString(CType(asmDoc, Document), p, kUserProps)
            End If
            Return GetPropertyString(oDoc, p, kUserProps)
        End If

        If inner.Equals("Phys", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("Phys.", StringComparison.OrdinalIgnoreCase) OrElse
           inner.Equals("Physical", StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("Physical.", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals("Phy", inner, StringComparison.OrdinalIgnoreCase) OrElse inner.StartsWith("Phy.", StringComparison.OrdinalIgnoreCase) Then
            Dim p As String = inner
            If p.StartsWith("Physical.", StringComparison.OrdinalIgnoreCase) AndAlso p.Length > 9 Then
                p = p.Substring(9)
            ElseIf p.StartsWith("Phys.", StringComparison.OrdinalIgnoreCase) AndAlso p.Length > 5 Then
                p = p.Substring(5)
            ElseIf p.StartsWith("Phy.", StringComparison.OrdinalIgnoreCase) AndAlso p.Length > 4 Then
                p = p.Substring(4)
            End If
            p = p.Trim()
            If oDoc IsNot Nothing AndAlso oDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Return GetPhysical(CType(CType(oDoc, PartDocument).ComponentDefinition, PartComponentDefinition), p, logLines, logger)
            End If
            If oDoc IsNot Nothing AndAlso oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Return GetPhysicalFromAsm(CType(oDoc, AssemblyDocument), p, logLines, logger)
            End If
        End If

        If isAssembly OrElse oDoc Is Nothing Then
            Return GetPropertyString(CType(asmDoc, Document), inner, kDesignProps)
        End If
        Return GetPropertyString(oDoc, inner, kDesignProps)
    End Function

    Private Function GetFileValue(inner As String, oDoc As Document, logLines As List(Of String), logger As Object) As Object
        If oDoc Is Nothing Then Return ""
        Dim key As String = inner
        If key.StartsWith("File.", StringComparison.OrdinalIgnoreCase) Then
            key = key.Substring(5).Trim()
        Else
            key = ""
        End If
        If String.IsNullOrEmpty(key) Then Return ""

        Dim fullPath As String = ""
        Try
            fullPath = oDoc.FullFileName
        Catch
        End Try
        If String.IsNullOrEmpty(fullPath) OrElse Not System.IO.File.Exists(fullPath) Then
            VLog("BOMExport: File lookup '" & inner & "' -> <file missing>", logLines, logger)
            Return ""
        End If

        Select Case key.ToLowerInvariant()
            Case "modified", "modifieddate", "lastwrite", "lastwritetime", "lastwriteutc"
                Dim dt As DateTime = System.IO.File.GetLastWriteTime(fullPath)
                VLog("BOMExport: File lookup '" & inner & "' -> " & dt.ToString("yyyy-MM-dd HH:mm:ss"), logLines, logger)
                Return dt
            Case "created", "createddate", "creationtime", "creationdate"
                Dim dt As DateTime = System.IO.File.GetCreationTime(fullPath)
                VLog("BOMExport: File lookup '" & inner & "' -> " & dt.ToString("yyyy-MM-dd HH:mm:ss"), logLines, logger)
                Return dt
            Case "name", "filename"
                Dim v As String = System.IO.Path.GetFileName(fullPath)
                VLog("BOMExport: File lookup '" & inner & "' -> '" & v & "'", logLines, logger)
                Return v
            Case "path", "fullpath"
                VLog("BOMExport: File lookup '" & inner & "' -> '" & fullPath & "'", logLines, logger)
                Return fullPath
            Case Else
                VLog("BOMExport: File lookup '" & inner & "' -> <unsupported>", logLines, logger)
                Return ""
        End Select
    End Function

    Private Function GetDrawingValue(inner As String, asmDoc As AssemblyDocument, srcDoc As Document, logLines As List(Of String), logger As Object) As Object
        If srcDoc Is Nothing Then Return ""

        Dim partNumber As String = GetPropertyString(srcDoc, "Part Number", kDesignProps)
        If String.IsNullOrEmpty(partNumber) Then
            VLog("BOMExport: Drawing lookup skipped - source document has no Part Number", logLines, logger)
            Return ""
        End If

        Dim info As BOMDrawingInfo = GetDrawingInfoForPart(partNumber, asmDoc, srcDoc, logLines, logger)
        If info Is Nothing OrElse Not info.Found Then
            VLog("BOMExport: Drawing lookup '" & inner & "' -> <not found> for part " & partNumber, logLines, logger)
            Return ""
        End If

        Dim key As String = inner
        If key.StartsWith("Drawing.", StringComparison.OrdinalIgnoreCase) Then
            key = key.Substring(8).Trim()
        Else
            key = ""
        End If

        Select Case key.ToLowerInvariant()
            Case "filename", "name", "file"
                Return info.FileName
            Case "path", "fullpath"
                Return info.FullPath
            Case "description"
                Return info.Description
            Case "partnumber", "part"
                Return info.PartNumber
            Case Else
                Return ""
        End Select
    End Function

    Private Function GetDrawingInfoForPart(partNumber As String, asmDoc As AssemblyDocument, srcDoc As Document, logLines As List(Of String), logger As Object) As BOMDrawingInfo
        If String.IsNullOrEmpty(partNumber) Then Return New BOMDrawingInfo()
        If m_DrawingCache.ContainsKey(partNumber) Then
            Return m_DrawingCache(partNumber)
        End If

        Dim info As BOMDrawingInfo = FindDrawingInfo(partNumber, asmDoc, srcDoc, logLines, logger)
        m_DrawingCache(partNumber) = info
        Return info
    End Function

    Private Function FindDrawingInfo(partNumber As String, asmDoc As AssemblyDocument, srcDoc As Document, logLines As List(Of String), logger As Object) As BOMDrawingInfo
        Dim result As New BOMDrawingInfo()
        Try
            ' 1) Check open drawings first (fast path)
            Dim app As Inventor.Application = asmDoc.Parent
            For Each d As Document In app.Documents
                If d.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    Dim dd As DrawingDocument = CType(d, DrawingDocument)
                    If Is1to1DrawingForPart(dd, partNumber) Then
                        FillDrawingInfo(dd, result)
                        result.Found = True
                        VLog("BOMExport: Found open 1:1 drawing for part " & partNumber & " -> " & result.FileName, logLines, logger)
                        Return result
                    End If
                End If
            Next

            ' 2) Search from part folder up to assembly folder boundary (depth-first)
            Dim startFolder As String = ""
            Dim limitFolder As String = ""
            Try : startFolder = System.IO.Path.GetDirectoryName(srcDoc.FullFileName) : Catch : End Try
            Try : limitFolder = System.IO.Path.GetDirectoryName(asmDoc.FullFileName) : Catch : End Try
            If String.IsNullOrEmpty(startFolder) OrElse Not System.IO.Directory.Exists(startFolder) Then
                startFolder = limitFolder
            End If
            If String.IsNullOrEmpty(limitFolder) OrElse Not System.IO.Directory.Exists(limitFolder) Then
                limitFolder = startFolder
            End If
            If String.IsNullOrEmpty(startFolder) OrElse String.IsNullOrEmpty(limitFolder) Then
                Return result
            End If

            Dim foundPath As String = FindDrawingPathOnDisk(partNumber, app, startFolder, limitFolder, logLines, logger)
            If Not String.IsNullOrEmpty(foundPath) Then
                Dim drawDoc As DrawingDocument = Nothing
                Dim openedByUs As Boolean = False
                Try
                    drawDoc = OpenOrGetDrawingDoc(app, foundPath, openedByUs)
                    If drawDoc IsNot Nothing Then
                        FillDrawingInfo(drawDoc, result)
                        result.Found = True
                        VLog("BOMExport: Found disk 1:1 drawing for part " & partNumber & " -> " & result.FileName, logLines, logger)
                    End If
                Finally
                    If openedByUs AndAlso drawDoc IsNot Nothing Then
                        Try : drawDoc.Close(True) : Catch : End Try
                    End If
                End Try
            End If
        Catch ex As Exception
            VLog("BOMExport: Drawing lookup error for part " & partNumber & ": " & ex.Message, logLines, logger)
        End Try
        Return result
    End Function

    Private Function FindDrawingPathOnDisk(partNumber As String, app As Inventor.Application, startFolder As String, limitFolder As String, logLines As List(Of String), logger As Object) As String
        Dim stack As New Stack(Of String)()
        Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        stack.Push(startFolder)

        While stack.Count > 0
            Dim dir As String = stack.Pop()
            If String.IsNullOrEmpty(dir) OrElse visited.Contains(dir) Then Continue While
            visited.Add(dir)

            ' Stay within boundary if possible
            If Not IsWithinBoundary(dir, limitFolder) Then Continue While

            Dim files() As String = {}
            Try
                files = System.IO.Directory.GetFiles(dir, "*.idw")
            Catch
            End Try
            For Each fp As String In files
                Dim drawDoc As DrawingDocument = Nothing
                Dim openedByUs As Boolean = False
                Try
                    drawDoc = OpenOrGetDrawingDoc(app, fp, openedByUs)
                    If drawDoc IsNot Nothing AndAlso Is1to1DrawingForPart(drawDoc, partNumber) Then
                        Return fp
                    End If
                Catch
                Finally
                    If openedByUs AndAlso drawDoc IsNot Nothing Then
                        Try : drawDoc.Close(True) : Catch : End Try
                    End If
                End Try
            Next

            Dim subDirs() As String = {}
            Try
                subDirs = System.IO.Directory.GetDirectories(dir)
            Catch
            End Try
            For Each sd As String In subDirs
                If sd.IndexOf("\OldVersions\", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
                If IsWithinBoundary(sd, limitFolder) Then
                    stack.Push(sd)
                End If
            Next

            ' Walk upward toward limit boundary to mimic depth-first-with-boundary behavior.
            Dim parent As String = ""
            Try : parent = System.IO.Directory.GetParent(dir).FullName : Catch : End Try
            If Not String.IsNullOrEmpty(parent) AndAlso IsWithinBoundary(parent, limitFolder) Then
                If Not visited.Contains(parent) Then stack.Push(parent)
            End If
        End While

        Return ""
    End Function

    Private Function IsWithinBoundary(path As String, boundaryRoot As String) As Boolean
        If String.IsNullOrEmpty(path) OrElse String.IsNullOrEmpty(boundaryRoot) Then Return True
        Dim p As String = path.TrimEnd("\"c).ToLowerInvariant()
        Dim b As String = boundaryRoot.TrimEnd("\"c).ToLowerInvariant()
        Return p.StartsWith(b)
    End Function

    Private Function OpenOrGetDrawingDoc(app As Inventor.Application, fullPath As String, ByRef openedByUs As Boolean) As DrawingDocument
        openedByUs = False
        For Each d As Document In app.Documents
            Try
                If d.DocumentType = DocumentTypeEnum.kDrawingDocumentObject AndAlso d.FullDocumentName.Equals(fullPath, StringComparison.OrdinalIgnoreCase) Then
                    Return CType(d, DrawingDocument)
                End If
            Catch
            End Try
        Next
        Dim o As Document = app.Documents.Open(fullPath, False)
        openedByUs = True
        If o Is Nothing Then Return Nothing
        If o.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
            Try : o.Close(True) : Catch : End Try
            openedByUs = False
            Return Nothing
        End If
        Return CType(o, DrawingDocument)
    End Function

    Private Function Is1to1DrawingForPart(drawDoc As DrawingDocument, partNumber As String) As Boolean
        If drawDoc Is Nothing OrElse String.IsNullOrEmpty(partNumber) Then Return False
        Dim storedPart As String = GetPropertyString(CType(drawDoc, Document), kDrawAssocPartNumber, kUserProps)
        If Not storedPart.Equals(partNumber, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If
        Dim t As String = GetPropertyString(CType(drawDoc, Document), kDrawAssocType, kUserProps)
        If String.IsNullOrEmpty(t) Then Return False
        Return t.Equals(kDrawAssocType1to1, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Sub FillDrawingInfo(drawDoc As DrawingDocument, target As BOMDrawingInfo)
        If drawDoc Is Nothing OrElse target Is Nothing Then Return
        target.FullPath = drawDoc.FullDocumentName
        target.FileName = System.IO.Path.GetFileName(drawDoc.FullDocumentName)
        target.PartNumber = GetPropertyString(CType(drawDoc, Document), "Part Number", kDesignProps)
        target.Description = GetPropertyString(CType(drawDoc, Document), "Description", kDesignProps)
    End Sub

    Private Function GetVaultLikeValue(propName As String, oDoc As Document, logLines As List(Of String), logger As Object) As Object
        If oDoc Is Nothing OrElse String.IsNullOrEmpty(propName) Then Return ""

        ' Common Vault-related aliases mapped to existing properties/date providers.
        Select Case propName.ToLowerInvariant()
            Case "modified", "modifieddate", "lastwrite"
                Return GetFileValue("File.ModifiedDate", oDoc, logLines, logger)
            Case "created", "createddate"
                Return GetFileValue("File.CreatedDate", oDoc, logLines, logger)
            Case "revision", "revisionnumber"
                Return GetPropertyString(oDoc, "Revision Number", kDesignProps)
            Case "checkedby", "checked by"
                Return GetPropertyString(oDoc, "Checked By", kDesignProps)
            Case "designer"
                Return GetPropertyString(oDoc, "Designer", kDesignProps)
        End Select

        ' Generic fallback lookup for Vault-populated metadata copied into iProperties.
        Dim v As String = GetPropertyString(oDoc, propName, kDesignProps)
        If Not String.IsNullOrEmpty(v) Then Return v
        v = GetPropertyString(oDoc, propName, kSummaryProps)
        If Not String.IsNullOrEmpty(v) Then Return v
        v = GetPropertyString(oDoc, propName, kDocSummaryProps)
        If Not String.IsNullOrEmpty(v) Then Return v
        v = GetPropertyString(oDoc, propName, kUserProps)
        Return v
    End Function

    Private Function GetBomValue(inner As String, bom As BOMRow, logLines As List(Of String), logger As Object) As Object
        If bom Is Nothing Then Return ""
        Dim key As String = inner
        If key.StartsWith("BOM.", StringComparison.OrdinalIgnoreCase) AndAlso key.Length > 4 Then
            key = key.Substring(4)
        End If
        key = key.Trim()
        Select Case key.ToLowerInvariant()
            Case "item", "itemnumber"
                Try
                    Dim v As String = CStr(bom.ItemNumber)
                    VLog("BOMExport: BOM lookup '" & inner & "' -> '" & LogValue(v) & "'", logLines, logger)
                    Return v
                Catch : Return "" : End Try
            Case "qty", "quantity", "itemquantity", "itemqty", "item_qty"
                Try
                    Dim v As Double = bom.ItemQuantity
                    VLog("BOMExport: BOM lookup '" & inner & "' -> " & v.ToString(), logLines, logger)
                    Return v
                Catch : Return 0.0 : End Try
            Case "total", "totalqty", "totqty"
                Try
                    Dim v As Double = bom.TotalQuantity
                    VLog("BOMExport: BOM lookup '" & inner & "' -> " & v.ToString(), logLines, logger)
                    Return v
                Catch : Return 0.0 : End Try
            Case "unitqty", "baseqty", "unit_qty"
                Try
                    Dim v As Double = bom.ItemQuantity
                    VLog("BOMExport: BOM lookup '" & inner & "' -> " & v.ToString(), logLines, logger)
                    Return v
                Catch
                    Return 0.0
                End Try
            Case Else
                VLog("BOMExport: BOM lookup '" & inner & "' -> <unsupported>", logLines, logger)
                Return ""
        End Select
    End Function

    Private Function GetPropertyString(oDoc As Document, propName As String, propSetName As String) As String
        If oDoc Is Nothing OrElse String.IsNullOrEmpty(propName) Then Return ""
        Try
            Dim ps As PropertySet = oDoc.PropertySets.Item(propSetName)
            If ps Is Nothing OrElse String.IsNullOrEmpty(propName) Then Return ""
            Dim v As String = CStr(ps.Item(propName).Value)
            VLog("BOMExport: Property lookup [" & propSetName & "].[" & propName & "] from '" & oDoc.DisplayName & "' -> '" & LogValue(v) & "'", Nothing, Nothing)
            Return v
        Catch
        End Try
        VLog("BOMExport: Property lookup [" & propSetName & "].[" & propName & "] -> <missing>", Nothing, Nothing)
        Return ""
    End Function

    Private Function GetPhysical(pDef As PartComponentDefinition, what As String, logLines As List(Of String), logger As Object) As Object
        If pDef Is Nothing OrElse what Is Nothing Then Return ""
        Dim w As String = what.Trim().ToLowerInvariant()
        Try
            Dim m As MassProperties = pDef.MassProperties
            If w = "mass" OrElse w = "weight" OrElse w = "kg" Then
                VLog("BOMExport: Physical lookup '" & what & "' -> " & m.Mass.ToString(), logLines, logger)
                Return m.Mass
            End If
            If w = "area" OrElse w = "surface" Then
                VLog("BOMExport: Physical lookup '" & what & "' -> " & m.Area.ToString(), logLines, logger)
                Return m.Area
            End If
            If w = "volume" OrElse w = "vol" Then
                VLog("BOMExport: Physical lookup '" & what & "' -> " & m.Volume.ToString(), logLines, logger)
                Return m.Volume
            End If
        Catch
        End Try
        VLog("BOMExport: Physical lookup '" & what & "' -> <unsupported/missing>", logLines, logger)
        Return ""
    End Function

    Private Function GetPhysicalFromAsm(asmAD As AssemblyDocument, what As String, logLines As List(Of String), logger As Object) As Object
        If asmAD Is Nothing OrElse what Is Nothing Then Return ""
        Try
            Return GetPhysical(asmAD.ComponentDefinition, what, logLines, logger)
        Catch
        End Try
        Return ""
    End Function

    ' Overload for assembly comp def (treats same as part for MassProperties on assembly)
    Private Function GetPhysical(aDef As AssemblyComponentDefinition, what As String, logLines As List(Of String), logger As Object) As Object
        If aDef Is Nothing OrElse what Is Nothing Then Return ""
        Dim w As String = what.Trim().ToLowerInvariant()
        Try
            Dim m As MassProperties = aDef.MassProperties
            If w = "mass" OrElse w = "weight" OrElse w = "kg" Then Return m.Mass
            If w = "area" OrElse w = "surface" Then Return m.Area
            If w = "volume" OrElse w = "vol" Then Return m.Volume
        Catch
        End Try
        Return ""
    End Function

#End Region

#Region "Expression"

    Private Function EvaluateExpressionString(
        expr As String, asmDoc As AssemblyDocument, bom As BOMRow, oDoc As Document, isAssembly As Boolean, logLines As List(Of String), logger As Object) As Object
        If String.IsNullOrEmpty(expr) Then Return ""
        Dim w0 As String = expr.Trim()
        If Regex.IsMatch(w0, "^\s*now\s*(\s*\(\s*)?\s*$", RegexOptions.IgnoreCase) Then
            VLog("BOMExport: Expr '" & expr & "' -> DateTime.Now", logLines, logger)
            Return System.DateTime.Now
        End If

        Dim s As String = PreprocessExpressionForCompute(w0, asmDoc, bom, oDoc, isAssembly, logLines, logger)
        s = s.Trim()
        If s = "" Then Return 0.0
        Try
            Dim dt As New DataTable()
            Dim v As Object = dt.Compute(s, Nothing)
            VLog("BOMExport: Expr '" & expr & "' preprocessed='" & s & "' -> '" & LogValue(v) & "'", logLines, logger)
            Return v
        Catch
            VLog("BOMExport: Expr '" & expr & "' preprocessed='" & s & "' failed, returning raw", logLines, logger)
            Return s
        End Try
    End Function

    ' Longest BOM tokens first (BOM.Item must come after BOM.ItemNumber).
    Private Function PreprocessExpressionForCompute(
        w As String, asmDoc As AssemblyDocument, bom As BOMRow, oDoc As Document, isAssembly As Boolean, logLines As List(Of String), logger As Object) As String
        Dim t As String = w
        t = t.Replace("BOM.ItemNumber", DoubleToString(GetBomItemNumeric(bom, isAssembly)))
        t = t.Replace("BOM.UnitQty", DoubleToString(If(Not (bom Is Nothing) AndAlso Not isAssembly, SafeBomItemQty(bom), 0.0)))
        t = t.Replace("BOM.TotalQty", DoubleToString(If(Not (bom Is Nothing) AndAlso Not isAssembly, SafeBomTotalQty(bom), 0.0)))
        t = t.Replace("BOM.Qty", DoubleToString(If(Not (bom Is Nothing) AndAlso Not isAssembly, SafeBomQty(bom), 0.0)))
        t = t.Replace("BOM.Item", DoubleToString(GetBomItemNumeric(bom, isAssembly)))
        t = CustomExpressionReplace(t, CType(asmDoc, Document), oDoc, isAssembly)
        t = PhysExpressionReplace(t, "Physical\.", asmDoc, bom, oDoc, isAssembly, logLines, logger)
        t = PhysExpressionReplace(t, "Phys\.", asmDoc, bom, oDoc, isAssembly, logLines, logger)
        t = PhysExpressionReplace(t, "Phy\.", asmDoc, bom, oDoc, isAssembly, logLines, logger)
        Return t
    End Function

    Private Function CustomExpressionReplace(t As String, asmDocD As Document, oDoc As Document, isAssembly As Boolean) As String
        Dim pat As String = "Custom\.([A-Za-z0-9_]+)"
        Do
            Dim m0 As Match = Regex.Match(t, pat, RegexOptions.IgnoreCase)
            If Not m0.Success Then Exit Do
            Dim g As String = m0.Groups(1).Value
            Dim repD As String
            If isAssembly OrElse oDoc Is Nothing Then
                repD = DoubleToString(TryParseDouble0(GetPropertyString(asmDocD, g, kUserProps)))
            Else
                repD = DoubleToString(TryParseDouble0(GetPropertyString(oDoc, g, kUserProps)))
            End If
            t = t.Substring(0, m0.Index) & repD & t.Substring(m0.Index + m0.Length)
        Loop
        Return t
    End Function

    Private Function PhysExpressionReplace(
        t As String, namePrefix As String, asmDoc As AssemblyDocument, bom As BOMRow, oDoc As Document, isAssembly As Boolean,
        logLines As List(Of String), logger As Object) As String
        Dim pat As String = namePrefix & "([A-Za-z0-9_]+)"
        Do
            Dim m0 As Match = Regex.Match(t, pat, RegexOptions.IgnoreCase)
            If Not m0.Success Then Exit Do
            Dim token As String
            If namePrefix = "Physical\." Then
                token = "Physical." & m0.Groups(1).Value
            ElseIf namePrefix = "Phys\." Then
                token = "Phys." & m0.Groups(1).Value
            Else
                token = "Phy." & m0.Groups(1).Value
            End If
            Dim o As Object = Nothing
            If oDoc IsNot Nothing Then
                o = ResolveToken(token, asmDoc, bom, oDoc, isAssembly, logLines, logger)
            End If
            Dim repD As String = DoubleToString(TryToDoubleForExpr(o, 0.0))
            t = t.Substring(0, m0.Index) & repD & t.Substring(m0.Index + m0.Length)
        Loop
        Return t
    End Function

#End Region

#Region "ExpressionHelpers"

    Private Function DoubleToString(d As Double) As String
        If Double.IsNaN(d) OrElse Double.IsInfinity(d) Then Return "0"
        Return d.ToString(System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Function TryParseDouble0(s As String) As Double
        Dim d As Double = 0.0
        If String.IsNullOrEmpty(s) Then Return 0.0
        If Double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, d) Then
            Return d
        End If
        If Double.TryParse(Replace(Replace(s, " ", ""), ",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, d) Then
            Return d
        End If
        Return 0.0
    End Function

    Private Function TryToDoubleForExpr(o As Object, def As Double) As Double
        If o Is Nothing OrElse Convert.IsDBNull(o) Then Return def
        If TypeOf o Is Double Then Return CDbl(o)
        If TypeOf o Is Single Then Return CDbl(CSng(o))
        If TypeOf o Is Integer OrElse TypeOf o Is Long OrElse TypeOf o Is Short Then
            Return CDbl(Convert.ToDouble(o))
        End If
        Return TryParseDouble0(Convert.ToString(o))
    End Function

    Private Function SafeBomQty(bom As BOMRow) As Double
        Try
            Return bom.ItemQuantity
        Catch
        End Try
        Return 0.0
    End Function

    Private Function SafeBomTotalQty(bom As BOMRow) As Double
        Try
            Return bom.TotalQuantity
        Catch
        End Try
        Return 0.0
    End Function

    Private Function SafeBomItemQty(bom As BOMRow) As Double
        Try
            Return bom.ItemQuantity
        Catch
        End Try
        Return 0.0
    End Function

    ' Numeric Item column for DataTable; ItemNumber is often "1" or "1.1" — best-effort parse
    Private Function GetBomItemNumeric(bom As BOMRow, isAs As Boolean) As Double
        If isAs OrElse bom Is Nothing Then Return 0.0
        Return TryParseDouble0(CStr(bom.ItemNumber))
    End Function

#End Region

End Module
