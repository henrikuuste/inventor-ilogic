' ============================================================================
' Lisa vaated - Add projected views to existing base view
' 
' Features:
' - Adds all orthographic projected views around an existing base view
' - Optional spacing parameter for view gaps
' - Optionally fits sheet to content after adding views
'
' Usage: Run on a drawing with at least one base view.
'        Select the base view or the first view will be used.
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/CAMDrawingLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    UtilsLib.LogInfo("Lisa vaated: Starting...")
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Lisa vaated: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Lisa vaated")
        Exit Sub
    End If
    
    Dim doc As Document = app.ActiveDocument
    
    If doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        UtilsLib.LogError("Lisa vaated: Not a drawing document")
        MessageBox.Show("See reegel töötab ainult joonisega.", "Lisa vaated")
        Exit Sub
    End If
    
    Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
    Dim sheet As Sheet = drawDoc.ActiveSheet
    
    If sheet.DrawingViews.Count = 0 Then
        UtilsLib.LogError("Lisa vaated: No views on sheet")
        MessageBox.Show("Lehel puuduvad vaated. Kasuta 'Loo 1:1 joonised' uue joonise loomiseks.", "Lisa vaated")
        Exit Sub
    End If
    
    ' Find the base view (from selection or first view)
    Dim baseView As DrawingView = Nothing
    
    ' Check selection
    Dim selectedViews As List(Of DrawingView) = CAMDrawingLib.GetSelectedViews(drawDoc)
    If selectedViews.Count > 0 Then
        baseView = selectedViews(0)
        UtilsLib.LogInfo("Lisa vaated: Using selected view: " & baseView.Name)
    Else
        ' Use first view on sheet
        baseView = sheet.DrawingViews.Item(1)
        UtilsLib.LogInfo("Lisa vaated: Using first view: " & baseView.Name)
    End If
    
    ' Get the referenced part document
    Dim partDoc As PartDocument = CAMDrawingLib.GetReferencedPartDocument(drawDoc)
    If partDoc Is Nothing Then
        UtilsLib.LogError("Lisa vaated: No referenced part document found")
        MessageBox.Show("Jooniselt ei leitud viidet detailile.", "Lisa vaated")
        Exit Sub
    End If
    
    ' Current view count
    Dim viewCountBefore As Integer = sheet.DrawingViews.Count
    
    ' Show options dialog
    Dim dimOffsetCm As Double = CAMDrawingLib.DEFAULT_DIMENSION_OFFSET
    Dim viewGapCm As Double = CAMDrawingLib.DEFAULT_VIEW_GAP
    Dim fitSheet As Boolean = True
    
    Dim dialogResult As DialogResult = ShowOptionsDialog(baseView, partDoc, dimOffsetCm, viewGapCm, fitSheet)
    
    If dialogResult <> DialogResult.OK Then
        UtilsLib.LogInfo("Lisa vaated: Cancelled by user")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Lisa vaated: Base view: " & baseView.Name)
    UtilsLib.LogInfo("Lisa vaated: Dimension offset: " & FormatNumber(dimOffsetCm * 10, 1) & " mm")
    UtilsLib.LogInfo("Lisa vaated: View gap: " & FormatNumber(viewGapCm * 10, 1) & " mm")
    
    ' Add projected views
    Dim allViews As List(Of DrawingView) = CAMDrawingLib.AddViewsToExistingBase(sheet, baseView, partDoc, app, dimOffsetCm, viewGapCm)
    Dim viewsAdded As Integer = allViews.Count - 1  ' Subtract base view
    
    ' Optionally fit sheet to content
    If fitSheet Then
        UtilsLib.LogInfo("Lisa vaated: Fitting sheet to content...")
        CAMDrawingLib.FitSheetToContent(sheet, app)
    End If
    
    ' Summary
    UtilsLib.LogInfo("Lisa vaated: ========================================")
    UtilsLib.LogInfo("Lisa vaated: COMPLETE")
    UtilsLib.LogInfo("Lisa vaated: Views added: " & viewsAdded)
    UtilsLib.LogInfo("Lisa vaated: Total views: " & sheet.DrawingViews.Count)
    UtilsLib.LogInfo("Lisa vaated: ========================================")
End Sub

' ============================================================================
' Options Dialog
' ============================================================================

Function ShowOptionsDialog(baseView As DrawingView, _
                            partDoc As PartDocument, _
                            ByRef dimOffsetCm As Double, _
                            ByRef viewGapCm As Double, _
                            ByRef fitSheet As Boolean) As DialogResult
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Lisa vaated"
    frm.Width = 450
    frm.Height = 320
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MinimizeBox = False
    frm.MaximizeBox = False
    
    Dim currentY As Integer = 20
    
    ' Base view info
    Dim lblBaseView As New System.Windows.Forms.Label()
    lblBaseView.Text = "Baasvaat: " & baseView.Name
    lblBaseView.Left = 20
    lblBaseView.Top = currentY
    lblBaseView.Width = 400
    frm.Controls.Add(lblBaseView)
    
    currentY += 25
    
    ' Part info
    Dim lblPart As New System.Windows.Forms.Label()
    lblPart.Text = "Detail: " & partDoc.DisplayName
    lblPart.Left = 20
    lblPart.Top = currentY
    lblPart.Width = 400
    frm.Controls.Add(lblPart)
    
    currentY += 35
    
    ' Separator
    Dim separator As New System.Windows.Forms.Label()
    separator.Text = "─────────────────────────────────────────────"
    separator.Left = 20
    separator.Top = currentY
    separator.Width = 400
    frm.Controls.Add(separator)
    
    currentY += 25
    
    ' Dimension offset input
    Dim lblDimOffset As New System.Windows.Forms.Label()
    lblDimOffset.Text = "Mõõdu kaugus (mm):"
    lblDimOffset.Left = 20
    lblDimOffset.Top = currentY + 3
    lblDimOffset.Width = 130
    frm.Controls.Add(lblDimOffset)
    
    Dim nudDimOffset As New System.Windows.Forms.NumericUpDown()
    nudDimOffset.Name = "nudDimOffset"
    nudDimOffset.Left = 160
    nudDimOffset.Top = currentY
    nudDimOffset.Width = 80
    nudDimOffset.Minimum = 5
    nudDimOffset.Maximum = 100
    nudDimOffset.DecimalPlaces = 0
    nudDimOffset.Value = CDec(dimOffsetCm * 10)
    nudDimOffset.Increment = 5
    frm.Controls.Add(nudDimOffset)
    
    Dim lblDimOffsetHint As New System.Windows.Forms.Label()
    lblDimOffsetHint.Text = "(ruum mõõdule)"
    lblDimOffsetHint.Left = 250
    lblDimOffsetHint.Top = currentY + 3
    lblDimOffsetHint.Width = 150
    frm.Controls.Add(lblDimOffsetHint)
    
    currentY += 35
    
    ' View gap input
    Dim lblViewGap As New System.Windows.Forms.Label()
    lblViewGap.Text = "Vaatevahik (mm):"
    lblViewGap.Left = 20
    lblViewGap.Top = currentY + 3
    lblViewGap.Width = 130
    frm.Controls.Add(lblViewGap)
    
    Dim nudViewGap As New System.Windows.Forms.NumericUpDown()
    nudViewGap.Name = "nudViewGap"
    nudViewGap.Left = 160
    nudViewGap.Top = currentY
    nudViewGap.Width = 80
    nudViewGap.Minimum = 0
    nudViewGap.Maximum = 100
    nudViewGap.DecimalPlaces = 0
    nudViewGap.Value = CDec(viewGapCm * 10)
    nudViewGap.Increment = 5
    frm.Controls.Add(nudViewGap)
    
    Dim lblViewGapHint As New System.Windows.Forms.Label()
    lblViewGapHint.Text = "(vaatevaheline ruum)"
    lblViewGapHint.Left = 250
    lblViewGapHint.Top = currentY + 3
    lblViewGapHint.Width = 150
    frm.Controls.Add(lblViewGapHint)
    
    currentY += 35
    
    ' Total spacing info
    Dim lblTotal As New System.Windows.Forms.Label()
    lblTotal.Name = "lblTotal"
    lblTotal.Text = "Kokku vaatevaheline kaugus: " & FormatNumber((dimOffsetCm + viewGapCm) * 10, 0) & " mm"
    lblTotal.Left = 20
    lblTotal.Top = currentY
    lblTotal.Width = 400
    frm.Controls.Add(lblTotal)
    
    ' Update total when values change
    AddHandler nudDimOffset.ValueChanged, Sub(s, e)
        lblTotal.Text = "Kokku vaatevaheline kaugus: " & FormatNumber(CDbl(nudDimOffset.Value) + CDbl(nudViewGap.Value), 0) & " mm"
    End Sub
    
    AddHandler nudViewGap.ValueChanged, Sub(s, e)
        lblTotal.Text = "Kokku vaatevaheline kaugus: " & FormatNumber(CDbl(nudDimOffset.Value) + CDbl(nudViewGap.Value), 0) & " mm"
    End Sub
    
    currentY += 35
    
    ' Fit sheet checkbox
    Dim chkFitSheet As New System.Windows.Forms.CheckBox()
    chkFitSheet.Name = "chkFitSheet"
    chkFitSheet.Text = "Kohanda lehe suurus sisule"
    chkFitSheet.Left = 20
    chkFitSheet.Top = currentY
    chkFitSheet.Width = 300
    chkFitSheet.Checked = fitSheet
    frm.Controls.Add(chkFitSheet)
    
    currentY += 45
    
    ' Buttons
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Lisa vaated"
    btnOK.Left = 240
    btnOK.Top = currentY
    btnOK.Width = 95
    btnOK.Height = 28
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 340
    btnCancel.Top = currentY
    btnCancel.Width = 85
    btnCancel.Height = 28
    btnCancel.DialogResult = DialogResult.Cancel
    frm.Controls.Add(btnCancel)
    frm.CancelButton = btnCancel
    
    ' Show dialog
    Dim result As DialogResult = frm.ShowDialog()
    
    ' Extract values
    If result = DialogResult.OK Then
        dimOffsetCm = CDbl(nudDimOffset.Value) / 10.0
        viewGapCm = CDbl(nudViewGap.Value) / 10.0
        fitSheet = chkFitSheet.Checked
    End If
    
    frm.Dispose()
    Return result
End Function
