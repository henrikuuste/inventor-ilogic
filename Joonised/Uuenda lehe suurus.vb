' ============================================================================
' Uuenda lehe suurus - Fit sheet to content with padding
' 
' Features:
' - Fits the sheet to all views, dimensions, and annotations
' - User can specify padding percentage (default 30%)
' - Option to exclude title block from bounds calculation
'
' Usage: Run on an open drawing to resize the sheet to fit content.
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/CAMDrawingLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    UtilsLib.LogInfo("Uuenda lehe suurus: Starting...")
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Uuenda lehe suurus: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Uuenda lehe suurus")
        Exit Sub
    End If
    
    Dim doc As Document = app.ActiveDocument
    
    If doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        UtilsLib.LogError("Uuenda lehe suurus: Not a drawing document")
        MessageBox.Show("See reegel töötab ainult joonisega.", "Uuenda lehe suurus")
        Exit Sub
    End If
    
    Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
    Dim sheet As Sheet = drawDoc.ActiveSheet
    
    ' Get current sheet info
    Dim viewCount As Integer = sheet.DrawingViews.Count
    Dim currentWidth As Double = sheet.Width * 10
    Dim currentHeight As Double = sheet.Height * 10
    
    If viewCount = 0 Then
        UtilsLib.LogWarn("Uuenda lehe suurus: No views on sheet")
        MessageBox.Show("Lehel puuduvad vaated.", "Uuenda lehe suurus")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Uuenda lehe suurus: Current sheet: " & FormatNumber(currentWidth, 1) & " x " & _
                FormatNumber(currentHeight, 1) & " mm, views: " & viewCount)
    
    ' Show dialog for padding
    Dim paddingPercent As Double = CAMDrawingLib.DEFAULT_SHEET_PADDING
    Dim excludeTitleBlock As Boolean = True
    
    Dim dialogResult As DialogResult = ShowPaddingDialog(sheet, paddingPercent, excludeTitleBlock)
    
    If dialogResult <> DialogResult.OK Then
        UtilsLib.LogInfo("Uuenda lehe suurus: Cancelled by user")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Uuenda lehe suurus: Padding: " & FormatNumber(paddingPercent * 100, 0) & "%, Exclude title block: " & excludeTitleBlock)
    
    ' Fit sheet to content
    CAMDrawingLib.FitSheetToContent(sheet, app, paddingPercent, CAMDrawingLib.DEFAULT_BORDER_PADDING, excludeTitleBlock)
    ' Get new sheet size
    Dim newWidth As Double = sheet.Width * 10
    Dim newHeight As Double = sheet.Height * 10
    
    ' Summary
    UtilsLib.LogInfo("Uuenda lehe suurus: ========================================")
    UtilsLib.LogInfo("Uuenda lehe suurus: COMPLETE")
    UtilsLib.LogInfo("Uuenda lehe suurus: " & FormatNumber(currentWidth, 0) & "x" & FormatNumber(currentHeight, 0) & _
                " -> " & FormatNumber(newWidth, 0) & "x" & FormatNumber(newHeight, 0) & " mm")
    UtilsLib.LogInfo("Uuenda lehe suurus: ========================================")
End Sub

' ============================================================================
' Padding Dialog
' ============================================================================

Function ShowPaddingDialog(sheet As Sheet, _
                            ByRef paddingPercent As Double, _
                            ByRef excludeTitleBlock As Boolean) As DialogResult
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Uuenda lehe suurus"
    frm.Width = 400
    frm.Height = 250
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MinimizeBox = False
    frm.MaximizeBox = False
    
    Dim currentY As Integer = 20
    
    ' Current sheet info
    Dim lblInfo As New System.Windows.Forms.Label()
    lblInfo.Text = "Praegune lehe suurus: " & FormatNumber(sheet.Width * 10, 1) & " x " & _
                   FormatNumber(sheet.Height * 10, 1) & " mm" & vbCrLf & _
                   "Vaateid: " & sheet.DrawingViews.Count
    lblInfo.Left = 20
    lblInfo.Top = currentY
    lblInfo.Width = 350
    lblInfo.Height = 40
    frm.Controls.Add(lblInfo)
    
    currentY += 55
    
    ' Padding input
    Dim lblPadding As New System.Windows.Forms.Label()
    lblPadding.Text = "Polsterdus (%):"
    lblPadding.Left = 20
    lblPadding.Top = currentY + 3
    lblPadding.Width = 100
    frm.Controls.Add(lblPadding)
    
    Dim nudPadding As New System.Windows.Forms.NumericUpDown()
    nudPadding.Name = "nudPadding"
    nudPadding.Left = 130
    nudPadding.Top = currentY
    nudPadding.Width = 80
    nudPadding.Minimum = 0
    nudPadding.Maximum = 200
    nudPadding.DecimalPlaces = 0
    nudPadding.Value = CDec(paddingPercent * 100)
    nudPadding.Increment = 5
    frm.Controls.Add(nudPadding)
    
    Dim lblPaddingHint As New System.Windows.Forms.Label()
    lblPaddingHint.Text = "(tüüpiline: 20-50%)"
    lblPaddingHint.Left = 220
    lblPaddingHint.Top = currentY + 3
    lblPaddingHint.Width = 120
    frm.Controls.Add(lblPaddingHint)
    
    currentY += 40
    
    ' Exclude title block checkbox
    Dim chkExcludeTitleBlock As New System.Windows.Forms.CheckBox()
    chkExcludeTitleBlock.Name = "chkExcludeTitleBlock"
    chkExcludeTitleBlock.Text = "Jäta nurgalehe raam arvestusest välja"
    chkExcludeTitleBlock.Left = 20
    chkExcludeTitleBlock.Top = currentY
    chkExcludeTitleBlock.Width = 300
    chkExcludeTitleBlock.Checked = excludeTitleBlock
    frm.Controls.Add(chkExcludeTitleBlock)
    
    currentY += 50
    
    ' Buttons
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Uuenda"
    btnOK.Left = 195
    btnOK.Top = currentY
    btnOK.Width = 85
    btnOK.Height = 28
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 285
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
        paddingPercent = CDbl(nudPadding.Value) / 100.0
        excludeTitleBlock = chkExcludeTitleBlock.Checked
    End If
    
    frm.Dispose()
    Return result
End Function
