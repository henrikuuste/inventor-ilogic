' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Lisa mõõdud - Add extent dimensions to drawing views
' 
' Features:
' - Adds horizontal and vertical extent dimensions to views
' - Works on selected views or all views if none selected
' - Optional parameter for dimension offset from view edge
'
' Usage: Run on a drawing with views.
'        Select specific views or leave empty to dimension all views.
' ============================================================================

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/FileSearchLib.vb"
AddVbFile "Lib/CAMDrawingLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    UtilsLib.LogInfo("Lisa mõõdud: Starting...")
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Lisa mõõdud: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Lisa mõõdud")
        Exit Sub
    End If
    
    Dim doc As Document = app.ActiveDocument
    
    If doc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
        UtilsLib.LogError("Lisa mõõdud: Not a drawing document")
        MessageBox.Show("See reegel töötab ainult joonisega.", "Lisa mõõdud")
        Exit Sub
    End If
    
    Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
    Dim sheet As Sheet = drawDoc.ActiveSheet
    
    If sheet.DrawingViews.Count = 0 Then
        UtilsLib.LogError("Lisa mõõdud: No views on sheet")
        MessageBox.Show("Lehel puuduvad vaated.", "Lisa mõõdud")
        Exit Sub
    End If
    
    ' Get views to dimension
    Dim views As List(Of DrawingView) = Nothing
    Dim useSelection As Boolean = False
    
    ' Check selection
    Dim selectedViews As List(Of DrawingView) = CAMDrawingLib.GetSelectedViews(drawDoc)
    If selectedViews.Count > 0 Then
        views = selectedViews
        useSelection = True
        UtilsLib.LogInfo("Lisa mõõdud: Using " & views.Count & " selected view(s)")
    Else
        ' Use all views on sheet
        views = CAMDrawingLib.GetAllViewsFromSheet(sheet)
        UtilsLib.LogInfo("Lisa mõõdud: Using all " & views.Count & " views on sheet")
    End If
    
    ' Count existing dimensions
    Dim dimCountBefore As Integer = sheet.DrawingDimensions.Count
    
    ' Show options dialog
    Dim offsetCm As Double = CAMDrawingLib.DEFAULT_DIMENSION_OFFSET
    
    Dim dialogResult As DialogResult = ShowOptionsDialog(views, useSelection, offsetCm)
    
    If dialogResult <> DialogResult.OK Then
        UtilsLib.LogInfo("Lisa mõõdud: Cancelled by user")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Lisa mõõdud: Dimension offset: " & FormatNumber(offsetCm * 10, 1) & " mm")
    UtilsLib.LogInfo("Lisa mõõdud: Processing " & views.Count & " view(s)")
    
    ' Add extent dimensions to views
    CAMDrawingLib.AddExtentDimensionsToViews(sheet, views, app, offsetCm)
    ' Count dimensions added
    Dim dimCountAfter As Integer = sheet.DrawingDimensions.Count
    Dim dimsAdded As Integer = dimCountAfter - dimCountBefore
    
    ' Summary
    UtilsLib.LogInfo("Lisa mõõdud: ========================================")
    UtilsLib.LogInfo("Lisa mõõdud: COMPLETE")
    UtilsLib.LogInfo("Lisa mõõdud: Views processed: " & views.Count)
    UtilsLib.LogInfo("Lisa mõõdud: Dimensions added: " & dimsAdded)
    UtilsLib.LogInfo("Lisa mõõdud: Total dimensions: " & dimCountAfter)
    UtilsLib.LogInfo("Lisa mõõdud: ========================================")
End Sub

' ============================================================================
' Options Dialog
' ============================================================================

Function ShowOptionsDialog(views As List(Of DrawingView), _
                            useSelection As Boolean, _
                            ByRef offsetCm As Double) As DialogResult
    
    Dim frm As New System.Windows.Forms.Form()
    frm.Text = "Lisa mõõdud"
    frm.Width = 430
    frm.Height = 280
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.MinimizeBox = False
    frm.MaximizeBox = False
    
    Dim currentY As Integer = 20
    
    ' Views info
    Dim lblViews As New System.Windows.Forms.Label()
    If useSelection Then
        lblViews.Text = "Valitud vaated (" & views.Count & "):"
    Else
        lblViews.Text = "Kõik vaated lehel (" & views.Count & "):"
    End If
    lblViews.Left = 20
    lblViews.Top = currentY
    lblViews.Width = 380
    frm.Controls.Add(lblViews)
    
    currentY += 25
    
    ' View names
    Dim viewNames As String = ""
    For i As Integer = 0 To Math.Min(views.Count - 1, 5)
        If viewNames.Length > 0 Then viewNames &= ", "
        viewNames &= views(i).Name
    Next
    If views.Count > 6 Then viewNames &= ", ..."
    
    Dim lblViewNames As New System.Windows.Forms.Label()
    lblViewNames.Text = viewNames
    lblViewNames.Left = 30
    lblViewNames.Top = currentY
    lblViewNames.Width = 360
    frm.Controls.Add(lblViewNames)
    
    currentY += 35
    
    ' Separator
    Dim separator As New System.Windows.Forms.Label()
    separator.Text = "─────────────────────────────────────────────"
    separator.Left = 20
    separator.Top = currentY
    separator.Width = 380
    frm.Controls.Add(separator)
    
    currentY += 25
    
    ' Dimension offset input
    Dim lblOffset As New System.Windows.Forms.Label()
    lblOffset.Text = "Mõõdu kaugus vaatest (mm):"
    lblOffset.Left = 20
    lblOffset.Top = currentY + 3
    lblOffset.Width = 170
    frm.Controls.Add(lblOffset)
    
    Dim nudOffset As New System.Windows.Forms.NumericUpDown()
    nudOffset.Name = "nudOffset"
    nudOffset.Left = 200
    nudOffset.Top = currentY
    nudOffset.Width = 80
    nudOffset.Minimum = 5
    nudOffset.Maximum = 100
    nudOffset.DecimalPlaces = 0
    nudOffset.Value = CDec(offsetCm * 10)
    nudOffset.Increment = 5
    frm.Controls.Add(nudOffset)
    
    currentY += 35
    
    ' Hint
    Dim lblHint As New System.Windows.Forms.Label()
    lblHint.Text = "Mõõtmed lisatakse vaate servast määratud " & vbCrLf & _
                   "kaugusele (horisontaal alla, vertikaal paremale)."
    lblHint.Left = 20
    lblHint.Top = currentY
    lblHint.Width = 380
    lblHint.Height = 35
    frm.Controls.Add(lblHint)
    
    currentY += 55
    
    ' Buttons
    Dim btnOK As New System.Windows.Forms.Button()
    btnOK.Text = "Lisa mõõdud"
    btnOK.Left = 220
    btnOK.Top = currentY
    btnOK.Width = 95
    btnOK.Height = 28
    btnOK.DialogResult = DialogResult.OK
    frm.Controls.Add(btnOK)
    frm.AcceptButton = btnOK
    
    Dim btnCancel As New System.Windows.Forms.Button()
    btnCancel.Text = "Tühista"
    btnCancel.Left = 320
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
        offsetCm = CDbl(nudOffset.Value) / 10.0
    End If
    
    frm.Dispose()
    Return result
End Function
