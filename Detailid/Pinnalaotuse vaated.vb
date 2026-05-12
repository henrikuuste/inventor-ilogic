' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Pinnalaotuse vaated — Design View Representations in the part only
'
' Creates/updates two DVRs on the active .ipt:
'   • Komponent  — bent/original solid(s); unwrap surface and manufactured flat body hidden
'   • Pinnalaotus — only the manufactured flat solid (Thicken or Extrude body you pick)
'
' Pick the manufactured body in the dialog. Default selection: saved BB_PinnalaotusSolidBodyName,
' else autodetect Thicken output if present.
'
' Joonised/Loo 1-1 joonised.vb — creates/updates .idw files and sheet views.
' ============================================================================

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/UnwrapLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Pinnalaotuse vaated: No active document")
        MessageBox.Show("Ava esmalt detail (.ipt).", "Pinnalaotuse vaated")
        Exit Sub
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        UtilsLib.LogError("Pinnalaotuse vaated: Not a part document")
        MessageBox.Show("See reegel töötab ainult detailiga.", "Pinnalaotuse vaated")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(app.ActiveDocument, PartDocument)
    
    If Not UnwrapLib.HasUnwrapFeature(partDoc) Then
        UtilsLib.LogWarn("Pinnalaotuse vaated: No Unwrap feature")
        MessageBox.Show("Detailil peab olema Unwrap (Pinnalaotus).", "Pinnalaotuse vaated")
        Exit Sub
    End If
    
    Dim pickedBody As SurfaceBody = PickManufacturedSolidBodyWithUi(partDoc)
    If pickedBody Is Nothing Then
        UtilsLib.LogInfo("Pinnalaotuse vaated: Cancelled or no valid body")
        Exit Sub
    End If
    
    UnwrapLib.SetManufacturedSolidBodyNameProperty(partDoc, pickedBody.Name)
    
    Dim kompOk As Boolean = (UnwrapLib.GetOrCreateKomponentDVR(partDoc, pickedBody) IsNot Nothing)
    Dim pinOk As Boolean = (UnwrapLib.GetOrCreatePinnalaotusDVR(partDoc, pickedBody) IsNot Nothing)
    
    If Not kompOk OrElse Not pinOk Then
        UtilsLib.LogError("Pinnalaotuse vaated: DVR creation failed (Komponent=" & kompOk.ToString() & ", Pinnalaotus=" & pinOk.ToString() & ")")
        MessageBox.Show("DVR-de loomine ebaõnnestus. Vaata logi.", "Pinnalaotuse vaated")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Pinnalaotuse vaated: Komponent ja Pinnalaotus DVR on detailis uuendatud (toodetud keha: " & pickedBody.Name & ").")
    MessageBox.Show("DVR-id uuendatud." & vbCrLf & "Toodetud keha: " & pickedBody.Name & vbCrLf & vbCrLf &
                    "Salvesta detail vajadusel." & vbCrLf & "Omadus " & UnwrapLib.PROP_PINNALAOTUS_BODY_NAME & " on uuendatud.",
                    "Pinnalaotuse vaated")
End Sub

''' <summary>
''' Dialog: choose manufactured solid (Thicken or Extrude). Default = property match, else Thicken body, else first non-unwrap body.
''' </summary>
Function PickManufacturedSolidBodyWithUi(partDoc As PartDocument) As SurfaceBody
    Dim bodies As New List(Of SurfaceBody)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    For Each b As SurfaceBody In compDef.SurfaceBodies
        bodies.Add(b)
    Next
    
    If bodies.Count = 0 Then
        MessageBox.Show("Detailis pole ühtegi keha.", "Pinnalaotuse vaated")
        Return Nothing
    End If
    
    Dim unwrapFeat As UnwrapFeature = UnwrapLib.GetUnwrapFeature(partDoc)
    Dim unwrapSurf As SurfaceBody = Nothing
    If unwrapFeat IsNot Nothing Then unwrapSurf = UnwrapLib.GetUnwrappedSurfaceBody(unwrapFeat)
    
    Dim frm As New Form()
    frm.Text = "Pinnalaotuse vaated — toodetud keha"
    frm.FormBorderStyle = FormBorderStyle.FixedDialog
    frm.StartPosition = FormStartPosition.CenterScreen
    frm.MinimizeBox = False
    frm.MaximizeBox = False
    frm.Width = 440
    frm.Height = 200
    
    Dim lbl As New Label()
    lbl.Left = 12
    lbl.Top = 12
    lbl.Width = 400
    lbl.Height = 48
    lbl.Text = "Vali toodetud lame keha (Thicken või Extrude)." & vbCrLf &
               "Vaikimisi: salvestatud omadus või automaatne Thicken, kui leitud."
    frm.Controls.Add(lbl)
    
    Dim cb As New ComboBox()
    cb.Left = 12
    cb.Top = 64
    cb.Width = 400
    cb.DropDownStyle = ComboBoxStyle.DropDownList
    For Each b As SurfaceBody In bodies
        cb.Items.Add(b.Name)
    Next
    frm.Controls.Add(cb)
    
    Dim defaultIdx As Integer = 0
    Dim propName As String = UnwrapLib.GetManufacturedSolidBodyNameProperty(partDoc)
    If Not String.IsNullOrWhiteSpace(propName) Then
        For i As Integer = 0 To bodies.Count - 1
            If String.Equals(bodies(i).Name, propName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                defaultIdx = i
                Exit For
            End If
        Next
    Else
        Dim thickBody As SurfaceBody = UnwrapLib.TryGetThickenManufacturedSolidBody(partDoc)
        If thickBody IsNot Nothing Then
            For i As Integer = 0 To bodies.Count - 1
                Try
                    If ReferenceEquals(bodies(i), thickBody) OrElse _
                       String.Equals(bodies(i).Name, thickBody.Name, StringComparison.OrdinalIgnoreCase) Then
                        defaultIdx = i
                        Exit For
                    End If
                Catch
                End Try
            Next
        Else
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
        End If
    End If
    
    cb.SelectedIndex = Math.Min(Math.Max(0, defaultIdx), cb.Items.Count - 1)
    
    Dim btnOk As New Button()
    btnOk.Text = "OK"
    btnOk.DialogResult = DialogResult.OK
    btnOk.Left = 240
    btnOk.Top = 110
    btnOk.Width = 80
    frm.Controls.Add(btnOk)
    frm.AcceptButton = btnOk
    
    Dim btnCancel As New Button()
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
    
    If unwrapSurf IsNot Nothing Then
        Try
            If ReferenceEquals(chosen, unwrapSurf) OrElse _
               String.Equals(chosen.Name, unwrapSurf.Name, StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("Unwrap väljundi pinda ei saa toodetud kehana valida. Vali Extrude/Thicken tahvelpaneeli keha.",
                                "Pinnalaotuse vaated")
                Return Nothing
            End If
        Catch
        End Try
    End If
    
    Return chosen
End Function
