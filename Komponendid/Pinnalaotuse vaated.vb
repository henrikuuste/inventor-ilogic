' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Pinnalaotuse vaated — Design View Representations in the part only
'
' Creates/updates two DVRs on the active .ipt (no drawing is created or saved here):
'   • Komponent  — bent/original solid(s); unwrap surface and thickened flat solid hidden
'   • Pinnalaotus — only the thickened unwrap solid (manufactured); Loo 1-1 joonised.vb uses this for 1:1 CAM drawings
'
' Joonised/Loo 1-1 joonised.vb — creates/updates .idw files and sheet views.
'
' Usage: Open the unwrap+thicken part, run this rule, then save the part if you want DVR changes kept on disk.
' ============================================================================

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/UnwrapLib.vb"

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
    
    If Not UnwrapLib.HasCompletePinnalaotus(partDoc) Then
        UtilsLib.LogWarn("Pinnalaotuse vaated: Part needs Unwrap + Thicken")
        MessageBox.Show("Detailil peavad olema Pinnalaotus (Unwrap) ja järgnev Thicken.", "Pinnalaotuse vaated")
        Exit Sub
    End If
    
    Dim kompOk As Boolean = (UnwrapLib.GetOrCreateKomponentDVR(partDoc) IsNot Nothing)
    Dim pinOk As Boolean = (UnwrapLib.GetOrCreatePinnalaotusDVR(partDoc) IsNot Nothing)
    
    If Not kompOk OrElse Not pinOk Then
        UtilsLib.LogError("Pinnalaotuse vaated: DVR creation failed (Komponent=" & kompOk.ToString() & ", Pinnalaotus=" & pinOk.ToString() & ")")
        MessageBox.Show("DVR-de loomine ebaõnnestus. Vaata logi.", "Pinnalaotuse vaated")
        Exit Sub
    End If
    
    UtilsLib.LogInfo("Pinnalaotuse vaated: Komponent ja Pinnalaotus DVR on detailis uuendatud.")
End Sub
