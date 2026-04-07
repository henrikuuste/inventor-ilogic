' ============================================================================
' BoundingBoxStock - Material Stock Size Calculator (Standalone)
' 
' Run this rule on a part document to configure and create bounding box
' iProperties (Width, Height, Thickness) for BOM use.
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/BoundingBoxStockLib.vb"

Sub Main()
    UtilsLib.SetLogger(Logger)
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        MessageBox.Show("No active document.", "Bounding Box Stock")
        Exit Sub
    End If

    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("This rule only works in part documents (.ipt).", "Bounding Box Stock")
        Exit Sub
    End If

    Dim partDoc As PartDocument = CType(doc, PartDocument)

    ' Process the part (standalone mode - no batch info, no skip button)
    Dim result As String = BoundingBoxStockLib.ProcessPartDocument(app, partDoc, "", False, iLogicVb.Automation)

    If result = "OK" Then
        MessageBox.Show( _
            "Created iProperties and auto-update rule." & vbCrLf & vbCrLf & _
            "You can override calculated values by creating parameters:" & vbCrLf & _
            "  WidthOverride, LengthOverride, ThicknessOverride", _
            "Bounding Box Stock")
    End If
End Sub
