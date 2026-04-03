' Lehtmetall - Teisenda tahke detail lehtmetalliks
' Teisendab aktiivse detaili lehtmetalliks, ekspordib Thickness iProperty'na,
' seab Width/Length kohandatud omadused lehtmetalli avaldistega
' ja loob sirge mustri peale A-külje pinna valimist.

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = ThisDoc.Document
    
    ' Kontrolli dokumendi tüüpi
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("Seda reeglit saab käivitada ainult detaili dokumendil.", "Lehtmetall")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    ' Kontrolli, kas on olemas tahke keha
    If compDef.SurfaceBodies.Count = 0 Then
        MessageBox.Show("Detailil puudub tahke keha. Veendu, et detail sisaldab geomeetriat enne lehtmetalliks teisendamist.", "Lehtmetall")
        Exit Sub
    End If
    
    ' Kontrolli, kas tahke keha on olemas (mitte ainult pinnad)
    Dim hasSolidBody As Boolean = False
    For Each body As SurfaceBody In compDef.SurfaceBodies
        If body.IsSolid Then
            hasSolidBody = True
            Exit For
        End If
    Next
    
    If Not hasSolidBody Then
        MessageBox.Show("Detailil puudub tahke keha. Leitud on ainult pinnad, mis ei sobi lehtmetalliks teisendamiseks.", "Lehtmetall")
        Exit Sub
    End If
    
    ' Kontrolli, kas on juba lehtmetall
    Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    If partDoc.SubType = SHEET_METAL_GUID Then
        MessageBox.Show("See detail on juba lehtmetall.", "Lehtmetall")
        Exit Sub
    End If
    
    ' Teisenda lehtmetalliks
    partDoc.SubType = SHEET_METAL_GUID
    partDoc.Update()
    
    ' Hangi lehtmetalli komponendi definitsioon
    Dim smCompDef As SheetMetalComponentDefinition = partDoc.ComponentDefinition
    
    ' Ekspordi Thickness parameeter iProperty'na
    ExportThicknessAsProperty(smCompDef)
    
    ' Sea Width ja Length kohandatud omadused avaldistega
    SetSheetMetalProperties(partDoc)
    
    ' Küsi kasutajalt A-külje pinna valik ja loo sirge muster
    CreateFlatPattern(app, smCompDef)
    
    partDoc.Update()
End Sub

Sub ExportThicknessAsProperty(smCompDef As SheetMetalComponentDefinition)
    Try
        Dim thicknessParam As Parameter = smCompDef.Thickness
        thicknessParam.ExposedAsProperty = True
        thicknessParam.CustomPropertyFormat.PropertyType = CustomPropertyTypeEnum.kTextPropertyType
        thicknessParam.CustomPropertyFormat.ShowUnitsString = True
    Catch ex As Exception
        MessageBox.Show("Hoiatus: Thickness parameetrit ei õnnestunud iProperty'na eksportida. " & ex.Message, "Lehtmetall")
    End Try
End Sub

Sub SetSheetMetalProperties(partDoc As PartDocument)
    Try
        Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
        
        ' Sea Width omadus avaldisega, mis viitab Sheet Metal Width parameetrile
        SetOrAddProperty(propSet, "Width", "=<Sheet Metal Width>")
        
        ' Sea Length omadus avaldisega, mis viitab Sheet Metal Length parameetrile
        SetOrAddProperty(propSet, "Length", "=<Sheet Metal Length>")
    Catch ex As Exception
        MessageBox.Show("Hoiatus: Width/Length omadusi ei õnnestunud seada. " & ex.Message, "Lehtmetall")
    End Try
End Sub

Sub SetOrAddProperty(propSet As PropertySet, propName As String, propValue As String)
    Try
        propSet.Item(propName).Value = propValue
    Catch
        Try
            propSet.Add(propValue, propName)
        Catch
        End Try
    End Try
End Sub

Sub CreateFlatPattern(app As Inventor.Application, smCompDef As SheetMetalComponentDefinition)
    Dim aSideFace As Face = Nothing
    
    Try
        aSideFace = app.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, _
            "Vali A-külje pind sirgeks mustriks (ESC tühistamiseks)")
    Catch
        ' Kasutaja tühistas
        MessageBox.Show("Pinna valik tühistatud. Sirget mustrit ei loodud.", "Lehtmetall")
        Exit Sub
    End Try
    
    If aSideFace Is Nothing Then
        MessageBox.Show("Pinda ei valitud. Sirget mustrit ei loodud.", "Lehtmetall")
        Exit Sub
    End If
    
    Try
        smCompDef.ASideFace = aSideFace
        smCompDef.Unfold()
    Catch ex As Exception
        MessageBox.Show("Sirget mustrit ei õnnestunud luua: " & ex.Message, "Lehtmetall")
    End Try
End Sub
