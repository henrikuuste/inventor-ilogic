' Copyright (c) 2026 Henri Kuuste
' Test1_Fingerprint.vb
' PURPOSE: Validate that fingerprinting is stable and deterministic
' 
' TESTS:
' 1. Same part, multiple calls = identical fingerprint
' 2. Different parts = different fingerprints
' 3. Bounding box dimensions are orientation-independent (sorted)
' 4. Sheet metal parts work correctly
' 5. Multibody parts - can fingerprint individual bodies
'
' RUN: Open any part file, then run this rule

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("Test1_Fingerprint: Open a part document first")
        MessageBox.Show("Ava esmalt detaili fail (.ipt)", "Test1")
        Return
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    Logger.Info("=== Test1_Fingerprint: Starting ===")
    Logger.Info("Part: " & partDoc.DisplayName)
    Logger.Info("")
    
    ' Count solid bodies
    Dim solidBodies As New List(Of SurfaceBody)
    For Each body As SurfaceBody In compDef.SurfaceBodies
        If body.IsSolid Then solidBodies.Add(body)
    Next
    
    Logger.Info("Solid bodies found: " & solidBodies.Count)
    If solidBodies.Count = 0 Then
        Logger.Warn("No solid bodies found in this part")
        Return
    End If
    
    ' === TEST 1: Determinism - same body, multiple calls ===
    Logger.Info("")
    Logger.Info("--- TEST 1: Determinism (same body, 3 calls) ---")
    
    Dim testBody As SurfaceBody = solidBodies(0)
    Dim fp1 As String = ComputePartFingerprint(testBody)
    Dim fp2 As String = ComputePartFingerprint(testBody)
    Dim fp3 As String = ComputePartFingerprint(testBody)
    
    Logger.Info("Call 1: " & fp1)
    Logger.Info("Call 2: " & fp2)
    Logger.Info("Call 3: " & fp3)
    
    If fp1 = fp2 AndAlso fp2 = fp3 Then
        Logger.Info("PASS: Fingerprints are deterministic")
    Else
        Logger.Error("FAIL: Fingerprints differ between calls!")
    End If
    
    ' === TEST 2: All bodies in part ===
    Logger.Info("")
    Logger.Info("--- TEST 2: All body fingerprints ---")
    
    Dim bodyFingerprints As New Dictionary(Of String, String)
    For Each body As SurfaceBody In solidBodies
        Dim fp As String = ComputePartFingerprint(body)
        bodyFingerprints.Add(body.Name, fp)
        Logger.Info("  " & body.Name & ": " & fp)
    Next
    
    ' === TEST 3: Whole-part fingerprint (aggregate) ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Whole-part fingerprint ---")
    
    Dim wholeFp As String = ComputeWholePartFingerprint(partDoc)
    Logger.Info("Whole part: " & wholeFp)
    
    ' === TEST 4: Raw values for debugging ===
    Logger.Info("")
    Logger.Info("--- TEST 4: Raw geometry values ---")
    
    For Each body As SurfaceBody In solidBodies
        Logger.Info("Body: " & body.Name)
        Try
            Dim vol As Double = body.Volume(0.001)
            Logger.Info("  Volume(0.001): " & vol.ToString("F8") & " cm³")
        Catch ex As Exception
            Logger.Error("  Volume FAILED: " & ex.Message)
        End Try
        
        Try
            ' SurfaceBody.SurfaceArea does NOT exist! Must iterate faces.
            Dim area As Double = 0
            For Each face As Face In body.Faces
                area += face.Evaluator.Area
            Next
            Logger.Info("  Sum of Face.Evaluator.Area: " & area.ToString("F8") & " cm²")
        Catch ex As Exception
            Logger.Error("  Face area sum FAILED: " & ex.Message)
        End Try
        
        Try
            Dim bb As Box = body.RangeBox
            Dim dx As Double = bb.MaxPoint.X - bb.MinPoint.X
            Dim dy As Double = bb.MaxPoint.Y - bb.MinPoint.Y
            Dim dz As Double = bb.MaxPoint.Z - bb.MinPoint.Z
            Logger.Info("  RangeBox (raw): " & dx.ToString("F6") & " x " & dy.ToString("F6") & " x " & dz.ToString("F6") & " cm")
            
            ' Sorted for orientation independence
            Dim dims() As Double = {dx, dy, dz}
            Array.Sort(dims)
            Logger.Info("  RangeBox (sorted): " & dims(0).ToString("F6") & " x " & dims(1).ToString("F6") & " x " & dims(2).ToString("F6") & " cm")
        Catch ex As Exception
            Logger.Error("  RangeBox FAILED: " & ex.Message)
        End Try
        
        Try
            Logger.Info("  Faces count: " & body.Faces.Count)
        Catch ex As Exception
            Logger.Error("  Faces count FAILED: " & ex.Message)
        End Try
    Next
    
    ' === TEST 5: Sheet metal check ===
    Logger.Info("")
    Logger.Info("--- TEST 5: Sheet metal detection ---")
    
    Dim isSheetMetal As Boolean = False
    Try
        Dim smCompDef As SheetMetalComponentDefinition = CType(compDef, SheetMetalComponentDefinition)
        isSheetMetal = True
        Logger.Info("This IS a sheet metal part")
        
        If smCompDef.HasFlatPattern Then
            Logger.Info("  Has flat pattern: Yes")
            Try
                Dim fpBox As Box = smCompDef.FlatPattern.RangeBox
                Dim fdx As Double = fpBox.MaxPoint.X - fpBox.MinPoint.X
                Dim fdy As Double = fpBox.MaxPoint.Y - fpBox.MinPoint.Y
                Dim fdz As Double = fpBox.MaxPoint.Z - fpBox.MinPoint.Z
                Logger.Info("  Flat pattern box: " & fdx.ToString("F4") & " x " & fdy.ToString("F4") & " x " & fdz.ToString("F4") & " cm")
            Catch ex As Exception
                Logger.Warn("  Could not get flat pattern box: " & ex.Message)
            End Try
        Else
            Logger.Info("  Has flat pattern: No")
        End If
    Catch
        Logger.Info("This is NOT a sheet metal part")
    End Try
    
    ' === TEST 6: MassProperties alternative ===
    Logger.Info("")
    Logger.Info("--- TEST 6: MassProperties (whole document) ---")
    
    Try
        Dim massProps As MassProperties = compDef.MassProperties
        Logger.Info("  MassProperties.Volume: " & massProps.Volume.ToString("F8") & " cm³")
        Logger.Info("  MassProperties.Area: " & massProps.Area.ToString("F8") & " cm²")
        
        ' Check if accessing MassProperties dirties the document
        Logger.Info("  Document dirty after MassProperties: " & partDoc.Dirty.ToString())
    Catch ex As Exception
        Logger.Error("  MassProperties FAILED: " & ex.Message)
    End Try
    
    Logger.Info("")
    Logger.Info("=== Test1_Fingerprint: Complete ===")
End Sub

' Compute fingerprint for a single body
' Format: V:volume|A:area|F:facecount|BB:d1xd2xd3 (sorted for orientation independence)
' NOTE: SurfaceBody.SurfaceArea does NOT exist! Use sum of face areas instead.
Function ComputePartFingerprint(body As SurfaceBody) As String
    Try
        Dim tol As Double = 0.001  ' cm tolerance
        
        ' Volume
        Dim vol As Double = 0
        Try : vol = Math.Round(body.Volume(tol), 6) : Catch : End Try
        
        ' Surface area - must iterate faces (SurfaceBody.SurfaceArea doesn't exist!)
        Dim area As Double = 0
        Try
            For Each face As Face In body.Faces
                area += face.Evaluator.Area
            Next
            area = Math.Round(area, 6)
        Catch : End Try
        
        ' Face count (additional discriminator)
        Dim faceCount As Integer = 0
        Try : faceCount = body.Faces.Count : Catch : End Try
        
        ' Bounding box - sorted for orientation independence
        Dim bb As Box = body.RangeBox
        Dim dims() As Double = {
            Math.Round(bb.MaxPoint.X - bb.MinPoint.X, 4),
            Math.Round(bb.MaxPoint.Y - bb.MinPoint.Y, 4),
            Math.Round(bb.MaxPoint.Z - bb.MinPoint.Z, 4)
        }
        Array.Sort(dims)
        
        Return String.Format("V:{0}|A:{1}|F:{2}|BB:{3}x{4}x{5}", 
            vol.ToString("F6", System.Globalization.CultureInfo.InvariantCulture),
            area.ToString("F6", System.Globalization.CultureInfo.InvariantCulture),
            faceCount,
            dims(0).ToString("F4", System.Globalization.CultureInfo.InvariantCulture),
            dims(1).ToString("F4", System.Globalization.CultureInfo.InvariantCulture),
            dims(2).ToString("F4", System.Globalization.CultureInfo.InvariantCulture))
    Catch ex As Exception
        Return "ERROR:" & ex.Message
    End Try
End Function

' Compute fingerprint for whole part (all solid bodies combined)
' Bodies are sorted by volume for consistency
Function ComputeWholePartFingerprint(partDoc As PartDocument) As String
    Try
        Dim bodyFps As New List(Of String)
        
        For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
            If body.IsSolid Then
                bodyFps.Add(ComputePartFingerprint(body))
            End If
        Next
        
        ' Sort so order doesn't matter
        bodyFps.Sort()
        
        ' Combine with separator
        Return String.Join("|", bodyFps.ToArray())
    Catch ex As Exception
        Return "ERROR:" & ex.Message
    End Try
End Function
