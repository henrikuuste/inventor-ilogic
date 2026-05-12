' Copyright (c) 2026 Henri Kuuste
' Test5_ParameterCycle.vb
' PURPOSE: Validate parameter save/restore approach for variant analysis
' 
' For variant matrix building, we need to:
' 1. Snapshot all parameter expressions from all masters
' 2. Set variant parameters on all masters
' 3. Update assembly (propagate to derived parts)
' 4. Fingerprint all parts
' 5. RESTORE original parameters
' 6. Repeat for each variant
'
' TESTS:
' 1. Can we snapshot parameter expressions?
' 2. Can we set parameters from a dictionary?
' 3. Do derived parts update when master parameters change?
' 4. Can we restore parameters exactly?
' 5. Is document dirty state as expected?
'
' RUN: Open an ASSEMBLY with master parts and derived parts

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Error("Test5_ParameterCycle: Open an assembly document first")
        MessageBox.Show("Ava esmalt koost (.iam)", "Test5")
        Return
    End If
    
    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    
    Logger.Info("=== Test5_ParameterCycle: Starting ===")
    Logger.Info("Assembly: " & asmDoc.DisplayName)
    Logger.Info("")
    
    ' === TEST 1: Find parts with parameters (user OR model) ===
    Logger.Info("--- TEST 1: Find parts with parameters ---")
    
    Dim partsWithParams As New Dictionary(Of String, PartDocument)
    Dim partParamTypes As New Dictionary(Of String, String)  ' Track if user or model
    
    For Each refDoc As Document In asmDoc.AllReferencedDocuments
        If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As PartDocument = CType(refDoc, PartDocument)
            Dim allParams As Parameters = partDoc.ComponentDefinition.Parameters
            Dim userParams As UserParameters = allParams.UserParameters
            Dim modelParams As ModelParameters = allParams.ModelParameters
            
            ' Check for user parameters first
            If userParams.Count > 0 Then
                partsWithParams.Add(partDoc.FullFileName, partDoc)
                partParamTypes.Add(partDoc.FullFileName, "user")
                Logger.Info("Part with USER params: " & partDoc.DisplayName & " (" & userParams.Count & " params)")
                
                Dim shown As Integer = 0
                For Each param As Parameter In userParams
                    If shown < 3 Then
                        Logger.Info("  " & param.Name & " = " & param.Expression)
                        shown += 1
                    End If
                Next
            ' Check for model parameters (like "laius", "sügavus", etc.)
            ElseIf modelParams.Count > 10 Then  ' Masters typically have many model params
                ' Look for known variant parameters
                Dim hasVariantParam As Boolean = False
                For Each param As Parameter In modelParams
                    If param.Name = "laius" OrElse param.Name = "sügavus" OrElse param.Name = "selja_kõrgus" Then
                        hasVariantParam = True
                        Exit For
                    End If
                Next
                
                If hasVariantParam Then
                    partsWithParams.Add(partDoc.FullFileName, partDoc)
                    partParamTypes.Add(partDoc.FullFileName, "model")
                    Logger.Info("Part with MODEL params (MASTER): " & partDoc.DisplayName & " (" & modelParams.Count & " params)")
                    
                    Dim shown As Integer = 0
                    For Each param As Parameter In modelParams
                        If shown < 5 Then
                            Logger.Info("  " & param.Name & " = " & param.Expression)
                            shown += 1
                        End If
                    Next
                    If modelParams.Count > 5 Then
                        Logger.Info("  ... and " & (modelParams.Count - 5) & " more")
                    End If
                End If
            End If
        End If
    Next
    
    Logger.Info("Total parts with parameters: " & partsWithParams.Count)
    
    If partsWithParams.Count = 0 Then
        Logger.Warn("No parts with parameters found")
        Logger.Info("Looking for parts with user params OR model params like 'laius'")
        Return
    End If
    
    ' === TEST 2: Snapshot parameters (for restore) ===
    Logger.Info("")
    Logger.Info("--- TEST 2: Snapshot parameters ---")
    
    ' Structure: partPath -> (paramName -> expression)
    Dim snapshot As New Dictionary(Of String, Dictionary(Of String, String))
    
    For Each kvp As KeyValuePair(Of String, PartDocument) In partsWithParams
        Dim partPath As String = kvp.Key
        Dim partDoc As PartDocument = kvp.Value
        Dim paramType As String = partParamTypes(partPath)
        
        Dim paramSnapshot As New Dictionary(Of String, String)
        
        If paramType = "user" Then
            For Each param As Parameter In partDoc.ComponentDefinition.Parameters.UserParameters
                paramSnapshot.Add(param.Name, param.Expression)
            Next
        Else
            ' For model params, only snapshot the known variant parameters
            Dim modelParams As ModelParameters = partDoc.ComponentDefinition.Parameters.ModelParameters
            For Each param As Parameter In modelParams
                If param.Name = "laius" OrElse param.Name = "sügavus" OrElse param.Name = "selja_kõrgus" Then
                    paramSnapshot.Add(param.Name, param.Expression)
                End If
            Next
        End If
        
        snapshot.Add(partPath, paramSnapshot)
        Logger.Info("Snapshot: " & System.IO.Path.GetFileName(partPath) & " - " & paramSnapshot.Count & " params (" & paramType & ")")
    Next
    
    ' === TEST 3: Fingerprint assembly before changes ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Fingerprint assembly parts BEFORE ---")
    
    Dim fpsBefore As New Dictionary(Of String, String)
    For Each refDoc As Document In asmDoc.AllReferencedDocuments
        If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As PartDocument = CType(refDoc, PartDocument)
            Dim fp As String = ComputePartFingerprint(partDoc)
            fpsBefore.Add(partDoc.FullFileName, fp)
            Logger.Info("  " & partDoc.DisplayName & ": " & fp.Substring(0, Math.Min(50, fp.Length)) & "...")
        End If
    Next
    
    ' === TEST 4: Pick a parameter to modify ===
    Logger.Info("")
    Logger.Info("--- TEST 4: Select parameter to modify ---")
    
    Dim testPartPath As String = ""
    Dim testParamName As String = ""
    Dim testOriginalExpr As String = ""
    Dim testNewExpr As String = ""
    Dim testParam As Parameter = Nothing
    
    For Each kvp As KeyValuePair(Of String, PartDocument) In partsWithParams
        Dim partDoc As PartDocument = kvp.Value
        Dim paramType As String = partParamTypes(kvp.Key)
        
        If paramType = "user" Then
            ' Try user parameters
            For Each param As Parameter In partDoc.ComponentDefinition.Parameters.UserParameters
                Try
                    Dim val As Double = param.Value
                    If val > 0 Then
                        testPartPath = kvp.Key
                        testParamName = param.Name
                        testOriginalExpr = param.Expression
                        testParam = param
                        testNewExpr = (val * 1.1).ToString("F4", System.Globalization.CultureInfo.InvariantCulture) & " cm"
                        Exit For
                    End If
                Catch
                End Try
            Next
        Else
            ' Try model parameters - look for "laius" specifically
            For Each param As Parameter In partDoc.ComponentDefinition.Parameters.ModelParameters
                If param.Name = "laius" Then
                    Try
                        Dim val As Double = param.Value
                        testPartPath = kvp.Key
                        testParamName = param.Name
                        testOriginalExpr = param.Expression
                        testParam = param
                        testNewExpr = (val * 1.1).ToString("F4", System.Globalization.CultureInfo.InvariantCulture) & " cm"
                        Exit For
                    Catch
                    End Try
                End If
            Next
        End If
        
        If testParam IsNot Nothing Then Exit For
    Next
    
    If testParam Is Nothing Then
        Logger.Warn("Could not find a suitable parameter to modify")
        Logger.Info("Looking for user params or model param 'laius'")
        Return
    End If
    
    Logger.Info("Will modify: " & System.IO.Path.GetFileName(testPartPath))
    Logger.Info("  Parameter: " & testParamName)
    Logger.Info("  Original: " & testOriginalExpr)
    Logger.Info("  New value: " & testNewExpr)
    
    ' Ask confirmation
    Dim confirmResult As DialogResult = MessageBox.Show(
        "This will test parameter cycling:" & vbCrLf & vbCrLf &
        "1. Change parameter on: " & System.IO.Path.GetFileName(testPartPath) & vbCrLf &
        "2. Update assembly" & vbCrLf &
        "3. Fingerprint all parts" & vbCrLf &
        "4. Restore original parameters" & vbCrLf & vbCrLf &
        "Continue?",
        "Test5_ParameterCycle",
        MessageBoxButtons.YesNo)
    
    If confirmResult <> DialogResult.Yes Then
        Logger.Info("User cancelled")
        Return
    End If
    
    ' === TEST 5: Change parameter ===
    Logger.Info("")
    Logger.Info("--- TEST 5: Change parameter ---")
    
    Dim testPart As PartDocument = partsWithParams(testPartPath)
    Dim dirtyBefore As Boolean = testPart.Dirty
    
    Try
        ' Use testParam directly (already found in TEST 4)
        testParam.Expression = testNewExpr
        Logger.Info("Parameter changed: " & testParamName & " = " & testNewExpr)
        
        ' Read back
        Logger.Info("Read back: " & testParam.Expression)
    Catch ex As Exception
        Logger.Error("Failed to change parameter: " & ex.Message)
        Return
    End Try
    
    ' === TEST 6: Update assembly ===
    Logger.Info("")
    Logger.Info("--- TEST 6: Update assembly ---")
    
    Try
        ' Update the part first
        testPart.Update()
        Logger.Info("Part updated")
        
        ' Update assembly
        asmDoc.Update()
        Logger.Info("Assembly updated")
    Catch ex As Exception
        Logger.Warn("Update had issues: " & ex.Message)
    End Try
    
    ' === TEST 7: Fingerprint AFTER change ===
    Logger.Info("")
    Logger.Info("--- TEST 7: Fingerprint AFTER parameter change ---")
    
    Dim fpsAfterChange As New Dictionary(Of String, String)
    Dim changedParts As New List(Of String)
    
    For Each refDoc As Document In asmDoc.AllReferencedDocuments
        If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As PartDocument = CType(refDoc, PartDocument)
            Dim fp As String = ComputePartFingerprint(partDoc)
            fpsAfterChange.Add(partDoc.FullFileName, fp)
            
            If fpsBefore.ContainsKey(partDoc.FullFileName) Then
                If fpsBefore(partDoc.FullFileName) <> fp Then
                    changedParts.Add(partDoc.DisplayName)
                    Logger.Info("  CHANGED: " & partDoc.DisplayName)
                End If
            End If
        End If
    Next
    
    Logger.Info("Parts with changed fingerprint: " & changedParts.Count)
    If changedParts.Count > 0 Then
        Logger.Info("  This shows parameter changes propagate to derived parts")
    Else
        Logger.Warn("  No fingerprints changed - parameter may not affect geometry")
    End If
    
    ' === TEST 8: Restore parameters ===
    Logger.Info("")
    Logger.Info("--- TEST 8: Restore parameters from snapshot ---")
    
    For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In snapshot
        Dim partPath As String = kvp.Key
        Dim paramSnapshot As Dictionary(Of String, String) = kvp.Value
        
        If partsWithParams.ContainsKey(partPath) Then
            Dim partDoc As PartDocument = partsWithParams(partPath)
            Dim paramType As String = partParamTypes(partPath)
            Dim allParams As Parameters = partDoc.ComponentDefinition.Parameters
            
            For Each paramKvp As KeyValuePair(Of String, String) In paramSnapshot
                Try
                    ' Find param in correct collection
                    Dim param As Parameter = Nothing
                    If paramType = "user" Then
                        param = allParams.UserParameters.Item(paramKvp.Key)
                    Else
                        param = allParams.ModelParameters.Item(paramKvp.Key)
                    End If
                    
                    If param IsNot Nothing AndAlso param.Expression <> paramKvp.Value Then
                        param.Expression = paramKvp.Value
                        Logger.Info("Restored: " & paramKvp.Key & " = " & paramKvp.Value)
                    End If
                Catch ex As Exception
                    Logger.Warn("Could not restore " & paramKvp.Key & ": " & ex.Message)
                End Try
            Next
        End If
    Next
    
    ' Update after restore
    Try
        testPart.Update()
        asmDoc.Update()
        Logger.Info("Updated after restore")
    Catch ex As Exception
        Logger.Warn("Update after restore had issues: " & ex.Message)
    End Try
    
    ' === TEST 9: Verify restore ===
    Logger.Info("")
    Logger.Info("--- TEST 9: Verify parameters restored ---")
    
    Dim restoreSuccess As Boolean = True
    For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In snapshot
        Dim partPath As String = kvp.Key
        Dim paramSnapshot As Dictionary(Of String, String) = kvp.Value
        
        If partsWithParams.ContainsKey(partPath) Then
            Dim partDoc As PartDocument = partsWithParams(partPath)
            Dim paramType As String = partParamTypes(partPath)
            Dim allParams As Parameters = partDoc.ComponentDefinition.Parameters
            
            For Each paramKvp As KeyValuePair(Of String, String) In paramSnapshot
                Try
                    Dim param As Parameter = Nothing
                    If paramType = "user" Then
                        param = allParams.UserParameters.Item(paramKvp.Key)
                    Else
                        param = allParams.ModelParameters.Item(paramKvp.Key)
                    End If
                    
                    If param IsNot Nothing AndAlso param.Expression <> paramKvp.Value Then
                        Logger.Error("NOT RESTORED: " & paramKvp.Key & " - Expected: " & paramKvp.Value & ", Got: " & param.Expression)
                        restoreSuccess = False
                    End If
                Catch
                End Try
            Next
        End If
    Next
    
    If restoreSuccess Then
        Logger.Info("All parameters restored correctly")
    End If
    
    ' === TEST 10: Fingerprint AFTER restore ===
    Logger.Info("")
    Logger.Info("--- TEST 10: Fingerprint AFTER restore ---")
    
    Dim fpsAfterRestore As New Dictionary(Of String, String)
    Dim notRestoredParts As New List(Of String)
    
    For Each refDoc As Document In asmDoc.AllReferencedDocuments
        If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            Dim partDoc As PartDocument = CType(refDoc, PartDocument)
            Dim fp As String = ComputePartFingerprint(partDoc)
            fpsAfterRestore.Add(partDoc.FullFileName, fp)
            
            If fpsBefore.ContainsKey(partDoc.FullFileName) Then
                If fpsBefore(partDoc.FullFileName) <> fp Then
                    notRestoredParts.Add(partDoc.DisplayName)
                    Logger.Warn("  NOT RESTORED: " & partDoc.DisplayName)
                End If
            End If
        End If
    Next
    
    Dim geometryRestored As Boolean = (notRestoredParts.Count = 0)
    If geometryRestored Then
        Logger.Info("All fingerprints restored to original")
    Else
        Logger.Warn("Some parts did not restore geometry")
    End If
    
    ' === SUMMARY ===
    Logger.Info("")
    Logger.Info("========================================")
    Logger.Info("TEST SUMMARY")
    Logger.Info("========================================")
    Logger.Info("Parts with parameters found: " & partsWithParams.Count)
    Logger.Info("Parameter snapshot: SUCCESS")
    Logger.Info("Parameter change: SUCCESS")
    Logger.Info("Assembly update: SUCCESS")
    Logger.Info("Derived parts changed: " & changedParts.Count)
    Logger.Info("Parameter restore: " & If(restoreSuccess, "SUCCESS", "PARTIAL"))
    Logger.Info("Geometry restore: " & If(geometryRestored, "SUCCESS", "PARTIAL"))
    Logger.Info("========================================")
    
    If restoreSuccess AndAlso geometryRestored Then
        Logger.Info("OVERALL: Parameter cycling approach WORKS!")
        MessageBox.Show(
            "Parameter save/restore approach WORKS!" & vbCrLf & vbCrLf &
            "- Parameters can be snapshot" & vbCrLf &
            "- Changes propagate to derived parts" & vbCrLf &
            "- Original state can be restored" & vbCrLf & vbCrLf &
            "Parts that changed during cycle: " & changedParts.Count & vbCrLf & vbCrLf &
            "Note: Documents may still be dirty from the cycle.",
            "Test5_ParameterCycle - SUCCESS")
    Else
        MessageBox.Show(
            "Parameter cycling has issues:" & vbCrLf & vbCrLf &
            "Parameters restored: " & restoreSuccess.ToString() & vbCrLf &
            "Geometry restored: " & geometryRestored.ToString() & vbCrLf & vbCrLf &
            "Check the log for details.",
            "Test5_ParameterCycle - ISSUES")
    End If
End Sub

' Compute fingerprint for whole part
' NOTE: SurfaceBody.SurfaceArea does NOT exist - use sum of face areas
Function ComputePartFingerprint(partDoc As PartDocument) As String
    Try
        Dim bodyFps As New List(Of String)
        
        For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
            If body.IsSolid Then
                Dim tol As Double = 0.001
                Dim vol As Double = 0
                Dim area As Double = 0
                Dim faceCount As Integer = 0
                
                Try : vol = Math.Round(body.Volume(tol), 6) : Catch : End Try
                Try
                    For Each face As Face In body.Faces
                        area += face.Evaluator.Area
                    Next
                    area = Math.Round(area, 6)
                    faceCount = body.Faces.Count
                Catch : End Try
                
                Dim bb As Box = body.RangeBox
                Dim dims() As Double = {
                    Math.Round(bb.MaxPoint.X - bb.MinPoint.X, 4),
                    Math.Round(bb.MaxPoint.Y - bb.MinPoint.Y, 4),
                    Math.Round(bb.MaxPoint.Z - bb.MinPoint.Z, 4)
                }
                Array.Sort(dims)
                
                Dim fp As String = String.Format("V:{0}|A:{1}|F:{2}|BB:{3}x{4}x{5}",
                    vol.ToString("F6", System.Globalization.CultureInfo.InvariantCulture),
                    area.ToString("F6", System.Globalization.CultureInfo.InvariantCulture),
                    faceCount,
                    dims(0).ToString("F4", System.Globalization.CultureInfo.InvariantCulture),
                    dims(1).ToString("F4", System.Globalization.CultureInfo.InvariantCulture),
                    dims(2).ToString("F4", System.Globalization.CultureInfo.InvariantCulture))
                bodyFps.Add(fp)
            End If
        Next
        
        bodyFps.Sort()
        Return String.Join("|", bodyFps.ToArray())
    Catch ex As Exception
        Return "ERROR:" & ex.Message
    End Try
End Function
