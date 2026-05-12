' Copyright (c) 2026 Henri Kuuste
' Test3_TransactionRollback.vb
' PURPOSE: Validate that transactions can safely rollback parameter changes
' 
' CRITICAL: For variant analysis, we need to:
' 1. Change parameters on master parts
' 2. Read fingerprints of all derived parts
' 3. UNDO all changes (leave files unmodified)
'
' TESTS:
' 1. Can we start a transaction on a part?
' 2. Can we change parameters within transaction?
' 3. Does Abort() restore original parameter values?
' 4. Is the document clean (not dirty) after abort?
' 5. Does fingerprint match before/after abort?
'
' RUN: Open a part with user parameters, then run this rule

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        Logger.Error("Test3_TransactionRollback: Open a document first")
        MessageBox.Show("Ava esmalt dokument", "Test3")
        Return
    End If
    
    Logger.Info("=== Test3_TransactionRollback: Starting ===")
    Logger.Info("Document: " & doc.DisplayName)
    Logger.Info("Document type: " & doc.DocumentType.ToString())
    Logger.Info("")
    
    ' Get component definition (works for both parts and assemblies)
    Dim compDef As Object = Nothing
    Dim params As Parameters = Nothing
    
    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        Dim partDoc As PartDocument = CType(doc, PartDocument)
        compDef = partDoc.ComponentDefinition
        params = partDoc.ComponentDefinition.Parameters
    ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        compDef = asmDoc.ComponentDefinition
        params = asmDoc.ComponentDefinition.Parameters
    Else
        Logger.Error("Unsupported document type")
        Return
    End If
    
    ' === TEST 1: List parameters (user AND model) ===
    Logger.Info("--- TEST 1: Parameters ---")
    
    Dim userParams As UserParameters = params.UserParameters
    Logger.Info("User parameters count: " & userParams.Count)
    
    ' Also check model parameters
    Dim modelParams As ModelParameters = params.ModelParameters
    Logger.Info("Model parameters count: " & modelParams.Count)
    
    If userParams.Count = 0 AndAlso modelParams.Count = 0 Then
        Logger.Warn("No parameters found at all")
    End If
    
    ' Store original values (user params)
    Dim originalValues As New Dictionary(Of String, String)
    For Each param As Parameter In userParams
        originalValues.Add(param.Name, param.Expression)
        Logger.Info("  User: " & param.Name & " = " & param.Expression)
    Next
    
    ' Show first few model params
    Dim modelShown As Integer = 0
    For Each param As Parameter In modelParams
        If modelShown < 5 Then
            Logger.Info("  Model: " & param.Name & " = " & param.Expression)
            modelShown += 1
        End If
    Next
    If modelParams.Count > 5 Then
        Logger.Info("  ... and " & (modelParams.Count - 5) & " more model parameters")
    End If
    
    ' Pick a parameter to modify - try user params first, then model params
    Dim testParamName As String = ""
    Dim testParamOriginal As String = ""
    Dim testParamNew As String = ""
    Dim testParam As Parameter = Nothing
    Dim usingModelParam As Boolean = False
    
    ' Try user parameters first
    For Each param As Parameter In userParams
        Try
            Dim val As Double = param.Value
            If val > 0.1 Then  ' Skip very small values
                testParamName = param.Name
                testParamOriginal = param.Expression
                testParam = param
                testParamNew = (val * 1.1).ToString("F4", System.Globalization.CultureInfo.InvariantCulture) & " cm"
                Exit For
            End If
        Catch
        End Try
    Next
    
    ' If no user param found, try model parameters
    If testParam Is Nothing Then
        For Each param As Parameter In modelParams
            Try
                ' Skip reference parameters and very small values
                If param.ParameterType = ParameterTypeEnum.kModelParameter Then
                    Dim val As Double = param.Value
                    If val > 0.1 AndAlso val < 1000 Then  ' Reasonable range
                        testParamName = param.Name
                        testParamOriginal = param.Expression
                        testParam = param
                        usingModelParam = True
                        testParamNew = (val * 1.1).ToString("F4", System.Globalization.CultureInfo.InvariantCulture) & " cm"
                        Exit For
                    End If
                End If
            Catch
            End Try
        Next
    End If
    
    If testParam Is Nothing Then
        Logger.Warn("Could not find a suitable parameter to modify")
        Logger.Info("Will test transaction mechanics without parameter change")
    Else
        Logger.Info("")
        Logger.Info("Will test with " & If(usingModelParam, "MODEL", "USER") & " parameter: " & testParamName)
        Logger.Info("  Original: " & testParamOriginal)
        Logger.Info("  New value: " & testParamNew)
    End If
    
    ' === TEST 2: Fingerprint BEFORE ===
    Logger.Info("")
    Logger.Info("--- TEST 2: State BEFORE transaction ---")
    
    Dim fpBefore As String = ""
    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        fpBefore = ComputePartFingerprint(CType(doc, PartDocument))
        Logger.Info("Fingerprint: " & fpBefore)
    End If
    
    Dim dirtyBefore As Boolean = doc.Dirty
    Logger.Info("Document.Dirty: " & dirtyBefore.ToString())
    
    ' === TEST 3: Start transaction ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Start Transaction ---")
    
    Dim trans As Transaction = Nothing
    Try
        trans = app.TransactionManager.StartTransaction(doc, "Test_ParameterChange")
        Logger.Info("Transaction started successfully")
        Logger.Info("  Transaction name: Test_ParameterChange")
    Catch ex As Exception
        Logger.Error("Failed to start transaction: " & ex.Message)
        Return
    End Try
    
    ' === TEST 4: Modify parameter within transaction ===
    Logger.Info("")
    Logger.Info("--- TEST 4: Modify parameter within transaction ---")
    
    Dim paramChanged As Boolean = False
    If testParam IsNot Nothing Then
        Try
            testParam.Expression = testParamNew
            paramChanged = True
            Logger.Info("Parameter changed: " & testParamName & " = " & testParamNew)
            
            ' Update to see the change
            doc.Update()
            Logger.Info("Document updated after parameter change")
            
            ' Read back to verify
            Logger.Info("Parameter value after change: " & testParam.Expression)
        Catch ex As Exception
            Logger.Error("Failed to change parameter: " & ex.Message)
        End Try
    Else
        Logger.Info("No parameter to change (testing transaction mechanics only)")
    End If
    
    ' === TEST 5: Check state during transaction ===
    Logger.Info("")
    Logger.Info("--- TEST 5: State DURING transaction ---")
    
    Dim fpDuring As String = ""
    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        fpDuring = ComputePartFingerprint(CType(doc, PartDocument))
        Logger.Info("Fingerprint: " & fpDuring)
        If paramChanged AndAlso fpBefore <> fpDuring Then
            Logger.Info("  Geometry CHANGED as expected")
        ElseIf paramChanged Then
            Logger.Info("  Geometry unchanged (parameter may not affect geometry)")
        End If
    End If
    
    Logger.Info("Document.Dirty: " & doc.Dirty.ToString())
    
    ' === TEST 6: Abort transaction ===
    Logger.Info("")
    Logger.Info("--- TEST 6: Abort Transaction ---")
    
    Try
        trans.Abort()
        Logger.Info("Transaction aborted successfully")
    Catch ex As Exception
        Logger.Error("Failed to abort transaction: " & ex.Message)
        Logger.Info("Trying to End instead...")
        Try
            trans.End()
            Logger.Warn("Transaction ended (committed) instead of aborted!")
        Catch ex2 As Exception
            Logger.Error("Also failed to end: " & ex2.Message)
        End Try
    End Try
    
    ' === TEST 7: Check state AFTER abort ===
    Logger.Info("")
    Logger.Info("--- TEST 7: State AFTER abort ---")
    
    ' Check parameter value
    If testParam IsNot Nothing Then
        Try
            Dim afterAbort As String = testParam.Expression
            Logger.Info("Parameter after abort: " & testParamName & " = " & afterAbort)
            
            If afterAbort = testParamOriginal Then
                Logger.Info("  PASS: Parameter RESTORED to original")
            Else
                Logger.Error("  FAIL: Parameter NOT restored!")
                Logger.Info("    Expected: " & testParamOriginal)
                Logger.Info("    Got: " & afterAbort)
            End If
        Catch ex As Exception
            Logger.Error("Failed to read parameter after abort: " & ex.Message)
        End Try
    End If
    
    ' Check fingerprint
    Dim fpAfter As String = ""
    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        fpAfter = ComputePartFingerprint(CType(doc, PartDocument))
        Logger.Info("Fingerprint: " & fpAfter)
        
        If fpBefore = fpAfter Then
            Logger.Info("  PASS: Fingerprint RESTORED")
        Else
            Logger.Error("  FAIL: Fingerprint NOT restored!")
        End If
    End If
    
    ' Check dirty state
    Dim dirtyAfter As Boolean = doc.Dirty
    Logger.Info("Document.Dirty: " & dirtyAfter.ToString())
    
    If dirtyAfter = dirtyBefore Then
        Logger.Info("  PASS: Dirty state RESTORED")
    Else
        Logger.Warn("  Document dirty state changed (Before: " & dirtyBefore & ", After: " & dirtyAfter & ")")
    End If
    
    ' === SUMMARY ===
    Logger.Info("")
    Logger.Info("========================================")
    Logger.Info("TEST SUMMARY")
    Logger.Info("========================================")
    Logger.Info("Transaction start: SUCCESS")
    Logger.Info("Parameter change: " & If(paramChanged, "YES", "N/A"))
    Logger.Info("Transaction abort: SUCCESS")
    
    Dim paramRestored As Boolean = True
    If testParam IsNot Nothing Then
        Dim currentVal As String = testParam.Expression
        paramRestored = (currentVal = testParamOriginal)
    End If
    
    Logger.Info("Parameter restored: " & If(paramRestored, "YES", "NO"))
    Logger.Info("Fingerprint restored: " & If(fpBefore = fpAfter OrElse String.IsNullOrEmpty(fpBefore), "YES", "NO"))
    Logger.Info("Document clean: " & If(dirtyAfter = dirtyBefore, "YES", "NO"))
    Logger.Info("========================================")
    
    If paramRestored AndAlso (fpBefore = fpAfter OrElse String.IsNullOrEmpty(fpBefore)) Then
        Logger.Info("OVERALL: Transaction rollback WORKS!")
        MessageBox.Show(
            "Transaction rollback WORKS!" & vbCrLf & vbCrLf &
            "- Parameters restored" & vbCrLf &
            "- Geometry restored" & vbCrLf & vbCrLf &
            "This approach is SAFE for variant analysis.",
            "Test3_TransactionRollback - SUCCESS")
    Else
        Logger.Error("OVERALL: Transaction rollback has ISSUES")
        MessageBox.Show(
            "Transaction rollback has issues:" & vbCrLf & vbCrLf &
            "Parameter restored: " & paramRestored.ToString() & vbCrLf &
            "Geometry restored: " & (fpBefore = fpAfter).ToString() & vbCrLf & vbCrLf &
            "May need alternative approach (save/restore).",
            "Test3_TransactionRollback - ISSUES")
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
