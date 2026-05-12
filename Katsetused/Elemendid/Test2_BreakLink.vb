' Copyright (c) 2026 Henri Kuuste
' Test2_BreakLink.vb
' PURPOSE: Validate that BreakLinkToFile() works for creating standalone parts
' 
' CRITICAL: This API is NOT used anywhere in the current codebase!
' The research doc says it exists but needs testing.
'
' TESTS:
' 1. Can we call BreakLinkToFile() on a DerivedPartComponent?
' 2. Does geometry survive after breaking the link?
' 3. Does the part become truly standalone (no ReferencedDocuments)?
' 4. Does fingerprint match before/after?
' 5. What happens if link is already broken or unresolved?
'
' RUN: Open a DERIVED part file (one created via Loo detailid), then run this rule

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("Test2_BreakLink: Open a derived part document first")
        MessageBox.Show("Ava esmalt tuletatud detaili fail (.ipt)", "Test2")
        Return
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    Logger.Info("=== Test2_BreakLink: Starting ===")
    Logger.Info("Part: " & partDoc.DisplayName)
    Logger.Info("Full path: " & partDoc.FullFileName)
    Logger.Info("")
    
    ' === TEST 1: Check if this is a derived part ===
    Logger.Info("--- TEST 1: Detect derived part components ---")
    
    Dim dpcs As DerivedPartComponents = compDef.ReferenceComponents.DerivedPartComponents
    Logger.Info("DerivedPartComponents count: " & dpcs.Count)
    
    If dpcs.Count = 0 Then
        Logger.Warn("This part has NO derived part components - nothing to break")
        Logger.Info("Try opening a part created with 'Loo detailid.vb'")
        
        ' Check ReferencedDocuments anyway
        Logger.Info("")
        Logger.Info("ReferencedDocuments count: " & partDoc.ReferencedDocuments.Count)
        For Each refDoc As Document In partDoc.ReferencedDocuments
            Logger.Info("  Referenced: " & refDoc.FullFileName)
        Next
        Return
    End If
    
    ' Log each derived component
    Dim dpcList As New List(Of DerivedPartComponent)
    For Each dpc As DerivedPartComponent In dpcs
        dpcList.Add(dpc)
        Logger.Info("DerivedPartComponent found:")
        Try
            Logger.Info("  Name: " & dpc.Name)
        Catch : Logger.Info("  Name: (could not read)") : End Try
        Try
            Logger.Info("  Type: " & dpc.Type.ToString())
        Catch : Logger.Info("  Type: (could not read)") : End Try
        Try
            Logger.Info("  ReferencedFile: " & dpc.ReferencedFile.FullFileName)
        Catch ex As Exception
            Logger.Info("  ReferencedFile: (could not read - " & ex.Message & ")")
        End Try
    Next
    
    ' === TEST 2: Check ReferencedDocuments before ===
    Logger.Info("")
    Logger.Info("--- TEST 2: ReferencedDocuments BEFORE break ---")
    
    Dim refDocsBefore As New List(Of String)
    For Each refDoc As Document In partDoc.ReferencedDocuments
        refDocsBefore.Add(refDoc.FullFileName)
        Logger.Info("  " & refDoc.FullFileName)
    Next
    Logger.Info("Total: " & refDocsBefore.Count)
    
    ' === TEST 3: Compute fingerprint BEFORE ===
    Logger.Info("")
    Logger.Info("--- TEST 3: Fingerprint BEFORE break ---")
    
    Dim fpBefore As String = ComputePartFingerprint(partDoc)
    Logger.Info("Fingerprint: " & fpBefore)
    
    ' === TEST 4: Attempt BreakLinkToFile ===
    Logger.Info("")
    Logger.Info("--- TEST 4: Attempting BreakLinkToFile() ---")
    
    ' Ask for confirmation
    Dim confirmResult As DialogResult = MessageBox.Show(
        "This will attempt to break the derivation link." & vbCrLf & vbCrLf &
        "The document will be modified. You can undo with Ctrl+Z." & vbCrLf & vbCrLf &
        "Continue?",
        "Test2_BreakLink",
        MessageBoxButtons.YesNo)
    
    If confirmResult <> DialogResult.Yes Then
        Logger.Info("User cancelled")
        Return
    End If
    
    Dim breakSuccess As Boolean = False
    Dim breakError As String = ""
    
    For Each dpc As DerivedPartComponent In dpcList
        Logger.Info("Attempting BreakLinkToFile on: " & dpc.Name)
        
        Try
            ' THE CRITICAL API CALL
            dpc.BreakLinkToFile()
            
            Logger.Info("  SUCCESS: BreakLinkToFile() completed without error")
            breakSuccess = True
        Catch ex As Exception
            Logger.Error("  FAILED: " & ex.Message)
            breakError = ex.Message
            
            ' Try to get more details
            If ex.Message.Contains("link") OrElse ex.Message.Contains("resolve") Then
                Logger.Info("  HINT: Link may need to be resolved first")
            End If
        End Try
    Next
    
    ' Update document
    Try
        partDoc.Update()
        Logger.Info("Document updated after break attempt")
    Catch ex As Exception
        Logger.Warn("Update after break had issue: " & ex.Message)
    End Try
    
    ' === TEST 5: Check ReferencedDocuments AFTER ===
    Logger.Info("")
    Logger.Info("--- TEST 5: ReferencedDocuments AFTER break ---")
    
    Dim refDocsAfter As New List(Of String)
    For Each refDoc As Document In partDoc.ReferencedDocuments
        refDocsAfter.Add(refDoc.FullFileName)
        Logger.Info("  " & refDoc.FullFileName)
    Next
    Logger.Info("Total: " & refDocsAfter.Count)
    
    ' === TEST 6: Check DerivedPartComponents count AFTER ===
    Logger.Info("")
    Logger.Info("--- TEST 6: DerivedPartComponents AFTER break ---")
    
    Dim dpcsAfter As DerivedPartComponents = compDef.ReferenceComponents.DerivedPartComponents
    Logger.Info("DerivedPartComponents count: " & dpcsAfter.Count)
    
    For Each dpc As DerivedPartComponent In dpcsAfter
        Logger.Info("  Still present: " & dpc.Name)
    Next
    
    ' === TEST 7: Fingerprint AFTER ===
    Logger.Info("")
    Logger.Info("--- TEST 7: Fingerprint AFTER break ---")
    
    Dim fpAfter As String = ComputePartFingerprint(partDoc)
    Logger.Info("Fingerprint: " & fpAfter)
    
    Dim fpMatch As Boolean = (fpBefore = fpAfter)
    If fpMatch Then
        Logger.Info("PASS: Fingerprints MATCH (geometry preserved)")
    Else
        Logger.Error("FAIL: Fingerprints DIFFER (geometry may have changed!)")
        Logger.Info("  Before: " & fpBefore)
        Logger.Info("  After:  " & fpAfter)
    End If
    
    ' === TEST 8: Document state ===
    Logger.Info("")
    Logger.Info("--- TEST 8: Document state ---")
    
    Logger.Info("Document.Dirty: " & partDoc.Dirty.ToString())
    
    ' === SUMMARY ===
    Logger.Info("")
    Logger.Info("========================================")
    Logger.Info("TEST SUMMARY")
    Logger.Info("========================================")
    Logger.Info("BreakLinkToFile() call: " & If(breakSuccess, "SUCCESS", "FAILED - " & breakError))
    Logger.Info("References before: " & refDocsBefore.Count)
    Logger.Info("References after: " & refDocsAfter.Count)
    Logger.Info("DerivedComponents before: " & dpcList.Count)
    Logger.Info("DerivedComponents after: " & dpcsAfter.Count)
    Logger.Info("Geometry preserved: " & If(fpMatch, "YES", "NO"))
    Logger.Info("Document dirty: " & partDoc.Dirty.ToString())
    Logger.Info("========================================")
    
    If breakSuccess AndAlso refDocsAfter.Count = 0 AndAlso fpMatch Then
        Logger.Info("OVERALL: BreakLinkToFile WORKS correctly!")
        MessageBox.Show(
            "BreakLinkToFile() WORKS!" & vbCrLf & vbCrLf &
            "- Derivation link broken" & vbCrLf &
            "- No more referenced documents" & vbCrLf &
            "- Geometry preserved" & vbCrLf & vbCrLf &
            "Document is dirty - save to keep or Ctrl+Z to undo.",
            "Test2_BreakLink - SUCCESS")
    ElseIf breakSuccess Then
        Logger.Warn("OVERALL: BreakLinkToFile called but results unexpected")
        MessageBox.Show(
            "BreakLinkToFile() called without error, but:" & vbCrLf & vbCrLf &
            "References remaining: " & refDocsAfter.Count & vbCrLf &
            "Geometry match: " & fpMatch.ToString() & vbCrLf & vbCrLf &
            "Check the log for details.",
            "Test2_BreakLink - PARTIAL")
    Else
        Logger.Error("OVERALL: BreakLinkToFile FAILED")
        MessageBox.Show(
            "BreakLinkToFile() FAILED:" & vbCrLf & vbCrLf &
            breakError & vbCrLf & vbCrLf &
            "This API may not work in this context." & vbCrLf &
            "Alternative: Delete DerivedPartComponent feature.",
            "Test2_BreakLink - FAILED")
    End If
End Sub

' Compute fingerprint for whole part (all solid bodies)
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
