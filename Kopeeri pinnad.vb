AddVbFile "Lib/UtilsLib.vb"

Imports Inventor

' Inventor 2026+ iLogic
' Create "A_" planes from selected work planes or planar faces with offset = 0
'
' Assembly (.iam):
'   - AddFixed is the only supported "Add" for work planes in an assembly.
'   - Position/orientation is controlled via constraints (Flush offset 0).
'   - Size: set AutoResize = True (UI-like sizing).
'
' Part (.ipt):
'   - AddByPlaneAndOffset is supported in parts; not supported in assemblies.
'   - Works with work planes and planar faces.

Sub Main()

    Dim app As Inventor.Application = ThisApplication
    
    ' Enable immediate logging
    UtilsLib.SetLogger(Logger)
    
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        UtilsLib.LogError("Kopeeri pinnad: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "Kopeeri pinnad")
        Exit Sub
    End If

    Dim sel As SelectSet = doc.SelectSet
    If sel Is Nothing OrElse sel.Count = 0 Then
        UtilsLib.LogWarn("Kopeeri pinnad: No work planes or faces selected.")
        MessageBox.Show("Vali üks või mitu tööpinda või tasapinda, seejärel käivita reegel.", "Kopeeri pinnad")
        Exit Sub
    End If

    Dim created As Integer = 0
    Dim skipped As Integer = 0

    If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        RunInAssembly(CType(doc, AssemblyDocument), sel, created, skipped)

    ElseIf doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        RunInPart(CType(doc, PartDocument), sel, created, skipped)

    Else
        UtilsLib.LogError("Kopeeri pinnad: Unsupported document type.")
        MessageBox.Show("Toetatud ainult koostes (.iam) ja detailis (.ipt).", "Kopeeri pinnad")
        Exit Sub
    End If

    UtilsLib.LogInfo("Kopeeri pinnad: Done. Created: " & created & ", Skipped: " & skipped)

End Sub

'========================
' Assembly implementation
'========================
Private Sub RunInAssembly(asmDoc As AssemblyDocument, sel As SelectSet, _
                          ByRef created As Integer, ByRef skipped As Integer)

    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

    For Each obj As Object In sel

        Dim srcName As String = Nothing
        Dim srcForConstraint As Object = Nothing   ' WorkPlane, WorkPlaneProxy, or FaceProxy
        Dim origin As Point = Nothing
        Dim xAxis As UnitVector = Nothing
        Dim yAxis As UnitVector = Nothing

        ' WorkPlaneProxy derives from WorkPlane; check proxy first
        If TypeOf obj Is WorkPlaneProxy Then
            Dim wpP As WorkPlaneProxy = CType(obj, WorkPlaneProxy)
            srcName = wpP.Name
            srcForConstraint = wpP
            Try
                wpP.GetPosition(origin, xAxis, yAxis)
            Catch ex As Exception
                skipped += 1
                UtilsLib.LogWarn("Kopeeri pinnad: Skipped '" & srcName & "': GetPosition failed: " & ex.Message)
                Continue For
            End Try

        ElseIf TypeOf obj Is WorkPlane Then
            Dim wp As WorkPlane = CType(obj, WorkPlane)
            srcName = wp.Name
            srcForConstraint = wp
            Try
                wp.GetPosition(origin, xAxis, yAxis)
            Catch ex As Exception
                skipped += 1
                UtilsLib.LogWarn("Kopeeri pinnad: Skipped '" & srcName & "': GetPosition failed: " & ex.Message)
                Continue For
            End Try

        ElseIf TypeOf obj Is FaceProxy Then
            Dim faceProxy As FaceProxy = CType(obj, FaceProxy)
            If faceProxy.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Then
                skipped += 1
                UtilsLib.LogWarn("Kopeeri pinnad: Skipped: face is not planar.")
                Continue For
            End If
            Dim plane As Plane = CType(faceProxy.Geometry, Plane)
            srcName = "Pind"
            srcForConstraint = faceProxy
            origin = plane.RootPoint
            xAxis = plane.XAxis
            yAxis = plane.YAxis

        ElseIf TypeOf obj Is Face Then
            Dim face As Face = CType(obj, Face)
            If face.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Then
                skipped += 1
                UtilsLib.LogWarn("Kopeeri pinnad: Skipped: face is not planar.")
                Continue For
            End If
            Dim plane As Plane = CType(face.Geometry, Plane)
            srcName = "Pind"
            srcForConstraint = face
            origin = plane.RootPoint
            xAxis = plane.XAxis
            yAxis = plane.YAxis

        Else
            skipped += 1
            UtilsLib.LogWarn("Kopeeri pinnad: Skipped: not a WorkPlane, WorkPlaneProxy, or planar Face.")
            Continue For
        End If

        Dim newWp As WorkPlane = Nothing
        Try
            newWp = asmDef.WorkPlanes.AddFixed(origin, xAxis, yAxis)
        Catch ex As Exception
            skipped += 1
            UtilsLib.LogWarn("Kopeeri pinnad: Failed AddFixed for '" & srcName & "': " & ex.Message)
            Continue For
        End Try

        Try
            asmDef.Constraints.AddFlushConstraint(newWp, srcForConstraint, 0)
        Catch ex As Exception
            UtilsLib.LogWarn("Kopeeri pinnad: Flush constraint failed for '" & srcName & "': " & ex.Message)
        End Try

        Try
            newWp.AutoResize = True
        Catch
        End Try

        Dim desired As String = "A_" & srcName
        Dim finalName As String = MakeUniqueWorkPlaneName(asmDef.WorkPlanes, desired)
        Try : newWp.Name = finalName : Catch : End Try

        created += 1
        UtilsLib.LogInfo("Kopeeri pinnad: Created '" & finalName & "' from '" & srcName & "' (assembly)")
    Next

End Sub

'====================
' Part implementation
'====================
Private Sub RunInPart(partDoc As PartDocument, sel As SelectSet, _
                      ByRef created As Integer, ByRef skipped As Integer)

    Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition

    For Each obj As Object In sel

        Dim srcName As String = Nothing
        Dim srcForPlane As Object = Nothing   ' WorkPlane or Face

        If TypeOf obj Is WorkPlane Then
            Dim wp As WorkPlane = CType(obj, WorkPlane)
            srcName = wp.Name
            srcForPlane = wp

        ElseIf TypeOf obj Is Face Then
            Dim face As Face = CType(obj, Face)
            If face.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Then
                skipped += 1
                UtilsLib.LogWarn("Kopeeri pinnad: Skipped: face is not planar.")
                Continue For
            End If
            srcName = "Pind"
            srcForPlane = face

        Else
            skipped += 1
            UtilsLib.LogWarn("Kopeeri pinnad: Skipped: not a WorkPlane or planar Face.")
            Continue For
        End If

        Dim newWp As WorkPlane = Nothing
        Try
            newWp = partDef.WorkPlanes.AddByPlaneAndOffset(srcForPlane, 0)
        Catch ex As Exception
            skipped += 1
            UtilsLib.LogWarn("Kopeeri pinnad: Failed AddByPlaneAndOffset for '" & srcName & "': " & ex.Message)
            Continue For
        End Try

        Try
            newWp.AutoResize = True
        Catch
        End Try

        Dim desired As String = "A_" & srcName
        Dim finalName As String = MakeUniqueWorkPlaneName(partDef.WorkPlanes, desired)
        Try : newWp.Name = finalName : Catch : End Try

        created += 1
        UtilsLib.LogInfo("Kopeeri pinnad: Created '" & finalName & "' from '" & srcName & "' (part)")
    Next

End Sub

'====================
' Name helpers
'====================
Private Function MakeUniqueWorkPlaneName(wps As WorkPlanes, desired As String) As String
    If Not WorkPlaneNameExists(wps, desired) Then Return desired

    Dim i As Integer = 1
    Do
        Dim candidate As String = desired & "_" & i.ToString("00")
        If Not WorkPlaneNameExists(wps, candidate) Then Return candidate
        i += 1
        If i > 999 Then Exit Do
    Loop

    Return desired & "_" & Guid.NewGuid().ToString("N").Substring(0, 6)
End Function

Private Function WorkPlaneNameExists(wps As WorkPlanes, testName As String) As Boolean
    For Each wp As WorkPlane In wps
        If String.Equals(wp.Name, testName, StringComparison.OrdinalIgnoreCase) Then Return True
    Next
    Return False
End Function
