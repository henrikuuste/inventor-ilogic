Imports Inventor
Imports System.Text

' Inventor 2026+ iLogic
' Create "A_" planes from selected planes with offset = 0
'
' Assembly (.iam):
'   - AddFixed is the only supported "Add" for work planes in an assembly. :contentReference[oaicite:3]{index=3}
'   - Position/orientation is controlled via constraints (Flush offset 0). :contentReference[oaicite:4]{index=4}
'   - Size: set AutoResize = True (UI-like sizing). :contentReference[oaicite:5]{index=5}
'
' Part (.ipt):
'   - AddByPlaneAndOffset is supported in parts; not supported in assemblies. :contentReference[oaicite:6]{index=6}
'   - Works also for reference/derived work planes (e.g., skeleton-derived), as long as they are selectable WorkPlanes.

Sub Main()

    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument

    If doc Is Nothing Then
        MessageBox.Show("No active document.", "iLogic")
        Exit Sub
    End If

    Dim sel As SelectSet = doc.SelectSet
    If sel Is Nothing OrElse sel.Count = 0 Then
        MessageBox.Show("Select one or more work planes, then run.", "iLogic")
        Exit Sub
    End If

    Dim created As Integer = 0
    Dim skipped As Integer = 0
    Dim log As New StringBuilder()

    If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        RunInAssembly(CType(doc, AssemblyDocument), sel, created, skipped, log)

    ElseIf doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        RunInPart(CType(doc, PartDocument), sel, created, skipped, log)

    Else
        MessageBox.Show("Supported only in Assembly (.iam) and Part (.ipt).", "iLogic")
        Exit Sub
    End If

    MessageBox.Show(
        "Done." & vbCrLf &
        "Created: " & created & vbCrLf &
        "Skipped: " & skipped & vbCrLf & vbCrLf &
        log.ToString(),
        "iLogic - Create A_ planes (0 offset)"
    )

End Sub

'========================
' Assembly implementation
'========================
Private Sub RunInAssembly(asmDoc As AssemblyDocument, sel As SelectSet, _
                          ByRef created As Integer, ByRef skipped As Integer, log As StringBuilder)

    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

    For Each obj As Object In sel

        Dim srcName As String = Nothing
        Dim srcForConstraint As Object = Nothing   ' WorkPlane or WorkPlaneProxy
        Dim srcForGeom As WorkPlane = Nothing      ' has GetPosition

        ' IMPORTANT: WorkPlaneProxy derives from WorkPlane; check proxy first
        If TypeOf obj Is WorkPlaneProxy Then
            Dim wpP As WorkPlaneProxy = CType(obj, WorkPlaneProxy)
            srcName = wpP.Name
            srcForConstraint = wpP
            srcForGeom = wpP

        ElseIf TypeOf obj Is WorkPlane Then
            Dim wp As WorkPlane = CType(obj, WorkPlane)
            srcName = wp.Name
            srcForConstraint = wp
            srcForGeom = wp

        Else
            skipped += 1
            log.AppendLine("Skipped: not a WorkPlane / WorkPlaneProxy.")
            Continue For
        End If

        ' Get coordinate system in assembly context
        Dim origin As Point = Nothing
        Dim xAxis As UnitVector = Nothing
        Dim yAxis As UnitVector = Nothing
        Try
            srcForGeom.GetPosition(origin, xAxis, yAxis)
        Catch ex As Exception
            skipped += 1
            log.AppendLine("Skipped '" & srcName & "': GetPosition failed: " & ex.Message)
            Continue For
        End Try

        ' Create new assembly plane using AddFixed (supported in assemblies) :contentReference[oaicite:7]{index=7}
        Dim newWp As WorkPlane = Nothing
        Try
            newWp = asmDef.WorkPlanes.AddFixed(origin, xAxis, yAxis)
        Catch ex As Exception
            skipped += 1
            log.AppendLine("Failed AddFixed for '" & srcName & "': " & ex.Message)
            Continue For
        End Try

        ' Constrain Flush offset 0 (keeps it coincident) :contentReference[oaicite:8]{index=8}
        Try
            asmDef.Constraints.AddFlushConstraint(newWp, srcForConstraint, 0)
        Catch ex As Exception
            log.AppendLine("Warning: Flush constraint failed for '" & srcName & "': " & ex.Message)
        End Try

        ' UI-like sizing: AutoResize (recommended for assembly work plane sizing) :contentReference[oaicite:9]{index=9}
        Try
            newWp.AutoResize = True
        Catch
        End Try

        ' Naming
        Dim desired As String = "A_" & srcName
        Dim finalName As String = MakeUniqueWorkPlaneName(asmDef.WorkPlanes, desired)
        Try : newWp.Name = finalName : Catch : End Try

        created += 1
        log.AppendLine("Created: '" & finalName & "' from '" & srcName & "' (assembly)")
    Next

End Sub

'====================
' Part implementation
'====================
Private Sub RunInPart(partDoc As PartDocument, sel As SelectSet, _
                      ByRef created As Integer, ByRef skipped As Integer, log As StringBuilder)

    Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition

    For Each obj As Object In sel

        If Not TypeOf obj Is WorkPlane Then
            skipped += 1
            log.AppendLine("Skipped: not a WorkPlane.")
            Continue For
        End If

        Dim src As WorkPlane = CType(obj, WorkPlane)
        Dim srcName As String = src.Name

        ' Create offset plane 0 (supported in parts; NOT supported in assemblies) :contentReference[oaicite:10]{index=10}
        Dim newWp As WorkPlane = Nothing
        Try
            newWp = partDef.WorkPlanes.AddByPlaneAndOffset(src, 0)
        Catch ex As Exception
            skipped += 1
            log.AppendLine("Failed AddByPlaneAndOffset for '" & srcName & "': " & ex.Message)
            Continue For
        End Try

        ' AutoResize gives sane display size like the UI typically does
        Try
            newWp.AutoResize = True
        Catch
        End Try

        ' Naming
        Dim desired As String = "A_" & srcName
        Dim finalName As String = MakeUniqueWorkPlaneName(partDef.WorkPlanes, desired)
        Try : newWp.Name = finalName : Catch : End Try

        created += 1
        log.AppendLine("Created: '" & finalName & "' from '" & srcName & "' (part)")
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
