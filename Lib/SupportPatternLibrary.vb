Option Strict Off
Imports Inventor
Imports System
Imports System.Windows.Forms

Public Module SupportPatternLibrary

    ' Single-line call:
    '   ApplySupports(ThisApplication, ThisDoc.Document, "Põõn")
    '
    ' Defaults derived from baseName:
    '   Occurrence:   baseName & ":1"
    '   Constraint:   baseName & "Offset"
    '   Pattern:      baseName & "Pattern"
    '   Width param:  baseName & "AlaSuurus"
    '   Max param:    baseName & "MaxVahe"
    '   Axis:         baseName & "Telg"
    '
    ' Optional overrides allow non-standard naming if needed.

    Public Sub ApplySupports( _
        invApp As Inventor.Application, _
        asmDoc As Inventor.AssemblyDocument, _
        baseName As String, _
        Optional seedOccName As String = "", _
        Optional offsetConstraintName As String = "", _
        Optional patternName As String = "", _
        Optional totalWidthParamName As String = "", _
        Optional maxSpanParamName As String = "", _
        Optional axisName As String = "")

        If seedOccName = "" Then seedOccName = baseName & ":1"
        If offsetConstraintName = "" Then offsetConstraintName = baseName & "Offset"
        If patternName = "" Then patternName = baseName & "Pattern"
        If totalWidthParamName = "" Then totalWidthParamName = baseName & "AlaSuurus"
        If maxSpanParamName = "" Then maxSpanParamName = baseName & "MaxVahe"
        If axisName = "" Then axisName = baseName & "Telg"

        Dim oDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim uom As UnitsOfMeasure = asmDoc.UnitsOfMeasure

        ' ---- Seed occurrence ----
        Dim seedOcc As ComponentOccurrence
        Try
            seedOcc = oDef.Occurrences.ItemByName(seedOccName)
        Catch
            MessageBox.Show("Seed occurrence not found: " & seedOccName, "ApplySupports")
            Exit Sub
        End Try

        ' ---- Find constraint by name ----
        Dim ac As AssemblyConstraint = Nothing
        For Each c As AssemblyConstraint In oDef.Constraints
            If c.Name = offsetConstraintName Then
                ac = c : Exit For
            End If
        Next
        If ac Is Nothing Then
            MessageBox.Show("Constraint not found: " & offsetConstraintName, "ApplySupports")
            Exit Sub
        End If

        ' ---- Read parameters by name (internal length units) ----
        Dim L As Double
        Dim MaxSpan As Double
        Try
            L = oDef.Parameters.Item(totalWidthParamName).Value
        Catch
            MessageBox.Show("Width parameter not found: " & totalWidthParamName, "ApplySupports")
            Exit Sub
        End Try

        Try
            MaxSpan = oDef.Parameters.Item(maxSpanParamName).Value
        Catch
            MessageBox.Show("Max-span parameter not found: " & maxSpanParamName, "ApplySupports")
            Exit Sub
        End Try

        ' ---- Compute n, spacing, firstOffset (MATCHES your original logic) ----
        ' n = floor(L/MaxSpan)
        ' spacing = L/(n+1)
        ' firstOffset = -L/2 + spacing
        Dim n As Integer = 0
        If MaxSpan > 0 Then
            n = CInt(Math.Floor(L / MaxSpan))
        End If

        Dim spacing As Double = 0
        Dim firstOffset As Double = 0
        If n > 0 Then
            spacing = L / (n + 1)
            firstOffset = -L / 2 + spacing
        End If

        ' ---- Helper: delete an existing pattern by name (if any) ----
        Dim existingPat As OccurrencePattern = Nothing
        For Each p As OccurrencePattern In oDef.OccurrencePatterns
            If p.Name = patternName Then
                existingPat = p : Exit For
            End If
        Next

        ' ---- 0 case: suppress seed and delete pattern ----
        If n <= 0 Then
            If Not seedOcc.Suppressed Then seedOcc.Suppress()
            If Not existingPat Is Nothing Then existingPat.Delete()
            Exit Sub
        End If

        ' ---- n > 0: ensure seed active ----
        If seedOcc.Suppressed Then seedOcc.Unsuppress()

        ' ---- Update offset on mate/flush constraint ----
        If TypeOf ac Is MateConstraint Then
            CType(ac, MateConstraint).Offset.Value = firstOffset
        ElseIf TypeOf ac Is FlushConstraint Then
            CType(ac, FlushConstraint).Offset.Value = firstOffset
        Else
            MessageBox.Show("Constraint '" & offsetConstraintName & "' is not Mate/Flush (no Offset).", "ApplySupports")
            Exit Sub
        End If

        ' ---- Pattern direction axis (work axis by name) ----
        Dim dirAxis As WorkAxis = Nothing
        Try
            dirAxis = oDef.WorkAxes.Item(axisName)
        Catch
            dirAxis = Nothing
        End Try
        If dirAxis Is Nothing Then
            MessageBox.Show("WorkAxis not found in top assembly: " & axisName, "ApplySupports")
            Exit Sub
        End If

        ' ---- Recreate pattern each time (robust + predictable) ----
        If Not existingPat Is Nothing Then existingPat.Delete()

        Dim occs As ObjectCollection = invApp.TransientObjects.CreateObjectCollection()
        occs.Add(seedOcc)

        ' IMPORTANT: AddRectangularPattern signature is ColumnOffset THEN ColumnCount
        ' Passing spacing as a string avoids numeric-unit assumptions.
        Dim spacingExpr As String = uom.GetStringFromValue(spacing, uom.LengthUnits)

        Dim rectPat As RectangularOccurrencePattern = _
            oDef.OccurrencePatterns.AddRectangularPattern( _
                occs, _
                dirAxis, _
                True, _
                spacingExpr, _   ' ColumnOffset
                n _              ' ColumnCount
            )

        rectPat.Name = patternName
    End Sub

End Module
