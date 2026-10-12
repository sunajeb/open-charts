Attribute VB_Name = "OpenCharts"
' open-charts PowerPoint add-in: the one VBA module in open-charts.ppam.
'
' Like the reference product's stub, this holds no chart logic. Every ribbon callback goes to
' Dispatch with the control id; Dispatch forwards the request to the engine (node, through the
' sandbox-approved AppleScriptTask bridge) and draws the native shapes the engine returns.
' The line format is documented in src/addin/format.ts, the request in src/addin/engine.ts.
'
' OCPREFIX is replaced at build time by a per-build prefix shared with customUI14.xml, so the
' callbacks never collide with another add-in's. Constants are numeric on purpose: the project
' references no Office type library.
Option Explicit

Private Const SCRIPT_FILE As String = "open-charts.scpt"
Private Const SCRIPT_HANDLER As String = "engine"

Private gRibbon As Object
' Set by the automation entry points: errors are logged instead of shown in a dialog.
Private gSilent As Boolean
Private gLastMessage As String

' ---- Ribbon callbacks (all route to Dispatch) ----------------------------------------------

Public Sub OCPREFIX_onLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub OCPREFIX_onAction(ByVal control As Object)
    Dispatch control.Id
End Sub

Public Sub OCPREFIX_onGallery(ByVal control As Object, selectedId As String, selectedIndex As Integer)
    Dispatch selectedId
End Sub

' ---- Automation entry points (AppleScript "run VB macro", end-to-end tests) -----------------

' Runs one control id without dialogs, e.g. "chart_stacked", "btnUpdate". The outcome is logged
' by the bridge script (~/Library/Application Support/open-charts/addin.log).
Public Sub OCPREFIX_Run(ByVal controlId As String)
    gSilent = True
    gLastMessage = ""
    Dispatch controlId
    Report "RESULT " & controlId & " " & IIf(gLastMessage = "", "OK", gLastMessage)
    gSilent = False
End Sub

' Selects the n-th open-charts chart on the active slide (1-based), for scripted updates.
Public Sub OCPREFIX_SelectChart(ByVal n As String)
    Dim sld As Object, shp As Object, k As Long
    On Error GoTo Failed
    Set sld = ActiveWindow.View.Slide
    For Each shp In sld.Shapes
        If IsChart(shp) Then
            k = k + 1
            If k = CLng(Val(n)) Then shp.Select: Exit Sub
        End If
    Next shp
Failed:
End Sub

' Saves a copy of the active presentation (format 0 = .pptx, 18 = one PNG per slide) to `dest`
' outside PowerPoint's sandbox: written to the sandbox's temp folder, copied out by the bridge.
Public Sub OCPREFIX_ExportCopy(ByVal dest As String, ByVal fileFormat As String)
    Dim tmp As String
    On Error GoTo Failed
    tmp = Environ("TMPDIR")
    If Right$(tmp, 1) <> "/" Then tmp = tmp & "/"
    tmp = tmp & "open-charts-export" & IIf(fileFormat = "18", "", ".pptx")
    ActivePresentation.SaveCopyAs tmp, CLng(Val(fileFormat))
    Report "EXPORT " & AppleScriptTask(SCRIPT_FILE, "copyOut", tmp & vbTab & dest)
    Exit Sub
Failed:
    Report "EXPORT failed: " & Err.Description
End Sub

' ---- Dispatcher ------------------------------------------------------------------------------

Private Sub Dispatch(ByVal id As String)
    On Error GoTo Failed
    If Left$(id, 6) = "chart_" Then
        InsertChart Mid$(id, 7)
    ElseIf id = "btnUpdate" Then
        UpdateChart
    ElseIf id = "btnEditData" Then
        EditData
    Else
        Tell "Unknown command: " & id
    End If
    Exit Sub
Failed:
    Tell "open-charts: " & Err.Description
End Sub

Private Sub InsertChart(ByVal itemId As String)
    Dim sld As Object, frame As String, response As String
    Set sld = CurrentSlide()
    If sld Is Nothing Then Exit Sub
    frame = DefaultFrame()
    response = Engine("insert" & vbLf & itemId & vbLf & frame & vbLf & "" & vbLf)
    If response = "" Then Exit Sub
    Dim grp As Object
    Set grp = Draw(sld, response)
    If Not grp Is Nothing Then grp.Select
End Sub

Private Sub UpdateChart()
    Dim old As Object, sld As Object, response As String, bounds As String
    Set old = SelectedChart()
    If old Is Nothing Then Exit Sub
    Set sld = old.Parent
    bounds = Fmt(old.Left) & "," & Fmt(old.Top) & "," & Fmt(old.Width) & "," & Fmt(old.Height)
    response = Engine("update" & vbLf & "" & vbLf & TagValue(old, "OC_FRAME") & vbLf & bounds & vbLf & SpecOf(old))
    If response = "" Then Exit Sub
    Dim grp As Object, z As Long
    Set grp = Draw(sld, response)
    If grp Is Nothing Then Exit Sub
    ' Same place in the stacking order, then the old chart goes.
    z = old.ZOrderPosition
    Do While grp.ZOrderPosition > z
        grp.ZOrder 3 ' msoSendBackward
    Loop
    old.Delete
    grp.Select
End Sub

Private Sub EditData()
    Dim chart As Object, response As String
    Set chart = SelectedChart()
    If chart Is Nothing Then Exit Sub
    response = Engine("edit" & vbLf & "" & vbLf & TagValue(chart, "OC_FRAME") & vbLf & "" & vbLf & SpecOf(chart))
    If response <> "" Then Draw Nothing, response
End Sub

' ---- Bridge to the engine -----------------------------------------------------------------

' Returns the engine's response, or "" after telling the user what went wrong.
Private Function Engine(ByVal request As String) As String
    Dim result As String
    On Error GoTo NoBridge
    result = AppleScriptTask(SCRIPT_FILE, SCRIPT_HANDLER, request)
    On Error GoTo 0
    If Left$(result, 3) <> "OC" & vbTab Then
        Tell "open-charts engine did not answer as expected:" & vbLf & Left$(result, 500)
        Exit Function
    End If
    Engine = result
    Exit Function
NoBridge:
    Tell "open-charts cannot reach its engine (" & Err.Description & ")." & vbLf & _
        "Run 'open-charts addin install' in Terminal, then restart PowerPoint."
End Function

Private Sub Report(ByVal text As String)
    On Error Resume Next
    Dim ignored As String
    ignored = AppleScriptTask(SCRIPT_FILE, "logLine", text)
End Sub

Private Sub Tell(ByVal text As String)
    gLastMessage = text
    If gSilent Then Exit Sub
    MsgBox text, vbOKOnly, "open-charts"
End Sub

' ---- Drawing --------------------------------------------------------------------------------

' Draws the response on the slide, groups it and writes the tags. Returns the group, or Nothing.
Private Function Draw(ByVal sld As Object, ByVal response As String) As Object
    Dim lines() As String, f() As String, i As Long, n As Long
    Dim names() As Variant, finalNames() As String, roles() As String
    Dim tagKeys() As String, tagVals() As String, nTags As Long
    Dim groupName As String, alt As String, stamp As String, shp As Object
    lines = Split(response, vbLf)
    ReDim names(0 To UBound(lines)): ReDim finalNames(0 To UBound(lines)): ReDim roles(0 To UBound(lines))
    ReDim tagKeys(0 To UBound(lines)): ReDim tagVals(0 To UBound(lines))
    stamp = "oc" & Format$(Timer * 100, "0") & "_"

    For i = 0 To UBound(lines)
        If lines(i) <> "" Then
            f = Split(lines(i), vbTab)
            Select Case f(0)
                Case "E"
                    Tell Unescape(f(1))
                    GoTo Cleanup
                Case "M"
                    Tell Unescape(f(1))
                Case "G"
                    groupName = Unescape(f(1)): alt = Unescape(f(2))
                Case "T"
                    tagKeys(nTags) = Unescape(f(1)): tagVals(nTags) = Unescape(f(2)): nTags = nTags + 1
                Case "S"
                    Set shp = DrawShape(sld, f)
                    names(n) = stamp & n
                    shp.Name = names(n)
                    roles(n) = Unescape(f(20))
                    finalNames(n) = Unescape(f(21))
                    n = n + 1
            End Select
        End If
    Next i
    If n = 0 Or sld Is Nothing Then Exit Function

    Dim grp As Object, k As Long
    If n = 1 Then
        Set grp = sld.Shapes(names(0))
    Else
        ReDim Preserve names(0 To n - 1)
        Set grp = sld.Shapes.Range(names).Group
    End If
    grp.Name = groupName
    On Error Resume Next
    grp.AlternativeText = alt
    On Error GoTo 0
    For k = 0 To nTags - 1
        grp.Tags.Add tagKeys(k), tagVals(k)
    Next k
    If n > 1 Then
        For k = 1 To grp.GroupItems.Count
            Set shp = grp.GroupItems(k)
            Dim idx As Long
            idx = CLng(Mid$(shp.Name, Len(stamp) + 1))
            If roles(idx) <> "" Then shp.Tags.Add "OC_ROLE", roles(idx)
            shp.Name = finalNames(idx)
        Next k
    End If
    Set Draw = grp
    Exit Function
Cleanup:
    ' An error record after some shapes were drawn: remove them again.
    For k = 0 To n - 1
        sld.Shapes(names(k)).Delete
    Next k
End Function

' Columns (src/addin/format.ts SHAPE_COLUMNS), f(0) = "S":
'  1 kind 2 x 3 y 4 w 5 h 6 fill 7 lineColor 8 lineWidth 9 dash 10 arrowEnd 11 text 12 font
'  13 size 14 bold 15 color 16 align 17 valign 18 wrap 19 autosize 20 role 21 name 22 points
'  23 textGeom
Private Function DrawShape(ByVal sld As Object, f() As String) As Object
    Dim kind As String, x As Single, y As Single, w As Single, h As Single, s As Object
    kind = f(1)
    x = Val(f(2)): y = Val(f(3)): w = Val(f(4)): h = Val(f(5))
    Select Case kind
        Case "line"
            Set s = sld.Shapes.AddLine(x, y, w, h)
            With s.Line
                .Visible = -1
                .ForeColor.RGB = HexColor(f(7))
                .Weight = Val(f(8))
                If f(9) = "1" Then .DashStyle = 4 ' msoLineDash
                If f(10) = "1" Then .EndArrowheadStyle = 2 ' msoArrowheadTriangle
            End With
        Case "text"
            If f(6) = "" And f(7) = "" Then
                Set s = sld.Shapes.AddTextbox(1, x, y, w, h) ' msoTextOrientationHorizontal
            Else
                Set s = sld.Shapes.AddShape(IIf(f(23) = "roundRect", 5, 1), x, y, w, h)
                If f(23) = "roundRect" Then s.Adjustments(1) = 0.5
            End If
            FillAndLine s, f(6), f(7), f(8), "0"
            SetText s, f
            s.Left = x: s.Top = y: s.Width = w: s.Height = h
        Case "path"
            Set s = DrawPath(sld, f(22))
            FillAndLine s, f(6), f(7), f(8), f(9)
        Case Else
            Set s = sld.Shapes.AddShape(AutoShapeType(kind), x, y, w, h)
            FillAndLine s, f(6), f(7), f(8), "0"
    End Select
    Set DrawShape = s
End Function

Private Function AutoShapeType(ByVal kind As String) As Long
    Select Case kind
        Case "ellipse": AutoShapeType = 9 ' msoShapeOval
        Case "roundRect": AutoShapeType = 5 ' msoShapeRoundedRectangle
        Case "diamond": AutoShapeType = 4 ' msoShapeDiamond
        Case "triangle": AutoShapeType = 7 ' msoShapeIsoscelesTriangle
        Case Else: AutoShapeType = 1 ' msoShapeRectangle
    End Select
End Function

Private Sub FillAndLine(ByVal s As Object, ByVal fillHex As String, ByVal lineHex As String, ByVal lineWidth As String, ByVal dash As String)
    If fillHex = "" Then
        s.Fill.Visible = 0
    Else
        s.Fill.Visible = -1
        s.Fill.Solid
        s.Fill.ForeColor.RGB = HexColor(fillHex)
        s.Fill.Transparency = 0
    End If
    If lineHex = "" Then
        s.Line.Visible = 0
    Else
        s.Line.Visible = -1
        s.Line.ForeColor.RGB = HexColor(lineHex)
        s.Line.Weight = Val(lineWidth)
        If dash = "1" Then s.Line.DashStyle = 4
    End If
    On Error Resume Next
    s.Shadow.Visible = 0
End Sub

Private Sub SetText(ByVal s As Object, f() As String)
    Dim tf As Object
    Set tf = s.TextFrame2
    tf.AutoSize = IIf(f(19) = "1", 1, 0) ' msoAutoSizeShapeToFitText / None
    tf.WordWrap = IIf(f(18) = "1", -1, 0)
    tf.MarginLeft = 0: tf.MarginRight = 0: tf.MarginTop = 0: tf.MarginBottom = 0
    Select Case f(17)
        Case "m": tf.VerticalAnchor = 3 ' msoAnchorMiddle
        Case "b": tf.VerticalAnchor = 4 ' msoAnchorBottom
        Case Else: tf.VerticalAnchor = 1 ' msoAnchorTop
    End Select
    ' Paragraphs are separated by CR in PowerPoint text.
    tf.TextRange.Text = Replace(Unescape(f(11)), vbLf, vbCr)
    With tf.TextRange
        .Font.Name = f(12)
        .Font.Size = Val(f(13))
        .Font.Bold = IIf(f(14) = "1", -1, 0)
        .Font.Fill.Visible = -1
        .Font.Fill.ForeColor.RGB = HexColor(f(15))
        Select Case f(16)
            Case "c": .ParagraphFormat.Alignment = 2 ' msoAlignCenter
            Case "r": .ParagraphFormat.Alignment = 3 ' msoAlignRight
            Case Else: .ParagraphFormat.Alignment = 1 ' msoAlignLeft
        End Select
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.SpaceWithin = 1
    End With
End Sub

' points: "x,y x,y ... [z]" per sub-path, sub-paths separated by "|". All sub-paths become one
' freeform; extra sub-paths are joined to the first with a straight seam (e.g. a full ring).
Private Function DrawPath(ByVal sld As Object, ByVal pointsText As String) As Object
    Dim subs() As String, pts() As String, xy() As String, i As Long, j As Long
    Dim ff As Object, started As Boolean, sx As Single, sy As Single, fx As Single, fy As Single
    subs = Split(pointsText, "|")
    For i = 0 To UBound(subs)
        pts = Split(Trim$(subs(i)), " ")
        For j = 0 To UBound(pts)
            If pts(j) = "z" Then
                ff.AddNodes 0, 0, sx, sy ' close: msoSegmentLine, msoEditingAuto
            ElseIf pts(j) <> "" Then
                xy = Split(pts(j), ",")
                If Not started Then
                    Set ff = sld.Shapes.BuildFreeform(1, Val(xy(0)), Val(xy(1))) ' msoEditingCorner
                    started = True
                    fx = Val(xy(0)): fy = Val(xy(1))
                Else
                    ff.AddNodes 0, 0, Val(xy(0)), Val(xy(1))
                End If
                If j = 0 Then sx = Val(xy(0)): sy = Val(xy(1))
            End If
        Next j
    Next i
    If UBound(subs) > 0 Then ff.AddNodes 0, 0, fx, fy
    Set DrawPath = ff.ConvertToShape
End Function

' ---- Reading charts back --------------------------------------------------------------------

Private Function IsChart(ByVal shp As Object) As Boolean
    On Error Resume Next
    IsChart = (shp.Tags("OC_V") <> "")
End Function

' The open-charts group the user selected (or a shape inside it), else Nothing after telling them.
Private Function SelectedChart() As Object
    Dim shp As Object, parentGroup As Object
    On Error Resume Next
    If ActiveWindow.Selection.Type = 2 Then ' ppSelectionShapes
        Set shp = ActiveWindow.Selection.ShapeRange(1)
        If ActiveWindow.Selection.HasChildShapeRange Then Set shp = ActiveWindow.Selection.ChildShapeRange(1)
        Set parentGroup = shp.ParentGroup
        If Not parentGroup Is Nothing Then Set shp = parentGroup
    ElseIf ActiveWindow.Selection.Type = 3 Then ' ppSelectionText: a label inside a chart
        Set shp = ActiveWindow.Selection.ShapeRange(1)
        Set parentGroup = shp.ParentGroup
        If Not parentGroup Is Nothing Then Set shp = parentGroup
    End If
    On Error GoTo 0
    If shp Is Nothing Then
        Tell "Select an open-charts chart first."
    ElseIf Not IsChart(shp) Then
        Tell "The selection is not an open-charts chart."
    Else
        Set SelectedChart = shp
    End If
End Function

Private Function TagValue(ByVal shp As Object, ByVal key As String) As String
    On Error Resume Next
    TagValue = shp.Tags(key)
End Function

' OC_SPEC, reassembled from OC_SPEC_0..n when it was chunked.
Private Function SpecOf(ByVal shp As Object) As String
    Dim s As String, n As Long, i As Long
    s = TagValue(shp, "OC_SPEC")
    If s = "" Then
        n = CLng(Val(TagValue(shp, "OC_SPEC_N")))
        For i = 0 To n - 1
            s = s & TagValue(shp, "OC_SPEC_" & i)
        Next i
    End If
    SpecOf = s
End Function

' ---- Helpers --------------------------------------------------------------------------------

Private Function CurrentSlide() As Object
    On Error Resume Next
    Set CurrentSlide = ActiveWindow.View.Slide
    On Error GoTo 0
    If CurrentSlide Is Nothing Then Tell "Open a slide in Normal view first."
End Function

' The chart area of a generated slide, scaled to this presentation's slide size.
Private Function DefaultFrame() As String
    Dim sw As Single, sh As Single
    sw = ActivePresentation.PageSetup.SlideWidth
    sh = ActivePresentation.PageSetup.SlideHeight
    DefaultFrame = Fmt(sw * 0.05) & "," & Fmt(sh * 0.178) & "," & Fmt(sw * 0.9) & "," & Fmt(sh * 0.756)
End Function

' Number with "." as decimal separator whatever the locale.
Private Function Fmt(ByVal v As Double) As String
    Fmt = Replace(Trim$(Str$(Round(v, 2))), ",", ".")
End Function

Private Function HexColor(ByVal h As String) As Long
    If Len(h) <> 6 Then HexColor = 0: Exit Function
    HexColor = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
End Function

' Reverses the engine's escaping: \\ \t \n \r.
Private Function Unescape(ByVal s As String) As String
    Dim i As Long, c As String, out As String
    If InStr(s, "\") = 0 Then Unescape = s: Exit Function
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c = "\" And i < Len(s) Then
            i = i + 1
            Select Case Mid$(s, i, 1)
                Case "t": out = out & vbTab
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case Else: out = out & Mid$(s, i, 1)
            End Select
        Else
            out = out & c
        End If
        i = i + 1
    Loop
    Unescape = out
End Function
