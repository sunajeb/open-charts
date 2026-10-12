Attribute VB_Name = "OpenCharts"
' open-charts PowerPoint add-in: the VBA in open-charts.ppam (plus the AppEvents class).
'
' Like the reference product's stub, this holds no chart logic. Every ribbon callback goes to
' Dispatch with the control id; Dispatch forwards a request to the engine (node, through the
' sandbox-approved AppleScriptTask bridge) and draws the native shapes the engine returns.
' The line format is documented in src/addin/format.ts, the request in src/addin/engine.ts.
'
' OCPREFIX is replaced at build time by a per-build prefix shared with customUI14.xml, so the
' callbacks never collide with another add-in's. Constants are numeric on purpose: the project
' references no Office type library.
Option Explicit

Private Const SCRIPT_FILE As String = "open-charts.scpt"
Private Const SCRIPT_HANDLER As String = "engine"
Private Const TITLE As String = "open-charts"

Private gRibbon As Object
Private gEvents As AppEvents
' True once PowerPoint has delivered a selection event: only then can the ribbon grey buttons out.
Private gEventsOk As Boolean
' Re-entrancy guard: drawing changes the selection, which raises the event again.
Private gBusy As Boolean
' Set by the automation entry points: messages are logged instead of shown in a dialog.
Private gSilent As Boolean
Private gLastMessage As String
' The chart last opened with Edit data, so Apply in the editor can update it in place.
Private gEditShape As Object
Private gEditId As String
' Filled by Draw from the engine's answer.
Private gNothing As Boolean
Private gSessionId As String
' Milliseconds spent in the engine call and in drawing, for the timing log.
Private gEngineMs As Single
Private gDrawMs As Single

' ---- Loading --------------------------------------------------------------------------------

Public Sub Auto_Open()
    StartEvents
End Sub

Private Sub StartEvents()
    On Error Resume Next
    If gEvents Is Nothing Then
        Set gEvents = New AppEvents
        Set gEvents.App = Application
    End If
End Sub

' ---- Ribbon callbacks (all route to Dispatch) ----------------------------------------------

Public Sub OCPREFIX_onLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
    StartEvents
End Sub

Public Sub OCPREFIX_onAction(ByVal control As Object)
    Dispatch control.Id
End Sub

' Edit data and Update chart need a selected chart. Until PowerPoint has shown that it reports
' selection changes, they stay enabled and explain themselves when nothing suitable is selected.
Public Sub OCPREFIX_getEnabled(ByVal control As Object, ByRef enabled)
    If Not gEventsOk Then
        enabled = True
    Else
        enabled = Not (FindSelectedChart() Is Nothing)
    End If
End Sub

' Called by AppEvents on every selection change.
Public Sub OnSelectionChange()
    gEventsOk = True
    If gBusy Then Exit Sub
    gBusy = True
    On Error Resume Next
    CheckResized
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "btnEditData"
        gRibbon.InvalidateControl "btnUpdate"
    End If
    gBusy = False
End Sub

' ---- Automation entry points (AppleScript "run VB macro", the add-in server, tests) ---------

' Runs one control id without dialogs, e.g. "chart_stacked", "btnUpdate". The outcome is logged
' by the bridge script (~/Library/Application Support/open-charts/addin.log).
Public Sub OCPREFIX_Run(ByVal controlId As String)
    Dim t As Single
    gSilent = True
    gLastMessage = ""
    t = Timer
    gEngineMs = 0: gDrawMs = 0
    Dispatch controlId
    Report "RESULT " & controlId & " " & IIf(gLastMessage = "", "OK", gLastMessage) & " (" & Format$((Timer - t) * 1000, "0") & " ms)" & Timing()
    gSilent = False
End Sub

' The editor applied an edit session: lay that chart out again with the new data.
Public Sub OCPREFIX_ApplyEdit(ByVal sessionId As String)
    Dim target As Object, t As Single
    t = Timer
    On Error Resume Next
    If sessionId = gEditId And Not gEditShape Is Nothing Then
        If gEditShape.Name <> "" Then Set target = gEditShape
    End If
    On Error GoTo 0
    If target Is Nothing Then Set target = FindSelectedChart()
    If target Is Nothing Then
        Report "APPLIED " & sessionId & " no chart"
        Tell "The edited data is ready. Select the chart in PowerPoint and click Update chart."
        Exit Sub
    End If
    gBusy = True
    gEngineMs = 0: gDrawMs = 0
    Set target = Relayout(target, True, sessionId)
    gBusy = False
    If Not target Is Nothing Then
        Set gEditShape = target
        gEditId = sessionId
    End If
    Report "APPLIED " & sessionId & " " & IIf(target Is Nothing, "failed: " & gLastMessage, "OK") & " (" & Format$((Timer - t) * 1000, "0") & " ms)" & Timing()
End Sub

' Selects the n-th open-charts chart on the active slide (1-based), or nothing for "0" (as a
' click on an empty part of the slide), for scripted tests.
Public Sub OCPREFIX_SelectChart(ByVal n As String)
    Dim shp As Object, k As Long
    On Error GoTo Failed
    If n = "0" Then ActiveWindow.Selection.Unselect: Exit Sub
    For Each shp In ActiveWindow.View.Slide.Shapes
        If IsChart(shp) Then
            k = k + 1
            If k = CLng(Val(n)) Then shp.Select: Exit Sub
        End If
    Next shp
Failed:
End Sub

' Resizes the n-th chart on the active slide by a factor, as a user dragging a corner would.
Public Sub OCPREFIX_ResizeChart(ByVal n As String, ByVal factor As String)
    Dim shp As Object, k As Long
    On Error GoTo Failed
    For Each shp In ActiveWindow.View.Slide.Shapes
        If IsChart(shp) Then
            k = k + 1
            If k = CLng(Val(n)) Then
                shp.LockAspectRatio = 0
                shp.Width = shp.Width * Val(factor)
                shp.Height = shp.Height * Val(factor)
                Exit Sub
            End If
        End If
    Next shp
Failed:
End Sub

' Exports the active presentation to `dest` outside PowerPoint's sandbox: kind "pptx" saves a
' copy, "png" one image per slide (slide1.png ...). PowerPoint writes into its sandbox temp
' folder and the bridge copies the result out.
Public Sub OCPREFIX_ExportCopy(ByVal dest As String, ByVal kind As String)
    Dim tmp As String, i As Long
    On Error GoTo Failed
    tmp = Environ("TMPDIR")
    If Right$(tmp, 1) <> "/" Then tmp = tmp & "/"
    If kind = "png" Then
        tmp = tmp & "open-charts-png"
        If Dir(tmp, vbDirectory) = "" Then MkDir tmp
        For i = 1 To ActivePresentation.Slides.Count
            ActivePresentation.Slides(i).Export tmp & "/slide" & i & ".png", "PNG", 1920, 1080
        Next i
    Else
        tmp = tmp & "open-charts-export.pptx"
        ActivePresentation.SaveCopyAs tmp ' format follows the .pptx extension
    End If
    Report "EXPORT " & AppleScriptTask(SCRIPT_FILE, "copyOut", tmp & vbTab & dest)
    Exit Sub
Failed:
    Report "EXPORT failed: " & Err.Description
End Sub

' ---- Dispatcher ------------------------------------------------------------------------------

Private Sub Dispatch(ByVal id As String)
    ' PowerPoint can deliver a click (or an AppleScript call) while an earlier chart is still being
    ' laid out: the bridge is busy then, so say so instead of failing.
    If gBusy Then Tell "open-charts is still updating a chart. Try again in a moment.": Exit Sub
    On Error GoTo Failed
    gBusy = True
    If Left$(id, 6) = "chart_" Then
        InsertChart Mid$(id, 7)
    ElseIf id = "btnUpdate" Then
        UpdateChart
    ElseIf id = "btnEditData" Then
        EditData
    Else
        Tell "This button is not known to this version of the add-in (" & id & ")."
    End If
    gBusy = False
    Exit Sub
Failed:
    gBusy = False
    Tell "Something went wrong: " & Err.Description
End Sub

Private Sub InsertChart(ByVal itemId As String)
    Dim sld As Object, frame As String, response As String, placeholder As Object, grp As Object
    Set sld = CurrentSlide()
    If sld Is Nothing Then Exit Sub
    Set placeholder = EmptyPlaceholder(sld)
    If placeholder Is Nothing Then
        frame = DefaultFrame()
    Else
        frame = Fmt(placeholder.Left) & "," & Fmt(placeholder.Top) & "," & Fmt(placeholder.Width) & "," & Fmt(placeholder.Height)
    End If
    response = Engine("insert" & vbLf & itemId & vbLf & frame & vbLf & "" & vbLf)
    If response = "" Then Exit Sub
    Set grp = Draw(sld, response)
    If grp Is Nothing Then Exit Sub
    If Not placeholder Is Nothing Then placeholder.Delete
    grp.Select
End Sub

Private Sub UpdateChart()
    Dim chart As Object
    Set chart = SelectedChart()
    If chart Is Nothing Then Exit Sub
    Set chart = Relayout(chart, True, "")
    If Not chart Is Nothing Then chart.Select
End Sub

Private Sub EditData()
    Dim chart As Object, response As String
    Set chart = SelectedChart()
    If chart Is Nothing Then Exit Sub
    response = Engine("edit" & vbLf & "" & vbLf & TagValue(chart, "OC_FRAME") & vbLf & "" & vbLf & SpecOf(chart))
    If response = "" Then Exit Sub
    gSessionId = ""
    Draw Nothing, response
    If gSessionId <> "" Then
        Set gEditShape = chart
        gEditId = gSessionId
    End If
End Sub

' Lays a chart out again where it now is (its group's bounds are its frame) and replaces it in
' place, same stacking order. force=False is the automatic check: the engine answers "nothing to
' do" when neither size nor data changed. Returns the new group (or the old one if unchanged).
Private Function Relayout(ByVal old As Object, ByVal force As Boolean, ByVal sessionId As String) As Object
    Dim response As String, bounds As String, grp As Object
    bounds = Fmt(old.Left) & "," & Fmt(old.Top) & "," & Fmt(old.Width) & "," & Fmt(old.Height)
    response = Engine(IIf(force, "update", "relayout") & vbLf & sessionId & vbLf & TagValue(old, "OC_FRAME") & vbLf & bounds & vbLf & SpecOf(old))
    If response = "" Then Exit Function
    Set grp = Draw(old.Parent, response)
    If gNothing Then Set Relayout = old: Exit Function
    If grp Is Nothing Then Exit Function
    Do While grp.ZOrderPosition > old.ZOrderPosition + 1
        grp.ZOrder 3 ' msoSendBackward
    Loop
    old.Delete
    Set Relayout = grp
End Function

' A chart whose group was resized on the slide is laid out again at its new size (text is never
' stretched); a chart that was only moved just records its new position.
Private Sub CheckResized()
    Dim sld As Object, shp As Object, charts As Collection, f() As String
    Set sld = ActiveWindow.View.Slide
    If sld Is Nothing Then Exit Sub
    Set charts = New Collection
    For Each shp In sld.Shapes
        If IsChart(shp) Then charts.Add shp
    Next shp
    Dim t As Single, result As Object, sel As Object, selectedId As Long, wasSelected As Boolean
    For Each shp In charts
        f = Split(TagValue(shp, "OC_FRAME"), ",")
        If UBound(f) = 3 Then
            If Abs(shp.Width - Val(f(2))) > 0.5 Or Abs(shp.Height - Val(f(3))) > 0.5 Then
                t = Timer
                gEngineMs = 0: gDrawMs = 0
                selectedId = -1
                Set sel = FindSelectedChart()
                If Not sel Is Nothing Then selectedId = sel.Id
                wasSelected = (selectedId = shp.Id)
                Set result = Relayout(shp, False, "")
                If wasSelected And Not result Is Nothing Then result.Select
                Report "RELAYOUT " & IIf(result Is Nothing, "failed: " & gLastMessage, "OK") & " (" & Format$((Timer - t) * 1000, "0") & " ms)" & Timing()
            ElseIf Abs(shp.Left - Val(f(0))) > 0.5 Or Abs(shp.Top - Val(f(1))) > 0.5 Then
                shp.Tags.Add "OC_FRAME", Fmt(shp.Left) & "," & Fmt(shp.Top) & "," & f(2) & "," & f(3)
            End If
        End If
    Next shp
End Sub

' ---- Bridge to the engine -----------------------------------------------------------------

' Returns the engine's response, or "" after telling the user what went wrong.
Private Function Engine(ByVal request As String) As String
    Dim result As String, t As Single
    On Error GoTo NoBridge
    t = Timer
    result = AppleScriptTask(SCRIPT_FILE, SCRIPT_HANDLER, request)
    gEngineMs = gEngineMs + (Timer - t) * 1000
    On Error GoTo 0
    If Left$(result, 3) <> "OC" & vbTab Then
        Tell "The chart engine gave an unexpected answer:" & vbLf & Left$(result, 300)
        Exit Function
    End If
    Engine = result
    Exit Function
NoBridge:
    Tell "open-charts can't reach its chart engine (" & Err.Description & ")." & vbLf & vbLf & _
        "Open Terminal and run:  open-charts addin install" & vbLf & "then restart PowerPoint."
End Function

Private Function Timing() As String
    If gEngineMs > 0 Or gDrawMs > 0 Then Timing = " engine " & Format$(gEngineMs, "0") & " ms, draw " & Format$(gDrawMs, "0") & " ms"
End Function

Private Sub Report(ByVal text As String)
    On Error Resume Next
    Dim ignored As String
    ignored = AppleScriptTask(SCRIPT_FILE, "logLine", text)
End Sub

Private Sub Tell(ByVal text As String)
    gLastMessage = text
    If gSilent Then Exit Sub
    MsgBox text, vbOKOnly, TITLE
End Sub

' ---- Drawing --------------------------------------------------------------------------------

' Draws the response on the slide, groups it and writes the tags. Returns the group, or Nothing.
Private Function Draw(ByVal sld As Object, ByVal response As String) As Object
    Dim lines() As String, f() As String, i As Long, n As Long
    Dim names() As Variant, finalNames() As String, roles() As String
    Dim tagKeys() As String, tagVals() As String, nTags As Long
    Dim groupName As String, alt As String, stamp As String, shp As Object
    Dim t As Single
    t = Timer
    gNothing = False
    lines = Split(response, vbLf)
    ReDim names(0 To UBound(lines) + 1): ReDim finalNames(0 To UBound(lines) + 1): ReDim roles(0 To UBound(lines) + 1)
    ReDim tagKeys(0 To UBound(lines)): ReDim tagVals(0 To UBound(lines))
    stamp = "oc" & Format$(Timer * 100, "0") & "_"

    For i = 0 To UBound(lines)
        If lines(i) <> "" Then
            f = Split(lines(i), vbTab)
            Select Case f(0)
                Case "E"
                    Tell Unescape(f(1))
                    GoTo Cleanup
                Case "N"
                    gNothing = True
                    Exit Function
                Case "M"
                    Tell Unescape(f(1))
                Case "I"
                    gSessionId = f(1)
                Case "G"
                    groupName = Unescape(f(1)): alt = Unescape(f(2))
                Case "T"
                    tagKeys(nTags) = Unescape(f(1)): tagVals(nTags) = Unescape(f(2)): nTags = nTags + 1
                Case "F"
                    ' Invisible frame rectangle: the group's bounds are the chart's frame.
                    Set shp = DrawFrame(sld, f(1))
                    names(n) = stamp & n: shp.Name = names(n)
                    roles(n) = "frame": finalNames(n) = "Frame"
                    n = n + 1
                Case "S"
                    Set shp = DrawShape(sld, f)
                    names(n) = stamp & n: shp.Name = names(n)
                    roles(n) = Unescape(f(20))
                    finalNames(n) = Unescape(f(21))
                    n = n + 1
            End Select
        End If
    Next i
    If n = 0 Then Exit Function

    Dim grp As Object, k As Long, idx As Long
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
            idx = CLng(Mid$(shp.Name, Len(stamp) + 1))
            If roles(idx) <> "" Then shp.Tags.Add "OC_ROLE", roles(idx)
            shp.Name = finalNames(idx)
        Next k
    End If
    Set Draw = grp
    gDrawMs = gDrawMs + (Timer - t) * 1000
    Exit Function
Cleanup:
    ' An error record after some shapes were drawn: remove them again.
    For k = 0 To n - 1
        sld.Shapes(names(k)).Delete
    Next k
End Function

Private Function DrawFrame(ByVal sld As Object, ByVal frame As String) As Object
    Dim v() As String, s As Object
    v = Split(frame, ",")
    Set s = sld.Shapes.AddShape(1, Val(v(0)), Val(v(1)), Val(v(2)), Val(v(3)))
    s.Fill.Visible = 0
    s.Line.Visible = 0
    On Error Resume Next
    s.Shadow.Visible = 0
    Set DrawFrame = s
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
        .Font.NameComplexScript = f(12) ' as the .pptx writer does (a:cs), so digits render the same
        .Font.Size = Val(f(13))
        .Font.Bold = IIf(f(14) = "1", -1, 0)
        ' No kerning, as the engine measured. Decks often kern text of 12 pt and up by default;
        ' Kerning is the size from which pairs are kerned (0 = all sizes), so push it past any text.
        .Font.Kerning = 4000
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

' The open-charts group that is selected (or a shape inside it), else Nothing. Never shows a message.
Private Function FindSelectedChart() As Object
    Dim shp As Object, parentGroup As Object, sel As Object
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    If sel.Type = 2 Or sel.Type = 3 Then ' ppSelectionShapes, ppSelectionText
        Set shp = sel.ShapeRange(1)
        If sel.HasChildShapeRange Then Set shp = sel.ChildShapeRange(1)
        Set parentGroup = shp.ParentGroup
        If Not parentGroup Is Nothing Then Set shp = parentGroup
    End If
    On Error GoTo 0
    If Not shp Is Nothing Then
        If IsChart(shp) Then Set FindSelectedChart = shp
    End If
End Function

' The selected chart, or Nothing after telling the user what to select.
Private Function SelectedChart() As Object
    Set SelectedChart = FindSelectedChart()
    If SelectedChart Is Nothing Then Tell "Select an open-charts chart first (click it once), then choose Edit data or Update chart."
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

' ---- Placement ------------------------------------------------------------------------------

Private Function CurrentSlide() As Object
    On Error Resume Next
    Set CurrentSlide = ActiveWindow.View.Slide
    On Error GoTo 0
    If CurrentSlide Is Nothing Then Tell "Click a slide (Normal view) first, then insert the chart."
End Function

' Where a new chart goes: a selected empty placeholder, else the slide's first empty content
' placeholder (the chart takes its place), else Nothing (use the default chart area).
Private Function EmptyPlaceholder(ByVal sld As Object) As Object
    Dim shp As Object
    On Error Resume Next
    If ActiveWindow.Selection.Type = 2 Then
        Set shp = ActiveWindow.Selection.ShapeRange(1)
        If IsEmptyContentPlaceholder(shp) Then Set EmptyPlaceholder = shp: Exit Function
    End If
    For Each shp In sld.Shapes
        If IsEmptyContentPlaceholder(shp) Then Set EmptyPlaceholder = shp: Exit Function
    Next shp
End Function

Private Function IsEmptyContentPlaceholder(ByVal shp As Object) As Boolean
    Dim kind As Long
    On Error GoTo No
    If shp.Type <> 14 Then Exit Function ' msoPlaceholder
    kind = shp.PlaceholderFormat.Type
    ' ppPlaceholderBody 2, Object 7, Chart 8, Table 12, VerticalObject 16, Picture 18
    If kind <> 2 And kind <> 7 And kind <> 8 And kind <> 12 And kind <> 16 And kind <> 18 Then Exit Function
    If shp.HasTextFrame Then
        If shp.TextFrame.TextRange.Length > 0 Then Exit Function
    End If
    IsEmptyContentPlaceholder = True
No:
End Function

' The chart area of a generated slide (src/deck.ts CHART_FRAME), scaled to this slide size.
Private Function DefaultFrame() As String
    Dim sw As Single, sh As Single
    sw = ActivePresentation.PageSetup.SlideWidth
    sh = ActivePresentation.PageSetup.SlideHeight
    DefaultFrame = Fmt(sw * 0.05) & "," & Fmt(sh * 0.178) & "," & Fmt(sw * 0.9) & "," & Fmt(sh * 0.756)
End Function

' ---- Helpers --------------------------------------------------------------------------------

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
