Attribute VB_Name = "OpenCharts"
' open-charts PowerPoint add-in: the VBA in open-charts.ppam (plus the AppEvents class).
'
' Like the reference product's stub, this holds no chart logic. Every ribbon callback goes to
' Dispatch with the control id; Dispatch forwards a request to the engine (node, through the
' sandbox-approved AppleScriptTask bridge) and draws the native shapes the engine returns.
' On Windows (#If Mac ... #Else) the same requests go over HTTP to the engine's local server
' (open-charts.exe, started here when it is not running); everything else is shared.
' The line format is documented in src/addin/format.ts, the request in src/addin/engine.ts.
'
' OCPREFIX is replaced at build time by a per-build prefix shared with customUI14.xml, so the
' callbacks never collide with another add-in's. Constants are numeric on purpose: the project
' references no Office type library.
Option Explicit

Private Const SCRIPT_FILE As String = "open-charts.scpt"
Private Const SCRIPT_HANDLER As String = "engine"
Private Const TITLE As String = "open-charts"
' The command that reinstalls the add-in; the build writes the launcher path of this install here.
Private Const REINSTALL As String = """$HOME/Library/Application Support/open-charts/bin/open-charts"" addin install"
' Chart types each command applies to, "|type|type|": the build writes them from src/core/spec.ts
' OPTIONS (totals, cagr / difference, scaleGroup), so the ribbon and right-click menu offer a command
' only where the engine draws it.
Private Const TOTALS_TYPES As String = "|stacked|"
Private Const ARROW_TYPES As String = "|stacked|"
Private Const SCALE_TYPES As String = "|stacked|"
' The directions each chart type can grow in ("stacked:up,down,left,right;line:up,left,right;..."),
' written by the build from DIRECTIONS in src/core/spec.ts.
Private Const DIRECTIONS As String = ""
#If Mac Then
#Else
' Windows: the engine's local server (127.0.0.1 only) and where install.ps1 puts the engine.
Private Const ENGINE_URL As String = "http://127.0.0.1:4174/api/addin/"
Private Const ENGINE_EXE As String = "\open-charts\open-charts.exe"
' The server's per-install token ("X-OC-Token: <hex>", src/addin/token.ts) and the user's settings.
Private Const TOKEN_FILE As String = "\open-charts\token"
Private Const SETTINGS_FILE As String = "\open-charts\settings.json"
Private Const WIN_REINSTALL As String = "Run install.ps1 from the open-charts-win folder again (README, Install on Windows), then restart PowerPoint."
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
#End If

Private gRibbon As Object
Private gEvents As AppEvents
Private gEngineStarted As Boolean
' True once PowerPoint has delivered a selection event: only then can the ribbon grey buttons out.
Private gEventsOk As Boolean
' Re-entrancy guard: drawing changes the selection, which raises the event again.
Private gBusy As Boolean
' True while gBusy is held by the automatic selection-change work (not a user action). A click that
' arrives then (PowerPoint delivers it while that work waits on the engine) runs at once and sets
' gAbortAuto; the automatic work then stops without drawing. A user action is never refused or queued.
Private gAutoBusy As Boolean
Private gAbortAuto As Boolean
' Set by the automation entry points: messages are logged instead of shown in a dialog.
Private gSilent As Boolean
' The user turned verbose logging on (`open-charts addin verbose on`): Verbose lines are logged too.
Private gVerbose As Boolean
Private gLastMessage As String
' The chart last opened with Edit data, so Apply in the editor can update it in place.
Private gEditShape As Object
Private gEditId As String
' The element mini toolbar: what NotifySelection last reported ("" = nothing selected).
Private gSelSig As String
' Filled by Draw from the engine's answer.
Private gNothing As Boolean
Private gSessionId As String
' Milliseconds spent in the engine call and in drawing, for the timing log.
Private gEngineMs As Single
Private gDrawMs As Single
' Finer phases for the timing log (" theme 12 ms, shapes 210 ms, ..."), summed per name.
Private gPhaseNames As Collection
Private gPhaseMs As Collection
' While ApplyEdit runs, Report collects lines for its answer instead of logging them itself.
Private gCollect As Boolean
Private gCollected As String
' The deck as the automatic agenda update last saw it (AgendaSignature).
Private gAgendaSig As String
' Agenda slides: the layout new ones use when the deck has it, and their title.
Private Const AGENDA_LAYOUT As String = "open-charts agenda"
Private Const AGENDA_TITLE As String = "Agenda"
' The direction buttons: the direction the next inserted chart grows in ("" = up, the default), and
' the selected chart's type and direction as of the last selection change ("" = no chart selected).
Private gNextDir As String
' How often PowerPoint asked a right-click item whether to show (OCPREFIX_Probe reports it).
Private gVisibleCalls As Long
Private gSelType As String
Private gSelDir As String

' ---- Loading --------------------------------------------------------------------------------

Public Sub Auto_Open()
    StartEvents
End Sub

' Starts the warm engine (the add-in's local server) in the background when the ribbon loads (the
' first document window), so the first chart inserted does not wait for node to start. Never blocks:
' the bridge only spawns it. Not from Auto_Open: an AppleScriptTask while PowerPoint is still
' launching makes `register add in` (addin install with PowerPoint closed) fail with -43.
Private Sub StartEngine()
    Dim ignored As String
    If gEngineStarted Then Exit Sub
    gEngineStarted = True
    On Error Resume Next
#If Mac Then
    ignored = AppleScriptTask(SCRIPT_FILE, "startServer", "")
    gVerbose = (ignored = "OK verbose")
#Else
    gVerbose = InStr(Replace(WinReadText(Environ("LOCALAPPDATA") & SETTINGS_FILE), " ", ""), """verbose"":true") > 0
    WinStartServer 0
#End If
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
    StartEngine
End Sub

Public Sub OCPREFIX_onAction(ByVal control As Object)
    Dispatch control.Id
End Sub

' The chart gallery: the tile's id is "chart_<item>".
Public Sub OCPREFIX_onGallery(ByVal control As Object, ByVal selectedId As String, ByVal selectedIndex As Integer)
    Dispatch selectedId
End Sub

' Direction buttons (dirUp, dirDown, dirLeft, dirRight): turn the selected chart, or pick the
' direction of the next insert. A toggle button reports its new state; the click is what counts.
Public Sub OCPREFIX_onDirection(ByVal control As Object, ByVal pressed As Boolean)
    Dispatch control.Id
End Sub

' Pressed: the selected chart's direction, or (nothing selected) the direction picked for the next insert.
Public Sub OCPREFIX_getDirPressed(ByVal control As Object, ByRef pressed)
    If Not gEventsOk Then RefreshSelection
    If gSelType <> "" Then
        pressed = (gSelDir = DirOf(control.Id))
    Else
        pressed = (gNextDir = DirOf(control.Id))
    End If
End Sub

' With a chart selected, only the directions its type can grow in; with none, all (for the next insert).
Public Sub OCPREFIX_getDirEnabled(ByVal control As Object, ByRef enabled)
    If Not gEventsOk Then RefreshSelection
    If gSelType = "" Then
        enabled = True
    Else
        enabled = CanGrow(gSelType, DirOf(control.Id))
    End If
End Sub

' The buttons need a selected chart (CommandApplies). Until PowerPoint has shown that it reports
' selection changes, they stay enabled and explain themselves when nothing suitable is selected.
' A piece of an ungrouped chart also enables Edit data and Update chart, so the click can say why
' it is not a chart.
Public Sub OCPREFIX_getEnabled(ByVal control As Object, ByRef enabled)
    If Not gEventsOk Then
        enabled = True
    ElseIf control.Id = "btnEditData" Or control.Id = "btnUpdate" Then
        enabled = Not (FindSelectedChart() Is Nothing) Or UngroupedSelected()
    Else
        enabled = CommandApplies(control.Id)
    End If
End Sub

' Right-click menu items (customUI contextMenus): shown only on open-charts charts, and only the
' commands that apply to what is selected.
Public Sub OCPREFIX_getVisible(ByVal control As Object, ByRef visible)
    gVisibleCalls = gVisibleCalls + 1
    visible = False
    On Error Resume Next
    visible = CommandApplies(CommandId(control.Id))
End Sub

' Called by AppEvents on every selection change.
Public Sub OnSelectionChange()
    gEventsOk = True
    If gBusy Then Exit Sub
    gBusy = True
    gAutoBusy = True
    gAbortAuto = False
    On Error Resume Next
    CheckResized
    AgendaAutoCheck
    RefreshSelection
    NotifySelection
    If Not gRibbon Is Nothing Then
        Dim c As Variant
        For Each c In Array("btnEditData", "btnUpdate", "btnTotals", "btnCagr", "btnDifference", "btnSameScale", "btnUnscale", "btnPlainText", "btnResetLabel")
            gRibbon.InvalidateControl c
        Next c
    End If
    InvalidateDirections
    gAutoBusy = False
    gAbortAuto = False
    gBusy = False
End Sub

' Reads the selected chart's type and direction (from its OC_SPEC) for the direction buttons.
Private Sub RefreshSelection()
    Dim chart As Object, spec As String
    gSelType = ""
    gSelDir = ""
    On Error Resume Next
    Set chart = FindSelectedChart()
    If chart Is Nothing Then Exit Sub
    spec = SpecOf(chart)
    gSelType = JsonString(spec, "type")
    gSelDir = JsonString(spec, "direction")
    If gSelDir = "" And JsonString(spec, "orientation") = "horizontal" Then gSelDir = "right"
    If gSelDir = "" Then gSelDir = "up"
End Sub

Private Sub InvalidateDirections()
    If gRibbon Is Nothing Then Exit Sub
    gRibbon.InvalidateControl "dirUp"
    gRibbon.InvalidateControl "dirDown"
    gRibbon.InvalidateControl "dirLeft"
    gRibbon.InvalidateControl "dirRight"
End Sub

' "dirLeft" -> "left".
Private Function DirOf(ByVal controlId As String) As String
    DirOf = LCase$(Mid$(controlId, 4))
End Function

' Whether a chart type can grow in a direction (DIRECTIONS; a type it does not list: up only).
Private Function CanGrow(ByVal chartType As String, ByVal dir As String) As Boolean
    Dim entry As Variant
    For Each entry In Split(DIRECTIONS, ";")
        If Left$(entry, Len(chartType) + 1) = chartType & ":" Then
            CanGrow = InStr(1, "," & Mid$(entry, Len(chartType) + 2) & ",", "," & dir & ",") > 0
            Exit Function
        End If
    Next entry
    CanGrow = (dir = "up")
End Function

' The first string value of `key` in compact JSON (OC_SPEC is written by JSON.stringify), or "".
Private Function JsonString(ByVal json As String, ByVal key As String) As String
    Dim at As Long, q As Long
    at = InStr(1, json, """" & key & """:""")
    If at = 0 Then Exit Function
    at = at + Len(key) + 4
    q = InStr(at, json, """")
    If q > at Then JsonString = Mid$(json, at, q - at)
End Function

' The element mini toolbar (src/addin/toolbar.ts) and the drag handles (src/addin/handles.ts): tells
' the add-in server which elements of an open-charts chart are selected (their OC_ROLE tags) and
' where, that the whole chart is selected (no roles), or that nothing is any more. Only when that
' changes, and never for other selections: one AppleScriptTask that starts a background curl and
' returns at once. Silent: an older bridge script without the handler is ignored.
Private Sub NotifySelection()
    Dim request As String, sig As String
    On Error Resume Next
#If Mac Then
    request = SelectionRequest(sig)
#End If
    If sig = gSelSig Then Exit Sub
    gSelSig = sig
#If Mac Then
    request = "seq=" & Fmt3(Timer) & vbLf & request
    AppleScriptTask SCRIPT_FILE, "elementSelection", request
#End If
    ' Windows: no element toolbar or drag handles yet (they are panels of the Mac pop-up helper), so
    ' no call at all: a selection change never waits on the engine.
End Sub

' The selection report (src/addin/toolbar.ts): "select", the selected roles (tab-separated; empty
' for the whole chart or nothing), then for a chart its context, screen= of the selection, view= of
' the chart group (its corners on screen, for the drag handles) and its OC_SPEC. `sig` identifies
' what is reported ("" = nothing selected).
Private Function SelectionRequest(ByRef sig As String) As String
    Dim sel As Object, chart As Object, shp As Object, roles As String, r As String, i As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    If sel.HasChildShapeRange Then
        Set chart = sel.ChildShapeRange(1).ParentGroup
        If Not chart Is Nothing Then
            If Not IsChart(chart) Then Set chart = Nothing
        End If
    ElseIf sel.Type = 2 Then ' ppSelectionShapes: the whole chart (one group)
        If sel.ShapeRange.Count = 1 Then
            Set chart = sel.ShapeRange(1)
            If Not IsChart(chart) Then Set chart = Nothing
        End If
    End If
    sig = ""
    If chart Is Nothing Then
        SelectionRequest = "select" & vbLf & "" & vbLf
        Exit Function
    End If
    x0 = 1E+30: y0 = 1E+30: x1 = -1E+30: y1 = -1E+30
    If sel.HasChildShapeRange Then
        For i = 1 To sel.ChildShapeRange.Count
            Set shp = sel.ChildShapeRange(i)
            r = TagValue(shp, "OC_ROLE")
            If r <> "" Then
                roles = roles & IIf(roles = "", "", vbTab) & r
                If shp.Left < x0 Then x0 = shp.Left
                If shp.Top < y0 Then y0 = shp.Top
                If shp.Left + shp.Width > x1 Then x1 = shp.Left + shp.Width
                If shp.Top + shp.Height > y1 Then y1 = shp.Top + shp.Height
            End If
        Next i
    End If
    If roles = "" Then x0 = chart.Left: y0 = chart.Top: x1 = chart.Left + chart.Width: y1 = chart.Top + chart.Height
    sig = chart.Id & "|" & IIf(roles = "", "*", roles) & "|" & Fmt(x0) & "," & Fmt(y0) & "," & Fmt(x1) & "," & Fmt(y1)
    SelectionRequest = "select" & vbLf & roles & vbLf & ChartContext(chart) & ScreenLinesFor(x0, y0, x1 - x0, y1 - y0) & ViewLine(chart) & SpecOf(chart)
End Function

' "view=left,top,right,bottom": the chart group's corners through PointsToScreenPixelsX/Y (the drag
' handles map the slide to the screen with it), or nothing when PowerPoint cannot tell.
Private Function ViewLine(ByVal chart As Object) As String
    Dim w As Object, s As String
    On Error Resume Next
    Set w = ActiveWindow
    If w Is Nothing Then Exit Function
    Err.Clear
    s = Fmt(w.PointsToScreenPixelsX(chart.Left)) & "," & Fmt(w.PointsToScreenPixelsY(chart.Top)) & "," & _
        Fmt(w.PointsToScreenPixelsX(chart.Left + chart.Width)) & "," & Fmt(w.PointsToScreenPixelsY(chart.Top + chart.Height))
    If Err.Number = 0 Then ViewLine = "view=" & s & vbLf
    Err.Clear
End Function

' ---- Entry points of the add-in server (run VB macro / COM Run; not on the ribbon) --------

' The selection report now (the add-in server re-measures the chart on screen after a scroll, zoom
' or window move: the drag handles hid themselves until then). Never shows a message.
Public Function OCPREFIX_SelectionInfo() As String
    Dim sig As String
    On Error Resume Next
    OCPREFIX_SelectionInfo = SelectionRequest(sig)
End Function

' The editor applied an edit session: the add-in server has laid the chart out with the new data for
' `context` (ChartContext: frame, bounds, theme, same-scale group) and passes the response, so this
' only draws it in place of the chart (and re-lays out its same-scale partners). If the chart's
' context is not `context` ("*" accepts any), it draws nothing and answers its context, and the
' server lays it out again for that. Silent: the server runs this and the editor shows the outcome,
' so no dialog may block it. Answers its log lines (the server writes them: a Report here would be
' another AppleScriptTask, about 0.2 s) and, after drawing, the new chart's context.
Public Function OCPREFIX_ApplyEdit(ByVal sessionId As String, ByVal context As String, ByVal response As String) As String
    Dim target As Object, t As Single, tp As Single, here As String, outcome As String
    t = Timer
    ResetTiming
    gSilent = True
    gLastMessage = ""
    gCollected = "": gCollect = True
    On Error Resume Next
    If sessionId = gEditId And Not gEditShape Is Nothing Then
        If gEditShape.Name <> "" Then Set target = gEditShape
    End If
    On Error GoTo Failed
    If target Is Nothing Then Set target = FindSelectedChart()
    If target Is Nothing Then
        OCPREFIX_ApplyEdit = "APPLIED " & sessionId & " no chart"
        GoTo Done
    End If
    here = ChartContext(target)
    If response = "" Or (context <> "*" And context <> here) Then
        OCPREFIX_ApplyEdit = ContextLine(here)
        GoTo Done
    End If
    gBusy = True
    NewUndoEntry
    Phase "find", t
    Dim sld As Object, oldGroup As String
    Set sld = target.Parent
    oldGroup = TagValue(target, "OC_SGROUP")
    Set target = ReplaceChart(target, response)
    tp = Timer
    If Not target Is Nothing Then SyncAfter sld, target, oldGroup
    Phase "sync", tp
    gBusy = False
    If Not target Is Nothing Then
        Set gEditShape = target
        gEditId = sessionId
        tp = Timer
        here = ChartContext(target)
        Phase "context", tp
    End If
    outcome = gCollected & "APPLIED " & sessionId & " " & IIf(target Is Nothing, "failed: " & gLastMessage, "OK") & " (" & Format$((Timer - t) * 1000, "0") & " ms)" & Timing()
    If Not target Is Nothing Then outcome = outcome & vbLf & ContextLine(here)
    OCPREFIX_ApplyEdit = outcome
    GoTo Done
Failed:
    ' Never leave an error dialog (it would block the server's call) or the busy flag behind.
    gBusy = False
    OCPREFIX_ApplyEdit = gCollected & "APPLIED " & sessionId & " failed: " & Err.Description
Done:
    gCollect = False
    gSilent = False
    ' When the macro started and ended (seconds since midnight), for the server's timing log.
    OCPREFIX_ApplyEdit = OCPREFIX_ApplyEdit & vbLf & "TIMER" & vbTab & Fmt3(t) & vbTab & Fmt3(Timer)
End Function

Private Function Fmt3(ByVal v As Double) As String
    Fmt3 = Replace(Format$(v, "0.000"), ",", ".")
End Function

' What a layout depends on besides the data: the request lines between the session id and the spec.
Private Function ChartContext(ByVal chart As Object) As String
    ChartContext = TagValue(chart, "OC_FRAME") & vbLf & Fmt(chart.Left) & "," & Fmt(chart.Top) & "," & Fmt(chart.Width) & "," & Fmt(chart.Height) & vbLf & _
        ThemeLine(chart.Parent) & ScaleLine(chart)
End Function

' Where the chart is on screen, for the Edit data pop-up (src/addin/popup.ts ChartWhere):
' "screen=left,top,right,bottom" through PointsToScreenPixelsX/Y, and "window=left,top,width,height,
' zoom" of the document window. Either is left out when PowerPoint cannot tell; the server only
' trusts the corners when they fall inside a PowerPoint window, else it places the pop-up by the window.
Private Function ScreenLines(ByVal chart As Object) As String
    ScreenLines = ScreenLinesFor(chart.Left, chart.Top, chart.Width, chart.Height)
End Function

' The same for a rectangle on the slide (pt).
Private Function ScreenLinesFor(ByVal x As Double, ByVal y As Double, ByVal cx As Double, ByVal cy As Double) As String
    Dim w As Object, s As String
    On Error Resume Next
    Set w = ActiveWindow
    If w Is Nothing Then Exit Function
    Err.Clear
    s = Fmt(w.PointsToScreenPixelsX(x)) & "," & Fmt(w.PointsToScreenPixelsY(y)) & "," & _
        Fmt(w.PointsToScreenPixelsX(x + cx)) & "," & Fmt(w.PointsToScreenPixelsY(y + cy))
    If Err.Number = 0 Then ScreenLinesFor = "screen=" & s & vbLf
    Err.Clear
    s = Fmt(w.Left) & "," & Fmt(w.Top) & "," & Fmt(w.Width) & "," & Fmt(w.Height) & "," & Fmt(w.View.Zoom)
    If Err.Number = 0 Then ScreenLinesFor = ScreenLinesFor & "window=" & s & vbLf
    Err.Clear
End Function

' A context as one answer line (src/addin/server.ts parseApplyAnswer).
Private Function ContextLine(ByVal context As String) As String
    ContextLine = "CONTEXT" & vbTab & Replace(context, vbLf, vbTab)
End Function

' ---- Dispatcher ------------------------------------------------------------------------------

Private Sub Dispatch(ByVal id As String)
    ' PowerPoint can deliver a click (or an AppleScript call) while an earlier chart is still being
    ' laid out: the bridge is busy then, so say so instead of failing.
    If gBusy Then
        ' Background selection work: this click wins; that work stops without drawing. A user action
        ' in progress: say so.
        If Not gAutoBusy Then Tell "open-charts is still updating a chart. Try again in a moment.": Exit Sub
        gAbortAuto = True
    End If
    On Error GoTo Failed
    gBusy = True
    id = CommandId(id)
    ' One user action, one undo entry (NewUndoEntry).
    NewUndoEntry
    If Left$(id, 6) = "chart_" Then
        InsertChart Mid$(id, 7)
    ElseIf id = "dirUp" Or id = "dirDown" Or id = "dirLeft" Or id = "dirRight" Then
        ChangeDirection DirOf(id)
    ElseIf id = "btnUpdate" Then
        UpdateChart
    ElseIf id = "btnEditData" Then
        EditData
    ElseIf Left$(id, 9) = "btnAgenda" Then
        AgendaCommand Mid$(id, 10)
    ElseIf id = "btnPlainText" Then
        ConvertToPlainText
    ElseIf id = "btnResetLabel" Then
        ChartCommand "unpin"
    ElseIf id = "btnTotals" Then
        ChartCommand "totals"
    ElseIf id = "btnCagr" Then
        ChartCommand "cagr"
    ElseIf id = "btnDifference" Then
        ChartCommand "difference"
    ElseIf id = "btnSameScale" Then
        SameScale True
    ElseIf id = "btnUnscale" Then
        SameScale False
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
    ' A direction picked beforehand applies to this one insert ("stacked@left"); the engine ignores a
    ' direction the chart type cannot take.
    If gNextDir <> "" Then itemId = itemId & "@" & gNextDir
    gNextDir = ""
    InvalidateDirections
    response = Engine("insert" & vbLf & itemId & vbLf & frame & vbLf & "" & vbLf & ThemeLine(sld))
    If response = "" Then Exit Sub
    Set grp = Draw(sld, response)
    If grp Is Nothing Then Exit Sub
    If Not placeholder Is Nothing Then placeholder.Delete
    grp.Select
End Sub

' A direction button: the selected chart is laid out again in its place, growing `dir`; with no
' chart selected, the button picks the direction of the next insert (a second click clears it).
Private Sub ChangeDirection(ByVal dir As String)
    Dim chart As Object, sld As Object, oldGroup As String, response As String
    Set chart = FindSelectedChart()
    If chart Is Nothing Then
        If UngroupedSelected() Then
            Tell "This chart was ungrouped, so it can't be edited as a chart. Undo (" & ChrW(&H2318) & "Z) the ungroup or insert it again."
            Exit Sub
        End If
        gNextDir = IIf(gNextDir = dir Or dir = "up", "", dir)
        InvalidateDirections
        Exit Sub
    End If
    Set sld = chart.Parent
    oldGroup = TagValue(chart, "OC_SGROUP")
    response = Engine("direction" & vbLf & dir & vbLf & ChartContext(chart) & SpecOf(chart))
    If response = "" Then Exit Sub
    Set chart = ReplaceChart(chart, response)
    If chart Is Nothing Then Exit Sub
    SyncAfter sld, chart, oldGroup
    chart.Select
    RefreshSelection
    InvalidateDirections
End Sub

Private Sub UpdateChart()
    Dim chart As Object, sld As Object, oldGroup As String
    Set chart = SelectedChart()
    If chart Is Nothing Then Exit Sub
    Set sld = chart.Parent
    oldGroup = TagValue(chart, "OC_SGROUP")
    Set chart = Relayout(chart, True)
    If chart Is Nothing Then Exit Sub
    SyncAfter sld, chart, oldGroup
    chart.Select
End Sub

Private Sub EditData()
    Dim chart As Object, response As String
    Set chart = SelectedChart()
    If chart Is Nothing Then Exit Sub
    ' The chart's context too: the server lays an Apply out for it by itself (ApplyEdit).
    ' And where it is on screen, so the Edit data pop-up opens next to it (ScreenLines).
    response = Engine("edit" & vbLf & "" & vbLf & ChartContext(chart) & ScreenLines(chart) & SpecOf(chart))
    If response = "" Then Exit Sub
    gSessionId = ""
    Draw Nothing, response
    If gSessionId <> "" Then
        Set gEditShape = chart
        gEditId = gSessionId
    End If
End Sub

' Convert to plain text: the selected label keeps the text it shows now, which data edits no longer
' rewrite; it stays in the chart and moves with it (engine op "detach").
Private Sub ConvertToPlainText()
    Dim chart As Object, sel As Object, shp As Object, role As String, response As String
    Set chart = SelectedChart()
    If chart Is Nothing Then Exit Sub
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    If sel.HasChildShapeRange Then Set shp = sel.ChildShapeRange(1)
    On Error GoTo 0
    If Not shp Is Nothing Then role = TagValue(shp, "OC_ROLE")
    If role = "" Then Tell "Click one label of the chart (so only it is selected), then convert it to plain text.": Exit Sub
    response = Engine("detach" & vbLf & role & vbLf & ChartContext(chart) & SpecOf(chart))
    If response = "" Then Exit Sub
    Set chart = ReplaceChart(chart, response)
    If Not chart Is Nothing Then chart.Select
End Sub

' ---- Chart commands (ribbon "Chart tools" group, right-click menus) ----------------------------

' A right-click menu item's id is "ctx<menu>_<ribbon button id>": the same command as the button.
Private Function CommandId(ByVal id As String) As String
    If Left$(id, 3) = "ctx" And InStr(id, "_") > 0 Then CommandId = Mid$(id, InStr(id, "_") + 1) Else CommandId = id
End Function

' One user action, one undo entry: PowerPoint for Windows starts a new entry here, so Ctrl+Z undoes
' the whole command and nothing before it. Late bound, and ignored where PowerPoint has no such
' method (docs/USER-GUIDE.md "Undo" says what PowerPoint for Mac does).
Private Sub NewUndoEntry()
    On Error Resume Next
    CallByName Application, "StartNewUndoEntry", VbMethod
End Sub

' Whether ribbon button `id` (or its right-click item) applies to the selection.
Private Function CommandApplies(ByVal id As String) As Boolean
    Dim chart As Object, shp As Variant
    On Error Resume Next
    Select Case id
        Case "btnSameScale"
            CommandApplies = (ScaleCandidates().Count >= 2)
            Exit Function
        Case "btnUnscale"
            For Each shp In SelectedCharts()
                If TagValue(shp, "OC_SGROUP") <> "" Then CommandApplies = True: Exit Function
            Next shp
            Exit Function
    End Select
    ' The other commands work on one chart: not with several selected.
    If SelectedCharts().Count > 1 Then Exit Function
    Set chart = FindSelectedChart()
    If chart Is Nothing Then Exit Function
    Select Case id
        Case "btnEditData", "btnUpdate", "btnResetLabel": CommandApplies = True
        Case "btnPlainText": CommandApplies = (Left$(SelectedRoles(chart), 6) = "label:")
        Case "btnTotals": CommandApplies = (InStr(TOTALS_TYPES, "|" & ChartTypeOf(chart) & "|") > 0)
        Case "btnCagr", "btnDifference": CommandApplies = (InStr(ARROW_TYPES, "|" & ChartTypeOf(chart) & "|") > 0)
    End Select
End Function

' The chart type in a chart's OC_SPEC ("" if it cannot be read).
Private Function ChartTypeOf(ByVal chart As Object) As String
    Dim s As String, at As Long, e As Long
    s = SpecOf(chart)
    at = InStr(s, """type"":""")
    If at = 0 Then Exit Function
    at = at + 8
    e = InStr(at, s, """")
    If e > at Then ChartTypeOf = Mid$(s, at, e - at)
End Function

' OC_ROLE values of the selected pieces of `chart`, comma-separated; "" when the whole chart is selected.
Private Function SelectedRoles(ByVal chart As Object) As String
    Dim sel As Object, i As Long, r As String, out As String, shp As Object
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    ' Compared, never negated with Not: PowerPoint for Mac can answer True as 1, and Not 1 is -2 (still True).
    If sel.HasChildShapeRange = 0 Then Exit Function
    For i = 1 To sel.ChildShapeRange.Count
        Set shp = sel.ChildShapeRange(i)
        r = TagValue(shp, "OC_ROLE")
        If r <> "" And r <> "frame" Then out = out & IIf(out = "", "", ",") & r
    Next i
    SelectedRoles = out
End Function

' The open-charts charts selected on the slide (several with Shift-click), or the one whose pieces are.
Private Function SelectedCharts() As Collection
    Dim sel As Object, i As Long, out As Collection, one As Object
    Set out = New Collection
    Set SelectedCharts = out
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    If sel.Type <> 2 And sel.Type <> 3 Then Exit Function ' ppSelectionShapes, ppSelectionText
    If sel.HasChildShapeRange Then
        Set one = FindSelectedChart()
        If Not one Is Nothing Then out.Add one
        Exit Function
    End If
    For i = 1 To sel.ShapeRange.Count
        If IsChart(sel.ShapeRange(i)) Then out.Add sel.ShapeRange(i)
    Next i
End Function

' The selected charts whose type can share a scale.
Private Function ScaleCandidates() As Collection
    Dim shp As Variant, out As Collection
    Set out = New Collection
    For Each shp In SelectedCharts()
        If InStr(SCALE_TYPES, "|" & ChartTypeOf(shp) & "|") > 0 Then out.Add shp
    Next shp
    Set ScaleCandidates = out
End Function

' Totals, CAGR / difference arrow, reset label position: the engine changes the chart's options or
' overrides (core/commands.ts runCommand), keeping formatting changed in PowerPoint, and the chart
' is drawn again in place. Its same-scale partners follow (totals can change its scale).
Private Sub ChartCommand(ByVal cmd As String)
    Dim chart As Object, sld As Object, response As String, oldGroup As String
    Set chart = SelectedChart()
    If chart Is Nothing Then Exit Sub
    Set sld = chart.Parent
    oldGroup = TagValue(chart, "OC_SGROUP")
    response = Engine("command" & vbLf & cmd & vbLf & ChartContext(chart) & "roles=" & SelectedRoles(chart) & vbLf & NativeLines(chart) & SpecOf(chart))
    If response = "" Then Exit Sub
    Set chart = ReplaceChart(chart, response)
    If chart Is Nothing Then Exit Sub
    SyncAfter sld, chart, oldGroup
    chart.Select
End Sub

' Same scale (joining = True): the selected charts get one same-scale group (options.scaleGroup in
' each OC_SPEC; an existing group of one of them is reused) and are laid out again so one value
' unit is equally long in all of them; later edits, resizes and updates keep it (SyncAfter).
' Remove from same scale (joining = False): the selected charts leave their groups; the charts left
' behind are laid out again at their remaining group's scale.
Private Sub SameScale(ByVal joining As Boolean)
    Dim charts As Collection, shp As Variant, g As String, sld As Object, response As String
    Dim oldGroups As Collection, old As String, done As Object, ids As String, n As Long, lastId As Long
    If joining Then
        Set charts = ScaleCandidates()
        If charts.Count < 2 Then
            Tell "Select two or more open-charts column, bar, waterfall, line, area or combination charts on this slide (Shift-click), then choose Same scale."
            Exit Sub
        End If
    Else
        Set charts = New Collection
        For Each shp In SelectedCharts()
            If TagValue(shp, "OC_SGROUP") <> "" Then charts.Add shp
        Next shp
        If charts.Count = 0 Then Tell "Select a chart that shares its scale with others, then choose Remove from same scale.": Exit Sub
    End If
    Set sld = charts(1).Parent
    If joining Then
        For Each shp In charts
            If g = "" Then g = TagValue(shp, "OC_SGROUP")
        Next shp
        If g = "" Then g = NewScaleGroup(sld)
    End If
    Set oldGroups = New Collection
    For Each shp In charts
        old = TagValue(shp, "OC_SGROUP")
        If old <> "" And old <> g Then oldGroups.Add old
        response = Engine("command" & vbLf & "scaleGroup" & vbLf & TagValue(shp, "OC_FRAME") & vbLf & _
            Fmt(shp.Left) & "," & Fmt(shp.Top) & "," & Fmt(shp.Width) & "," & Fmt(shp.Height) & vbLf & _
            ThemeLine(sld) & ScaleLineFor(shp, g) & "group=" & g & vbLf & NativeLines(shp) & SpecOf(shp))
        If response = "" Then Exit For
        Set done = ReplaceChart(shp, response)
        If done Is Nothing Then Exit For
        n = n + 1: lastId = done.Id: ids = ids & "|" & done.Id
    Next shp
    ' Every member now carries its own scale: lay the others out again at the group's smallest.
    If joining And n > 0 Then SyncGroup sld, g, lastId
    For Each shp In oldGroups
        SyncGroup sld, CStr(shp), 0
    Next shp
    Report "SAMESCALE " & IIf(joining, g, "-") & " " & n & " of " & charts.Count & " charts"
    ' Select the charts again: the group's members (the sync replaced them), or the ones that left.
    If n > 0 Then SelectCharts sld, g, ids & "|"
End Sub

' Selects the charts of same-scale group g on the slide, or (g = "") those with an id in "|id|id|".
Private Sub SelectCharts(ByVal sld As Object, ByVal g As String, ByVal ids As String)
    Dim shp As Object, first As Boolean, hit As Boolean
    On Error Resume Next
    first = True
    For Each shp In sld.Shapes
        If IsChart(shp) Then
            If g = "" Then hit = (InStr(ids, "|" & shp.Id & "|") > 0) Else hit = (TagValue(shp, "OC_SGROUP") = g)
            If hit Then
                shp.Select IIf(first, -1, 0) ' msoTrue replaces the selection, msoFalse adds to it
                first = False
            End If
        End If
    Next shp
End Sub

' A same-scale group name no chart on the slide uses yet ("scale1", "scale2", ...).
Private Function NewScaleGroup(ByVal sld As Object) As String
    Dim i As Long, shp As Object, used As String
    For Each shp In sld.Shapes
        If IsChart(shp) Then used = used & "|" & TagValue(shp, "OC_SGROUP")
    Next shp
    used = used & "|"
    Do
        i = i + 1
    Loop While InStr(used, "|scale" & i & "|") > 0
    NewScaleGroup = "scale" & i
End Function

' Lays a chart out again where it now is (its group's bounds are its frame) and replaces it in
' place, same stacking order. force=False is the automatic check: the engine answers "nothing to
' do" when neither size nor data changed. Returns the new group (or the old one if unchanged).
Private Function Relayout(ByVal old As Object, ByVal force As Boolean) As Object
    Dim response As String, request As String, tp As Single
    tp = Timer
    request = IIf(force, "update", "relayout") & vbLf & "" & vbLf & ChartContext(old) & NativeLines(old) & SpecOf(old)
    Phase "read", tp
    response = Engine(request)
    If response = "" Then Exit Function
    ' The automatic check gives way to a user action that arrived while it waited on the engine.
    If Not force And gAbortAuto Then Exit Function
    Set Relayout = ReplaceChart(old, response)
End Function

'  Formatting changed with PowerPoint's own tools is kept: the request carries the chart's OC_BASE
' (its formatting as last drawn) and what its shapes look like now ("base=" and "native=" lines);
' the engine turns every difference into an override in the spec (core/overrides.ts mergeNativeEdits).
' "" for a chart drawn before OC_BASE existed.
Private Function NativeLines(ByVal chart As Object) As String
    Dim base As String, k As Long, shp As Object, role As String, out As String, item As String
    base = ChunkedTag(chart, "OC_BASE")
    If base = "" Then Exit Function
    For k = 1 To chart.GroupItems.Count
        Set shp = chart.GroupItems(k)
        role = TagValue(shp, "OC_ROLE")
        If role <> "" And role <> "frame" Then
            item = ShapeEntry(shp, role)
            If item <> "" Then out = out & IIf(out = "", "", ",") & item
        End If
    Next k
    NativeLines = "base=" & base & vbLf & "native=[" & out & "]" & vbLf
End Function

' One shape as {"role", "kind": text | mark | line, "props"} (core/overrides.ts RenderedEntry).
Private Function ShapeEntry(ByVal shp As Object, ByVal role As String) As String
    Dim p As String, kind As String, tr As Object, o As Long
    On Error GoTo Failed
    If shp.Type = 9 Then ' msoLine
        kind = "line"
        p = """line"":" & JsonStr(ColorOf(shp.Line.ForeColor, shp.Line.Visible)) & ",""lineWidth"":" & Fmt3(shp.Line.Weight)
    Else
        kind = "mark"
        p = """fill"":" & JsonStr(ColorOf(shp.Fill.ForeColor, shp.Fill.Visible)) & ",""outline"":" & JsonStr(ColorOf(shp.Line.ForeColor, shp.Line.Visible))
        If shp.Line.Visible Then p = p & ",""outlineWidth"":" & Fmt3(shp.Line.Weight)
        If shp.HasTextFrame Then
            If shp.TextFrame2.HasText Then
                kind = "text"
                Set tr = shp.TextFrame2.TextRange
                o = shp.TextFrame2.Orientation
                p = p & ",""text"":" & JsonStr(Replace(tr.Text, vbCr, vbLf)) & ",""font"":" & JsonStr(tr.Font.Name) & _
                    ",""size"":" & Fmt3(tr.Font.Size) & ",""bold"":" & IIf(tr.Font.Bold = -1, "true", "false") & _
                    ",""color"":" & JsonStr(ColorOf(tr.Font.Fill.ForeColor, -1)) & ",""rotate"":" & IIf(o = 2, "-90", IIf(o = 3, "90", "0"))
            End If
        End If
    End If
    ShapeEntry = "{""role"":" & JsonStr(role) & ",""kind"":""" & kind & """,""props"":{" & p & "}}"
    Exit Function
Failed:
    ShapeEntry = ""
End Function

' A colour as "RRGGBB", or "RRGGBB@accent1-lighter40" for a theme colour; "none" when not visible.
Private Function ColorOf(ByVal cf As Object, ByVal visible As Long) As String
    Dim c As Long, t As Long, nm As String, ts As Single
    If visible = 0 Then ColorOf = "none": Exit Function
    c = cf.RGB
    ColorOf = Right$("0" & Hex$(c And 255), 2) & Right$("0" & Hex$((c \ 256) And 255), 2) & Right$("0" & Hex$((c \ 65536) And 255), 2)
    On Error Resume Next
    t = cf.ObjectThemeColor
    nm = ThemeName(t)
    If nm = "" Then Exit Function
    ts = cf.TintAndShade
    If ts > 0.004 Then nm = nm & "-lighter" & CStr(CLng(ts * 100))
    If ts < -0.004 Then nm = nm & "-darker" & CStr(CLng(-ts * 100))
    ColorOf = ColorOf & "@" & nm
End Function

' Theme colour name of an msoThemeColorIndex (the inverse of ThemeIndex), "" for none.
Private Function ThemeName(ByVal t As Long) As String
    Select Case t
        Case 1: ThemeName = "dk1"
        Case 2: ThemeName = "lt1"
        Case 3: ThemeName = "dk2"
        Case 4: ThemeName = "lt2"
        Case 5 To 10: ThemeName = "accent" & (t - 4)
        Case 11: ThemeName = "hlink"
        Case 12: ThemeName = "folHlink"
        Case 13: ThemeName = "tx1"
        Case 14: ThemeName = "bg1"
        Case 15: ThemeName = "tx2"
        Case 16: ThemeName = "bg2"
    End Select
End Function

' A JSON string literal.
Private Function JsonStr(ByVal v As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(v)
        ch = Mid$(v, i, 1)
        code = AscW(ch)
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 10: out = out & "\n"
            Case 9: out = out & "\t"
            Case 0 To 31: out = out & "\u" & Right$("000" & Hex$(code), 4)
            Case Else: out = out & ch
        End Select
    Next i
    JsonStr = """" & out & """"
End Function

' A tag written in chunks by the engine (render/tags.ts chunkedTags): NAME, or NAME_0..n + NAME_N.
Private Function ChunkedTag(ByVal shp As Object, ByVal key As String) As String
    Dim s As String, n As Long, i As Long
    s = TagValue(shp, key)
    If s = "" Then
        n = CLng(Val(TagValue(shp, key & "_N")))
        For i = 0 To n - 1
            s = s & TagValue(shp, key & "_" & i)
        Next i
    End If
    ChunkedTag = s
End Function

' Draws an engine response in place of chart `old` (same stacking order) and deletes it. Returns the
' new group, `old` when the response is "nothing to do", or Nothing.
Private Function ReplaceChart(ByVal old As Object, ByVal response As String) As Object
    Dim grp As Object, tp As Single
    Set grp = Draw(old.Parent, response)
    If gNothing Then Set ReplaceChart = old: Exit Function
    If grp Is Nothing Then Exit Function
    tp = Timer
    Do While grp.ZOrderPosition > old.ZOrderPosition + 1
        grp.ZOrder 3 ' msoSendBackward
    Loop
    Phase "zorder", tp: tp = Timer
    old.Delete
    Phase "delete", tp
    Set ReplaceChart = grp
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
    Dim groups As Collection, g As Variant
    Set groups = New Collection
    For Each shp In charts
        f = Split(TagValue(shp, "OC_FRAME"), ",")
        If UBound(f) = 3 Then
            If Abs(shp.Width - Val(f(2))) > 0.5 Or Abs(shp.Height - Val(f(3))) > 0.5 Then
                t = Timer
                ResetTiming
                selectedId = -1
                Set sel = FindSelectedChart()
                If Not sel Is Nothing Then selectedId = sel.Id
                wasSelected = (selectedId = shp.Id)
                Set result = Relayout(shp, False)
                If Not result Is Nothing Then
                    If TagValue(result, "OC_SGROUP") <> "" Then groups.Add TagValue(result, "OC_SGROUP") & vbTab & result.Id
                End If
                If wasSelected And Not result Is Nothing Then result.Select
                Report "RELAYOUT " & IIf(result Is Nothing, "failed: " & gLastMessage, "OK") & " (" & Format$((Timer - t) * 1000, "0") & " ms)" & Timing()
            ElseIf Abs(shp.Left - Val(f(0))) > 0.5 Or Abs(shp.Top - Val(f(1))) > 0.5 Then
                shp.Tags.Add "OC_FRAME", Fmt(shp.Left) & "," & Fmt(shp.Top) & "," & f(2) & "," & f(3)
            End If
        End If
    Next shp
    ' A resized chart changes its own scale: its same-scale partners follow (after the loop, which
    ' still holds the shapes the re-layout replaces).
    For Each g In groups
        If gAbortAuto Then Exit For
        SyncGroup sld, Split(g, vbTab)(0), CLng(Split(g, vbTab)(1))
    Next g
End Sub

' ---- Bridge to the engine -----------------------------------------------------------------

' Returns the engine's response, or "" after telling the user what went wrong.
Private Function Engine(ByVal request As String) As String
    Dim result As String, t As Single
    On Error GoTo NoBridge
    t = Timer
#If Mac Then
    result = AppleScriptTask(SCRIPT_FILE, SCRIPT_HANDLER, request)
    ' AppleScriptTask now and then hands back "" although the engine answered (never a valid answer;
    ' an engine request only lays out, so asking again is safe).
    If result = "" Then
        Report "ENGINE empty answer from AppleScriptTask, asked again"
        result = AppleScriptTask(SCRIPT_FILE, SCRIPT_HANDLER, request)
    End If
#Else
    result = WinEngine(request)
#End If
    gEngineMs = gEngineMs + (Timer - t) * 1000
    On Error GoTo 0
    If Left$(result, 3) <> "OC" & vbTab Then
        Tell "The chart engine gave an unexpected answer:" & vbLf & Left$(result, 300)
        Exit Function
    End If
    Engine = result
    Exit Function
NoBridge:
#If Mac Then
    Tell "open-charts can't reach its chart engine (" & Err.Description & ")." & vbLf & vbLf & _
        "Open Terminal and run:" & vbLf & REINSTALL & vbLf & "then restart PowerPoint."
#Else
    Tell "open-charts can't reach its chart engine (" & Err.Description & ")." & vbLf & vbLf & WIN_REINSTALL
#End If
End Function

Private Function Timing() As String
    Dim i As Long
    If gEngineMs > 0 Or gDrawMs > 0 Then Timing = " engine " & Format$(gEngineMs, "0") & " ms, draw " & Format$(gDrawMs, "0") & " ms"
    If gPhaseNames Is Nothing Then Exit Function
    If gPhaseNames.Count > 0 Then Timing = Timing & " ["
    For i = 1 To gPhaseNames.Count
        Timing = Timing & IIf(i > 1, ", ", "") & gPhaseNames(i) & " " & Format$(gPhaseMs(i), "0")
    Next i
    If gPhaseNames.Count > 0 Then Timing = Timing & " ms]"
End Function

Private Sub ResetTiming()
    gEngineMs = 0: gDrawMs = 0
    Set gPhaseNames = New Collection
    Set gPhaseMs = New Collection
End Sub

' Adds the time since t0 (a Timer value) to phase `name` of the timing log.
Private Sub Phase(ByVal name As String, ByVal t0 As Single)
    Dim i As Long, ms As Single
    ms = (Timer - t0) * 1000
    If gPhaseNames Is Nothing Then ResetTiming
    For i = 1 To gPhaseNames.Count
        If gPhaseNames(i) = name Then
            ms = ms + gPhaseMs(i)
            gPhaseMs.Remove i
            If i <= gPhaseMs.Count Then gPhaseMs.Add ms, Before:=i Else gPhaseMs.Add ms
            Exit Sub
        End If
    Next i
    gPhaseNames.Add name: gPhaseMs.Add ms
End Sub

Private Sub Report(ByVal text As String)
    If gCollect Then gCollected = gCollected & text & vbLf: Exit Sub
    On Error Resume Next
#If Mac Then
    Dim ignored As String
    ignored = AppleScriptTask(SCRIPT_FILE, "logLine", text)
#Else
    WinLog text
#End If
End Sub

' A diagnostic line: logged only with verbose logging on (results and errors use Report).
Private Sub Verbose(ByVal text As String)
    If gVerbose Then Report text
End Sub

Private Sub Tell(ByVal text As String)
    gLastMessage = text
    If gSilent Then Exit Sub
    MsgBox text, vbOKOnly, TITLE
End Sub

#If Mac Then
#Else
' ---- Windows bridge: HTTP to the engine's local server --------------------------------------

' Returns the engine's response. Starts the server (open-charts.exe) if it is not running and waits
' for it; raises an error with the reason when the engine cannot be reached.
Private Function WinEngine(ByVal request As String) As String
    Dim ok As Boolean, result As String
    result = WinHttp("run", request, 30000, ok)
    If ok Then WinEngine = result: Exit Function
    If Not WinStartServer(20000) Then
        If Dir(Environ("LOCALAPPDATA") & ENGINE_EXE) = "" Then Err.Raise vbObjectError + 1, , Environ("LOCALAPPDATA") & ENGINE_EXE & " is missing"
        Err.Raise vbObjectError + 2, , "the engine did not start within 20 seconds"
    End If
    result = WinHttp("run", request, 30000, ok)
    If Not ok Then Err.Raise vbObjectError + 3, , result
    WinEngine = result
End Function

' Starts the engine's server unless it answers; waits up to waitMs for it (0: does not wait).
Private Function WinStartServer(ByVal waitMs As Long) As Boolean
    Dim ok As Boolean, exe As String, waited As Long
    ' A running server accepts at once; a closed port takes Windows about a second to refuse, so the
    ' first check gives up after 300 ms (it runs as the ribbon loads).
    WinHttp "ping", "", 300, ok
    If ok Then WinStartServer = True: Exit Function
    exe = Environ("LOCALAPPDATA") & ENGINE_EXE
    If Dir(exe) = "" Then Exit Function
    On Error Resume Next
    Shell """" & exe & """ addin serve", vbHide
    On Error GoTo 0
    Do While waited < waitMs
        Sleep 100
        DoEvents
        waited = waited + 100
        WinHttp "ping", "", 500, ok
        If ok Then WinStartServer = True: Exit Function
    Loop
End Function

' One call to the server: GET for "ping", else POST of body (sent as UTF-8) with the install's token
' (the server refuses calls without it). ok is True on 200; the result is then the response,
' otherwise what went wrong. Never through a proxy.
Private Function WinHttp(ByVal path As String, ByVal body As String, ByVal timeoutMs As Long, ByRef ok As Boolean) As String
    Dim http As Object
    ok = False
    On Error GoTo Failed
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error Resume Next
    http.setProxy 1 ' SXH_PROXY_SET_DIRECT
    On Error GoTo Failed
    http.setTimeouts 1000, IIf(timeoutMs < 1000, timeoutMs, 1000), timeoutMs, timeoutMs
    If path = "ping" Then
        http.Open "GET", ENGINE_URL & path, False
        http.send
    Else
        http.Open "POST", ENGINE_URL & path, False
        http.setRequestHeader "Content-Type", "text/plain; charset=utf-8"
        http.setRequestHeader "X-OC-Token", WinToken()
        http.send body
    End If
    If http.Status = 200 Then
        ok = True
        WinHttp = http.responseText
    Else
        WinHttp = "the engine answered " & http.Status & ": " & Left$(http.responseText, 200)
    End If
    Exit Function
Failed:
    WinHttp = Err.Description
End Function

' The install's token: the file holds "X-OC-Token: <hex>" ("" when it is missing).
Private Function WinToken() As String
    Dim t As String
    t = Replace(Replace(WinReadText(Environ("LOCALAPPDATA") & TOKEN_FILE), vbCr, ""), vbLf, "")
    If Left$(t, 12) = "X-OC-Token: " Then WinToken = Trim$(Mid$(t, 13))
End Function

' A small text file's contents ("" when it cannot be read).
Private Function WinReadText(ByVal path As String) As String
    Dim f As Integer
    On Error GoTo Failed
    If Dir(path) = "" Then Exit Function
    f = FreeFile
    Open path For Binary Access Read As #f
    WinReadText = Space$(LOF(f))
    Get #f, , WinReadText
    Close #f
    Exit Function
Failed:
    WinReadText = ""
    On Error Resume Next
    If f <> 0 Then Close #f
End Function

' Appends one line to %LOCALAPPDATA%\open-charts\addin.log, as the Mac bridge's logLine does.
Private Sub WinLog(ByVal text As String)
    Dim f As Integer, folder As String
    On Error Resume Next
    folder = Environ("LOCALAPPDATA") & "\open-charts"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    f = FreeFile
    Open folder & "\addin.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss") & " " & Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " ")
    Close #f
End Sub
#End If

' ---- Drawing --------------------------------------------------------------------------------

' Draws the response on the slide, groups it and writes the tags. Returns the group, or Nothing.
Private Function Draw(ByVal sld As Object, ByVal response As String) As Object
    Dim lines() As String, f() As String, i As Long, n As Long
    Dim names() As Variant, finalNames() As String, roles() As String
    Dim tagKeys() As String, tagVals() As String, nTags As Long
    Dim groupName As String, alt As String, stamp As String, shp As Object
    Dim t As Single, tp As Single
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
    Phase "shapes", t: tp = Timer

    Dim grp As Object, k As Long, idx As Long
    If n = 1 Then
        Set grp = sld.Shapes(names(0))
    Else
        ReDim Preserve names(0 To n - 1)
        Set grp = sld.Shapes.Range(names).Group
    End If
    Phase "group", tp: tp = Timer
    grp.Name = groupName
    On Error Resume Next
    grp.AlternativeText = alt
    On Error GoTo 0
    For k = 0 To nTags - 1
        grp.Tags.Add tagKeys(k), tagVals(k)
    Next k
    Phase "tags", tp: tp = Timer
    If n > 1 Then
        For k = 1 To grp.GroupItems.Count
            Set shp = grp.GroupItems(k)
            idx = CLng(Mid$(shp.Name, Len(stamp) + 1))
            If roles(idx) <> "" Then shp.Tags.Add "OC_ROLE", roles(idx)
            shp.Name = finalNames(idx)
        Next k
    End If
    Phase "names", tp
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
'  23 textGeom 24 rotate 25 arrowStart 26 pattern (per-element overrides; absent from older engines)
'  27 level 28 bullet 29 bulletAt 30 bulletHang (agenda headings: list paragraph)
Private Function DrawShape(ByVal sld As Object, f() As String) As Object
    Dim kind As String, x As Single, y As Single, w As Single, h As Single, s As Object
    kind = f(1)
    x = Val(f(2)): y = Val(f(3)): w = Val(f(4)): h = Val(f(5))
    Select Case kind
        Case "line"
            Set s = sld.Shapes.AddLine(x, y, w, h)
            With s.Line
                .Visible = -1
                SetColor .ForeColor, f(7)
                .Weight = Val(f(8))
                If f(9) = "1" Then .DashStyle = 4 ' msoLineDash
                If f(9) = "dot" Then .DashStyle = 2 ' msoLineSquareDot
                If f(10) = "1" Then .EndArrowheadStyle = 2 ' msoArrowheadTriangle
                If Col(f, 25) = "1" Then .BeginArrowheadStyle = 2
            End With
        Case "text"
            If f(6) = "" And f(7) = "" Then
                Set s = sld.Shapes.AddTextbox(Orientation(Col(f, 24)), x, y, w, h)
            Else
                Set s = sld.Shapes.AddShape(IIf(f(23) = "roundRect", 5, 1), x, y, w, h)
                If f(23) = "roundRect" Then s.Adjustments(1) = 0.5
            End If
            FillAndLine s, f(6), f(7), f(8), "0"
            If Col(f, 24) <> "" Then s.TextFrame2.Orientation = Orientation(Col(f, 24))
            SetText s, f
            s.Left = x: s.Top = y: s.Width = w: s.Height = h
        Case "path"
            Set s = DrawPath(sld, f(22))
            FillAndLine s, f(6), f(7), f(8), f(9)
            If Col(f, 26) <> "" Then SetPattern s, f(6), Col(f, 26)
        Case Else
            Set s = sld.Shapes.AddShape(AutoShapeType(kind), x, y, w, h)
            FillAndLine s, f(6), f(7), f(8), f(9)
            If Col(f, 26) <> "" Then SetPattern s, f(6), Col(f, 26)
    End Select
    Set DrawShape = s
End Function

' A pattern fill (an override): OOXML preset `prst` in colour `fillHex` on the theme background.
Private Sub SetPattern(ByVal s As Object, ByVal fillHex As String, ByVal prst As String)
    Dim n As Long
    n = PatternType(prst)
    If n = 0 Or fillHex = "" Then Exit Sub
    On Error Resume Next
    With s.Fill
        .Patterned n
        SetColor .ForeColor, fillHex
        .BackColor.ObjectThemeColor = 14 ' msoThemeColorBackground1
    End With
End Sub

' msoPatternType of an OOXML preset pattern (DrawingML ST_PresetPatternVal), 0 if unknown.
Private Function PatternType(ByVal prst As String) As Long
    Select Case prst
        Case "pct5": PatternType = 1
        Case "pct10": PatternType = 2
        Case "pct20": PatternType = 3
        Case "pct25": PatternType = 4
        Case "pct30": PatternType = 5
        Case "pct40": PatternType = 6
        Case "pct50": PatternType = 7
        Case "pct60": PatternType = 8
        Case "pct70": PatternType = 9
        Case "pct75": PatternType = 10
        Case "pct80": PatternType = 11
        Case "pct90": PatternType = 12
        Case "dkHorz": PatternType = 13
        Case "dkVert": PatternType = 14
        Case "dkDnDiag": PatternType = 15
        Case "dkUpDiag": PatternType = 16
        Case "smCheck": PatternType = 17
        Case "trellis": PatternType = 18
        Case "ltHorz": PatternType = 19
        Case "ltVert": PatternType = 20
        Case "ltDnDiag": PatternType = 21
        Case "ltUpDiag": PatternType = 22
        Case "smGrid": PatternType = 23
        Case "dotDmnd": PatternType = 24
        Case "wdDnDiag": PatternType = 25
        Case "wdUpDiag": PatternType = 26
        Case "dashUpDiag": PatternType = 27
        Case "dashDnDiag": PatternType = 28
        Case "narVert": PatternType = 29
        Case "narHorz": PatternType = 30
        Case "dashVert": PatternType = 31
        Case "dashHorz": PatternType = 32
        Case "lgConfetti": PatternType = 33
        Case "lgGrid": PatternType = 34
        Case "horzBrick": PatternType = 35
        Case "lgCheck": PatternType = 36
        Case "smConfetti": PatternType = 37
        Case "zigZag": PatternType = 38
        Case "solidDmnd": PatternType = 39
        Case "diagBrick": PatternType = 40
        Case "openDmnd": PatternType = 41
        Case "plaid": PatternType = 42
        Case "sphere": PatternType = 43
        Case "weave": PatternType = 44
        Case "dotGrid": PatternType = 45
        Case "divot": PatternType = 46
        Case "shingle": PatternType = 47
        Case "wave": PatternType = 48
        Case "horz": PatternType = 49
        Case "vert": PatternType = 50
        Case "cross": PatternType = 51
        Case "dnDiag": PatternType = 52
        Case "upDiag": PatternType = 53
        Case "diagCross": PatternType = 54
    End Select
End Function

' Column i of a record, "" when an older engine did not send it.
Private Function Col(f() As String, ByVal i As Long) As String
    If i <= UBound(f) Then Col = f(i)
End Function

' Text direction (rotate column): -90 reads bottom to top (msoTextOrientationUpward), 90 top to
' bottom (msoTextOrientationDownward), else horizontal.
Private Function Orientation(ByVal rotate As String) As Long
    Select Case rotate
        Case "-90": Orientation = 2
        Case "90": Orientation = 3
        Case Else: Orientation = 1
    End Select
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
    ' One FillFormat / LineFormat object each: every s.Fill or s.Line is another call into PowerPoint.
    With s.Fill
        If fillHex = "" Then
            .Visible = 0
        Else
            .Visible = -1
            .Solid
            SetColor .ForeColor, fillHex
            .Transparency = 0
        End If
    End With
    With s.Line
        If lineHex = "" Then
            .Visible = 0
        Else
            .Visible = -1
            SetColor .ForeColor, lineHex
            .Weight = Val(lineWidth)
            If dash = "1" Then .DashStyle = 4
            If dash = "dot" Then .DashStyle = 2 ' msoLineSquareDot
        End If
    End With
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
    Dim tr As Object
    Set tr = tf.TextRange
    With tr.Font
        .Name = f(12)
        .NameComplexScript = f(12) ' as the .pptx writer does (a:cs), so digits render the same
        .Size = Val(f(13))
        .Bold = IIf(f(14) = "1", -1, 0)
        ' No kerning, as the engine measured. Decks often kern text of 12 pt and up by default;
        ' Kerning is the size from which pairs are kerned (0 = all sizes), so push it past any text.
        .Kerning = 4000
        With .Fill
            .Visible = -1
            SetColor .ForeColor, f(15)
        End With
    End With
    With tr.ParagraphFormat
        Select Case f(16)
            Case "c": .Alignment = 2 ' msoAlignCenter
            Case "r": .Alignment = 3 ' msoAlignRight
            Case Else: .Alignment = 1 ' msoAlignLeft
        End Select
        .SpaceBefore = 0
        .SpaceAfter = 0
        .SpaceWithin = 1
    End With
    If Col(f, 28) <> "" Then SetBullet s, tr.ParagraphFormat, Col(f, 27), Col(f, 28), Val(Col(f, 29)), Val(Col(f, 30))
End Sub

' A list paragraph (agenda heading): text level, bullet "u<code point>" or number "n<start>", the
' bullet `at` pt from the left edge and the text `hang` pt after it. The indents are set after the
' bullet (turning a bullet on brings back the level's default indents), on the paragraph and on the
' text frame's ruler level, so PowerPoint draws the text where the engine measured it.
Private Sub SetBullet(ByVal s As Object, ByVal pf As Object, ByVal level As String, ByVal bullet As String, ByVal at As Single, ByVal hang As Single)
    Dim lvl As Long
    On Error Resume Next
    lvl = CLng(Val(level))
    If lvl < 1 Then lvl = 1
    If lvl > 1 Then pf.IndentLevel = lvl
    With pf.Bullet
        .Visible = -1
        If Left$(bullet, 1) = "n" Then
            .Type = 2 ' msoBulletNumbered
            .Style = 3 ' msoBulletArabicPeriod
            .StartValue = CLng(Val(Mid$(bullet, 2)))
        Else
            .Type = 1 ' msoBulletUnnumbered
            .Character = CLng(Val(Mid$(bullet, 2)))
        End If
        .UseTextColor = -1
        .UseTextFont = -1
    End With
    With s.TextFrame.Ruler.Levels(lvl)
        .FirstMargin = at
        .LeftMargin = at + hang
    End With
    pf.LeftIndent = at + hang
    pf.FirstLineIndent = -hang
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

' True when the selection is a piece of an open-charts chart that was ungrouped: its shapes keep
' their OC_ROLE tags, but no group carries the chart's data (OC_SPEC) any more, and regrouping
' does not bring it back.
Private Function UngroupedSelected() As Boolean
    Dim sel As Object, shp As Object
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    If sel.Type <> 2 And sel.Type <> 3 Then Exit Function ' ppSelectionShapes, ppSelectionText
    Set shp = sel.ShapeRange(1)
    If sel.HasChildShapeRange Then Set shp = sel.ChildShapeRange(1)
    If shp Is Nothing Then Exit Function
    If TagValue(shp, "OC_ROLE") <> "" Then UngroupedSelected = True: Exit Function
    If shp.Type = 6 Then ' msoGroup: the pieces regrouped
        If TagValue(shp.GroupItems(1), "OC_ROLE") <> "" Then UngroupedSelected = True
    End If
End Function

' The selected chart, or Nothing after telling the user what to select.
Private Function SelectedChart() As Object
    Set SelectedChart = FindSelectedChart()
    If Not SelectedChart Is Nothing Then Exit Function
    If UngroupedSelected() Then
#If Mac Then
        Tell "This chart was ungrouped, so it can't be edited as a chart. Undo (" & ChrW(&H2318) & "Z) the ungroup or insert it again."
#Else
        Tell "This chart was ungrouped, so it can't be edited as a chart. Undo (Ctrl+Z) the ungroup or insert it again."
#End If
    Else
        Tell "Select an open-charts chart first (click it once), then choose Edit data or Update chart."
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

' A colour field of the line format: RRGGBB, or a theme colour name. A theme colour is set as such
' (ObjectThemeColor), so the chart follows the presentation's theme like PowerPoint's own shapes.
Private Sub SetColor(ByVal cf As Object, ByVal v As String)
    Dim t As Long, at As Long, shade As Single
    ' A theme colour with a tint / shade (per-element overrides): "accent1-lighter40" / "accent1-darker25".
    at = InStr(v, "-lighter")
    If at > 0 Then
        shade = Val(Mid$(v, at + 8)) / 100
    Else
        at = InStr(v, "-darker")
        If at > 0 Then shade = -Val(Mid$(v, at + 7)) / 100
    End If
    If at > 0 Then v = Left$(v, at - 1)
    t = ThemeIndex(v)
    If t > 0 Then
        cf.ObjectThemeColor = t
        If at > 0 Then cf.TintAndShade = shade
    Else
        cf.RGB = HexColor(v)
    End If
End Sub

' msoThemeColorIndex of a theme colour name (src/core/theme.ts), 0 for anything else.
Private Function ThemeIndex(ByVal v As String) As Long
    Select Case v
        Case "dk1": ThemeIndex = 1 ' msoThemeColorDark1
        Case "lt1": ThemeIndex = 2 ' msoThemeColorLight1
        Case "dk2": ThemeIndex = 3 ' msoThemeColorDark2
        Case "lt2": ThemeIndex = 4 ' msoThemeColorLight2
        Case "accent1": ThemeIndex = 5 ' msoThemeColorAccent1 ... Accent6 = 10
        Case "accent2": ThemeIndex = 6
        Case "accent3": ThemeIndex = 7
        Case "accent4": ThemeIndex = 8
        Case "accent5": ThemeIndex = 9
        Case "accent6": ThemeIndex = 10
        Case "hlink": ThemeIndex = 11 ' msoThemeColorHyperlink
        Case "folHlink": ThemeIndex = 12 ' msoThemeColorFollowedHyperlink
        Case "tx1": ThemeIndex = 13 ' msoThemeColorText1: follows the slide master's colour map
        Case "bg1": ThemeIndex = 14 ' msoThemeColorBackground1
        Case "tx2": ThemeIndex = 15 ' msoThemeColorText2
        Case "bg2": ThemeIndex = 16 ' msoThemeColorBackground2
    End Select
End Function

' Request lines with the slide's theme: its colours and its fonts.
Private Function ThemeLine(ByVal sld As Object) As String
    ThemeLine = ThemeColorsLine(sld) & ThemeFontsLine(sld)
End Function

' Request line with the slide's theme colours (ThemeColorScheme order), so the engine decides
' label contrast on the colours PowerPoint will show. "" if PowerPoint cannot say.
Private Function ThemeColorsLine(ByVal sld As Object) As String
    Dim scheme As Object, i As Long, c As Long, out As String
    On Error GoTo Done
    Set scheme = sld.ThemeColorScheme
    For i = 1 To 12
        c = scheme.Colors(i).RGB
        out = out & IIf(i > 1, ",", "") & Right$("0" & Hex$(c And 255), 2) & Right$("0" & Hex$((c \ 256) And 255), 2) & Right$("0" & Hex$((c \ 65536) And 255), 2)
    Next i
    ThemeColorsLine = "theme=" & out & vbLf
Done:
End Function

' Request lines with the deck's theme fonts (headings, body: what +mj-lt / +mn-lt stand for), so the
' engine measures text in the fonts PowerPoint will draw. "" if PowerPoint cannot say.
Private Function ThemeFontsLine(ByVal sld As Object) As String
    Dim fs As Object, major As String, minor As String
    On Error Resume Next
    Set fs = sld.Design.SlideMaster.Theme.ThemeFontScheme
    If fs Is Nothing Then Set fs = ActivePresentation.SlideMaster.Theme.ThemeFontScheme
    On Error GoTo Done
    If fs Is Nothing Then Exit Function
    major = fs.MajorFont.Item(1).Name ' msoThemeLatin
    minor = fs.MinorFont.Item(1).Name
    If major <> "" Then ThemeFontsLine = "majorfont=" & major & vbLf
    If minor <> "" Then ThemeFontsLine = ThemeFontsLine & "minorfont=" & minor & vbLf
Done:
End Function

' ---- Same-scale groups ------------------------------------------------------------------------
' Charts on one slide whose spec has the same options.scaleGroup carry OC_SGROUP (the group) and
' OC_SCALE (their own scale, pt per unit). Each one is laid out with the smallest OC_SCALE of its
' partners, so equal values are equally tall across the group.

' Request line "scale=" with the smallest own scale of the chart's partners, "" outside a group.
Private Function ScaleLine(ByVal chart As Object) As String
    ScaleLine = ScaleLineFor(chart, TagValue(chart, "OC_SGROUP"))
End Function

' The same for the chart as a member of group g.
Private Function ScaleLineFor(ByVal chart As Object, ByVal g As String) As String
    Dim shp As Object, best As String, v As Double
    If g = "" Then Exit Function
    For Each shp In chart.Parent.Shapes
        If IsChart(shp) Then
            If shp.Id <> chart.Id And TagValue(shp, "OC_SGROUP") = g Then
                v = Val(TagValue(shp, "OC_SCALE"))
                If v > 0 And (best = "" Or v < Val(best)) Then best = TagValue(shp, "OC_SCALE")
            End If
        End If
    Next shp
    If best <> "" Then ScaleLineFor = "scale=" & best & vbLf
End Function

' Re-lays out every chart of group g on the slide except the one with id skipId.
Private Sub SyncGroup(ByVal sld As Object, ByVal g As String, ByVal skipId As Long)
    Dim shp As Object, peers As Collection, n As Long
    If g = "" Then Exit Sub
    Set peers = New Collection
    For Each shp In sld.Shapes
        If IsChart(shp) Then
            If shp.Id <> skipId And TagValue(shp, "OC_SGROUP") = g Then peers.Add shp
        End If
    Next shp
    For Each shp In peers
        If Not Relayout(shp, True) Is Nothing Then n = n + 1
    Next shp
    If peers.Count > 0 Then Report "SYNC " & g & " " & n & " of " & peers.Count & " charts"
End Sub

' After a chart was laid out again: its group follows, and so does the group it just left.
Private Sub SyncAfter(ByVal sld As Object, ByVal chart As Object, ByVal oldGroup As String)
    Dim g As String
    g = TagValue(chart, "OC_SGROUP")
    SyncGroup sld, g, chart.Id
    If oldGroup <> "" And oldGroup <> g Then SyncGroup sld, oldGroup, 0
End Sub

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

' ---- Agenda ---------------------------------------------------------------------------------
' Agenda slides start chapters (docs/USER-GUIDE.md "Agenda"). This sends the engine a snapshot of
' the deck (src/addin/agenda.ts: D, L and H lines) and applies its answer: per agenda slide a new
' agenda element (or only its tags), its chapter id and hidden state (P records), and the slides
' to insert, move or delete (X records). Slides are found by their tags: OC_KIND, OC_CHAPTER,
' OC_SPLIT; the element by OC_AGENDA, each heading by OC_ROLE "chapter:<id>".


Private Sub AgendaCommand(ByVal cmd As String)
    Dim sld As Object, engineCmd As String, response As String, newId As String, moved As Boolean
    Select Case cmd
        Case "Placeholder"
            AgendaPlaceholder
            Exit Sub
        Case "Split", "Join"
            Set sld = CurrentSlide()
            If sld Is Nothing Then Exit Sub
            If cmd = "Join" And TagValue(sld, "OC_SPLIT") <> "1" Then
                Tell "This slide does not start an agenda section. Go to the slide with the agenda split (the marker above the slide) and click Join Agendas there."
                Exit Sub
            End If
            SetSplit sld, cmd = "Split"
            engineCmd = "update"
        Case "Chapter": engineCmd = "chapter"
        Case "Toc": engineCmd = "toc"
        Case "Up": engineCmd = "up"
        Case "Down": engineCmd = "down"
        Case "Indent": engineCmd = "indent"
        Case "Outdent": engineCmd = "outdent"
        Case "RemoveHidden": engineCmd = "removeHidden"
        Case "InsertHidden": engineCmd = "insertHidden"
        Case Else: engineCmd = "update"
    End Select
    response = Engine(AgendaRequest(engineCmd))
    If gSilent Then Report "AGENDA " & engineCmd & ": engine answered " & Len(response) & " chars"
    If response = "" Then Exit Sub
    moved = ApplyAgenda(response, newId)
    If gSilent Then Report "AGENDA " & engineCmd & ": applied" & IIf(moved, ", slides inserted or moved", "")
    ' After slides were inserted or moved, every agenda slide is laid out again.
    If moved Then
        response = Engine(AgendaRequest("update"))
        If response <> "" Then ApplyAgenda response, ""
    End If
    If newId <> "" Then SelectHeading newId
    gAgendaSig = AgendaSignature()
End Sub

' The request: command, slide size, the layout's agenda placeholder, theme, then the snapshot.
Private Function AgendaRequest(ByVal cmd As String) As String
    Dim pres As Object, themeText As String
    Set pres = ActivePresentation
    If pres.Slides.Count > 0 Then themeText = ThemeLine(pres.Slides(1))
    AgendaRequest = "agenda" & vbLf & cmd & vbLf & "0,0," & Fmt(pres.PageSetup.SlideWidth) & "," & Fmt(pres.PageSetup.SlideHeight) & vbLf & _
        PlaceholderFrame() & vbLf & themeText & AgendaSnapshot()
End Function

Private Function AgendaSnapshot() As String
    Dim pres As Object, sld As Object, grp As Object, i As Long, cur As Long, parts() As String
    Dim bounds As String, json As String, hash As String, drawnAt As String
    Set pres = ActivePresentation
    cur = CurrentSlideIndex()
    ReDim parts(0 To pres.Slides.Count)
    parts(0) = "D" & vbTab & cur & vbTab & Esc(SelectedHeadingChapter())
    For i = 1 To pres.Slides.Count
        Set sld = pres.Slides(i)
        Set grp = AgendaGroup(sld)
        bounds = "": json = "": hash = "": drawnAt = ""
        If Not grp Is Nothing Then
            bounds = Fmt(grp.Left) & "," & Fmt(grp.Top) & "," & Fmt(grp.Width) & "," & Fmt(grp.Height)
            json = TagValue(grp, "OC_AGENDA"): hash = TagValue(grp, "OC_HASH"): drawnAt = TagValue(grp, "OC_FRAME")
        End If
        parts(i) = "L" & vbTab & i & vbTab & Esc(TagValue(sld, "OC_KIND")) & vbTab & Esc(TagValue(sld, "OC_CHAPTER")) & vbTab & _
            IIf(TagValue(sld, "OC_SPLIT") = "1", "1", "0") & vbTab & IIf(sld.SlideShowTransition.Hidden, "1", "0") & vbTab & _
            bounds & vbTab & Esc(json) & vbTab & Esc(hash) & vbTab & drawnAt
        If Not grp Is Nothing Then parts(i) = parts(i) & HeadingLines(grp, i, i = cur)
    Next i
    AgendaSnapshot = Join(parts, vbLf)
End Function

' H lines for the headings of an agenda element as they are now (text, text level; format too on
' the current slide, where a restyle is looked for).
Private Function HeadingLines(ByVal grp As Object, ByVal n As Long, ByVal withStyle As Boolean) As String
    Dim k As Long, shp As Object, role As String, tr As Object, ch As Object, out As String, st As String
    On Error Resume Next
    For k = 1 To grp.GroupItems.Count
        Set shp = grp.GroupItems(k)
        role = TagValue(shp, "OC_ROLE")
        If Left$(role, 8) = "chapter:" Then
            Set tr = shp.TextFrame2.TextRange
            out = out & vbLf & "H" & vbTab & n & vbTab & Esc(Mid$(role, 9)) & vbTab & Esc(HeadingText(tr)) & vbTab & tr.Paragraphs(1).ParagraphFormat.IndentLevel
            If withStyle Then
                Set ch = tr.Characters(1, 1)
                st = vbTab & IIf(shp.Fill.Visible, HexOf(shp.Fill.ForeColor.RGB), "") & vbTab & IIf(shp.Line.Visible, HexOf(shp.Line.ForeColor.RGB), "")
                st = st & vbTab & HexOf(ch.Font.Fill.ForeColor.RGB) & vbTab & IIf(ch.Font.Bold, "1", "0") & vbTab & Fmt(ch.Font.Size)
                out = out & st
            End If
        End If
    Next k
    HeadingLines = out
End Function

Private Function HeadingText(ByVal tr As Object) As String
    HeadingText = Trim$(Replace(Replace(Replace(tr.Text, vbCr, " "), vbLf, " "), Chr$(11), " "))
End Function

' Applies P records (in order), then X records: new and move slides (returns True: lay out again),
' then deletes. newId: the chapter of a slide inserted by Chapter, to select its heading.
Private Function ApplyAgenda(ByVal response As String, ByRef newId As String) As Boolean
    Dim lines() As String, f() As String, i As Long, records As String, cur() As String, haveP As Boolean
    Dim structural As Collection, x As Variant
    lines = Split(response, vbLf)
    Set structural = New Collection
    For i = 0 To UBound(lines)
        If lines(i) <> "" Then
            f = Split(lines(i), vbTab)
            Select Case f(0)
                Case "P"
                    If haveP Then ApplySlide cur, records
                    cur = f: records = "": haveP = True
                Case "G", "F", "T", "S"
                    records = records & lines(i) & vbLf
                Case "X"
                    structural.Add lines(i)
                Case "M"
                    Tell Unescape(f(1))
                Case "E"
                    Tell Unescape(f(1))
                    Exit Function
            End Select
        End If
    Next i
    If haveP Then ApplySlide cur, records
    For Each x In structural
        f = Split(x, vbTab)
        If f(1) = "new" Then
            NewAgendaSlideAt CLng(f(2)), f(3), Unescape(f(4)), CLng(f(5))
            If f(3) = "chapter" And newId = "" Then newId = Unescape(f(4))
            ApplyAgenda = True
        ElseIf f(1) = "move" Then
            ActivePresentation.Slides.Range(SlideNumbers(CLng(f(2)), CLng(f(3)))).MoveTo CLng(f(4))
            ApplyAgenda = True
        End If
    Next x
    For Each x In structural
        f = Split(x, vbTab)
        If f(1) = "delete" Then ActivePresentation.Slides(CLng(f(2))).Delete
    Next x
End Function

Private Function SlideNumbers(ByVal from As Long, ByVal count As Long) As Variant
    Dim a() As Variant, k As Long
    ReDim a(0 To count - 1)
    For k = 0 To count - 1
        a(k) = from + k
    Next k
    SlideNumbers = a
End Function

' One P record: slide tags, hidden state, and the new element (drawn where the old one was in the
' stacking order) or the old element's new tags.
Private Sub ApplySlide(f() As String, ByVal records As String)
    Dim sld As Object, old As Object, grp As Object, lines() As String, t() As String, i As Long
    Set sld = ActivePresentation.Slides(CLng(f(1)))
    If f(2) <> "" Then sld.Tags.Add "OC_CHAPTER", Unescape(f(2))
    If UBound(f) >= 5 Then
        If f(5) <> "" Then sld.Tags.Add "OC_KIND", f(5)
    End If
    sld.SlideShowTransition.Hidden = IIf(f(3) = "1", -1, 0)
    Set old = AgendaGroup(sld)
    If f(4) = "1" Then
        Set grp = Draw(sld, records)
        If grp Is Nothing Then Exit Sub
        If Not old Is Nothing Then
            Do While grp.ZOrderPosition > old.ZOrderPosition + 1
                grp.ZOrder 3 ' msoSendBackward
            Loop
            old.Delete
        End If
    ElseIf Not old Is Nothing Then
        lines = Split(records, vbLf)
        For i = 0 To UBound(lines)
            If Left$(lines(i), 2) = "T" & vbTab Then
                t = Split(lines(i), vbTab)
                old.Tags.Add Unescape(t(1)), Unescape(t(2))
            End If
        Next i
    End If
End Sub

' Inserts an agenda slide so it becomes slide `pos`: a copy of slide `roleModel` (the closest agenda
' slide: same master, layout and shapes), else a new slide on the "open-charts agenda" layout or a
' title-only layout.
Private Sub NewAgendaSlideAt(ByVal pos As Long, ByVal kind As String, ByVal id As String, ByVal roleModel As Long)
    Dim pres As Object, sld As Object
    Set pres = ActivePresentation
    If roleModel > 0 And roleModel <= pres.Slides.Count Then
        pres.Slides(roleModel).Duplicate.MoveTo pos
        Set sld = pres.Slides(pos)
        SetSplit sld, False
        sld.SlideShowTransition.Hidden = 0
    Else
        Set sld = NewAgendaSlide(pos)
    End If
    sld.Tags.Add "OC_KIND", kind
    If id <> "" Then
        sld.Tags.Add "OC_CHAPTER", id
    Else
        On Error Resume Next
        sld.Tags.Delete "OC_CHAPTER"
        On Error GoTo 0
    End If
    On Error Resume Next
    ActiveWindow.View.GotoSlide pos
End Sub

Private Function NewAgendaSlide(ByVal pos As Long) As Object
    Dim pres As Object, lay As Object, ref As Object, sld As Object, k As Long, shp As Object
    Set pres = ActivePresentation
    Set lay = AgendaLayout()
    If pres.Slides.Count > 0 Then Set ref = pres.Slides(IIf(pos > 1, pos - 1, 1))
    If lay Is Nothing Then Set lay = TitleOnlyLayout(ref)
    If lay Is Nothing And Not ref Is Nothing Then Set lay = ref.CustomLayout
    If lay Is Nothing Then Set lay = pres.SlideMaster.CustomLayouts(1)
    Set sld = pres.Slides.AddSlide(pos, lay)
    For k = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(k)
        If shp.Type = 14 Then ' msoPlaceholder
            If shp.PlaceholderFormat.Type = 1 Or shp.PlaceholderFormat.Type = 3 Then ' title, center title
                shp.TextFrame.TextRange.Text = AGENDA_TITLE
            ElseIf shp.HasTextFrame Then
                If shp.TextFrame.TextRange.Length = 0 Then shp.Delete
            End If
        End If
    Next k
    If sld.Shapes.HasTitle = 0 Then ' msoFalse (Not on a Mac True of 1 would still be True)
        ' A layout without a title: a title text box like the CLI writes.
        Set shp = sld.Shapes.AddTextbox(1, 36, 24, pres.PageSetup.SlideWidth - 72, 44)
        shp.Name = "Title"
        With shp.TextFrame2.TextRange
            .Text = AGENDA_TITLE
            .Font.Size = 24
            .Font.Bold = -1
        End With
    End If
    Set NewAgendaSlide = sld
End Function

' The layout named "open-charts agenda" in any slide master of the presentation, else Nothing.
Private Function AgendaLayout() As Object
    Dim d As Object, cl As Object
    On Error Resume Next
    For Each d In ActivePresentation.Designs
        For Each cl In d.SlideMaster.CustomLayouts
            If cl.Name = AGENDA_LAYOUT Then Set AgendaLayout = cl: Exit Function
        Next cl
    Next d
End Function

Private Function TitleOnlyLayout(ByVal ref As Object) As Object
    Dim cl As Object, master As Object
    On Error Resume Next
    If ref Is Nothing Then Set master = ActivePresentation.SlideMaster Else Set master = ref.Design.SlideMaster
    For Each cl In master.CustomLayouts
        If cl.Name = "Title Only" Then Set TitleOnlyLayout = cl: Exit Function
    Next cl
End Function

' "x,y,w,h" of the agenda placeholder on the "open-charts agenda" layout, "" if none.
Private Function PlaceholderFrame() As String
    Dim lay As Object, shp As Object
    Set lay = AgendaLayout()
    If lay Is Nothing Then Exit Function
    On Error Resume Next
    For Each shp In lay.Shapes
        If TagValue(shp, "OC_AGENDA_PH") = "1" Then
            PlaceholderFrame = Fmt(shp.Left) & "," & Fmt(shp.Top) & "," & Fmt(shp.Width) & "," & Fmt(shp.Height)
            Exit Function
        End If
    Next shp
End Function

' Agenda Placeholder (Slide Master view): an invisible rectangle on the layout that marks where
' agenda elements go on agenda slides made from it.
Private Sub AgendaPlaceholder()
    Dim lay As Object, shp As Object, w As Single, h As Single
    On Error Resume Next
    Set lay = ActiveWindow.View.Slide
    On Error GoTo 0
    If lay Is Nothing Then
        Tell "Agenda Placeholder works on a slide layout: choose View > Slide Master, select the layout named """ & AGENDA_LAYOUT & """, then click Agenda Placeholder again."
        Exit Sub
    End If
    If TypeName(lay) <> "CustomLayout" Then
        Tell "Agenda Placeholder works on a slide layout: choose View > Slide Master, select the layout named """ & AGENDA_LAYOUT & """, then click Agenda Placeholder again."
        Exit Sub
    End If
    For Each shp In lay.Shapes
        If TagValue(shp, "OC_AGENDA_PH") = "1" Then shp.Select: Exit Sub
    Next shp
    w = ActivePresentation.PageSetup.SlideWidth: h = ActivePresentation.PageSetup.SlideHeight
    Set shp = lay.Shapes.AddShape(1, w * 0.3, h * 0.3, w * 0.4, h * 0.4)
    shp.Fill.Visible = 0
    shp.Line.Visible = 0
    shp.Name = "Agenda placeholder"
    shp.Tags.Add "OC_AGENDA_PH", "1"
    shp.Select
    If lay.Name <> AGENDA_LAYOUT Then Tell "Agenda placeholder added. Name this layout """ & AGENDA_LAYOUT & """ (exactly) so Chapter uses it."
End Sub

' Marks (or unmarks) a slide as the start of a new agenda section; the marker sits just above the
' slide, so it is seen while editing but never in the slide show or a printout.
Private Sub SetSplit(ByVal sld As Object, ByVal turnOn As Boolean)
    Dim shp As Object, k As Long
    For k = sld.Shapes.Count To 1 Step -1
        If TagValue(sld.Shapes(k), "OC_ROLE") = "agenda:split" Then sld.Shapes(k).Delete
    Next k
    If Not turnOn Then
        On Error Resume Next
        sld.Tags.Delete "OC_SPLIT"
        Exit Sub
    End If
    sld.Tags.Add "OC_SPLIT", "1"
    Set shp = sld.Shapes.AddTextbox(1, 0, -24, 220, 18)
    shp.Name = "Agenda split"
    shp.Tags.Add "OC_ROLE", "agenda:split"
    With shp.TextFrame2.TextRange
        .Text = "Agenda split: a new agenda starts here"
        .Font.Size = 9
        .Font.Fill.ForeColor.RGB = RGB(127, 127, 127)
    End With
End Sub

Private Function AgendaGroup(ByVal sld As Object) As Object
    Dim shp As Object
    For Each shp In sld.Shapes
        If TagValue(shp, "OC_AGENDA") <> "" Then Set AgendaGroup = shp: Exit Function
    Next shp
End Function

Private Function CurrentSlideIndex() As Long
    On Error Resume Next
    CurrentSlideIndex = ActiveWindow.View.Slide.SlideIndex
    If CurrentSlideIndex = 0 Then CurrentSlideIndex = ActiveWindow.Selection.SlideRange(1).SlideIndex
End Function

' The chapter of the selected heading (a shape of an agenda element, or text in it), else "".
Private Function SelectedHeadingChapter() As String
    Dim sel As Object, shp As Object, role As String
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    If sel.Type <> 2 And sel.Type <> 3 Then Exit Function
    Set shp = sel.ShapeRange(1)
    If sel.HasChildShapeRange Then Set shp = sel.ChildShapeRange(1)
    role = TagValue(shp, "OC_ROLE")
    If Left$(role, 8) = "chapter:" Then SelectedHeadingChapter = Mid$(role, 9)
End Function

' Selects the text of chapter id's heading on its own agenda slide, ready to type over.
Private Sub SelectHeading(ByVal id As String)
    Dim sld As Object, grp As Object, k As Long
    On Error Resume Next
    For Each sld In ActivePresentation.Slides
        If TagValue(sld, "OC_CHAPTER") = id Then
            ActiveWindow.View.GotoSlide sld.SlideIndex
            Set grp = AgendaGroup(sld)
            If grp Is Nothing Then Exit Sub
            For k = 1 To grp.GroupItems.Count
                If TagValue(grp.GroupItems(k), "OC_ROLE") = "chapter:" & id Then grp.GroupItems(k).TextFrame2.TextRange.Select: Exit Sub
            Next k
        End If
    Next sld
End Sub

' What the automatic update watches: each slide's id, kind, chapter and split flag, and whether the
' current slide's headings differ from what was drawn (OC_SIG: a rename or a new text level).
Private Function AgendaSignature() As String
    Dim pres As Object, sld As Object, parts() As String, i As Long, kind As String, anyAgenda As Boolean, grp As Object
    Set pres = ActivePresentation
    ReDim parts(0 To pres.Slides.Count)
    For i = 1 To pres.Slides.Count
        Set sld = pres.Slides(i)
        kind = TagValue(sld, "OC_KIND")
        If kind <> "" Then anyAgenda = True
        parts(i) = sld.SlideID & ":" & kind & ":" & TagValue(sld, "OC_CHAPTER") & ":" & TagValue(sld, "OC_SPLIT")
    Next i
    If Not anyAgenda Then Exit Function
    On Error Resume Next
    Set grp = AgendaGroup(ActiveWindow.View.Slide)
    On Error GoTo 0
    If Not grp Is Nothing Then
        If HeadingSig(grp) <> Replace(Replace(TagValue(grp, "OC_SIG"), vbCrLf, vbLf), vbCr, vbLf) Then parts(0) = "edited"
    End If
    AgendaSignature = pres.FullName & "|" & Join(parts, "|")
End Function

' The headings as signature lines: id, text level, text (src/core/agenda.ts headingSignature).
Private Function HeadingSig(ByVal grp As Object) As String
    Dim k As Long, shp As Object, role As String, tr As Object, out As String
    On Error Resume Next
    For k = 1 To grp.GroupItems.Count
        Set shp = grp.GroupItems(k)
        role = TagValue(shp, "OC_ROLE")
        If Left$(role, 8) = "chapter:" Then
            Set tr = shp.TextFrame2.TextRange
            out = out & IIf(out = "", "", vbLf) & Mid$(role, 9) & vbTab & tr.Paragraphs(1).ParagraphFormat.IndentLevel & vbTab & HeadingText(tr)
        End If
    Next k
    HeadingSig = out
End Function

' Selection changed: if agenda slides were added, duplicated, moved or deleted, or a heading was
' renamed or nested, update every agenda slide. Never while the user is typing in a heading.
Private Sub AgendaAutoCheck()
    Dim sig As String, presName As String, response As String
    If SelectedHeadingChapter() <> "" Then
        If ActiveWindow.Selection.Type = 3 Then Exit Sub ' ppSelectionText: still typing
    End If
    sig = AgendaSignature()
    If sig = "" Then gAgendaSig = "": Exit Sub
    presName = Split(sig, "|")(0)
    If gAgendaSig = "" Or Split(gAgendaSig & "|", "|")(0) <> presName Then
        ' First look at this presentation: remember it, unless a heading was edited.
        gAgendaSig = sig
        If InStr(sig, "|edited|") = 0 Then Exit Sub
    End If
    If sig = gAgendaSig And InStr(sig, "|edited|") = 0 Then Exit Sub
    response = Engine(AgendaRequest("update"))
    If gAbortAuto Then Exit Sub ' a user action arrived meanwhile; check again on the next selection
    If response <> "" Then ApplyAgenda response, ""
    gAgendaSig = AgendaSignature()
    Verbose "AGENDA auto-update"
End Sub

' A field of a request line: \ tab CR LF escaped as the engine reads them (format.ts unescapeField).
Private Function Esc(ByVal s As String) As String
    Esc = Replace(Replace(Replace(Replace(s, "\", "\\"), vbTab, "\t"), vbCr, "\r"), vbLf, "\n")
End Function

Private Function HexOf(ByVal c As Long) As String
    HexOf = Right$("0" & Hex$(c And 255), 2) & Right$("0" & Hex$((c \ 256) And 255), 2) & Right$("0" & Hex$((c \ 65536) And 255), 2)
End Function

' ---- TEST ONLY: entry points for the scripted live tests (tools/e2e-*.ts) ----------------------
' Called by name through AppleScript "run VB macro" or COM Application.Run; Public only for that.
' No ribbon control, menu item or other code of the add-in calls them (test/addin.test.ts),
' they take arguments (so PowerPoint's Macros dialog does not list them) and never show a dialog.

' Test only: runs one control id without dialogs, e.g. "chart_stacked", "btnUpdate". The outcome is logged
' by the bridge script (~/Library/Application Support/open-charts/addin.log).
Public Sub OCPREFIX_Run(ByVal controlId As String)
    Dim t As Single
    gSilent = True
    gLastMessage = ""
    t = Timer
    ResetTiming
    Dispatch controlId
    Report "RESULT " & controlId & " " & IIf(gLastMessage = "", "OK", gLastMessage) & " (" & Format$((Timer - t) * 1000, "0") & " ms)" & Timing()
    gSilent = False
End Sub

' Test only: selects the n-th open-charts chart on the active slide (1-based), or nothing for "0" (as a
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

' Test only: selects several charts of the active slide, e.g. "1,2" (as Shift-click would), for scripted tests.
Public Sub OCPREFIX_SelectCharts(ByVal list As String)
    Dim shp As Object, k As Long, first As Boolean
    On Error GoTo Failed
    first = True
    For Each shp In ActiveWindow.View.Slide.Shapes
        If IsChart(shp) Then
            k = k + 1
            If InStr("," & list & ",", "," & k & ",") > 0 Then shp.Select IIf(first, -1, 0): first = False
        End If
    Next shp
Failed:
End Sub

' Test only: selects the piece with OC_ROLE `role` of the n-th chart (as clicking into the group would), for
' scripted tests.
Public Sub OCPREFIX_SelectRole(ByVal n As String, ByVal role As String)
    Dim shp As Object, k As Long, i As Long
    On Error GoTo Failed
    For Each shp In ActiveWindow.View.Slide.Shapes
        If IsChart(shp) Then
            k = k + 1
            If k = CLng(Val(n)) Then
                ' From no selection: selecting a piece of the group that is selected already selects nothing on a Mac.
                ActiveWindow.Selection.Unselect
                For i = 1 To shp.GroupItems.Count
                    If TagValue(shp.GroupItems(i), "OC_ROLE") = role Then shp.GroupItems(i).Select: Exit Sub
                Next i
                Report "SELECTROLE " & role & " not in chart " & n
            End If
        End If
    Next shp
    Exit Sub
Failed:
    Report "SELECTROLE " & role & ": " & Err.Description
End Sub

' Test only: logs where the n-th chart is on screen ("WHERE screen=left,top,right,bottom"), and which
' right-click items and ribbon commands apply to the selection ("APPLIES btnTotals btnCagr ..."),
' for scripted tests.
Public Sub OCPREFIX_Probe(ByVal n As String)
    Dim shp As Object, k As Long, c As Variant, out As String
    On Error Resume Next
    For Each shp In ActiveWindow.View.Slide.Shapes
        If IsChart(shp) Then
            k = k + 1
            If k = CLng(Val(n)) Then Report "WHERE " & Replace(ScreenLines(shp), vbLf, " ")
        End If
    Next shp
    For Each c In Array("btnEditData", "btnUpdate", "btnTotals", "btnCagr", "btnDifference", "btnSameScale", "btnUnscale", "btnPlainText", "btnResetLabel")
        If CommandApplies(CStr(c)) Then out = out & " " & c
    Next c
    Report "APPLIES" & out
    Report "SELECTION " & ProbeSelection()
    Report "VISIBLECALLS " & gVisibleCalls
End Sub

' Test only: what PowerPoint reports as selected (type, pieces of a group, their roles).
Private Function ProbeSelection() As String
    Dim sel As Object, i As Long, shp As Object, out As String
    On Error Resume Next
    Set sel = ActiveWindow.Selection
    out = "type=" & sel.Type & " child=" & sel.HasChildShapeRange
    If sel.HasChildShapeRange Then
        out = out & " n=" & sel.ChildShapeRange.Count
        For i = 1 To sel.ChildShapeRange.Count
            Set shp = sel.ChildShapeRange(i)
            out = out & " [" & shp.Name & "|" & TagValue(shp, "OC_ROLE") & "]"
        Next i
    ElseIf sel.Type = 2 Then
        out = out & " shapes=" & sel.ShapeRange.Count & " [" & sel.ShapeRange(1).Name & "]"
    End If
    ProbeSelection = out & " roles=" & SelectedRoles(FindSelectedChart())
End Function

' Test only: resizes the n-th chart on the active slide by a factor, as a user dragging a corner would.
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

' Test only: exports the active presentation to `dest` outside PowerPoint's sandbox: kind "pptx" saves a
' copy, "png" one image per slide (slide1.png ...). PowerPoint writes into its sandbox temp
' folder and the bridge copies the result out.
Public Sub OCPREFIX_ExportCopy(ByVal dest As String, ByVal kind As String)
    Dim tmp As String, i As Long
    On Error GoTo Failed
#If Mac Then
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
#Else
    ' No sandbox on Windows: write straight to dest.
    If kind = "png" Then
        If Dir(dest, vbDirectory) = "" Then MkDir dest
        For i = 1 To ActivePresentation.Slides.Count
            ActivePresentation.Slides(i).Export dest & "\slide" & i & ".png", "PNG", 1920, 1080
        Next i
    Else
        ActivePresentation.SaveCopyAs dest
    End If
    Report "EXPORT OK " & dest
#End If
    Exit Sub
Failed:
    Report "EXPORT failed: " & Err.Description
End Sub

' Test only: converts the label with role `role` of the selected chart to plain text.
Public Sub OCPREFIX_ConvertToPlainText(ByVal role As String)
    Dim chart As Object, response As String
    Set chart = FindSelectedChart()
    If chart Is Nothing Then Report "PLAINTEXT failed: no chart selected": Exit Sub
    response = Engine("detach" & vbLf & role & vbLf & ChartContext(chart) & SpecOf(chart))
    If response <> "" Then Set chart = ReplaceChart(chart, response)
    Report "PLAINTEXT " & IIf(chart Is Nothing, "failed: " & gLastMessage, "OK")
End Sub

' Test only: types `text` over chapter `title`'s heading on slide n and sets its text level (0 = unchanged).
Public Sub OCPREFIX_AgendaEdit(ByVal n As String, ByVal heading As String, ByVal newText As String, ByVal level As String)
    Dim grp As Object, k As Long, tr As Object
    On Error GoTo Failed
    Set grp = AgendaGroup(ActivePresentation.Slides(CLng(n)))
    For k = 1 To grp.GroupItems.Count
        Set tr = grp.GroupItems(k).TextFrame2.TextRange
        If HeadingText(tr) = heading Then
            If newText <> "" Then tr.Text = newText
            If Val(level) > 0 Then tr.Paragraphs(1).ParagraphFormat.IndentLevel = CLng(Val(level))
            Report "AGENDAEDIT OK"
            Exit Sub
        End If
    Next k
Failed:
    Report "AGENDAEDIT failed: no heading """ & heading & """ on slide " & n
End Sub

' Test only: goes to slide n and selects the heading `title` there ("" = select nothing), as a click would.
Public Sub OCPREFIX_AgendaSelect(ByVal n As String, ByVal heading As String)
    Dim grp As Object, k As Long
    On Error Resume Next
    ActiveWindow.View.GotoSlide CLng(n)
    ActiveWindow.Selection.Unselect
    If heading = "" Then Exit Sub
    Set grp = AgendaGroup(ActivePresentation.Slides(CLng(n)))
    For k = 1 To grp.GroupItems.Count
        If HeadingText(grp.GroupItems(k).TextFrame2.TextRange) = heading Then grp.GroupItems(k).Select: Exit Sub
    Next k
End Sub

' Test only: slide sorter operations: "move" from to, "duplicate" n, "delete" n, "none"; then the selection
' change PowerPoint raises (the automatic update).
Public Sub OCPREFIX_AgendaSorter(ByVal op As String, ByVal a As String, ByVal b As String)
    On Error GoTo Failed
    Select Case op
        Case "move": ActivePresentation.Slides(CLng(a)).MoveTo CLng(b)
        Case "duplicate": ActivePresentation.Slides(CLng(a)).Duplicate
        Case "delete": ActivePresentation.Slides(CLng(a)).Delete
    End Select
    gSilent = True
    OnSelectionChange
    gSilent = False
    Report "SORTER " & op & " OK"
    Exit Sub
Failed:
    gSilent = False
    Report "SORTER " & op & " failed: " & Err.Description
End Sub

' Test only: lists every slide for the test log: n, kind, hidden, the agenda headings (">" per level, "*" current).
Public Sub OCPREFIX_AgendaDump(ByVal withIndents As String)
    Dim sld As Object, grp As Object, k As Long, out As String, row As String, role As String, tr As Object
    For Each sld In ActivePresentation.Slides
        row = "SLIDE " & sld.SlideIndex & " " & IIf(TagValue(sld, "OC_KIND") = "", "-", TagValue(sld, "OC_KIND")) & IIf(sld.SlideShowTransition.Hidden, " hidden", "") & IIf(TagValue(sld, "OC_SPLIT") = "1", " split", "")
        Set grp = AgendaGroup(sld)
        If Not grp Is Nothing Then
            row = row & ":"
            For k = 1 To grp.GroupItems.Count
                role = TagValue(grp.GroupItems(k), "OC_ROLE")
                If Left$(role, 8) = "chapter:" Then
                    Set tr = grp.GroupItems(k).TextFrame2.TextRange
                    row = row & " " & String$(tr.Paragraphs(1).ParagraphFormat.IndentLevel - 1, ">") & HeadingText(tr) & IIf(Mid$(role, 9) = TagValue(sld, "OC_CHAPTER"), "*", "")
                    If withIndents = "1" Then row = row & "[" & Fmt(tr.Paragraphs(1).ParagraphFormat.LeftIndent) & "/" & Fmt(tr.Paragraphs(1).ParagraphFormat.FirstLineIndent) & "/" & Fmt(grp.GroupItems(k).Width) & "]"
                End If
            Next k
        End If
        out = out & row & " ;; "
    Next sld
    Report "DUMP " & out
End Sub

' Test only: Logs every text of the n-th chart as PowerPoint draws it, one line each: "TEXT role|font|bound width|
' box width|bound height|box height" (pt), so a scripted test can compare
' PowerPoint's text with what the engine measured (tools/e2e-fonts-live.ts).
Public Sub OCPREFIX_TextBounds(ByVal n As String)
    Dim shp As Object, s As Object, tr As Object, k As Long, i As Long
    On Error Resume Next
    For Each shp In ActiveWindow.View.Slide.Shapes
        If IsChart(shp) Then
            k = k + 1
            If k = CLng(Val(n)) Then
                For i = 1 To shp.GroupItems.Count
                    Set s = shp.GroupItems(i)
                    If s.HasTextFrame Then
                        If s.TextFrame2.HasText Then
                            Set tr = s.TextFrame2.TextRange
                            Report "TEXT " & TagValue(s, "OC_ROLE") & "|" & tr.Font.Name & "|" & Fmt(tr.BoundWidth) & "|" & _
                                Fmt(s.Width) & "|" & Fmt(tr.BoundHeight) & "|" & Fmt(s.Height) & "|" & _
                                Fmt(tr.Characters(1, tr.Length).BoundWidth) & "|" & tr.Length & "|" & Replace(tr.Text, " ", "_")
                        End If
                    End If
                Next i
            End If
        End If
    Next shp
End Sub
