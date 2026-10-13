// open-charts pop-up helper (Mac): a background app (LSUIElement, no Dock icon) that hosts the
// add-in's floating panels — the "Edit data" datasheet, and later the element mini toolbar — as
// always-on-top NSPanels over PowerPoint, each with a WKWebView whose page is preloaded from the
// add-in server, so showing one is a JavaScript call, not a page load.
//
// Started by the add-in server (src/addin/popup.ts), which talks to it over stdin/stdout, one JSON
// object per line (the protocol is documented there). It exits when stdin closes, so it never
// outlives the server. Built by `npm run build:popup` (swiftc, Command Line Tools; no Xcode project)
// and ad-hoc signed.
//
// Pages talk back through window.webkit.messageHandlers.openCharts.postMessage({type, …}):
//   {type:"ready"}   the session is loaded and drawn: the panel may be shown
//   {type:"close"}   hide the panel (Done, Cancel, Escape)
import Cocoa
import WebKit

setvbuf(stdout, nil, _IOLBF, 0)

func emit(_ obj: [String: Any]) {
    guard let data = try? JSONSerialization.data(withJSONObject: obj), var line = String(data: data, encoding: .utf8) else { return }
    line += "\n"
    FileHandle.standardOutput.write(line.data(using: .utf8)!)
}

/// Global screen rectangles with the origin at the top left of the main display (as CGWindowList
/// reports windows and the server computes positions), converted to and from Cocoa's bottom-left.
func mainHeight() -> CGFloat { NSScreen.screens.first?.frame.height ?? 900 }
func topLeft(_ r: NSRect) -> [Double] { [Double(r.minX), Double(mainHeight() - r.maxY), Double(r.width), Double(r.height)] }
func cocoa(_ a: [Double]) -> NSRect { NSRect(x: a[0], y: Double(mainHeight()) - a[1] - a[3], width: a[2], height: a[3]) }

final class Panel: NSPanel {
    var onCancel: (() -> Void)?
    override var canBecomeKey: Bool { true }
    override var canBecomeMain: Bool { false }
    // Escape that the page did not handle (focus outside the web view).
    override func cancelOperation(_ sender: Any?) { onCancel?() }
    // An app without a main menu gets no Edit-menu key equivalents: route ⌘C/⌘X/⌘V/⌘A that the
    // page did not handle to the web view (the editor's own copy/paste listens for these events).
    override func performKeyEquivalent(with event: NSEvent) -> Bool {
        if super.performKeyEquivalent(with: event) { return true }
        guard event.modifierFlags.intersection(.deviceIndependentFlagsMask) == .command, let key = event.charactersIgnoringModifiers else { return false }
        let actions: [String: Selector] = ["c": #selector(NSText.copy(_:)), "x": #selector(NSText.cut(_:)), "v": #selector(NSText.paste(_:)), "a": #selector(NSText.selectAll(_:))]
        guard let action = actions[key] else { return false }
        return NSApp.sendAction(action, to: nil, from: self)
    }
}

/// One floating panel and its web page.
final class Host: NSObject, WKScriptMessageHandler, WKNavigationDelegate, NSWindowDelegate {
    let name: String
    let panel: Panel
    let web: WKWebView
    var loaded = false
    var open = false            // a session is showing (or about to)
    var suspended = false       // hidden because another app came to the front
    var pending: (start: Date, focus: Bool, timer: Timer?)?

    init(name: String, utility: Bool) {
        self.name = name
        let config = WKWebViewConfiguration()
        let ucc = WKUserContentController()
        config.userContentController = ucc
        config.suppressesIncrementalRendering = false
        web = WKWebView(frame: NSRect(x: 0, y: 0, width: 620, height: 360), configuration: config)
        // The datasheet: a small utility window (title bar to drag it, close = Done, resizable). The
        // toolbar: borderless, never takes the keyboard from PowerPoint.
        let style: NSWindow.StyleMask = utility ? [.nonactivatingPanel, .titled, .closable, .resizable, .utilityWindow] : [.nonactivatingPanel, .borderless]
        panel = Panel(contentRect: NSRect(x: 0, y: 0, width: 620, height: 360), styleMask: style, backing: .buffered, defer: false)
        super.init()
        ucc.add(self, name: "openCharts")
        web.navigationDelegate = self
        web.autoresizingMask = [.width, .height]
        panel.title = "Edit data"
        panel.hasShadow = true
        panel.level = .floating
        panel.isFloatingPanel = true
        panel.hidesOnDeactivate = false
        panel.becomesKeyOnlyIfNeeded = false
        panel.isReleasedWhenClosed = false
        panel.collectionBehavior = [.moveToActiveSpace, .fullScreenAuxiliary]
        panel.contentView = web
        panel.delegate = self
        panel.onCancel = { [weak self] in self?.js("window.ocPopupClose ? ocPopupClose('cancel') : 0") }
    }

    func js(_ source: String, done: ((Any?, Error?) -> Void)? = nil) {
        web.evaluateJavaScript(source) { r, e in done?(r, e) }
    }

    func load(_ url: String) {
        guard let u = URL(string: url) else { return }
        loaded = false
        web.load(URLRequest(url: u, cachePolicy: .reloadIgnoringLocalCacheData))
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        loaded = true
        emit(["event": "loaded", "panel": name])
    }
    func webView(_ webView: WKWebView, didFail navigation: WKNavigation!, withError error: Error) {
        emit(["event": "error", "panel": name, "message": error.localizedDescription])
    }
    func webView(_ webView: WKWebView, didFailProvisionalNavigation navigation: WKNavigation!, withError error: Error) {
        emit(["event": "error", "panel": name, "message": error.localizedDescription])
    }
    // The page went away (WebContent process crashed): load it again next time.
    func webViewWebContentProcessDidTerminate(_ webView: WKWebView) { loaded = false }

    func show(id: Any, url: String, frame: [Double], focus: Bool) {
        placing = true
        panel.setFrame(cocoa(frame), display: false)
        placing = false
        open = true
        suspended = false
        pending?.timer?.invalidate()
        let t = Timer.scheduledTimer(withTimeInterval: 0.6, repeats: false) { [weak self] _ in self?.reveal(ready: false) }
        pending = (Date(), focus, t)
        let arg = String(data: try! JSONSerialization.data(withJSONObject: [id]), encoding: .utf8)!
        if loaded {
            js("typeof window.ocPopupOpen === 'function' ? (window.ocPopupOpen(...\(arg)), true) : false") { [weak self] r, _ in
                if (r as? Bool) != true { self?.load(url) }
            }
        } else {
            load(url)
        }
    }

    func reveal(ready: Bool) {
        guard let p = pending else { return }
        p.timer?.invalidate()
        pending = nil
        if p.focus {
            panel.makeKeyAndOrderFront(nil)
            panel.makeFirstResponder(web)
        } else {
            panel.orderFrontRegardless()
        }
        emit(["event": "shown", "panel": name, "ms": Int(Date().timeIntervalSince(p.start) * 1000), "ready": ready])
    }

    func hide(reason: String) {
        pending?.timer?.invalidate()
        pending = nil
        open = false
        suspended = false
        panel.orderOut(nil)
        emit(["event": "closed", "panel": name, "reason": reason])
    }

    func userContentController(_ ucc: WKUserContentController, didReceive message: WKScriptMessage) {
        guard let body = message.body as? [String: Any], let type = body["type"] as? String else { return }
        switch type {
        case "ready":
            if let title = body["title"] as? String, !title.isEmpty { panel.title = title }
            if open { reveal(ready: true) }
        case "close": hide(reason: (body["reason"] as? String) ?? "page")
        case "resize":
            // Keep the top-left corner where the server placed it.
            if let w = body["w"] as? Double, let h = body["h"] as? Double, w > 20, h > 20 {
                let f = panel.frame
                placing = true
                panel.setFrame(NSRect(x: f.minX, y: f.maxY - h, width: min(w, 900), height: h), display: true)
                placing = false
            }
        default: emit(["event": "page", "panel": name, "body": body])
        }
    }

    // The close button: Done (keeps what was applied), handled by the page like ⌘↵.
    // Where the user put the window (dragged or resized it), remembered for every chart and across
    // restarts: the server places the next pop-up by it (info "saved").
    var placing = false
    var savedKey: String { "frame.\(name)" }
    func remember() {
        // Only the user moves a window with the mouse button down; anything else (placement, macOS
        // pulling a window onto a display) is not a choice to remember.
        guard open, !placing, pending == nil, panel.isVisible, NSEvent.pressedMouseButtons & 1 == 1 else { return }
        UserDefaults.standard.set(topLeft(panel.frame), forKey: savedKey)
    }
    func windowDidMove(_ notification: Notification) { remember() }
    func windowDidEndLiveResize(_ notification: Notification) { remember() }

    func windowShouldClose(_ sender: NSWindow) -> Bool {
        js("window.ocPopupClose ? (ocPopupClose('done'), true) : false") { [weak self] r, _ in
            if (r as? Bool) != true { self?.hide(reason: "closed") }
        }
        return false
    }
}

// ---- Drag handles: a transparent overlay over the selected chart (src/addin/handles.ts) ---------
//
// A borderless, non-activating, transparent panel over the chart. It ignores the mouse (every click
// goes to PowerPoint) except while the pointer is over one of its handles: a global mouse-moved
// monitor (no Accessibility needed for mouse events) turns ignoresMouseEvents off there. A press on
// a handle then belongs to the panel: AppKit sends it the drag and the release (also outside the
// panel), PowerPoint sees none of them — the overlay holds the mouse for the whole gesture. Each move
// goes to the server ({"event":"handle","phase":"move"}), which answers with the re-laid-out chart
// ({"cmd":"ghost"}), drawn here over the chart; the release commits ("up"); a double-click resets.
//
// It hides while another app is in front, during PowerPoint's own drags, on scrolling or zooming
// and when PowerPoint's window moves, and asks to be re-measured ({"event":"remeasure"}) once that
// settles: the server then asks PowerPoint where the chart is now.
//
// server -> helper:
//   {"cmd":"handles","frame":[x,y,w,h],"handles":[{"id","rect":[x,y,w,h],"marker":bool,"cursor":"ew|ns|move"}]}
//   {"cmd":"handles","hide":true}
//   {"cmd":"ghost","shapes":[…],"tip":S} / {"cmd":"ghost","clear":true}
//   {"cmd":"handlesState"} -> {"event":"handlesState",…} (tests)
// Rectangles are global screen points, top left of the main display.

struct HandleBox { let id: String; let rect: NSRect; let marker: Bool; let cursor: String }

final class OverlayView: NSView {
    weak var overlay: HandleOverlay?
    override var isFlipped: Bool { true }
    override func acceptsFirstMouse(for event: NSEvent?) -> Bool { true }
    override func draw(_ dirtyRect: NSRect) { overlay?.draw() }
    override func mouseDown(with event: NSEvent) { overlay?.down(event) }
    override func mouseDragged(with event: NSEvent) { overlay?.dragged(event) }
    override func mouseUp(with event: NSEvent) { overlay?.up(event) }
    override func rightMouseDown(with event: NSEvent) { overlay?.cancel() }
    override func updateTrackingAreas() {
        super.updateTrackingAreas()
        for t in trackingAreas { removeTrackingArea(t) }
        addTrackingArea(NSTrackingArea(rect: .zero, options: [.activeAlways, .inVisibleRect, .mouseMoved, .cursorUpdate], owner: self, userInfo: nil))
    }
    override func mouseMoved(with event: NSEvent) { overlay?.hover() }
    override func cursorUpdate(with event: NSEvent) { overlay?.setCursor() }
}

final class HandleOverlay {
    let panel: NSPanel
    let view = OverlayView()
    let powerPoint: String
    var frameTL = NSRect.zero          // the panel, top-left global points
    var handles: [HandleBox] = []
    var ghost: [[String: Any]] = []
    var tip = ""
    var active = false                 // the server wants handles shown
    var stale = false                  // hidden until re-measured (scroll, zoom, PowerPoint's own drag)
    var suspended = false              // another app is in front
    var dragging: String?
    var hovering: String?
    var dragPoint = NSPoint.zero       // last pointer position while dragging (top-left global)
    var ghostTimer: Timer?
    var quietTimer: Timer?
    var pollTimer: Timer?
    var windowsSeen: [[Double]] = []
    var pptDrag = false
    var onDrag: ((Bool) -> Void)?      // the mini toolbar steps aside while a handle is dragged
    var windows: () -> [[Double]] = { [] }

    init(powerPoint: String) {
        self.powerPoint = powerPoint
        panel = NSPanel(contentRect: NSRect(x: 0, y: 0, width: 10, height: 10), styleMask: [.nonactivatingPanel, .borderless], backing: .buffered, defer: false)
        panel.isOpaque = false
        panel.backgroundColor = .clear
        panel.hasShadow = false
        panel.level = .floating
        panel.isFloatingPanel = true
        panel.hidesOnDeactivate = false
        panel.becomesKeyOnlyIfNeeded = true
        panel.isReleasedWhenClosed = false
        panel.ignoresMouseEvents = true
        panel.acceptsMouseMovedEvents = true
        panel.collectionBehavior = [.moveToActiveSpace, .fullScreenAuxiliary]
        view.overlay = self
        view.wantsLayer = true
        panel.contentView = view
        watch()
    }

    var visible: Bool { active && !suspended && (!stale || dragging != nil) }

    static func mouseTL() -> NSPoint {
        let p = NSEvent.mouseLocation
        return NSPoint(x: p.x, y: mainHeight() - p.y)
    }

    func hit(_ p: NSPoint) -> HandleBox? {
        guard visible, ghost.isEmpty || dragging != nil else { return nil }
        // Markers are small: a little slack around them.
        return handles.last { h in (h.marker ? h.rect.insetBy(dx: -3, dy: -3) : h.rect).contains(p) }
    }

    /// Mouse privilege only over a handle: everywhere else the panel lets clicks through.
    func hover() {
        if dragging != nil { return }
        let h = hit(HandleOverlay.mouseTL())
        let ignore = h == nil
        if panel.ignoresMouseEvents != ignore { panel.ignoresMouseEvents = ignore }
        if h?.id != hovering { hovering = h?.id; view.needsDisplay = true; setCursor() }
    }

    func setCursor() {
        guard let id = dragging ?? hovering, let h = handles.first(where: { $0.id == id }) else { return }
        switch h.cursor {
        case "ew": NSCursor.resizeLeftRight.set()
        case "ns": NSCursor.resizeUpDown.set()
        default: (dragging != nil ? NSCursor.closedHand : NSCursor.openHand).set()
        }
    }

    func refresh() {
        if visible {
            if !panel.isVisible { panel.orderFrontRegardless() }
        } else {
            if panel.isVisible { panel.orderOut(nil) }
            panel.ignoresMouseEvents = true
            hovering = nil
        }
        view.needsDisplay = true
    }

    // -- Commands

    func command(_ m: [String: Any]) {
        switch m["cmd"] as? String {
        case "handles":
            ghostTimer?.invalidate()
            ghost = []
            tip = ""
            if (m["hide"] as? Bool) == true {
                active = false
                handles = []
                if dragging != nil { dragging = nil; onDrag?(false) }
            } else if let f = m["frame"] as? [Double], f.count == 4 {
                frameTL = NSRect(x: f[0], y: f[1], width: f[2], height: f[3])
                handles = ((m["handles"] as? [[String: Any]]) ?? []).compactMap { h in
                    guard let id = h["id"] as? String, let r = h["rect"] as? [Double], r.count == 4 else { return nil }
                    return HandleBox(id: id, rect: NSRect(x: r[0], y: r[1], width: r[2], height: r[3]), marker: (h["marker"] as? Bool) ?? true, cursor: (h["cursor"] as? String) ?? "move")
                }
                panel.setFrame(cocoa([f[0], f[1], f[2], f[3]]), display: false)
                active = true
                stale = false
                windowsSeen = windows()
                if let app = NSWorkspace.shared.frontmostApplication {
                    suspended = !(app.bundleIdentifier == powerPoint || app.processIdentifier == ProcessInfo.processInfo.processIdentifier)
                }
            }
            refresh()
            hover()
            poll()
        case "ghost":
            if (m["clear"] as? Bool) == true {
                ghost = []
                tip = ""
            } else {
                ghost = (m["shapes"] as? [[String: Any]]) ?? []
                tip = (m["tip"] as? String) ?? ""
                if dragging == nil { keepGhostBriefly() }
            }
            view.needsDisplay = true
        case "handlesState":
            emit(["event": "handlesState", "active": active, "visible": panel.isVisible, "stale": stale, "suspended": suspended,
                  "ignoresMouse": panel.ignoresMouseEvents, "frame": [Double(frameTL.minX), Double(frameTL.minY), Double(frameTL.width), Double(frameTL.height)],
                  "handles": handles.map { ["id": $0.id, "rect": [Double($0.rect.minX), Double($0.rect.minY), Double($0.rect.width), Double($0.rect.height)]] },
                  "ghost": ghost.count, "dragging": dragging ?? ""])
        default: break
        }
    }

    /// After a commit the ghost stands in for the chart until the server sends the redrawn handles.
    func keepGhostBriefly() {
        ghostTimer?.invalidate()
        ghostTimer = Timer.scheduledTimer(withTimeInterval: 4, repeats: false) { [weak self] _ in
            guard let self, self.dragging == nil else { return }
            self.ghost = []; self.tip = ""; self.view.needsDisplay = true
        }
    }

    // -- The drag (the panel holds the mouse from press to release)

    func send(_ phase: String, _ id: String, _ p: NSPoint) {
        emit(["event": "handle", "phase": phase, "id": id, "x": Double(p.x), "y": Double(p.y)])
    }

    func down(_ e: NSEvent) {
        let p = HandleOverlay.mouseTL()
        guard let h = hit(p) else { panel.ignoresMouseEvents = true; return }
        if e.clickCount >= 2 { send("reset", h.id, p); return }
        dragging = h.id
        dragPoint = p
        onDrag?(true)
        setCursor()
        send("down", h.id, p)
        view.needsDisplay = true
    }

    func dragged(_ e: NSEvent) {
        guard let id = dragging else { return }
        let p = HandleOverlay.mouseTL()
        dragPoint = p
        send("move", id, p)
        view.needsDisplay = true
    }

    func up(_ e: NSEvent) {
        guard let id = dragging else { return }
        let p = HandleOverlay.mouseTL()
        dragging = nil
        send("up", id, p)
        onDrag?(false)
        keepGhostBriefly()
        hover()
    }

    func cancel() {
        guard let id = dragging else { return }
        dragging = nil
        send("cancel", id, HandleOverlay.mouseTL())
        ghost = []; tip = ""
        onDrag?(false)
        refresh()
    }

    // -- Hide rules

    /// Hidden until the server has measured the chart again (after things settle).
    func goStale(after: TimeInterval = 0.3) {
        guard active, dragging == nil else { return }
        if !stale { stale = true; refresh() }
        quietTimer?.invalidate()
        quietTimer = Timer.scheduledTimer(withTimeInterval: after, repeats: false) { [weak self] _ in
            guard let self, self.active, !self.pptDrag else { return }
            emit(["event": "remeasure"])
        }
    }

    func watch() {
        // PowerPoint's own drags (moving shapes, the zoom slider, scroll bars): out of the way.
        NSEvent.addGlobalMonitorForEvents(matching: [.mouseMoved]) { [weak self] _ in self?.hover() }
        NSEvent.addLocalMonitorForEvents(matching: [.mouseMoved]) { [weak self] e in self?.hover(); return e }
        NSEvent.addGlobalMonitorForEvents(matching: [.leftMouseDragged]) { [weak self] _ in
            guard let self, self.active, self.dragging == nil else { return }
            self.pptDrag = true
            if !self.stale { self.stale = true; self.refresh() }
        }
        NSEvent.addGlobalMonitorForEvents(matching: [.leftMouseUp]) { [weak self] _ in
            guard let self, self.active else { return }
            let wasDrag = self.pptDrag
            self.pptDrag = false
            // A drag, or a click outside the chart (ribbon, zoom buttons): the chart may have moved.
            // A click on the slide changes the selection, which re-reports it anyway.
            if wasDrag || !self.frameTL.contains(HandleOverlay.mouseTL()) { self.goStale(after: 0.25) }
        }
        NSEvent.addGlobalMonitorForEvents(matching: [.scrollWheel, .magnify, .smartMagnify]) { [weak self] _ in self?.goStale() }
    }

    /// While shown: PowerPoint's windows moved or resized -> re-measure.
    func poll() {
        pollTimer?.invalidate()
        guard active else { return }
        pollTimer = Timer.scheduledTimer(withTimeInterval: 0.5, repeats: true) { [weak self] t in
            guard let self, self.active else { t.invalidate(); return }
            let now = self.windows()
            if now != self.windowsSeen { self.windowsSeen = now; self.goStale() }
        }
    }

    func appChanged(ours: Bool) {
        let s = !ours
        if s && dragging != nil { cancel() }
        if s != suspended { suspended = s; refresh() }
    }

    // -- Drawing (view coordinates: flipped, origin at the panel's top left)

    static func color(_ v: Any?, alpha: CGFloat = 1) -> NSColor? {
        guard let s = v as? String, s.count == 6, let n = UInt32(s, radix: 16) else { return nil }
        return NSColor(srgbRed: CGFloat((n >> 16) & 255) / 255, green: CGFloat((n >> 8) & 255) / 255, blue: CGFloat(n & 255) / 255, alpha: alpha)
    }
    func d(_ m: [String: Any], _ k: String) -> CGFloat { CGFloat((m[k] as? Double) ?? Double((m[k] as? Int) ?? 0)) }
    func local(_ x: CGFloat, _ y: CGFloat) -> NSPoint { NSPoint(x: x - frameTL.minX, y: y - frameTL.minY) }

    func draw() {
        guard visible else { return }
        for s in ghost { drawShape(s) }
        if !ghost.isEmpty, dragging != nil, !tip.isEmpty { drawTip() }
        if ghost.isEmpty { drawHandles() }
    }

    func drawHandles() {
        let ring = NSColor(srgbRed: 0.35, green: 0.25, blue: 0, alpha: 1)
        for h in handles {
            let r = NSRect(origin: local(h.rect.minX, h.rect.minY), size: h.rect.size)
            let hot = h.id == hovering || h.id == dragging
            if h.marker {
                // PowerPoint's adjustment-handle idiom: a yellow dot.
                let p = NSBezierPath(ovalIn: hot ? r.insetBy(dx: -1.5, dy: -1.5) : r)
                NSColor(srgbRed: 1, green: 0.83, blue: 0.1, alpha: 1).setFill()
                p.fill()
                ring.setStroke()
                p.lineWidth = 1
                p.stroke()
            } else if hot {
                let p = NSBezierPath(rect: r.insetBy(dx: -1, dy: -1))
                p.setLineDash([3, 2], count: 2, phase: 0)
                NSColor(srgbRed: 0.95, green: 0.65, blue: 0, alpha: 1).setStroke()
                p.lineWidth = 1
                p.stroke()
            }
        }
    }

    func drawTip() {
        let attrs: [NSAttributedString.Key: Any] = [.font: NSFont.systemFont(ofSize: 11, weight: .medium), .foregroundColor: NSColor.white]
        let text = NSAttributedString(string: tip, attributes: attrs)
        let size = text.size()
        var o = local(dragPoint.x + 14, dragPoint.y + 16)
        o.x = min(o.x, view.bounds.width - size.width - 10)
        o.y = min(o.y, view.bounds.height - size.height - 6)
        let box = NSRect(x: o.x, y: o.y, width: size.width + 10, height: size.height + 4)
        NSColor(white: 0.15, alpha: 0.88).setFill()
        NSBezierPath(roundedRect: box, xRadius: 4, yRadius: 4).fill()
        text.draw(at: NSPoint(x: box.minX + 5, y: box.minY + 2))
    }

    func drawShape(_ s: [String: Any]) {
        switch s["t"] as? String {
        case "r":
            let o = local(d(s, "x"), d(s, "y"))
            let r = NSRect(x: o.x, y: o.y, width: d(s, "w"), height: d(s, "h"))
            let p: NSBezierPath
            switch s["g"] as? String {
            case "ellipse": p = NSBezierPath(ovalIn: r)
            case "roundRect": p = NSBezierPath(roundedRect: r, xRadius: min(r.width, r.height) * 0.17, yRadius: min(r.width, r.height) * 0.17)
            case "diamond":
                p = NSBezierPath(); p.move(to: NSPoint(x: r.midX, y: r.minY)); p.line(to: NSPoint(x: r.maxX, y: r.midY)); p.line(to: NSPoint(x: r.midX, y: r.maxY)); p.line(to: NSPoint(x: r.minX, y: r.midY)); p.close()
            case "triangle":
                p = NSBezierPath(); p.move(to: NSPoint(x: r.midX, y: r.minY)); p.line(to: NSPoint(x: r.maxX, y: r.maxY)); p.line(to: NSPoint(x: r.minX, y: r.maxY)); p.close()
            default: p = NSBezierPath(rect: r)
            }
            if let f = HandleOverlay.color(s["f"]) { f.setFill(); p.fill() }
            if let c = HandleOverlay.color(s["s"]) { c.setStroke(); p.lineWidth = max(0.5, d(s, "sw")); p.stroke() }
        case "l":
            guard let c = HandleOverlay.color(s["c"]) else { return }
            let a = local(d(s, "x1"), d(s, "y1")), b = local(d(s, "x2"), d(s, "y2"))
            let p = NSBezierPath(); p.move(to: a); p.line(to: b)
            p.lineWidth = max(0.5, d(s, "w"))
            if s["d"] != nil { p.setLineDash([3, 2], count: 2, phase: 0) }
            c.set(); p.stroke()
            let head = { (from: NSPoint, to: NSPoint) in
                let ang = atan2(to.y - from.y, to.x - from.x), len = max(5, p.lineWidth * 4)
                let h = NSBezierPath(); h.move(to: to)
                h.line(to: NSPoint(x: to.x - len * cos(ang - 0.4), y: to.y - len * sin(ang - 0.4)))
                h.line(to: NSPoint(x: to.x - len * cos(ang + 0.4), y: to.y - len * sin(ang + 0.4))); h.close(); h.fill()
            }
            if s["ae"] != nil { head(a, b) }
            if s["as"] != nil { head(b, a) }
        case "p":
            guard let pts = s["p"] as? [Double], pts.count >= 4 else { return }
            let p = NSBezierPath()
            p.move(to: local(CGFloat(pts[0]), CGFloat(pts[1])))
            var i = 2
            while i + 1 < pts.count { p.line(to: local(CGFloat(pts[i]), CGFloat(pts[i + 1]))); i += 2 }
            if (s["z"] as? Int) == 1 { p.close() }
            if let f = HandleOverlay.color(s["f"]) { f.setFill(); p.fill() }
            if let c = HandleOverlay.color(s["s"]) { c.setStroke(); p.lineWidth = max(0.5, d(s, "sw")); p.stroke() }
        case "x":
            let o = local(d(s, "x"), d(s, "y"))
            var r = NSRect(x: o.x, y: o.y, width: d(s, "w"), height: d(s, "h"))
            let fill = HandleOverlay.color(s["f"]), stroke = HandleOverlay.color(s["k"])
            if fill != nil || stroke != nil {
                let p = (s["o"] != nil) ? NSBezierPath(roundedRect: r, xRadius: min(r.height / 2, 6), yRadius: min(r.height / 2, 6)) : NSBezierPath(rect: r)
                if let f = fill { f.setFill(); p.fill() }
                if let c = stroke { c.setStroke(); p.lineWidth = 0.75; p.stroke() }
            }
            let size = max(1, d(s, "z"))
            let bold = s["b"] != nil
            let font = NSFont(name: bold ? "Arial-BoldMT" : "ArialMT", size: size) ?? NSFont.systemFont(ofSize: size, weight: bold ? .bold : .regular)
            let para = NSMutableParagraphStyle()
            para.alignment = (s["a"] as? String) == "l" ? .left : (s["a"] as? String) == "r" ? .right : .center
            para.lineBreakMode = .byWordWrapping
            let text = NSAttributedString(string: (s["s"] as? String) ?? "", attributes: [.font: font, .foregroundColor: HandleOverlay.color(s["c"]) ?? NSColor.black, .paragraphStyle: para])
            let ctx = NSGraphicsContext.current?.cgContext
            let rot = CGFloat((s["r"] as? Int) ?? 0)
            if rot != 0, let ctx {
                // The box is the turned text's footprint: draw unturned in a swapped box about its centre.
                ctx.saveGState()
                ctx.translateBy(x: r.midX, y: r.midY)
                ctx.rotate(by: rot * .pi / 180)
                r = NSRect(x: -r.height / 2, y: -r.width / 2, width: r.height, height: r.width)
            }
            let h = text.boundingRect(with: NSSize(width: r.width + 1, height: 10_000), options: [.usesLineFragmentOrigin]).height
            let y = (s["v"] as? String) == "t" ? r.minY : (s["v"] as? String) == "b" ? r.maxY - h : r.midY - h / 2
            text.draw(with: NSRect(x: r.minX, y: y, width: r.width + 1, height: h + 1), options: [.usesLineFragmentOrigin])
            if rot != 0 { ctx?.restoreGState() }
        default: break
        }
    }
}

final class Helper: NSObject {
    var hosts: [String: Host] = [:]
    let powerPoint = "com.microsoft.Powerpoint"
    lazy var overlay: HandleOverlay = {
        let o = HandleOverlay(powerPoint: powerPoint)
        o.windows = { [weak self] in self?.powerPointWindows() ?? [] }
        // The mini toolbar steps aside while a handle is dragged (the reference suppresses it while
        // its privileged mouse handler is active).
        o.onDrag = { [weak self] on in
            guard let self, let t = self.hosts["toolbar"], t.open, !t.suspended else { return }
            if on { self.dragHidden = true; t.panel.orderOut(nil) } else { self.restoreToolbar() }
        }
        return o
    }()

    func host(_ name: String) -> Host {
        if let h = hosts[name] { return h }
        let h = Host(name: name, utility: name == "datasheet")
        hosts[name] = h
        return h
    }

    override init() {
        super.init()
        // Panels float over every app; keep them with PowerPoint: hide while another app is in front.
        NSWorkspace.shared.notificationCenter.addObserver(forName: NSWorkspace.didActivateApplicationNotification, object: nil, queue: .main) { [weak self] note in
            guard let self, let app = note.userInfo?[NSWorkspace.applicationUserInfoKey] as? NSRunningApplication else { return }
            let ours = app.processIdentifier == ProcessInfo.processInfo.processIdentifier
            self.overlay.appChanged(ours: ours || app.bundleIdentifier == self.powerPoint)
            for h in self.hosts.values where h.open {
                if app.bundleIdentifier == self.powerPoint || ours {
                    if h.suspended { h.suspended = false; h.panel.orderFrontRegardless() }
                } else if !h.suspended && h.pending == nil {
                    h.suspended = true
                    h.panel.orderOut(nil)
                }
            }
        }
    }

    /// The toolbar steps aside while the user drags (moving or resizing shapes): hidden on the first
    /// drag event, back after the mouse is released unless the server has moved or hidden it since.
    var dragHidden = false
    var dragTimer: Timer?
    func watchDrags() {
        NSEvent.addGlobalMonitorForEvents(matching: [.leftMouseDragged]) { [weak self] _ in
            guard let self, let t = self.hosts["toolbar"], t.open, !t.suspended, t.panel.isVisible else { return }
            self.dragHidden = true
            t.panel.orderOut(nil)
        }
        NSEvent.addGlobalMonitorForEvents(matching: [.leftMouseUp]) { [weak self] _ in self?.restoreToolbar() }
    }

    func restoreToolbar() {
        guard dragHidden else { return }
        dragTimer?.invalidate()
        dragTimer = Timer.scheduledTimer(withTimeInterval: 0.6, repeats: false) { [weak self] _ in
            guard let self, self.dragHidden else { return }
            self.dragHidden = false
            if let t = self.hosts["toolbar"], t.open, !t.suspended { t.panel.orderFrontRegardless() }
        }
    }

    /// PowerPoint's document windows (bounds, top-left global points), front-most first.
    func powerPointWindows() -> [[Double]] {
        var windows: [[Double]] = []
        // Bounds and owner of on-screen windows need no Screen Recording permission (names would).
        if let list = CGWindowListCopyWindowInfo([.optionOnScreenOnly, .excludeDesktopElements], kCGNullWindowID) as? [[String: Any]] {
            let pids = Set(NSRunningApplication.runningApplications(withBundleIdentifier: powerPoint).map { $0.processIdentifier })
            for w in list {
                guard let pid = w[kCGWindowOwnerPID as String] as? Int32, pids.contains(pid),
                      (w[kCGWindowLayer as String] as? Int) == 0,
                      let b = w[kCGWindowBounds as String] as? [String: Double],
                      let x = b["X"], let y = b["Y"], let width = b["Width"], let height = b["Height"],
                      width > 200, height > 150 else { continue }
                windows.append([x, y, width, height])
            }
        }
        return windows
    }

    func info() -> [String: Any] {
        let screens = NSScreen.screens.map { s in ["frame": topLeft(s.frame), "visible": topLeft(s.visibleFrame), "scale": Double(s.backingScaleFactor)] as [String: Any] }
        let windows = powerPointWindows()
        var saved: [String: [Double]] = [:]
        for name in ["datasheet", "toolbar"] {
            if let f = UserDefaults.standard.array(forKey: "frame.\(name)") as? [Double], f.count == 4 { saved[name] = f }
        }
        return ["event": "info", "screens": screens, "windows": windows, "saved": saved]
    }

    func handle(_ line: String) {
        guard let data = line.data(using: .utf8), let m = (try? JSONSerialization.jsonObject(with: data)) as? [String: Any], let cmd = m["cmd"] as? String else {
            emit(["event": "error", "message": "bad request"])
            return
        }
        let name = (m["panel"] as? String) ?? "datasheet"
        if name == "toolbar" && (cmd == "show" || cmd == "hide") { dragHidden = false; dragTimer?.invalidate() }
        switch cmd {
        case "handles", "ghost", "handlesState":
            overlay.command(m)
        case "load":
            if let url = m["url"] as? String { host(name).load(url) }
        case "info":
            emit(info())
        case "show":
            guard let id = m["id"], JSONSerialization.isValidJSONObject([id]), let url = m["url"] as? String, let frame = m["frame"] as? [Double], frame.count == 4 else {
                emit(["event": "error", "message": "show needs id, url and frame"])
                return
            }
            host(name).show(id: id, url: url, frame: frame, focus: (m["focus"] as? Bool) ?? true)
        case "hide":
            host(name).hide(reason: "server")
        case "eval":
            host(name).js((m["js"] as? String) ?? "") { r, e in
                var out: [String: Any] = ["event": "eval"]
                if let e { out["error"] = e.localizedDescription } else if let r, JSONSerialization.isValidJSONObject([r]) { out["result"] = r }
                emit(out)
            }
        case "quit":
            exit(0)
        default:
            emit(["event": "error", "message": "unknown command \(cmd)"])
        }
    }
}

let app = NSApplication.shared
app.setActivationPolicy(.accessory)
let helper = Helper()
helper.watchDrags()

// stdin: one JSON request per line; EOF (the server is gone) ends the helper.
var pendingInput = Data()
FileHandle.standardInput.readabilityHandler = { handle in
    let chunk = handle.availableData
    if chunk.isEmpty { DispatchQueue.main.async { exit(0) } ; handle.readabilityHandler = nil; return }
    pendingInput.append(chunk)
    while let nl = pendingInput.firstIndex(of: 0x0A) {
        let lineData = pendingInput.subdata(in: pendingInput.startIndex..<nl)
        pendingInput.removeSubrange(pendingInput.startIndex...nl)
        if let line = String(data: lineData, encoding: .utf8), !line.isEmpty {
            DispatchQueue.main.async { helper.handle(line) }
        }
    }
}
emit(["event": "started", "pid": Int(ProcessInfo.processInfo.processIdentifier)])
app.run()
