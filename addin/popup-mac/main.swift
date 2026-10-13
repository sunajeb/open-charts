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

final class Helper: NSObject {
    var hosts: [String: Host] = [:]
    let powerPoint = "com.microsoft.Powerpoint"

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
        NSEvent.addGlobalMonitorForEvents(matching: [.leftMouseUp]) { [weak self] _ in
            guard let self, self.dragHidden else { return }
            self.dragTimer?.invalidate()
            self.dragTimer = Timer.scheduledTimer(withTimeInterval: 0.6, repeats: false) { [weak self] _ in
                guard let self, self.dragHidden else { return }
                self.dragHidden = false
                if let t = self.hosts["toolbar"], t.open, !t.suspended { t.panel.orderFrontRegardless() }
            }
        }
    }

    func info() -> [String: Any] {
        let screens = NSScreen.screens.map { s in ["frame": topLeft(s.frame), "visible": topLeft(s.visibleFrame), "scale": Double(s.backingScaleFactor)] as [String: Any] }
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
