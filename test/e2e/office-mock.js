// A recording stand-in for office.js (Office + PowerPoint JS), served in place of the CDN script by
// test/e2e/officejs.spec.ts. It models one presentation (slides → shapes, groups, tags, selection)
// closely enough to exercise the task pane's contract: which API calls it makes, in how many syncs,
// and what ends up on the slide. Everything executes immediately; `sync()` only counts (and fails
// when a queued read was invalid, like PowerPoint's ItemNotFound for a shape with no parent group).
// Configure before the page loads: window.__OC_MOCK_CONFIG = { maxSet: '1.10', slides: [...], exportBase64 }.
(() => {
  const cfg = window.__OC_MOCK_CONFIG ?? {};
  const calls = (window.__OC_CALLS = []);
  const call = (s) => calls.push(s);
  let nextId = 100;
  const ver = (v) => v.split('.').map(Number);
  const atMost = (a, b) => { const [x1, y1] = ver(a), [x2, y2] = ver(b); return x1 < x2 || (x1 === x2 && y1 <= y2); };

  function tagsOf(init = {}) {
    const items = Object.entries(init).map(([key, value]) => ({ key: key.toUpperCase(), value }));
    return {
      items,
      add(k, v) { call(`tags.add ${k.toUpperCase()}`); const t = items.find((x) => x.key === k.toUpperCase()); if (t) t.value = v; else items.push({ key: k.toUpperCase(), value: v }); },
      getItemOrNullObject(k) { const t = items.find((x) => x.key === k.toUpperCase()); return { isNullObject: !t, value: t?.value, load() { return this; } }; },
      load() { return this; },
    };
  }

  function makeShape(slide, props) {
    const s = {
      id: String(nextId++), name: '', type: 'GeometricShape', left: 0, top: 0, width: 0, height: 0, rotation: 0,
      slide, parent: null, children: null, geom: null, text: null,
      fill: { color: null, setSolidColor(c) { this.color = c; }, clear() { this.color = null; } },
      lineFormat: { visible: true, color: null, weight: null, dashStyle: null },
      adjustments: { values: [], set(i, v) { this.values[i] = v; } },
      load() { return this; },
      ...props,
    };
    s.tags = tagsOf(props.tagInit);
    s.textFrame = { wordWrap: true, autoSizeSetting: null, verticalAlignment: null, textRange: { text: props.text ?? '', font: {}, paragraphFormat: {} } };
    Object.defineProperty(s, 'zOrderPosition', { get() { return slide.shapes.list.indexOf(s); } });
    Object.defineProperty(s, 'parentGroup', {
      get() {
        if (s.parent) return s.parent;
        return { tags: { items: [], load() {} }, load() { slide.pres.failNext = 'ItemNotFound: the shape is not in a group'; return this; } };
      },
    });
    s.delete = () => { call(`delete ${s.id}`); const l = (s.parent?.children ?? slide.shapes.list); l.splice(l.indexOf(s), 1); };
    s.setZOrder = (op) => {
      call(`setZOrder ${op}`);
      const l = slide.shapes.list, i = l.indexOf(s);
      if (op === 'SendBackward' && i > 0) { l.splice(i, 1); l.splice(i - 1, 0, s); }
    };
    if (props.children) for (const c of props.children) c.parent = s;
    return s;
  }

  function makeSlide(pres, def) {
    const slide = { id: def.id, pres };
    const list = [];
    slide.shapes = {
      list,
      get items() { return list; },
      load() { return this; },
      getItem(id) { const s = findShape(slide, id); if (!s) throw new Error(`no shape ${id}`); return s; },
      addGeometricShape(geom, o) { call(`addGeometricShape ${geom}`); const s = makeShape(slide, { geom, ...o }); list.push(s); return s; },
      addLine(type, o) { call(`addLine ${type}`); const s = makeShape(slide, { type: 'Line', geom: 'line', ...o }); list.push(s); return s; },
      addTextBox(text, o) { call('addTextBox'); const s = makeShape(slide, { type: 'TextBox', text, ...o }); list.push(s); return s; },
      addGroup(members) {
        call(`addGroup ${members.length}`);
        const kids = members.map((m) => (typeof m === 'string' ? findShape(slide, m) : m));
        for (const k of kids) list.splice(list.indexOf(k), 1);
        const x0 = Math.min(...kids.map((k) => k.left)), y0 = Math.min(...kids.map((k) => k.top));
        const x1 = Math.max(...kids.map((k) => k.left + k.width)), y1 = Math.max(...kids.map((k) => k.top + k.height));
        const g = makeShape(slide, { type: 'Group', left: x0, top: y0, width: x1 - x0, height: y1 - y0, children: kids });
        list.push(g);
        return g;
      },
    };
    slide.setSelectedShapes = (ids) => { call(`setSelectedShapes ${ids.join(',')}`); pres.selection = { slideIds: [slide.id], shapeIds: ids }; };
    slide.exportAsBase64 = () => { call('exportAsBase64'); window.__OC_EXPORTED_NAMES = list.map((s) => s.name); return { value: cfg.exportBase64 ?? '' }; };
    slide.themeColorScheme = { getThemeColor(n) { call(`getThemeColor ${n}`); return { value: (cfg.theme ?? {})[n] ?? '#000000' }; } };
    slide.delete = () => { call(`deleteSlide ${slide.id}`); pres.slideList.splice(pres.slideList.indexOf(slide), 1); };
    slide.load = () => slide;
    for (const sd of def.shapes ?? []) list.push(buildShape(slide, sd));
    return slide;
  }

  function buildShape(slide, sd) {
    const children = sd.children?.map((c) => buildShape(slide, c));
    return makeShape(slide, { ...sd, tagInit: sd.tags, children, type: sd.type ?? (children ? 'Group' : 'GeometricShape') });
  }

  function findShape(slide, id) {
    const walk = (l) => { for (const s of l) { if (s.id === id) return s; if (s.children) { const f = walk(s.children); if (f) return f; } } };
    return walk(slide.shapes.list);
  }

  const pres = { slideList: [], selection: { slideIds: [], shapeIds: [] }, failNext: null };
  window.__OC_PRES = pres;
  for (const def of cfg.slides ?? [{ id: '256#', shapes: [] }]) pres.slideList.push(makeSlide(pres, def));
  pres.selection.slideIds = [pres.slideList[0].id];
  const collection = (items) => ({ items, getItemAt: (i) => items[i], load() { return this; } });
  const slideById = (id) => pres.slideList.find((s) => s.id === id);

  const presentation = {
    getSelectedSlides() { return collection(pres.selection.slideIds.map(slideById).filter(Boolean)); },
    getSelectedShapes() {
      const slide = slideById(pres.selection.slideIds[0]);
      return collection(slide ? pres.selection.shapeIds.map((id) => findShape(slide, id)).filter(Boolean) : []);
    },
    get slides() {
      return { get items() { return pres.slideList; }, load() { return this; }, getItem: (id) => slideById(id) ?? (() => { throw new Error(`no slide ${id}`); })() };
    },
    setSelectedSlides(ids) { call(`setSelectedSlides ${ids.join(',')}`); pres.selection = { slideIds: ids, shapeIds: [] }; },
    insertSlidesFromBase64(b64, opts) {
      call(`insertSlidesFromBase64 ${opts?.formatting} ${opts?.targetSlideId ?? ''}`);
      (window.__OC_INSERTED ??= []).push(b64);
      const at = opts?.targetSlideId ? pres.slideList.indexOf(slideById(opts.targetSlideId)) + 1 : pres.slideList.length;
      pres.slideList.splice(at, 0, makeSlide(pres, { id: `${300 + pres.slideList.length}#`, shapes: [] }));
    },
  };

  let syncs = 0;
  window.__OC_SYNCS = () => syncs;
  window.PowerPoint = {
    async run(f) {
      const ctx = {
        presentation,
        async sync() {
          syncs++; call('sync');
          if (pres.failNext) { const m = pres.failNext; pres.failNext = null; throw new Error(m); }
        },
      };
      return f(ctx);
    },
  };

  const handlers = [];
  window.Office = {
    HostType: { PowerPoint: 'PowerPoint' },
    EventType: { DocumentSelectionChanged: 'documentSelectionChanged' },
    context: {
      requirements: { isSetSupported: (name, v) => name === 'PowerPointApi' && atMost(v ?? '1.1', cfg.maxSet ?? '1.10') },
      document: { addHandlerAsync: (type, h) => { call(`addHandlerAsync ${type}`); handlers.push(h); } },
    },
    onReady(cb) { setTimeout(() => cb({ host: 'PowerPoint', platform: 'Mac' }), 0); return Promise.resolve({ host: 'PowerPoint' }); },
  };
  // Test hooks: what the user does in PowerPoint.
  window.__OC_MOCK = {
    select(ids, slideId) { pres.selection = { slideIds: [slideId ?? pres.selection.slideIds[0]], shapeIds: ids }; handlers.forEach((h) => h({})); },
    shapes(slideId) { return slideById(slideId ?? pres.selection.slideIds[0])?.shapes.list ?? []; },
    selected() { return pres.selection; },
  };
})();
