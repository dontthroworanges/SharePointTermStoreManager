namespace SharePointTermStoreManager.Services;

/// <summary>
/// The page the HTML export fills in: {{TITLE}} and {{DATA}} (the JSON written by
/// <see cref="ExportWriter"/>). Self-contained, with no outside references, so it opens from anywhere, offline.
/// </summary>
internal static class HtmlTemplate
{
    public const string Page = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{TITLE}}</title>
<style>
  :root {
    color-scheme: light dark;
    --bg: #f3f3f3; --card: #ffffff; --text: #1b1b1b; --muted: #616161; --line: #e1e1e1;
    --hover: rgba(0, 0, 0, .05); --sel: #dcebf8; --accent: #0f6cbd; --mark: #fff0a0;
    --g: #b88a00; --s: #0f6cbd; --t: #5c5c5c; --dep: #9a9a9a; --chip: #f0f0f0;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #202020; --card: #2b2b2b; --text: #f3f3f3; --muted: #b0b0b0; --line: #3c3c3c;
      --hover: rgba(255, 255, 255, .06); --sel: #24425e; --accent: #4cc2ff; --mark: #6b5800;
      --g: #ffd335; --s: #4cc2ff; --t: #c8c8c8; --dep: #8a8a8a; --chip: #383838;
    }
  }
  * { box-sizing: border-box; }
  html, body { height: 100%; }
  body { margin: 0; display: flex; flex-direction: column; font: 14px/1.45 "Segoe UI", system-ui, -apple-system, sans-serif; background: var(--bg); color: var(--text); }
  header { padding: 16px 24px 4px; }
  h1 { font-size: 20px; font-weight: 600; margin: 0; }
  .meta { color: var(--muted); font-size: 12px; margin-top: 2px; }
  .toolbar { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; padding: 10px 24px 12px; }
  .toolbar input { flex: 0 1 380px; min-width: 200px; padding: 6px 10px; font: inherit; color: var(--text); background: var(--card); border: 1px solid var(--line); border-bottom: 2px solid var(--accent); border-radius: 4px; outline: none; }
  .tb { padding: 5px 12px; font: inherit; color: var(--text); background: var(--card); border: 1px solid var(--line); border-radius: 4px; cursor: pointer; }
  .tb:hover { background: var(--hover); }
  #count { color: var(--muted); font-size: 12px; }
  main { flex: 1; min-height: 0; display: grid; grid-template-columns: minmax(280px, 2fr) 3fr; gap: 12px; padding: 0 24px 24px; }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 8px; overflow: auto; min-height: 0; }
  #tree { padding: 8px 6px; }
  #tree ul { list-style: none; margin: 0; padding-left: 20px; }
  #tree > ul { padding-left: 0; }
  #tree li:not(.open) > ul { display: none; }
  #tree.filtering li:not(.vis) { display: none; }
  .row { display: flex; align-items: center; min-width: 0; }
  .tw { flex: none; width: 22px; height: 24px; padding: 0; border: 0; border-radius: 4px; background: none; color: var(--muted); cursor: pointer; }
  .tw::before { content: "\203A"; display: inline-block; font-size: 18px; line-height: 1; transition: transform .12s; }
  li.open > .row > .tw::before { transform: rotate(90deg); }
  .tw.leaf { visibility: hidden; }
  .tw:hover, .nm:hover { background: var(--hover); }
  .nm { display: flex; align-items: center; gap: 6px; min-width: 0; padding: 2px 8px 2px 4px; border: 0; border-radius: 4px; background: none; color: inherit; font: inherit; text-align: left; cursor: pointer; }
  .nm .l { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .nm .c { color: var(--muted); font-size: 12px; }
  .nm.sel { background: var(--sel); }
  .nm.dep .l { color: var(--dep); text-decoration: line-through; }
  .nm:focus-visible, .tb:focus-visible, .link:focus-visible { outline: 2px solid var(--accent); outline-offset: -2px; }
  .ic { flex: none; width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.2; stroke-linejoin: round; }
  .ic.g { color: var(--g); } .ic.s { color: var(--s); } .ic.t { color: var(--t); }
  mark { background: var(--mark); color: inherit; border-radius: 2px; }
  #details { padding: 20px 24px; }
  #details .empty { color: var(--muted); margin-top: 40px; text-align: center; }
  .crumbs { display: flex; flex-wrap: wrap; align-items: center; gap: 2px; color: var(--muted); font-size: 13px; }
  .crumbs .sep { margin: 0 2px; }
  .link { padding: 1px 4px; border: 0; border-radius: 4px; background: none; color: var(--accent); font: inherit; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; text-align: left; }
  .link:hover { background: var(--hover); text-decoration: underline; }
  h2 { display: flex; align-items: center; gap: 8px; font-size: 22px; font-weight: 600; margin: 6px 0 0; word-break: break-word; }
  h2 .ic { width: 22px; height: 22px; }
  h3 { font-size: 14px; font-weight: 600; margin: 20px 0 6px; }
  .kind { color: var(--muted); }
  .flags { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }
  .flag, .chip { padding: 2px 10px; border-radius: 12px; background: var(--chip); font-size: 12px; }
  .chips { display: flex; flex-wrap: wrap; gap: 6px; }
  dl { display: grid; grid-template-columns: max-content 1fr; gap: 8px 20px; margin: 18px 0 0; }
  dt { color: var(--muted); }
  dd { margin: 0; white-space: pre-line; word-break: break-word; }
  .mono { font-family: Consolas, "Cascadia Mono", monospace; user-select: all; }
  table { border-collapse: collapse; }
  td { padding: 2px 16px 2px 0; vertical-align: top; }
  td:first-child { color: var(--muted); }
  ul.children { list-style: none; margin: 0; padding: 0; columns: 2 220px; }
  ul.children li { break-inside: avoid; }
  .more { color: var(--muted); padding: 2px 4px; }
  @media (max-width: 760px) {
    body { height: auto; }
    main { grid-template-columns: 1fr; }
    .card { max-height: 70vh; }
  }
</style>
</head>
<body>
<svg width="0" height="0" style="position:absolute" aria-hidden="true">
  <symbol id="i-g" viewBox="0 0 16 16"><path d="M1.5 4.5a1 1 0 0 1 1-1h3.6l1.5 1.5h5.9a1 1 0 0 1 1 1v6.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1z"/></symbol>
  <symbol id="i-s" viewBox="0 0 16 16"><path d="M1.5 2.5h4.8l6.2 6.2-4.6 4.6-6.4-6.2z"/><path d="M8.3 2.5h1.2l5 5-4.2 4.2"/><circle cx="4.3" cy="5.3" r=".9"/></symbol>
  <symbol id="i-t" viewBox="0 0 16 16"><path d="M2 2.5h5.3l6.7 6.7-4.8 4.8L2.5 7.3z"/><circle cx="5" cy="5.5" r="1"/></symbol>
</svg>
<header>
  <h1 id="title">Term store</h1>
  <div class="meta" id="meta"></div>
</header>
<div class="toolbar">
  <input id="q" type="search" placeholder="Search names, synonyms, descriptions and IDs" aria-label="Search the term store">
  <button class="tb" type="button" id="expand">Expand all</button>
  <button class="tb" type="button" id="collapse">Collapse all</button>
  <span id="count" role="status"></span>
</div>
<main>
  <nav class="card" id="tree" aria-label="Term store"><ul></ul></nav>
  <section class="card" id="details" aria-live="polite"><p class="empty">Select an item to see its details.</p></section>
</main>
<script type="application/json" id="data">{{DATA}}</script>
<script>
(() => {
  const D = JSON.parse(document.getElementById('data').textContent);
  const N = D.nodes;
  const kids = N.map(() => []);
  const roots = [];
  N.forEach((x, i) => (x.p >= 0 ? kids[x.p] : roots).push(i));
  const TYPE = { g: 'Term group', s: 'Term set', t: 'Term' };
  const $ = s => document.querySelector(s);
  const items = new Map();            // rendered tree items, by node index
  let selected = -1, query = '', hits = null, keep = null;

  const count = t => N.reduce((n, x) => n + (x.t === t ? 1 : 0), 0);
  const plural = (n, one, many) => `${n.toLocaleString()} ${n === 1 ? one : many}`;
  const summary = [plural(count('g'), 'term group', 'term groups'), plural(count('s'), 'term set', 'term sets'), plural(count('t'), 'term', 'terms')].join(' · ');
  let host = D.site; try { host = new URL(D.site).host; } catch (e) { }
  $('#title').textContent = `Term store: ${host}`;
  $('#meta').textContent = `${D.site} · exported ${D.exported} · ${D.language} · ${summary}`;
  $('#count').textContent = summary;

  function h(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }
  function icon(t) {
    const ns = 'http://www.w3.org/2000/svg', s = document.createElementNS(ns, 'svg'), u = document.createElementNS(ns, 'use');
    s.setAttribute('class', 'ic ' + t);
    s.setAttribute('viewBox', '0 0 16 16');
    s.setAttribute('aria-hidden', 'true');
    u.setAttribute('href', '#i-' + t);
    s.append(u);
    return s;
  }
  // The name, with the search text highlighted.
  function nameInto(el, name) {
    el.textContent = '';
    if (!query) { el.textContent = name; return; }
    const lower = name.toLowerCase();
    let at = 0, i;
    while ((i = lower.indexOf(query, at)) >= 0) {
      el.append(name.slice(at, i), h('mark', null, name.slice(i, i + query.length)));
      at = i + query.length;
    }
    el.append(name.slice(at));
  }
  function hitAbove(k) {
    for (let p = N[k].p; p >= 0; p = N[p].p) if (hits.has(p)) return true;
    return false;
  }
  // While searching, matches, the items above them and everything below a match stay visible.
  const visible = k => !hits || keep.has(k) || hitAbove(k);

  function makeItem(i) {
    const x = N[i], li = h('li'), row = h('div', 'row');
    const tw = h('button', 'tw');
    tw.type = 'button';
    tw.tabIndex = -1;
    if (kids[i].length) { tw.title = 'Expand or collapse'; tw.onclick = () => toggle(i); } else tw.classList.add('leaf');
    const nm = h('button', 'nm' + (x.f && x.f.includes('Deprecated') ? ' dep' : '') + (i === selected ? ' sel' : ''));
    nm.type = 'button';
    const label = h('span', 'l');
    nameInto(label, x.n);
    nm.append(icon(x.t), label);
    if (kids[i].length) nm.append(h('span', 'c', kids[i].length.toLocaleString()));
    nm.onclick = () => select(i, false);
    nm.ondblclick = () => toggle(i);
    nm.onkeydown = e => keys(e, i);
    row.append(tw, nm);
    li.append(row);
    if (visible(i)) li.classList.add('vis');
    items.set(i, li);
    return li;
  }
  function expand(i, open) {
    const li = items.get(i);
    if (!li || !kids[i].length) return;
    if (open && !li.querySelector(':scope > ul')) {
      const ul = h('ul');
      kids[i].forEach(c => ul.append(makeItem(c)));
      li.append(ul);
    }
    li.classList.toggle('open', open);
  }
  const toggle = i => expand(i, !items.get(i).classList.contains('open'));
  function ancestors(i) {
    const a = [];
    for (let p = N[i].p; p >= 0; p = N[p].p) a.unshift(p);
    return a;
  }
  const reveal = i => ancestors(i).forEach(a => expand(a, true));
  const button = i => items.get(i).querySelector(':scope > .row > .nm');

  function select(i, scroll) {
    if (selected >= 0 && items.has(selected)) button(selected).classList.remove('sel');
    selected = i;
    reveal(i);
    const b = button(i);
    b.classList.add('sel');
    if (scroll) { b.scrollIntoView({ block: 'center' }); b.focus({ preventScroll: true }); }
    history.replaceState(null, '', '#' + N[i].id);
    show(i);
  }

  // Arrow keys move through the tree; right and left open and close.
  function keys(e, i) {
    if (e.key === 'ArrowRight') {
      if (kids[i].length && !items.get(i).classList.contains('open')) expand(i, true);
      else if (kids[i].length) select(kids[i].find(visible) ?? kids[i][0], true);
    } else if (e.key === 'ArrowLeft') {
      if (items.get(i).classList.contains('open')) expand(i, false);
      else if (N[i].p >= 0) select(N[i].p, true);
    } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      const all = [...document.querySelectorAll('#tree .nm')].filter(b => b.offsetParent !== null);
      const next = all[all.indexOf(e.currentTarget) + (e.key === 'ArrowDown' ? 1 : -1)];
      if (next) select(+next.closest('li').dataset.k, true);
    } else return;
    e.preventDefault();
  }

  function show(i) {
    const x = N[i], d = $('#details');
    d.textContent = '';
    const crumbs = h('nav', 'crumbs');
    crumbs.setAttribute('aria-label', 'Path');
    ancestors(i).forEach(a => {
      const b = h('button', 'link', N[a].n);
      b.type = 'button';
      b.onclick = () => select(a, true);
      crumbs.append(b, h('span', 'sep', '›'));
    });
    const title = h('h2');
    title.append(icon(x.t), h('span', null, x.n));
    d.append(crumbs, title, h('div', 'kind', TYPE[x.t]));
    if (x.f) {
      const f = h('div', 'flags');
      x.f.forEach(t => f.append(h('span', 'flag', t)));
      d.append(f);
    }
    const dl = h('dl');
    const add = (k, v) => {
      if (v == null || v === '') return;
      const dd = h('dd');
      if (v instanceof Node) dd.append(v); else dd.textContent = v;
      dl.append(h('dt', null, k), dd);
    };
    const table = o => {
      const t = h('table');
      Object.entries(o).forEach(([k, v]) => { const tr = h('tr'); tr.append(h('td', null, k), h('td', null, v)); t.append(tr); });
      return t;
    };
    add('Description', x.d);
    if (x.sy) { const c = h('div', 'chips'); x.sy.forEach(s => c.append(h('span', 'chip', s))); add('Synonyms', c); }
    if (x.ol) add('Other languages', x.ol.join('\n'));
    add('Owner', x.ow);
    add('Contact', x.co);
    if (x.cp) add(x.t === 't' ? 'Shared custom properties' : 'Custom properties', table(x.cp));
    if (x.lp) add('Local custom properties', table(x.lp));
    add('ID', h('span', 'mono', x.id));
    add('Created', x.cr);
    add('Last modified', x.lm);
    d.append(dl);

    if (kids[i].length) {
      d.append(h('h3', null, `${x.t === 'g' ? 'Term sets' : 'Terms'} (${kids[i].length.toLocaleString()})`));
      const ul = h('ul', 'children');
      kids[i].slice(0, 500).forEach(c => {
        const li = h('li'), b = h('button', 'link');
        b.type = 'button';
        b.append(icon(N[c].t), h('span', null, N[c].n));
        b.onclick = () => select(c, true);
        li.append(b);
        ul.append(li);
      });
      if (kids[i].length > 500) ul.append(h('li', 'more', `…and ${(kids[i].length - 500).toLocaleString()} more`));
      d.append(ul);
    }
    d.scrollTop = 0;
  }

  // Search: names, synonyms, descriptions and IDs. The first 300 matches are opened in the tree.
  function search(text) {
    query = text.trim().toLowerCase();
    const tree = $('#tree');
    if (!query) {
      hits = keep = null;
      tree.classList.remove('filtering');
      $('#count').textContent = summary;
    } else {
      hits = new Set();
      N.forEach((x, i) => {
        if (x.n.toLowerCase().includes(query) || x.id.includes(query) ||
            (x.sy && x.sy.some(s => s.toLowerCase().includes(query))) ||
            (x.d && x.d.toLowerCase().includes(query))) hits.add(i);
      });
      keep = new Set();
      hits.forEach(i => {
        keep.add(i);
        for (let p = N[i].p; p >= 0 && !keep.has(p); p = N[p].p) keep.add(p);
      });
      [...hits].slice(0, 300).forEach(reveal);
      tree.classList.add('filtering');
      $('#count').textContent = hits.size === 0 ? 'No matches'
        : plural(hits.size, 'match', 'matches') + (hits.size > 300 ? ' (the first 300 are opened)' : '');
    }
    items.forEach((li, k) => {
      li.classList.toggle('vis', visible(k));
      nameInto(li.querySelector(':scope > .row > .nm > .l'), N[k].n);
    });
  }
  let timer;
  $('#q').addEventListener('input', e => { clearTimeout(timer); timer = setTimeout(() => search(e.target.value), 200); });
  $('#expand').onclick = () => {
    // Opening a level draws the next, so walk the whole tree.
    const stack = [...roots];
    while (stack.length) { const i = stack.pop(); if (kids[i].length) { expand(i, true); stack.push(...kids[i]); } }
  };
  $('#collapse').onclick = () => items.forEach(li => li.classList.remove('open'));

  const list = $('#tree > ul');
  roots.forEach(i => list.append(makeItem(i)));
  roots.forEach(i => expand(i, true)); // show each group's term sets

  // A link to an item (#its-id) opens at that item.
  const fromHash = decodeURIComponent(location.hash.slice(1)).toLowerCase();
  const start = fromHash ? N.findIndex(x => x.id === fromHash) : -1;
  if (start >= 0) select(start, true);
})();
</script>
</body>
</html>
""";
}
