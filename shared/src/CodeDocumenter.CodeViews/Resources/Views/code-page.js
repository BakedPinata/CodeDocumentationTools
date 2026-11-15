// code-page.js: a page of code, for the views in this library: its data, the selection, what the source panel shows,
// the explorer, Back and Forward, and search. The frame around it (theme, tabs, side panel, Load folder, source text)
// is the shell's (shell.js in CodeDocumenter.Shared).
//
// The data is the code model (CodeModel in this library): files, groups, units and members, the edges between them,
// and the language's description of itself (data.language: its kinds of unit, member and relationship, each with a
// label and how it's drawn, and its words for things). Nothing here or in the views names a language's kinds: what
// a unit is called, what kinds there are and how each looks all come from that description.
//
// The tool says what's its own before this script runs, in window.codeViewsPage:
//   tool           the tool's short name, such as "vbadoc": its data block, its stored settings, and a global of its
//                  own for the page (window.vbadoc), the same object as window.codeViews
//   sourceWord     what the page documents, in a word, for the shell's messages: "workbook", "solution"
//   source(data)   how its source text is read: { decode(buffer), note, highlight(lines), files }, where files
//                  completes "Choose the folder that contains this page and its …" (".bas, .cls and .frm files")
//   exports(page)  for a tool with the shell's Export tab, what shell.exports.init takes, worked out from page:
//                  { data, unitsById, unitEdges, unitOfId(id) }. Left out, there's no Export tab to fill
//
// The views (dependency graph, class diagram, Calls) talk to it through window.codeViews:
//   page.tool / data / language        the tool's name, the code model and its language's description
//   page.unitsById / membersById       Maps of the model's units and members
//   page.unitKind(unit)                the language's description of a unit's kind
//   page.unitKindLabel(unit) / memberKindLabel(kind) / roleLabel(role)   what to call them
//   page.memberKind(kind)              the language's description of a kind of member
//   page.memberCategories              [{ id, label, shape, under, kinds }]: the kinds of member the Show… menu chooses between
//   page.relationship(id)              the language's description of a kind of edge
//   page.capital(text)                 text with a capital first letter, for a word of the language's at the start
//   page.select(id, other)             select a unit or member by id, or null; other is true for a Ctrl+click, which
//                                      does the other thing from focus on click
//   page.selection()                   the selected id, or null
//   page.onSelectionChange(fn)         fn(id) after the selection changes
//   page.focused() / onFocusChange(fn) whether the views show only the selection and what it's connected to; fn
//                                      when that changes and the selection doesn't
//   page.focusOnClick() / setFocusOnClick(on)  the page's Focus on click setting (the views' Focus buttons)
//   page.theme() / onThemeChange(fn)   "light" or "dark"
//   page.activeView() / onViewChange(fn)  "graph", "classes" or "calls"
//   page.favourites() / isFavourite(id) / toggleFavourite(id) / onFavouritesChange(fn)  the favourite units and members (the shell's)
//   page.inFilter(id) / filterActive() / filterKey() / onFilterChange(fn)  whether the explorer's filter shows a unit
//                                      or member (a ticked unit brings its members, a ticked member its unit)
//   page.hidden() / isHidden(id) / toggleHidden(id)   the items hidden from the views (a hidden unit takes its
//                                      members with it)
//   page.shows(id) / shownKey() / onShownChange(fn)   whether the views show a unit or member: in the filter and
//                                      not hidden. The key changes with either
//   page.showsItself(id)               whether the graph's member level shows a unit's own node: as shows(id), but
//                                      while the view is locked, only a unit ticked or kept on its own
//   page.locked() / onLockChange(fn) / lockSource(view, fn)   the shell's Lock view: whether it's on; fn(on) when
//                                      it changes, just before onShownChange; and a view's say in what locking ticks:
//                                      fn() gives { ids, alone } for what it shows now, when it's the open view
//   page.kept(id)                      whether the Show… choices keep a unit or member (its unit's kind, its own
//                                      kind, private members, references); the views apply the rest themselves
//   page.choice(id) / choiceValue(name) / choicesKey() / choicesCount() / onChoicesChange(fn)   the explorer's Show…
//                                      choices: "unit:<kind>" and "member:<category>" for kinds, "line:<relationship>"
//                                      for lines (and any of the class diagram's own), and "unconnected" for items
//                                      with no connections
//   page.focusFavourites() / setFocusFavourites(on)   whether the views show only the favourites and what they use
//                                      (with the focused selection's); changes come through onFocusChange
//   page.preparePrint(which, options) / finishPrint() / viewText(options)   the shell's print document and copied
//                                      text, made without printing or copying: for tests
//   page.viewPart(name, part)          a view's part of every saved view (the shell's views.part): its save(),
//                                      apply(value) and ids(value). Opening a saved view sets each part and the
//                                      selection, then the views hear of it through onShownChange, once, and read
//                                      the selection and its focus as they draw (onSelectionChange isn't called)
(function () {
  "use strict";

  var config = window.codeViewsPage;
  var shell = window.CodeDocumenterShell.create({ tool: config.tool, panel: "source", rememberView: true });
  var data = shell.data;
  var language = data.language;
  var unitsById = new Map();
  var membersById = new Map();
  var filesById = new Map();
  var unfollowedByMember = new Map();
  data.units.forEach(function (u) { unitsById.set(u.id, u); });
  data.members.forEach(function (m) { membersById.set(m.id, m); });
  data.files.forEach(function (f) { filesById.set(f.id, f); });
  data.unfollowed.forEach(function (l) {
    if (!unfollowedByMember.has(l.from)) unfollowedByMember.set(l.from, []);
    unfollowedByMember.get(l.from).push(l);
  });

  // ---- The language ----
  //
  // Its kinds by id, and its words. A kind the description doesn't list is shown by its id, so a mistake in a
  // language shows up as an odd label, not a broken page.

  function byId(list) {
    var map = {};
    (list || []).forEach(function (item) { map[item.id] = item; });
    return map;
  }
  var unitKinds = byId(language.unitKinds);
  var memberKinds = byId(language.memberKinds);
  var roles = byId(language.roles);
  var relationships = byId(language.relationships);
  var UNKNOWN_UNIT_KIND = { id: "unknown", label: "", plural: "", shape: "circle", colour: 1, stereotype: "" };

  function capital(text) { return text.charAt(0).toUpperCase() + text.slice(1); }
  function unitKind(unit) { return unitKinds[unit.kind] || Object.assign({}, UNKNOWN_UNIT_KIND, { id: unit.kind, label: unit.kind, stereotype: unit.kind }); }
  function unitKindLabel(unit) { return unit.kindLabel || unitKind(unit).label; }
  function memberKindLabel(kind) { return memberKinds[kind] ? memberKinds[kind].label : kind; }
  function roleLabel(role) { return roles[role] ? roles[role].label : role; }
  function fileName(id) { var file = filesById.get(id); return file ? file.name : id; }

  // The kinds of member the Show… menu chooses between, each with the kinds in it.
  var MEMBER_CATEGORIES = language.memberCategories.map(function (c) {
    return {
      id: c.id, label: c.label, shape: c.shape, under: c.under || null,
      kinds: language.memberKinds.filter(function (k) { return k.category === c.id; }).map(function (k) { return k.id; })
    };
  });
  var categoryOfKind = {};
  MEMBER_CATEGORIES.forEach(function (c) { c.kinds.forEach(function (k) { categoryOfKind[k] = c.id; }); });

  // How each kind is drawn that the stylesheet can't know: a unit kind's colour on its class diagram box, and each
  // relationship's line, arrowhead and tag. Colours are slots of the views' palette (views.css), so they follow the theme.
  (function () {
    var DASHES = { solid: "none", dashed: "6 4", dotted: "2 3" };
    var rules = [];
    language.unitKinds.forEach(function (k) {
      rules.push(".class-box.kind-" + k.id + " .box-outline { stroke: var(--kind-" + k.colour + "); }");
    });
    language.relationships.forEach(function (r) {
      var colour = "var(--edge-" + r.colour + ")";
      rules.push(".graph-art .edge." + r.id + " { stroke: " + colour + "; stroke-dasharray: " + DASHES[r.line] + "; }");
      rules.push(r.arrow === "hollow"
        ? ".graph-art .arrow-" + r.id + " { fill: var(--bg); stroke: " + colour + "; stroke-width: 1.25px; }"
        : ".graph-art .arrow-" + r.id + " { fill: " + colour + "; }");
      if (!r.plain) rules.push(".edge-tag." + r.id + " { border-color: " + colour + "; border-style: " + (r.line === "solid" ? "solid" : r.line) + "; }");
    });
    language.classRelationships.forEach(function (r) {
      var colour = r.colour > 0 ? "var(--edge-" + r.colour + ")" : "var(--text-muted)";
      rules.push(".class-edge." + r.id + " { stroke: " + colour + "; stroke-dasharray: " + (r.line === "dashed" ? "5 3" : "none") + "; }");
      rules.push("#classes-arrow-" + r.id + " path { fill: " + (r.arrow === "triangle" ? "var(--bg)" : "none") + "; stroke: " + colour + "; stroke-width: 1.25px; }");
    });
    var style = document.createElement("style");
    style.id = "code-views-kinds";
    style.textContent = rules.join("\n");
    // First in the head, so the stylesheet's own rules for the same things (a selected box's outline) still win.
    document.head.insertBefore(style, document.head.firstChild);
  })();

  // Each node's edges, for the source panel's Dependencies list. Unit ids are also member-level node ids
  // (references to the unit itself), so the two levels are indexed separately.
  function indexEdges(edges) {
    var index = { uses: new Map(), usedBy: new Map() };
    function add(map, id, edge) {
      if (!map.has(id)) map.set(id, []);
      map.get(id).push(edge);
    }
    edges.forEach(function (e) {
      add(index.uses, e.from, e);
      add(index.usedBy, e.to, e);
    });
    return index;
  }
  var unitEdges = indexEdges(data.unitEdges);
  var memberEdges = indexEdges(data.memberEdges);

  function $(id) { return document.getElementById(id); }

  var escapeHtml = shell.escapeHtml;
  var settings = shell.settings;

  var listeners = { selection: [], shown: [], lock: [] };
  function emit(name, value) {
    listeners[name].forEach(function (fn) {
      try { fn(value); } catch (e) { console.error(e); }
    });
  }

  // ---- Source text ----
  //
  // Embedded in the page, or read with Load folder (see shell.js), a file at a time: units and members say which
  // file and lines they're in. How the text is decoded and highlighted is the tool's to say.

  var sourceReader = config.source(data);

  shell.source.init({
    mode: data.source.sourceMode,
    attribute: "data-file",
    files: function () { return data.files.map(function (f) { return { id: f.id, file: f.name, hash: f.hash }; }); },
    decode: sourceReader.decode,
    note: sourceReader.note || "",
    highlight: sourceReader.highlight,
    noFiles: "No source files for this project in that folder. Choose the folder that contains this page.",
    mismatch: "lines may not match",
    loaded: function () { renderSource(); }
  });

  // ---- Selection and the source panel ----

  var selection = null;
  var showWholeUnit = false;
  var shownPart = 0; // which part of a unit in several places the panel shows
  var panelHeader = shell.panel.header;
  var panelBody = shell.panel.body;

  function select(id, other) {
    id = id || null;
    if (id !== null && !unitsById.has(id) && !membersById.has(id)) return false;
    // First, so the panel is open when its source is drawn and scrolled. The same selection may change its focus.
    if (!shell.selectionChanged(id, other)) return true;
    selection = id;
    showWholeUnit = false;
    shownPart = 0;
    renderSource();
    rememberInHistory(id);
    emit("selection", selection);
    return true;
  }

  // ---- The explorer's filter ----
  //
  // While anything is ticked in the explorer, the views show only that (the shell keeps the ticks). A ticked unit
  // brings all its members, and a ticked member its unit: at unit level the unit stands for it.

  var unitsWithTickedMembers = new Set();

  function refreshFilter() {
    unitsWithTickedMembers = new Set();
    shell.filter.ids().forEach(function (id) {
      var member = membersById.get(id);
      if (member) unitsWithTickedMembers.add(member.unit);
    });
  }

  // First, so the views see the new filter when they hear of it.
  shell.onFilterChange(refreshFilter);

  function inFilter(id) {
    if (!shell.filter.active() || shell.filter.has(id) || shell.filter.alone(id)) return true;
    var member = membersById.get(id);
    return member ? shell.filter.has(member.unit) : unitsWithTickedMembers.has(id);
  }

  // ---- Show… choices ----
  //
  // One set, in the explorer, for the tree and every view. A choice applies wherever it means something: member kinds,
  // private members and references to the tree, the graph at member level and the class diagram's boxes; a line to
  // the views that draw it. The kinds and lines are the language's.

  // A line of the class diagram's own has a choice of its own, after the graph's.
  var lineChoices = language.relationships.map(function (r) { return { id: "line:" + r.id, label: r.showLabel || r.label }; });
  language.classRelationships.forEach(function (r) {
    if (r.choiceLabel && !lineChoices.some(function (c) { return c.id === r.choice; })) lineChoices.push({ id: r.choice, label: r.choiceLabel });
  });

  shell.choices.init({
    sections: [
      { title: capital(language.unit.many), choices: language.unitKinds.map(function (k) { return { id: "unit:" + k.id, label: k.plural }; }) },
      {
        title: capital(language.member.many),
        choices: MEMBER_CATEGORIES.map(function (c) { return { id: "member:" + c.id, label: c.label }; })
          .concat([{ id: "private", label: "Private " + language.member.many }])
      },
      {
        title: "References", radio: "references",
        choices: [{ value: "all", label: "All " + language.member.many }, { value: "unreferenced", label: "Unreferenced only" }, { value: "referenced", label: "Referenced only" }]
      },
      { title: "Lines", choices: lineChoices },
      { title: "Other", choices: [{ id: "unconnected", label: "Items with no connections" }] }
    ]
  });

  function kept(id) {
    var member = membersById.get(id);
    var unit = unitsById.get(member ? member.unit : id);
    if (!unit || !shell.choices.on("unit:" + unit.kind)) return false;
    if (!member) return true;
    if (!shell.choices.on("member:" + categoryOfKind[member.kind])) return false;
    if (member.visibility === "private" && !shell.choices.on("private")) return false;
    var references = shell.choices.value("references");
    return references === "all" || (references === "unreferenced") === !!member.unreferenced;
  }

  // ---- What the views show ----
  //
  // What's in the filter and not hidden. A hidden unit takes its members with it; a hidden member goes alone.

  var hiddenNow = new Set();

  function isHidden(id) {
    if (hiddenNow.has(id)) return true;
    var member = membersById.get(id);
    return !!member && hiddenNow.has(member.unit);
  }

  function shows(id) { return inFilter(id) && !isHidden(id); }

  // A unit's own node at the graph's member level. Unlocked, a ticked member brings it; locked, only what locking
  // found on screen is kept, so the unit has to be ticked or kept on its own.
  function showsItself(id) {
    if (!shows(id)) return false;
    return !shell.lock.on() || !shell.filter.active() || shell.filter.has(id) || shell.filter.alone(id);
  }

  function filterKey() {
    var alone = shell.filter.aloneIds();
    return shell.filter.ids().slice().sort().join("|") + (alone.length > 0 ? "\u0001" + alone.slice().sort().join("|") : "");
  }

  function shownKey() {
    return filterKey() + (shell.lock.on() ? "\u0001locked" : "") + "\u0000" + shell.hidden.list().slice().sort().join("|");
  }

  var lockSources = {};

  shell.onFilterChange(function () { emit("shown"); });
  // Locking and unlocking change the filter with the lock: the views hear of the lock, then draw once.
  shell.onLockChange(function (on) {
    refreshFilter();
    emit("lock", on);
    emit("shown");
  });
  shell.onHiddenChange(function () {
    hiddenNow = new Set(shell.hidden.list());
    if (selection !== null) renderHeader(unitsById.get(membersById.has(selection) ? membersById.get(selection).unit : selection), membersById.get(selection) || null);
    emit("shown");
  });
  // A saved view sets the selection, the filter, hiding and the Show… choices at once: the views draw once, each
  // reading the selection and its focus as it does. The shell has taken the selection already.
  shell.onViewApplied(function (id) {
    refreshFilter();
    hiddenNow = new Set(shell.hidden.list());
    selection = id;
    shownPart = 0;
    renderSource();
    rememberInHistory(id);
    emit("shown");
  });

  // How many of the items the selection is linked to, either way, are hidden.
  function hiddenNear(id, index) {
    var near = new Set();
    (index.uses.get(id) || []).forEach(function (e) { near.add(e.to); });
    (index.usedBy.get(id) || []).forEach(function (e) { near.add(e.from); });
    near.delete(id);
    var count = 0;
    near.forEach(function (other) { if (isHidden(other)) count++; });
    return count;
  }

  // ---- Back and Forward ----
  //
  // Each selection is added to the browser's history as the page address's #fragment (e.g. Sales-vba.html#Order.Save),
  // so Back and Forward step through selections, and an address with a fragment opens with that selection.

  var followingHistory = false;

  function idFromAddress() {
    var fragment = location.hash.slice(1);
    if (!fragment) return null;
    try { return decodeURIComponent(fragment); } catch (e) { return null; }
  }

  function rememberInHistory(id) {
    if (followingHistory || idFromAddress() === id) return;
    try {
      history.pushState(null, "", "#" + (id === null ? "" : encodeURIComponent(id)));
    } catch (e) { /* some browsers refuse history changes for pages opened from disk; selection still works */ }
  }

  function followAddress() {
    followingHistory = true;
    try { select(idFromAddress()); } finally { followingHistory = false; }
  }

  window.addEventListener("popstate", followAddress);

  // ---- Where a unit is ----
  //
  // A unit is the whole of its file unless it says where its parts are: a file can hold several units, and a unit
  // can be in several files.

  function partsOf(unit) {
    if (unit.parts && unit.parts.length > 0) return unit.parts;
    var file = filesById.get(unit.file);
    return [{ file: unit.file, first: 1, last: file ? file.lines : 0, whole: true }];
  }

  function partText(part) {
    return fileName(part.file) + (part.whole ? " · " + part.last + " lines" : " · lines " + part.first + " to " + part.last);
  }

  // What the panel and a printed item say of where a unit is: its file and how long it is, or each of its parts.
  function unitWhere(unit) { return partsOf(unit).map(partText).join("; "); }

  // The part of a member's unit that holds it, or the member's whole file.
  function partHolding(unit, member) {
    var holding = partsOf(unit).filter(function (p) { return p.file === member.file && member.start >= p.first && member.end <= p.last; })[0];
    if (holding) return holding;
    var file = filesById.get(member.file);
    return { file: member.file, first: 1, last: file ? file.lines : member.end };
  }

  // ---- Dependencies list ----
  //
  // What the selection uses and what uses it: direct links only, at the selection's own level (a unit lists
  // unit edges, a member lists member edges), whatever the graph's filters. Each entry selects its target,
  // so the list is also a way to walk the graph. Collapsed by default; whether it's open is remembered.

  var dependenciesOpen = settings.load("dependencies-open") === "true";

  // How an entry reads: a member is shown by name, with its unit when that differs from the selection's.
  function nodeText(id, ownUnit) {
    var member = membersById.get(id);
    if (!member) {
      var unit = unitsById.get(id);
      return { name: id, detail: ownUnit !== null ? "the " + language.unit.one + " itself" : unitKindLabel(unit) };
    }
    return {
      name: member.unit === ownUnit ? member.name : member.unit + "." + member.name,
      detail: memberKindLabel(member.kind)
    };
  }

  // The unit a node is in: a member's, or a unit node's own.
  function unitOfId(id) { var member = membersById.get(id); return member ? member.unit : id; }

  // The file a node's lines are in.
  function fileOfId(id) { var member = membersById.get(id); return member ? member.file : unitsById.get(id).file; }

  // Member edges by the units at their ends, for a unit's references.
  var memberEdgesByUnits = new Map();
  data.memberEdges.forEach(function (e) {
    var key = unitOfId(e.from) + "\u0000" + unitOfId(e.to);
    if (!memberEdgesByUnits.has(key)) memberEdgesByUnits.set(key, []);
    memberEdgesByUnits.get(key).push(e);
  });

  // Where member edges' references are: each line once, in order, in the member that holds it. A preview of one
  // (the shell's Peek) can show all of that member: its lines.
  function referencesOf(edges) {
    var seen = new Set();
    var refs = [];
    edges.forEach(function (e) {
      (e.lines || []).forEach(function (line) {
        var key = e.from + "\u0000" + line;
        if (seen.has(key)) return;
        seen.add(key);
        var member = membersById.get(e.from);
        var ref = { "in": e.from, source: fileOfId(e.from), line: line, inName: member ? member.unit + "." + member.name : e.from };
        if (member) { ref.first = member.start; ref.last = member.end; }
        refs.push(ref);
      });
    });
    return refs.sort(function (a, b) { return a.source < b.source ? -1 : a.source > b.source ? 1 : a.line - b.line; });
  }

  // One entry per item at the other end, whatever the kinds of link: a tag for each kind other than the plain one, how
  // many references when more than one, and the lines they're on. At unit level the lines are its members' references.
  function dependencyList(id, edges, direction, otherEnd, ownUnit) {
    var byOther = new Map();
    edges.forEach(function (e) {
      var other = otherEnd(e);
      if (!byOther.has(other)) byOther.set(other, []);
      byOther.get(other).push(e);
    });
    var atUnitLevel = ownUnit === null;
    var entries = Array.from(byOther.keys()).map(function (other) {
      var group = byOther.get(other);
      var text = nodeText(other, ownUnit);
      var tags = [];
      var count = group.reduce(function (sum, e) { return sum + (relationships[e.kind] && relationships[e.kind].plain ? e.count : 0); }, 0);
      if (count > 1) tags.push({ label: "×" + count, title: "Number of references", className: "muted" });
      group.forEach(function (e) {
        var relationship = relationships[e.kind];
        if (!relationship || relationship.plain) return;
        tags.push({ label: relationship[direction + "Tag"] || relationship.label, title: relationship[direction + "Title"] || "", className: "edge-tag " + e.kind });
      });
      var references = atUnitLevel
        ? memberEdgesByUnits.get(direction === "uses" ? id + "\u0000" + other : other + "\u0000" + id) || []
        : group;
      return { id: other, name: text.name, detail: text.detail, tags: tags, hidden: isHidden(other), refs: referencesOf(references) };
    });
    // Case-insensitive, without localeCompare, so the order is the same in every browser language.
    function key(entry) { return entry.name.toLowerCase() + " | " + entry.id; }
    entries.sort(function (a, b) { return key(a) < key(b) ? -1 : key(a) > key(b) ? 1 : 0; });
    return { count: entries.length, html: shell.references.listHtml(entries, "Nothing") };
  }

  function dependenciesHtml(id, index, ownUnit) {
    // Recursion is a self-loop: noted, not listed (selecting the selection does nothing).
    var recursive = false;
    function notSelf(e) { if (e.from === e.to) { recursive = true; return false; } return true; }
    var uses = dependencyList(id, (index.uses.get(id) || []).filter(notSelf), "uses", function (e) { return e.to; }, ownUnit);
    var usedBy = dependencyList(id, (index.usedBy.get(id) || []).filter(notSelf), "usedBy", function (e) { return e.from; }, ownUnit);
    return "<details class=\"dependencies\" id=\"source-dependencies\"" + (dependenciesOpen ? " open" : "") + ">" +
      "<summary>Dependencies <span class=\"muted\">uses " + uses.count + " · used by " + usedBy.count + "</span></summary>" +
      (recursive ? "<p class=\"muted recursive\">Calls itself (recursion).</p>" : "") +
      "<h3>Uses (" + uses.count + ")</h3>" +
      "<div class=\"dependency-list\" data-direction=\"uses\">" + uses.html + "</div>" +
      "<h3>Used by (" + usedBy.count + ")</h3>" +
      "<div class=\"dependency-list\" data-direction=\"usedBy\">" + usedBy.html + "</div>" +
      "</details>";
  }

  // What a unit or member says about itself, which the language's Core works out.
  function descriptionHtml(item) {
    return item.description ? "<p class=\"description\" id=\"source-description\">" + escapeHtml(item.description) + "</p>" : "";
  }

  function renderHeader(unit, member) {
    var html;
    if (member) {
      html = "<h2>" + escapeHtml(member.name) + shell.favourites.buttonHtml(member.id) + shell.hidden.buttonHtml(member.id) + "<span class=\"kind\">" + escapeHtml(memberKindLabel(member.kind)) +
        " in <a href=\"#\" data-select=\"" + escapeHtml(unit.id) + "\">" + escapeHtml(unit.id) + "</a></span></h2>" +
        "<p class=\"signature\">" + escapeHtml(member.signature) + "</p>" + descriptionHtml(member);
      var badges = [];
      if (member.role !== "none") {
        badges.push("<span class=\"badge\">" + escapeHtml(roleLabel(member.role) + (member.roleNote ? " (" + member.roleNote + ")" : "")) + "</span>");
      }
      if (member.unreferenced) {
        badges.push("<span class=\"badge warning\" title=\"" + escapeHtml(language.unreferencedWhy) + "\">Unreferenced</span>");
      }
      if (badges.length) html += "<div class=\"badges\">" + badges.join("") + "</div>";

      var unfollowed = unfollowedByMember.get(member.id) || [];
      if (unfollowed.length) {
        html += "<div class=\"muted\">" + escapeHtml(language.unfollowedHeading) + "</div><ul class=\"late-bound\">" + unfollowed.map(function (l) {
          return "<li><code>" + escapeHtml(l.expression) + "</code> (line " + l.line + ")</li>";
        }).join("") + "</ul>";
      }
      if (shell.source.has(member.file)) {
        html += "<div class=\"actions\"><button type=\"button\" class=\"button\" id=\"toggle-context\">" +
          (showWholeUnit ? "Show this " + escapeHtml(language.member.one) + " only" : "Show in the whole " + escapeHtml(language.unit.one)) + "</button></div>";
      }
    } else {
      var parts = partsOf(unit);
      html = "<h2>" + escapeHtml(unit.id) + shell.favourites.buttonHtml(unit.id) + shell.hidden.buttonHtml(unit.id) + "<span class=\"kind\">" + escapeHtml(unitKindLabel(unit)) + "</span></h2>" +
        "<p class=\"signature\">" + escapeHtml(unitWhere(unit)) + "</p>" + descriptionHtml(unit);
      // A unit in several places shows one at a time.
      if (parts.length > 1) {
        html += "<div class=\"actions unit-parts\" role=\"group\" aria-label=\"Parts\">" + parts.map(function (part, i) {
          return "<button type=\"button\" class=\"button\" data-part=\"" + i + "\" aria-pressed=\"" + (i === shownPart) + "\" title=\"" +
            escapeHtml("Show " + partText(part)) + "\">" + escapeHtml(fileName(part.file)) + "</button>";
        }).join("") + "</div>";
      }
    }
    var shownId = member ? member.id : unit.id;
    var index = member ? memberEdges : unitEdges;
    // A member hidden with its unit says so as the unit would.
    html += shell.hidden.noteHtml(member && !shell.hidden.has(member.id) && isHidden(member.id) ? unit.id : shownId, hiddenNear(shownId, index));
    html += dependenciesHtml(shownId, index, member ? unit.id : null);
    panelHeader.innerHTML = html;

    var toggle = $("toggle-context");
    if (toggle) toggle.addEventListener("click", function () { showWholeUnit = !showWholeUnit; renderSource(); });
    panelHeader.querySelectorAll("[data-part]").forEach(function (button) {
      button.addEventListener("click", function () { shownPart = Number(button.getAttribute("data-part")); renderSource(); });
    });
    // Noted on click, not on the "toggle" event, which comes later: by then a new selection may have redrawn the header.
    var dependencies = $("source-dependencies");
    dependencies.querySelector("summary").addEventListener("click", function () {
      dependenciesOpen = !dependencies.open;
      settings.save("dependencies-open", dependenciesOpen ? "true" : "false");
    });
    panelHeader.querySelectorAll("[data-select]").forEach(function (link) {
      link.addEventListener("click", function (e) { e.preventDefault(); select(link.getAttribute("data-select"), e.ctrlKey || e.metaKey); });
    });
  }

  function renderSource() {
    if (selection === null) {
      panelHeader.innerHTML = "";
      panelBody.innerHTML = "<p class=\"source-message\">Select a " + escapeHtml(language.unit.one) + " or " + escapeHtml(language.member.one) + " to see its source.</p>";
      return;
    }

    var member = membersById.get(selection) || null;
    var unit = unitsById.get(member ? member.unit : selection);
    var parts = partsOf(unit);
    if (shownPart >= parts.length) shownPart = 0;
    renderHeader(unit, member);

    var file = member ? member.file : parts[shownPart].file;
    if (!shell.source.has(file)) {
      if (data.source.sourceMode === "folder") {
        // The shell opens Load folder from the button.
        panelBody.innerHTML = "<div class=\"source-message\"><p>The source isn't loaded yet. Choose the folder that " +
          "contains this page and its " + escapeHtml(sourceReader.files) + ".</p><button type=\"button\" class=\"button primary\" id=\"load-folder-prompt\">Load folder</button></div>";
      } else {
        panelBody.innerHTML = "<p class=\"source-message\">No source was found for " + escapeHtml(unit.id) + ".</p>";
      }
      return;
    }

    // A unit shows the part chosen. A member shows its own lines, or the part of its unit that holds it with its
    // lines marked.
    var count = shell.source.lines(file).length;
    var first, last, marked = null;
    if (!member) {
      first = parts[shownPart].first;
      last = parts[shownPart].last;
    } else if (showWholeUnit) {
      var holding = partHolding(unit, member);
      first = holding.first;
      last = holding.last;
      marked = { first: member.start, last: member.end };
    } else {
      first = member.start;
      last = member.end;
    }
    panelBody.innerHTML = shell.source.codeHtml(file, Math.max(1, first), Math.min(count, last), marked);

    var firstInMember = panelBody.querySelector(".marked");
    if (firstInMember) firstInMember.scrollIntoView({ block: "start" });
    else panelBody.scrollTop = 0;
  }

  // A saved view keeps what the source panel shows of the selection: a member alone or in its whole unit, and
  // whether its Dependencies are open. The shell keeps the selection.
  shell.views.part("source", {
    save: function () { return { whole: showWholeUnit, dependencies: dependenciesOpen }; },
    apply: function (value) {
      value = value && typeof value === "object" ? value : {};
      showWholeUnit = value.whole === true;
      if (typeof value.dependencies === "boolean") {
        dependenciesOpen = value.dependencies;
        settings.save("dependencies-open", dependenciesOpen ? "true" : "false");
      }
    }
  });

  // ---- Explorer ----
  //
  // Two layouts: units in their groups (folders, namespaces, packages: the language says what they're called), and
  // units by kind. Groups come first when the model has any; without, there's only the kinds. Each unit lists its
  // members in source order, with the members that belong under another (an Enum's values) under it.

  function byLowerName(a, b) {
    var x = a.toLowerCase();
    var y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : a < b ? -1 : a > b ? 1 : 0;
  }

  function memberNode(member, children) {
    return {
      id: member.id,
      label: member.name,
      detail: memberKindLabel(member.kind),
      badge: member.role !== "none" ? roleLabel(member.role) : null,
      className: member.unreferenced ? "unreferenced" : "",
      title: member.signature + (member.unreferenced ? " (nothing in the code refers to it)" : ""),
      children: children
    };
  }

  function unitNode(unit, showKind, membersOf) {
    var members = (membersOf.get(unit.id) || []).slice().sort(function (a, b) {
      return a.file === b.file ? a.start - b.start || a.declaration - b.declaration : byLowerName(a.file, b.file);
    });
    // A member that names a parent among the unit's other members goes under it.
    var names = new Set(members.filter(function (m) { return !m.parent; }).map(function (m) { return m.name; }));
    var under = new Map();
    members.forEach(function (m) {
      if (!m.parent || !names.has(m.parent)) return;
      if (!under.has(m.parent)) under.set(m.parent, []);
      under.get(m.parent).push(m);
    });
    var children = members
      .filter(function (m) { return !(m.parent && names.has(m.parent)); })
      .map(function (m) { return memberNode(m, under.has(m.name) && !m.parent ? under.get(m.name).map(function (v) { return memberNode(v); }) : undefined); });
    return {
      id: unit.id, label: unit.id, title: unitWhere(unit), children: children,
      detail: showKind ? unitKindLabel(unit) : unit.note || ""
    };
  }

  // ---- Widely used ----
  //
  // Units used by at least a share of the other units (Core counts them, as usedBy), most used first, each with its
  // own members that are, under it. A member is never used by more units than its unit, so the unit stands for it.

  function widelyUsed(share) {
    var others = data.units.length - 1;
    var least = Math.max(2, Math.ceil(share * others / 100));
    function title(count) { return "Used by " + count + " of the " + others + " other " + language.unit.many; }
    function byUse(a, b) { return b.usedBy - a.usedBy || byLowerName(a.id, b.id); }
    var membersOf = new Map();
    data.members.forEach(function (m) {
      if ((m.usedBy || 0) < least) return;
      if (!membersOf.has(m.unit)) membersOf.set(m.unit, []);
      membersOf.get(m.unit).push(m);
    });
    return data.units.filter(function (m) { return (m.usedBy || 0) >= least; }).sort(byUse).map(function (m) {
      return {
        id: m.id, label: m.id, detail: "used by " + m.usedBy, title: title(m.usedBy),
        children: (membersOf.get(m.id) || []).sort(byUse).map(function (member) {
          return { id: member.id, label: member.name, detail: "used by " + member.usedBy, title: title(member.usedBy) };
        })
      };
    });
  }

  function explorerLayouts() {
    var membersOf = new Map();
    data.members.forEach(function (m) {
      if (!membersOf.has(m.unit)) membersOf.set(m.unit, []);
      membersOf.get(m.unit).push(m);
    });
    var units = data.units.slice().sort(function (a, b) { return byLowerName(a.id, b.id); });

    var types = {
      name: "types", label: "Types", title: capital(language.unit.many) + " by kind",
      nodes: language.unitKinds.map(function (kind) {
        var inGroup = units.filter(function (u) { return u.kind === kind.id; });
        return { label: kind.plural, detail: String(inGroup.length), children: inGroup.map(function (u) { return unitNode(u, false, membersOf); }) };
      }).filter(function (group) { return group.children.length > 0; })
    };
    if (data.groups.length === 0) return [types];
    return [{ name: "folders", label: language.groupsLabel, title: language.groupsTitle, nodes: groupTree(units, membersOf) }, types];
  }

  // Groups nest as the model says. A unit in no group goes in a group of its own at the end.
  function groupTree(units, membersOf) {
    var childrenOf = new Map();
    var unitsOf = new Map();
    data.groups.forEach(function (g) {
      var parent = g.parent || "";
      if (!childrenOf.has(parent)) childrenOf.set(parent, []);
      childrenOf.get(parent).push(g);
    });
    var loose = [];
    units.forEach(function (u) {
      if (!u.group) { loose.push(u); return; }
      if (!unitsOf.has(u.group)) unitsOf.set(u.group, []);
      unitsOf.get(u.group).push(u);
    });

    function count(group) {
      var total = (unitsOf.get(group.id) || []).length;
      (childrenOf.get(group.id) || []).forEach(function (child) { total += count(child); });
      return total;
    }

    function groupNodes(parent) {
      return (childrenOf.get(parent) || []).slice().sort(function (a, b) { return byLowerName(a.name, b.name); }).map(function (group) {
        return {
          label: group.name,
          folder: true,
          detail: String(count(group)),
          children: groupNodes(group.id).concat((unitsOf.get(group.id) || []).map(function (u) { return unitNode(u, true, membersOf); }))
        };
      });
    }

    var nodes = groupNodes("");
    if (loose.length > 0) {
      nodes.push({ label: language.ungrouped, detail: String(loose.length), children: loose.map(function (u) { return unitNode(u, true, membersOf); }) });
    }
    return nodes;
  }

  // ---- Search ----

  var searchBox = $("search");
  var idsByLowerCase = new Map();
  var options = [];
  data.units.forEach(function (u) {
    idsByLowerCase.set(u.id.toLowerCase(), u.id);
    options.push("<option value=\"" + escapeHtml(u.id) + "\">" + escapeHtml(unitKindLabel(u)) + "</option>");
  });
  data.members.forEach(function (m) {
    idsByLowerCase.set(m.id.toLowerCase(), m.id);
    options.push("<option value=\"" + escapeHtml(m.id) + "\">" + escapeHtml(m.signature) + "</option>");
  });
  $("search-list").innerHTML = options.join("");

  function searchFor(text, allowPartial) {
    var wanted = text.trim().toLowerCase();
    if (!wanted) return;
    var id = idsByLowerCase.get(wanted);
    if (!id && allowPartial) {
      var keys = Array.from(idsByLowerCase.keys());
      var match = keys.find(function (k) { return k.indexOf(wanted) === 0; }) ||
                  keys.find(function (k) { return k.indexOf(wanted) >= 0; });
      id = match ? idsByLowerCase.get(match) : null;
    }
    if (id) select(id);
  }

  searchBox.addEventListener("input", function () { searchFor(searchBox.value, false); }); // picking from the list
  searchBox.addEventListener("keydown", function (e) { if (e.key === "Enter") searchFor(searchBox.value, true); });

  // ---- Start ----

  var api = {
    tool: config.tool,
    data: data,
    language: language,
    unitsById: unitsById,
    membersById: membersById,
    unitKind: unitKind,
    unitKindLabel: unitKindLabel,
    memberKind: function (kind) { return memberKinds[kind] || null; },
    memberKindLabel: memberKindLabel,
    memberCategories: MEMBER_CATEGORIES,
    relationship: function (id) { return relationships[id] || null; },
    roleLabel: roleLabel,
    capital: capital,
    select: select,
    selection: function () { return selection; },
    onSelectionChange: function (fn) { listeners.selection.push(fn); },
    theme: shell.theme,
    onThemeChange: shell.onThemeChange,
    activeView: shell.activeView,
    onViewChange: shell.onViewChange,
    panelFollowing: shell.panel.following,
    focused: shell.focused,
    onFocusChange: shell.onFocusChange,
    focusOnClick: shell.focusOnClick,
    setFocusOnClick: shell.setFocusOnClick,
    favourites: shell.favourites.list,
    isFavourite: shell.favourites.has,
    toggleFavourite: shell.favourites.toggle,
    inFilter: inFilter,
    filterActive: shell.filter.active,
    filterKey: filterKey,
    onFilterChange: shell.onFilterChange,
    locked: shell.lock.on,
    onLockChange: function (fn) { listeners.lock.push(fn); },
    lockSource: function (view, fn) { lockSources[view] = fn; },
    showsItself: showsItself,
    hidden: shell.hidden.list,
    isHidden: isHidden,
    toggleHidden: shell.hidden.toggle,
    shows: shows,
    shownKey: shownKey,
    onShownChange: function (fn) { listeners.shown.push(fn); },
    focusFavourites: shell.focusFavourites,
    setFocusFavourites: shell.setFocusFavourites,
    escapeHtml: shell.escapeHtml,
    showLine: shell.source.showLine,
    kept: kept,
    choice: shell.choices.on,
    choiceValue: shell.choices.value,
    choicesKey: shell.choices.key,
    choicesCount: shell.choices.count,
    onChoicesChange: shell.onChoicesChange,
    onFavouritesChange: shell.onFavouritesChange,
    viewPart: shell.views.part,
    preparePrint: shell.print.prepare,
    finishPrint: shell.print.finish,
    viewText: shell.copy.text
  };
  // The views' name for it, and the tool's own.
  window.codeViews = api;
  window[config.tool] = api;

  // Favourites and the filter are per source. In the explorer's Favourites group a member shows its unit, as it
  // isn't under it. The filter counts the units it shows.
  shell.items.init({
    source: data.source.title,
    sourceWord: config.sourceWord,
    kept: kept,
    widelyUsed: widelyUsed,
    exists: function (id) { return unitsById.has(id) || membersById.has(id); },
    // What Lock view ticks: what the open view shows, if it's one that can say.
    shown: function () {
      var source = lockSources[shell.activeView()];
      return source ? source() : null;
    },
    node: function (id) {
      var member = membersById.get(id);
      if (!member) return { label: id, detail: unitKindLabel(unitsById.get(id)), title: unitWhere(unitsById.get(id)) };
      var node = memberNode(member);
      node.label = member.unit + "." + member.name;
      return node;
    },
    describeFilter: function (ids) {
      var units = new Set(ids.map(unitOfId));
      return units.size + " of " + data.units.length + " " + (data.units.length === 1 ? language.unit.one : language.unit.many);
    }
  });
  refreshFilter();
  hiddenNow = new Set(shell.hidden.list());
  shell.explorer.setLayouts(explorerLayouts(), select);

  // The Export tab is the tool's to fill, if it has one.
  if (config.exports) {
    shell.exports.init(config.exports({ data: data, unitsById: unitsById, unitEdges: unitEdges, unitOfId: unitOfId }));
  }

  // ---- Printing ----
  //
  // The shell makes the document (see shell.js); this says when the open view is drawn and still, which element is
  // its diagram, and what there is to say about an item.

  // Kept once the test passes, looked at ten times a second, or after eight seconds whatever it says.
  function whenStill(test) {
    return new Promise(function (resolve) {
      var tries = 0;
      (function check() {
        window.setTimeout(function () { if (test() || ++tries > 80) resolve(); else check(); }, 100);
      })();
    });
  }

  function namesAt(edges, otherEnd, ownUnit) {
    var names = new Set();
    (edges || []).forEach(function (e) { if (e.from !== e.to) names.add(nodeText(otherEnd(e), ownUnit).name); });
    return Array.from(names).sort(byLowerName);
  }

  shell.print.init({
    ready: function () {
      var view = shell.activeView();
      return whenStill(function () {
        if (view === "graph") return !!window.codeViewsGraph && !window.codeViewsGraph.moving() && !window.codeViewsGraph.sliding();
        if (view === "classes") return !!window.codeViewsClasses && window.codeViewsClasses.drawn() && !window.codeViewsClasses.moving();
        return true;
      });
    },
    element: function () {
      var view = shell.activeView();
      if (view === "graph") return $("graph-svg");
      if (view === "classes") return $("classes-svg");
      if (view === "calls") return $("calls-sequence").hidden ? $("calls-tree") : $("calls-sequence");
      return null;
    },
    // For Copy: the Calls tab adds the tree as it's opened, and code blocks are the language's.
    language: language.codeBlock,
    text: function () { return shell.activeView() === "calls" && window.codeViewsCalls ? window.codeViewsCalls.text() : ""; },
    describe: function (id) {
      var member = membersById.get(id);
      var unit = unitsById.get(member ? member.unit : id);
      if (!unit) return null;
      var index = member ? memberEdges : unitEdges;
      var own = member ? unit.id : null;
      var first = partsOf(unit)[0];
      return {
        name: member ? unit.id + "." + member.name : unit.id,
        kind: member ? memberKindLabel(member.kind) : unitKindLabel(unit),
        signature: member ? member.signature : "",
        note: member ? "" : unitWhere(unit),
        description: (member || unit).description || "",
        uses: namesAt(index.uses.get(id), function (e) { return e.to; }, own),
        usedBy: namesAt(index.usedBy.get(id), function (e) { return e.from; }, own),
        source: member ? { id: member.file, first: member.start, last: member.end }
          : first.whole ? { id: first.file } : { id: first.file, first: first.first, last: first.last }
      };
    }
  });

  shell.start();
  // The views read the selection when they first draw, so an address's selection is made before they do.
  if (idFromAddress() !== null) followAddress();
  renderSource();
})();
