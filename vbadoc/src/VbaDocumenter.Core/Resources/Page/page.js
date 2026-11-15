// page.js: what's vbadoc's own in its page. The page itself (the selection, the source panel, the explorer, search)
// and its views (dependency graph, class diagram, Calls) are the code views library's (code-page.js and the views'
// scripts in CodeDocumenter.CodeViews), which reads this before it starts: see the top of code-page.js.
//
// vbadoc's own is what's about a workbook's VBA project: its text is in the project's code page, and whole modules
// can be taken out of it for another workbook (the shell's Export tab). What that needs beyond the code model is in
// the model's extra (the project's code page and references) and each unit's (PageProject and PageModuleExport).
(function () {
  "use strict";

  window.codeViewsPage = {
    tool: "vbadoc",
    sourceWord: "workbook",

    // Module files are written in the project's code page, so Load folder decodes them with it.
    source: function (data) {
      var decoder;
      var note = "";
      try {
        decoder = new TextDecoder(data.extra.encoding);
      } catch (e) {
        decoder = new TextDecoder("windows-1252");
        note = " This browser can't decode code page " + data.extra.codePage + "; non-ASCII characters may be wrong.";
      }
      return {
        decode: function (buffer) { return decoder.decode(buffer); },
        note: note,
        highlight: function (lines) { return window.VbaHighlight.highlightLines(lines); },
        files: ".bas, .cls and .frm files"
      };
    },

    // The Export tab takes whole modules. Core has worked out which can come and what each relies on.
    exports: function (page) {
      // What Core says of a module for this tab is in the unit's extra: why it can't come, and what it relies on.
      function exportOf(id) { var unit = page.unitsById.get(id); return unit ? unit.extra || {} : null; }

      function whyNotTaken(id) {
        var module = exportOf(id);
        return module ? module.whyNot || null : "Only whole modules can be taken";
      }

      function modulesLinked(edges, otherEnd) {
        return Array.from(new Set((edges || []).map(otherEnd))).filter(function (id) { return !whyNotTaken(id); });
      }

      return {
        what: "modules and classes",
        whyNot: whyNotTaken,
        uses: function (id) { return modulesLinked(page.unitEdges.uses.get(id), function (e) { return e.to; }); },
        needs: function (id) { return exportOf(id).needs || []; },
        offers: function (id) {
          return modulesLinked((page.unitEdges.usedBy.get(id) || []).filter(function (e) { return e.kind === "implements"; }), function (e) { return e.from; });
        },
        pickFor: page.unitOfId,
        references: page.data.extra.references,
        writeTo: "--new Shared.xlsm"
      };
    }
  };
})();
