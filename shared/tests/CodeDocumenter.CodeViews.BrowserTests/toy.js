// toy.js: what a tool for the made-up language Toy tells the page of code about itself (see the top of
// code-page.js). Its files are UTF-8, and its highlighter marks nothing.
(function () {
  "use strict";

  function escapeHtml(text) {
    return text.replace(/[&<>"]/g, function (c) { return c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : "&quot;"; });
  }

  window.codeViewsPage = {
    tool: "toydoc",
    sourceWord: "toybox",
    source: function () {
      var decoder = new TextDecoder("utf-8");
      return {
        decode: function (buffer) { return decoder.decode(buffer); },
        highlight: function (lines) { return lines.map(escapeHtml); },
        files: ".toy files"
      };
    }
  };
})();
