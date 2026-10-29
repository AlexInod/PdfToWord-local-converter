(function () {
  var COPY = {
    pdf_to_word: { title: "Drop PDF files here", hint: "or " },
    word_to_pdf: { title: "Drop Word files here", hint: ".docx or .doc, or " }
  };

  // Call a pywebview API method; resolves to undefined when the API is missing.
  function api(name) {
    var args = Array.prototype.slice.call(arguments, 1);
    try {
      if (window.pywebview && window.pywebview.api && window.pywebview.api[name]) {
        return Promise.resolve(window.pywebview.api[name].apply(null, args));
      }
    } catch (e) { /* fall through */ }
    return Promise.resolve();
  }

  function $(id) { return document.getElementById(id); }

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined) n.textContent = text;
    return n;
  }

  function renderStatus(st, job) {
    st.className = "st";
    st.textContent = "";
    if (job.status === "waiting") {
      st.textContent = "Waiting";
    } else if (job.status === "converting") {
      var bar = el("span", "prog"), fill = el("div");
      bar.appendChild(fill);
      if (job.pages > 0) {
        st.appendChild(document.createTextNode("Page " + job.page + " of " + job.pages));
        fill.style.width = Math.min(100, Math.max(0, job.page / job.pages * 100)) + "%";
      } else {
        st.appendChild(document.createTextNode("Converting…"));
        bar.className = "prog indet";
      }
      st.appendChild(bar);
    } else if (job.status === "done") {
      var link = el("a", "", "Show");
      link.href = "#";
      link.onclick = function (e) { e.preventDefault(); api("show_file", job.id); };
      st.appendChild(document.createTextNode("Done"));
      st.appendChild(link);
    } else if (job.status === "failed") {
      st.className = "st failed";
      st.textContent = job.message || "";
    }
  }

  var app = {
    mode: "pdf_to_word",

    onJob: function (job) {
      var row = document.querySelector('[data-id="' + job.id + '"]');
      if (!row) {
        row = el("div", "row");
        row.setAttribute("data-id", job.id);
        row.appendChild(el("span", "name"));
        row.appendChild(el("span", "st"));
        $("list").appendChild(row);
      }
      var name = row.firstChild;
      name.textContent = job.name;
      name.title = job.name;
      renderStatus(row.lastChild, job);
      if (job.status === "done") $("open-folder").disabled = false;
    },

    setMode: function (mode, initial) {
      app.mode = mode;
      var word = mode === "word_to_pdf";
      $("mode-pdf").className = word ? "" : "on";
      $("mode-word").className = word ? "on" : "";
      $("mode-pdf").setAttribute("aria-selected", String(!word));
      $("mode-word").setAttribute("aria-selected", String(word));
      $("banner").hidden = true;
      document.querySelector(".seg").classList.toggle("word", word);
      $("drop-title").textContent = COPY[mode].title;
      $("drop-hint").textContent = COPY[mode].hint;
      if (initial) return;
      api("set_mode", mode);
      if (word) {
        api("office_available").then(function (ok) {
          if (app.mode === "word_to_pdf") $("banner").hidden = ok !== false;
        });
      }
    }
  };
  window.app = app;

  function init() {
    $("mode-pdf").onclick = function () { app.setMode("pdf_to_word"); };
    $("mode-word").onclick = function () { app.setMode("word_to_pdf"); };
    $("choose").onclick = function (e) { e.preventDefault(); api("choose_files"); };
    $("open-folder").onclick = function () { api("open_folder"); };

    var drop = $("drop");
    ["dragenter", "dragover"].forEach(function (t) {
      drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add("hover"); });
    });
    drop.addEventListener("dragleave", function () { drop.classList.remove("hover"); });
    drop.addEventListener("drop", function (e) { e.preventDefault(); drop.classList.remove("hover"); });

    app.setMode("pdf_to_word", true); // render only; no API call before pywebviewready
  }

  document.addEventListener("DOMContentLoaded", init);
})();
