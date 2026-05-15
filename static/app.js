// MCQ Shuffler — UI controller
// Communicates with the Flask backend at the same origin.

const $ = (sel) => document.querySelector(sel);

const state = {
  paper_id: null,
  name: null,
  n_questions: 0,
  has_katex: false,
  persisted: false,
};

// --- Upload -----------------------------------------------------------------

$("#upload-form").addEventListener("submit", async (ev) => {
  ev.preventDefault();
  const file = $("#file-input").files[0];
  if (!file) return;

  const persist = $("#persist").checked;
  const name = $("#paper-name").value.trim();

  const fd = new FormData();
  fd.append("file", file);
  fd.append("persist", persist ? "true" : "false");
  if (name) fd.append("name", name);

  setError("#upload-error", null);
  const btn = $("#upload-btn");
  btn.disabled = true;
  btn.textContent = "Parsing…";

  try {
    const r = await fetch("/upload", { method: "POST", body: fd });
    const data = await r.json();
    if (!r.ok) {
      setError("#upload-error", data.error || `HTTP ${r.status}`);
      return;
    }
    state.paper_id   = data.paper_id;
    state.name       = data.name;
    state.n_questions = data.n_questions;
    state.has_katex  = data.has_katex;
    state.persisted  = data.persisted;
    showUploadResult();
    enableGenerateStep();
    if (data.persisted) refreshSavedList();
  } catch (e) {
    setError("#upload-error", e.message || "Upload failed.");
  } finally {
    btn.disabled = false;
    btn.textContent = "Upload & parse";
  }
});

function showUploadResult() {
  const el = $("#upload-result");
  el.hidden = false;
  el.innerHTML =
    `Parsed <strong>${state.n_questions}</strong> questions from ` +
    `<strong>${escapeHtml(state.name)}</strong>` +
    (state.has_katex
      ? ` — KaTeX math detected (will render in Word output).`
      : ` — no KaTeX math detected.`) +
    (state.persisted ? ` <em>Saved to local database.</em>` : "");
}

// --- Generate ---------------------------------------------------------------

function enableGenerateStep() {
  $("#step-generate").hidden = false;
  $("#paper-summary").innerHTML =
    `Working with <strong>${escapeHtml(state.name)}</strong> — ` +
    `${state.n_questions} questions` +
    (state.has_katex ? `, KaTeX present` : ``) +
    `.`;
  window.scrollTo({
    top: $("#step-generate").offsetTop - 16,
    behavior: "smooth",
  });
}

$("#generate-form").addEventListener("submit", async (ev) => {
  ev.preventDefault();
  setError("#generate-error", null);
  setStatus("#generate-status", null);

  const n_sets = parseInt($("#n-sets").value, 10);
  if (!(n_sets >= 1 && n_sets <= 20)) {
    setError("#generate-error", "Number of sets must be between 1 and 20.");
    return;
  }

  const fmt = document.querySelector('input[name="format"]:checked').value;
  const headerMode = (document.querySelector('input[name="header_mode"]:checked') || {}).value || "none";

  // Multipart body so the header image can ride along.
  const fd = new FormData();
  fd.append("paper_id", state.paper_id);
  fd.append("n_sets", String(n_sets));
  fd.append("shuffle_questions", $("#shuffle-questions").checked ? "true" : "false");
  fd.append("shuffle_options",   $("#shuffle-options").checked   ? "true" : "false");
  fd.append("format", fmt);
  fd.append("persist", state.persisted ? "true" : "false");
  fd.append("math_in_docx", document.querySelector('input[name="math_in_docx"]:checked').value);
  fd.append("math_in_data", document.querySelector('input[name="math_in_data"]:checked').value);
  fd.append("header_mode", headerMode);
  if (headerMode === "custom") {
    const f = $("#header-image-file").files[0];
    if (!f) {
      setError("#generate-error", "Custom header chosen but no image file selected.");
      return;
    }
    fd.append("header_image", f);
  }

  const btn = $("#generate-btn");
  btn.disabled = true;
  btn.textContent = "Building ZIP…";
  setStatus("#generate-status",
            fmt.startsWith("pdf_")
              ? "PDF generation runs LibreOffice; this can take 10–30 seconds per set…"
              : "Generating sets, this may take a few seconds for math-heavy papers…");

  try {
    const r = await fetch("/generate", { method: "POST", body: fd });
    if (!r.ok) {
      let msg = `HTTP ${r.status}`;
      try { const j = await r.json(); msg = j.error || msg; } catch {}
      setError("#generate-error", msg);
      return;
    }
    const blob = await r.blob();
    const cd = r.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename="?([^"]+)"?/);
    const fname = m ? m[1] : `${state.name}_sets.zip`;
    triggerDownload(blob, fname);
    setStatus("#generate-status",
              `Done — ${n_sets} set${n_sets > 1 ? "s" : ""} downloaded as ${fname}.`);
  } catch (e) {
    setError("#generate-error", e.message || "Generation failed.");
  } finally {
    btn.disabled = false;
    btn.textContent = "Generate & download ZIP";
  }
});

// Show/hide controls based on selections
function refreshConditionalControls() {
  const fmt = (document.querySelector('input[name="format"]:checked') || {}).value;
  const isDocxOrPdf = fmt && (fmt.startsWith("docx_") || fmt.startsWith("pdf_"));
  const docxFs = $("#math-docx-fieldset");
  const dataFs = $("#math-data-fieldset");
  const hdrFs = $("#header-fieldset");
  if (docxFs) docxFs.hidden = !isDocxOrPdf;
  if (dataFs) dataFs.hidden = isDocxOrPdf;
  if (hdrFs)  hdrFs.hidden  = !isDocxOrPdf;

  const headerMode = (document.querySelector('input[name="header_mode"]:checked') || {}).value;
  const customField = $("#custom-header-field");
  if (customField) customField.hidden = headerMode !== "custom";
}
document.addEventListener("change", (e) => {
  if (e.target && (e.target.name === "format" || e.target.name === "header_mode")) {
    refreshConditionalControls();
  }
});
refreshConditionalControls();

// --- Sample files -----------------------------------------------------------

async function refreshSamples() {
  const el = $("#samples-list");
  try {
    const r = await fetch("/samples");
    const data = await r.json();
    const list = data.samples || [];
    if (list.length === 0) {
      el.innerHTML = `<em class="muted">No sample files available.</em>`;
      return;
    }
    el.innerHTML = "";
    for (const s of list) {
      const div = document.createElement("div");
      div.className = "saved-row";
      div.innerHTML = `
        <div>
          <div><strong>${escapeHtml(s.filename)}</strong></div>
          <div class="meta">${escapeHtml(s.description)}</div>
        </div>
        <div>
          <a class="use" href="/samples/${encodeURIComponent(s.filename)}" download>Download</a>
        </div>`;
      el.appendChild(div);
    }
  } catch (e) {
    el.innerHTML = `<em class="muted">Could not load samples: ${escapeHtml(e.message || e)}</em>`;
  }
}
refreshSamples();

$("#reset-btn").addEventListener("click", () => {
  state.paper_id = null;
  state.name = null;
  state.n_questions = 0;
  state.has_katex = false;
  state.persisted = false;
  $("#step-generate").hidden = true;
  $("#upload-form").reset();
  $("#upload-result").hidden = true;
  setError("#upload-error", null);
  setError("#generate-error", null);
  setStatus("#generate-status", null);
});

// --- Saved papers -----------------------------------------------------------

async function refreshSavedList() {
  const el = $("#saved-list");
  try {
    const r = await fetch("/papers");
    const data = await r.json();
    const papers = data.papers || [];
    if (papers.length === 0) {
      el.innerHTML = `<em class="muted">Nothing saved yet. Tick the "save in local database" box on upload to keep papers here.</em>`;
      return;
    }
    el.innerHTML = "";
    for (const p of papers) {
      const div = document.createElement("div");
      div.className = "saved-row";
      const ts = new Date(p.created_at * 1000);
      div.innerHTML = `
        <div>
          <div><strong>${escapeHtml(p.name)}</strong></div>
          <div class="meta">${escapeHtml(p.source_filename)} — saved ${ts.toLocaleString()}</div>
        </div>
        <div>
          <button class="use">Use</button>
          <button class="del">Delete</button>
        </div>`;
      div.querySelector(".use").addEventListener("click", () => usePaper(p));
      div.querySelector(".del").addEventListener("click", () => deletePaper(p.id));
      el.appendChild(div);
    }
  } catch (e) {
    el.innerHTML = `<em class="muted">Could not load saved papers: ${escapeHtml(e.message || e)}</em>`;
  }
}

async function usePaper(p) {
  // Re-resolve by hitting /papers/<id>/sets just to confirm it exists; for
  // generation we only need the paper_id and a question count. We don't have a
  // direct "fetch metadata" endpoint, so we use the list entry's fields and
  // run a single dry parse on the server via /generate's first step.
  state.paper_id = p.id;
  state.name = p.name;
  state.persisted = true;
  // We don't know n_questions until generate runs. Show a friendly placeholder.
  state.n_questions = "?";
  state.has_katex = false;
  $("#upload-result").hidden = false;
  $("#upload-result").innerHTML =
    `Loaded saved paper <strong>${escapeHtml(p.name)}</strong>. ` +
    `Choose options below and generate.`;
  enableGenerateStep();
}

async function deletePaper(id) {
  if (!confirm("Delete this saved paper? This cannot be undone.")) return;
  try {
    const r = await fetch(`/papers/${id}`, { method: "DELETE" });
    if (!r.ok) throw new Error((await r.json()).error || `HTTP ${r.status}`);
    refreshSavedList();
  } catch (e) {
    alert("Delete failed: " + (e.message || e));
  }
}

// --- Helpers ----------------------------------------------------------------

function setError(sel, msg) {
  const el = $(sel);
  if (!msg) { el.hidden = true; el.textContent = ""; return; }
  el.hidden = false;
  el.textContent = msg;
}

function setStatus(sel, msg) {
  const el = $(sel);
  if (!msg) { el.hidden = true; el.textContent = ""; return; }
  el.hidden = false;
  el.textContent = msg;
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

// --- Init -------------------------------------------------------------------

refreshSavedList();
