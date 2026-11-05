const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const dmy = iso => (iso || "").split("-").reverse().join("-");
const cat = r => r.cur === "\u00d7" ? "bad" : (r.red ? "gaps" : "ok");
let report = null, filter = "all", query = "";
const shut = new Set();

/* ---- client picker ---- */
function drawClients(q = "") {
  $("#client-list").innerHTML = CLIENTS.filter(c => c.name.toLowerCase().includes(q.toLowerCase()))
    .map(c => `<label><input type="checkbox" name="clients" value="${c.id}" ${c.on === false ? "" : "checked"}> ${esc(c.name)}<small>${c.channels}</small></label>`).join("")
    || '<p class="hint" style="padding:8px 10px">No client matches.</p>';
}
function syncChecks() {
  document.querySelectorAll("#client-list input").forEach(i => { CLIENTS.find(c => c.id == i.value).on = i.checked; });
  $("#partial-hint").hidden = CLIENTS.every(c => c.on !== false);
}
$("#client-list").addEventListener("change", syncChecks);
$("#client-q").addEventListener("input", e => drawClients(e.target.value));
$("#all").onclick = () => { CLIENTS.forEach(c => c.on = true); drawClients($("#client-q").value); syncChecks(); };
$("#none").onclick = () => { CLIENTS.forEach(c => c.on = false); drawClients($("#client-q").value); syncChecks(); };

/* ---- run + progress ---- */
$("#run-form").addEventListener("submit", async e => {
  e.preventDefault();
  const picked = CLIENTS.filter(c => c.on !== false);
  if (!picked.length) return showError("Select at least one client.");
  const fd = new FormData(e.target);
  fd.delete("clients");
  if (picked.length < CLIENTS.length) picked.forEach(c => fd.append("clients", c.id));
  showError("");
  const r = await fetch("/api/run", { method: "POST", body: fd });
  if (!r.ok) return showError((await r.json()).error);
  $("#run").disabled = true;
  poll();
});
async function poll() {
  const s = await (await fetch("/api/status")).json();
  $("#progress").hidden = s.state !== "running";
  if (s.state === "running") {
    const pct = s.total ? Math.round(s.done / s.total * 100) : 3;
    $("#bar").style.width = pct + "%";
    $("#ptext").textContent = s.total ? `Checked ${s.done} of ${s.total} channels` : "Starting";
    return setTimeout(poll, 600);
  }
  $("#run").disabled = false;
  if (s.state === "error") return showError(s.error);
  if (s.state === "done") await loadReports(s.result.file);
}
function showError(msg) { const e = $("#error"); e.textContent = msg; e.hidden = !msg; }

/* ---- reports ---- */
async function loadReports(select) {
  const list = await (await fetch("/api/reports")).json();
  $("#past").innerHTML = list.map(r =>
    `<option value="${esc(r.file)}">${dmy(r.previous)} to ${dmy(r.current)}${r.demo ? " (demo)" : ""}</option>`).join("") || "<option>None yet</option>";
  const pick = select || (list[0] && list[0].file);
  if (pick) { $("#past").value = pick; openReport(pick); }
}
$("#past").addEventListener("change", e => openReport(e.target.value));
async function openReport(file) {
  const r = await fetch("/api/report/" + encodeURIComponent(file));
  if (!r.ok) return;
  report = await r.json();
  report.file = file;
  $("#dl").href = "/download/" + encodeURIComponent(file);
  $("#dl").hidden = false;
  $("#title").textContent = `${dmy(report.previous)} to ${dmy(report.current)}`;
  $("#sub").textContent = `Created ${report.generated.replace("T", " ")}${report.demo ? " with demo data" : ""}`;
  render();
}

/* ---- table ---- */
$("#q").addEventListener("input", e => { query = e.target.value.toLowerCase(); render(); });
$("#seg").addEventListener("click", e => { if (e.target.dataset.f) setFilter(e.target.dataset.f); });
function setFilter(f) {
  filter = f;
  document.querySelectorAll("#seg button").forEach(b => b.classList.toggle("on", b.dataset.f === f));
  render();
}
function render() {
  if (!report) return;
  const sm = report.summary, t = sm.total || 1;
  $("#summary").hidden = $("#tools").hidden = false;
  $("#split").innerHTML = [["ok", sm.ok], ["gaps", sm.gaps], ["bad", sm.bad]]
    .map(([k, n]) => `<span class="seg-${k}" style="width:${n / t * 100}%"></span>`).join("");
  $("#legend").innerHTML =
    `<span><b>${sm.ok}</b>complete</span><span class="r-red"><b>${sm.gaps}</b>\u221a with gaps</span><span class="s-bad"><b>${sm.bad}</b>no or old data</span>`;

  const rows = report.rows.filter(r => (filter === "all" || cat(r) === filter) &&
    (!query || (r.client + r.meter + r.channel + r.reason).toLowerCase().includes(query)));
  $("#tbl").hidden = !rows.length;
  $("#empty").hidden = !!rows.length;
  $("#empty").textContent = "No channel matches this filter.";

  const groups = new Map();
  rows.forEach(r => { if (!groups.has(r.client_id)) groups.set(r.client_id, []); groups.get(r.client_id).push(r); });
  let html = "<thead>" + $("#tbl thead").innerHTML + "</thead>";
  groups.forEach((list, id) => {
    const bad = list.filter(r => r.cur === "\u00d7").length;
    html += `<tbody class="${shut.has(id) ? "shut" : ""}"><tr class="client" data-id="${id}"><th colspan="5">${esc(list[0].client)}<small>${list.length} channels${bad ? `, ${bad} with no or old data` : ""}</small></th></tr>`;
    let lastMeter = null;
    list.forEach(r => {
      const sc = v => v === "\u221a" ? "s-ok" : v === "\u00d7" ? "s-bad" : "s-none";
      html += `<tr><td class="meter">${r.meter === lastMeter ? "" : esc(r.meter)}</td><td>${esc(r.channel)}</td>` +
        `<td class="c ${sc(r.prev)}">${esc(r.prev)}</td><td class="c ${sc(r.cur)}">${esc(r.cur)}</td>` +
        `<td class="${r.red ? "r-red" : ""}">${esc(r.reason)}</td></tr>`;
      lastMeter = r.meter;
    });
    html += "</tbody>";
  });
  $("#tbl").innerHTML = html;
}
$("#tbl").addEventListener("click", e => {
  const row = e.target.closest("tr.client");
  if (!row) return;
  const id = Number(row.dataset.id);
  shut.has(id) ? shut.delete(id) : shut.add(id);
  row.parentElement.classList.toggle("shut");
});

/* ---- start ---- */
drawClients();
loadReports();
fetch("/api/status").then(r => r.json()).then(s => { if (s.state === "running") { $("#run").disabled = true; poll(); } });
