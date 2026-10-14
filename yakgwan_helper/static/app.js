// 약관 맞춤 요약 - 화면 동작
const $ = (s) => document.querySelector(s);
let docId = null;
let busy = false;
let password = "";
try { password = localStorage.getItem("appPassword") || ""; } catch {}

// 비밀번호가 걸린 서버면 요청마다 붙여 보냄. 틀리면 다시 물어봄.
async function api(url, opts = {}) {
  opts.headers = { ...(opts.headers || {}), "X-App-Password": password };
  let r = await fetch(url, opts);
  if (r.status === 401) {
    password = prompt("이 사이트의 비밀번호를 입력하세요") || "";
    try { localStorage.setItem("appPassword", password); } catch {}
    opts.headers["X-App-Password"] = password;
    r = await fetch(url, opts);
  }
  return r;
}

// ---------------------------------------------------------------
// 간단한 마크다운 → HTML (외부 라이브러리 없이, HTML은 모두 이스케이프)
// ---------------------------------------------------------------
function esc(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function inline(s) {
  return esc(s)
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, "$1<em>$2</em>")
    .replace(/\[(p\.\s?\d+(?:\s?[-~,]\s?(?:p\.\s?)?\d+)*)\]/g, '<span class="cite">[$1]</span>');
}
function renderMarkdown(md) {
  const lines = md.replace(/\r/g, "").split("\n");
  const out = [];
  let i = 0;
  const isTableSep = (l) => /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(l);
  const cells = (l) => l.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());

  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { i++; continue; }

    let m;
    if ((m = line.match(/^(#{1,4})\s+(.*)/))) {
      const lv = Math.min(4, m[1].length + 1);   // # → h2
      out.push(`<h${lv}>${inline(m[2])}</h${lv}>`); i++; continue;
    }
    if (/^\s*(-{3,}|\*{3,})\s*$/.test(line)) { out.push("<hr>"); i++; continue; }
    if (line.includes("|") && i + 1 < lines.length && isTableSep(lines[i + 1])) {
      const head = cells(line);
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].includes("|") && lines[i].trim()) rows.push(cells(lines[i++]));
      out.push("<table><thead><tr>" + head.map((c) => `<th>${inline(c)}</th>`).join("") +
        "</tr></thead><tbody>" + rows.map((r) => "<tr>" + r.map((c) => `<td>${inline(c)}</td>`).join("") + "</tr>").join("") +
        "</tbody></table>");
      continue;
    }
    if (/^\s*>/.test(line)) {
      const buf = [];
      while (i < lines.length && /^\s*>/.test(lines[i])) buf.push(lines[i++].replace(/^\s*>\s?/, ""));
      out.push(`<blockquote>${renderMarkdown(buf.join("\n"))}</blockquote>`);
      continue;
    }
    if (/^\s*([-*+]|\d+[.)])\s+/.test(line)) {
      const ordered = /^\s*\d+[.)]/.test(line);
      const items = [];
      while (i < lines.length && /^\s*([-*+]|\d+[.)])\s+/.test(lines[i])) {
        const indent = lines[i].match(/^\s*/)[0].length;
        const text = lines[i].replace(/^\s*([-*+]|\d+[.)])\s+/, "").replace(/^\[ \]\s*/, "☐ ").replace(/^\[x\]\s*/i, "☑ ");
        items.push(`<li${indent >= 2 ? ' style="margin-left:1.2em"' : ""}>${inline(text)}</li>`);
        i++;
      }
      out.push(ordered ? `<ol>${items.join("")}</ol>` : `<ul>${items.join("")}</ul>`);
      continue;
    }
    const buf = [];
    while (i < lines.length && lines[i].trim() && !/^(#{1,4}\s|\s*>|\s*([-*+]|\d+[.)])\s+)/.test(lines[i]) &&
           !(lines[i].includes("|") && i + 1 < lines.length && isTableSep(lines[i + 1]))) {
      buf.push(inline(lines[i++]));
    }
    out.push(`<p>${buf.join("<br>")}</p>`);
  }
  return out.join("\n");
}

// ---------------------------------------------------------------
// 1. 업로드
// ---------------------------------------------------------------
const drop = $("#drop");
$("#file").addEventListener("change", (e) => e.target.files[0] && upload(e.target.files[0]));
drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault(); drop.classList.remove("over");
  const f = e.dataTransfer.files[0];
  if (f) upload(f);
});

async function upload(file) {
  const info = $("#docInfo");
  info.hidden = false;
  info.textContent = `"${file.name}" 읽는 중… (수백 쪽이면 몇십 초 걸릴 수 있어요)`;
  docId = null;
  updateButton();
  const fd = new FormData();
  fd.append("file", file);
  try {
    const r = await api("/api/upload", { method: "POST", body: fd });
    const j = await r.json();
    if (!r.ok) throw new Error(j.detail || "업로드 실패");
    docId = j.doc_id;
    let html = `<b>✔ ${esc(j.filename)}</b> · ${j.page_count}쪽`;
    if (j.title) html += `<br>추정 상품명: ${esc(j.title)}`;
    if (j.tokens) html += ` · 약 ${Math.round(j.tokens / 1000).toLocaleString()}천 토큰`;
    if (j.scanned) html += `<br>⚠️ 글자를 추출할 수 없는 스캔본이라 PDF 이미지를 그대로 분석합니다 (조금 더 느릴 수 있어요).`;
    if (j.will_trim) html += `<br>ℹ️ 약관이 매우 길어(약 ${Math.round(j.tokens / 1000)}천 토큰) 관련성이 높은 페이지 위주로 분석합니다.`;
    info.innerHTML = html;
  } catch (err) {
    info.innerHTML = `❌ ${esc(err.message)}`;
  }
  updateButton();
}

// ---------------------------------------------------------------
// 2~3. 방식 / 내 정보
// ---------------------------------------------------------------
document.querySelectorAll('input[name=mode]').forEach((r) => r.addEventListener("change", () => {
  $("#profileCard").hidden = mode() !== "personal";
}));
document.querySelectorAll('input[name=surgery]').forEach((r) => r.addEventListener("change", () => {
  $("#surgeryDetail").hidden = r.value !== "있음" || !r.checked;
}));
$("select[name=gender]").addEventListener("change", (e) => {
  $("#pregnantRow").hidden = e.target.value === "남성";
});

function mode() { return document.querySelector("input[name=mode]:checked").value; }

function readProfile() {
  const form = $("#profile");
  const p = {};
  form.querySelectorAll("input[name]:not([type=radio]):not([type=checkbox]), select[name], textarea[name]").forEach((el) => {
    if (el.value.trim()) p[el.name] = el.name === "age" ? Number(el.value) : el.value.trim();
  });
  form.querySelectorAll("input[type=radio]:checked").forEach((el) => (p[el.name] = el.value));
  form.querySelectorAll(".chips").forEach((box) => {
    const vals = [...box.querySelectorAll("input:checked")].map((x) => x.value);
    if (vals.length) p[box.dataset.name] = vals;
  });
  if (p.gender === "남성") delete p.pregnant;
  if (p.surgery !== "있음") delete p.surgery_detail;
  return p;
}

function updateButton() {
  $("#run").disabled = !docId || busy;
  $("#askBtn").disabled = !docId || busy;
}

// ---------------------------------------------------------------
// 요약 요청 (스트리밍으로 받아서 바로 보여줌)
// ---------------------------------------------------------------
async function streamTo(body, el, onDone) {
  busy = true; updateButton();
  el.classList.add("cursor");
  let text = "";
  let last = 0;
  try {
    const r = await api("/api/summarize", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!r.ok) {
      const j = await r.json().catch(() => ({}));
      throw new Error(j.detail || `오류 (${r.status})`);
    }
    const reader = r.body.getReader();
    const dec = new TextDecoder();
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      text += dec.decode(value, { stream: true });
      if (Date.now() - last > 120) { el.innerHTML = renderMarkdown(text); last = Date.now(); }
    }
    text += dec.decode();
    el.innerHTML = renderMarkdown(text);
  } catch (err) {
    el.innerHTML = renderMarkdown(text + `\n\n> ❌ ${err.message}`);
  } finally {
    el.classList.remove("cursor");
    busy = false; updateButton();
    onDone && onDone(text);
  }
  return text;
}

let lastResult = "";
$("#run").addEventListener("click", async () => {
  if (!docId) return;
  const m = mode();
  const body = { doc_id: docId, mode: m };
  if (m === "personal") body.profile = readProfile();
  $("#resultCard").hidden = false;
  $("#answers").innerHTML = "";
  $("#status").textContent = "약관을 읽고 있어요… 긴 약관은 1~3분 정도 걸립니다.";
  $("#result").innerHTML = "";
  $("#resultCard").scrollIntoView({ behavior: "smooth" });
  const t0 = Date.now();
  lastResult = await streamTo(body, $("#result"));
  $("#status").textContent = `완료 (${Math.round((Date.now() - t0) / 1000)}초)`;
});

$("#askBtn").addEventListener("click", ask);
$("#question").addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.isComposing) ask(); });
async function ask() {
  const q = $("#question").value.trim();
  if (!q || !docId || busy) return;
  const box = document.createElement("div");
  box.className = "answer";
  box.innerHTML = `<div class="q">Q. ${esc(q)}</div><div class="md"></div>`;
  $("#answers").prepend(box);
  $("#question").value = "";
  const body = { doc_id: docId, mode: "question", question: q };
  if (mode() === "personal") body.profile = readProfile();
  await streamTo(body, box.querySelector(".md"));
}

$("#copy").addEventListener("click", async () => {
  try { await navigator.clipboard.writeText(lastResult); $("#copy").textContent = "복사됨"; }
  catch { $("#copy").textContent = "복사 실패"; }
  setTimeout(() => ($("#copy").textContent = "복사"), 1500);
});
$("#print").addEventListener("click", () => window.print());

// 서버 상태 확인
fetch("/api/health").then((r) => r.json()).then((j) => { $("#keyWarn").hidden = j.api_key; }).catch(() => {});
updateButton();
