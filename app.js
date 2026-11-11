import { firebaseConfig, ADMIN_ID, EMAIL_DOMAIN } from "./firebase-config.js";
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged, createUserWithEmailAndPassword, updatePassword } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { getFirestore, doc, getDoc, getDocs, setDoc, updateDoc, deleteDoc, collection, query, where, writeBatch, serverTimestamp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const secondary = initializeApp(firebaseConfig, "secondary");
const secAuth = getAuth(secondary);

const DIFFS = ["Easy", "Medium", "Hard"];
const DIFF_AR = { Easy: "سهل", Medium: "متوسط", Hard: "صعب" };
const PASS = 70;
const emailOf = id => `${String(id).trim().toLowerCase()}@${EMAIL_DOMAIN}`;

const S = { user: null, tab: "results", exam: null, loginErr: "", selExam: "",
  users: [], questions: {}, exams: [], attempts: {}, nurseList: [] };

/* ---------- helpers ---------- */
const $ = s => document.querySelector(s);
const $$ = s => document.querySelectorAll(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const pct = (a, b) => b ? Math.round(a * 100 / b) : 0;
const tsDate = t => t && t.toDate ? t.toDate() : null;
const fmtDate = t => { const d = tsDate(t); return d ? d.toLocaleString("ar-EG", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) : "—"; };
function shuffle(a) { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; }
function toast(m, ms = 2600) { const d = document.createElement("div"); d.className = "toast"; d.textContent = m; document.body.appendChild(d); setTimeout(() => d.remove(), ms); }
function busy(m) { $("#app").innerHTML = `<section class="card"><p class="muted" style="margin:0">${esc(m)}</p></section>`; }
function modal(html) { $("#layer").innerHTML = `<div class="modal"><div class="card">${html}</div></div>`; }
const closeModal = () => $("#layer").innerHTML = "";
function confirmBox(msg, onYes) {
  modal(`<p>${esc(msg)}</p><div class="row"><button class="btn primary" id="cy">تأكيد</button><button class="btn" id="cn">إلغاء</button></div>`);
  $("#cy").onclick = () => { closeModal(); onYes(); }; $("#cn").onclick = closeModal;
}
function errText(e) {
  const c = e && e.code || "";
  if (c.includes("invalid-credential") || c.includes("wrong-password") || c.includes("user-not-found") || c.includes("invalid-email")) return "الرقم الوظيفي أو كلمة المرور غير صحيحة";
  if (c.includes("too-many-requests")) return "محاولات كثيرة. انتظر قليلاً ثم حاول مرة أخرى";
  if (c.includes("network")) return "لا يوجد اتصال بالإنترنت";
  if (c.includes("permission-denied")) return "غير مسموح بهذه العملية";
  if (c.includes("weak-password")) return "كلمة المرور يجب ألا تقل عن 6 أحرف";
  return "حدث خطأ: " + (c || (e && e.message) || "");
}
function download(name, text) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob(["﻿" + text], { type: "text/csv;charset=utf-8" }));
  a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
async function readSheet(file, prefer) {
  const wb = XLSX.read(await file.arrayBuffer());
  const name = wb.SheetNames.includes(prefer) ? prefer : wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: false });
}
async function commitBatches(ops) { // ops: [ (batch)=>void ]
  for (let i = 0; i < ops.length; i += 400) {
    const b = writeBatch(db); ops.slice(i, i + 400).forEach(f => f(b)); await b.commit();
  }
}

/* ---------- auth ---------- */
onAuthStateChanged(auth, async u => {
  if (!u) { S.user = null; render(); return; }
  busy("جاري التحميل…");
  try {
    const ref = doc(db, "users", u.uid);
    let snap = await getDoc(ref);
    if (!snap.exists() && u.email === emailOf(ADMIN_ID)) {
      await setDoc(ref, { empId: String(ADMIN_ID), name: "مسؤول التدريب", unit: "Nursing Education", role: "admin", active: true });
      snap = await getDoc(ref);
    }
    if (!snap.exists() || snap.data().active === false) {
      S.loginErr = "هذا الحساب غير مفعّل. راجع مسؤول التدريب."; await signOut(auth); return;
    }
    S.user = { uid: u.uid, ...snap.data() };
    if (S.user.role === "admin") await loadAdmin(); else await loadNurse();
    render();
  } catch (e) { S.loginErr = errText(e); await signOut(auth); }
});

function render() {
  $("#logout").hidden = !S.user;
  $("#pwbtn").hidden = !S.user || !!S.exam;
  $("#who").textContent = S.user ? `${S.user.name} · ${S.user.empId}` : "تدريب وتقييم ومتابعة";
  if (!S.user) return renderLogin();
  if (S.exam) return renderExam();
  S.user.role === "admin" ? renderAdmin() : renderNurse();
}

function renderLogin() {
  $("#app").innerHTML = `
  <section class="card">
    <h1>تسجيل الدخول</h1>
    <form id="lf" style="display:grid;gap:12px">
      <label for="uid">الرقم الوظيفي<input id="uid" inputmode="numeric" autocomplete="username" required></label>
      <label for="pw">كلمة المرور<input id="pw" type="password" autocomplete="current-password" required></label>
      <p class="small" id="lerr" style="color:var(--bad);margin:0">${esc(S.loginErr)}</p>
      <button class="btn primary block" id="lb">دخول</button>
    </form>
  </section>`;
  $("#lf").onsubmit = async e => {
    e.preventDefault(); $("#lb").disabled = true; S.loginErr = "";
    try { await signInWithEmailAndPassword(auth, emailOf($("#uid").value), $("#pw").value); }
    catch (err) { $("#lerr").textContent = errText(err); $("#lb").disabled = false; }
  };
}
$("#logout").onclick = async () => { if (S.exam) { toast("أنهِ الامتحان أولاً"); return; } S.loginErr = ""; await signOut(auth); };
$("#pwbtn").onclick = () => {
  modal(`<h3>تغيير كلمة المرور</h3><label for="np1">كلمة المرور الجديدة (6 أحرف على الأقل)<input id="np1" type="password" autocomplete="new-password"></label>
   <p class="small" id="perr" style="color:var(--bad);margin:0"></p><div class="row"><button class="btn primary" id="ps">حفظ</button><button class="btn" id="pc">إلغاء</button></div>`);
  $("#pc").onclick = closeModal;
  $("#ps").onclick = async () => {
    const v = $("#np1").value; if (v.length < 6) { $("#perr").textContent = "6 أحرف على الأقل"; return; }
    try { await updatePassword(auth.currentUser, v); closeModal(); toast("تم تغيير كلمة المرور"); }
    catch (e) { $("#perr").textContent = e.code && e.code.includes("recent-login") ? "سجّل خروج ثم دخول مرة أخرى، ثم غيّر كلمة المرور" : errText(e); }
  };
};

/* ---------- nurse ---------- */
async function loadNurse() {
  const ex = await getDocs(collection(db, "exams"));
  const list = [];
  for (const d of ex.docs) {
    const e = { id: d.id, ...d.data() };
    if (!e.open && !e.released) continue;
    try {
      const a = await getDoc(doc(db, "attempts", `${e.id}_${S.user.empId}`));
      if (a.exists()) list.push({ e, a: { id: a.id, ...a.data() } });
    } catch (err) { /* not assigned to this nurse */ }
  }
  S.nurseList = list.sort((x, y) => String(y.e.created).localeCompare(String(x.e.created)));
}
function renderNurse() {
  const L = S.nurseList;
  $("#app").innerHTML = `
  <section class="card"><h1>أهلاً ${esc(S.user.name.split(" ")[0])}</h1><p class="muted small" style="margin:0">${esc(S.user.unit)}</p></section>
  <section class="card"><div class="row"><h2 class="grow">الامتحانات</h2><button class="btn sm" id="rf">تحديث</button></div>
  ${L.length ? L.map(({ e, a }, i) => {
    let right;
    if (a.status === "assigned") right = e.open ? `<button class="btn primary sm" data-start="${i}">ابدأ</button>` : `<span class="pill">مغلق</span>`;
    else if (a.status === "in-progress") right = `<span class="pill warn">انقطع، راجع المسؤول</span>`;
    else if (e.released && a.score != null) right = `<span class="pill ${pct(a.score, a.total) >= PASS ? "good" : "bad"}">${a.score}/${a.total}</span>${a.review ? ` <button class="btn sm" data-review="${i}">المراجعة</button>` : ""}`;
    else right = `<span class="pill">تم التسليم</span>`;
    return `<div class="exam-item"><div class="grow"><b>${esc(e.title)}</b><div class="muted small">${a.total} سؤال · ${e.sec} ثانية للسؤال</div></div>${right}</div>`;
  }).join("") : `<p class="muted">لا توجد امتحانات متاحة لك الآن.</p>`}
  </section>`;
  $("#rf").onclick = async () => { busy("جاري التحديث…"); await loadNurse(); render(); };
  $$("[data-start]").forEach(b => b.onclick = () => preExam(S.nurseList[+b.dataset.start]));
  $$("[data-review]").forEach(b => b.onclick = () => showReview(S.nurseList[+b.dataset.review].a, S.nurseList[+b.dataset.review].e, render));
}
function preExam({ e, a }) {
  $("#app").innerHTML = `<section class="card"><h1>${esc(e.title)}</h1>
   <ul class="rules">
    <li>عدد الأسئلة: ${a.total}، ولكل سؤال ${e.sec} ثانية.</li>
    <li>سؤال واحد في الشاشة، ولا يمكن الرجوع للسؤال السابق.</li>
    <li>عند انتهاء الوقت ينتقل الامتحان للسؤال التالي تلقائياً.</li>
    <li>لك محاولة واحدة فقط. الخروج من الامتحان أو قفل الصفحة يُحسب محاولة.</li>
    <li>الخروج من التطبيق أثناء الامتحان يُسجَّل ويظهر للمسؤول.</li>
    <li>تأكد أن الإنترنت متصل قبل البدء.</li>
   </ul>
   <div class="row"><button class="btn primary grow" id="go">ابدأ الامتحان الآن</button><button class="btn" id="back">رجوع</button></div></section>`;
  $("#back").onclick = render;
  $("#go").onclick = async () => {
    $("#go").disabled = true;
    const ref = doc(db, "attempts", a.id);
    try { await updateDoc(ref, { status: "in-progress", started: serverTimestamp() }); }
    catch (err) { toast(errText(err)); $("#go").disabled = false; return; }
    S.exam = { ref, a, e, i: 0, answers: [], sel: null, left: e.sec, tabs: 0, timer: null, sending: false };
    render();
  };
}
function renderExam() {
  const X = S.exam, it = X.a.items[X.i];
  const wm = Array(40).fill(`${esc(S.user.empId)} · ${esc(S.user.name)}`).map(t => `<span>${t}</span>`).join("");
  $("#app").innerHTML = `<div class="exam"><div class="wm" aria-hidden="true">${wm}</div>
   <section class="card">
    <div class="qhead"><span class="small muted">سؤال ${X.i + 1} من ${X.a.items.length}</span><div class="prog"><i style="width:${pct(X.i, X.a.items.length)}%"></i></div><span class="timer${X.left <= 10 ? " low" : ""}" id="tm">${X.left}</span></div>
    ${X.tabs ? `<div class="warnbar">تم تسجيل خروجك من الامتحان ${X.tabs} مرة. هذا يظهر للمسؤول.</div>` : ""}
    <div class="stem">${esc(it.stem)}</div>
    <div class="opts">${it.opts.map((o, i) => `<button class="opt" data-i="${i}" aria-pressed="${X.sel === i}"><span class="k">${"abcd"[i]}</span><span>${esc(o)}</span></button>`).join("")}</div>
    <button class="btn primary block" id="nx" ${X.sel === null ? "disabled" : ""}>${X.i + 1 === X.a.items.length ? "إنهاء وتسليم" : "تأكيد والتالي"}</button>
   </section></div>`;
  $$(".opt").forEach(b => b.onclick = () => { X.sel = +b.dataset.i; $$(".opt").forEach(o => o.setAttribute("aria-pressed", o === b)); $("#nx").disabled = false; });
  $("#nx").onclick = () => commit(false);
  clearInterval(X.timer);
  X.timer = setInterval(() => {
    X.left--; const t = $("#tm"); if (t) { t.textContent = X.left; t.classList.toggle("low", X.left <= 10); }
    if (X.left <= 0) commit(true);
  }, 1000);
}
function commit(timeout) {
  const X = S.exam; if (X.sending) return;
  clearInterval(X.timer);
  X.answers.push(timeout && X.sel === null ? -1 : X.sel);
  X.i++; X.sel = null; X.left = X.e.sec;
  if (X.i < X.a.items.length) {
    updateDoc(X.ref, { answers: X.answers.slice() }).catch(err => console.warn(err));
    renderExam(); return;
  }
  submitFinal();
}
async function submitFinal() {
  const X = S.exam; X.sending = true;
  $("#app").innerHTML = `<section class="card"><p style="margin:0">جاري تسليم الإجابات…</p></section>`;
  try {
    await updateDoc(X.ref, { answers: X.answers.slice(), status: "done", finished: serverTimestamp() });
    S.exam = null;
    $("#app").innerHTML = `<section class="card" style="text-align:center;justify-items:center">
      <span class="pill good">تم التسليم</span><h2>تم تسليم إجاباتك بنجاح</h2>
      <p class="muted small">النتيجة والمراجعة تظهران هنا بعد أن يعلنها مسؤول التدريب.</p>
      <button class="btn primary" id="home">الرئيسية</button></section>`;
    $("#home").onclick = async () => { busy("جاري التحميل…"); await loadNurse(); render(); };
  } catch (err) {
    X.sending = false;
    $("#app").innerHTML = `<section class="card"><p class="warnbar">${esc(errText(err))}. إجاباتك محفوظة على الجهاز، لا تغلق الصفحة.</p><button class="btn primary" id="retry">إعادة المحاولة</button></section>`;
    $("#retry").onclick = submitFinal;
  }
}
document.addEventListener("visibilitychange", () => {
  const X = S.exam; if (!X || X.sending) return;
  if (document.hidden) { X.tabs++; updateDoc(X.ref, { tabs: X.tabs }).catch(() => {}); }
  else renderExam();
});
window.addEventListener("beforeunload", e => { if (S.exam) { e.preventDefault(); e.returnValue = ""; } });
document.addEventListener("contextmenu", e => { if (S.exam) e.preventDefault(); });
document.addEventListener("copy", e => { if (S.exam) e.preventDefault(); });

function showReview(a, e, back, adminView) {
  const rv = a.review;
  $("#app").innerHTML = `<section class="card"><div class="row"><h2 class="grow">${esc(e.title)}</h2><button class="btn sm" id="bk">رجوع</button></div>
   <p class="muted small" style="margin:0">${esc(a.name)} · ${a.score ?? "—"}/${a.total}</p>
   ${a.items.map((it, i) => {
     const ch = a.answers?.[i] ?? -1;
     const c = adminView ? it.order.indexOf("ABCD".indexOf(S.questions[it.qid]?.correct)) : rv?.[i]?.c;
     const rat = adminView ? S.questions[it.qid]?.rat : rv?.[i]?.rat;
     const ok = ch === c;
     return `<div class="review ${ok ? "ok" : "no"}"><b>${i + 1}. ${esc(it.stem)}</b>
      <span>Answer: ${ch >= 0 ? esc(it.opts[ch]) : "— (time out)"} ${ok ? "✓" : "✗"}</span>
      ${ok || c == null || c < 0 ? "" : `<span style="color:var(--good)">Correct: ${esc(it.opts[c])}</span>`}
      ${rat ? `<span class="muted small">${esc(rat)}</span>` : ""}${adminView ? `<span class="muted small">${esc(it.qid)}</span>` : ""}</div>`;
   }).join("")}</section>`;
  $("#bk").onclick = back;
}

/* ---------- admin data ---------- */
async function loadAdmin() {
  const [u, q, e] = await Promise.all([getDocs(collection(db, "users")), getDocs(collection(db, "questions")), getDocs(collection(db, "exams"))]);
  S.users = u.docs.map(d => ({ uid: d.id, ...d.data() }));
  S.questions = Object.fromEntries(q.docs.map(d => [d.id, d.data()]));
  S.exams = e.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => String(b.created).localeCompare(String(a.created)));
  S.attempts = {};
  if (!S.selExam && S.exams[0]) S.selExam = S.exams[0].id;
}
const nurses = () => S.users.filter(u => u.role === "nurse");
const examById = id => S.exams.find(e => e.id === id);
async function loadAttempts(eid, force) {
  if (!eid) return [];
  if (S.attempts[eid] && !force) return S.attempts[eid];
  const snap = await getDocs(query(collection(db, "attempts"), where("examId", "==", eid)));
  S.attempts[eid] = snap.docs.map(d => ({ id: d.id, ...d.data() }));
  await gradeAttempts(eid);
  return S.attempts[eid];
}
function reviewFor(a) {
  return a.items.map(it => { const q = S.questions[it.qid]; return { c: q ? it.order.indexOf("ABCD".indexOf(q.correct)) : -1, rat: q ? q.rat : "" }; });
}
async function gradeAttempts(eid) {
  const e = examById(eid), ops = [];
  for (const a of S.attempts[eid]) {
    if (a.status !== "done") continue;
    const upd = {};
    if (a.score == null) {
      const ok = a.items.map((it, i) => { const q = S.questions[it.qid]; const ch = a.answers?.[i] ?? -1; return !!q && ch >= 0 && "ABCD"[it.order[ch]] === q.correct; });
      Object.assign(upd, { ok, score: ok.filter(Boolean).length, total: a.items.length });
    }
    if (e?.released && !a.review) upd.review = reviewFor(a);
    if (Object.keys(upd).length) { Object.assign(a, upd); ops.push(b => b.update(doc(db, "attempts", a.id), upd)); }
  }
  if (ops.length) await commitBatches(ops);
}
function drawItems(exam) {
  const all = Object.entries(S.questions).map(([id, q]) => ({ id, ...q }));
  const out = [];
  exam.picks.forEach(p => {
    shuffle(all.filter(q => q.topic === p.topic && q.diff === p.diff)).slice(0, p.n).forEach(q => {
      const order = shuffle([0, 1, 2, 3]);
      out.push({ qid: q.id, stem: q.stem, opts: order.map(k => q.opts[k]), order });
    });
  });
  return shuffle(out);
}
const targetsOf = e => nurses().filter(u => u.active !== false && (!e.units?.length || e.units.includes(u.unit)));
async function distribute(e, onlyEmp) {
  const existing = new Set((await loadAttempts(e.id)).map(a => a.empId));
  const list = targetsOf(e).filter(u => (onlyEmp ? u.empId === onlyEmp : true) && !existing.has(u.empId));
  const created = [];
  const ops = list.map(u => {
    const items = drawItems(e);
    const a = { examId: e.id, empId: u.empId, name: u.name, unit: u.unit, items, answers: [], status: "assigned", tabs: 0, started: null, finished: null, score: null, total: items.length, ok: null, review: null };
    created.push({ id: `${e.id}_${u.empId}`, ...a });
    return b => b.set(doc(db, "attempts", `${e.id}_${u.empId}`), a);
  });
  await commitBatches(ops);
  S.attempts[e.id].push(...created);
  return list.length;
}

/* ---------- admin UI ---------- */
function renderAdmin() {
  const tabs = [["results", "النتائج"], ["items", "تحليل الأسئلة"], ["exams", "الامتحانات"], ["users", "الممرضين"], ["bank", "بنك الأسئلة"]];
  $("#app").innerHTML = `<nav class="tabs" role="tablist">${tabs.map(([k, l]) => `<button role="tab" aria-selected="${S.tab === k}" data-tab="${k}">${l}</button>`).join("")}</nav><div id="pane" style="display:grid;gap:14px"></div>`;
  $$("[data-tab]").forEach(b => b.onclick = () => { S.tab = b.dataset.tab; renderAdmin(); });
  ({ results: aResults, items: aItems, exams: aExams, users: aUsers, bank: aBank })[S.tab]();
}
function examPicker() {
  return S.exams.length ? `<label for="fe">الامتحان<select id="fe">${S.exams.map(e => `<option value="${e.id}" ${S.selExam === e.id ? "selected" : ""}>${esc(e.title)}</option>`).join("")}</select></label>` : `<p class="muted" style="margin:0">لا توجد امتحانات بعد. أنشئ امتحاناً من تبويب "الامتحانات".</p>`;
}
function bindPicker() { const f = $("#fe"); if (f) f.onchange = ev => { S.selExam = ev.target.value; renderAdmin(); }; }
const durSec = a => { const s = tsDate(a.started), f = tsDate(a.finished); return s && f ? Math.round((f - s) / 1000) : null; };

async function aResults() {
  const e = examById(S.selExam);
  $("#pane").innerHTML = `<section class="card">${examPicker()}</section><div id="rbody"></div>`;
  bindPicker(); if (!e) return;
  $("#rbody").innerHTML = `<p class="muted">جاري تحميل النتائج…</p>`;
  let A; try { A = await loadAttempts(e.id); } catch (err) { $("#rbody").innerHTML = `<p class="warnbar">${esc(errText(err))}</p>`; return; }
  const unit = S.unitFilter || "";
  const units = [...new Set(A.map(a => a.unit))].sort();
  const rs = A.filter(a => !unit || a.unit === unit);
  const done = rs.filter(a => a.status === "done" && a.score != null);
  const avg = done.length ? Math.round(done.reduce((s, a) => s + pct(a.score, a.total), 0) / done.length) : 0;
  const passN = done.filter(a => pct(a.score, a.total) >= PASS).length;
  const limit = s => s > e.sec * (rs[0]?.total || 0) + 120;
  const flagged = rs.filter(a => a.tabs > 0 || a.status === "in-progress" || (durSec(a) != null && limit(durSec(a)))).length;
  const missing = targetsOf(e).filter(u => !A.some(a => a.empId === u.empId));
  const stat = a => a.status === "done" ? (a.score != null ? `${a.score}/${a.total}` : "—") : a.status === "in-progress" ? `<span class="pill warn">انقطع</span>` : `<span class="pill">لم يبدأ</span>`;
  $("#rbody").innerHTML = `
  <section class="card"><div class="row">
    <label class="grow" for="fu">الوحدة<select id="fu"><option value="">كل الوحدات</option>${units.map(u => `<option ${unit === u ? "selected" : ""}>${esc(u)}</option>`).join("")}</select></label>
    <span class="pill ${e.open ? "good" : ""}">${e.open ? "مفتوح" : "مغلق"}</span><span class="pill ${e.released ? "good" : "warn"}">${e.released ? "النتائج معلنة" : "النتائج غير معلنة"}</span>
  </div></section>
  <div class="tiles">
   <div class="tile"><div class="n">${done.length}<span class="muted" style="font-size:15px"> / ${rs.length}</span></div><div class="l">سلّموا الامتحان</div></div>
   <div class="tile"><div class="n">${avg}%</div><div class="l">متوسط الدرجات</div></div>
   <div class="tile"><div class="n" style="color:var(--good)">${pct(passN, done.length)}%</div><div class="l">نسبة النجاح (${PASS}%)</div></div>
   <div class="tile"><div class="n" style="color:${flagged ? "var(--warn)" : "inherit"}">${flagged}</div><div class="l">محاولات عليها ملاحظة</div></div>
  </div>
  ${missing.length ? `<section class="card"><div class="row"><p class="grow" style="margin:0">${missing.length} ممرض مستهدف لم يُوزَّع عليهم الامتحان بعد (أُضيفوا بعد إنشائه).</p><button class="btn primary sm" id="dist">توزيع عليهم</button></div></section>` : ""}
  <section class="card"><div class="row"><h2 class="grow">النتائج</h2><button class="btn sm" id="rl">تحديث</button><button class="btn sm" id="xl">تنزيل Excel</button></div>
   <div class="tablewrap"><table><thead><tr><th>الاسم</th><th>الرقم</th><th>الوحدة</th><th>الدرجة</th><th>%</th><th>خروج</th><th>المدة</th><th>التاريخ</th><th></th></tr></thead><tbody>
   ${rs.sort((x, y) => (tsDate(y.finished || y.started)?.getTime() || 0) - (tsDate(x.finished || x.started)?.getTime() || 0)).map(a => {
     const p = pct(a.score, a.total), d = durSec(a);
     return `<tr><td>${esc(a.name)}</td><td>${esc(a.empId)}</td><td>${esc(a.unit)}</td><td>${stat(a)}</td>
      <td>${a.status === "done" && a.score != null ? `<span class="pill ${p >= PASS ? "good" : "bad"}">${p}%</span>` : "—"}</td>
      <td>${a.tabs ? `<span class="pill warn">${a.tabs}</span>` : "0"}</td>
      <td>${d == null ? "—" : `<span class="${limit(d) ? "pill warn" : ""}">${Math.floor(d / 60)}:${String(d % 60).padStart(2, "0")}</span>`}</td>
      <td>${fmtDate(a.finished || a.started)}</td>
      <td>${a.status === "done" ? `<button class="btn sm" data-rv="${esc(a.id)}">التفاصيل</button> ` : ""}${a.status !== "assigned" ? `<button class="btn sm danger" data-reset="${esc(a.id)}">إعادة السماح</button>` : ""}</td></tr>`;
   }).join("") || `<tr><td colspan="9" class="muted">لا توجد محاولات.</td></tr>`}
   </tbody></table></div>
   <p class="small muted" style="margin:0">"المدة" تظهر بلون مميز إذا زادت عن الوقت المسموح بأكثر من دقيقتين، وقد تعني أن الممرض تلاعب بالتوقيت.</p>
  </section>`;
  $("#fu").onchange = ev => { S.unitFilter = ev.target.value; aResults(); };
  $("#rl").onclick = async () => { await loadAttempts(e.id, true); aResults(); };
  if ($("#dist")) $("#dist").onclick = async () => { $("#dist").disabled = true; try { const n = await distribute(e); toast(`تم التوزيع على ${n}`); } catch (err) { toast(errText(err)); } aResults(); };
  $$("[data-rv]").forEach(b => b.onclick = () => { const a = A.find(x => x.id === b.dataset.rv); showReview(a, e, renderAdmin, true); });
  $$("[data-reset]").forEach(b => b.onclick = () => {
    const a = A.find(x => x.id === b.dataset.reset);
    confirmBox(`حذف محاولة ${a.name} وإعطاؤه امتحاناً جديداً بأسئلة مختلفة؟`, async () => {
      try { await deleteDoc(doc(db, "attempts", a.id)); S.attempts[e.id] = A.filter(x => x.id !== a.id); await distribute(e, a.empId); toast("تم. يمكنه الدخول للامتحان الآن"); }
      catch (err) { toast(errText(err)); }
      aResults();
    });
  });
  $("#xl").onclick = () => {
    const rows = [["Name", "ID", "Unit", "Status", "Score", "Total", "Percent", "Tab switches", "Duration (sec)", "Finished"]].concat(
      rs.map(a => [a.name, a.empId, a.unit, a.status, a.score ?? "", a.total, a.score != null ? pct(a.score, a.total) : "", a.tabs, durSec(a) ?? "", tsDate(a.finished)?.toLocaleString("en-GB") || ""]));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Results");
    XLSX.writeFile(wb, `${e.title}.xlsx`);
  };
}

async function aItems() {
  const e = examById(S.selExam);
  $("#pane").innerHTML = `<section class="card">${examPicker()}<p class="small muted" style="margin:0">نسبة الإجابة الصحيحة لكل سؤال. سؤال صعب يجيبه الجميع قد يكون تسرّب، وسؤال يخطئ فيه الأغلب قد تكون صياغته غير واضحة.</p></section><div id="ibody"></div>`;
  bindPicker(); if (!e) return;
  let A; try { A = await loadAttempts(e.id); } catch (err) { $("#ibody").innerHTML = `<p class="warnbar">${esc(errText(err))}</p>`; return; }
  const st = {};
  A.filter(a => a.status === "done" && a.ok).forEach(a => a.items.forEach((it, i) => { (st[it.qid] ??= { n: 0, ok: 0 }).n++; if (a.ok[i]) st[it.qid].ok++; }));
  const rows = Object.entries(st).map(([id, s]) => ({ id, q: S.questions[id], ...s, p: pct(s.ok, s.n) })).filter(x => x.q).sort((a, b) => a.p - b.p);
  const flag = x => x.n < 5 ? `<span class="pill">عينة صغيرة</span>` : x.q.diff === "Hard" && x.p >= 90 ? `<span class="pill warn">احتمال تسريب</span>` : x.p < 30 ? `<span class="pill bad">راجع السؤال</span>` : `<span class="pill good">طبيعي</span>`;
  $("#ibody").innerHTML = `<section class="card"><div class="tablewrap"><table><thead><tr><th>السؤال</th><th>الموضوع</th><th>الصعوبة</th><th>ظهر</th><th>صحيح</th><th></th><th>الحالة</th></tr></thead><tbody>
  ${rows.map(x => `<tr><td class="ltr"><b>${esc(x.id)}</b> ${esc(x.q.stem.slice(0, 90))}${x.q.stem.length > 90 ? "…" : ""}</td><td>${esc(x.q.topic)}</td><td>${DIFF_AR[x.q.diff] || ""}</td><td>${x.n}</td><td>${x.p}%</td><td><div class="bar"><i style="width:${x.p}%"></i></div></td><td>${flag(x)}</td></tr>`).join("") || `<tr><td colspan="7" class="muted">لا توجد بيانات بعد.</td></tr>`}
  </tbody></table></div></section>`;
}

function aExams() {
  const all = Object.values(S.questions);
  const topics = [...new Set(all.map(q => q.topic))].sort();
  const avail = (t, d) => all.filter(q => q.topic === t && q.diff === d).length;
  const units = [...new Set(nurses().map(u => u.unit))].sort();
  $("#pane").innerHTML = `<section class="card"><h2>الامتحانات</h2>
   ${S.exams.map(e => `<div class="exam-item" style="flex-wrap:wrap"><div class="grow"><b>${esc(e.title)}</b><div class="muted small">${e.picks.reduce((s, p) => s + p.n, 0)} سؤال · ${e.sec} ث/سؤال · ${e.units?.length ? esc(e.units.join("، ")) : "كل الوحدات"}</div></div>
     <button class="btn sm" data-open="${e.id}">${e.open ? "إغلاق" : "فتح"}</button>
     <button class="btn sm ${e.released ? "" : "primary"}" data-rel="${e.id}">${e.released ? "إخفاء النتائج" : "إعلان النتائج"}</button></div>`).join("") || `<p class="muted">لا توجد امتحانات بعد.</p>`}
  </section>
  <section class="card"><h2>امتحان جديد</h2>
   ${all.length ? `
   <label for="et">اسم الامتحان<input id="et" placeholder="مثال: التعليم المستمر - ديسمبر"></label>
   <label for="es">ثواني لكل سؤال<input id="es" type="number" min="15" value="60"></label>
   <fieldset style="border:1px solid var(--line);border-radius:8px;padding:10px;display:flex;flex-wrap:wrap;gap:10px"><legend class="small muted">الوحدات المستهدفة (بدون اختيار = الكل)</legend>
    ${units.map((u, i) => `<label style="display:flex;gap:6px;align-items:center;color:var(--ink)" for="u${i}"><input type="checkbox" id="u${i}" data-unit="${esc(u)}" style="width:auto">${esc(u)}</label>`).join("") || `<span class="muted small">أضف ممرضين أولاً</span>`}</fieldset>
   <div class="grid3 small"><span class="muted">الموضوع</span>${DIFFS.map(d => `<span class="muted" style="text-align:center">${DIFF_AR[d]}</span>`).join("")}
    ${topics.map((t, ti) => `<span dir="ltr" style="text-align:right">${esc(t)}</span>${DIFFS.map(d => `<input type="number" min="0" max="${avail(t, d)}" value="0" data-t="${ti}" data-d="${d}" aria-label="${esc(t)} ${DIFF_AR[d]}" title="المتاح ${avail(t, d)}">`).join("")}`).join("")}
   </div>
   <p class="small muted" style="margin:0" id="esum">اختر عدد الأسئلة من كل موضوع وصعوبة. كل ممرض سيحصل على أسئلة مختلفة من نفس التوزيع.</p>
   <button class="btn primary" id="ec">إنشاء الامتحان وتوزيعه</button>` : `<p class="muted" style="margin:0">ارفع بنك الأسئلة أولاً من تبويب "بنك الأسئلة".</p>`}
  </section>`;
  $$("[data-open]").forEach(b => b.onclick = async () => { const e = examById(b.dataset.open); try { await updateDoc(doc(db, "exams", e.id), { open: !e.open }); e.open = !e.open; } catch (err) { toast(errText(err)); } aExams(); });
  $$("[data-rel]").forEach(b => b.onclick = async () => {
    const e = examById(b.dataset.rel); b.disabled = true;
    try {
      await updateDoc(doc(db, "exams", e.id), { released: !e.released }); e.released = !e.released;
      if (e.released) { await loadAttempts(e.id, true); toast("تم إعلان النتائج"); } else toast("تم إخفاء النتائج");
    } catch (err) { toast(errText(err)); }
    aExams();
  });
  if (!all.length) return;
  const upd = () => { let n = 0, bad = false; $$("[data-t]").forEach(i => { const v = +i.value || 0; n += v; if (v > +i.max) bad = true; }); $("#esum").textContent = bad ? "عدد أكبر من المتاح في البنك لأحد المربعات." : `إجمالي ${n} سؤال.`; return { n, bad }; };
  $$("[data-t]").forEach(i => i.oninput = upd);
  $("#ec").onclick = async () => {
    const { n, bad } = upd(), title = $("#et").value.trim();
    if (!title || !n || bad) { toast(!title ? "اكتب اسم الامتحان" : !n ? "اختر أسئلة" : "راجع الأعداد"); return; }
    const picks = []; $$("[data-t]").forEach(i => { if (+i.value > 0) picks.push({ topic: topics[+i.dataset.t], diff: i.dataset.d, n: +i.value }); });
    const unitsSel = [...$$("[data-unit]")].filter(c => c.checked).map(c => c.dataset.unit);
    const e = { title, picks, sec: Math.max(15, +$("#es").value || 60), open: true, released: false, units: unitsSel, created: new Date().toISOString() };
    const id = "E" + Date.now();
    $("#ec").disabled = true; $("#esum").textContent = "جاري الإنشاء والتوزيع…";
    try {
      await setDoc(doc(db, "exams", id), e); S.exams.unshift({ id, ...e }); S.attempts[id] = [];
      const cnt = await distribute({ id, ...e });
      S.selExam = id; toast(`تم إنشاء الامتحان وتوزيعه على ${cnt} ممرض`); S.tab = "results"; renderAdmin();
    } catch (err) { toast(errText(err)); $("#ec").disabled = false; }
  };
}

function aUsers() {
  const N = nurses().sort((a, b) => String(a.unit).localeCompare(String(b.unit)) || String(a.name).localeCompare(String(b.name)));
  $("#pane").innerHTML = `<section class="card"><h2>إضافة ممرضين من Excel</h2>
   <p class="small muted" style="margin:0">ملف فيه 4 أعمدة بالترتيب: الرقم الوظيفي، الاسم، الوحدة، كلمة المرور (6 أحرف على الأقل). الصف الأول عناوين. الممرض الموجود بالفعل يتم تحديث اسمه ووحدته فقط.</p>
   <input type="file" id="uf" accept=".xlsx,.xls,.csv">
   <p class="small muted" style="margin:0">Firebase يسمح بإنشاء 100 حساب في الساعة. لو العدد أكبر، ارفع نفس الملف مرة ثانية بعد ساعة وسيكمل من حيث توقف.</p>
   <div id="ulog" class="small"></div></section>
  <section class="card"><h2>إضافة ممرض واحد</h2>
   <div class="row"><label class="grow" for="n1">الرقم الوظيفي<input id="n1" inputmode="numeric"></label><label class="grow" for="n2">الاسم<input id="n2"></label></div>
   <div class="row"><label class="grow" for="n3">الوحدة<input id="n3"></label><label class="grow" for="n4">كلمة المرور<input id="n4"></label></div>
   <button class="btn primary" id="nadd">إضافة</button></section>
  <section class="card"><h2>الممرضين (${N.length})</h2><div class="tablewrap"><table><thead><tr><th>الرقم</th><th>الاسم</th><th>الوحدة</th><th>الحالة</th><th></th></tr></thead><tbody>
   ${N.map(u => `<tr><td>${esc(u.empId)}</td><td>${esc(u.name)}</td><td>${esc(u.unit)}</td><td>${u.active === false ? `<span class="pill bad">موقوف</span>` : `<span class="pill good">نشط</span>`}</td>
     <td><button class="btn sm" data-act="${u.uid}">${u.active === false ? "تفعيل" : "إيقاف"}</button> <button class="btn sm danger" data-del="${u.uid}">حذف</button></td></tr>`).join("") || `<tr><td colspan="5" class="muted">لا يوجد ممرضون بعد.</td></tr>`}
   </tbody></table></div>
   <p class="small muted" style="margin:0">نسي ممرض كلمة المرور؟ احذف حسابه من Firebase (Authentication ← Users)، ثم احذفه من هنا، ثم أضفه من جديد بكلمة مرور جديدة. نتائجه السابقة تبقى كما هي.</p></section>`;
  const log = m => { $("#ulog").innerHTML += `<div>${esc(m)}</div>`; };
  async function addOne(id, name, unit, pass) {
    id = String(id).trim(); name = String(name).trim(); unit = String(unit).trim(); pass = String(pass).trim();
    if (!id || !name || !unit) return "skip";
    if (String(id).toLowerCase() === String(ADMIN_ID).toLowerCase()) return "skip";
    const ex = S.users.find(u => u.empId === id);
    if (ex) { await updateDoc(doc(db, "users", ex.uid), { name, unit }); Object.assign(ex, { name, unit }); return "updated"; }
    if (pass.length < 6) return "short";
    const cred = await createUserWithEmailAndPassword(secAuth, emailOf(id), pass);
    const p = { empId: id, name, unit, role: "nurse", active: true };
    await setDoc(doc(db, "users", cred.user.uid), p); await signOut(secAuth);
    S.users.push({ uid: cred.user.uid, ...p }); return "created";
  }
  $("#uf").onchange = async ev => {
    const f = ev.target.files[0]; if (!f) return;
    $("#ulog").innerHTML = "";
    let rows; try { rows = await readSheet(f); } catch (err) { log("تعذر قراءة الملف"); return; }
    if (rows.length && !/^\d/.test(String(rows[0][0]).trim())) rows = rows.slice(1);
    rows = rows.filter(r => String(r[0]).trim());
    const c = { created: 0, updated: 0, skip: 0, short: 0 }; const bad = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      $("#ulog").dataset.p = `${i + 1}/${rows.length}`; log(`(${i + 1}/${rows.length}) ${r[0]}`);
      try { const res = await addOne(r[0], r[1], r[2], r[3]); c[res]++; if (res === "short") bad.push(`${r[0]}: كلمة مرور أقل من 6`); }
      catch (err) {
        if (String(err.code).includes("too-many-requests")) { log("وصلت للحد المسموح (100 حساب/ساعة). ارفع نفس الملف بعد ساعة لإكمال الباقي."); break; }
        if (String(err.code).includes("email-already-in-use")) bad.push(`${r[0]}: الحساب موجود في Firebase بدون بيانات. احذفه من Authentication وأعد الرفع`);
        else bad.push(`${r[0]}: ${errText(err)}`);
      }
    }
    $("#ulog").innerHTML = `<p class="hint" style="margin:0">جديد: ${c.created} · تم تحديثه: ${c.updated} · تم تخطيه: ${c.skip + c.short}</p>${bad.map(b => `<div style="color:var(--bad)">${esc(b)}</div>`).join("")}`;
    setTimeout(() => { if (S.tab === "users") { const keep = $("#ulog").innerHTML; aUsers(); $("#ulog").innerHTML = keep; } }, 50);
  };
  $("#nadd").onclick = async () => {
    $("#nadd").disabled = true;
    try { const r = await addOne($("#n1").value, $("#n2").value, $("#n3").value, $("#n4").value);
      toast({ created: "تمت الإضافة", updated: "تم تحديث البيانات", skip: "أكمل كل الخانات", short: "كلمة المرور 6 أحرف على الأقل" }[r]); if (r === "created" || r === "updated") aUsers(); }
    catch (err) { toast(String(err.code).includes("email-already-in-use") ? "الرقم مسجل في Firebase. احذفه من Authentication أولاً" : errText(err)); }
    $("#nadd").disabled = false;
  };
  $$("[data-act]").forEach(b => b.onclick = async () => { const u = S.users.find(x => x.uid === b.dataset.act); try { await updateDoc(doc(db, "users", u.uid), { active: u.active === false }); u.active = u.active === false; } catch (err) { toast(errText(err)); } aUsers(); });
  $$("[data-del]").forEach(b => b.onclick = () => { const u = S.users.find(x => x.uid === b.dataset.del);
    confirmBox(`حذف ${u.name} من التطبيق؟ لن يستطيع الدخول. نتائجه السابقة تبقى.`, async () => { try { await deleteDoc(doc(db, "users", u.uid)); S.users = S.users.filter(x => x.uid !== u.uid); } catch (err) { toast(errText(err)); } aUsers(); }); });
}

function aBank() {
  const all = Object.values(S.questions), topics = [...new Set(all.map(q => q.topic))].sort();
  $("#pane").innerHTML = `<section class="card"><h2>بنك الأسئلة (${all.length})</h2>
   <div class="tablewrap"><table><thead><tr><th>الموضوع</th>${DIFFS.map(d => `<th>${DIFF_AR[d]}</th>`).join("")}<th>الإجمالي</th></tr></thead><tbody>
   ${topics.map(t => { const c = DIFFS.map(d => all.filter(q => q.topic === t && q.diff === d).length); return `<tr><td dir="ltr" style="text-align:right">${esc(t)}</td>${c.map(x => `<td>${x}</td>`).join("")}<td><b>${c.reduce((a, b) => a + b, 0)}</b></td></tr>`; }).join("") || `<tr><td colspan="5" class="muted">البنك فارغ.</td></tr>`}
   </tbody></table></div></section>
  <section class="card"><h2>رفع أسئلة من Excel</h2>
   <p class="small muted" style="margin:0">ارفع ملف بنك الأسئلة بنفس شكل الشيت (ID, Topic, Level, Difficulty, Type, Question, A, B, C, D, Correct, Rationale). السؤال الذي له نفس الـ ID يتم تحديثه.</p>
   <label style="display:flex;gap:8px;align-items:center;color:var(--ink)" for="appr"><input type="checkbox" id="appr" style="width:auto">ارفع الأسئلة المعتمدة فقط (Reviewer status = Approved)</label>
   <input type="file" id="bf" accept=".xlsx,.xls,.csv"><div id="blog" class="small"></div></section>`;
  $("#bf").onchange = async ev => {
    const f = ev.target.files[0]; if (!f) return;
    let rows; try { rows = await readSheet(f, "Question Bank"); } catch (err) { $("#blog").textContent = "تعذر قراءة الملف"; return; }
    const H = rows[0].map(h => String(h).trim().toLowerCase());
    const col = n => H.indexOf(n);
    const need = ["id", "topic", "level", "difficulty", "type", "question", "a", "b", "c", "d", "correct", "rationale"];
    const miss = need.filter(n => col(n) < 0);
    if (miss.length) { $("#blog").innerHTML = `<p class="warnbar">أعمدة ناقصة: ${esc(miss.join(", "))}</p>`; return; }
    const rs = col("reviewer status"), onlyAppr = $("#appr").checked;
    const ops = [], bad = []; let skipped = 0;
    rows.slice(1).forEach((r, i) => {
      const g = n => String(r[col(n)] ?? "").trim();
      if (!g("id")) return;
      const status = rs >= 0 ? String(r[rs]).trim().toLowerCase() : "";
      if (status === "delete" || (onlyAppr && status !== "approved")) { skipped++; return; }
      const q = { topic: g("topic"), level: g("level"), diff: g("difficulty"), type: g("type"), stem: g("question"), opts: [g("a"), g("b"), g("c"), g("d")], correct: g("correct").toUpperCase(), rat: g("rationale") };
      if (!DIFFS.includes(q.diff) || !"ABCD".includes(q.correct) || q.correct.length !== 1 || !q.stem || q.opts.some(o => !o) || /[\/.#\[\]]/.test(g("id"))) { bad.push(g("id") || `row ${i + 2}`); return; }
      ops.push([g("id"), q]);
    });
    $("#blog").textContent = "جاري الرفع…";
    try {
      await commitBatches(ops.map(([id, q]) => b => b.set(doc(db, "questions", id), q)));
      ops.forEach(([id, q]) => S.questions[id] = q);
      aBank();
      $("#blog").innerHTML = `<p class="hint" style="margin:0">تم رفع ${ops.length} سؤال${skipped ? ` · تم تخطي ${skipped}` : ""}${bad.length ? ` · ${bad.length} صف به خطأ: ${esc(bad.join(", "))}` : ""}</p>`;
    } catch (err) { $("#blog").innerHTML = `<p class="warnbar">${esc(errText(err))}</p>`; }
  };
}
