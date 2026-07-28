/* app.js — 物流倒计时预警台 前端逻辑
 * 依赖：window.QEngine (engine.js) 与 window.QXlsxIO (xlsx_io.js)
 */
(function () {
  "use strict";

  const E = window.QEngine;
  const X = window.QXlsxIO;

  const PALETTE = [
    "#FF9EBB", "#6FD3A8", "#FFB07C", "#9ED8FF", "#C9B6FF",
    "#FFE08A", "#F2A6C2", "#7FD8C6", "#FFC9A8", "#B8C9FF",
  ];

  const STATUS_CLASS = {
    "⏰ 超36h未发货": "st-timeout",
    "⚠️ 距36h不足12h": "st-urgent",
    "✅ 正常": "st-normal",
    "✅ 已完成": "st-done",
    "⛔ 已取消": "st-cancel",
    "⚠️ 缺付款时间": "st-nopay",
  };

  const FB_KEY = "qiaofei_logi_feedback_v1";
  const FB_STATUS = ["待确认", "已发货", "未发货"];
  const HISTORY_KEY = "qiaofei_logi_history_v1";
  const ROLE_KEY = "qiaofei_logi_role_v1";

  let orderFile = null;
  let trackFile = null;
  let currentResult = null;
  let feedback = loadFeedback();      // { 订单号: {status, reason, time, supplier} }
  let urgentOnly = false;             // 主表"只看紧急订单"开关
  let fbSupplier = "";                // 反馈视图当前供应商，""=管理总览
  let fbSearch = "";                  // 反馈表搜索关键字（订单号/商品名称/快递单号/反馈状态）
  let role = (localStorage.getItem(ROLE_KEY) || "manager"); // "manager"=管理总览 | "supplier"=供应商视图
  let lockedSupplier = "";            // 供应商视图锁定的供应商名（来自快照/上传）

  // 分页状态（解决大批量订单渲染卡顿）
  const MAIN_PAGE_SIZE = 80;
  const FB_PAGE_SIZE = 60;
  let mainPage = 1;
  let fbPage = 1;

  const $ = (id) => document.getElementById(id);

  // 防抖：输入停止 ~180ms 后再执行，避免每次按键都全量重渲染表格
  function debounce(fn, ms) {
    let t = null;
    return function (...args) {
      if (t) clearTimeout(t);
      t = setTimeout(() => { t = null; fn.apply(this, args); }, ms);
    };
  }

  function esc(v) {
    v = v == null ? "" : String(v);
    return v.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function setStatus(msg, type) {
    const el = $("status");
    el.hidden = false;
    el.className = "status" + (type ? " " + type : "");
    el.textContent = msg;
  }

  /* ================= 反馈持久化 ================= */
  function loadFeedback() {
    try {
      const raw = localStorage.getItem(FB_KEY);
      if (!raw) return {};
      const obj = JSON.parse(raw);
      return obj && typeof obj === "object" ? obj : {};
    } catch (e) { return {}; }
  }
  function saveFeedback() {
    try { localStorage.setItem(FB_KEY, JSON.stringify(feedback)); } catch (e) {}
  }
  function fbBadge(status, urgent) {
    if (status === "已发货") return `<span class="badge st-fb-done">✅已发货</span>`;
    if (status === "未发货") return `<span class="badge st-fb-none">❌未发货</span>`;
    if (status === "待确认") return `<span class="badge st-fb-wait">⏳待确认</span>`;
    return urgent ? `<span class="badge st-fb-wait">⏳待确认</span>` : `<span class="badge st-fb-na">—</span>`;
  }

  /* ================= 累计汇总（每日历史存储） ================= */
  // 结构：{ "YYYY-MM-DD": { date, perSupplier:{ 供应商:{total,timeout,urgent12,normal,done,cancel,missingPay,shipped,notShipped,pendingFb} }, overall:{...} } }
  function loadHistory() {
    try {
      const raw = localStorage.getItem(HISTORY_KEY);
      if (!raw) return {};
      const obj = JSON.parse(raw);
      return obj && typeof obj === "object" ? obj : {};
    } catch (e) { return {}; }
  }
  function saveHistory(h) {
    try { localStorage.setItem(HISTORY_KEY, JSON.stringify(h)); } catch (e) {}
  }
  function todayStr(d) {
    d = d || new Date();
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  }
  // 文件名用：MMDD_HHMMSS（月日 + 具体时间）
  function nowStamp(d) {
    d = d || new Date();
    const p = (n) => String(n).padStart(2, "0");
    return `${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  }

  // 从当前结果 + 反馈 计算当日 per-supplier 指标
  function computeDailyMetrics(res, fb) {
    const perSupplier = {};
    const blank = () => ({ total:0, timeout:0, urgent12:0, normal:0, done:0, cancel:0, missingPay:0, shipped:0, notShipped:0, pendingFb:0 });
    const ensure = (s) => { if (!perSupplier[s]) perSupplier[s] = blank(); return perSupplier[s]; };
    res.countdown.forEach((r) => {
      const sup = normSupplierName(r["供应商名称"]);
      const m = ensure(sup);
      m.total++;
      const st = r["状态监控"] || "";
      if (st === "⏰ 超36h未发货") m.timeout++;
      else if (st === "⚠️ 距36h不足12h") m.urgent12++;
      else if (st === "✅ 正常") m.normal++;
      else if (st === "✅ 已完成") m.done++;
      else if (st === "⛔ 已取消") m.cancel++;
      else if (st === "⚠️ 缺付款时间") m.missingPay++;
      // 反馈统计（仅对紧急/超时单计 pendingFb）
      if (st === "⏰ 超36h未发货" || st === "⚠️ 距36h不足12h") {
        const f = fb[r["订单号"]];
        if (f && f.status === "已发货") m.shipped++;
        else if (f && f.status === "未发货") m.notShipped++;
        else m.pendingFb++;
      }
    });
    const overall = blank();
    Object.keys(perSupplier).forEach((s) => {
      const m = perSupplier[s];
      ["total","timeout","urgent12","normal","done","cancel","missingPay","shipped","notShipped","pendingFb"].forEach((k) => overall[k] += m[k]);
    });
    return { perSupplier, overall };
  }
  function normSupplierName(s) {
    s = (s == null ? "" : String(s)).trim();
    return s || "(未填供应商)";
  }

  // 归档今日看板到历史（同日期会覆盖当日；累计为跨日求和）
  function archiveToday() {
    if (!currentResult) { setStatus("请先处理数据再归档今日看板。", "err"); return; }
    const h = loadHistory();
    const day = todayStr();
    const metrics = computeDailyMetrics(currentResult, feedback);
    // 只归档「有风险」的供应商（超36h未发货>0 或 距36h不足12h>0），无风险的当日不进累计
    const riskSuppliers = {};
    let riskTotal = 0;
    Object.keys(metrics.perSupplier).forEach((s) => {
      const m = metrics.perSupplier[s];
      if (m.timeout > 0 || m.urgent12 > 0) {
        riskSuppliers[s] = m;
        riskTotal += m.total;
      }
    });
    h[day] = { date: day, perSupplier: riskSuppliers, overall: metrics.overall, archivedAt: new Date().toISOString() };
    saveHistory(h);
    const totalSup = Object.keys(metrics.perSupplier).length;
    setStatus(`📅 已归档 ${day} 看板：今日共 ${totalSup} 个供应商，其中 ${Object.keys(riskSuppliers).length} 个有风险订单（超时/不足12h），已纳入累计；无风险供应商不计入。`, "ok");
    renderCumulative();
  }

  // 累计汇总看板：按供应商累加所有日期的 超时/未发货 等指标
  function renderCumulative() {
    const h = loadHistory();
    const days = Object.keys(h).sort();
    const secBody = $("cumulative-body");
    if (!days.length) {
      secBody.innerHTML = `<div class="empty-hint">还没有归档任何日期的看板。处理完每日数据后，点击「📅 归档今日看板」即可开始累计统计。</div>`;
      renderCumulativeChart([], []);
      return;
    }
    // 累加 per-supplier
    const acc = {};
    const blank = () => ({ total:0, timeout:0, urgent12:0, normal:0, done:0, cancel:0, missingPay:0, shipped:0, notShipped:0, pendingFb:0 });
    days.forEach((d) => {
      const ps = h[d].perSupplier || {};
      Object.keys(ps).forEach((s) => {
        if (!acc[s]) acc[s] = blank();
        const m = ps[s];
        ["total","timeout","urgent12","normal","done","cancel","missingPay","shipped","notShipped","pendingFb"].forEach((k) => acc[s][k] += (m[k] || 0));
      });
    });
    const suppliers = Object.keys(acc).sort((a, b) => acc[b].timeout - acc[a].timeout);
    // 汇总行
    const sum = blank();
    suppliers.forEach((s) => ["total","timeout","urgent12","normal","done","cancel","missingPay","shipped","notShipped","pendingFb"].forEach((k) => sum[k] += acc[s][k]));

    let html = `<div class="cum-meta">已归档 <b>${days.length}</b> 天（${days[0]} ~ ${days[days.length-1]}）。累计统计各供应商发货预警情况：</div>`;
    html += `<table class="tbl-cum"><thead><tr>` +
      ["供应商","累计总单","累计超时","累计不足12h","累计未发货反馈","累计已发货反馈","累计完成","累计取消","超时率"].map((c) => `<th>${esc(c)}</th>`).join("") +
      `</tr></thead><tbody>`;
    suppliers.forEach((s, i) => {
      const m = acc[s];
      const rate = m.total ? ((m.timeout / m.total) * 100).toFixed(1) + "%" : "—";
      const color = PALETTE[i % PALETTE.length];
      html += `<tr>` +
        `<td><span class="dot" style="background:${color}"></span>${esc(s)}</td>` +
        `<td>${m.total}</td>` +
        `<td class="num-neg">${m.timeout}</td>` +
        `<td class="num-warn">${m.urgent12}</td>` +
        `<td>${m.notShipped}</td>` +
        `<td>${m.shipped}</td>` +
        `<td>${m.done}</td>` +
        `<td>${m.cancel}</td>` +
        `<td>${rate}</td>` +
        `</tr>`;
    });
    html += `<tr class="cum-sum"><td>合计</td><td>${sum.total}</td><td class="num-neg">${sum.timeout}</td><td class="num-warn">${sum.urgent12}</td><td>${sum.notShipped}</td><td>${sum.shipped}</td><td>${sum.done}</td><td>${sum.cancel}</td><td>${sum.total?((sum.timeout/sum.total)*100).toFixed(1)+"%":"—"}</td></tr>`;
    html += `</tbody></table>`;
    secBody.innerHTML = html;

    // 图表：各供应商累计超时 vs 累计未发货反馈
    const items = suppliers.map((s, i) => ({ label: s, timeout: acc[s].timeout, notShipped: acc[s].notShipped, color: PALETTE[i % PALETTE.length] }));
    renderCumulativeChart(items, days);
    // 更新概览累计卡片
    updateCumulativeSummary(sum, days.length);
  }

  function renderCumulativeChart(items, days) {
    const canvas = $("chart-cum");
    if (!canvas) return;
    const legendEl = $("legend-cum");
    if (!items.length) { drawBar(canvas, legendEl, []); return; }
    // 用分组柱：每个供应商画 累计超时 一根（简化用单色柱展示超时，图例说明）
    const barItems = items.map((it) => ({ label: it.label.length > 6 ? it.label.slice(0, 6) + "…" : it.label, value: it.timeout, color: it.color }));
    drawBar(canvas, legendEl, barItems);
    if (legendEl) legendEl.innerHTML = `<span class="lg"><span class="dot" style="background:#FF6B81"></span>柱高 = 累计超时单数（共 ${items.length} 家供应商，归档 ${days.length} 天）</span>`;
  }
  function updateCumulativeSummary(sum, dayCount) {
    const el = $("cum-summary");
    if (!el) return;
    const cards = [
      { c: "s-timeout", icon: "⏰", num: sum.timeout, lab: "累计超时单" },
      { c: "s-urgent", icon: "⚠️", num: sum.urgent12, lab: "累计不足12h" },
      { c: "s-fb-none", icon: "❌", num: sum.notShipped, lab: "累计未发货反馈" },
      { c: "s-fb-done", icon: "✅", num: sum.shipped, lab: "累计已发货反馈" },
      { c: "s-final", icon: "📋", num: sum.total, lab: "累计总单" },
      { c: "s-done", icon: "📅", num: dayCount, lab: "归档天数" },
    ];
    el.innerHTML = cards.map((k) => `<div class="sum ${k.c}"><div class="sum-ico">${k.icon}</div><div class="num">${k.num}</div><div class="lab">${k.lab}</div></div>`).join("");
  }

  // 导出累计汇总（Excel）
  function exportCumulative() {
    const h = loadHistory();
    const days = Object.keys(h).sort();
    if (!days.length) { alert("还没有累计数据可导出。"); return; }
    const acc = {};
    const blank = () => ({ total:0, timeout:0, urgent12:0, normal:0, done:0, cancel:0, missingPay:0, shipped:0, notShipped:0, pendingFb:0 });
    days.forEach((d) => { const ps = h[d].perSupplier || {}; Object.keys(ps).forEach((s) => { if (!acc[s]) acc[s] = blank(); const m = ps[s]; ["total","timeout","urgent12","normal","done","cancel","missingPay","shipped","notShipped","pendingFb"].forEach((k) => acc[s][k] += (m[k] || 0)); }); });
    const suppliers = Object.keys(acc).sort((a, b) => acc[b].timeout - acc[a].timeout);
    const cols = ["供应商","累计总单","累计超时","累计不足12h","累计未发货反馈","累计已发货反馈","累计完成","累计取消","累计缺付款时间","超时率"];
    const out = suppliers.map((s) => {
      const m = acc[s];
      return { 供应商:s, 累计总单:m.total, 累计超时:m.timeout, 累计不足12h:m.urgent12, 累计未发货反馈:m.notShipped, 累计已发货反馈:m.shipped, 累计完成:m.done, 累计取消:m.cancel, 累计缺付款时间:m.missingPay, 超时率: m.total?((m.timeout/m.total)*100).toFixed(1)+"%":"—" };
    });
    download(X.toXLSX(out, cols), `累计汇总_${todayStr()}.xlsx`);
  }

  /* ================= 角色模式（管理总览 / 供应商视图） ================= */
  // 当前视图下应展示的 countdown 行：供应商视图仅显示 lockedSupplier
  function viewRows() {
    if (!currentResult) return [];
    if (role === "supplier" && lockedSupplier) {
      return currentResult.countdown.filter((r) => normSupplierName(r["供应商名称"]) === lockedSupplier);
    }
    return currentResult.countdown;
  }
  function applyRoleUI() {
    document.body.classList.toggle("role-supplier", role === "supplier");
    document.body.classList.toggle("role-manager", role === "manager");
    const toggle = $("role-toggle");
    if (toggle) toggle.value = role;
    // 切换模式时，显示/隐藏管理专属区块
    const managerOnlyEls = document.querySelectorAll(".manager-only");
    managerOnlyEls.forEach((el) => { el.hidden = (role !== "manager"); });
    const supplierOnlyEls = document.querySelectorAll(".supplier-only");
    supplierOnlyEls.forEach((el) => { el.hidden = (role !== "supplier"); });
    const banner = $("role-banner");
    if (banner) {
      if (role === "supplier") {
        banner.hidden = false;
        banner.textContent = lockedSupplier
          ? `🔒 供应商视图：当前仅显示「${lockedSupplier}」的物流数据。反馈提交后将导出反馈文件回传经理。`
          : `🔒 供应商视图：请导入经理发来的专属快照，或上传你方源数据。只展示你方物流与反馈。`;
      } else {
        banner.hidden = true;
      }
    }
  }
  function setRole(r) {
    role = r;
    localStorage.setItem(ROLE_KEY, r);
    applyRoleUI();
    // 供应商视图：自动展开「供应商反馈」并滚动到该区，形成独立工作台体验
    if (r === "supplier" && currentResult) {
      const fb = $("sec-feedback");
      if (fb) { fb.classList.remove("collapsed"); const s = loadCollapsed() || {}; s["sec-feedback"] = false; saveCollapsed(s); }
      setTimeout(() => { const t = $("sec-feedback"); if (t) t.scrollIntoView({ behavior: "smooth", block: "start" }); }, 60);
    }
    if (currentResult) {
      if (r === "manager") renderCumulative();
      render(res0());
    }
  }
  // 取出 currentResult 的轻量副本用于重渲染
  function res0() { return currentResult; }

  // 经理：导出某供应商专属快照（不含其他供应商数据）
  function exportSupplierSnapshot() {
    if (!currentResult) { setStatus("请先处理数据再导出供应商快照。", "err"); return; }
    if (!fbSupplier) { setStatus("请先在「供应商反馈」下拉中选择要导出的供应商。", "err"); return; }
    const sup = fbSupplier;
    const rows = currentResult.step1.concat(currentResult.removedRows || [], currentResult.dupRows || []).concat(currentResult.final)
      .filter((r) => normSupplierName(r["供应商名称"]) === sup);
    const seen = new Set(); const dedup = [];
    rows.forEach((r) => { const no = r["订单号"]; if (!no || seen.has(no)) return; seen.add(no); dedup.push(r); });
    const snap = {
      type: "qiaofei_logi_snapshot",
      version: 1,
      exportedAt: new Date().toISOString(),
      targetSupplier: sup,
      orderRows: dedup,
      feedback: {},
    };
    download(new Blob([JSON.stringify(snap)], { type: "application/json;charset=utf-8" }),
      `供应商快照_${sup}_${todayStr()}.json`);
    setStatus(`📤 已导出「${sup}」专属快照（${dedup.length} 行，仅含该供应商），发给该供应商即可。`, "ok");
  }

  $("role-toggle").addEventListener("change", (e) => setRole(e.target.value));
  $("btn-export-supplier-snap").addEventListener("click", exportSupplierSnapshot);

  /* ================= 文件选择 / 拖拽 ================= */
  function bindFileInput(inputId, nameId, store) {
    const input = $(inputId);
    input.addEventListener("change", () => {
      if (input.files && input.files[0]) {
        store.v = input.files[0];
        $(nameId).textContent = input.files[0].name;
        $(nameId).style.color = "var(--mint-accent)";
      }
    });
    const dz = input.closest(".dropzone");
    ["dragover", "dragenter"].forEach((ev) =>
      dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); })
    );
    ["dragleave", "drop"].forEach((ev) =>
      dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); })
    );
    dz.addEventListener("drop", (e) => {
      const f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) { store.v = f; $(nameId).textContent = f.name; $(nameId).style.color = "var(--mint-accent)"; }
    });
  }
  const orderStore = { v: null }, trackStore = { v: null };
  bindFileInput("file-order", "name-order", orderStore);
  bindFileInput("file-track", "name-track", trackStore);

  /* ================= 运行 ================= */
  async function run(orderRows, trackRows) {
    $("btn-run").disabled = true;
    try {
      const res = E.processOrders(orderRows, trackRows || []);
      currentResult = res;
      applyRoleUI();
      render(res);
      if (res.warnings && res.warnings.length) {
        setStatus("⚠️ " + res.warnings.join("；"), "err");
      } else {
        setStatus("✅ 处理完成！共 " + res.final.length + " 行进入最终模板。", "ok");
      }
    } catch (err) {
      setStatus("❌ 处理出错：" + (err && err.message ? err.message : err), "err");
      console.error(err);
    } finally {
      $("btn-run").disabled = false;
    }
  }

  $("btn-run").addEventListener("click", async () => {
    if (!orderStore.v) { setStatus("请先选择「订单源文件」。", "err"); return; }
    setStatus("⏳ 正在读取并解析文件…", null);
    try {
      const orderRows = await X.readFileAuto(orderStore.v);
      let trackRows = [];
      if (trackStore.v) trackRows = await X.readFileAuto(trackStore.v);
      if (!orderRows.length) { setStatus("订单文件未识别到数据行，请检查表头。", "err"); return; }

      // ⑧ 上传前预检：检查必备列是否齐全
      const cols = Object.keys(orderRows[0]);
      const check = E.validateColumns(cols, ["订单号", "供应商", "订单状态", "物流状态", "付款时间", "快递单号"]);
      if (!check.ok) {
        const tips = check.missing.map((m) => {
          const s = m.hints.length ? `（疑似列名：${m.hints.join("、")}）` : "（未找到相似列）";
          return `「${m.intent}」${s}`;
        }).join("；");
        const ok = window.confirm(
          "⚠️ 列名预检提醒：未识别到以下必备列 ——\n" + tips +
          "\n\n仍可继续运行（缺失列将留空或跳过对应筛选），是否继续？"
        );
        if (!ok) { setStatus("已取消。请修正源文件列名后重新上传。", "err"); return; }
        setStatus("⚠️ 部分必备列未识别，已按现有列继续：" + tips, "err");
      }

      await run(orderRows, trackRows);
      // 供应商视图：上传的是自家源数据，锁定到文件中的供应商（取出现最多的）
      if (role === "supplier") {
        const sups = orderRows.map((r) => normSupplierName(r["供应商名称"])).filter((v) => v !== "(未填供应商)");
        if (sups.length) { lockedSupplier = sups.sort((a,b)=>sups.filter(x=>x===b).length-sups.filter(x=>x===a).length)[0]; }
        applyRoleUI();
      }
    } catch (err) {
      setStatus("❌ 读取文件失败：" + (err && err.message ? err.message : err), "err");
      console.error(err);
    }
  });

  $("btn-sample").addEventListener("click", () => {
    const s = makeSample();
    setStatus("🎲 已载入示例数据，演示处理效果。", "ok");
    run(s.orderRows, s.trackRows);
  });

  // ① 下载示例数据（生成真实 xlsx 样例，便于上手）
  $("btn-download-sample").addEventListener("click", () => {
    const s = makeSample();
    const orderCols = ["订单号", "供应商名称", "商品名称", "订单状态", "物流状态", "快递单号", "付款时间"];
    const trackCols = ["快递单号"];
    download(X.toXLSX(s.orderRows, orderCols), "示例_订单表.xlsx");
    setTimeout(() => download(X.toXLSX(s.trackRows, trackCols), "示例_快递单号表.xlsx"), 300);
    setStatus("⬇️ 已生成示例数据：示例_订单表.xlsx 与 示例_快递单号表.xlsx，请用它们尝试上传。", "ok");
  });

  /* ================= 渲染：概览 ================= */
  function countStatus(rows, st) {
    return rows.filter((r) => (r["状态监控"] || "") === st).length;
  }
  function pendingFeedbackCount() {
    if (!currentResult) return 0;
    let n = 0;
    viewRows().forEach((r) => {
      const st = r["状态监控"] || "";
      if (st === "⏰ 超36h未发货" || st === "⚠️ 距36h不足12h") {
        const f = feedback[r["订单号"]];
        if (!f || !f.status || f.status === "待确认") n++;
      }
    });
    return n;
  }
  function renderSummary(res) {
    const cd = viewRows();
    const cards = [
      { c: "s-origin", icon: "📥", num: res.step1.length + res.removed + res.dupRemoved, lab: "原始订单" },
      { c: "s-keep", icon: "✅", num: res.step1.length, lab: "筛选后保留" },
      { c: "s-remove", icon: "🗑️", num: res.removed, lab: "筛选剔除" },
      { c: "s-dup", icon: "🔁", num: res.dupRemoved, lab: "去重剔除" },
      { c: "s-final", icon: "📋", num: res.final.length, lab: "最终模板" },
      { c: "s-timeout", icon: "⏰", num: countStatus(cd, "⏰ 超36h未发货"), lab: "⏰ 超36h未发货" },
      { c: "s-urgent", icon: "⚠️", num: countStatus(cd, "⚠️ 距36h不足12h"), lab: "⚠️ 距36h不足12h" },
      { c: "s-done", icon: "🏁", num: countStatus(cd, "✅ 已完成") + countStatus(cd, "⛔ 已取消") + countStatus(cd, "⚠️ 缺付款时间"), lab: "已完成/取消/缺时" },
      { c: "s-fb", icon: "📨", num: pendingFeedbackCount(), lab: "待供应商确认" },
    ];
    $("summary").innerHTML = cards.map((k) =>
      `<div class="sum ${k.c}"><div class="sum-ico">${k.icon}</div><div class="num">${k.num}</div><div class="lab">${k.lab}</div></div>`
    ).join("");
  }

  /* ================= 渲染：图表 ================= */
  function setupCanvas(canvas) {
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    const w = rect.width || canvas.clientWidth || 300;
    const h = rect.height || 200;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    return { ctx, w, h };
  }

  function drawDonut(canvas, legendEl, items) {
    if (!canvas || canvas.offsetParent === null) return; // 折叠/不可见时不绘制（展开时再画）
    const { ctx, w, h } = setupCanvas(canvas);
    ctx.clearRect(0, 0, w, h);
    const total = items.reduce((a, b) => a + b.value, 0);
    const cx = w / 2, cy = h / 2, R = Math.min(w, h) / 2 - 8, r0 = R * 0.58;
    if (total === 0) {
      ctx.fillStyle = "#C9C2BC"; ctx.font = "14px sans-serif"; ctx.textAlign = "center";
      ctx.fillText("无数据", cx, cy); legendEl.innerHTML = ""; return;
    }
    let ang = -Math.PI / 2;
    items.forEach((it) => {
      const frac = it.value / total;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, R, ang, ang + frac * 2 * Math.PI);
      ctx.closePath();
      ctx.fillStyle = it.color; ctx.fill();
      ang += frac * 2 * Math.PI;
    });
    ctx.beginPath(); ctx.arc(cx, cy, r0, 0, 2 * Math.PI);
    ctx.fillStyle = "#FFFDFB"; ctx.fill();
    ctx.fillStyle = "#6B5B52"; ctx.textAlign = "center"; ctx.textBaseline = "middle";
    ctx.font = "bold 20px sans-serif"; ctx.fillText(String(total), cx, cy - 8);
    ctx.font = "12px sans-serif"; ctx.fillStyle = "#9A8C82"; ctx.fillText("总数", cx, cy + 12);

    legendEl.innerHTML = items.map((it) =>
      `<span class="lg"><span class="dot" style="background:${it.color}"></span>${esc(it.label)} <b>${it.value}</b></span>`
    ).join("");
  }

  function drawBar(canvas, legendEl, items) {
    if (!canvas || canvas.offsetParent === null) return; // 折叠/不可见时不绘制（展开时再画）
    const { ctx, w, h } = setupCanvas(canvas);
    ctx.clearRect(0, 0, w, h);
    if (!items.length) {
      ctx.fillStyle = "#C9C2BC"; ctx.font = "14px sans-serif"; ctx.textAlign = "center";
      ctx.fillText("无数据", w / 2, h / 2); legendEl.innerHTML = ""; return;
    }
    const max = Math.max.apply(null, items.map((i) => i.value)) || 1;
    const padL = 10, padR = 10, padB = 26, padT = 10;
    const cw = w - padL - padR;
    const ch = h - padB - padT;
    const n = items.length;
    const gap = 8;
    const bw = (cw - gap * (n - 1)) / n;
    ctx.fillStyle = "#F2E6DD"; ctx.fillRect(padL, padT + ch, cw, 2);
    items.forEach((it, i) => {
      const bh = (it.value / max) * ch;
      const x = padL + i * (bw + gap);
      const y = padT + ch - bh;
      ctx.fillStyle = it.color;
      roundRect(ctx, x, y, bw, bh, 5); ctx.fill();
      ctx.fillStyle = "#6B5B52"; ctx.font = "bold 11px sans-serif"; ctx.textAlign = "center";
      ctx.fillText(String(it.value), x + bw / 2, y - 4);
      const label = it.label.length > 5 ? it.label.slice(0, 5) + "…" : it.label;
      ctx.fillStyle = "#9A8C82"; ctx.font = "10px sans-serif";
      ctx.fillText(label, x + bw / 2, h - 12);
    });
    legendEl.innerHTML = items.map((it) =>
      `<span class="lg"><span class="dot" style="background:${it.color}"></span>${esc(it.label)} <b>${it.value}</b></span>`
    ).join("");
  }

  function roundRect(ctx, x, y, w, h, r) {
    if (h < 0) { y += h; h = -h; }
    r = Math.min(r, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
  }

  function toItems(map) {
    return Object.keys(map).map((k, i) => ({
      label: k, value: map[k], color: PALETTE[i % PALETTE.length],
    }));
  }

  function renderCharts(res) {
    const cd = viewRows();
    const orderMap = E.countBy(cd, "订单状态");
    const logMap = E.countBy(cd, "物流状态");
    const statusMap = {};
    cd.forEach((r) => {
      const s = r["状态监控"] || "未知";
      statusMap[s] = (statusMap[s] || 0) + 1;
    });
    drawDonut($("chart-order"), $("legend-order"), toItems(orderMap));
    drawDonut($("chart-log"), $("legend-log"), toItems(logMap));
    drawDonut($("chart-status"), $("legend-status"), toItems(statusMap));
    drawBar($("chart-top"), $("legend-top"), res.top10.map((t, i) => ({
      label: t.sup, value: t.risk, color: PALETTE[i % PALETTE.length],
    })));
  }

  // 展开被折叠的看板区时，补画之前因不可见而跳过的图表
  function redrawVisibleCharts() {
    if (!currentResult) return;
    const chartsSec = $("sec-charts");
    if (chartsSec && !chartsSec.classList.contains("collapsed")) renderCharts(currentResult);
    const cumSec = $("sec-cum");
    if (cumSec && !cumSec.classList.contains("collapsed")) renderCumulative();
  }

  /* ================= 渲染：表格（含筛选） ================= */
  let filterState = {};
  function getFilterCols(res) {
    const cols = res.countdown[0] ? Object.keys(res.countdown[0]) : [];
    const kws = [["供应商"], ["订单状态"], ["物流状态"], ["状态监控"]];
    const out = [];
    kws.forEach((kw) => {
      const c = cols.find((x) => kw.some((k) => x.indexOf(k) !== -1));
      if (c && out.indexOf(c) === -1) out.push(c);
    });
    return out;
  }
  function buildFilters(res) {
    const fcols = getFilterCols(res);
    const box = $("col-filters");
    box.innerHTML = "";
    filterState = {};
    fcols.forEach((col) => {
      const values = Array.from(new Set(res.countdown.map((r) => r[col] || "").filter((v) => v !== "")));
      values.sort();
      const sel = document.createElement("select");
      sel.dataset.col = col;
      let html = `<option value="">${esc(col)}：全部</option>`;
      values.forEach((v) => { html += `<option value="${esc(v)}">${esc(v)}</option>`; });
      sel.innerHTML = html;
      sel.addEventListener("change", () => { filterState[col] = sel.value; mainPage = 1; renderMainTable(); });
      box.appendChild(sel);
      filterState[col] = "";
    });
  }
  function filteredRows() {
    if (!currentResult) return [];
    const q = ($("search").value || "").trim().toLowerCase();
    return viewRows().filter((r) => {
      if (urgentOnly) {
        const st = r["状态监控"] || "";
        if (st !== "⏰ 超36h未发货" && st !== "⚠️ 距36h不足12h") return false;
      }
      for (const col in filterState) {
        const fv = filterState[col];
        if (fv && (r[col] || "") !== fv) return false;
      }
      if (q) {
        const hit = Object.keys(r).some((k) => String(r[k] == null ? "" : r[k]).toLowerCase().indexOf(q) !== -1);
        if (!hit) return false;
      }
      return true;
    });
  }
  function isUrgent(r) {
    const st = r["状态监控"] || "";
    return st === "⏰ 超36h未发货" || st === "⚠️ 距36h不足12h";
  }
  function numClass(hours) {
    const n = parseFloat(hours);
    if (isNaN(n)) return "";
    if (n < 0) return "num-neg";
    if (n < 12) return "num-warn";
    return "num-ok";
  }
  function mainDisplayCols() {
    const base = currentResult && currentResult.countdown[0] ? Object.keys(currentResult.countdown[0]) : [];
    const out = base.slice();
    if (out.indexOf("供应商反馈") === -1) out.push("供应商反馈");
    return out;
  }
  function renderPager(containerId, page, totalPages, onGoto, totalCount) {
    const el = $(containerId);
    if (!el) return;
    if (totalPages <= 1) { el.innerHTML = ""; return; }
    el.innerHTML =
      `<button class="pg-btn" data-pg="prev"${page <= 1 ? " disabled" : ""}>‹ 上一页</button>` +
      `<span class="pg-info">第 <b>${page}</b> / ${totalPages} 页 · 共 ${totalCount} 条</span>` +
      `<button class="pg-btn" data-pg="next"${page >= totalPages ? " disabled" : ""}>下一页 ›</button>`;
    el.querySelectorAll(".pg-btn").forEach((b) => {
      if (b.disabled) return;
      b.addEventListener("click", () => {
        const dir = b.dataset.pg;
        onGoto(dir === "prev" ? Math.max(1, page - 1) : Math.min(totalPages, page + 1));
      });
    });
  }

  function renderMainTable() {
    const res = currentResult;
    const cols = mainDisplayCols();
    const rows = filteredRows();
    const table = $("table-main");
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / MAIN_PAGE_SIZE));
    if (mainPage > totalPages) mainPage = totalPages;
    if (!cols.length) {
      table.innerHTML = '<tbody><tr><td>无数据</td></tr></tbody>';
      renderPager("table-pager", 1, 1, () => {}, 0);
      return;
    }
    const start = (mainPage - 1) * MAIN_PAGE_SIZE;
    const pageRows = rows.slice(start, start + MAIN_PAGE_SIZE);
    let html = "<thead><tr>" + cols.map((c) => `<th>${esc(c)}</th>`).join("") + "</tr></thead><tbody>";
    pageRows.forEach((r) => {
      const urgent = isUrgent(r);
      const rowCls = r["状态监控"] === "⏰ 超36h未发货" ? "row-timeout" : (urgent ? "row-urgent" : "");
      html += `<tr class="${rowCls}">`;
      cols.forEach((c) => {
        let v = r[c];
        if (c === "状态监控") {
          const cls = STATUS_CLASS[v] || "";
          html += `<td><span class="badge ${cls}">${esc(v)}</span></td>`;
        } else if (c === "供应商反馈") {
          const f = feedback[r["订单号"]];
          html += `<td>${fbBadge(f && f.status, urgent)}</td>`;
        } else if (c === "剩余(小时)" || c === "剩余(天)") {
          html += `<td class="${numClass(v)}">${esc(v)}</td>`;
        } else {
          html += `<td>${esc(v)}</td>`;
        }
      });
      html += "</tr>";
    });
    html += "</tbody>";
    table.innerHTML = html;
    renderPager("table-pager", mainPage, totalPages, (p) => { mainPage = p; renderMainTable(); }, total);
  }

  function renderRemovedTable(res) {
    const all = (res.removedRows || []).concat(res.dupRows || []);
    $("removed-count").textContent = all.length;
    const table = $("table-removed");
    if (!all.length) { table.innerHTML = '<tbody><tr><td>没有剔除任何行 🎉</td></tr></tbody>'; return; }
    const cols = Array.from(new Set(all.flatMap((r) => Object.keys(r))));
    const reasonIdx = cols.indexOf("剔除原因");
    if (reasonIdx !== -1) { cols.splice(reasonIdx, 1); cols.push("剔除原因"); }
    let html = "<thead><tr>" + cols.map((c) => `<th>${esc(c)}</th>`).join("") + "</tr></thead><tbody>";
    all.forEach((r) => {
      html += "<tr>";
      cols.forEach((c) => {
        if (c === "剔除原因") html += `<td><span class="badge st-cancel">${esc(r[c])}</span></td>`;
        else html += `<td>${esc(r[c])}</td>`;
      });
      html += "</tr>";
    });
    html += "</tbody>";
    table.innerHTML = html;
  }

  /* ================= 紧急待办条 ================= */
  function renderUrgentCallout(res) {
    const urgent = res.countdown.filter(isUrgent);
    const box = $("urgent-callout");
    if (!urgent.length) { box.hidden = true; return; }
    box.hidden = false;
    const timeout = urgent.filter((r) => (r["状态监控"] || "") === "⏰ 超36h未发货").length;
    const soon = urgent.length - timeout;
    $("uc-desc").textContent =
      `共 ${urgent.length} 笔订单需要尽快处理：超36h未发货 ${timeout} 笔、距36h不足12h ${soon} 笔。请优先联系对应供应商确认发货。`;
  }

  /* ================= 供应商联动反馈 ================= */
  function populateSupplierSelect() {
    const sel = $("fb-supplier");
    if (role === "supplier" && lockedSupplier) {
      // 供应商视图：下拉只显示自己，禁用切换
      sel.innerHTML = `<option value="${esc(lockedSupplier)}">${esc(lockedSupplier)}（本视图）</option>`;
      sel.value = lockedSupplier;
      fbSupplier = lockedSupplier;
      sel.disabled = true;
      return;
    }
    sel.disabled = false;
    const suppliers = currentResult
      ? Array.from(new Set(currentResult.countdown.map((r) => r["供应商名称"] || "").filter((v) => v)))
      : [];
    // 按"风险订单数"（超36h未发货 + 距36h不足12h）从多到少排序；并列则按名称
    const riskCount = {};
    if (currentResult) {
      currentResult.countdown.forEach((r) => {
        const s = r["供应商名称"] || "";
        if (!s) return;
        riskCount[s] = (riskCount[s] || 0) + (isUrgent(r) ? 1 : 0);
      });
    }
    suppliers.sort((a, b) => (riskCount[b] || 0) - (riskCount[a] || 0) || (a < b ? -1 : (a > b ? 1 : 0)));
    let html = `<option value="">— 管理总览（查看全部反馈）—</option>`;
    suppliers.forEach((s) => {
      const n = riskCount[s] || 0;
      html += `<option value="${esc(s)}">${esc(s)}（风险${n}）</option>`;
    });
    sel.innerHTML = html;
    sel.value = fbSupplier && suppliers.indexOf(fbSupplier) !== -1 ? fbSupplier : "";
    fbSupplier = sel.value;
  }

  function renderFeedbackTable() {
    const table = $("table-feedback");
    if (!currentResult) { table.innerHTML = '<tbody><tr><td>请先上传并处理源数据。</td></tr></tbody>'; return; }
    // 供应商视图：锁定到 lockedSupplier
    const viewSup = role === "supplier" && lockedSupplier ? lockedSupplier : fbSupplier;
    // 联动反馈只针对「距36h不足12h」与「超36h未发货」的订单
    const q = fbSearch.trim().toLowerCase();
    const rows = viewRows()
      .filter(isUrgent)
      .filter((r) => !viewSup || (r["供应商名称"] || "") === viewSup)
      .filter((r) => {
        if (!q) return true;
        const sup = (r["供应商名称"] || "").toLowerCase();
        return sup.indexOf(q) !== -1;
      })
      .slice()
      .sort((a, b) => parseFloat(a["剩余(小时)"]) - parseFloat(b["剩余(小时)"]));
    if (!rows.length) {
      const msg = q
        ? `没有匹配「${esc(fbSearch)}」的「距36h不足12h / 超36h未发货」订单。`
        : `当前视角下没有「距36h不足12h / 超36h未发货」的订单需要反馈。`;
      table.innerHTML = `<tbody><tr><td>${msg}</td></tr></tbody>`;
      renderPager("fb-pager", 1, 1, () => {}, 0);
      return;
    }

    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / FB_PAGE_SIZE));
    if (fbPage > totalPages) fbPage = totalPages;
    const start = (fbPage - 1) * FB_PAGE_SIZE;
    const pageRows = rows.slice(start, start + FB_PAGE_SIZE);

    const cols = ["订单号", "供应商名称", "商品名称", "快递公司", "快递单号", "订单状态", "物流状态", "状态监控", "剩余(小时)", "最晚发货时间", "是否已完成发货", "反馈原因", "反馈时间"];
    let html = "<thead><tr>" + cols.map((c) => `<th>${esc(c)}</th>`).join("") + "</tr></thead><tbody>";
    pageRows.forEach((r) => {
      const no = r["订单号"];
      const f = feedback[no] || {};
      const urgent = isUrgent(r);
      const fstatus = f.status || (urgent ? "待确认" : "");
      const rowCls = r["状态监控"] === "⏰ 超36h未发货" ? "row-timeout" : (urgent ? "row-urgent" : "");
      const selCls = fstatus === "已发货" ? "done" : (fstatus === "未发货" ? "none" : (fstatus === "待确认" ? "wait" : ""));
      html += `<tr class="${rowCls}" data-no="${esc(no)}">`;
      cols.forEach((c) => {
        if (c === "订单号") html += `<td>${esc(no)}</td>`;
        else if (c === "供应商名称") html += `<td>${esc(r["供应商名称"])}</td>`;
        else if (c === "商品名称") html += `<td>${esc(r["商品名称"] || "")}</td>`;
        else if (c === "快递公司") html += `<td>${esc(r["快递公司"] || "")}</td>`;
        else if (c === "快递单号") html += `<td>${esc(r["快递单号"] || "")}</td>`;
        else if (c === "订单状态") html += `<td>${esc(r["订单状态"] || "")}</td>`;
        else if (c === "物流状态") html += `<td>${esc(r["物流状态"] || "")}</td>`;
        else if (c === "状态监控") { const sc = STATUS_CLASS[r["状态监控"]] || ""; html += `<td><span class="badge ${sc}">${esc(r["状态监控"])}</span></td>`; }
        else if (c === "剩余(小时)") html += `<td class="${numClass(r["剩余(小时)"])}">${esc(r["剩余(小时)"])}</td>`;
        else if (c === "最晚发货时间") html += `<td>${esc(r["最晚发货时间"])}</td>`;
        else if (c === "是否已完成发货") {
          html += `<td><select class="fb-status ${selCls}" data-no="${esc(no)}" data-fld="status">` +
            FB_STATUS.map((s) => `<option value="${s}"${s === fstatus ? " selected" : ""}>${s}</option>`).join("") +
            `</select></td>`;
        }
        else if (c === "反馈原因") html += `<td><input class="fb-reason" data-no="${esc(no)}" data-fld="reason" value="${esc(f.reason || "")}" placeholder="未发货请填原因" /></td>`;
        else if (c === "反馈时间") html += `<td class="fb-time">${esc(f.time || (urgent ? "⚠需反馈" : "—"))}</td>`;
      });
      html += "</tr>";
    });
    html += "</tbody>";
    table.innerHTML = html;

    table.querySelectorAll(".fb-status").forEach((el) => {
      el.addEventListener("change", () => onFeedbackChange(el.dataset.no, "status", el.value));
    });
    table.querySelectorAll(".fb-reason").forEach((el) => {
      el.addEventListener("input", () => onFeedbackChange(el.dataset.no, "reason", el.value));
    });
    renderPager("fb-pager", fbPage, totalPages, (p) => { fbPage = p; renderFeedbackTable(); }, total);
  }

  function onFeedbackChange(no, fld, val) {
    if (!no) return;
    const sup = currentResult ? (currentResult.countdown.find((r) => r["订单号"] === no) || {})["供应商名称"] || "" : "";
    if (!feedback[no]) feedback[no] = { status: "", reason: "", time: "", supplier: sup };
    feedback[no][fld] = val;
    if (fld === "status" && val) feedback[no].time = new Date().toLocaleString("zh-CN");
    if (fld === "status" && val === "待确认") feedback[no].time = "";
    feedback[no].supplier = sup;
    saveFeedback();
    // 同步主表与概览
    renderMainTable();
    renderSummary(currentResult);
  }

  function exportFeedbackRows(rows, cols, filename) {
    download(X.toXLSX(rows, cols), filename);
  }

  // 导出待确认清单（发给供应商）
  $("btn-fb-export-list").addEventListener("click", () => {
    if (!currentResult) return;
    const viewSup = role === "supplier" && lockedSupplier ? lockedSupplier : fbSupplier;
    const list = viewRows().filter((r) => {
      if (viewSup && (r["供应商名称"] || "") !== viewSup) return false;
      // 仅含紧急/超时订单（即需要反馈的对象）
      return isUrgent(r);
    });
    if (!list.length) { alert("当前选择下没有「距36h不足12h / 超36h未发货」的订单需要反馈。"); return; }
    const cols = ["订单号", "供应商名称", "商品名称", "快递公司", "订单状态", "物流状态", "快递单号", "付款时间", "最晚发货时间", "剩余(小时)", "状态监控", "是否已完成发货", "反馈原因"];
    const out = list.map((r) => {
      const o = {};
      cols.forEach((c) => {
        if (c === "是否已完成发货") o[c] = "待确认";
        else if (c === "反馈原因") o[c] = "";
        else o[c] = r[c] || "";
      });
      return o;
    });
    const name = fbSupplier ? fbSupplier : "全部供应商";
    download(X.toXLSX(out, cols), "待确认清单_" + name + "_" + nowStamp() + ".xlsx");
  });

  // 导出全部反馈（自己留存汇总）
  $("btn-fb-export-all").addEventListener("click", () => {
    if (!currentResult) return;
    const cols = ["订单号", "供应商名称", "商品名称", "是否已完成发货", "反馈原因", "反馈时间"];
    const out = currentResult.countdown
      .filter((r) => feedback[r["订单号"]])
      .map((r) => {
        const f = feedback[r["订单号"]] || {};
        return {
          订单号: r["订单号"], 供应商名称: r["供应商名称"], 商品名称: r["商品名称"] || "",
          是否已完成发货: f.status || "", 反馈原因: f.reason || "", 反馈时间: f.time || "",
        };
      });
    if (!out.length) { alert("还没有任何供应商反馈记录。"); return; }
    download(X.toXLSX(out, cols), "供应商反馈汇总.xlsx");
  });

  // 把供应商填回的状态文本归一化为 已发货 / 未发货 / 待确认
  function normFbStatus(raw) {
    raw = (raw == null ? "" : String(raw)).trim();
    if (raw === "已发货" || raw === "已发" || raw === "已" || raw === "发货了") return "已发货";
    if (raw === "未发货" || raw === "未发" || raw === "没发" || raw === "未") return "未发货";
    if (raw === "待确认" || raw === "") return "待确认";
    return "待确认"; // 其他奇奇怪怪的文本按待确认处理，不丢数据
  }

  // 导入反馈文件（接收供应商回传的 Excel / CSV）
  $("btn-fb-import").addEventListener("click", () => $("file-fb").click());
  $("file-fb").addEventListener("change", async () => {
    const f = $("file-fb").files && $("file-fb").files[0];
    if (!f) return;
    if (!currentResult) {
      setStatus("❌ 导入反馈前请先「① 上传源数据」并点「▶ 开始处理」，拿到当前看板后再导入供应商回传的反馈文件。", "err");
      $("file-fb").value = "";
      return;
    }
    try {
      const rows = await X.readFileAuto(f);
      if (!rows.length) { setStatus("⚠️ 反馈文件没有可读到数据行，请确认文件含表头与数据。", "err"); $("file-fb").value = ""; return; }
      let merged = 0, unknown = 0;
      rows.forEach((r) => {
        const no = r["订单号"];
        // 兼容多种表头写法
        const statusRaw = (r["是否已完成发货"] || r["是否发货"] || r["已完成发货"] || r["反馈状态"] || r["状态"] || "").toString().trim();
        if (!no) return;
        const status = normFbStatus(statusRaw);
        const sup = (currentResult.countdown.find((x) => x["订单号"] === no) || {})["供应商名称"] || (r["供应商名称"] || "");
        feedback[no] = {
          status,
          reason: r["反馈原因"] || r["原因"] || (feedback[no] && feedback[no].reason) || "",
          time: r["反馈时间"] || new Date().toLocaleString("zh-CN"),
          supplier: sup,
        };
        if (currentResult.countdown.some((x) => x["订单号"] === no)) merged++;
        else unknown++;
      });
      saveFeedback();
      fbPage = 1;
      renderFeedbackTable();
      renderMainTable();
      renderSummary(currentResult);
      let msg = `📥 已导入反馈 ${merged} 条`;
      if (unknown) msg += `（其中 ${unknown} 条订单号在当前看板未找到，可能是不同批次数据，已暂存，切到对应批次后可显示）。`;
      if (merged === 0 && unknown > 0) {
        setStatus("⚠️ " + msg + " 请确认：①已先处理源数据；②回传的是供应商填好的「我的反馈」文件（不是你发出的待确认清单空表）。", "err");
      } else {
        setStatus(msg, "ok");
      }
    } catch (err) {
      setStatus("❌ 导入反馈失败：" + (err && err.message ? err.message : err), "err");
    } finally {
      $("file-fb").value = "";
    }
  });

  $("fb-supplier").addEventListener("change", () => {
    fbSupplier = $("fb-supplier").value;
    fbPage = 1;
    renderFeedbackTable();
  });

  $("fb-search").addEventListener("input", debounce(() => {
    fbSearch = $("fb-search").value || "";
    fbPage = 1;
    renderFeedbackTable();
  }, 180));

  // 供应商视图：导出「我的反馈」文件（仅反馈，回传经理）
  $("btn-export-myfb").addEventListener("click", () => {
    if (!currentResult) { setStatus("请先处理数据再导出反馈。", "err"); return; }
    const viewSup = role === "supplier" && lockedSupplier ? lockedSupplier : fbSupplier;
    const cols = ["订单号", "供应商名称", "是否已完成发货", "反馈原因", "反馈时间"];
    const out = viewRows()
      .filter((r) => feedback[r["订单号"]])
      .map((r) => {
        const f = feedback[r["订单号"]] || {};
        return { 订单号: r["订单号"], 供应商名称: r["供应商名称"] || "", 是否已完成发货: f.status || "", 反馈原因: f.reason || "", 反馈时间: f.time || "" };
      });
    if (!out.length) { alert("还没有任何反馈记录可导出。"); return; }
    const name = viewSup || "供应商";
    download(X.toXLSX(out, cols), `我的反馈_${name}_${todayStr()}.xlsx`);
    setStatus(`📤 已导出「${name}」反馈文件（${out.length} 条），请回传给经理。`, "ok");
  });

  /* ================= 导出 ================= */
  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1200);
  }
  function mainColumns() {
    return mainDisplayCols();
  }
  function augmentWithFeedback(rows) {
    return rows.map((r) => {
      const f = feedback[r["订单号"]];
      const urgent = isUrgent(r);
      const o = Object.assign({}, r);
      o["供应商反馈"] = f && f.status ? f.status : (urgent ? "待确认" : "");
      return o;
    });
  }
  $("btn-xlsx").addEventListener("click", () => {
    if (!currentResult) return;
    const cols = mainColumns();
    download(X.toXLSX(augmentWithFeedback(currentResult.countdown), cols), "物流倒计时预警_最终模板.xlsx");
  });
  $("btn-csv").addEventListener("click", () => {
    if (!currentResult) return;
    const cols = mainColumns();
    download(new Blob(["\ufeff" + X.toCSV(augmentWithFeedback(currentResult.countdown), cols)], { type: "text/csv;charset=utf-8" }),
      "物流倒计时预警_最终模板.csv");
  });
  $("btn-removed-xlsx").addEventListener("click", () => {
    if (!currentResult) return;
    const all = (currentResult.removedRows || []).concat(currentResult.dupRows || []);
    const cols = Array.from(new Set(all.flatMap((r) => Object.keys(r))));
    const ri = cols.indexOf("剔除原因"); if (ri !== -1) { cols.splice(ri, 1); cols.push("剔除原因"); }
    download(X.toXLSX(all, cols), "物流倒计时预警_已剔除明细.xlsx");
  });
  $("btn-removed-csv").addEventListener("click", () => {
    if (!currentResult) return;
    const all = (currentResult.removedRows || []).concat(currentResult.dupRows || []);
    const cols = Array.from(new Set(all.flatMap((r) => Object.keys(r))));
    const ri = cols.indexOf("剔除原因"); if (ri !== -1) { cols.splice(ri, 1); cols.push("剔除原因"); }
    download(new Blob(["\ufeff" + X.toCSV(all, cols)], { type: "text/csv;charset=utf-8" }),
      "物流倒计时预警_已剔除明细.csv");
  });

  // 导出数据快照（发给供应商：含订单原始行 + 你的反馈记录，几 KB）
  $("btn-export-snapshot").addEventListener("click", () => {
    if (!currentResult) { setStatus("请先处理数据再导出快照。", "err"); return; }
    const d = new Date();
    const p = (n) => String(n).padStart(2, "0");
    const dateTag = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`;
    const snap = {
      type: "qiaofei_logi_snapshot",
      version: 1,
      exportedAt: d.toISOString(),
      orderRows: currentResult.step1.concat(currentResult.removedRows || [], currentResult.dupRows || [])
        .concat(currentResult.final),
      feedback: feedback,
    };
    // 去重：按订单号保留（优先 final，其次 step1）
    const seen = new Set();
    const dedup = [];
    snap.orderRows.forEach((r) => {
      const no = r["订单号"];
      if (!no || seen.has(no)) return;
      seen.add(no); dedup.push(r);
    });
    snap.orderRows = dedup;
    const str = JSON.stringify(snap);
    download(new Blob([str], { type: "application/json;charset=utf-8" }),
      `数据快照_${dateTag}.json`);
    setStatus(`📤 已导出数据快照（${dedup.length} 行订单 + ${Object.keys(feedback).length} 条反馈），发给供应商即可同步最新看板。`, "ok");
  });

  // 导入数据快照（接收你/同事发来的最新看板）
  $("btn-import-snapshot").addEventListener("click", () => $("file-snapshot").click());
  $("file-snapshot").addEventListener("change", async () => {
    const f = $("file-snapshot").files && $("file-snapshot").files[0];
    if (!f) return;
    try {
      const text = await f.text();
      const snap = JSON.parse(text);
      if (!snap || snap.type !== "qiaofei_logi_snapshot" || !Array.isArray(snap.orderRows)) {
        setStatus("❌ 文件不是有效的数据快照。", "err"); return;
      }
      // 合并反馈（保留本地已有的，快照覆盖同订单号）
      const merged = Object.assign({}, feedback, snap.feedback || {});
      feedback = merged;
      saveFeedback();
      // 供应商专属快照：锁定视图到该供应商
      if (snap.targetSupplier) { lockedSupplier = snap.targetSupplier; }
      // 用快照里的订单原始行重新跑引擎（源数据最权威）
      await run(snap.orderRows, []);
      if (role === "supplier" && lockedSupplier) applyRoleUI();
      setStatus("📥 已导入数据快照，看板已更新为最新。" +
        (snap.exportedAt ? `（导出时间：${new Date(snap.exportedAt).toLocaleString("zh-CN")}）` : ""), "ok");
    } catch (err) {
      setStatus("❌ 导入快照失败：" + (err && err.message ? err.message : err), "err");
    } finally {
      $("file-snapshot").value = "";
    }
  });

  /* ================= 主渲染 ================= */
  function render(res) {
    $("result-section").hidden = false;
    mainPage = 1; fbPage = 1;
    renderSummary(res);
    renderCharts(res);
    buildFilters(res);
    populateSupplierSelect();
    renderMainTable();
    renderRemovedTable(res);
    renderUrgentCallout(res);
    renderFeedbackTable();
    if (role === "manager") renderCumulative();
    window.scrollTo({ top: $("result-section").offsetTop - 10, behavior: "smooth" });
  }

  // 归档今日看板 + 导出累计（仅管理总览模式可用）
  $("btn-archive-today").addEventListener("click", archiveToday);
  $("btn-export-cum").addEventListener("click", exportCumulative);
  $("btn-clear-cum").addEventListener("click", () => {
    const h = loadHistory();
    if (!Object.keys(h).length) { setStatus("当前没有累计数据可清空。", "err"); return; }
    if (!window.confirm("确定要清空全部累计汇总看板数据吗？\n此操作会删除所有已归档日期的统计，且无法恢复。")) return;
    try { localStorage.removeItem(HISTORY_KEY); } catch (e) {}
    renderCumulative();
    updateCumulativeSummary({ total:0, timeout:0, urgent12:0, notShipped:0, shipped:0 }, 0);
    setStatus("🗑 已清空全部累计汇总数据。", "ok");
  });

  $("search").addEventListener("input", debounce(() => { mainPage = 1; renderMainTable(); }, 180));
  $("btn-urgent-jump").addEventListener("click", () => {
    urgentOnly = !urgentOnly;
    $("btn-urgent-jump").textContent = urgentOnly ? "查看全部订单" : "只看紧急订单";
    mainPage = 1;
    renderMainTable();
    if (urgentOnly) $("sec-table").scrollIntoView({ behavior: "smooth" });
  });

  /* ================= 顶部导航 scroll-spy ================= */
  function initNavSpy() {
    const links = Array.from(document.querySelectorAll(".nav-links a"));
    const map = {};
    links.forEach((a) => {
      const id = a.getAttribute("href").slice(1);
      const el = document.getElementById(id);
      if (el) map[id] = a;
    });
    const sections = Object.keys(map).map((id) => document.getElementById(id));
    function onScroll() {
      const y = window.scrollY + 90;
      let activeId = sections.length ? sections[0].id : null;
      sections.forEach((s) => { if (s.offsetTop <= y) activeId = s.id; });
      links.forEach((a) => a.classList.remove("active"));
      if (activeId && map[activeId]) map[activeId].classList.add("active");
    }
    window.addEventListener("scroll", onScroll, { passive: true });
    onScroll();
  }
  initNavSpy();

  /* ================= 板块折叠/展开 ================= */
  const COLLAPSE_KEY = "qiaofei_sec_collapsed_v1";
  function loadCollapsed() {
    try { const raw = localStorage.getItem(COLLAPSE_KEY); return raw ? JSON.parse(raw) : null; }
    catch (e) { return null; }
  }
  function saveCollapsed(obj) {
    try { localStorage.setItem(COLLAPSE_KEY, JSON.stringify(obj)); } catch (e) {}
  }
  function applyCollapse(sec, collapsed) {
    if (collapsed) sec.classList.add("collapsed");
    else sec.classList.remove("collapsed");
  }
  function initCollapse() {
    const secs = Array.from(document.querySelectorAll("section.card[data-collapse]"));
    const stored = loadCollapsed();
    secs.forEach((sec) => {
      const head = sec.querySelector(".card-head");
      if (!head) return;
      // 首次进入（无存储状态）：除「上传」外全部收起；有存储则按存储
      let collapsed;
      if (stored && Object.prototype.hasOwnProperty.call(stored, sec.id)) {
        collapsed = stored[sec.id];
      } else {
        collapsed = sec.id !== "sec-upload"; // 上传区默认展开，其余默认收起
      }
      applyCollapse(sec, collapsed);
      head.addEventListener("click", (e) => {
        // 避免点击内部交互元素（如按钮）误触
        if (e.target.closest("button") && e.target.classList.contains("collapse-btn") === false && e.target.tagName !== "BUTTON") return;
        if (e.target.closest("a, input, select, label")) return;
        const willCollapse = !sec.classList.contains("collapsed");
        applyCollapse(sec, willCollapse);
        const s = loadCollapsed() || {};
        s[sec.id] = willCollapse;
        saveCollapsed(s);
        // 展开看板/累计区时，补画之前因不可见跳过的图表
        if (!willCollapse) redrawVisibleCharts();
      });
    });
  }
  initCollapse();

  /* ================= 启动初始化 ================= */
  // 重开页面时立即恢复：① 角色视图 UI（含经理专属区块显隐）；② 累计归档看板
  // （localStorage 中昨日/历史归档数据在打开页面时就直接显示，无需先处理数据）
  applyRoleUI();
  if (role === "manager") renderCumulative();

  /* ================= 示例数据 ================= */
  function makeSample() {
    const now = new Date();
    const fmt = (d) => {
      const p = (n) => String(n).padStart(2, "0");
      return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
    };
    const ago = (h) => new Date(now.getTime() - h * 3600000);
    const orderRows = [
      { 订单号: "O20260720001", 供应商名称: "广东俏妃合作供应商有限公司", 商品名称: "俏妃卫生巾日用组合装", 订单状态: "待发货", 物流状态: "空白", 快递单号: "YT880000000111", 付款时间: fmt(ago(50)) },
      { 订单号: "O20260720002", 供应商名称: "浙江杭州某供应链有限公司", 商品名称: "纯棉夜用卫生巾", 订单状态: "待发货", 物流状态: "空白", 快递单号: "YT880000000222", 付款时间: fmt(ago(30)) },
      { 订单号: "O20260720003", 供应商名称: "山东青岛某贸易有限公司", 商品名称: "经期护理套装", 订单状态: "待发货", 物流状态: "无物流", 快递单号: "YT880000000333", 付款时间: fmt(ago(10)) },
      { 订单号: "O20260720004", 供应商名称: "福建厦门某电子商务有限公司", 商品名称: "安睡裤超值包", 订单状态: "待发货", 物流状态: "无物流 无物流", 快递单号: "YT880000000444", 付款时间: fmt(ago(5)) },
      { 订单号: "O20260720005", 供应商名称: "北京某科技发展有限公司", 商品名称: "护垫便携装", 订单状态: "已发货", 物流状态: "空白", 快递单号: "YT880000000555", 付款时间: fmt(ago(40)) },
      { 订单号: "O20260720006", 供应商名称: "湖南俏妃卫生用品有限公司", 商品名称: "俏妃经期裤", 订单状态: "待发货", 物流状态: "空白", 快递单号: "YT880000000666", 付款时间: fmt(now) },
      { 订单号: "O20260720007", 供应商名称: "河南郑州某商贸有限公司", 商品名称: "日用卫生巾家庭装", 订单状态: "待发货", 物流状态: "已签收", 快递单号: "YT880000000777", 付款时间: fmt(now) },
      { 订单号: "0123456789", 供应商名称: "江苏南京某实业有限公司", 商品名称: "迷你巾试用装", 订单状态: "待发货", 物流状态: "空白", 快递单号: "YT880000000888", 付款时间: fmt(ago(1)) },
      { 订单号: "O20260720009", 供应商名称: "上海某物流服务商有限公司", 商品名称: "加长夜用卫生巾", 订单状态: "待发货", 物流状态: "空白", 快递单号: "SF_DUP_999", 付款时间: fmt(ago(2)) },
    ];
    const trackRows = [{ 快递单号: "SF_DUP_999" }, { 快递单号: "OTHER_TRACK_123" }];
    return { orderRows, trackRows };
  }
})();
