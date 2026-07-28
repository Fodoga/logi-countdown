/* engine.js — 物流倒计时预警 核心处理引擎（浏览器/Node 通用）
 * 逻辑严格对齐本地 process_orders.py：9家湖南供应商剔除、
 * 订单状态白名单(仅留 已发货/待发货)、物流状态白名单(仅留 空/无物流任意重复)、
 * 快递单号去重、最晚发货=付款时间+1.5天、剩余天/小时、状态监控、Top10。
 */
(function (global) {
  "use strict";

  const REMOVE_SUPPLIERS = new Set([
    "湖南俏妃卫生用品有限公司",
    "湖南婕妙生物科技有限公司",
    "湖南逸红颜生物科技有限公司",
    "湖南卫蕾贸易有限公司",
    "湖南优感觉生物科技有限公司",
    "湖南紫色风铃贸易有限公司",
    "湖南护宫福生物科技有限公司",
    "湖南小良知网络科技有限公司",
    "湖南小黄巾卫生用品有限公司",
  ]);

  // 订单状态：与 process_orders.py 一致的【白名单】——仅保留 已发货 / 待发货，
  // 其余（取消/退款/已完成/已签收/运输中等）一律剔除。
  const KEEP_ORDER_STATUS = new Set(["已发货", "待发货"]);
  // 物流状态：与 process_orders.py 一致的【白名单】——仅保留 空白(空单元格) /
  // 字面"空白" / 仅由"无物流"重复任意次组成；运输中/派送中/已签收等一律剔除。
  const KEEP_LOGISTICS = new Set(["空白"]);

  const KW_SUPPLIER = ["供应商"];
  const KW_ORDER_STATUS = ["订单状态"];
  const KW_LOGISTICS = ["物流状态", "物流"];
  const KW_TRACKING = ["快递单号", "快递", "单号"];
  const KW_ORDER_NO = ["订单号"];

  function normVal(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && isNaN(v)) return "";
    return String(v).trim();
  }

  function cleanId(v) {
    let s = normVal(v).replace(/^'+/, ""); // 去掉前导单引号
    if (/^\d+\.0$/.test(s)) s = s.slice(0, -2);
    return s;
  }

  // 物流状态白名单（对齐 process_orders.py keep_logistics）：
  // 仅保留 空单元格 / 字面"空白" / 仅由"无物流"重复任意次组成（去空白后）；
  // 其余（运输中/派送中/已签收/已完结等）一律剔除。
  function keepLogistics(v) {
    let s = normVal(v).replace(/^'+/, ""); // 去前导单引号（Excel 文本标记）
    s = s.replace(/\s+/g, "");             // 去所有空白（空格/换行/全角空格/重复间隔）
    if (s === "") return true;             // 空单元格 → 保留（需预警）
    if (KEEP_LOGISTICS.has(s)) return true; // 字面"空白" → 保留
    return s.replace(/无物流/g, "") === ""; // 仅由"无物流"重复组成 → 保留；其余删除
  }

  function findCol(columns, keywords) {
    if (!columns) return null;
    for (const kw of keywords) {
      for (const c of columns) {
        if (c != null && String(c).indexOf(kw) !== -1) return c;
      }
    }
    return null;
  }

  // 常见列名别名（用于模糊匹配）。键为标准化意图，值为【具体】写法。
  // 注意：宽泛词（如"状态""快递""单号"）不放这里，否则会误中"付款状态""快递公司"等，
  // 故单独放在 GENERIC_ALIASES 仅作兜底。
  const COL_ALIASES = {
    订单号: ["订单编号", "订单id", "订单编号/单号", "orderno", "order_no"],
    供应商: ["供应商名称", "供货商", "商家", "店铺", "厂商", "供应商名", "供货商名称"],
    订单状态: ["订单的状态", "单据状态", "orderstatus"],
    物流状态: [],
    付款时间: ["支付时间", "付款时刻", "下单时间", "成交时间", "paytime", "支付时刻", "付款日期", "下单日期", "实付时间", "买家付款时间"],
    快递单号: ["运单号", "物流单号", "顺丰单号", "快递编号", "trackno", "运单"],
  };
  // 宽泛兜底别名：仅当具体别名全部未命中时才使用，避免误匹配
  // （订单状态 的"状态"会误中"付款状态"；快递单号 的"快递/单号"会误中"快递公司"）。
  const GENERIC_ALIASES = {
    订单号: ["单号"],
    订单状态: ["状态"],
    物流状态: ["物流"],
    快递单号: ["快递", "单号"],
  };

  // 标准化：去空格、转小写，便于比较
  function normColName(c) {
    return String(c == null ? "" : c).replace(/\s+/g, "").toLowerCase();
  }

  // 模糊匹配：优先【意图本体 + 具体别名】，其次才用【宽泛兜底别名】。
  // 这样既能兼容不同表头写法，又不会误中"付款状态""快递公司"等。
  function findColSmart(columns, intent) {
    if (!columns || !columns.length) return null;
    const specific = (COL_ALIASES[intent] || []).concat([intent]); // 意图本体优先
    const generic = GENERIC_ALIASES[intent] || [];
    const normCols = columns.map((c) => ({ raw: c, n: normColName(c) }));
    // 1) 具体别名 / 意图本体（优先）
    for (const kw of specific) {
      const nk = normColName(kw);
      const hit = normCols.find((x) => x.n.indexOf(nk) !== -1);
      if (hit) return hit.raw;
    }
    // 2) 宽泛兜底别名（仅当具体别名都未命中）
    for (const kw of generic) {
      const nk = normColName(kw);
      const hit = normCols.find((x) => x.n.indexOf(nk) !== -1);
      if (hit) return hit.raw;
    }
    return null;
  }

  // 上传前预检：检查必备列是否齐全，返回 { ok, missing:[{intent, hints:[]}], suggestions:{} }
  function validateColumns(columns, intentList) {
    const list = intentList || ["订单号", "供应商", "订单状态", "物流状态", "付款时间", "快递单号"];
    const missing = [];
    const suggestions = {};
    list.forEach((intent) => {
      const found = findColSmart(columns, intent);
      if (!found) {
        // 给出候选提示：名称里含任意别名词的字段
        const hints = (columns || []).filter((c) => {
          const n = normColName(c);
          return COL_ALIASES[intent].some((a) => n.indexOf(normColName(a)) !== -1);
        });
        missing.push({ intent, hints });
      } else {
        suggestions[intent] = found;
      }
    });
    return { ok: missing.length === 0, missing, suggestions };
  }

  function extractTrackingNumbers(trackRows) {
    const set = new Set();
    if (!trackRows || !trackRows.length) return set;
    const cols = Object.keys(trackRows[0]);
    const tcol = findCol(cols, KW_TRACKING);
    trackRows.forEach((r) => {
      const vals = tcol ? [r[tcol]] : Object.values(r);
      vals.forEach((v) => {
        const s = cleanId(v);
        if (s) set.add(s.toLowerCase());
      });
    });
    return set;
  }

  // 付款时间解析：支持 Excel 序列号 与 常见字符串格式
  function parsePayTime(v) {
    const s = normVal(v);
    if (!s) return null;
    if (/^\d+(\.\d+)?$/.test(s)) {
      const serial = parseFloat(s);
      if (isFinite(serial)) {
        // Excel 1900 日期系统：1899-12-30 为 0，1970-01-01 为 25569
        return new Date((serial - 25569) * 86400000);
      }
    }
    let d = new Date(s);
    if (!isNaN(d.getTime())) return d;
    let m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
    m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    return null;
  }

  function fmtDateTime(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return "";
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }

  function computeCountdown(rows, cols) {
    const now = new Date();
    const payCol = cols.pay;
    const stCol = cols.orderStatus;
    const logCol = cols.logistics;
    return rows.map((r) => {
      const out = Object.assign({}, r);
      const st = normVal(r[stCol]).replace(/[\s\r\n]+/g, "");
      const logi = logCol ? normVal(r[logCol]) : "";
      let status = "", latest = null, remDays = "", remHours = "";
      // 已发货 但 物流状态含"无物流"时，视为未真正揽收，继续按倒计时监控
      const shipped = st === "已发货" || /已发货|已发出|已寄出|已交运/.test(st);
      if (shipped && !/无物流/.test(logi)) status = "已完成";
      else if (st === "已取消") status = "已取消";
      else {
        // 待发货 / 已发货但无物流 / 其他未明确发货的状态：按付款时间算倒计时预警
        const pay = parsePayTime(r[payCol]);
        if (!pay) status = "缺付款时间";
        else {
          latest = new Date(pay.getTime() + 1.5 * 86400000);
          const rem_h = (latest - now) / 3600000;
          remHours = rem_h.toFixed(1);
          remDays = (rem_h / 24).toFixed(1);
          if (rem_h < 0) status = "已超时";
          else if (rem_h < 12) status = "不足12小时";
          else status = "正常";
        }
      }
      out["最晚发货时间"] = latest ? fmtDateTime(latest) : "";
      out["剩余(天)"] = remDays;
      out["剩余(小时)"] = remHours;
      out["状态监控"] = status;
      return out;
    });
  }

  function top10(rows, supCol) {
    const counter = {};
    rows.forEach((r) => {
      const sup = normVal(r[supCol]) || "(未知)";
      const st = normVal(r["状态监控"]);
      if (!counter[sup]) counter[sup] = { 已超时: 0, 紧急: 0 };
      if (st === "已超时") counter[sup].已超时++;
      else if (st === "不足12小时") counter[sup].紧急++;
    });
    return Object.keys(counter)
      .map((sup) => ({ sup, 已超时: counter[sup].已超时, 紧急: counter[sup].紧急, risk: counter[sup].已超时 + counter[sup].紧急 }))
      .sort((a, b) => b.risk - a.risk)
      .slice(0, 10);
  }

  // 主流程：返回 { cols, step1, final, removed, dupRemoved, countdown, top10, warnings }
  function processOrders(orderRows, trackRows) {
    const warnings = [];
    if (!orderRows || !orderRows.length) {
      warnings.push("订单表为空或未识别到数据行。");
      return { cols: {}, step1: [], final: [], removed: 0, dupRemoved: 0, countdown: [], top10: [], warnings };
    }
    const columns = Object.keys(orderRows[0]);
    const cols = {
      supplier: findColSmart(columns, "供应商"),
      orderStatus: findColSmart(columns, "订单状态"),
      logistics: findColSmart(columns, "物流状态"),
      tracking: findColSmart(columns, "快递单号"),
      orderNo: findColSmart(columns, "订单号"),
      pay: findColSmart(columns, "付款时间"),
    };
    if (!cols.supplier) warnings.push("未识别“供应商”列，跳过供应商剔除。");
    if (!cols.orderStatus) warnings.push("未识别“订单状态”列，跳过订单状态筛选（默认保留全部）。");
    if (!cols.logistics) warnings.push("未识别“物流状态”列，跳过物流状态筛选（默认保留全部）。");
    if (!cols.tracking) warnings.push("未识别“快递单号”列，跳过去重。");

    // 读取防御：若某列识别成功，但 90% 以上行该列值为空，判定为解析异常，发出警告
    [["供应商", cols.supplier], ["订单状态", cols.orderStatus], ["物流状态", cols.logistics], ["付款时间", cols.pay]]
      .forEach(([name, col]) => {
        if (!col) return;
        let emptyN = 0;
        for (const r of orderRows) if (normVal(r[col]) === "") emptyN++;
        if (orderRows.length && emptyN / orderRows.length > 0.9) {
          warnings.push(`⚠️ 列「${name}」识别到但95%以上为空，可能文件解析异常，请检查源文件或改传 CSV。`);
        }
      });

    // 筛选（白名单，与 process_orders.py 一致）：供应商剔除 + 订单状态白名单 + 物流状态白名单
    const step1 = orderRows.filter((r) => {
      const okSup = cols.supplier ? !REMOVE_SUPPLIERS.has(normVal(r[cols.supplier])) : true;
      const okOs = cols.orderStatus ? KEEP_ORDER_STATUS.has(normVal(r[cols.orderStatus])) : true; // 白名单：仅保留 已发货/待发货
      const okLog = cols.logistics ? keepLogistics(r[cols.logistics]) : true;
      return okSup && okOs && okLog;
    });

    // 清洗标识列
    step1.forEach((r) => {
      if (cols.orderNo != null && r[cols.orderNo] != null) r[cols.orderNo] = cleanId(r[cols.orderNo]);
      if (cols.tracking != null && r[cols.tracking] != null) r[cols.tracking] = cleanId(r[cols.tracking]);
    });

    const trackSet = extractTrackingNumbers(trackRows);
    const final = cols.tracking
      ? step1.filter((r) => !trackSet.has(cleanId(r[cols.tracking]).toLowerCase()))
      : step1.slice();

    const countdown = computeCountdown(final, cols);
    const tp = top10(countdown, cols.supplier);

    // 被筛选剔除的行（附原因）
    const removedRows = orderRows
      .filter((r) => !step1.includes(r))
      .map((r) => {
        const reasons = [];
        if (cols.supplier && REMOVE_SUPPLIERS.has(normVal(r[cols.supplier]))) reasons.push("湖南供应商剔除");
        if (cols.orderStatus && !KEEP_ORDER_STATUS.has(normVal(r[cols.orderStatus]))) reasons.push("订单状态不符(非已发货/待发货)");
        if (cols.logistics && !keepLogistics(r[cols.logistics])) reasons.push("物流状态不符(非空且非无物流)");
        return Object.assign({}, r, { 剔除原因: reasons.join("；") || "未知" });
      });

    // 被去重剔除的行（附原因）
    const dupRows = cols.tracking
      ? step1
          .filter((r) => trackSet.has(cleanId(r[cols.tracking]).toLowerCase()))
          .map((r) => Object.assign({}, r, { 剔除原因: "快递单号已在快递表(去重)" }))
      : [];

    return {
      cols,
      step1,
      final,
      removed: orderRows.length - step1.length,
      dupRemoved: step1.length - final.length,
      countdown,
      top10: tp,
      removedRows,
      dupRows,
      warnings,
    };
  }

  // 简单计数（用于可视化）
  function countBy(rows, col) {
    const m = {};
    rows.forEach((r) => {
      const k = normVal(r[col]) || "(空)";
      m[k] = (m[k] || 0) + 1;
    });
    return m;
  }

  const Engine = {
    REMOVE_SUPPLIERS, KEEP_ORDER_STATUS, KEEP_LOGISTICS,
    KW_SUPPLIER, KW_ORDER_STATUS, KW_LOGISTICS, KW_TRACKING, KW_ORDER_NO,
    COL_ALIASES, normVal, cleanId, keepLogistics, findCol, findColSmart, validateColumns,
    extractTrackingNumbers,
    parsePayTime, fmtDateTime, computeCountdown, top10, processOrders, countBy,
  };

  if (typeof module !== "undefined" && module.exports) module.exports = Engine;
  else global.QEngine = Engine;
})(typeof window !== "undefined" ? window : globalThis);
