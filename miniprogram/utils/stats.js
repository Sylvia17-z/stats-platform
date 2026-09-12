// utils/stats.js
// 移植自数据统计平台 app.js 的纯计算逻辑（无 DOM 依赖），供微信小程序调用。
// 依赖 SheetJS：npm i xlsx 后「构建 npm」；或把 xlsx.full.min.js 放到本目录并改成 require('./xlsx.full.min.js')。

try {
  // 迷你程序运行环境通常没有 process，给 SheetJS 一个最小垫片，避免它初始化时报错
  if (typeof process === 'undefined') {
    var _g = (typeof global !== 'undefined') ? global : (typeof this !== 'undefined' ? this : {});
    _g.process = { browser: true, env: {} };
  }
} catch (e) {}

// 优先用本地打包的 xlsx.full.min.js（无需「构建 npm」即可运行），缺失时回退到 npm 包
let XLSX;
try { XLSX = require('./xlsx.full.min.js'); } catch (e) {
  try { XLSX = require('xlsx'); } catch (e2) { throw new Error('未找到 xlsx：请把 xlsx.full.min.js 放到 utils/ 下，或在开发者工具执行「构建 npm」'); }
}

const NOTE_TEXT = "统计说明：数据来源 领健系统每周一至周六数据（截止时间：周六 12:00）\n1、初诊客量~周区间内到访类型为初诊 复诊客量~周区间内到访类型为复诊/疗程内/再消费/复查/为空\n2、成交率 现款支付不为0并且不是退款的消费客量 占总客量的比例\n3、如果一个客户 周一来体检 周三来解读报告 计算客量为1";

const HI = ['南媛', '系统管理员'];

function pad(n) { return (n < 10 ? "0" : "") + n; }
function parseDateVal(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) return v.getFullYear() + "-" + pad(v.getMonth() + 1) + "-" + pad(v.getDate());
  var s = String(v).replace(/\//g, "-").replace("T", " ").trim();
  var m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? (m[1] + "-" + pad(+m[2]) + "-" + pad(+m[3])) : null;
}
function money(x) { return (Math.round((x || 0) * 100) / 100).toFixed(2); }
function pctStr(r) { return (r * 100).toFixed(0) + "%"; }
function weekday(d) {
  if (!d) return "";
  var p = d.split("-");
  var WD = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
  return WD[new Date(+p[0], +p[1] - 1, +p[2]).getDay()];
}
function findCol(headers, keys) {
  var best = -1, bestScore = 0;
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i] != null ? String(headers[i]) : "";
    var score = 0; for (var k = 0; k < keys.length; k++) { if (h.indexOf(keys[k]) >= 0) score++; }
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return bestScore > 0 ? best : -1;
}
function deptAppt(src) {
  var s = String(src || "");
  if (s.indexOf("营销部") >= 0) return "营销部";
  if (s.indexOf("市场部") >= 0 || s.indexOf("商务部") >= 0) return "商务部";
  return "其他(公司所有)";
}
function deptOrder(src) {
  var s = String(src || "");
  if (s.indexOf("营销部") >= 0) return "营销部";
  if (s.indexOf("市场部") >= 0 || s.indexOf("商务部") >= 0) return "商务部";
  return "其他(公司所有)";
}
// 在所有 sheet × 前几行里找关键列命中最全的一组作为表头（领健导出结构不固定）
function scanSheets(wb, colDefs, maxScan) {
  var best = null;
  for (var s = 0; s < wb.SheetNames.length; s++) {
    var name = wb.SheetNames[s], ws = wb.Sheets[name];
    if (!ws) continue;
    var aoa;
    try { aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }); } catch (e) { continue; }
    for (var i = 0; i < Math.min(aoa.length, maxScan || 6); i++) {
      var row = aoa[i]; if (!row) continue;
      var H = row.map(function (c) { return c == null ? "" : String(c); });
      var ci = {}, hit = 0, reqFail = false;
      colDefs.forEach(function (d) {
        var idx = findCol(H, d.names);
        ci[d.key] = idx;
        if (idx >= 0) hit++; else if (d.req) reqFail = true;
      });
      if (reqFail) continue;
      if (!best || hit > best.hit) { best = { ws: ws, ci: ci, hrow: i, hit: hit }; }
    }
  }
  return best;
}
function parseAppt(wb) {
  var b = scanSheets(wb, [
    { key: "date", names: ["日期"] },
    { key: "visit", names: ["是否到访"], req: true },
    { key: "mem", names: ["会员号"], req: true },
    { key: "type", names: ["预约类型"] },
    { key: "source", names: ["客户来源"], req: true },
    { key: "butler", names: ["预约健康管家"] }
  ]);
  if (!b) throw new Error("预约列表缺少关键列（会员号/是否到访/客户来源）");
  var aoa = XLSX.utils.sheet_to_json(b.ws, { header: 1, defval: null, raw: true });
  var ci = b.ci, out = [];
  for (var r = b.hrow + 1; r < aoa.length; r++) {
    var row = aoa[r]; if (!row) continue;
    var get = function (i) { return (i >= 0 && i < row.length) ? row[i] : null; };
    var mem = get(ci.mem); if (mem == null || String(mem).trim() === "") continue;
    out.push({
      date: parseDateVal(get(ci.date)), visit: String(get(ci.visit) || "").trim(),
      mem: String(mem).trim(), type: String(get(ci.type) || "").trim(),
      source: String(get(ci.source) || "").trim(), butler: String(get(ci.butler) || "").trim()
    });
  }
  return out;
}
function parseOrder(wb) {
  var b = scanSheets(wb, [
    { key: "mem", names: ["会员号"], req: true },
    { key: "consultant", names: ["开单咨询师"], req: true },
    { key: "cash", names: ["现款支付"], req: true },
    { key: "dept", names: ["客户来源渠道分类", "渠道分类"] },
    { key: "time", names: ["收款时间"] },
    { key: "emp", names: ["收款员工"] },
    { key: "type", names: ["收款类型"] }
  ]);
  if (!b) throw new Error("订单明细缺少关键列（会员号/开单咨询师/现款支付）");
  var aoa = XLSX.utils.sheet_to_json(b.ws, { header: 1, defval: null, raw: true });
  var ci = b.ci, out = [];
  for (var r = b.hrow + 1; r < aoa.length; r++) {
    var row = aoa[r]; if (!row) continue;
    var get = function (i) { return (i >= 0 && i < row.length) ? row[i] : null; };
    var mem = get(ci.mem); if (mem == null || String(mem).trim() === "") continue;
    var cash = get(ci.cash); cash = (typeof cash === "number") ? cash : (parseFloat(cash) || 0);
    out.push({
      date: parseDateVal(get(ci.time)), mem: String(mem).trim(),
      consultant: String(get(ci.consultant) || "").trim(),
      deptSrc: String(get(ci.dept) || "").trim(), cash: cash,
      employee: String(get(ci.emp) || "").trim(),
      type: (ci.type >= 0 ? String(get(ci.type) || "") : "").trim()
    });
  }
  return out;
}
function inRange(d, RANGE) { return (!RANGE.start || (d && d >= RANGE.start)) && (!RANGE.end || (d && d <= RANGE.end)); }
function compute(APPT, ORDER, RANGE) {
  var depts = ["营销部", "商务部", "其他(公司所有)"];
  var kv = {}; depts.forEach(function (d) { kv[d] = { tot: new Set(), chu: new Set(), cash: new Set() }; });
  APPT.forEach(function (r) {
    if (r.visit !== "是" || !inRange(r.date, RANGE)) return;
    var d = deptAppt(r.source); if (!kv[d]) return;
    kv[d].tot.add(r.mem);
    if (r.type === "初诊") kv[d].chu.add(r.mem);
  });
  ORDER.forEach(function (r) {
    if (r.cash === 0 || r.type === "退款" || !inRange(r.date, RANGE)) return;
    var d = deptOrder(r.deptSrc); if (!kv[d]) return;
    kv[d].cash.add(r.mem);
  });
  var kvR = {};
  depts.forEach(function (d) {
    var tot = kv[d].tot.size, chu = kv[d].chu.size, cash = kv[d].cash.size;
    kvR[d] = { tot: tot, chu: chu, fu: tot - chu, cash: cash, rate: tot ? cash / tot : 0 };
  });
  var gt = 0, gc = 0, gf = 0, gk = 0;
  depts.forEach(function (d) { gt += kvR[d].tot; gc += kvR[d].chu; gf += kvR[d].fu; gk += kvR[d].cash; });
  kvR["合计"] = { tot: gt, chu: gc, fu: gf, cash: gk, rate: gt ? gk / gt : 0 };

  var gaoDay = {}, jingDay = {}, visitDay = {}, bookDay = {};
  ORDER.forEach(function (r) {
    if (!inRange(r.date, RANGE)) return;
    var add = Math.round(r.cash * 100) / 100;
    if (HI.indexOf(r.employee) >= 0) gaoDay[r.date] = Math.round(((gaoDay[r.date] || 0) + add) * 100) / 100;
    else jingDay[r.date] = Math.round(((jingDay[r.date] || 0) + add) * 100) / 100;
  });
  APPT.forEach(function (r) {
    if (!inRange(r.date, RANGE) || !r.mem) return;
    (bookDay[r.date] = bookDay[r.date] || new Set()).add(r.mem);
    if (r.visit === "是") (visitDay[r.date] = visitDay[r.date] || new Set()).add(r.mem);
  });
  var daySet = {}, k;
  [gaoDay, jingDay, visitDay, bookDay].forEach(function (o) { for (k in o) daySet[k] = 1; });
  var days = Object.keys(daySet).sort();
  var perfRows = days.map(function (d) {
    var g = gaoDay[d] || 0, j = jingDay[d] || 0;
    var va = (visitDay[d] || new Set()).size, vb = (bookDay[d] || new Set()).size;
    return { date: d, gao: g, jing: j, total: g + j, rank: 0, visit: va, book: vb, rate: vb ? va / vb : 0 };
  });
  var sorted = perfRows.slice().sort(function (a, b) { return b.total - a.total; });
  sorted.forEach(function (r, i) { for (var x = 0; x < perfRows.length; x++) { if (perfRows[x].date === r.date) { perfRows[x].rank = i + 1; break; } } });
  var wkVisit = new Set(), wkBook = new Set();
  APPT.forEach(function (r) { if (!inRange(r.date, RANGE) || !r.mem) return; wkBook.add(r.mem); if (r.visit === "是") wkVisit.add(r.mem); });
  var pSum = { gao: 0, jing: 0, total: 0 };
  perfRows.forEach(function (r) { pSum.gao += r.gao; pSum.jing += r.jing; pSum.total += r.total; });
  var maxDay = "", minDay = "", maxV = -1, minV = 1e18;
  perfRows.forEach(function (r) { if (r.total > maxV) { maxV = r.total; maxDay = r.date; } if (r.total < minV) { minV = r.total; minDay = r.date; } });
  var avg = perfRows.length ? pSum.total / perfRows.length : 0;

  var perfByName = {};
  ORDER.forEach(function (r) {
    if (!inRange(r.date, RANGE)) return;
    var add = Math.round(r.cash * 100) / 100;
    var c = r.consultant || "";
    // 管家业绩纯按"开单咨询师"分组；仅王勤红/空咨询师归入"无"
    var key = (c === "王勤红" || c === "") ? "无" : c;
    perfByName[key] = Math.round(((perfByName[key] || 0) + add) * 100) / 100;
  });
  var bBook = {}, bVisit = {};
  APPT.forEach(function (r) {
    if (!inRange(r.date, RANGE) || !r.mem) return; if (!r.butler) return;
    (bBook[r.butler] = bBook[r.butler] || new Set()).add(r.mem);
    if (r.visit === "是") (bVisit[r.butler] = bVisit[r.butler] || new Set()).add(r.mem);
  });
  var named = Object.keys(perfByName).filter(function (n) { return n !== "无"; }).sort(function (a, b) { return perfByName[b] - perfByName[a]; });
  var namedBook = 0, namedVisit = 0;
  named.forEach(function (n) { namedBook += (bBook[n] ? bBook[n].size : 0); namedVisit += (bVisit[n] ? bVisit[n].size : 0); });
  var noBook = Math.max(0, wkBook.size - namedBook), noVisit = Math.max(0, wkVisit.size - namedVisit);
  var butlerRows = named.map(function (n) {
    var bk = bBook[n] ? bBook[n].size : 0, vt = bVisit[n] ? bVisit[n].size : 0;
    return { name: n, perf: perfByName[n], rank: 0, book: bk, visit: vt, rate: bk ? vt / bk : null };
  });
  butlerRows.forEach(function (r, i) { r.rank = i + 1; });
  butlerRows.push({ name: "无", perf: perfByName["无"] || 0, rank: "/", book: noBook, visit: noVisit, rate: noBook ? noVisit / noBook : null });
  var bSumP = 0, bSumB = 0, bSumV = 0;
  butlerRows.forEach(function (r) { bSumP += r.perf; bSumB += r.book; bSumV += r.visit; });

  return {
    kv: kvR, depts: depts,
    perf: { rows: perfRows, sum: pSum, maxDay: maxDay, minDay: minDay, avg: avg, wkVisit: wkVisit.size, wkBook: wkBook.size },
    butler: { rows: butlerRows, sumP: bSumP, sumB: bSumB, sumV: bSumV }
  };
}
// 应用手工调整覆盖并重算联动指标（排名/合计/最高最低平均）；OV 结构与网页版一致
function effectiveR(APPT, ORDER, RANGE, OV) {
  OV = OV || {};
  var R = compute(APPT, ORDER, RANGE);
  R.perf.rows.forEach(function (r) {
    var o = OV.perf && OV.perf[r.date];
    if (o) { if (o.gao != null) r.gao = o.gao; if (o.jing != null) r.jing = o.jing; r.total = r.gao + r.jing; }
  });
  var ps = R.perf.rows.slice().sort(function (a, b) { return b.total - a.total; });
  ps.forEach(function (r, i) { for (var x = 0; x < R.perf.rows.length; x++) { if (R.perf.rows[x].date === r.date) { R.perf.rows[x].rank = i + 1; break; } } });
  var sg = 0, sj = 0, st = 0; R.perf.rows.forEach(function (r) { sg += r.gao; sj += r.jing; st += r.total; });
  R.perf.sum = { gao: sg, jing: sj, total: st };
  var maxDay = "", minDay = "", maxV = -1, minV = 1e18;
  R.perf.rows.forEach(function (r) { if (r.total > maxV) { maxV = r.total; maxDay = r.date; } if (r.total < minV) { minV = r.total; minDay = r.date; } });
  R.perf.avg = R.perf.rows.length ? st / R.perf.rows.length : 0; R.perf.maxDay = maxDay; R.perf.minDay = minDay;
  R.butler.rows.forEach(function (r) { var o = OV.butler && OV.butler[r.name]; if (o && o.perf != null) r.perf = o.perf; });
  var bs = R.butler.rows.filter(function (r) { return r.name !== "无"; }).slice().sort(function (a, b) { return b.perf - a.perf; });
  bs.forEach(function (r, i) { for (var x = 0; x < R.butler.rows.length; x++) { if (R.butler.rows[x].name === r.name) { R.butler.rows[x].rank = i + 1; break; } } });
  var sp = 0; R.butler.rows.forEach(function (r) { sp += r.perf; }); R.butler.sumP = sp;
  return R;
}

// 给页面用的格式化工具（money/pctStr/weekday 已在文件上方定义）
function fmtMD(d) { if (!d) return ""; var p = d.split("-"); return (+p[1]) + "." + (+p[2]); }
// 三表标题用的区间文本（带截止时间后缀）
function periodTitle(RANGE) {
  if (!RANGE || !RANGE.start || !RANGE.end) return "";
  return fmtMD(RANGE.start) + "~" + fmtMD(RANGE.end) + " 12:00";
}
function escapeHtml(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

// VAB 会员名单：会员号 + 组别（健管1/2/3组、营养科、女性健康中心等），用于缓存与组别分布
function parseVab(wb) {
  var b = scanSheets(wb, [
    { key: "mem", names: ["会员号", "会员编号", "卡号"], req: true },
    { key: "group", names: ["组别", "分组", "部门", "健管组", "类型"] }
  ]);
  if (!b) throw new Error("VAB名单缺少关键列（会员号）");
  var aoa = XLSX.utils.sheet_to_json(b.ws, { header: 1, defval: null, raw: true });
  var ci = b.ci, out = [];
  for (var r = b.hrow + 1; r < aoa.length; r++) {
    var row = aoa[r]; if (!row) continue;
    var get = function (i) { return (i >= 0 && i < row.length) ? row[i] : null; };
    var mem = get(ci.mem); if (mem == null || String(mem).trim() === "") continue;
    out.push({ mem: String(mem).trim(), group: (ci.group >= 0 ? String(get(ci.group) || "").trim() : "") });
  }
  return out;
}

// 依据内容自动识别：订单明细含「开单咨询师/现款支付」，预约列表含「是否到访」
function classify(wb) {
  try { return { kind: "order", data: parseOrder(wb) }; } catch (e) { }
  try { return { kind: "appt", data: parseAppt(wb) }; } catch (e) { }
  return { kind: null };
}
// 从数据里取最小/最大日期作为默认区间
function autoRange(APPT, ORDER) {
  var min = null, max = null;
  APPT.concat(ORDER).forEach(function (r) {
    if (!r.date) return;
    if (min == null || r.date < min) min = r.date;
    if (max == null || r.date > max) max = r.date;
  });
  return min ? { start: min, end: max } : null;
}

module.exports = {
  NOTE_TEXT: NOTE_TEXT, XLSX: XLSX,
  parseAppt: parseAppt, parseOrder: parseOrder, parseVab: parseVab, compute: compute, effectiveR: effectiveR,
  classify: classify, autoRange: autoRange,
  parseDateVal: parseDateVal, money: money, pctStr: pctStr, weekday: weekday, fmtMD: fmtMD, periodTitle: periodTitle, escapeHtml: escapeHtml
};
