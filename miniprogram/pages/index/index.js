const stats = require('../../utils/stats.js');

Page({
  data: {
    apptCount: 0,
    orderCount: 0,
    rangeStart: '',
    rangeEnd: '',
    periodText: '',
    statusMsg: '请先导入《预约列表》与《订单项目收退款明细》两份 Excel 文件。',
    note: stats.NOTE_TEXT,
    hasData: false,
    // 三张表视图数据
    kvRows: [],
    perfRows: [],
    perfTotal: {},
    perfSummary: {},
    butlerRows: [],
    butlerTotal: {},
    // 导入缓存状态（一次上传，自动记住，可覆盖）
    apptCached: false,
    orderCached: false,
    vabCached: false,
    vabCount: 0,
    vabMeta: null
  },

  // ---- 非渲染状态（不放进 data，避免无谓的 setData）----
  APPT: [],
  ORDER: [],
  OV: { perf: {}, butler: {} },
  RANGE: { start: null, end: null },
  lastR: null,

  // ---------- 启动：自动加载本地缓存 ----------
  onLoad() {
    const appt = wx.getStorageSync('dt_appt');
    const order = wx.getStorageSync('dt_order');
    const vab = wx.getStorageSync('dt_vab');
    const vmeta = wx.getStorageSync('dt_vab_meta');
    const range = wx.getStorageSync('dt_range');
    if (appt && appt.length) { this.APPT = appt; this.setData({ apptCount: appt.length, apptCached: true }); }
    if (order && order.length) { this.ORDER = order; this.setData({ orderCount: order.length, orderCached: true }); }
    if (vab && vab.length) { this.VAB = vab; this.setData({ vabCount: vab.length, vabCached: true, vabMeta: vmeta || this.buildVabMeta(vab) }); }
    if (range && range.start && range.end) this.RANGE = range;
    const ovRaw = wx.getStorageSync('dt_ov');
    if (ovRaw && ovRaw.period === this.periodKey() && ovRaw.ov) this.OV = ovRaw.ov;
    if (this.APPT.length && this.ORDER.length) this.recompute();
  },

  // ---------- 导入（分类型，支持覆盖更新）----------
  chooseAppt() { this.pickOne('appt'); },
  chooseOrder() { this.pickOne('order'); },
  chooseVab() { this.pickOne('vab'); },
  pickOne(kind) {
    wx.chooseMessageFile({
      count: 1,
      type: 'file',
      extension: ['xlsx', 'xls'],
      success: (res) => { if (res.tempFiles && res.tempFiles[0]) this.importOne(kind, res.tempFiles[0]); },
      fail: () => {}
    });
  },

  readFileBuf(path) {
    return new Promise((resolve, reject) => {
      wx.getFileSystemManager().readFile({
        filePath: path,
        success: (r) => resolve(r.data),
        fail: reject
      });
    });
  },

  async importOne(kind, file) {
    wx.showLoading({ title: '解析中…', mask: true });
    try {
      const buf = await this.readFileBuf(file.path);
      const wb = stats.XLSX.read(buf, { type: 'array', cellDates: true });
      let data;
      if (kind === 'appt') data = stats.parseAppt(wb);
      else if (kind === 'order') data = stats.parseOrder(wb);
      else data = stats.parseVab(wb);
      if (!data || !data.length) {
        wx.showToast({ title: '未识别到有效数据，请检查列结构', icon: 'none' });
        return;
      }
      if (kind === 'appt') {
        this.APPT = data; wx.setStorageSync('dt_appt', data);
        this.setData({ apptCount: data.length, apptCached: true });
      } else if (kind === 'order') {
        this.ORDER = data; wx.setStorageSync('dt_order', data);
        this.setData({ orderCount: data.length, orderCached: true });
      } else {
        this.VAB = data; const meta = this.buildVabMeta(data);
        wx.setStorageSync('dt_vab', data); wx.setStorageSync('dt_vab_meta', meta);
        this.setData({ vabCount: data.length, vabCached: true, vabMeta: meta });
      }
      wx.showToast({ title: '已导入 ' + data.length + ' 行', icon: 'success' });
      if (this.APPT.length && this.ORDER.length) {
        const auto = stats.autoRange(this.APPT, this.ORDER);
        if (auto) { this.RANGE = auto; this.setData({ rangeStart: auto.start, rangeEnd: auto.end }); }
        this.OV = this.restoreOV();
        this.recompute();
      } else {
        this.setData({
          hasData: false,
          statusMsg: '还需导入《预约列表》与《订单明细》才能计算（VAB名单为可选附加）。当前：预约列表 ' + this.APPT.length + ' 行，订单明细 ' + this.ORDER.length + ' 行。'
        });
      }
    } catch (e) {
      wx.showToast({ title: '导入失败：' + (e.message || e), icon: 'none' });
    } finally {
      wx.hideLoading();
    }
  },

  buildVabMeta(data) {
    const groups = {};
    data.forEach((r) => { const g = r.group || '未分类'; groups[g] = (groups[g] || 0) + 1; });
    const groupsArr = Object.keys(groups).map((k) => ({ name: k, count: groups[k] })).sort((a, b) => b.count - a.count);
    return { time: this.fmtNow(), total: data.length, groups: groupsArr };
  },
  fmtNow() {
    const d = new Date();
    const p = (n) => (n < 10 ? '0' + n : '' + n);
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
  },
  clearCache() {
    wx.showModal({
      title: '清除本地缓存',
      content: '将清除已缓存的预约列表、订单明细、VAB名单与统计区间，下次需重新导入。',
      success: (res) => {
        if (!res.confirm) return;
        wx.removeStorageSync('dt_appt'); wx.removeStorageSync('dt_order');
        wx.removeStorageSync('dt_vab'); wx.removeStorageSync('dt_vab_meta'); wx.removeStorageSync('dt_range');
        wx.removeStorageSync('dt_ov');
        this.APPT = []; this.ORDER = []; this.VAB = []; this.RANGE = { start: null, end: null }; this.OV = { perf: {}, butler: {} };
        this.setData({
          apptCount: 0, orderCount: 0, vabCount: 0, apptCached: false, orderCached: false, vabCached: false,
          vabMeta: null, hasData: false, rangeStart: '', rangeEnd: '',
          statusMsg: '缓存已清除，请重新导入《预约列表》与《订单明细》。'
        });
      }
    });
  },

  // ---------- 区间 ----------
  onStartChange(e) { this.setData({ rangeStart: e.detail.value }); },
  onEndChange(e) { this.setData({ rangeEnd: e.detail.value }); },

  // ---------- 计算 ----------
  recompute() {
    if (!this.APPT.length || !this.ORDER.length) {
      this.setData({ hasData: false, statusMsg: '需同时导入《预约列表》与《订单项目收退款明细》两份文件后才能计算。' });
      return;
    }
    try {
      if (!this.RANGE.start || !this.RANGE.end) {
        const auto = stats.autoRange(this.APPT, this.ORDER);
        if (auto) { this.RANGE = auto; this.setData({ rangeStart: auto.start, rangeEnd: auto.end }); }
      }
      wx.setStorageSync('dt_range', this.RANGE);

      const pt = stats.periodTitle(this.RANGE); // 格式：M.D~M.D 12:00
      const R = stats.effectiveR(this.APPT, this.ORDER, this.RANGE, this.OV);
      this.lastR = R;
      this.mapToView(R);
      const kv = R.kv;
      this.setData({
        hasData: true,
        periodText: pt,
        statusMsg: '已导入 预约列表 ' + this.APPT.length + ' 行 / 订单明细 ' + this.ORDER.length + ' 行 · 数据区间 ' + pt +
          ' · 周客量合计到访 ' + kv['合计'].tot + ' 人，成交率 ' + stats.pctStr(kv['合计'].rate) +
          '；周业绩合计 ' + stats.money(R.perf.sum.total) + ' 元。'
      });
    } catch (err) {
      this.setData({ hasData: false, statusMsg: '计算失败：' + (err.message || err) });
    }
  },

  mapToView(R) {
    const d = R.depts; // ["营销部","商务部","其他(公司所有)"]
    const kv = R.kv;
    const kvRows = [
      { key: 'm', cat: '线上', dept: '营销部', tot: kv[d[0]].tot, chu: kv[d[0]].chu, fu: kv[d[0]].fu, cash: kv[d[0]].cash, rate: stats.pctStr(kv[d[0]].rate) },
      { key: 's', cat: '线下', dept: '商务部', tot: kv[d[1]].tot, chu: kv[d[1]].chu, fu: kv[d[1]].fu, cash: kv[d[1]].cash, rate: stats.pctStr(kv[d[1]].rate) },
      { key: 'o', cat: '', dept: '其他(公司所有)', tot: kv[d[2]].tot, chu: kv[d[2]].chu, fu: kv[d[2]].fu, cash: kv[d[2]].cash, rate: stats.pctStr(kv[d[2]].rate) },
      { key: 't', cat: '合计', dept: '', tot: kv['合计'].tot, chu: kv['合计'].chu, fu: kv['合计'].fu, cash: kv['合计'].cash, rate: stats.pctStr(kv['合计'].rate), isTotal: true }
    ];

    const perfRows = R.perf.rows.map((r) => ({
      date: r.date,
      dateLabel: r.date.replace(/-/g, '/') + '(' + stats.weekday(r.date) + ')',
      gao: stats.money(r.gao),
      jing: stats.money(r.jing),
      total: stats.money(r.total),
      rank: r.rank,
      visit: r.visit,
      book: r.book,
      rate: stats.pctStr(r.rate)
    }));
    const perfTotal = {
      gao: stats.money(R.perf.sum.gao),
      jing: stats.money(R.perf.sum.jing),
      total: stats.money(R.perf.sum.total),
      visit: R.perf.wkVisit,
      book: R.perf.wkBook,
      rate: stats.pctStr(R.perf.wkBook ? R.perf.wkVisit / R.perf.wkBook : 0)
    };
    const perfSummary = {
      maxDay: R.perf.maxDay ? (stats.fmtMD(R.perf.maxDay) + '(' + stats.weekday(R.perf.maxDay) + ')') : '—',
      minDay: R.perf.minDay ? (stats.fmtMD(R.perf.minDay) + '(' + stats.weekday(R.perf.minDay) + ')') : '—',
      avg: stats.money(R.perf.avg)
    };

    const butlerRows = R.butler.rows.map((r) => ({
      name: r.name,
      perf: stats.money(r.perf),
      rank: r.rank,
      book: r.book,
      visit: r.visit,
      rate: (r.rate == null) ? '/' : stats.pctStr(r.rate)
    }));
    const butlerTotal = {
      perf: stats.money(R.butler.sumP),
      book: R.butler.sumB,
      visit: R.butler.sumV,
      rate: stats.pctStr(R.butler.sumB ? R.butler.sumV / R.butler.sumB : 0)
    };

    this.setData({ kvRows, perfRows, perfTotal, perfSummary, butlerRows, butlerTotal });
  },

  // ---------- 手工调整（点击可编辑单元格）----------
  // ---------- 手工调整持久化（按数据区间 key，刷新/重新导入后仍在）----------
  periodKey() {
    return (this.RANGE.start && this.RANGE.end) ? (this.RANGE.start + '~' + this.RANGE.end) : '';
  },
  saveOV() {
    try { wx.setStorageSync('dt_ov', { period: this.periodKey(), ov: this.OV }); } catch (e) {}
  },
  restoreOV() {
    try {
      const raw = wx.getStorageSync('dt_ov');
      if (raw && raw.period === this.periodKey() && raw.ov) return raw.ov;
    } catch (e) {}
    return { perf: {}, butler: {} };
  },
  clearOV() {
    this.OV = { perf: {}, butler: {} };
    try { wx.removeStorageSync('dt_ov'); } catch (e) {}
    this.recompute();
  },

  editCell(e) {
    const ds = e.currentTarget.dataset;
    const table = ds.table, field = ds.field;
    const key = (table === 'perf') ? ds.date : ds.name;
    let cur = '';
    if (this.lastR) {
      if (table === 'perf') {
        const row = this.lastR.perf.rows.find((x) => x.date === key);
        if (row) cur = (field === 'gao') ? String(row.gao) : String(row.jing);
      } else {
        const row = this.lastR.butler.rows.find((x) => x.name === key);
        if (row) cur = String(row.perf);
      }
    }
    const label = (table === 'perf')
      ? (field === 'gao' ? '高新(元)' : '经开(元)')
      : '总业绩(元)';
    wx.showModal({
      title: label + ' — ' + key,
      editable: true,
      placeholderText: '输入金额，留空则清除手工调整',
      content: cur,
      success: (res) => {
        if (!res.confirm) return;
        const raw = (res.content || '').trim();
        if (raw === '') {
          if (table === 'perf') {
            if (this.OV.perf[key]) {
              delete this.OV.perf[key][field];
              if (Object.keys(this.OV.perf[key]).length === 0) delete this.OV.perf[key];
            }
          } else {
            delete this.OV.butler[key];
          }
        } else {
          let num;
          if (raw.charAt(0) === '=' && /[+*/-]/.test(raw.slice(1))) {
            // 算式运算（不使用 eval）：@ 表示当前单元格原值
            num = this.calcExpr(raw.slice(1), this.cellBase(table, key, field));
            if (isNaN(num)) num = 0;
          } else {
            const ns = raw.replace(/[,\s¥￥]/g, '').replace(/[^\d.\-]/g, '');
            num = parseFloat(ns);
            if (isNaN(num)) num = 0;
          }
          if (table === 'perf') {
            this.OV.perf[key] = this.OV.perf[key] || {};
            this.OV.perf[key][field] = num;
          } else {
            this.OV.butler[key] = this.OV.butler[key] || {};
            this.OV.butler[key].perf = num;
          }
        }
        this.recompute();
        this.saveOV();
      }
    });
  },

  // 安全算式求值（不使用 eval）：递归下降解析，支持 + - * / ( ) 及 @（当前单元格原值）
  calcExpr(expr, base) {
    expr = String(expr == null ? '' : expr).replace(/@/g, String(Number(base) || 0));
    if (!/^[\d\s.+*/()-]+$/.test(expr)) return NaN;
    const s = expr.replace(/\s+/g, '');
    if (s === '') return NaN;
    let pos = 0;
    const parseExpr = () => {
      let v = parseTerm();
      while (pos < s.length && (s.charAt(pos) === '+' || s.charAt(pos) === '-')) {
        const op = s.charAt(pos++);
        const r = parseTerm();
        v = op === '+' ? v + r : v - r;
      }
      return v;
    };
    const parseTerm = () => {
      let v = parseFactor();
      while (pos < s.length && (s.charAt(pos) === '*' || s.charAt(pos) === '/')) {
        const op = s.charAt(pos++);
        const r = parseFactor();
        v = op === '*' ? v * r : (r === 0 ? 0 : v / r);
      }
      return v;
    };
    const parseFactor = () => {
      if (pos >= s.length) return NaN;
      const ch = s.charAt(pos);
      if (ch === '+') { const p = s.charAt(pos - 1); if (p === '+' || p === '*' || p === '/') return NaN; pos++; return parseFactor(); }
      if (ch === '-') { pos++; return -parseFactor(); }
      if (ch === '(') { pos++; const v = parseExpr(); if (s.charAt(pos) !== ')') return NaN; pos++; return v; }
      const m = s.slice(pos).match(/^\d+\.?\d*/);
      if (!m) return NaN;
      pos += m[0].length;
      return parseFloat(m[0]);
    };
    const result = parseExpr();
    if (pos < s.length) return NaN;
    return isFinite(result) ? result : NaN;
  },
  // 取当前单元格有效数值（@ 引用）
  cellBase(table, key, field) {
    const R = this.lastR;
    if (!R) return 0;
    if (table === 'perf') {
      const row = R.perf.rows.find((x) => x.date === key);
      return row ? Number(row[field]) || 0 : 0;
    }
    const b = R.butler.rows.find((x) => x.name === key);
    return b ? Number(b.perf) || 0 : 0;
  },

  // ---------- 导出 Excel ----------
  exportExcel() {
    if (!this.lastR) { wx.showToast({ title: '请先导入并计算', icon: 'none' }); return; }
    const XLSX = stats.XLSX;
    const R = this.lastR;
    const pt = this.data.periodText;
    const kv = R.kv;

    const kvAoa = [
      ['周客量统计(数据区间：' + pt + ')'],
      ['部门', '总到访客量', '初诊客量', '复诊客量', '有成交客量', '成交率'],
      ['线上', '营销部', kv['营销部'].tot, kv['营销部'].chu, kv['营销部'].fu, kv['营销部'].cash, stats.pctStr(kv['营销部'].rate)],
      ['线下', '商务部', kv['商务部'].tot, kv['商务部'].chu, kv['商务部'].fu, kv['商务部'].cash, stats.pctStr(kv['商务部'].rate)],
      ['', '其他(公司所有)', kv['其他(公司所有)'].tot, kv['其他(公司所有)'].chu, kv['其他(公司所有)'].fu, kv['其他(公司所有)'].cash, stats.pctStr(kv['其他(公司所有)'].rate)],
      ['合计', '', kv['合计'].tot, kv['合计'].chu, kv['合计'].fu, kv['合计'].cash, stats.pctStr(kv['合计'].rate)]
    ];

    const perfAoa = [
      ['周业绩统计(数据区间：' + pt + ')'],
      ['日期', '高新(元)', '经开(元)', '总业绩(元)', '业绩排名', '总到访客量', '总预约客量', '到访率']
    ];
    R.perf.rows.forEach((r) => {
      perfAoa.push([
        r.date.replace(/-/g, '/') + '(' + stats.weekday(r.date) + ')',
        r.gao, r.jing, r.total, r.rank, r.visit, r.book, stats.pctStr(r.rate)
      ]);
    });
    perfAoa.push(['合计', R.perf.sum.gao, R.perf.sum.jing, R.perf.sum.total, '/', R.perf.wkVisit, R.perf.wkBook, stats.pctStr(R.perf.wkBook ? R.perf.wkVisit / R.perf.wkBook : 0)]);
    perfAoa.push(['业绩最高值 ' + stats.fmtMD(R.perf.maxDay) + '　业绩最低值 ' + stats.fmtMD(R.perf.minDay) + '　业绩平均值：' + stats.money(R.perf.avg) + ' 元']);

    const butlerAoa = [
      ['管家业绩统计(数据区间：' + pt + ')'],
      ['姓名', '总业绩(元)', '业绩排名', '总预约量', '总到访量', '到访率']
    ];
    R.butler.rows.forEach((r) => {
      butlerAoa.push([r.name, r.perf, r.rank, r.book, r.visit, (r.rate == null) ? '/' : stats.pctStr(r.rate)]);
    });
    butlerAoa.push(['合计', R.butler.sumP, '/', R.butler.sumB, R.butler.sumV, stats.pctStr(R.butler.sumB ? R.butler.sumV / R.butler.sumB : 0)]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(kvAoa), '周客量统计');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(perfAoa), '周业绩统计');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(butlerAoa), '管家业绩统计');

    const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const fs = wx.getFileSystemManager();
    const safe = (pt || '导出').replace(/[^\d~/]/g, '');
    const filePath = `${wx.env.USER_DATA_PATH}/周客量业绩统计_${safe}.xlsx`;
    fs.writeFile({
      filePath: filePath,
      data: out,
      success: () => {
        wx.openDocument({ filePath: filePath, fileType: 'xlsx', showMenu: true, fail: () => wx.showToast({ title: '已生成，但预览失败', icon: 'none' }) });
      },
      fail: (e) => wx.showToast({ title: '导出失败：' + (e.message || e), icon: 'none' })
    });
  },

  // ---------- 导出 PNG（画布绘制三表）----------
  exportPng() {
    if (!this.lastR) { wx.showToast({ title: '请先导入并计算', icon: 'none' }); return; }
    const q = wx.createSelectorQuery();
    q.select('#exportCanvas').fields({ node: true, size: true }).exec((res) => {
      if (!res || !res[0] || !res[0].node) { wx.showToast({ title: '画布初始化失败', icon: 'none' }); return; }
      try { this.drawAndSave(res[0].node); }
      catch (e) { wx.showToast({ title: '导出图片失败：' + (e.message || e), icon: 'none' }); }
    });
  },

  drawAndSave(canvas) {
    const R = this.lastR;
    const dpr = (wx.getSystemInfoSync && wx.getSystemInfoSync().pixelRatio) || 2;
    const NAVY = '#01418B', BLACK = '#000000';
    const fsBase = 13;
    const padX = 16, lineH = 30, titleH = 34, gap = 16;

    // 列定义：w 为像素宽
    const tables = [
      {
        title: '周客量统计（数据区间：' + this.data.periodText + '）',
        head: ['类别', '部门', '总到访客量', '初诊客量', '复诊客量', '有成交客量', '成交率'],
        cols: [60, 110, 90, 80, 80, 90, 70],
        rows: [
          ['线上', '营销部', R.kv['营销部'].tot, R.kv['营销部'].chu, R.kv['营销部'].fu, R.kv['营销部'].cash, stats.pctStr(R.kv['营销部'].rate)],
          ['线下', '商务部', R.kv['商务部'].tot, R.kv['商务部'].chu, R.kv['商务部'].fu, R.kv['商务部'].cash, stats.pctStr(R.kv['商务部'].rate)],
          ['', '其他(公司所有)', R.kv['其他(公司所有)'].tot, R.kv['其他(公司所有)'].chu, R.kv['其他(公司所有)'].fu, R.kv['其他(公司所有)'].cash, stats.pctStr(R.kv['其他(公司所有)'].rate)],
          ['合计', '', R.kv['合计'].tot, R.kv['合计'].chu, R.kv['合计'].fu, R.kv['合计'].cash, stats.pctStr(R.kv['合计'].rate), true]
        ]
      },
      {
        title: '周业绩统计（数据区间：' + this.data.periodText + '）',
        head: ['日期', '高新(元)', '经开(元)', '总业绩(元)', '业绩排名', '总到访客量', '总预约客量', '到访率'],
        cols: [130, 80, 80, 100, 70, 90, 90, 70],
        rows: R.perf.rows.map((r) => [
          r.date.replace(/-/g, '/') + '(' + stats.weekday(r.date) + ')', stats.money(r.gao), stats.money(r.jing), stats.money(r.total), r.rank, r.visit, r.book, stats.pctStr(r.rate)
        ]).concat([[
          '合计', stats.money(R.perf.sum.gao), stats.money(R.perf.sum.jing), stats.money(R.perf.sum.total), '/', R.perf.wkVisit, R.perf.wkBook, stats.pctStr(R.perf.wkBook ? R.perf.wkVisit / R.perf.wkBook : 0), true
        ]]).concat([[
          '业绩最高值 ' + stats.fmtMD(R.perf.maxDay) + '　业绩最低值 ' + stats.fmtMD(R.perf.minDay) + '　业绩平均值：' + stats.money(R.perf.avg) + ' 元'
        ]])
      },
      {
        title: '管家业绩统计（数据区间：' + this.data.periodText + '）',
        head: ['姓名', '总业绩(元)', '业绩排名', '总预约量', '总到访量', '到访率'],
        cols: [100, 100, 80, 90, 90, 80],
        rows: R.butler.rows.map((r) => [
          r.name, stats.money(r.perf), r.rank, r.book, r.visit, (r.rate == null) ? '/' : stats.pctStr(r.rate)
        ]).concat([[
          '合计', stats.money(R.butler.sumP), '/', R.butler.sumB, R.butler.sumV, stats.pctStr(R.butler.sumB ? R.butler.sumV / R.butler.sumB : 0), true
        ]])
      }
    ];

    // 计算每张表宽度与总高度
    let totalW = 0;
    tables.forEach((t) => { totalW = Math.max(totalW, t.cols.reduce((a, b) => a + b, 0) + padX * 2); });
    let totalH = padX;
    tables.forEach((t) => { totalH += titleH + lineH * (t.head.length ? (t.rows.length + 1) : t.rows.length) + gap; });
    totalH += padX;

    canvas.width = totalW * dpr;
    canvas.height = totalH * dpr;
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, totalW, totalH);
    ctx.textBaseline = 'middle';

    let y = padX;
    tables.forEach((t) => {
      // 标题
      ctx.fillStyle = NAVY;
      ctx.font = 'bold ' + (fsBase + 2) + 'px sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText(t.title, padX, y + titleH / 2);
      y += titleH;

      const totalRowW = t.cols.reduce((a, b) => a + b, 0);
      const x0 = padX + (totalW - padX * 2 - totalRowW) / 2;

      const drawRow = (cells, isHead, isTotal) => {
        let x = x0;
        ctx.font = (isHead || isTotal ? 'bold ' : '') + fsBase + 'px sans-serif';
        for (let c = 0; c < t.cols.length; c++) {
          const w = t.cols[c];
          const text = (cells[c] == null ? '' : String(cells[c]));
          if (isHead) { ctx.fillStyle = NAVY; ctx.fillRect(x, y, w, lineH); }
          else if (isTotal) { ctx.fillStyle = '#e7eefb'; ctx.fillRect(x, y, w, lineH); }
          ctx.strokeStyle = BLACK; ctx.lineWidth = 1;
          ctx.strokeRect(x + 0.5, y + 0.5, w, lineH);
          ctx.fillStyle = (isHead || isTotal) ? '#000000' : '#1f2733';
          if (isHead) ctx.fillStyle = '#ffffff';
          ctx.textAlign = 'center';
          ctx.fillText(text, x + w / 2, y + lineH / 2);
          x += w;
        }
        y += lineH;
      };

      drawRow(t.head, true, false);
      t.rows.forEach((row) => { drawRow(row, false, row[row.length - 1] === true); });
      y += gap;
    });

    wx.canvasToTempFilePath({
      canvas: canvas,
      success: (r) => {
        wx.previewImage({
          urls: [r.tempFilePath],
          current: r.tempFilePath,
          fail: () => wx.showToast({ title: '已生成图片，预览失败', icon: 'none' })
        });
      },
      fail: (e) => wx.showToast({ title: '导出图片失败：' + (e.message || e), icon: 'none' })
    });
  }
});
