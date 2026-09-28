/**
 * 数据采集明细 —— 数据采集管理 > 业务界面 > 数据采集明细
 *
 * 采集流水视角：一行 = 一条采集到的参数值。
 *   查询条件：数据时间（范围查询）+ 数据年度 + 数据月度（都可清空，留空 = 不限）
 *   列表按钮：只有「导出」→ 导出当前列表（筛选后）的数据为 Excel
 *   列表列：采集时间(年月日时分秒) / 数据年度 / 数据月度 / 生产线类型 / 生产线名称 /
 *           物料名称 / 参数名称 / 数据值 / 参数单位
 *
 * 数据落在 localStorage 键 DC_ROW_KEY。
 *
 * 口径说明：
 *   生产线类型 —— 用户给定清单（全厂 / 掺烧自产二次能源的化石燃料发电设施 / 转炉炼钢工序 /
 *                 电炉炼钢工序 / 炼铁工序 / 球团工序 / 烧结工序 / 焦化工序）；
 *   物料名称   —— 钢铁行业化石燃料标准口径（与 js/report-config-store.js 的 RC_MATERIALS 同源）；
 *   参数名称   —— 消耗量 / 产量。
 */

/* ---------- 常量 ---------- */

/** 行数据存储键（本项目内独立的一套，不与其他业务线共享）
 * v2：气体类物料的计量单位从「万m³」改成「10⁴Nm³」（数量级一样，但需求要求用后一种写法）。
 *     unit 是 dcRowsValid 的校验字段之一，但只校验非空 —— 旧数据能过校验、单位却是旧的，
 *     所以必须升键，让演示数据重新播种。
 */
const DC_ROW_KEY = 'emission-mgmt-collection-detail-rows-v2';

/** 生产线类型（用户给定清单） */
const DC_LINE_TYPES = [
  '全厂',
  '掺烧自产二次能源的化石燃料发电设施',
  '转炉炼钢工序',
  '电炉炼钢工序',
  '炼铁工序',
  '球团工序',
  '烧结工序',
  '焦化工序',
];

/**
 * 物料名称 —— 钢铁行业化石燃料标准口径。
 * 与 js/report-config-store.js 的 RC_MATERIALS 同源（引用它，改口径只改一处）。
 */
const DC_MATERIALS = RC_MATERIALS;

/** 参数名称 */
const DC_PARAMS = ['消耗量', '产量'];

/**
 * 生产线类型 → 生产线名称（演示数据用）：
 * 每个类型下给几条具体生产线，让「类型 / 名称」两级看起来像真的。
 */
const DC_LINE_NAMES = {
  '全厂': ['全厂汇总'],
  '掺烧自产二次能源的化石燃料发电设施': ['1#掺烧发电机组', '2#掺烧发电机组'],
  '转炉炼钢工序': ['1#转炉', '2#转炉', '3#转炉'],
  '电炉炼钢工序': ['1#电炉', '2#电炉'],
  '炼铁工序': ['1#高炉', '2#高炉', '3#高炉'],
  '球团工序': ['1#球团回转窑'],
  '烧结工序': ['1#烧结机', '2#烧结机'],
  '焦化工序': ['1#焦炉', '2#焦炉', '3#焦炉'],
};

/** 演示数据：按「生产线上有产能的才消耗对应物料」给映射，避免出现「球团工序烧焦炉煤气」这种假数据 */
const DC_TYPE_MATERIALS = {
  '全厂': ['无烟煤', '烟煤', '洗精煤', '焦炭', '焦炉煤气', '高炉煤气', '转炉煤气', '天然气'],
  '掺烧自产二次能源的化石燃料发电设施': ['高炉煤气', '转炉煤气', '焦炉煤气'],
  '转炉炼钢工序': ['焦炉煤气', '转炉煤气'],
  '电炉炼钢工序': ['天然气', '焦炉煤气'],
  '炼铁工序': ['焦炭', '无烟煤', '高炉煤气'],
  '球团工序': ['无烟煤', '天然气'],
  '烧结工序': ['无烟煤', '焦炭', '高炉煤气'],
  '焦化工序': ['洗精煤', '高炉煤气'],
};

/**
 * 各物料的常用计量单位（演示数据用）。
 * 气体类（焦炉煤气 / 高炉煤气 / 转炉煤气 / 天然气）统一用 10⁴Nm³
 * —— 煤气计量按标准立方米（Nm³）统计，量级大所以用万倍（10⁴）。
 */
const DC_MATERIAL_UNITS = {
  '无烟煤': 't', '烟煤': 't', '炼焦煤': 't', '洗精煤': 't', '其他洗煤': 't',
  '型煤': 't', '煤矸石': 't', '焦炭': 't', '石油焦': 't',
  '焦炉煤气': '10⁴Nm³', '高炉煤气': '10⁴Nm³', '转炉煤气': '10⁴Nm³',
  '原油': 't', '汽油': 't', '柴油': 't', '燃料油': 't',
  '液化石油气': 't', '天然气': '10⁴Nm³',
  '其他石油制品': 't', '其他焦化产品': 't',
};

/** 演示数据覆盖的日期：今天倒推 5 天（与煤气数据管理页同一条采集线） */
const DC_DEMO_DAYS = 5;

/** 确定性伪随机（sin 哈希）：同一个 seed 永远得到同一个数 */
function dcNoise(seed) {
  const x = Math.sin(seed * 97 + 13) * 10000;
  return x - Math.floor(x);
}

/** 今天 'YYYY-MM-DD'（本地时区） */
function dcToday() {
  const d = new Date();
  const mm = d.getMonth() + 1;
  const dd = d.getDate();
  return d.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
}

/** 今天起倒推 n 天的 'YYYY-MM-DD' */
function dcDaysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  const mm = d.getMonth() + 1;
  const dd = d.getDate();
  return d.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
}

/** 修正常见笔误键（防手误），找不到原样返回 */
function dcFixMaterial(name) {
  if (name === '洗洗精煤') return '洗精煤';
  return name;
}

/**
 * 演示数据：今天倒推 5 天，每天给一批采集记录。
 * 结构：类型 × 该类型的具体生产线 × 该类型的物料 × 参数（消耗量必有，主要物料带产量）。
 * 同一天同一组合只有一条，采集时间 = 当天上午 6~9 点错开。
 */
function dcDefaultRows() {
  const rows = [];
  for (let d = DC_DEMO_DAYS - 1; d >= 0; d--) {
    const iso = dcDaysAgo(d);
    const date = new Date(iso + 'T00:00:00');
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    let seq = 0;   // 采集顺序：用于把时间错开

    DC_LINE_TYPES.forEach(function (type, ti) {
      const names = DC_LINE_NAMES[type] || [];
      const mats = (DC_TYPE_MATERIALS[type] || []).map(dcFixMaterial);
      names.forEach(function (name, ni) {
        mats.forEach(function (mat, mi) {
          DC_PARAMS.forEach(function (param) {
            // 产量只给「生产型」的组合（物料有产出才谈产量）：全厂只给消耗量
            if (param === '产量' && (type === '全厂' || mi % 2 !== 0)) return;

            const seed = year * 10000 + month * 100 + d * 10 + ti * 3 + ni + mi * 7;
            const base = mat.indexOf('煤气') >= 0 ? 30000
              : (mat === '天然气' ? 8000 : 20000);
            const value = (base * (0.6 + dcNoise(seed) * 0.8)).toFixed(2);

            const hh = 6 + ((seq + ni) % 3);
            const mm2 = 10 + ((seq * 7) % 49);
            const ss = 10 + ((seq * 13) % 49);

            rows.push({
              time: iso + ' ' + (hh < 10 ? '0' : '') + hh + ':' + (mm2 < 10 ? '0' : '') + mm2 + ':' + (ss < 10 ? '0' : '') + ss,
              year: year,
              month: month,
              lineType: type,
              lineName: name,
              material: mat,
              param: param,
              value: value,
              unit: DC_MATERIAL_UNITS[mat] || 't',
            });
            seq++;
          });
        });
      });
    });
  }
  return rows;
}

/* ---------- 行数据读写 ---------- */

let dcRows = [];

/** 一行的唯一键：采集时间 × 类型 × 生产线 × 物料 × 参数 */
function dcKeyOf(r) {
  return [r.time, r.lineType, r.lineName, r.material, r.param].join('|');
}

/**
 * 行数据形状校验：字段齐 + 值合法。旧口径（time/point/medium 那套）会被拦下重新播种。
 */
function dcRowsValid(arr) {
  return Array.isArray(arr) && arr.length > 0 && arr.every(function (r) {
    return r && typeof r === 'object'
      && typeof r.time === 'string' && /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(r.time)
      && typeof r.lineType === 'string' && r.lineType !== ''
      && typeof r.lineName === 'string' && r.lineName !== ''
      && typeof r.material === 'string' && r.material !== ''
      && typeof r.param === 'string' && r.param !== ''
      && r.value !== undefined && r.value !== null && String(r.value) !== ''
      && typeof r.unit === 'string' && r.unit !== '';
  });
}

/**
 * 读全部行。没存过（或旧格式）就播种演示数据（并落库）。
 * 存过就一律以存过的为准 —— 哪怕用户清不掉，也不能自己长回来。
 */
function dcLoad() {
  try {
    const raw = localStorage.getItem(DC_ROW_KEY);
    if (raw) {
      const arr = JSON.parse(raw);
      if (dcRowsValid(arr)) return arr;
    }
  } catch (e) { /* 存储损坏就当没有，回落到演示数据 */ }

  const seed = dcDefaultRows();
  dcSave(seed);
  return seed;
}

function dcSave(rows) {
  try {
    localStorage.setItem(DC_ROW_KEY, JSON.stringify(rows == null ? dcRows : rows));
    return true;
  } catch (e) {
    toast('保存失败：浏览器存储空间不足');
    return false;
  }
}

/* ---------- 查询条件 ---------- */

/**
 * 数据时间是**范围查询**：起 / 止两端都可清空，留空 = 该端不限制。
 * 年度 / 月度下拉留空（「全部」）= 不按该条件过滤。
 */
function dcQueryStart() {
  return document.getElementById('dc-date-start').value || null;
}

function dcQueryEnd() {
  return document.getElementById('dc-date-end').value || null;
}

function dcQueryYear() {
  const v = document.getElementById('dc-year').value;
  return v === '' || v == null ? null : Number(v);
}

function dcQueryMonth() {
  const v = document.getElementById('dc-month').value;
  return v === '' || v == null ? null : Number(v);
}

/** 查询条件的中文描述（空态文案 / 导出文件名 / 查询 toast 共用） */
function dcQueryLabel() {
  const s = dcQueryStart();
  const e = dcQueryEnd();
  const y = dcQueryYear();
  const m = dcQueryMonth();
  const parts = [];
  if (s || e) {
    parts.push((s ? dcIsoToText(s) : '早期') + ' 至 ' + (e ? dcIsoToText(e) : '今'));
  }
  if (y !== null) parts.push(y + '年');
  if (m !== null) parts.push(m + '月');
  return parts.length ? parts.join('，') : '全部数据';
}

/** '2026-09-25' → '2026年9月25日' */
function dcIsoToText(iso) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(iso || ''));
  if (!m) return String(iso || '');
  return m[1] + '年' + Number(m[2]) + '月' + Number(m[3]) + '日';
}

/** '2026-09-25 08:20:00' → '2026年9月25日 08:20:00'（采集时间给人看） */
function dcTimeText(t) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2}) (.+)$/.exec(String(t || ''));
  if (!m) return String(t || '');
  return m[1] + '年' + Number(m[2]) + '月' + Number(m[3]) + '日 ' + m[4];
}

/**
 * 当前查询条件下要显示的行。
 * 数据时间按采集时间的日期部分过滤；年度 / 月度按行上的 year / month。
 * **保持存储顺序**（演示数据按时间正序播种），不额外排序。
 */
function dcFilteredRows() {
  const s = dcQueryStart();
  const e = dcQueryEnd();
  const y = dcQueryYear();
  const m = dcQueryMonth();
  return dcRows.filter(function (r) {
    const day = r.time.slice(0, 10);
    if (s && day < s) return false;
    if (e && day > e) return false;
    if (y !== null && Number(r.year) !== y) return false;
    if (m !== null && Number(r.month) !== m) return false;
    return true;
  });
}

/* ---------- 翻页 ---------- */

/** 每页条数档位与默认值 */
const DC_PAGE_SIZE_DEFAULT = 10;
const DC_PAGE_SIZES = [10, 20, 50];

/** 当前页码（1 起）与每页条数 */
let dcPage = 1;
let dcPageSize = DC_PAGE_SIZE_DEFAULT;

function dcPageCount(total) {
  return Math.max(1, Math.ceil(total / dcPageSize));
}

/**
 * 夹住页码。查询条件变化之后都要过一遍 ——
 * 否则会出现「当前在第 3 页，筛选后只剩 1 页」从而渲染出空白列表。
 */
function dcClampPage(total) {
  const max = dcPageCount(total);
  if (dcPage > max) dcPage = max;
  if (dcPage < 1) dcPage = 1;
  return dcPage;
}

/** 当前页要显示的行 */
function dcPagedRows(list) {
  const start = (dcClampPage(list.length) - 1) * dcPageSize;
  return list.slice(start, start + dcPageSize);
}

/**
 * 页码序列，超过 7 页时折叠成 1 … 4 5 6 … 20。
 * 返回数字或 '…'。
 */
function dcPageNumbers(cur, total) {
  if (total <= 7) {
    const all = [];
    for (let i = 1; i <= total; i++) all.push(i);
    return all;
  }
  const out = [1];
  const lo = Math.max(2, cur - 1);
  const hi = Math.min(total - 1, cur + 1);
  if (lo > 2) out.push('…');
  for (let i = lo; i <= hi; i++) out.push(i);
  if (hi < total - 1) out.push('…');
  out.push(total);
  return out;
}

function dcRenderPager(total) {
  const box = document.getElementById('dc-pager');
  if (!box) return;

  // 没有数据时不显示翻页条
  if (!total) { box.innerHTML = ''; return; }

  const pages = dcPageCount(total);
  const cur = dcClampPage(total);

  const nums = dcPageNumbers(cur, pages).map(function (n) {
    if (n === '…') return '<span class="dc-page-gap">…</span>';
    return '<button type="button" class="dc-page-btn' + (n === cur ? ' is-active' : '') + '"'
      + ' data-page="' + n + '">' + n + '</button>';
  }).join('');

  const sizes = DC_PAGE_SIZES.map(function (n) {
    return '<option value="' + n + '"' + (n === dcPageSize ? ' selected' : '') + '>' + n + ' 条/页</option>';
  }).join('');

  box.innerHTML =
    '<span class="dc-page-total">共 ' + total + ' 条</span>'
    + '<select class="dc-page-size" id="dc-page-size">' + sizes + '</select>'
    + '<div class="dc-page-nav">'
    + '<button type="button" class="dc-page-btn dc-page-prev"' + (cur <= 1 ? ' disabled' : '')
    + ' data-page="' + (cur - 1) + '">‹</button>'
    + nums
    + '<button type="button" class="dc-page-btn dc-page-next"' + (cur >= pages ? ' disabled' : '')
    + ' data-page="' + (cur + 1) + '">›</button>'
    + '</div>';
}

/* ---------- 渲染 ---------- */

function dcRenderTable() {
  const tbody = document.getElementById('dc-tbody');
  const list = dcFilteredRows();

  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="9" class="dc-empty">'
      + '「' + dcQueryLabel() + '」下没有采集明细'
      + '</td></tr>';
    dcRenderPager(0);
    return;
  }

  tbody.innerHTML = dcPagedRows(list).map(function (r) {
    return '<tr data-key="' + rcEsc(dcKeyOf(r)) + '">'
      + '<td class="is-num">' + rcEsc(dcTimeText(r.time)) + '</td>'
      + '<td class="is-num">' + rcEsc(r.year) + '年</td>'
      + '<td class="is-num">' + rcEsc(r.month) + '月</td>'
      + '<td>' + rcEsc(r.lineType) + '</td>'
      + '<td>' + rcEsc(r.lineName) + '</td>'
      + '<td>' + rcEsc(r.material) + '</td>'
      + '<td>' + rcEsc(r.param) + '</td>'
      + '<td class="dc-value">' + rcEsc(r.value) + '</td>'
      + '<td>' + rcEsc(r.unit) + '</td>'
      + '</tr>';
  }).join('');

  dcRenderPager(list.length);
}

function dcRenderAll() {
  dcRenderTable();
}

/* ---------- 导出 ---------- */

/**
 * 导出当前列表（筛选后）的数据为 Excel。
 * 列结构与列表一致；文件名带查询条件，便于留档。
 */
function dcExportXlsx() {
  const list = dcFilteredRows();
  if (!list.length) {
    toast('当前条件下没有数据可导出');
    return;
  }

  const head = ['采集时间', '数据年度', '数据月度', '生产线类型', '生产线名称', '物料名称', '参数名称', '数据值', '参数单位'];
  const rows = [head];
  list.forEach(function (r) {
    rows.push([r.time, r.year, r.month, r.lineType, r.lineName, r.material, r.param, r.value, r.unit]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [
    { wch: 22 }, { wch: 10 }, { wch: 10 }, { wch: 32 }, { wch: 16 },
    { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 10 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '数据采集明细');
  XLSX.writeFile(wb, '数据采集明细_' + dcQueryLabel() + '.xlsx');
}

/* ---------- 事件 ---------- */

function bindDcEvents() {
  document.getElementById('dc-search').addEventListener('click', function () {
    dcPage = 1;                     // 换了查询条件必须回到第 1 页，否则可能停在一个空页上
    dcRenderTable();
    toast('已查询「' + dcQueryLabel() + '」，共 ' + dcFilteredRows().length + ' 条');
  });

  // 清空查询条件 = 看全部数据
  document.getElementById('dc-reset').addEventListener('click', function () {
    document.getElementById('dc-date-start').value = '';
    document.getElementById('dc-date-end').value = '';
    document.getElementById('dc-year').value = '';
    document.getElementById('dc-month').value = '';
    dcPage = 1;
    dcRenderTable();
    toast('已重置为全部数据');
  });

  // 改了条件自动重查（日期 / 年度 / 月度都支持），并回第 1 页
  ['dc-date-start', 'dc-date-end', 'dc-year', 'dc-month'].forEach(function (id) {
    document.getElementById(id).addEventListener('change', function () {
      dcPage = 1;
      dcRenderTable();
    });
  });

  // 翻页：页码 / 上一页 / 下一页 / 每页条数
  document.getElementById('dc-pager').addEventListener('click', function (e) {
    const btn = e.target.closest('.dc-page-btn');
    if (!btn || btn.disabled) return;
    dcPage = Number(btn.dataset.page) || 1;
    dcRenderTable();
  });

  document.getElementById('dc-pager').addEventListener('change', function (e) {
    if (e.target.id !== 'dc-page-size') return;
    dcPageSize = Number(e.target.value) || DC_PAGE_SIZE_DEFAULT;
    dcPage = 1;                     // 换了每页条数也回第 1 页，不然位置会乱跳
    dcRenderTable();
  });

  // 导出当前列表数据（导出筛选后的**全部**数据，不受翻页影响）
  document.getElementById('dc-export').addEventListener('click', function () {
    dcExportXlsx();
    toast('已导出 ' + dcFilteredRows().length + ' 条数据');
  });
}

/* ---------- 初始化 ---------- */

function initCollectionDetailPage() {
  initLayout('collection-detail', { moduleId: 'carbon' });

  dcRows = dcLoad();

  // 年度 / 月度下拉：候选取数据里实际出现过的（去重），默认「全部」
  const years = [];
  const months = [];
  dcRows.forEach(function (r) {
    if (years.indexOf(Number(r.year)) < 0) years.push(Number(r.year));
    if (months.indexOf(Number(r.month)) < 0) months.push(Number(r.month));
  });
  years.sort(function (a, b) { return b - a; });
  months.sort(function (a, b) { return a - b; });

  document.getElementById('dc-year').innerHTML = '<option value="">全部</option>'
    + years.map(function (y) { return '<option value="' + y + '">' + y + '年</option>'; }).join('');
  document.getElementById('dc-month').innerHTML = '<option value="">全部</option>'
    + months.map(function (m) { return '<option value="' + m + '">' + m + '月</option>'; }).join('');

  dcPage = 1;
  dcRenderAll();
  bindDcEvents();
}

if (document.querySelector('.app-shell')) initCollectionDetailPage();
