/**
 * 煤气数据管理 —— 数据采集管理 > 业务界面 > 煤气数据管理
 *
 * 「一份日数据文件 = 一行」的档案视角：
 *   查询条件：数据日期（范围查询，两端可清空）
 *   列表按钮：只有「导入」→ 弹窗（数据日期 + 上传文件 + 确认/取消）
 *   列表列：数据日期(年月日) / 操作时间 / 原始文件(点击查看) / 操作(详情 / 删除)
 *   详情弹窗：展示该文件解析出的数据表格
 *
 * 数据落在 localStorage 键 GAS_ROW_KEY。
 * 行结构：{ date: 'YYYY-MM-DD', opTime: 'YYYY-MM-DD HH:mm:ss',
 *          fileName, fileType, fileData(base64 DataURL), data: [...解析结果] }
 *
 * 演示数据：从今天倒推 5 天，每天一条；data 里是解析出的量能明细。
 * 原始文件查看：Excel/CSV 用 XLSX 解析后按工作表渲染成表格；PDF 用 iframe 原生预览。
 */

/* ---------- 常量 ---------- */

/**
 * 行数据存储键（本项目内独立的一套，不与其他业务线共享）。
 * v3：详情表列口径改成「生产线类型/物料/参数名称/数值/单位」时升的版 ——
 * 之前版本存的 data 行字段对不上（旧的是 type/self/out/release，新的是 line/material/param/value），
 * 行级校验拦不住行内字段变化，详情弹窗就会渲染成空表。升键 + 行内形状校验双保险。
 * v4：详情表的「单位」列口径从 10⁴Nm³ 改成 Nm³，并新增一条全厂行 ——
 * unit 不在 gasDetailRowValid 的校验字段里，旧数据能过校验但单位是旧的，
 * 所以同样要升键，让演示数据重新播种。
 */
const GAS_ROW_KEY = 'emission-mgmt-gas-rows-v4';

/** 原型口径：文件以 base64 存行里，单文件上限 4MB（base64 后约 5.4MB） */
const GAS_FILE_MAX = 4 * 1024 * 1024;

/** '2026-09-25' → '2026年9月25日'（给人看） */
function gasDateText(iso) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(iso || ''));
  if (!m) return String(iso || '');
  return Number(m[1]) + '年' + Number(m[2]) + '月' + Number(m[3]) + '日';
}

/** 今天 'YYYY-MM-DD'（本地时区） */
function gasToday() {
  const d = new Date();
  const mm = d.getMonth() + 1;
  const dd = d.getDate();
  return d.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
}

/** 今天起倒推 n 天的 'YYYY-MM-DD' */
function gasDaysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  const mm = d.getMonth() + 1;
  const dd = d.getDate();
  return d.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
}

/** 演示数据共用的文件元信息 */
const GAS_DEMO_FILE = {
  fileName: '煤气日数据采集表.xlsx',
  fileType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

/**
 * 导入文件解析出的二维数组 → 详情行对象数组。
 *
 * 为什么要转：`XLSX.sheet_to_json(..., {header:1})` / `gasParseCsv()` 给的都是二维数组，
 * 而详情弹窗按 `{line, material, param, value, unit}` 取值。直接存二维数组的话，
 * 所有行的 line 都是 undefined —— 详情表全空，而且「相邻同值合并」会把它们
 * 全并成一个跨 N 行的空格子，比不合并还难看。
 *
 * 认列顺序（第 1 行像表头就跳过）：生产线类型 / 物料 / 参数名称 / 数值 / 单位。
 * 表格里一列都认不出来的行（空行、说明行）直接丢掉。
 */
function gasNormalizeRows(raw) {
  if (!Array.isArray(raw)) return [];

  const out = [];
  raw.forEach(function (row, i) {
    // 已经是对象了（演示数据 / 以后改结构）就原样留
    if (row && !Array.isArray(row) && typeof row === 'object') {
      if (gasDetailRowValid(row)) out.push(row);
      return;
    }
    if (!Array.isArray(row)) return;

    const cell = function (k) { return row[k] === undefined || row[k] === null ? '' : String(row[k]).trim(); };
    // 表头行：第一格写着「生产线类型」之类 → 跳过（只可能出现在首行）
    if (i === 0 && /生产线|类型|物料/.test(cell(0) + cell(1))) return;

    const line = cell(0);
    const material = cell(1);
    const param = cell(2);
    const value = cell(3);
    if (!line || !material || !param || value === '') return;   // 凑不齐一行有效数据就丢

    out.push({
      line: line,
      material: material,
      param: param,
      value: value,
      unit: cell(4) || 'Nm³',
    });
  });
  return out;
}

/**
 * 数值展示口径：能当数字用就**固定两位小数**；转不成数字（比如导入的 Excel 给的是文本）就原样输出。
 * 详情弹窗、导入预览都走这里，保证「同一列小数位一致」，不会出现 199540 和 199540.00 混排。
 */
function gasValueText(v) {
  if (v === undefined || v === null || String(v).trim() === '') return '';
  const n = Number(String(v).replace(/,/g, ''));
  return isFinite(n) ? n.toFixed(2) : String(v);
}

/**
 * 演示数据的基准表（今天这一天与需求样例完全一致）。
 * 列：生产线类型 / 物料 / 参数名称 / 数值 / 单位（单位统一 Nm³）。
 * 第一行是「全厂」汇总行（来自用户样例）。
 * 往前推的几天数值围绕基准做确定性小扰动，结构不变 —— 每天看得见差异又不是乱数。
 * 注意：基准表**按生产线类型聚簇排列**（同类相邻），详情弹窗才能把同类合并成一个单元格。
 */
const GAS_DEMO_BASE = [
  { line: '全厂', material: '转炉煤气', param: '消耗量', value: -128178 },
  { line: '焦化工序', material: '焦炉煤气', param: '消耗量', value: 199540 },
  { line: '焦化工序', material: '焦炉煤气', param: '产量', value: 236483 },
  { line: '烧结工序', material: '高炉煤气', param: '消耗量', value: 3498 },
  { line: '烧结工序', material: '焦炉煤气', param: '消耗量', value: 922604 },
  { line: '炼铁工序', material: '焦炉煤气', param: '消耗量', value: 481961 },
  { line: '炼铁工序', material: '高炉煤气', param: '消耗量', value: 500494 },
  { line: '炼铁工序', material: '高炉煤气', param: '产量', value: 37418 },
  { line: '转炉炼钢工序', material: '焦炉煤气', param: '消耗量', value: 35768 },
  { line: '转炉炼钢工序', material: '转炉煤气', param: '产量', value: 1383 },
  { line: '掺烧自产二次能源的化石燃料发电设施', material: '转炉煤气', param: '消耗量', value: 19196 },
  { line: '掺烧自产二次能源的化石燃料发电设施', material: '焦炉煤气', param: '消耗量', value: 1784046 },
  { line: '掺烧自产二次能源的化石燃料发电设施', material: '高炉煤气', param: '消耗量', value: 2138961 },
];

/**
 * 某天的演示数据（详情弹窗里展示的表格，即解析出的文件内容）。
 * dayIdx = 0（今天）→ 原样输出基准表；往前推的每天 ±5% 内确定性扰动。
 */
function gasDemoDetail(dayIdx, iso) {
  // 确定性伪随机：同一天永远同一批数
  const noise = function (seed) {
    const x = Math.sin(seed * 97 + 13) * 10000;
    return x - Math.floor(x);
  };
  return GAS_DEMO_BASE.map(function (b, k) {
    const factor = dayIdx === 0 ? 1 : (1 + (noise(dayIdx * 12 + k) - 0.5) * 0.1);
    return {
      line: b.line,
      material: b.material,
      param: b.param,
      value: gasValueText(Math.round(b.value * factor)),
      unit: 'Nm³',
    };
  });
}

/**
 * 演示数据：今天倒推 5 天，每天一条（5 条）。
 * 操作时间 = 当天上午 8 点到 10 点之间（按天错开，看得出行间顺序）。
 */
function gasDefaultRows() {
  const rows = [];
  for (let i = 0; i < 5; i++) {
    const iso = gasDaysAgo(i);
    rows.push({
      date: iso,
      opTime: gasDaysAgo(i) + ' 0' + (8 + (i % 3)) + ':2' + (i % 6) + ':00',
      fileName: GAS_DEMO_FILE.fileName,
      fileType: GAS_DEMO_FILE.fileType,
      fileData: '',              // 演示行不带文件体，「查看文件」给演示内容
      demoFile: true,            // 标记：演示文件（查看文件 / 详情都用演示数据）
      data: gasDemoDetail(i, iso),
    });
  }
  return rows;
}

/* ---------- 行数据读写 ---------- */

let gasRows = [];
let gasImportFile = null;        // 导入弹窗里待上传的文件（{name, type, dataUrl, rows?}）

/** 一行的唯一键：数据日期（同一天只有一份文件） */
function gasKeyOf(r) {
  return r.date;
}

/**
 * 详情数据行的形状校验：必须有 line / material / param / value 字段。
 * 行级校验（date/opTime/fileName/data）拦不住 data 内部字段口径变化
 * （上一版是 type/self/out/release），所以 data 也要逐行查 —— 否则详情弹窗渲染成空表。
 */
function gasDetailRowValid(row) {
  return row && typeof row === 'object'
    && typeof row.line === 'string' && row.line !== ''
    && typeof row.material === 'string' && row.material !== ''
    && typeof row.param === 'string' && row.param !== ''
    && row.value !== undefined && row.value !== null && String(row.value) !== '';
}

/**
 * 行数据形状校验：新结构一行必须有 date / opTime / fileName / data（数组），
 * 且 data 每行都要过 gasDetailRowValid。
 * data 允许为空数组（导入 PDF 这类解析不出表的文件是合法情况，不能因此把整批数据判废）。
 * 读到旧格式（不管是台账行还是上一版详情结构）就整个当没有，回落到演示数据。
 */
function gasRowsValid(arr) {
  return Array.isArray(arr) && arr.length > 0 && arr.every(function (r) {
    return r && typeof r === 'object'
      && typeof r.date === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(r.date)
      && typeof r.opTime === 'string'
      && typeof r.fileName === 'string'
      && Array.isArray(r.data)
      && r.data.every(gasDetailRowValid);
  });
}

/**
 * 读全部行。没存过就播种演示数据（并落库）。
 * 存过就一律以存过的为准 —— 哪怕用户把行删光了，也不能自己长回来。
 */
function gasLoad() {
  try {
    const raw = localStorage.getItem(GAS_ROW_KEY);
    if (raw) {
      const arr = JSON.parse(raw);
      if (gasRowsValid(arr)) return arr;
    }
  } catch (e) { /* 存储损坏就当没有，回落到演示数据 */ }

  const seed = gasDefaultRows();
  gasSave(seed);
  return seed;
}

function gasSave(rows) {
  try {
    localStorage.setItem(GAS_ROW_KEY, JSON.stringify(rows == null ? gasRows : rows));
    return true;
  } catch (e) {
    // 最常见的是超出 localStorage 配额（文件太大），给一句能懂的话
    toast('保存失败：浏览器存储空间不足，文件可能过大');
    return false;
  }
}

/* ---------- 查询条件 ---------- */

/**
 * 数据日期是**范围查询**：起 / 止两端都可清空，留空 = 该端不限制。
 * 都为空 = 全部数据。
 */
function gasQueryStart() {
  return document.getElementById('gas-date-start').value || null;
}

function gasQueryEnd() {
  return document.getElementById('gas-date-end').value || null;
}

/** 查询条件的中文描述（空态文案 / 查询 toast 共用） */
function gasQueryLabel() {
  const s = gasQueryStart();
  const e = gasQueryEnd();
  if (s && e) return gasDateText(s) + ' 至 ' + gasDateText(e);
  if (s) return gasDateText(s) + ' 起';
  if (e) return gasDateText(e) + ' 止';
  return '全部数据';
}

/** 当前查询条件下要显示的行：按数据日期倒序（新的在前） */
function gasFilteredRows() {
  const s = gasQueryStart();
  const e = gasQueryEnd();
  return gasRows.filter(function (r) {
    if (s && r.date < s) return false;
    if (e && r.date > e) return false;
    return true;
  }).sort(function (a, b) {
    return a.date < b.date ? 1 : -1;
  });
}

/* ---------- 渲染 ---------- */

function gasRenderTable() {
  const tbody = document.getElementById('gas-tbody');
  const list = gasFilteredRows();

  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="gas-empty">'
      + '「' + gasQueryLabel() + '」下没有煤气数据，可点右上角「导入」上传'
      + '</td></tr>';
    return;
  }

  tbody.innerHTML = list.map(function (r) {
    return '<tr data-key="' + rcEsc(gasKeyOf(r)) + '">'
      + '<td class="is-num">' + rcEsc(gasDateText(r.date)) + '</td>'
      + '<td class="is-num">' + rcEsc(r.opTime) + '</td>'
      + '<td class="gas-file-cell"><button type="button" class="rc-link gas-file" title="查看文件">'
      + rcEsc(r.fileName)
      + '</button></td>'
      + '<td class="gas-act">'
      + '<button type="button" class="rc-link gas-detail">详情</button>'
      + '<button type="button" class="rc-link rc-link-danger gas-del">删除</button>'
      + '</td>'
      + '</tr>';
  }).join('');
}

function gasRenderAll() {
  gasRenderTable();
}

/* ---------- 导入弹窗 ---------- */

function gasOpenImport() {
  gasImportFile = null;
  // 数据日期默认今天
  document.getElementById('gas-import-date').value = gasToday();
  const box = document.getElementById('gas-upload-box');
  box.innerHTML = '点击选择要上传的数据文件';
  box.classList.remove('has-file');
  openModal('gas-import-modal');
}

function gasApplyImport() {
  const date = document.getElementById('gas-import-date').value;
  if (!date) { toast('请选择数据日期'); return; }
  if (!gasImportFile) { toast('请先选择要上传的数据文件'); return; }

  // 同一天重复导入 = 用新文件覆盖旧文件（口径：一天一份日数据文件）
  const exist = gasRows.filter(function (x) { return x.date === date; })[0];

  const row = {
    date: date,
    opTime: nowText(),
    fileName: gasImportFile.name,
    fileType: gasImportFile.type,
    fileData: gasImportFile.dataUrl,
    demoFile: false,
    data: gasNormalizeRows(gasImportFile.rows),   // Excel/CSV 解析出的行；PDF 解析不出就是空数组
  };

  if (exist) {
    const idx = gasRows.indexOf(exist);
    gasRows[idx] = row;
  } else {
    gasRows.push(row);
  }

  if (!gasSave()) return;
  closeModal('gas-import-modal');
  gasRenderTable();
  toast(exist ? '已覆盖导入 ' + gasDateText(date) + ' 的数据' : '导入成功');
}

/* ---------- 原始文件点击 ---------- */

/**
 * 原始文件列：点击只给一句文字提示（toast），不弹窗。
 * 文件内容走「详情」看。
 */
function gasOpenFile(key) {
  const r = gasRows.filter(function (x) { return gasKeyOf(x) === key; })[0];
  if (!r) { toast('这一条已经不存在了'); return; }
  toast('可以预览文件内容');
}

/* ---------- 详情弹窗 ---------- */

/**
 * 详情弹窗：展示该天文件解析出的数据表。
 * 列（与需求样例一致）：生产线类型 / 物料 / 参数名称 / 数值 / 单位。
 *
 * 两个展示规则：
 *   1. 「生产线类型」列**上下相邻同值合并**（rowspan）—— 同类工序只出现一次，
 *      下面被覆盖的行不再输出该格（不能输出空 td，否则列会错位）；
 *   2. 数值列统一**保留两位小数**（走 gasValueText）。
 *
 * ⚠️ 合并只认**相邻**同值，不认全表同值：所以基准表必须按类型聚簇排列。
 */
function gasOpenDetail(key) {
  const r = gasRows.filter(function (x) { return gasKeyOf(x) === key; })[0];
  if (!r) { toast('这一条已经不存在了'); return; }

  // 演示行没有真实文件体：拿演示基准表演示；真实行就用自己的解析结果（解析不出就是空的，
  // 不能拿演示数据顶上 —— 那是假数据，会让用户以为文件解析成功了）
  let rows = (r.demoFile && (!Array.isArray(r.data) || !r.data.length))
    ? gasDemoDetail(0, r.date)
    : (Array.isArray(r.data) ? r.data : []);
  rows = rows.filter(gasDetailRowValid);   // 兜一层：脏行不进表，避免合并逻辑被空值带偏

  const box = document.getElementById('gas-detail-body');

  if (!rows.length) {
    box.innerHTML = '<div class="gas-detail-empty">该文件没有解析出数据行'
      + '<br /><span class="gas-detail-empty-hint">'
      + '仅 Excel / CSV 能解析出表格，PDF 只做预览；也可能是表头与'
      + '「生产线类型 / 物料 / 参数名称 / 数值 / 单位」对不上。</span></div>';
    openModal('gas-detail-modal');
    return;
  }

  const head = ['生产线类型', '物料', '参数名称', '数值', '单位'];
  const html = ['<table class="data-table gas-detail-table"><thead><tr>'];
  head.forEach(function (h) { html.push('<th>' + h + '</th>'); });
  html.push('</tr></thead><tbody>');

  rows.forEach(function (row, idx) {
    const prev = rows[idx - 1];
    // 与上一行同类型 → 本行不输出「生产线类型」格，由上一行的 rowspan 覆盖
    const covered = idx > 0 && prev && prev.line === row.line;

    html.push('<tr>');
    if (!covered) {
      let span = 1;
      while (idx + span < rows.length && rows[idx + span].line === row.line) span++;
      html.push('<td class="gas-line"'
        + (span > 1 ? ' rowspan="' + span + '"' : '')
        + '>' + rcEsc(row.line) + '</td>');
    }
    html.push('<td>' + rcEsc(row.material) + '</td>'
      + '<td>' + rcEsc(row.param) + '</td>'
      + '<td class="gas-value">' + rcEsc(gasValueText(row.value)) + '</td>'
      + '<td>' + rcEsc(row.unit || 'Nm³') + '</td>'
      + '</tr>');
  });
  html.push('</tbody></table>');
  box.innerHTML = html.join('');
  openModal('gas-detail-modal');
}

/* ---------- 删除 ---------- */

/** 待删除行的键（确认弹窗点「删除」时用它） */
let gasDeletingKey = '';

function gasDelete(key) {
  const r = gasRows.filter(function (x) { return gasKeyOf(x) === key; })[0];
  if (!r) return;

  // 删除不可逆，先确认一次。自绘弹窗而不是 window.confirm：原生对话框和页面风格不一致。
  gasDeletingKey = key;
  document.getElementById('gas-confirm-text').textContent =
    '确定删除 ' + gasDateText(r.date) + ' 的煤气数据吗？删除后不可恢复。';
  document.getElementById('gas-confirm-close').focus();
  openModal('gas-confirm-modal');
}

function gasApplyDelete() {
  const key = gasDeletingKey;
  gasDeletingKey = '';
  closeModal('gas-confirm-modal');

  if (!key) return;
  const before = gasRows.length;
  gasRows = gasRows.filter(function (x) { return gasKeyOf(x) !== key; });
  if (gasRows.length === before) return;   // 没删掉（比如重复点了）

  gasSave();
  gasRenderTable();
  toast('已删除');
}

/* ---------- 事件 ---------- */

function bindGasEvents() {
  document.getElementById('gas-search').addEventListener('click', function () {
    gasRenderTable();
    toast('已查询「' + gasQueryLabel() + '」，共 ' + gasFilteredRows().length + ' 条');
  });

  // 清空查询条件 = 看全部数据
  document.getElementById('gas-reset').addEventListener('click', function () {
    document.getElementById('gas-date-start').value = '';
    document.getElementById('gas-date-end').value = '';
    gasRenderTable();
    toast('已重置为全部数据');
  });

  // 改了日期自动重查（两个日期框都支持）
  ['gas-date-start', 'gas-date-end'].forEach(function (id) {
    document.getElementById(id).addEventListener('change', gasRenderTable);
  });

  // 导入弹窗
  document.getElementById('gas-import').addEventListener('click', gasOpenImport);
  document.getElementById('gas-import-ok').addEventListener('click', gasApplyImport);
  document.getElementById('gas-import-cancel').addEventListener('click', function () { closeModal('gas-import-modal'); });
  document.getElementById('gas-import-close').addEventListener('click', function () { closeModal('gas-import-modal'); });

  // 上传框：点一下选文件
  const file = document.getElementById('gas-file');
  document.getElementById('gas-upload-box').addEventListener('click', function () {
    file.value = '';   // 同一个文件连选两次也要能触发 change
    file.click();
  });

  file.addEventListener('change', function () {
    const f = file.files && file.files[0];
    if (!f) return;

    if (f.size > GAS_FILE_MAX) {
      toast('文件超过 4MB 上限，请压缩后再试');
      file.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = function () {
      gasImportFile = {
        name: f.name,
        type: f.type || '',
        dataUrl: String(reader.result),
        rows: null,
      };
      // Excel / CSV 立即解析出数据行（详情弹窗直接用）；PDF 不解析
      if (!/\.pdf$/i.test(f.name)) {
        try {
          let rows;
          if (/\.csv$/i.test(f.name)) {
            rows = gasParseCsv(String(reader.result.replace(/^data:[^,]+,/, '')));
          } else {
            const b64 = gasImportFile.dataUrl.replace(/^data:[^;]+;base64,/, '');
            const bin = atob(b64);
            const bytes = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            const wb = XLSX.read(bytes, { type: 'array' });
            rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, blankrows: false, defval: '' });
          }
          gasImportFile.rows = rows;
        } catch (e) { /* 解析失败不拦导入，详情里会提示 */ }
      }

      const box = document.getElementById('gas-upload-box');
      box.innerHTML = '📄 ' + rcEsc(f.name)
        + '<br /><span class="gas-upload-hint">' + (f.size / 1024).toFixed(1) + ' KB，点击可重新选择</span>';
      box.classList.add('has-file');
    };
    reader.readAsDataURL(f);
  });

  // CSV 文本读出来是纯文本，readAsDataURL 会变成 base64 —— CSV 单独按文本读再转
  // （简化处理：CSV 也走 DataURL，导入解析时再解码；上面 rows 分支已处理）

  // 列表：查看文件 / 详情 / 删除
  document.getElementById('gas-tbody').addEventListener('click', function (e) {
    const tr = e.target.closest('tr[data-key]');
    if (!tr) return;
    if (e.target.closest('.gas-file')) gasOpenFile(tr.dataset.key);
    else if (e.target.closest('.gas-detail')) gasOpenDetail(tr.dataset.key);
    else if (e.target.closest('.gas-del')) gasDelete(tr.dataset.key);
  });

  // 详情弹窗关闭
  document.getElementById('gas-detail-ok').addEventListener('click', function () { closeModal('gas-detail-modal'); });
  document.getElementById('gas-detail-close').addEventListener('click', function () { closeModal('gas-detail-modal'); });

  // 删除确认弹窗
  document.getElementById('gas-confirm-ok').addEventListener('click', gasApplyDelete);
  document.getElementById('gas-confirm-cancel').addEventListener('click', function () {
    gasDeletingKey = '';
    closeModal('gas-confirm-modal');
  });
  document.getElementById('gas-confirm-close').addEventListener('click', function () {
    gasDeletingKey = '';
    closeModal('gas-confirm-modal');
  });

  // 点遮罩关弹窗
  ['gas-import-modal', 'gas-detail-modal', 'gas-confirm-modal'].forEach(function (id) {
    document.getElementById(id).addEventListener('click', function (e) {
      if (e.target.id === id) {
        if (id === 'gas-confirm-modal') gasDeletingKey = '';
        closeModal(id);
      }
    });
  });
}

/** 拆 CSV 文本为二维数组（处理引号包裹、引号转义、\r\n） */
function gasParseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuote = false;

  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuote) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQuote = false;
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuote = true;
    } else if (c === ',') {
      row.push(field); field = '';
    } else if (c === '\n') {
      row.push(field); rows.push(row); row = []; field = '';
    } else if (c !== '\r') {
      field += c;
    }
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/* ---------- 初始化 ---------- */

function initGasDataPage() {
  initLayout('gas-data', { moduleId: 'carbon' });

  gasRows = gasLoad();

  gasRenderAll();
  bindGasEvents();
}

if (document.querySelector('.app-shell')) initGasDataPage();
