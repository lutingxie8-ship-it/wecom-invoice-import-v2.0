// wecom_invoice_import.js — 企微文档批量录入开票记录（V2.0 独立脚本版）
//
// V2.0 变化：
//   1. 粘贴列数从 10 列改为 18 列（col0..col17），新增 col17「是否邮件开票」。
//      col17 由 read_excel_to_tsv.py 按开票人规则产出：
//        符瑞香 → 空；伍惠娟 → "是"；其他开票人已在 Python 侧拦截为异常。
//   2. 输入/输出文件改名，避免与 V1 冲突。
//   3. 导航 API 从 sheet.setActiveCell 迁移到 app.view.setCurrentSelection（引擎已移除 setActiveCell）。
//   4. 查重/刷新验证前等待「数据加载稳定」（企微表格数据是渐进加载的，立即扫描会漏行）。
//
// 输入：~/.dev-browser/tmp/wecom_import_v2_input.json
//   {
//     "tsv": "<18列TSV，由 read_excel_to_tsv.py 产出>",
//     "doc_url": "https://doc.weixin.qq.com/sheet/...",   // 可选，缺省用默认
//     "force": false,                                      // 可选，true=即使有重复也强制录入全部
//     "skip_duplicates": false                             // 可选，true=跳过重复行，只录入新增
//   }
// 输出：~/.dev-browser/tmp/wecom_import_v2_output.json
//
// 18列TSV列布局（引擎列，0-based）：
//   col0 空 | col1 公司 | col2 开票日期 | col3 发票代码 | col4 发票号码
//   col5 发票类型 | col6 开票名称 | col7 纳税人识别号 | col8 开票金额 | col9 订单ID
//   col10 备注 | col11 归属项目 | col12 备注_1 | col13 订单ID是否重复
//   col14 年 | col15 月 | col16 货物或应税劳务名称 | col17 是否邮件开票
//
// ⚠️ 事故教训（务必保留）：
//   1. 幽灵粘贴：上一次导入尝试遗留的未提交 mutation 会被本次键盘操作触发提交。
//      → 本脚本开篇先 goto 全新加载清残留态；粘贴前强制读回目标行确认空，非空即停手。
//   2. 列偏移：锚点偏一列会导致整块错位。
//      → 本脚本直接粘贴完整 18 列 TSV，从 A 列(col0)起粘，零偏移风险。
//   3. 重复录入：只查"最近2日期"跨天补录易漏。
//      → 本脚本一次 evaluate 读全表发票号码做全量查重。
//   4. 数据渐进加载：企微表格 canvas 数据是异步分批渲染的，goto 后立即 scan 只能读到部分行，
//      导致 lastRow 算小、漏查重，粘贴时会覆盖还没加载出来的已有数据。
//      → 查重/刷新验证前必须用 waitForDataStable 轮询到 lastRow 稳定。

var PAGE = "wecom-doc";
// 真实文档链接不入库：运行时从输入 JSON 的 doc_url 或环境变量 WECOM_DOC_URL 注入。
var DEFAULT_DOC_URL = "https://doc.weixin.qq.com/sheet/REPLACE_WITH_YOUR_DOC_ID";
var INPUT_PATH = "wecom_import_v2_input.json";
var OUTPUT_PATH = "wecom_import_v2_output.json";

// 列常量（引擎 0-based）
var NUM_COLS = 18;   // 粘贴列数
var DATE_COL = 2;    // 开票日期
var NUM_COL = 4;     // 发票号码
var ORDER_COL = 9;   // 订单ID
var MAIL_COL = 17;   // 是否邮件开票（V2.0 新增）

// 统一进度日志：[步骤 n/N] 描述
function step(n, total, msg) {
  console.log("[步骤 " + n + "/" + total + "] " + msg);
}

function out(obj) {
  console.log(JSON.stringify(obj, null, 2));
  try { writeFile(OUTPUT_PATH, JSON.stringify(obj, null, 2)); } catch (e) {}
}

// ---- 读取输入 ----
var input = null;
try {
  var raw = await readFile(INPUT_PATH);
  input = JSON.parse(raw);
} catch (e) {
  out({ status: "input_read_failed", detail: String(e) });
}

if (!input || !input.tsv || !String(input.tsv).trim()) {
  out({ status: "no_input", hint: "请在 wecom_import_v2_input.json 提供 tsv 字段（18列TSV）" });
} else {
  await main();
}

// 全表扫描：读所有现有记录的 发票号码(col4)，并以 col4 为准找最后一条记录所在行。
// 注意：lastRow 以「发票号码」为锚，而不是「任意列非空」——表格底部可能存在杂散数据
// （如某处残留的日期 2099/12/31），用任意列会把它误判成最后一行，导致粘贴位置错误。
async function fullScan(page) {
  return await page.evaluate(function () {
    var app = window.SpreadsheetApp;
    var sid = app.workbook.worksheetManager.activeSheetId;
    var sheet = app.workbook.worksheetManager.getSheetBySheetId(sid);
    var total = sheet.getRowCount();
    var existingNums = [];
    var lastRow = 0;
    for (var r = 1; r < total; r++) {
      var nCell = sheet.getCellDataAtPosition(r, 4);
      var num = nCell && nCell.formattedValue ? nCell.formattedValue.value : "";
      if (num) {
        existingNums.push(num);
        lastRow = r;
      }
    }
    return { existingNums: existingNums, lastRow: lastRow, total: total };
  });
}

// 等待数据加载稳定：连续两次扫描 lastRow 一致即返回
async function waitForDataStable(page, maxAttempts) {
  var prev = -1;
  var scan = null;
  for (var i = 0; i < maxAttempts; i++) {
    scan = await fullScan(page);
    if (scan.lastRow === prev) { return scan; }
    prev = scan.lastRow;
    console.log("  数据加载中 lastRow=" + scan.lastRow + "，继续等待稳定...");
    await page.waitForTimeout(700);
  }
  return scan;
}

// 探活：轮询等待 app.view 就绪（它是异步初始化的），再检查关键接口 + 判定导航方式（兼容新旧引擎）。轻量、毫秒级。
async function probeEngine(page, timeoutMs) {
  var start = Date.now();
  var last = null;
  while (Date.now() - start < timeoutMs) {
    last = await page.evaluate(function () {
      var app = window.SpreadsheetApp;
      var sid = app.workbook.worksheetManager.activeSheetId;
      var sheet = app.workbook.worksheetManager.getSheetBySheetId(sid);
      var view = app.view;
      function has(o, k) { return o && typeof o[k] === "function"; }

      var missing = [];
      if (!has(sheet, "getCellDataAtPosition")) missing.push("sheet.getCellDataAtPosition");
      if (!has(sheet, "getRowCount")) missing.push("sheet.getRowCount");
      if (!has(sheet, "setCellDataAtPosition")) missing.push("sheet.setCellDataAtPosition");
      if (!view) missing.push("SpreadsheetApp.view");

      // 导航接口：优先用新的 view.setCurrentSelection，退回老的 sheet.setActiveCell
      var hasNewNav = has(view, "setCurrentSelection");
      var hasOldNav = has(sheet, "setActiveCell");
      var nav = hasNewNav ? "setCurrentSelection" : (hasOldNav ? "setActiveCell" : null);
      if (!nav) missing.push("导航接口(view.setCurrentSelection / sheet.setActiveCell)");

      return { missing: missing, nav: nav };
    });
    // view 就绪且导航方式确定 → 通过
    if (last.nav && last.missing.length === 0) {
      return { ok: true, missing: [], nav: last.nav };
    }
    await page.waitForTimeout(500);
  }
  // 超时仍不满足 → 报异常
  return { ok: false, missing: last ? last.missing : [], nav: last ? last.nav : null };
}

// dump 详情：仅在探活发现接口缺失/异常时调用，把 sheet/view 的方法与相关函数源码倒出来，便于定位新写法。
async function dumpEngine(page) {
  return await page.evaluate(function () {
    var app = window.SpreadsheetApp;
    var sid = app.workbook.worksheetManager.activeSheetId;
    var sheet = app.workbook.worksheetManager.getSheetBySheetId(sid);
    var view = app.view;
    function methodNames(o) {
      var m = [];
      if (!o) return m;
      for (var k in o) { try { if (typeof o[k] === "function") m.push(k); } catch (e) {} }
      return m;
    }
    function src(o, k) {
      return o && typeof o[k] === "function" ? String(o[k]).slice(0, 500) : null;
    }
    return {
      sheetType: sheet && sheet.getSheetType ? sheet.getSheetType() : null,
      sheetMethods: methodNames(sheet),
      viewMethods: methodNames(view),
      setCurrentSelectionSrc: src(view, "setCurrentSelection"),
      setSelectionSrc: src(view, "setSelection"),
      setActiveCellSrc: src(sheet, "setActiveCell")
    };
  });
}

async function main() {
  var tsvRaw = String(input.tsv);
  var linesRaw = tsvRaw.split("\n").filter(function (l) { return l.trim() !== ""; });

  // 归一化：每行拆成 18 列，不足补空（防末尾空列被 trim 掉导致的列数不足）
  var lines = linesRaw.map(function (l) {
    var cols = l.split("\t");
    while (cols.length < NUM_COLS) cols.push("");
    return cols;
  });
  var pasteCount = lines.length;
  var tsv18 = lines.map(function (cols) { return cols.join("\t"); }).join("\n");

  var docUrl = input.doc_url || DEFAULT_DOC_URL;
  var force = input.force === true;
  var TOTAL = 8;

  step(1, TOTAL, "全新加载企微文档（清残留态防幽灵粘贴）" + docUrl);
  var page = await browser.getPage(PAGE);
  await page.goto(docUrl, { waitUntil: "domcontentloaded" });

  step(2, TOTAL, "等待 SpreadsheetApp 引擎就绪...");
  var appReady = await waitForAppReady(page, 20000);
  if (!appReady.ok) {
    out({ status: "app_not_ready", detail: "引擎未就绪，可能未登录", elapsed_ms: appReady.elapsed });
    return;
  }
  var sheetReady = await waitForSheetReady(page, 15000);
  if (!sheetReady.ok) {
    out({ status: "app_not_ready", detail: "sheet 数据未就绪", elapsed_ms: sheetReady.elapsed });
    return;
  }
  console.log("  ✓ 引擎+Sheet 就绪");

  // ---- 接口探活：动手前先确认关键接口都在、并确定导航方式（轮询等 view 就绪）----
  var probe = await probeEngine(page, 15000);
  if (!probe.ok) {
    var dump = await dumpEngine(page);
    out({
      status: "engine_api_changed",
      missing: probe.missing,
      dump: dump,
      hint: "企微表格引擎接口已变化，缺少上述接口。请按 dump 里的方法列表/函数源码定位新写法，更新脚本后重跑。"
    });
    return;
  }
  console.log("  ✓ 接口探活通过，导航方式: " + probe.nav);

  // ---- 全表查重（等待数据加载稳定后扫描）----
  step(3, TOTAL, "全表查重（等待数据加载稳定，读现有发票号码比对）...");
  var scan = await waitForDataStable(page, 15);

  // 提取 TSV 的发票号码（第5列，index 4）
  var tsvNums = lines.map(function (cols) { return cols[NUM_COL] || ""; }).filter(function (n) { return n !== ""; });
  var existingSet = {};
  scan.existingNums.forEach(function (n) { existingSet[n] = true; });
  var duplicates = tsvNums.filter(function (n) { return existingSet[n]; });
  var dupSet = {};
  duplicates.forEach(function (n) { dupSet[n] = true; });
  var newCount = lines.filter(function (cols) { return !dupSet[cols[NUM_COL] || ""]; }).length;
  var mailCount = lines.filter(function (cols) { return (cols[MAIL_COL] || "") === "是"; }).length;

  console.log("  现有记录 " + scan.existingNums.length + " 条，TSV " + tsvNums.length + " 条，重复 " + duplicates.length + " 条，新增 " + newCount + " 条");
  console.log("  其中「是否邮件开票=是」: " + mailCount + " 条");

  if (duplicates.length > 0 && !force) {
    if (input.skip_duplicates === true) {
      // 跳过重复：过滤掉已在表格里的行，只保留新增（用于"删了记录要补回"等场景）
      lines = lines.filter(function (cols) { return !dupSet[cols[NUM_COL] || ""]; });
      pasteCount = lines.length;
      tsv18 = lines.map(function (cols) { return cols.join("\t"); }).join("\n");
      tsvNums = lines.map(function (cols) { return cols[NUM_COL] || ""; }).filter(function (n) { return n !== ""; });
      mailCount = lines.filter(function (cols) { return (cols[MAIL_COL] || "") === "是"; }).length;
      console.log("  ⏭ 跳过重复 " + duplicates.length + " 条，仅录入新增 " + pasteCount + " 条");
      if (pasteCount === 0) {
        out({
          status: "success",
          existing_count: scan.existingNums.length,
          tsv_count: tsvNums.length,
          duplicates_count: duplicates.length,
          new_count: 0,
          mail_count: 0,
          pasted_rows: 0,
          detail: "全部记录都已在表格中，无新增需要录入。"
        });
        return;
      }
    } else {
      out({
        status: "dedup_blocked",
        existing_count: scan.existingNums.length,
        tsv_count: tsvNums.length,
        duplicates_count: duplicates.length,
        new_count: newCount,
        mail_count: mailCount,
        duplicates: duplicates.slice(0, 30),
        hint: "发现重复，已停手。确认要强制录入全部请设 force=true；或只录新增请设 skip_duplicates=true 重跑。"
      });
      return;
    }
  }

  // ---- 导航到 lastRow+1 空行（按探活结果选导航方式）----
  var targetRow = scan.lastRow + 1;
  step(4, TOTAL, "导航到空行 row " + targetRow + "（" + probe.nav + " 直跳）...");
  await page.cua.click({ x: 25, y: 200 });
  await page.waitForTimeout(500);
  if (probe.nav === "setCurrentSelection") {
    await page.evaluate(function (r) {
      var app = window.SpreadsheetApp;
      app.view.setCurrentSelection({ yRange: [r, r], xRange: [0, 0] });
    }, targetRow);
  } else if (probe.nav === "setActiveCell") {
    await page.evaluate(function (r) {
      var app = window.SpreadsheetApp;
      var sid = app.workbook.worksheetManager.activeSheetId;
      var sheet = app.workbook.worksheetManager.getSheetBySheetId(sid);
      sheet.setActiveCell(r, 0);
    }, targetRow);
  }
  await page.waitForTimeout(300);

  // ---- ⭐ 防幽灵粘贴：粘贴前读回目标行确认空 ----
  step(5, TOTAL, "粘贴前校验目标行 row " + targetRow + " 为空（防幽灵粘贴）...");
  var preCheck = await page.evaluate(function (r) {
    var app = window.SpreadsheetApp;
    var sid = app.workbook.worksheetManager.activeSheetId;
    var sheet = app.workbook.worksheetManager.getSheetBySheetId(sid);
    var nonEmpty = 0;
    for (var c = 0; c < 18; c++) {
      var cell = sheet.getCellDataAtPosition(r, c);
      var v = cell && cell.formattedValue ? cell.formattedValue.value : (cell && cell.value != null ? cell.value : "");
      if (v !== "" && v != null) { nonEmpty++; }
    }
    return { nonEmptyCols: nonEmpty };
  }, targetRow);
  if (preCheck.nonEmptyCols > 0) {
    out({
      status: "target_row_not_empty",
      target_row: targetRow,
      non_empty_cols: preCheck.nonEmptyCols,
      hint: "目标行非空，可能服务器已有数据或残留态未清。已停手，避免覆盖。"
    });
    return;
  }
  console.log("  ✓ 目标行为空，可安全粘贴");

  // ---- 写剪贴板 + Ctrl+V ----
  step(6, TOTAL, "粘贴 " + pasteCount + " 条(18列)到 row " + targetRow + "...");
  var clipOk = await page.evaluate(function (text) {
    try { return navigator.clipboard.writeText(text).then(function () { return "ok"; }); }
    catch (e) { return "err:" + e.message; }
  }, tsv18);
  await page.waitForTimeout(300);
  await page.keyboard.press("Control+V");
  await page.waitForTimeout(3000);

  // ---- 读回验证列对齐（日期col2/发票号col4/订单号col9/邮件开票col17）----
  step(7, TOTAL, "读回验证列对齐（日期col2/发票号col4/订单号col9/邮件开票col17）...");
  var readBack = await page.evaluate(function (args) {
    var start = args.start, n = args.n;
    var app = window.SpreadsheetApp;
    var sid = app.workbook.worksheetManager.activeSheetId;
    var sheet = app.workbook.worksheetManager.getSheetBySheetId(sid);
    var rows = [];
    for (var r = start; r < start + n; r++) {
      var row = [];
      for (var c = 0; c < 18; c++) {
        var cell = sheet.getCellDataAtPosition(r, c);
        row.push(cell && cell.formattedValue ? cell.formattedValue.value : "");
      }
      rows.push(row);
    }
    return rows;
  }, { start: targetRow, n: pasteCount });

  var aligned = 0;
  for (var k = 0; k < readBack.length; k++) {
    var rb = readBack[k];
    var exp = lines[k];
    var dateOk = rb[2] && rb[2] !== "";
    var numOk = rb[4] && /^\d{8,}$/.test(String(rb[4]));
    var orderOk = rb[9] && rb[9] !== "";
    var expMail = (exp[MAIL_COL] || "");
    var gotMail = (rb[MAIL_COL] || "");
    var mailOk = gotMail === expMail;
    if (dateOk && numOk && orderOk && mailOk) aligned++;
  }
  console.log("  列对齐 " + aligned + "/" + pasteCount + " 条");
  if (aligned < pasteCount) {
    out({
      status: "alignment_failed",
      target_row: targetRow,
      pasted: pasteCount,
      aligned: aligned,
      readback_sample: readBack.slice(0, 3),
      hint: "部分行列未对齐，请人工核查。"
    });
    return;
  }

  // ---- 刷新验证持久化 ----
  step(8, TOTAL, "刷新页面验证持久化...");
  await page.goto(docUrl, { waitUntil: "domcontentloaded" });
  var r2 = await waitForAppReady(page, 20000);
  var s2 = await waitForSheetReady(page, 15000);
  if (!r2.ok || !s2.ok) {
    out({ status: "persistence_check_failed", detail: "刷新后引擎未就绪", target_row: targetRow, pasted: pasteCount });
    return;
  }
  // 刷新后同样等待数据加载稳定，再读回粘贴行
  await waitForDataStable(page, 15);
  var persist = await page.evaluate(function (args) {
    var start = args.start, n = args.n;
    var app = window.SpreadsheetApp;
    var sid = app.workbook.worksheetManager.activeSheetId;
    var sheet = app.workbook.worksheetManager.getSheetBySheetId(sid);
    var stillThere = 0;
    for (var r = start; r < start + n; r++) {
      var nCell = sheet.getCellDataAtPosition(r, 4);
      var num = nCell && nCell.formattedValue ? nCell.formattedValue.value : "";
      if (num && /^\d{8,}$/.test(String(num))) stillThere++;
    }
    return stillThere;
  }, { start: targetRow, n: pasteCount });

  var persistOk = persist === pasteCount;
  out({
    status: persistOk ? "success" : "persistence_failed",
    tsv_total: tsvNums.length,
    existing_count: scan.existingNums.length,
    duplicates_count: duplicates.length,
    new_count: newCount,
    mail_count: mailCount,
    pasted_rows: pasteCount,
    target_row: targetRow,
    aligned: aligned,
    persisted: persist,
    persistence_ok: persistOk,
    detail: persistOk
      ? "录入成功并已持久化。" + pasteCount + " 条落位 row " + targetRow + "-" + (targetRow + pasteCount - 1) + "（其中邮件开票=是 " + mailCount + " 条）"
      : "持久化异常：预期 " + pasteCount + " 条，刷新后实读 " + persist + " 条，请人工核查"
  });
}

// ---- 等待引擎就绪（轮询，非盲等）----
async function waitForAppReady(page, timeoutMs) {
  var start = Date.now();
  while (Date.now() - start < timeoutMs) {
    var ok = await page.evaluate(function () {
      return typeof window.SpreadsheetApp !== "undefined"
        && window.SpreadsheetApp
        && !!window.SpreadsheetApp.workbook
        && !!window.SpreadsheetApp.workbook.worksheetManager;
    });
    if (ok) return { ok: true, elapsed: Date.now() - start };
    await page.waitForTimeout(300);
  }
  return { ok: false, elapsed: timeoutMs };
}

async function waitForSheetReady(page, timeoutMs) {
  var start = Date.now();
  while (Date.now() - start < timeoutMs) {
    var ok = await page.evaluate(function () {
      try {
        var app = window.SpreadsheetApp;
        var sid = app.workbook.worksheetManager.activeSheetId;
        var sheet = app.workbook.worksheetManager.getSheetBySheetId(sid);
        return !!(sheet && typeof sheet.getRowCount === "function");
      } catch (e) { return false; }
    });
    if (ok) return { ok: true, elapsed: Date.now() - start };
    await page.waitForTimeout(500);
  }
  return { ok: false, elapsed: timeoutMs };
}
