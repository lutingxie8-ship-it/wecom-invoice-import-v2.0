---
name: wecom-invoice-import-v2.0
description: "把税务局导出的发票 Excel 批量录入到企微在线表格，并按开票人自动填写「是否邮件开票」列。当用户上传/提供一个税务局导出的 xlsx、或说'同步昨天的开票记录'、'开票实习生同步'、'把发票录入企微'、'导入发票到企微文档'、'批量录入开票记录'、'发票登记'、'录发票'时使用——只要是在同步/录入税务局开票记录，即使用户没明说'发票'也要用这个 skill。规则：开票人=符瑞香→直接录入；开票人=伍惠娟→录入并在「是否邮件开票」填'是'；其他开票人→报异常转人工。用 dev-browser 浏览器自动化+剪贴板粘贴，绕过企微无 API 权限/无企业认证的限制。"
version: 2.0.0
tier: write_with_safety_guard
priority: high
agent_created: true
---

# 企微发票录入 V2.0

把税务局导出的Excel发票记录，按开票人规则自动录入到企微在线表格末尾。

> 本 skill 是 `wecom-invoice-import`（V1）的升级版。**唯一区别**：录入前先按「开票人」分流，决定是否填写「是否邮件开票」列。其余（读Excel→查重→导航空行→粘贴→验证）与 V1 完全一致。

## 与 V1 的差异（一句话）

| | V1 `wecom-invoice-import` | V2.0 本 skill |
|---|---|---|
| 开票人筛选 | 无，全量录入 | ✅ 按开票人分流（见下） |
| 是否邮件开票列(col17) | 不填 | ✅ 伍惠娟填"是"，符瑞香留空 |
| 粘贴列数 | 10 列 | 18 列（新增 col17） |

## 开票人规则（V2.0 核心，必须遵守）

| 开票人 | 是否邮件开票(col17) | 处理 |
|---|---|---|
| `符瑞香` | 留空 | 直接录入（同 V1） |
| `伍惠娟` | `是` | 录入 + 邮件开票标记 |
| **其他任何值** | — | **不录入，报异常转人工判断** |

> ⚠️ 注意开票人姓名用字：税务局导出数据里是「伍**惠**娟」（恩惠的惠），不是「伍慧娟」。已按实际数据「伍惠娟」匹配。
> 若出现符瑞香/伍惠娟之外的开票人，`read_excel_to_tsv.py` 会打印异常明细并以退出码 2 结束，**停止录入**，把异常记录报告给用户转人工。

## 核心原理（必须理解）

企微在线表格是 **canvas渲染 + 协同编辑(mutation)** 架构：
- 引擎层 `setCellDataAtPosition` 只改内存，**不会提交到服务器**（刷新即丢）
- `keyboard.type()` 在canvas表格里**不响应**键盘事件
- `page.locator().click()` 被 `operate-board` 覆盖层**拦截**

**唯一可靠的写入方式**：`page.cua.click` 聚焦 → `navigator.clipboard.writeText` 写TSV → `Ctrl+V` 粘贴。粘贴走表格的正常paste事件处理，自动触发mutation提交到服务器。因此「是否邮件开票」也必须随 TSV 一起粘贴，不能单独 setCellData。

## 前提条件

1. **首次使用运行环境检查**：`python "<skill目录>/scripts/setup.py"`，自动检查并安装缺失依赖（openpyxl 等）。
2. **首次需扫码登录**：dev-browser 打开企微文档后用户扫码，登录态保持在 profile 里。

## 初始化配置（首次使用时向用户确认以下信息）

| # | 信息 | 说明 |
|---|------|------|
| 1 | 企微文档分享链接 | `https://doc.weixin.qq.com/sheet/xxx`。脚本支持从输入 JSON 的 `doc_url` 字段传入；缺省用脚本内默认链接 |

**Excel 来源**：由用户直接上传/提供 Excel 文件，取其绝对路径作为第1步的输入。**无需**配置文件夹路径或文件名规则。

## 执行流程（2步）

### 第1步：读取用户上传的Excel生成18列TSV（含开票人分流）

用户直接上传/提供 Excel 文件后，取其绝对路径运行：

```bash
python "<skill目录>/scripts/read_excel_to_tsv.py" "<用户上传的excel文件路径>" > tsv.tsv
```

脚本读取「信息汇总表」sheet，跳过表头和合计行，按开票人规则分流，输出 **18列TSV** 到 stdout，统计信息与异常明细到 stderr。

**⚠️ 务必检查 stderr**：
- 若出现 `⚠️ 发现 N 条异常开票人记录，需转人工判断！`（退出码 2）→ **停止**，把异常明细报告给用户转人工，不要录入异常记录。
- 正常情况 stderr 打印：`符瑞香（直接录入）X 条` / `伍惠娟（邮件开票=是）Y 条`。

**字段映射**（Excel「信息汇总表」 → 企微表格，列号均为企微表格引擎 0-based，即从 A=0 起）：

| 企微列 | Excel列(序号) | 处理 |
|--------|-------------|------|
| col0(空) | — | 空 |
| col1(公司) | — | 留空 |
| 开票日期(col2) | 9 | `2026-07-24 18:22:02` → `2026/7/24` |
| 发票代码(col3) | — | 留空 |
| 发票号码(col4) | 4 | 直接取 |
| 发票类型(col5) | 22 | 直接取 |
| 开票名称(col6) | 8 | 直接取 |
| 纳税人识别号(col7) | 7 | None→空字符串 |
| 开票金额(col8) | 20 | 直接取；**可能为负（红冲发票），照填** |
| 订单ID(col9) | 27 | 直接取；**可能为空（客户要求，属正常，不阻塞录入）** |
| col10~col16 | — | 留空（备注/归属项目/备注_1/订单ID是否重复/年/月/货物或应税劳务名称） |
| **是否邮件开票(col17)** | **26(开票人)** | **伍惠娟→"是"，符瑞香→""** |

### 第2步：写输入JSON + 运行录入脚本

**写输入文件**（把第1步的 18列TSV 内容塞进 `tsv` 字段）：

```
~/.dev-browser/tmp/wecom_import_v2_input.json
```
```json
{
  "tsv": "<第1步的18列TSV全文>",
  "doc_url": "https://doc.weixin.qq.com/sheet/...",
  "force": false,
  "skip_duplicates": false
}
```
- `tsv` 必填，18列TSV
- `doc_url` 可选，缺省用脚本内默认链接
- `force` 可选，默认 false；true=即使发现重复也强制录入全部（慎用，会重复）
- `skip_duplicates` 可选，默认 false；true=跳过重复行，**只录入新增**（适合"删了部分记录要补回"的场景，不会重复录入还在表格里的记录）

**运行脚本**：

```bash
dev-browser --browser wecom --idle-timeout 30m --timeout 240 run "<skill目录>/scripts/wecom_invoice_import.js"
```

> 本脚本是独立版（`waitForAppReady`/`waitForSheetReady` 等函数已内联），**无需 build_all.py 合并**，直接 `dev-browser run` 即可。

脚本分 8 步执行，每步打印 `[步骤 n/8]` 进度日志：
1. 全新加载企微文档（清残留态防幽灵粘贴）
2. 等待引擎+Sheet 就绪，**并探活关键接口**（自动探测新旧引擎结构、读写/导航接口，自动判定用 `setCurrentSelection` 还是 `setActiveCell`，并探测是否支持 `insertDimension` 插行）
3. 全表查重（**先等数据加载稳定**，再一次 evaluate 读全表发票号码，与 TSV 比对）
3.5 **行满检测**：若「目标行 + 待录条数」超出表格总行数，自动 `insertDimension` 扩行（详见技术要点 #17）
4. 导航到空行（按探活结果选导航方式，A列起，不Tab）
5. **粘贴前校验目标行为空**（防幽灵粘贴，非空即停手）
6. 粘贴 18列TSV
7. 读回验证列对齐（日期col2/发票号col4/订单号col9/**邮件开票col17**）
8. 刷新验证持久化

**读输出**：

```
~/.dev-browser/tmp/wecom_import_v2_output.json
```

| status | 含义 | 处理 |
|--------|------|------|
| `success` | 录入成功并已持久化（含 mail_count） | 完成 |
| `dedup_blocked` | 发现重复，已停手 | 看输出 `duplicates` 列表；如确要强录设 `force=true` 重跑；或由调用方过滤掉重复条目后重跑 |
| `target_row_not_empty` | 目标行非空（残留态或服务器已有数据） | 已停手避免覆盖，人工核查 |
| `alignment_failed` | 部分行列未对齐 | 看 `readback_sample`，人工核查 |
| `persistence_failed` | 刷新后数据丢失 | 粘贴可能未触发提交，重跑 |
| `app_not_ready` / `no_input` | 引擎未就绪 / 无输入 | 检查登录态或输入文件 |
| `engine_api_changed` | 探活发现关键接口缺失 | 看输出 `missing`（缺哪个接口）和 `dump`（现有方法列表/函数源码），据此更新脚本 |
| `row_full_manual_expand` | 表格行满，且引擎未暴露插行 API | 看输出 `need_rows`，手动在表格底部加相应行数后重试 |

## 关键技术要点（为什么这样做）

1. **全新加载防幽灵粘贴**：脚本开篇先 `page.goto(docUrl)` 全新加载，清掉浏览器内存里上一次导入尝试遗留的未提交 mutation。这是 v1 heredoc 版发生 row 61-84 脏数据事故的根因——键盘操作会触发残留 mutation 提交。

2. **18列TSV从A列起粘（防列偏移）**：直接从 col0 起粘完整 18 列，日期必落 col2、发票号必落 col4、订单号必落 col9、**是否邮件开票必落 col17**，零偏移风险。中间 col10~col16 全是空列占位。

3. **是否邮件开票必须随 TSV 粘贴**：`setCellDataAtPosition` 只改内存不提交，刷新即丢。所以 col17 必须在 TSV 里一起粘贴，由粘贴事件触发 mutation 提交。

4. **粘贴前读回确认空**：导航后、粘贴前，用引擎 API 读目标行 col0-17，有任何非空就停手——防幽灵粘贴和防误覆盖的最后一道闸。

5. **全量查重**：一次 evaluate 读全表所有发票号码（col4），与 TSV 比对。比"只查最近2日期"更安全，能拦住跨天补录的重复。

6. **⚠️ 数据渐进加载，查重/验证前必须等稳定**：企微表格的 canvas 数据是**异步分批渲染**的——`page.goto` 后立即扫描，`getCellDataAtPosition` 只能读到已加载的部分行（实测：刚加载完只读到 60 条，约 1.5s 后才到 82 条）。若不等稳定就查重，会漏掉尚未加载的记录，`lastRow` 也会算小，导致粘贴时**覆盖已有数据**。脚本用 `waitForDataStable()` 轮询 `lastRow`，连续两次一致才继续。

7. **⚠️ lastRow 必须以「发票号码(col4)」为锚，不能看「任意列非空」**：表格底部可能残留杂散数据（实测：row 5842 有一格孤立的日期 `2099/12/31`）。若用"任意列非空"判断最后一行，会把底部杂散数据误判成最后一条记录，导致粘贴到错误位置（row 5843 而不是 row 159）。所以 `fullScan` 里 `lastRow` 只在 col4（发票号码）非空时才更新，杂散数据自然被忽略。

8. **⚠️ 接口探活 + 导航双写**：企微引擎会不定期改内部接口（本次 `sheet.setActiveCell` 被移除）。脚本开头 `probeEngine()` 会检查关键接口是否都在，并自动判定导航方式——**优先 `app.view.setCurrentSelection({yRange:[r,r], xRange:[c,c]})`，退回老接口 `sheet.setActiveCell(r,0)`**。若关键接口缺失，`dumpEngine()` 会把 sheet/view 的方法列表和函数源码倒出来，输出 `engine_api_changed` 供定位新写法，而不是中途崩。注意 `app.view` 是**异步初始化**的，探活要轮询等它就绪。dev-browser 的 `page.evaluate` 只允许传 **1 个参数**，多参数要包进一个对象 `{start, n}`。

9. **开票人分流在 Python 侧完成**：Python 读 Excel 时顺带取「开票人」列(col26)，直接决定 col17 填"是"还是空。异常开票人在进入 JS 前就被拦下，JS 无需关心开票人。

10. **page.cua.click vs page.locator.click**：企微有 `operate-board` 覆盖层拦截DOM点击。page.cua.click 通过 CDP 发送原始鼠标事件，绕过拦截。

11. **Ctrl+V vs keyboard.type**：canvas表格编辑模式不响应 type 的键盘事件。粘贴走浏览器原生 paste 事件，表格有完整处理逻辑，自动触发 mutation 提交。

12. **必须刷新验证**：只有刷新后数据还在，才确认提交到服务器。

13. **引擎行号 vs UI行号**：引擎 row 0 = UI row 1（表头），引擎 row N = UI row N+1。`setCurrentSelection` 的 `yRange` 与引擎 `getCellDataAtPosition` 的 row 索引一致。

14. **订单号(col9)允许为空（客户要求）**：对齐校验只看 日期(col2) + 发票号(col4) + 邮件开票(col17)，**订单号不作为对齐判据**。若存在订单号为空的记录，照常录入，但完成后会在输出里带 `empty_order_count` / `empty_order_nums` 告知用户（哪些发票号订单号为空），供人工知晓。

15. **红冲发票（金额为负）直接录入、不中断**：红字发票（红冲）金额为负（col8，如 `-2380`），是冲销之前发票的正常业务，**照常录入、不中途打断询问**，金额为负也照填。脚本会在读回时统计红冲条数，完成后在输出里带 `redletter_count` / `redletter_list`（发票号+金额）供事后汇报，不在录入过程中停顿。

16. **⚠️ 引擎结构探测式兼容（新旧引擎取 sheet）**：企微表格引擎存在两套「取 sheet」结构，不同文档/版本走不同结构，**硬编码会 `app_not_ready`**：

    | | 旧引擎 | 新引擎 |
    |---|---|---|
    | 取 sheet 入口 | `app.workbook.worksheetManager` | `app.spreadsheet.sheetManager` |
    | 拿当前 sheet id | `activeSheetId`（**属性**） | `getActiveSheetId()`（**方法**） |
    | getSheetBySheetId / getRowCount / getCellDataAtPosition | ✅ | ✅ 一致 |

    脚本统一走 `window.__wbSheet.resolve(app)`（`installHelpers()` 挂到 window 全局，`goto` 刷新后需重挂），自动识别新旧引擎再取 sheet。`waitForAppReady`/`waitForSheetReady` 内联了同样的兼容判断（它们最先跑、helper 还没挂）。探活时会把 `engine` 类型打进日志。

17. **⚠️ 行满自动扩行（`sheet.insertDimension`，且是「尽力而为」）**：当表格贴满到最后一行时（`lastRow+1 >= getRowCount()`，即没有空行），脚本会检测「目标行 + 待录条数 超出总行数」，用 `sheet.insertDimension({index: 总行数, dimensionType: 1, count: 需要行数})` 在末尾内存扩行，再正常粘贴。

    - `dimensionType` 是**数字枚举**：`1` = 行，`0` = 列（实测确认，传字符串 `"ROW"` 等不生效）。
    - **已知限制**：`insertDimension` 和 `setCellDataAtPosition` 一样**只改内存、不保证提交**（实测刷新后行数回退）。所以脚本把它定位为「尽力而为」：若企微的粘贴提交恰好连带提交扩行，则全自动成功；若没有，**刷新验证会兜底**——数据丢了会报 `persistence_failed` 而不是静默失败。若引擎没暴露 `insertDimension`，则报 `row_full_manual_expand` 提示用户手动加行。
    - 想做到「真正可靠的行满自动加行」，需走 UI 模拟（点击表格底部加行按钮）或 command/mutation 提交，需在真实行满文档上进一步验证。

## 故障排查

| 问题 | 原因 | 解决 |
|------|------|------|
| SpreadsheetApp undefined | 未登录/页面未加载完 | 等待重试，或截图检查登录状态 |
| clipboard 返回 err | 页面无焦点/非HTTPS | 确保脚本开篇 cua.click 提供用户手势 |
| target_row_not_empty | 残留态未清或服务器已有数据 | 重跑（脚本会全新加载）；仍非空则人工核查该行 |
| alignment_failed | 粘贴锚点偏移（订单号为空不会再触发，已允许空） | 看 readback_sample 定位偏几列，检查 TSV 是否 18 列 |
| 刷新后数据丢失 | 粘贴没触发提交 | 确认选中了正确单元格再粘贴，重跑 |
| col17「是否邮件开票」没填上 | TSV 没带 col17 或列数不对 | 确认 read_excel_to_tsv.py 输出了 18 列，且第18列是"是"/空 |
| 退出码 2 | 发现异常开票人 | 看 stderr 异常明细，转人工，不要录入异常记录 |
| `sheet.setActiveCell is not a function` | 引擎已移除该 API | 改用 `app.view.setCurrentSelection({yRange:[r,r], xRange:[c,c]})` |
| 查重漏行 / lastRow 算小 | 数据渐进加载，扫描过早 | 脚本已用 `waitForDataStable` 等 lastRow 稳定；若仍漏，增大轮询次数/间隔 |
| `Too many arguments` (evaluate) | dev-browser 只允许 1 个参数 | 多参数包成对象 `{start, n}` 传入 |
| 新文档报 `app_not_ready` | 新文档用了新引擎结构 `app.spreadsheet.sheetManager` | 脚本已做新旧引擎探测兼容，无需手动改；若仍失败看 `engine_api_changed` 的 dump |
| 表格贴满贴不进（用户需手动加行） | 目标行超出总行数 | 脚本已加行满检测 + `insertDimension` 内存扩行；若引擎无插行 API 会报 `row_full_manual_expand`，仍需手动加行 |
| 扩行后刷新数据丢失 | `insertDimension` 只改内存未提交 | 属已知限制，脚本会报 `persistence_failed`；可靠方案需 UI 模拟加行或 command 提交 |

## 批量粘贴注意事项

- TSV 多行时，企微表格会自动从选中单元格开始向下填充多行
- 粘贴大量数据（如60行）后脚本等待 3 秒让表格处理完
- 如果数据量很大（100+行），考虑分批粘贴（每批50行）
- 符瑞香与伍惠娟的记录会混在同一批里，col17 各自正确填写，无需分两批

## 录入完成后告知用户的话术（订单号为空 / 红冲发票）

当输出 `empty_order_count > 0` 或 `redletter_count > 0` 时，录入照常完成，按下面模板向用户报告（先报结果 → 温和提示 → 给清单）：

**① 订单号为空**

> ⚠️ 其中 {M} 条记录的「订单号」为空（客户要求，属正常情况），已照常录入、没有跳过：
>
> | 发票号码 | 情况 |
> |---|---|
> | {发票号} | 订单号为空 |

**② 红冲发票（金额为负）**

> ⚠️ 其中 {R} 条为红冲发票（金额为负），已照常录入、未中断：
>
> | 发票号码 | 金额 |
> |---|---|
> | {发票号} | {负金额} |

统一开头：`✅ 录入完成，{N} 条已同步到企微文档。` 结尾：`其余均正常，不影响本次同步。`

**措辞要点**：
- 用「订单号为空（客户要求，属正常情况）」「红冲发票（金额为负）」这类中性描述，不要说「异常/错误」。
- 先报结果再报提示，避免用户误以为出错。
- 给具体发票号（+ 金额）清单，方便定位核对。

**数据来源**：脚本输出 `empty_order_count` / `empty_order_nums`（订单号空），`redletter_count` / `redletter_list`（红冲，含发票号+金额）。
