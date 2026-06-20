# 人机作业分析 Excel VBA 自动化模板 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 交付一份最多支持 5 人、10 台设备、每台 20 步，具备技能约束、移动时间、四种优化目标、甘特图和改善对比的 Excel VBA `.xlsm` 模板。

**Architecture:** 使用 PowerShell + Excel COM 构建和验证宏启用工作簿；VBA 源码以可审查的 `.bas` 文件保存。核心排程模块只处理内存模型，不直接访问单元格；工作簿读写、校验、优化、图表和方案存档分别封装。测试通过 VBA 自测宏与 PowerShell 自动打开 Excel、运行宏、读取结果完成。

**Tech Stack:** Excel VBA、PowerShell、Excel COM Automation、结构化 Excel 表格、条件格式、堆积条形图、Git。

---

## 文件结构

```text
tools/man_machine_analysis/
  build_template.ps1                 # 创建工作簿、导入 VBA、绑定按钮并保存 xlsm
  verify_template.ps1                # 自动运行自测、检查工作表/表格/图表/公式错误
  vba/
    modDomain.bas                    # 任务、人员、设备、排程结果的数据结构与常量
    modValidation.bas                # 输入、技能、前置关系、锁定和矩阵校验
    modScheduler.bas                 # 可行初始排程与资源占用逻辑
    modScoring.bas                   # 四种目标的指标和评分
    modOptimizer.bas                 # 邻域生成、可行性检查和启发式搜索
    modWorkbookIO.bas                # 工作表与内存模型之间的转换
    modScenario.bas                  # 撤回、改善前、改善后方案快照
    modVisualization.bas             # 排程明细、甘特辅助数据、图表和报告刷新
    modCommands.bas                  # 按钮入口、状态栏和错误提示
    modSelfTest.bas                  # 无交互 VBA 回归测试
  fixtures/
    baseline_1p3m.csv                # 1 人 3 机基准步骤
outputs/man_machine_analysis/
  人机作业分析自动排程模板_v1.xlsm
```

生成和验证产生的日志、临时工作簿放在 `.codex_tmp/man_machine_analysis/`，不提交。禁止批量删除；若需要清理，只逐个删除明确文件。

### Task 1: 建立排程领域模型与基准测试入口

**Files:**
- Create: `tools/man_machine_analysis/vba/modDomain.bas`
- Create: `tools/man_machine_analysis/vba/modSelfTest.bas`
- Create: `tools/man_machine_analysis/fixtures/baseline_1p3m.csv`

- [ ] **Step 1: 写入 1 人 3 机基准数据**

```csv
DeviceId,StepNo,StepName,StepType,DurationSec,Skill
M1,1,上料,MANUAL,5,通用
M1,2,自动运行1,AUTO,20,
M1,3,翻面,MANUAL,6,通用
M1,4,自动运行2,AUTO,15,
M1,5,下料,MANUAL,5,通用
```

测试装载时复制到 M2、M3，并生成 3 个连续循环。

- [ ] **Step 2: 编写会失败的领域模型自测**

```vb
Public Sub Test_DomainConstants()
    AssertEqual STEP_MANUAL, "MANUAL", "人工步骤常量"
    AssertEqual STEP_AUTO, "AUTO", "自动步骤常量"
    AssertEqual MAX_PEOPLE, 5, "人员上限"
    AssertEqual MAX_DEVICES, 10, "设备上限"
    AssertEqual MAX_STEPS_PER_DEVICE, 20, "单设备步骤上限"
End Sub
```

- [ ] **Step 3: 在 Excel 中导入当前模块并运行测试，确认失败**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/build_template.ps1 -TestOnly
```

Expected: 失败，指出 `STEP_MANUAL` 或领域类型尚未定义。

- [ ] **Step 4: 实现领域模型和公共常量**

```vb
Public Const MAX_PEOPLE As Long = 5
Public Const MAX_DEVICES As Long = 10
Public Const MAX_STEPS_PER_DEVICE As Long = 20
Public Const STEP_MANUAL As String = "MANUAL"
Public Const STEP_AUTO As String = "AUTO"
Public Const STEP_JOINT As String = "JOINT"
Public Const STEP_WAIT As String = "WAIT"

Public Type TaskDef
    TaskId As String
    DeviceId As String
    CycleNo As Long
    StepNo As Long
    StepName As String
    StepType As String
    DurationSec As Double
    RequiredSkill As String
    PredecessorId As String
    LockedPersonId As String
    LockedStartSec As Double
    HasLockedStart As Boolean
End Type

Public Type ScheduledTask
    Definition As TaskDef
    PersonId As String
    StartSec As Double
    EndSec As Double
    MoveSec As Double
    WaitSec As Double
    IsLocked As Boolean
End Type
```

- [ ] **Step 5: 运行领域模型自测**

Run: `powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/build_template.ps1 -TestOnly`

Expected: `Test_DomainConstants PASS`。

- [ ] **Step 6: 提交**

```powershell
git add tools/man_machine_analysis/vba/modDomain.bas tools/man_machine_analysis/vba/modSelfTest.bas tools/man_machine_analysis/fixtures/baseline_1p3m.csv
git commit -m "test: add man-machine scheduling domain model"
```

### Task 2: 构建 5 页工作簿骨架与输入表

**Files:**
- Create: `tools/man_machine_analysis/build_template.ps1`
- Create: `tools/man_machine_analysis/vba/modWorkbookIO.bas`
- Modify: `tools/man_machine_analysis/vba/modSelfTest.bas`

- [ ] **Step 1: 添加工作簿结构自测**

```vb
Public Sub Test_WorkbookStructure()
    AssertSheetExists "基础设置"
    AssertSheetExists "作业步骤"
    AssertSheetExists "自动排程"
    AssertSheetExists "人机作业图"
    AssertSheetExists "改善对比报告"
    AssertTableExists "基础设置", "tblPeople"
    AssertTableExists "基础设置", "tblDevices"
    AssertTableExists "基础设置", "tblMoveTime"
    AssertTableExists "作业步骤", "tblSteps"
End Sub
```

- [ ] **Step 2: 运行并确认缺少工作表而失败**

Run: `powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/build_template.ps1 -TestOnly`

Expected: `基础设置 sheet missing`。

- [ ] **Step 3: 实现工作簿创建器**

`build_template.ps1` 必须：

```powershell
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$book = $excel.Workbooks.Add()
$sheetNames = @('基础设置','作业步骤','自动排程','人机作业图','改善对比报告')
```

并创建：

- `tblParameters`：方案名称、循环数、优化目标、搜索迭代、四项综合权重。
- `tblPeople`：人员编号、人员名称、技能、可操作设备、启用。
- `tblDevices`：设备编号、设备名称、产品、关系类型、前置设备、启用。
- `tblMoveTime`：起点及 M1-M10 移动秒数矩阵。
- `tblSteps`：设备、步骤号、名称、类型、工时、前置步骤、技能、锁定人员、锁定开始、允许等待、备注。

使用微软雅黑；必填黄色 `RGB(255,242,204)`，选填浅蓝 `RGB(221,235,247)`，输出灰色 `RGB(242,242,242)`。

- [ ] **Step 4: 添加数据验证与命名区域**

```powershell
$stepTypeFormula = '"人工,自动运行,人机协同,等待"'
$targetFormula = '"最高产能,人员均衡,设备少等待,综合优化"'
$relationFormula = '"独立循环,有先后顺序"'
```

为步骤类型、优化目标、设备关系、启用状态添加下拉；建立 `nmOptimizationTarget`、`nmCycleCount` 和四个权重命名区域。

- [ ] **Step 5: 运行结构自测**

Expected: `Test_WorkbookStructure PASS`，工作簿恰好 5 个可见工作表。

- [ ] **Step 6: 提交**

```powershell
git add tools/man_machine_analysis/build_template.ps1 tools/man_machine_analysis/vba/modWorkbookIO.bas tools/man_machine_analysis/vba/modSelfTest.bas
git commit -m "feat: build five-sheet man-machine workbook"
```

### Task 3: 实现输入读取与严格校验

**Files:**
- Create: `tools/man_machine_analysis/vba/modValidation.bas`
- Modify: `tools/man_machine_analysis/vba/modWorkbookIO.bas`
- Modify: `tools/man_machine_analysis/vba/modSelfTest.bas`

- [ ] **Step 1: 添加校验失败测试**

```vb
Public Sub Test_ValidationRejectsBadInputs()
    LoadFixture "DUPLICATE_STEP"
    AssertContains ValidateCurrentModel(), "步骤序号重复", "重复步骤"
    LoadFixture "NO_SKILLED_PERSON"
    AssertContains ValidateCurrentModel(), "无合格人员", "技能不匹配"
    LoadFixture "CYCLIC_PREDECESSOR"
    AssertContains ValidateCurrentModel(), "循环依赖", "前置循环"
    LoadFixture "NEGATIVE_MOVE"
    AssertContains ValidateCurrentModel(), "移动时间不能为负数", "移动矩阵"
End Sub
```

- [ ] **Step 2: 运行并确认校验器未定义**

Expected: FAIL with `Sub or Function not defined: ValidateCurrentModel`。

- [ ] **Step 3: 实现工作表到模型的读取**

`ReadPeople`、`ReadDevices`、`ReadMoveMatrix`、`ReadSteps` 逐行读取结构化表格；跳过“启用=否”的资源；将显示文字映射为内部常量：

```vb
Select Case displayType
    Case "人工": internalType = STEP_MANUAL
    Case "自动运行": internalType = STEP_AUTO
    Case "人机协同": internalType = STEP_JOINT
    Case "等待": internalType = STEP_WAIT
End Select
```

- [ ] **Step 4: 实现校验规则**

`ValidateCurrentModel` 返回换行分隔的全部错误，覆盖：

- 上限、空工时、非正工时、重复步骤。
- 前置步骤不存在或循环。
- 人工/协同步骤无合格人员。
- 锁定人员技能不符。
- 锁定开始时间早于前置完成。
- 移动矩阵缺值或负值。
- 设备关系循环。

- [ ] **Step 5: 运行校验测试**

Expected: `Test_ValidationRejectsBadInputs PASS`。

- [ ] **Step 6: 提交**

```powershell
git add tools/man_machine_analysis/vba/modValidation.bas tools/man_machine_analysis/vba/modWorkbookIO.bas tools/man_machine_analysis/vba/modSelfTest.bas
git commit -m "feat: validate scheduling inputs"
```

### Task 4: 实现可行初始排程

**Files:**
- Create: `tools/man_machine_analysis/vba/modScheduler.bas`
- Modify: `tools/man_machine_analysis/vba/modSelfTest.bas`

- [ ] **Step 1: 添加资源冲突和基准案例测试**

```vb
Public Sub Test_BaselineScheduleIsFeasible()
    LoadFixture "BASELINE_1P3M"
    Dim result As ScheduleResult
    result = BuildInitialSchedule(ReadCurrentModel())
    AssertTrue result.IsFeasible, "基准方案可排"
    AssertNoPersonOverlap result
    AssertNoDeviceOverlap result
    AssertPrecedenceHeld result
End Sub
```

- [ ] **Step 2: 运行并确认 `BuildInitialSchedule` 未定义**

Expected: FAIL。

- [ ] **Step 3: 实现最早可行时间计算**

```vb
candidateStart = Max3(predecessorEnd, deviceAvailable(deviceId), _
    personAvailable(personId) + MoveTime(lastLocation(personId), deviceId))
```

人工步骤占用人员；自动步骤仅占设备；协同步骤同时占用人员和设备；等待步骤只推进工艺时间。

- [ ] **Step 4: 实现人员选择**

对所有合格人员计算“最早开始时间 + 移动时间 + 当前负荷惩罚”，选择最小值；锁定人员优先且不可被替换。

- [ ] **Step 5: 实现连续循环任务展开**

按 `计划分析循环数` 复制任务，并将设备同一步骤的前后循环关系连接，确保可识别稳定循环边界和产出。

- [ ] **Step 6: 运行初始排程测试**

Expected: 所有资源与前置关系测试 PASS。

- [ ] **Step 7: 提交**

```powershell
git add tools/man_machine_analysis/vba/modScheduler.bas tools/man_machine_analysis/vba/modSelfTest.bas
git commit -m "feat: generate feasible man-machine schedules"
```

### Task 5: 实现指标、四种评分与启发式优化

**Files:**
- Create: `tools/man_machine_analysis/vba/modScoring.bas`
- Create: `tools/man_machine_analysis/vba/modOptimizer.bas`
- Modify: `tools/man_machine_analysis/vba/modSelfTest.bas`

- [ ] **Step 1: 添加指标和目标差异测试**

```vb
Public Sub Test_ScoringTargets()
    Dim base As ScheduleResult, optimized As ScheduleResult
    base = BuildInitialSchedule(ReadFixtureModel("BASELINE_1P3M"))
    optimized = OptimizeSchedule(base, TARGET_THROUGHPUT, 500, 12345)
    AssertTrue optimized.Metrics.CycleTimeSec <= base.Metrics.CycleTimeSec, "产能目标不恶化周期"
    AssertTrue ScoreSchedule(base, TARGET_BALANCE) <> ScoreSchedule(base, TARGET_DEVICE_WAIT), "目标评分不同"
End Sub
```

- [ ] **Step 2: 运行并确认评分模块未定义**

Expected: FAIL。

- [ ] **Step 3: 实现指标**

```vb
metrics.TaktSec = metrics.CycleTimeSec / metrics.OutputCount
metrics.HourlyCapacity = 3600# / metrics.TaktSec
metrics.PersonLoad(personId) = workSec / analysisSec
metrics.DeviceUtil(deviceId) = productiveSec / analysisSec
```

同时计算移动、人员空闲、设备等待人员、设备空闲、瓶颈人员、瓶颈设备和最长等待步骤。

- [ ] **Step 4: 实现四种评分**

```vb
Select Case target
    Case TARGET_THROUGHPUT
        score = cycleTime + 0.1 * totalWait + 0.05 * totalMove
    Case TARGET_BALANCE
        score = 1000 * personLoadStdDev + cyclePenalty + overloadPenalty
    Case TARGET_DEVICE_WAIT
        score = totalDeviceWait + 2 * maxDeviceWait + cyclePenalty
    Case TARGET_COMPOSITE
        score = wCycle * NormCycle + wBalance * NormBalance _
              + wWait * NormWait + wMove * NormMove
End Select
```

- [ ] **Step 5: 实现确定性邻域搜索**

候选动作：

- 交换两个未锁定、同优先级任务。
- 将任务改派给另一合格人员。
- 调整设备首次启动偏移。

使用固定随机种子重现测试；每个候选必须通过 `IsScheduleFeasible`，仅接受更低评分。

- [ ] **Step 6: 运行四目标测试**

Expected: 指标计算 PASS；同一随机种子结果一致；锁定任务未变化。

- [ ] **Step 7: 提交**

```powershell
git add tools/man_machine_analysis/vba/modScoring.bas tools/man_machine_analysis/vba/modOptimizer.bas tools/man_machine_analysis/vba/modSelfTest.bas
git commit -m "feat: optimize schedules for four objectives"
```

### Task 6: 实现按钮命令、锁定、撤回与方案存档

**Files:**
- Create: `tools/man_machine_analysis/vba/modScenario.bas`
- Create: `tools/man_machine_analysis/vba/modCommands.bas`
- Modify: `tools/man_machine_analysis/build_template.ps1`
- Modify: `tools/man_machine_analysis/vba/modSelfTest.bas`

- [ ] **Step 1: 添加锁定和快照测试**

```vb
Public Sub Test_LocksAndSnapshots()
    Dim before As ScheduleResult, after As ScheduleResult
    before = BuildInitialSchedule(ReadFixtureModel("BASELINE_1P3M"))
    LockTask before.Tasks(1).Definition.TaskId, before.Tasks(1).PersonId, before.Tasks(1).StartSec
    SaveUndoSnapshot before
    after = OptimizeSchedule(before, TARGET_COMPOSITE, 200, 7)
    AssertLockedTaskUnchanged before, after, 1
    AssertScheduleEqual RestoreUndoSnapshot(), before
End Sub
```

- [ ] **Step 2: 运行并确认快照函数未定义**

Expected: FAIL。

- [ ] **Step 3: 实现方案快照**

将排程结果序列化到隐藏命名区域或 `VeryHidden` 数据区，保存：

- `UNDO`
- `BEFORE`
- `AFTER`

不得增加第六个可见工作表。

- [ ] **Step 4: 实现用户命令**

```vb
Public Sub CmdCheckData()
Public Sub CmdBuildInitial()
Public Sub CmdOptimize()
Public Sub CmdToggleLock()
Public Sub CmdUndo()
Public Sub CmdSaveBefore()
Public Sub CmdSaveAfter()
Public Sub CmdRefreshReport()
```

所有命令使用统一错误处理，状态写入`自动排程`页，不吞掉错误。

- [ ] **Step 5: 在构建器中创建并绑定按钮**

按钮绑定以上宏，统一蓝色主题；锁定按钮根据选择行切换状态。

- [ ] **Step 6: 运行快照与按钮存在测试**

Expected: 宏可由 `Application.Run` 调用；锁定不变；撤回恢复原方案。

- [ ] **Step 7: 提交**

```powershell
git add tools/man_machine_analysis/vba/modScenario.bas tools/man_machine_analysis/vba/modCommands.bas tools/man_machine_analysis/build_template.ps1 tools/man_machine_analysis/vba/modSelfTest.bas
git commit -m "feat: add scheduling controls and scenarios"
```

### Task 7: 生成人机作业图与改善对比报告

**Files:**
- Create: `tools/man_machine_analysis/vba/modVisualization.bas`
- Modify: `tools/man_machine_analysis/vba/modCommands.bas`
- Modify: `tools/man_machine_analysis/vba/modSelfTest.bas`

- [ ] **Step 1: 添加输出测试**

```vb
Public Sub Test_VisualizationOutputs()
    LoadFixture "BASELINE_1P3M"
    CmdBuildInitial
    RefreshGanttView "全部"
    AssertTrue SheetTableRowCount("自动排程", "tblSchedule") > 0, "排程明细"
    AssertChartExists "人机作业图", "chtManMachine"
    AssertMetricIsNumeric "改善对比报告", "循环周期"
End Sub
```

- [ ] **Step 2: 运行并确认图表和指标缺失**

Expected: FAIL。

- [ ] **Step 3: 输出排程明细**

填充 `tblSchedule`：任务、循环、设备、步骤、人员、开始、结束、持续、类型、前置、锁定和等待原因。

- [ ] **Step 4: 生成甘特辅助数据和图表**

按资源逐行构造“开始偏移 + 持续时间”的堆积条形图；开始偏移系列无填充；任务系列按类型着色。提供`全部`、`仅人员`、`仅设备`视图；时间轴上限按排程最大结束时间向上取整。

- [ ] **Step 5: 生成改善对比**

报告表至少包含：

```text
指标 | 改善前 | 改善后 | 差值 | 改善率
循环周期(s)
平均节拍(s/件)
小时产能(pcs/h)
平均人员负荷率
最大人员负荷率
设备等待人员(s)
总移动时间(s)
```

改善建议使用可追溯规则，例如：`最大人员负荷率 > 90%` 时提示检查人员过载；`设备等待人员占比 > 10%` 时提示调整服务顺序或技能覆盖。

- [ ] **Step 6: 运行可视化测试**

Expected: 表格非空、图表存在、指标为数值、三种视图可刷新。

- [ ] **Step 7: 提交**

```powershell
git add tools/man_machine_analysis/vba/modVisualization.bas tools/man_machine_analysis/vba/modCommands.bas tools/man_machine_analysis/vba/modSelfTest.bas
git commit -m "feat: add man-machine gantt and comparison report"
```

### Task 8: 完成模板构建、基准案例和自动验证

**Files:**
- Create: `tools/man_machine_analysis/verify_template.ps1`
- Modify: `tools/man_machine_analysis/build_template.ps1`
- Modify: `tools/man_machine_analysis/vba/modSelfTest.bas`
- Create: `outputs/man_machine_analysis/人机作业分析自动排程模板_v1.xlsm`

- [ ] **Step 1: 在构建器中导入全部 VBA 模块**

启用 Excel 对 VBA 工程对象模型的信任后，逐个导入 `.bas`；如果权限未启用，脚本必须给出明确提示并停止，不生成残缺模板。

- [ ] **Step 2: 写入完整 1 人 3 机示例**

基础设置预置 1 人、3 机、通用技能、设备间移动时间 0；作业步骤预置三台设备的五步流程；说明区注明视频案例预期为约 56 秒/3 件、18.67 秒/件。

- [ ] **Step 3: 保存为新文件**

```powershell
$output = Join-Path $repoRoot 'outputs\man_machine_analysis\人机作业分析自动排程模板_v1.xlsm'
$book.SaveAs($output, 52) # xlOpenXMLWorkbookMacroEnabled
```

不得覆盖现有 v16 或动作要素 v3。

- [ ] **Step 4: 实现自动验证脚本**

`verify_template.ps1` 打开最终 `.xlsm`，执行：

```powershell
$excel.Run("'人机作业分析自动排程模板_v1.xlsm'!RunAllSelfTests")
$excel.Run("'人机作业分析自动排程模板_v1.xlsm'!CmdBuildInitial")
$excel.Run("'人机作业分析自动排程模板_v1.xlsm'!CmdOptimize")
```

并检查：

- 5 个工作表名称正确。
- 所有必需表格和按钮存在。
- VBA 自测结果全 PASS。
- 排程无人员/设备重叠。
- 基准循环约 56 秒，平均节拍约 18.67 秒/件，使用允许的浮点容差。
- 甘特图存在且数据范围非空。
- 关键区域无 `#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#N/A`。
- 最终文件为 `.xlsm` 且 VBA 工程存在。

- [ ] **Step 5: 运行完整构建与验证**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/build_template.ps1
powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/verify_template.ps1
```

Expected:

```text
SELF_TESTS: PASS
BASELINE_CYCLE_SEC: 56.0
BASELINE_TAKT_SEC: 18.67
FORMULA_ERRORS: 0
VERIFY: PASS
```

- [ ] **Step 6: 人工视觉检查**

用 Excel 打开最终文件并逐页检查：

- 表头、输入色彩和微软雅黑字体。
- 关键文字和数字无裁切。
- 甘特图颜色、图例、时间轴和标签可读。
- 改善报告在常用缩放比例下可阅读。
- 宏按钮可点击，错误提示指向具体输入行。

- [ ] **Step 7: 提交最终实现**

```powershell
git add tools/man_machine_analysis outputs/man_machine_analysis/人机作业分析自动排程模板_v1.xlsm
git commit -m "feat: deliver automated man-machine analysis workbook"
```

### Task 9: 最终回归与项目文档更新

**Files:**
- Modify: `README.md`
- Modify: `AGENTS.md`

- [ ] **Step 1: 更新 README 使用入口**

添加最终模板路径、5 个工作表、启用宏要求、四种优化目标和常用操作顺序：

```text
检查数据 → 生成初始方案 → 选择优化目标 → 优化排程
→ 保存改善前/后 → 刷新报告
```

- [ ] **Step 2: 更新 AGENTS.md 维护说明**

记录：

- 推荐模板版本与输出路径。
- VBA 源码和构建/验证脚本路径。
- 上限、颜色约定和指标定义。
- 修改后必须运行的两个 PowerShell 命令。
- 不覆盖历史模板和禁止批量删除规则。

- [ ] **Step 3: 运行最终回归**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/build_template.ps1
powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/verify_template.ps1
git diff --check
git status --short
```

Expected: 验证全部 PASS；无空白错误；仅预期文件发生变化。

- [ ] **Step 4: 提交文档**

```powershell
git add README.md AGENTS.md
git commit -m "docs: document man-machine analysis workbook"
```

## 实施约束

- 不使用 `Remove-Item -Recurse`、`rm -rf` 或任何批量删除命令。
- 不覆盖现有 Excel 模板。
- 自动排程必须先通过数据校验。
- 锁定约束发生冲突时必须报告无解，不能擅自放松。
- 测试使用固定随机种子；用户运行优化可以使用可记录的随机种子。
- 工作簿保持 5 个可见工作表；内部快照使用命名区域或 `VeryHidden` 区域。
- 启发式算法不宣称数学全局最优，只交付可行、可解释、可复现的较优方案。
