# 工时测量与负荷山积模板

## 人机作业分析自动排程模板

推荐文件：

`outputs/人机作业分析表/人机作业分析自动排程模板_v2.xlsm`

该模板支持最多 5 人、10 台设备、每台设备 20 个步骤，包含人员技能、设备间移动时间、自动排程、人机作业时间轴和改善前后对比。工作簿共 6 个页面：`使用说明`、`基础设置`、`作业步骤`、`自动排程`、`人机作业图`、`改善对比报告`。

打开文件后需启用宏。模板默认不带案例数据，可在`使用说明`点击`载入1人3机示例并试算`体验完整流程，或点击`新建空白分析`恢复干净模板。推荐操作顺序：

```text
新建空白 → 填基础设置 → 填作业步骤 → 检查数据
→ 生成初始方案 → 选择优化目标 → 优化排程
→ 保存改善前/后 → 刷新报告
```

四种优化目标为：最高产能、人员均衡、设备少等待、综合优化。内置 1 人 3 机案例的基准结果为 56 秒完成 3 件，平均 18.67 秒/件。

重新生成与验证：

```powershell
powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/build_template.ps1
powershell -ExecutionPolicy Bypass -File tools/man_machine_analysis/verify_template.ps1
```

本仓库用于整理和维护 Excel 工时测量、标准工时计算、工序负荷山积图分析模板，并提供一个移动端优先的在线工时测量工具。

项目主要资产包括 Excel 模板文件、`移动网页/` 在线录入应用、Netlify/Vercel 服务端接口，以及少量用于生成和验证模板的脚本。

## 推荐使用

### 在线工时测量工具

`移动网页/` 目录包含一个 Vite + React 应用，可独立部署到 Netlify 或 Vercel，用于在线录入现场测时数据并导出 Excel。

主要功能：

- 访问码保护，避免公开入口被随意使用。
- 按线体和型号新建测量单。
- 当前启用 `A线`、`B线`、`C线`、`D线`、`E线`、`H线` 和 `底座线`。
- `A线` 保留线体专用机型与工序预设；`B线`、`C线` 使用带 `盘座班` 的波轮结构；`D线`、`E线`、`H线` 使用滚筒结构；`底座线` 使用独立的 `底座线` 班组。
- 按班组录入工序名称、ST1-ST5 测量值、人数和备注。
- 支持工序新增、删除和拖动排序。
- 自动保存测量单，并可复制测量单链接继续填写。
- 支持导出当前班组或整机 Excel。
- 移动端界面已针对窄屏状态栏、操作菜单、工序横向录入和底部导出栏优化。
- 能识别移动键盘状态，避免底部导出栏遮挡录入区域。
- 导出文件基于 `移动网页/outputs/worktime_new_template/工时测量与负荷山积自动扩展模板_v16.xlsx` 生成。

本地运行：

```powershell
Set-Location 移动网页
npm install
npm run dev
```

测试、构建与导出校验：

```powershell
Set-Location 移动网页
npm test
npm run build
npm run verify:export
```

### Vercel 部署

从仓库连接 Vercel 时，Root Directory 选择 `移动网页`。

必需环境变量：

```text
UPSTASH_REDIS_REST_URL=Upstash Redis REST URL
UPSTASH_REDIS_REST_TOKEN=Upstash Redis REST Token
VITE_REQUIRE_ACCESS_CODE=false
```

也兼容 Vercel KV 风格的 `KV_REST_API_URL` 和 `KV_REST_API_TOKEN`。

Vercel 版将测量单保存到 Upstash Redis，并对创建、读取、保存和导出接口分别实施滑动窗口限流。若需要从旧 Netlify 站点按需迁移测量单，可额外配置：

```text
NETLIFY_MIGRATION_BASE_URL=https://旧站点域名
NETLIFY_MIGRATION_ACCESS_CODE=旧站点访问码
```

### Netlify 部署

从仓库连接 Netlify 时，站点 Base directory 选择 `移动网页`，并配置：

```text
SITE_ACCESS_CODE=你的访问码
```

### 工序负荷山积模板

推荐文件：

`outputs/worktime_new_template/工时测量与负荷山积自动扩展模板_v16.xlsx`

主要功能：

- 录入工序、人员、测量值等基础数据。
- 自动计算平均实测、标准工时、负荷 ST、TT 基准。
- 按负荷 ST 与 TT 基准判断状态：负荷低、正常、超节拍。
- 计算最大 CT 和波动风险。
- 新增简洁版工时表，自动汇总工序号、工序名称、平均实测和标准工时。
- 山积图随工序测量表实际行数自动扩展。
- 超过 TT 基准的工序在山积图中以红色显示。
- 山积分析页显示总工时、线平衡、瓶颈时间、TT 时间四项指标。
- 平均实测、标准工时、负荷 ST 及山积图柱形标签按一位小数显示。

### 动作要素堆叠山积模板

推荐文件：

`outputs/worktime_action_element_template/动作要素工时与堆叠山积图模板_v3.xlsm`

主要功能：

- 支持按工序录入多个动作要素。
- 每个工序默认 5 个动作要素，每个动作可填写 ST1-ST5 测量值。
- 以堆叠柱形图展示动作要素级平均 ST。
- 通过宏按钮支持新增工序、删除工序、刷新报表。
- 员工负荷表横向展示工序、动作、时间、总工时、目标节拍和实际节拍。

打开 `.xlsm` 文件时，需要在 Excel 中启用宏才能使用新增、删除和刷新按钮。

## 目录说明

- `移动网页/`
  - 在线工时测量工具独立子项目。
  - 包含前端源码、Netlify/Vercel 接口、部署配置、测试、导出校验脚本和在线导出所需的 v16 模板副本。

- `移动网页/src/`
  - 在线工时测量工具前端代码。

- `移动网页/netlify/functions/`
  - 访问校验、测量单保存和 Excel 导出的 Netlify Functions。

- `移动网页/api/`
  - Vercel Functions，包括 Redis 会话存储、接口限流和 Excel 导出。

- `移动网页/shared/`
  - Netlify 与 Vercel 共用的会话规范化和 Excel 导出逻辑。

- `移动网页/scripts/`
  - 导出结果校验脚本。

- `outputs/worktime_new_template/`
  - 工时测量与负荷山积自动扩展模板。
  - 当前推荐版本为 `v16`。

- `outputs/worktime_action_element_template/`
  - 动作要素工时与堆叠山积图模板。
  - 当前推荐版本为 `v3.xlsm`。

- `outputs/worktime_logic_test/`
  - 早期公式、山积图和表格逻辑的测试文件目录。
  - 当前仓库不再保留旧测试工作簿。

- `.codex_tmp/`
  - 生成、检查和验证模板时使用的辅助脚本。
  - 不是普通用户的使用入口。

- `AGENTS.md`
  - 项目维护说明、模板逻辑、历史问题处理记录和本地操作规则。

## 维护原则

- 不直接覆盖历史模板，优先另存新版本。
- 修改模板后重点检查公式、图表引用范围、TT 折线、负荷 ST 柱形图和线平衡率。
- 修改在线工具后在 `移动网页/` 下运行 `npm test` 和 `npm run build`，涉及导出逻辑时同时运行 `npm run verify:export`。
- 不批量删除文件或目录；如需删除文件，只删除单个明确路径的文件。
