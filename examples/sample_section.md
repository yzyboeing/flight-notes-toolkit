---
id: "9.1"
title_cn: "示例——源格式演示"
aircraft: ["NG", "738"]
source: ["演示用，非真实手册数据"]
related: ["9.2"]
updated: "2026-09-04"
---

## 9.1　示例——源格式演示

> 本文件仅演示源格式的各种写法，**数据全部为虚构占位值，不可用于任何实际运行**。

### 一、故障处置类表格（主判据合并列 + 前提行 + 末尾补充行）

<table class="ftn">
<tr class="hdr"><th>类型</th><th>现象</th><th>处置</th></tr>
<tr class="premise"><td colspan="3">正常离场 → <strong>限速 <em>230kt</em></strong> → 执行《<strong>示例故障</strong>》检查单。<br>
<em>本行是「共性前提」：对表内每一行都成立的入口条件，浅灰底 + 左侧竖条。</em></td></tr>
<tr><td rowspan="2">手柄在 <strong>UP</strong> 位</td>
    <td>所有<strong>红色和绿色</strong>指示灯亮</td>
    <td>核实盖板完全盖好 → 手柄 <strong>DN，然后 UP</strong> → 任一红灯亮 → <em>考虑返场落地</em>。</td></tr>
<tr><td>任何其他组合指示灯亮</td>
    <td><strong>无法判断部件具体位置</strong>（由 <em>A 系统压力</em>保持）→ <em>考虑返场落地</em>。</td></tr>
<tr><td>手柄在 <strong>OFF</strong> 位</td>
    <td>曾移到 UP 位</td>
    <td>手柄<strong>再次移到 UP 位</strong> → 任一红灯亮 → <em>考虑返场落地</em>。</td></tr>
<tr class="note"><td colspan="3">
<strong>本行是「末尾补充」：</strong>只在特定条件下用得上的经验、跨类别通用事项，放区块末尾通栏。<br>
<strong>三红三绿：</strong>可能是<em>弹簧老化</em>导致手柄无法自动回位，可尝试前推并下压。
</td></tr>
</table>

详见 9.2

### 二、参数表（数值列自动居中、短列自动收窄）

<table class="ftn">
<tr class="hdr"><th>运行条件</th><th>型号 A</th><th>型号 B</th></tr>
<tr><td>最大起飞限值</td><td><em>950℃</em></td><td><em>1038℃</em></td></tr>
<tr><td>最大连续限值</td><td>925℃</td><td>1013℃</td></tr>
<tr><td>最大地面起动限值</td><td>725℃</td><td>753℃</td></tr>
<tr><td>最大空中起动限值</td><td>遵守红线限制</td><td>883℃</td></tr>
</table>

### 三、时序流程类（序号列 + 环节 + 内容）

<table class="ftn">
<tr class="hdr"><th>步骤</th><th>操作</th></tr>
<tr class="premise"><td colspan="2"><strong>状况：</strong>示例触发条件之一成立。</td></tr>
<tr><td>①</td><td>自动油门（如接通）→ <strong>脱开</strong></td></tr>
<tr><td>②</td><td>推力手柄 → <strong>核实 → 收回</strong>，直至<em>参数不再超限</em></td></tr>
<tr><td>③</td><td>如果 <strong>示例警告灯</strong>保持亮 → 执行<strong>下一级检查单</strong>记忆项目</td></tr>
</table>

### 四、PF / PM 协同（角色作列）

<table class="ftn">
<tr class="hdr"><th>时机</th><th>PF</th><th>PM</th></tr>
<tr><td><em>400ft</em></td>
    <td>「<strong>示例喊话</strong>」→ 执行相应动作，<em>提前 10° 改出</em></td>
    <td>「<strong>400ft</strong>」；<br>核实<strong>示例参数</strong></td></tr>
<tr><td>稳定后</td>
    <td>「<strong>稳定进近</strong>」</td>
    <td>报告 ATC：「<em>示例通报用语</em>」</td></tr>
</table>

### 五、写法要点速览

<table class="ftn">
<tr class="hdr"><th>规则</th><th>说明</th></tr>
<tr><td>行首无项目符号</td><td>不用 <strong>▪ • - *</strong> 起项——成品复制进 OneNote 会被转成列表格式。层级用<strong>换行 + 加粗前缀标签</strong>表达</td></tr>
<tr><td>箭头串联时序</td><td>写成 <strong>A → B → C</strong> 单行，不拆成多行编号</td></tr>
<tr><td>序号只用于真步骤</td><td>用 ①②③ 放在独立的「序号」列，不写在正文行首</td></tr>
<tr><td>原理解释入括号</td><td>跟在动作后不另起行</td></tr>
<tr><td>时间用缩写</td><td><em>3min</em>、<em>5min</em>，不写「3 分钟」</td></tr>
<tr><td>并列项用「·」分隔</td><td>甲·乙·丙·丁，不用全角空格拉开</td></tr>
<tr><td>机型差异用【】</td><td>「关车前冷车时间【型号 A】」</td></tr>
<tr><td>检查单名用《》</td><td>便于检索</td></tr>
<tr class="note"><td colspan="2">
<strong>同名标签不可自嵌：</strong><code>&lt;em&gt;</code> 里不能再有 <code>&lt;em&gt;</code>。<code>&lt;em&gt;</code> 与 <code>&lt;strong&gt;</code> 可互嵌，红色优先。
</td></tr>
</table>
