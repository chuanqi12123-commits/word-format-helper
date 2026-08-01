<script setup>
import { ref, reactive, computed } from 'vue'
import UploadZone from './components/UploadZone.vue'
import HeadingTree from './components/HeadingTree.vue'
import LogPanel from './components/LogPanel.vue'
import { CLEAN_DEFAULT, CLEAN_ITEMS, TEMPLATES } from './utils/templates.js'
import { previewHeadings, processFile } from './utils/pipeline.js'
import { downloadBlob } from './utils/docxIO.js'
import { logger } from './utils/logger.js'

// ========== 状态 ==========
const file = ref(null)
const busy = ref(false)
const tree = ref([])
const result = ref(null)
const parsed = ref(false)
const report = ref(null)

// ========== 流程区块（单页滚动，锚点定位） ==========
const SECTIONS = [
  { id: 'upload', no: '1', name: '上传文档', icon: '📄', desc: '选择要处理的 Word 文件' },
  { id: 'clean', no: '2', name: '格式清洗', icon: '🧹', desc: '选择需要清理的项目' },
  { id: 'inspect', no: '3', name: '处理结果', icon: '🔍', desc: '查看检测与处理数据' },
  { id: 'export', no: '4', name: '导出下载', icon: '📤', desc: '下载处理好的文档' }
]

// 锚点滚动导航（单页无需跳转，直接滚动到对应区块）
function scrollTo(id) {
  const el = document.getElementById('sec-' + id)
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' })
}

// ========== 清洗配置 ==========
const cleanEnabled = ref(true)
const cleanOpts = reactive({ ...CLEAN_DEFAULT })
const cleanCount = computed(() => CLEAN_ITEMS.filter((i) => cleanOpts[i.key] === true).length)

// 一键预设
const PRESETS = [
  { id: 'all', name: '全面清洗', desc: '启用全部清洗项', apply: () => Object.assign(cleanOpts, { ...CLEAN_DEFAULT }) },
  { id: 'chars', name: '仅字符清理', desc: '零宽/软连字符/控制字符', apply: () => Object.assign(cleanOpts, { ...CLEAN_DEFAULT }, {
    leadingSpace: false, trailingSpace: false, nbspToSpace: false, tabToSpace: false,
    clearHyperlinks: false, clearColor: false, clearHighlight: false, hiddenText: 'keep', mergeSpaces: false
  }) },
  { id: 'layout', name: '排版清理', desc: '行首行尾空格 + 制表符', apply: () => Object.assign(cleanOpts, {
    zeroWidth: false, softHyphen: false, controlChars: false, nbspToSpace: false,
    clearHyperlinks: false, clearColor: false, clearHighlight: false, hiddenText: 'keep', mergeSpaces: false
  }) },
  { id: 'none', name: '不清洗', desc: '仅执行后续步骤', apply: () => Object.assign(cleanOpts, { ...CLEAN_DEFAULT }, {
    zeroWidth: false, softHyphen: false, controlChars: false, leadingSpace: false, trailingSpace: false,
    nbspToSpace: false, tabToSpace: false, clearHyperlinks: false, clearColor: false, clearHighlight: false,
    hiddenText: 'keep', mergeSpaces: false
  }) }
]
const presetSel = ref('all')
function applyPreset(p) {
  presetSel.value = p.id
  p.apply()
}

// 清洗规则分组
const CLEAN_GROUPS = [
  { name: '字符清理', keys: ['zeroWidth', 'softHyphen', 'controlChars', 'nbspToSpace', 'tabToSpace', 'mergeSpaces'] },
  { name: '排版清理', keys: ['leadingSpace', 'trailingSpace'] },
  { name: '格式清理', keys: ['clearHyperlinks', 'clearColor', 'clearHighlight', 'hiddenText', 'emptyNumbered'] }
]
const cleanGroups = CLEAN_GROUPS.map((g) => ({ ...g, items: g.keys.map((k) => CLEAN_ITEMS.find((i) => i.key === k)).filter(Boolean) }))
const groupCount = (g) => g.keys.filter((k) => cleanOpts[k] === true).length

// ========== 模板格式化 ==========
const formatEnabled = ref(true)
const templateId = ref('special')
const borderSz = ref('8')
const headerFill = ref('D9D9D9')
const cellAlign = ref('center')
const headerBold = ref(true)
const tblAlign = ref('left')
const tblLayout = ref('autofit')
const numberEnabled = ref(true)
const textDetect = ref(true)
const bulletsEnabled = ref(true)
const template = computed(() => TEMPLATES[templateId.value])

const templateDetails = {
  special: [
    '纸张：A4；页边距：上/下 2.54cm，左/右 3.18cm',
    '字体：中文宋体，英文 Times New Roman；行距统一 1.5 倍',
    '一级标题：二号宋体加粗居中，编号 一、二、三…',
    '二级标题：三号宋体加粗左对齐，编号 （一）（二）…',
    '三级标题：四号宋体加粗左对齐，编号 1. 2. …',
    '四级标题：小四号宋体加粗左对齐，编号 （1）（2）…',
    '正文：小四号宋体，两端对齐，首行缩进 2 字符',
    '页脚：底部居中页码'
  ],
  general: [
    '纸张：A4；页边距：上/下 2.40cm，左/右 2.50cm',
    '字体：微软雅黑；正文 10.5pt，首行缩进 0.74cm（2 字符）',
    '一级标题：18pt 微软雅黑加粗；二级 16pt；三级 15pt；四级 14pt',
    '标题行距 1.2 倍，段前 2.5pt、段后 1pt，无首行缩进',
    '页脚：底部居中页码'
  ]
}

// ========== 质检统计 ==========
const headingTotal = computed(() => {
  let n = 0
  const walk = (arr) => {
    for (const x of arr) {
      n++
      if (x.children && x.children.length) walk(x.children)
    }
  }
  walk(tree.value)
  return n
})
function countLevel(level) {
  let n = 0
  const walk = (arr) => {
    for (const x of arr) {
      if (x.level === level) n++
      if (x.children && x.children.length) walk(x.children)
    }
  }
  walk(tree.value)
  return n
}
const sizeText = (n) => (n > 1048576 ? (n / 1048576).toFixed(2) + ' MB' : (n / 1024).toFixed(1) + ' KB')

const CLEAN_STAT_LABELS = {
  zeroWidth: '零宽字符', softHyphen: '软连字符', controlChars: '控制字符',
  leadingSpace: '行首空格', trailingSpace: '行尾空格', nbspToSpace: '不间断空格',
  tabToSpace: '制表符', clearHyperlinks: '超链接', clearColor: '手动颜色',
  clearHighlight: '手动高亮', hiddenDelete: '隐藏文字删除', hiddenDisplay: '隐藏文字显示',
  mergeSpaces: '多余空格', emptyNumbered: '空编号行'
}
const cleanStatItems = computed(() => {
  const items = []
  if (report.value?.clean?.items) {
    for (const [k, v] of Object.entries(report.value.clean.items)) {
      if (v > 0) items.push({ label: CLEAN_STAT_LABELS[k] || k, value: v })
    }
  }
  return items
})

// ========== 事件 ==========
async function onFileSelect(f) {
  file.value = f
  result.value = null
  report.value = null
  tree.value = []
  parsed.value = false
  logger.clear()
  logger.info(`已选择文件：${f.name}`)
  busy.value = true
  try {
    const { tree: t } = await previewHeadings(f)
    tree.value = t
    parsed.value = true
    logger.info(`标题检测完成，共 ${t.length} 个一级标题`)
  } catch (e) {
    logger.error('文件解析失败：' + (e.message || e))
  } finally {
    busy.value = false
  }
}

function clearFile() {
  file.value = null
  result.value = null
  report.value = null
  tree.value = []
  parsed.value = false
  logger.clear()
}

async function run() {
  if (!file.value || busy.value) return
  busy.value = true
  result.value = null
  report.value = null
  logger.info('🚀 开始处理……')
  try {
    const res = await processFile(file.value, {
      clean: cleanEnabled.value,
      cleanOptions: { ...cleanOpts },
      number: numberEnabled.value,
      numberOptions: { textDetect: textDetect.value, bulletsEnabled: bulletsEnabled.value },
      format: formatEnabled.value,
      template: template.value,
      formatOptions: {
        borderSz: Number(borderSz.value),
        headerFill: headerFill.value,
        cellAlign: cellAlign.value,
        headerBold: headerBold.value,
        tblAlign: tblAlign.value,
        tblLayout: tblLayout.value
      }
    })
    result.value = res
    report.value = res.report
    tree.value = res.tree
    // 处理完成后自动滚动到「处理结果」区块
    scrollTo('inspect')
  } catch (e) {
    logger.error('处理失败：' + (e.message || e))
  } finally {
    busy.value = false
  }
}

function goInspect() {
  run()
}

function download() {
  if (result.value) {
    downloadBlob(result.value.blob, result.value.outName)
    logger.info('⬇️ 已触发下载')
  }
}

// 生成质检报告文本（txt）
function exportReport() {
  if (!report.value) return
  const r = report.value
  const lines = []
  lines.push('Word 格式化小助手 - 质检报告')
  lines.push('='.repeat(40))
  lines.push(`处理文件：${file.value?.name || '未知'}`)
  lines.push(`标题总数：${r.headings.total}（一级 ${r.headings.byLevel[0] || 0} / 二级 ${r.headings.byLevel[1] || 0} / 三级 ${r.headings.byLevel[2] || 0} / 四级 ${r.headings.byLevel[3] || 0}）`)
  lines.push(`清洗总数：${r.clean.total} 处`)
  for (const it of cleanStatItems.value) lines.push(`  - ${it.label}：${it.value}`)
  lines.push(`模板：${r.format.template || '未应用'}`)
  lines.push(`格式化：标题 ${r.format.headings} / 标记 ${r.format.markers} / 表格 ${r.format.tables}`)
  lines.push('='.repeat(40))
  lines.push('生成时间：' + new Date().toLocaleString())
  const blob = new Blob([lines.join('\n')], { type: 'text/plain;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = (file.value?.name || 'report').replace(/\.[^.]+$/, '') + '-质检报告.txt'
  a.click()
  URL.revokeObjectURL(url)
  logger.info('⬇️ 已导出质检报告（.txt）')
}

const btnText = computed(() => (busy.value ? '处理中…' : '开始处理'))
</script>

<template>
  <div class="app">
    <!-- 顶部导航栏 -->
    <header class="topbar">
      <div class="brand">
        <div class="brand-text">
          <span class="name">Word 格式化小助手</span>
          <span class="sub">格式清洗 · 模板格式化 · 标题自动编号</span>
        </div>
      </div>
      <div class="topbar-right">
        <span class="top-stat" v-if="file">📄 {{ file.name.length > 14 ? file.name.slice(0, 14) + '…' : file.name }}</span>
        <span class="top-stat">清洗 {{ cleanCount }}/{{ CLEAN_ITEMS.length }}</span>
        <span class="top-stat" v-if="file && parsed">标题 {{ headingTotal }}</span>
      </div>
    </header>

    <div class="body">
      <!-- 左侧导航（大字号锚点菜单） -->
      <aside class="sidebar">
        <div class="sidebar-title">操作指引</div>
        <nav class="menu">
          <button v-for="s in SECTIONS" :key="s.id" class="menu-item" @click="scrollTo(s.id)">
            <span class="mi-icon">{{ s.icon }}</span>
            <span class="mi-text">{{ s.name }}</span>
            <span class="mi-no">{{ s.no }}</span>
          </button>
        </nav>
        <div class="sidebar-foot" v-if="file">
          <div class="file-mini">
            <div class="fm-name">{{ file.name }}</div>
            <div class="fm-size">{{ sizeText(file.size) }}</div>
          </div>
          <button class="btn ghost side-btn" @click="clearFile">更换文件</button>
        </div>
      </aside>

      <!-- 右侧主内容区（单页滚动） -->
      <main class="content">
        <!-- 区块 1：上传文档 -->
        <section class="section" id="sec-upload">
          <div class="sec-head">
            <span class="sec-no">1</span>
            <div class="sec-titles">
              <h2>📄 上传文档</h2>
              <p>选择要处理的 Word 文件（支持 .docx / .doc）</p>
            </div>
          </div>
          <UploadZone :file="file" :busy="busy" @select="onFileSelect" @clear="clearFile" />
          <div class="valid-result" v-if="file && parsed">
            <div class="vr-title">标题层级预检测结果</div>
            <div class="vr-stats">
              <div class="vr-item"><b>{{ headingTotal }}</b><span>标题总数</span></div>
              <div class="vr-item" v-for="(_, i) in [0, 1, 2, 3]" :key="i">
                <b>{{ countLevel(i + 1) }}</b><span>{{ ['一级', '二级', '三级', '四级'][i] }}</span>
              </div>
            </div>
            <HeadingTree :tree="tree" :show-numbers="true" />
          </div>
        </section>

        <!-- 区块 2：格式清洗 -->
        <section class="section" id="sec-clean">
          <div class="sec-head">
            <span class="sec-no">2</span>
            <div class="sec-titles">
              <h2>🧹 格式清洗</h2>
              <p>选择需要清理的项目，可用一键预设快速完成</p>
            </div>
          </div>

          <div class="presets">
            <button v-for="p in PRESETS" :key="p.id" class="preset" :class="{ sel: presetSel === p.id }" @click="applyPreset(p)">
              <span class="p-name">{{ p.name }}</span>
              <span class="p-desc">{{ p.desc }}</span>
            </button>
          </div>

          <div class="clean-grid">
            <div v-for="g in cleanGroups" :key="g.name" class="clean-group">
              <div class="cg-head">
                <span class="cg-title">{{ g.name }}</span>
                <span class="cg-count">{{ groupCount(g) }}/{{ g.items.length }}</span>
              </div>
              <div class="cg-items">
                <label v-for="item in g.items" :key="item.key" class="item" :title="item.desc">
                  <input type="checkbox" v-model="cleanOpts[item.key]" />
                  <span>{{ item.label }}</span>
                </label>
              </div>
            </div>
          </div>

          <div class="sub-opt">
            <span>隐藏文字处理：</span>
            <select v-model="cleanOpts.hiddenText">
              <option value="delete">删除隐藏文字</option>
              <option value="display">显示隐藏文字</option>
              <option value="keep">保留</option>
            </select>
          </div>

          <!-- 独立表格设置（不随模板切换，置于模板设置上部） -->
          <div class="opt-card">
            <h3>📊 表格设置</h3>
            <p class="tbl-hint">表格格式为独立设置，不随模板切换而改变，对两种模板统一生效</p>
            <div class="opt-grid">
              <label><span>表格边框</span><select v-model="borderSz" :disabled="!formatEnabled"><option value="8">1.0pt</option><option value="6">0.75pt</option></select></label>
              <label><span>表头背景</span><select v-model="headerFill" :disabled="!formatEnabled"><option value="E6E6E6">灰色 10%</option><option value="D9D9D9">灰色 15%</option></select></label>
              <label><span>单元格对齐</span><select v-model="cellAlign" :disabled="!formatEnabled"><option value="center">居中</option><option value="both">两端对齐</option></select></label>
              <label><span>表头加粗</span><select v-model="headerBold" :disabled="!formatEnabled"><option :value="true">加粗</option><option :value="false">不加粗</option></select></label>
              <label><span>表格整体对齐</span><select v-model="tblAlign" :disabled="!formatEnabled"><option value="left">居左</option><option value="center">居中</option></select></label>
              <label><span>列宽调整</span><select v-model="tblLayout" :disabled="!formatEnabled"><option value="autofit">自动适应</option><option value="fixed">固定宽度</option></select></label>
            </div>
            <p class="tbl-fixed">固定项：黑色实线边框 · 行高最小值 0.6cm</p>
          </div>

          <div class="opt-card">
            <h3>🎨 模板格式化配置</h3>
            <label class="switch-row"><span>启用模板格式化</span><input type="checkbox" v-model="formatEnabled" class="switch" /></label>
            <div class="tpl-grid" :class="{ off: !formatEnabled }">
              <div v-for="(t, id) in TEMPLATES" :key="id" class="tpl" :class="{ sel: templateId === id }" @click="formatEnabled && (templateId = id)">
                <div class="tpl-head">
                  <span class="radio" :class="{ on: templateId === id }"></span>
                  <span class="tpl-name">{{ t.name }}</span>
                  <span class="tpl-desc">{{ t.desc }}</span>
                </div>
                <details class="tpl-details" @click.stop>
                  <summary>查看模板参数</summary>
                  <ul><li v-for="(d, i) in templateDetails[id]" :key="i">{{ d }}</li></ul>
                </details>
              </div>
            </div>
            <div class="sub-opts">
              <label><input type="checkbox" v-model="numberEnabled" /><span>标题自动编号（删除旧编号，四层级，父子联动）</span></label>
              <label><input type="checkbox" v-model="textDetect" :disabled="!numberEnabled" /><span>按文本编号识别标题</span></label>
              <label><input type="checkbox" v-model="bulletsEnabled" :disabled="!numberEnabled" /><span>正文编号转实心圆点 ●（删除列举项编号，统一为 ●）</span></label>
            </div>
          </div>

          <div class="run-bar">
            <button class="btn primary run-btn" @click="run()" :disabled="busy || !file">
              <span v-if="busy" class="spin"></span><span v-else>🚀</span>{{ btnText }}
            </button>
            <p class="run-hint" v-if="!file">请先在上方上传文档</p>
          </div>
        </section>

        <!-- 区块 3：处理结果 -->
        <section class="section" id="sec-inspect">
          <div class="sec-head">
            <span class="sec-no">3</span>
            <div class="sec-titles">
              <h2>🔍 处理结果</h2>
              <p>结构检测与清洗检测数据一览</p>
            </div>
          </div>
          <div v-if="!report" class="no-data">
            <p>尚未执行处理，请在「格式清洗」区块点击「开始处理」。</p>
            <button class="btn primary" @click="goInspect()">立即处理</button>
          </div>
          <template v-else>
            <div class="kpi-row">
              <div class="kpi blue"><b>{{ report.headings.total }}</b><span>标题总数</span></div>
              <div class="kpi green"><b>{{ report.clean.total }}</b><span>清洗处数</span></div>
              <div class="kpi orange"><b>{{ report.format.tables }}</b><span>表格数</span></div>
              <div class="kpi purple"><b>{{ report.format.markers }}</b><span>特殊标记</span></div>
            </div>
            <div class="qa-grid">
              <div class="qa-card">
                <h3>📑 结构检测</h3>
                <ul class="qa-list">
                  <li><span>一级标题</span><b>{{ report.headings.byLevel[0] || 0 }}</b></li>
                  <li><span>二级标题</span><b>{{ report.headings.byLevel[1] || 0 }}</b></li>
                  <li><span>三级标题</span><b>{{ report.headings.byLevel[2] || 0 }}</b></li>
                  <li><span>四级标题</span><b>{{ report.headings.byLevel[3] || 0 }}</b></li>
                </ul>
              </div>
              <div class="qa-card">
                <h3>🧹 清洗检测</h3>
                <ul class="qa-list" v-if="cleanStatItems.length">
                  <li v-for="(it, i) in cleanStatItems" :key="i"><span>{{ it.label }}</span><b>{{ it.value }}</b></li>
                </ul>
                <p v-else class="qa-empty">未检测到需清洗内容</p>
              </div>
              <div class="qa-card">
                <h3>🎨 模板应用</h3>
                <ul class="qa-list">
                  <li><span>应用模板</span><b>{{ report.format.template || '未应用' }}</b></li>
                  <li><span>标题格式</span><b>{{ report.format.headings }}</b></li>
                  <li><span>表格</span><b>{{ report.format.tables }}</b></li>
                </ul>
              </div>
            </div>
            <HeadingTree :tree="tree" :show-numbers="true" />
          </template>
        </section>

        <!-- 区块 4：导出下载 -->
        <section class="section" id="sec-export">
          <div class="sec-head">
            <span class="sec-no">4</span>
            <div class="sec-titles">
              <h2>📤 导出下载</h2>
              <p>下载处理好的文档与质检报告</p>
            </div>
          </div>
          <div v-if="!result" class="no-data">
            <p>暂无可导出内容，请先完成处理。</p>
          </div>
          <template v-else>
            <div class="summary">
              <div class="sum-item"><span class="si-label">处理文件</span><span class="si-value">{{ file.name }}</span></div>
              <div class="sum-item"><span class="si-label">标题数量</span><span class="si-value">{{ report.headings.total }}</span></div>
              <div class="sum-item"><span class="si-label">清洗处理</span><span class="si-value">{{ report.clean.total }} 处</span></div>
              <div class="sum-item"><span class="si-label">应用模板</span><span class="si-value">{{ report.format.template }}</span></div>
              <div class="sum-item"><span class="si-label">输出文件</span><span class="si-value">{{ result.outName }}</span></div>
            </div>
            <div class="export-actions">
              <button class="btn success big" @click="download">⬇️ 导出 Word 文档</button>
              <button class="btn ghost big" @click="exportReport">📝 导出质检报告（.txt）</button>
            </div>
            <p class="tip">所有处理均在本地浏览器完成，文件不会上传到任何服务器</p>
          </template>
        </section>

        <!-- 日志面板 -->
        <div class="log-panel">
          <LogPanel />
        </div>
      </main>
    </div>
  </div>
</template>

<style scoped>
.app { min-height: 100vh; display: flex; flex-direction: column; }

/* ---------- 顶部导航栏 ---------- */
.topbar {
  height: 64px;
  background: linear-gradient(90deg, #1b6bb5, #2f8bd0);
  color: #fff;
  display: flex;
  align-items: center;
  padding: 0 24px;
  box-shadow: 0 2px 10px rgba(13, 67, 119, 0.25);
  flex: none;
  z-index: 10;
}
.brand { display: flex; align-items: center; gap: 12px; }
.brand-text { display: flex; flex-direction: column; line-height: 1.25; }
.name { font-size: 20px; font-weight: 700; }
.sub { font-size: 13px; opacity: 0.9; }
.topbar-right { margin-left: auto; display: flex; gap: 10px; flex-wrap: wrap; }
.top-stat {
  font-size: 14px;
  font-weight: 600;
  background: rgba(255, 255, 255, 0.18);
  border: 1px solid rgba(255, 255, 255, 0.4);
  border-radius: 999px;
  padding: 6px 16px;
}

/* ---------- 主体布局 ---------- */
.body { display: flex; flex: 1; min-height: 0; }

/* ---------- 左侧导航 ---------- */
.sidebar {
  width: 220px;
  flex: none;
  background: var(--sidebar-bg);
  color: #d7e3f0;
  display: flex;
  flex-direction: column;
  padding: 18px 12px;
}
.sidebar-title {
  font-size: 16px;
  color: #9db8d6;
  padding: 0 12px 14px;
  font-weight: 700;
  letter-spacing: 1px;
}
.menu { display: flex; flex-direction: column; gap: 8px; flex: 1; }
.menu-item {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 14px 14px;
  border: none;
  border-radius: 10px;
  background: transparent;
  color: #d7e3f0;
  font-size: 17px;
  font-weight: 600;
  cursor: pointer;
  text-align: left;
  transition: all 0.15s;
}
.menu-item:hover { background: rgba(255, 255, 255, 0.12); }
.menu-item:active { background: #1b6bb5; color: #fff; }
.mi-icon { font-size: 22px; }
.mi-text { flex: 1; }
.mi-no {
  font-size: 15px;
  font-weight: 700;
  width: 28px;
  height: 28px;
  border-radius: 50%;
  border: 1px solid #5a7699;
  display: flex;
  align-items: center;
  justify-content: center;
}
.sidebar-foot { padding-top: 14px; border-top: 1px solid rgba(255, 255, 255, 0.15); }
.file-mini { padding: 8px 12px; font-size: 15px; }
.fm-name { color: #e9f1fa; word-break: break-all; }
.fm-size { color: #9db8d6; margin-top: 3px; }
.side-btn { width: 100%; margin-top: 6px; }

/* ---------- 右侧主内容区 ---------- */
.content { flex: 1; padding: 20px 24px 40px; min-width: 0; overflow-y: auto; }

/* ---------- 区块 ---------- */
.section {
  background: var(--card-bg);
  border: 1px solid var(--border);
  border-radius: 10px;
  box-shadow: var(--shadow);
  padding: 24px;
  margin-bottom: 20px;
}
.sec-head {
  display: flex;
  align-items: center;
  gap: 14px;
  margin-bottom: 20px;
  padding-bottom: 14px;
  border-bottom: 2px solid var(--primary-100);
}
.sec-no {
  flex: none;
  width: 40px;
  height: 40px;
  border-radius: 10px;
  background: var(--primary);
  color: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 22px;
  font-weight: 700;
}
.sec-titles h2 { margin: 0 0 2px; font-size: var(--fs-xl); color: var(--text-strong); }
.sec-titles p { margin: 0; font-size: var(--fs-sm); color: var(--muted); }

/* ---------- 文件校验 ---------- */
.valid-result { margin-top: 20px; }
.vr-title { font-size: 19px; font-weight: 700; color: var(--primary-600); margin-bottom: 12px; }
.vr-stats { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
.vr-item {
  flex: 1;
  min-width: 100px;
  background: var(--primary-50);
  border-radius: 10px;
  padding: 16px;
  text-align: center;
}
.vr-item b { font-size: 30px; color: var(--primary-700); display: block; }
.vr-item span { font-size: 16px; color: var(--text); }

/* ---------- 清洗规则 ---------- */
.presets { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }
.preset {
  flex: 1;
  min-width: 130px;
  border: 2px solid var(--border);
  background: #fbfcfe;
  border-radius: 10px;
  padding: 14px 16px;
  cursor: pointer;
  text-align: left;
  transition: all 0.15s;
}
.preset:hover { border-color: var(--primary); }
.preset.sel { border-color: var(--primary); background: var(--primary-50); box-shadow: 0 0 0 3px var(--primary-100); }
.p-name { display: block; font-weight: 700; color: var(--text-strong); font-size: 18px; }
.p-desc { display: block; font-size: 15px; color: var(--muted); margin-top: 3px; }
.clean-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 16px; margin-bottom: 18px; }
.clean-group { border: 1px solid var(--border); border-radius: 10px; padding: 16px; }
.cg-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
.cg-title { font-size: 18px; font-weight: 700; color: var(--primary-700); }
.cg-count { font-size: 15px; color: var(--primary-600); background: var(--primary-50); border-radius: 999px; padding: 2px 12px; }
.cg-items { display: flex; flex-direction: column; gap: 6px; }
.item {
  display: flex;
  align-items: center;
  gap: 10px;
  font-size: 18px;
  padding: 9px 10px;
  border-radius: 8px;
  cursor: pointer;
}
.item:hover { background: var(--primary-50); }
.item input {
  width: 22px;
  height: 22px;
  accent-color: var(--primary);
  cursor: pointer;
}
.sub-opt { display: flex; align-items: center; gap: 12px; font-size: 18px; margin-bottom: 18px; }
.opt-card { border: 1px solid var(--border); border-radius: 10px; padding: 20px; margin-bottom: 18px; background: #fbfcfe; }
.opt-card h3 { margin: 0 0 16px; font-size: 20px; color: var(--text-strong); }
.switch-row { display: flex; align-items: center; justify-content: space-between; font-size: 18px; font-weight: 700; margin-bottom: 14px; }
.switch { appearance: none; width: 54px; height: 30px; border-radius: 16px; background: #c6d2e2; position: relative; cursor: pointer; flex: none; }
.switch:checked { background: var(--primary); }
.switch::after { content: ''; position: absolute; top: 3px; left: 3px; width: 24px; height: 24px; border-radius: 50%; background: #fff; transition: left 0.2s; }
.switch:checked::after { left: 27px; }
.tpl-grid { display: flex; flex-direction: column; gap: 12px; margin-bottom: 14px; }
.tpl-grid.off { opacity: 0.5; pointer-events: none; }
.tpl { border: 1px solid var(--border); border-radius: 10px; padding: 14px; cursor: pointer; }
.tpl.sel { border-color: var(--primary); background: var(--primary-50); }
.tpl-head { display: flex; align-items: center; gap: 10px; }
.radio { width: 20px; height: 20px; border-radius: 50%; border: 2px solid #a9b8cb; flex: none; }
.radio.on { border-color: var(--primary); background: radial-gradient(circle, var(--primary) 0 6px, #fff 7px); }
.tpl-name { font-size: 18px; font-weight: 700; }
.tpl-desc { font-size: 14px; color: var(--muted); margin-left: auto; }
.tpl-details { margin-top: 10px; font-size: 15px; color: #4a5668; }
.tpl-details summary { cursor: pointer; color: var(--primary); font-weight: 700; font-size: 16px; }
.tpl-details ul { margin: 8px 0 0; padding-left: 20px; line-height: 1.9; }
.opt-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 12px; margin-bottom: 14px; }
.opt-grid label { display: flex; flex-direction: column; gap: 6px; font-size: 16px; color: #3a4658; }
.tbl-hint { margin: 0 0 12px; font-size: 15px; color: var(--muted); }
.tbl-fixed { margin: 12px 0 0; font-size: 15px; color: var(--muted); padding-top: 12px; border-top: 1px dashed var(--border); }
.sub-opts { display: flex; flex-direction: column; gap: 10px; font-size: 17px; }
.sub-opts label { display: flex; gap: 10px; align-items: flex-start; cursor: pointer; }
.sub-opts input { width: 22px; height: 22px; accent-color: var(--primary); margin-top: 2px; }

/* ---------- 处理按钮条 ---------- */
.run-bar {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 10px;
  padding-top: 18px;
  border-top: 1px solid var(--border);
}
.run-btn {
  width: 100%;
  padding: 18px;
  font-size: 22px;
}
.run-hint { margin: 0; font-size: 16px; color: var(--muted); }

/* ---------- 质检 ---------- */
.kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 14px; margin-bottom: 18px; }
.kpi { border-radius: 10px; padding: 20px; text-align: center; color: #fff; }
.kpi b { display: block; font-size: 32px; }
.kpi span { font-size: 15px; opacity: 0.95; }
.kpi.blue { background: linear-gradient(135deg, #1b6bb5, #3d8fd0); }
.kpi.green { background: linear-gradient(135deg, #1e8f54, #46b377); }
.kpi.orange { background: linear-gradient(135deg, #cf7a1e, #eea243); }
.kpi.purple { background: linear-gradient(135deg, #6a4fb5, #8d76d8); }
.qa-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 16px; margin-bottom: 18px; }
.qa-card { border: 1px solid var(--border); border-radius: 10px; padding: 16px; }
.qa-card h3 { margin: 0 0 12px; font-size: 18px; color: var(--text-strong); }
.qa-list { list-style: none; margin: 0; padding: 0; }
.qa-list li { display: flex; justify-content: space-between; padding: 9px 0; border-bottom: 1px dashed var(--border); font-size: 17px; }
.qa-list li:last-child { border-bottom: none; }
.qa-list li span { color: var(--muted); }
.qa-list li b { color: var(--primary-700); font-size: 18px; }
.qa-empty { color: var(--muted); font-size: 16px; }
.no-data { text-align: center; padding: 48px 20px; color: var(--muted); }
.no-data p { margin-bottom: 20px; font-size: 18px; }

/* ---------- 导出 ---------- */
.summary { border: 1px solid var(--border); border-radius: 10px; margin-bottom: 24px; overflow: hidden; }
.sum-item { display: flex; justify-content: space-between; align-items: center; padding: 16px 20px; border-bottom: 1px solid var(--border); }
.sum-item:last-child { border-bottom: none; }
.si-label { color: var(--muted); font-size: 17px; }
.si-value { font-weight: 700; color: var(--text-strong); font-size: 17px; word-break: break-all; text-align: right; }
.export-actions { display: flex; gap: 14px; flex-wrap: wrap; justify-content: center; }
.tip { text-align: center; font-size: 15px; color: var(--muted); margin-top: 18px; }

/* ---------- 按钮 ---------- */
.btn {
  border: none;
  border-radius: 10px;
  padding: 14px 28px;
  font-size: 19px;
  font-weight: 700;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  transition: all 0.15s;
}
.btn:disabled { opacity: 0.5; cursor: not-allowed; }
.btn.primary { background: var(--primary); color: #fff; }
.btn.primary:hover:not(:disabled) { background: var(--primary-600); }
.btn.ghost { background: #fff; color: var(--muted); border: 2px solid var(--border); }
.btn.ghost:hover { color: var(--primary); border-color: var(--primary); }
.btn.success { background: var(--green); color: #fff; }
.btn.success:hover { background: var(--green-600); }
.btn.big { padding: 18px 36px; font-size: 20px; }
.spin {
  width: 20px;
  height: 20px;
  border: 3px solid rgba(255, 255, 255, 0.4);
  border-top-color: #fff;
  border-radius: 50%;
  animation: rot 0.8s linear infinite;
}
@keyframes rot { to { transform: rotate(360deg); } }

/* ---------- 日志 ---------- */
.log-panel { background: #fff; border: 1px solid var(--border); border-radius: 10px; box-shadow: var(--shadow); padding: 16px 20px; }
select {
  padding: 10px 14px;
  border: 2px solid #b9c8da;
  border-radius: 8px;
  background: #fff;
  font-size: 17px;
  color: var(--text);
  outline: none;
}
select:focus { border-color: var(--primary); box-shadow: 0 0 0 3px var(--primary-100); }
@media (max-width: 860px) {
  .sidebar { display: none; }
  .content { padding: 14px; }
}
</style>
