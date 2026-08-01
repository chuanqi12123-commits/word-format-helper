// 标题编号：检测标题层级 → 清除旧编号 → 按层级自动编号（父子依赖联动）
// 编号规则：一级 一、二、三…；二级 （一）（二）…；三级 1. 2. …；四级 （1）（2）…

import { W, isW, child, bodyParas, directRuns, getPText, wEl, textEl, ppr } from './xml.js'
import { logger } from './logger.js'

const CN_DIGITS = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']

// 阿拉伯数字 → 中文数字
export function toCn(n) {
  if (n <= 0) return '零'
  if (n < 10) return CN_DIGITS[n]
  if (n < 20) return '十' + (n % 10 ? CN_DIGITS[n % 10] : '')
  if (n < 100) {
    const tens = Math.floor(n / 10)
    const rest = n % 10
    return CN_DIGITS[tens] + '十' + (rest ? CN_DIGITS[rest] : '')
  }
  if (n < 1000) {
    const h = Math.floor(n / 100)
    const rest = n % 100
    let s = CN_DIGITS[h] + '百'
    if (rest) {
      if (rest < 10) s += '零' + CN_DIGITS[rest]
      else s += toCn(rest)
    }
    return s
  }
  return String(n)
}

// 各级别旧编号的正则（用于清除）
const STRIP_PATTERNS = [
  /^第[一二三四五六七八九十百零〇]+[章节篇][、.．,，]?\s*/,
  /^[一二三四五六七八九十百千零〇]+[、.．,，]\s*/,
  /^（[一二三四五六七八九十百千零〇]+）\s*/,
  /^\([一二三四五六七八九十百千零〇]+\)\s*/,
  /^\(\d{1,4}\)\s*/,
  /^（\d{1,4}）\s*/,
  // 连续多级阿拉伯编号，如 1.1、1.1.1、2.1.1（需整体删除，避免残留）
  /^\d{1,4}([.．]\d{1,4})+[.．、]?\s*/,
  // 单级阿拉伯编号，如 1. / 2.
  /^\d{1,4}[.．、]\s*/
]

// 文本编号检测 → 级别（用于无样式文档的自动识别）。
// 规则：只把符合章节标题格式的文本编号识别为标题；
//       数字编号（N. / N.N / N.N.N / （1）等）多属正文列举项，不识别为标题
//       （它们会由"编号转圆点"步骤统一改为 ● 项目符号）。
//   一、/第一章            → 1 级
//   （一）/(一)             → 2 级
//   其余（数字编号、列举项）→ 0 不识别
const LEVEL_RE = [
  { level: 1, re: /^第[一二三四五六七八九十百零〇]+[章节篇]|^[一二三四五六七八九十百千零〇]+[、]/ },
  { level: 2, re: /^（[一二三四五六七八九十百千零〇]+）|^\([一二三四五六七八九十百千零〇]+\)/ }
]

// 正文列举项开头的编号前缀（用于"编号转圆点"），需覆盖常见格式：
//   3.1.1 / 7.1 / 2. / （1） / (1) / 一、 等
const BODY_NUM_RE = /^(\（\d{1,4}\）|\（[一二三四五六七八九十百千零〇]+\）|\(\d{1,4}\)|第[一二三四五六七八九十百零〇]+[章节篇][、.．,，]?|[一二三四五六七八九十百千零〇]+[、.．,，]|\d{1,4}([.．]\d{1,4})+[.．、]?|\d{1,4}[.．、])\s*/

function textLevel(text) {
  if (!text) return 0
  const t = text.trim()
  if (!t || t.length > 50) return 0
  if (/[。；;！？!?，,]$/.test(t)) return 0
  for (const p of LEVEL_RE) {
    if (p.re.test(t)) return p.level
  }
  return 0
}

const HEADING_NAME_RE = /^(heading|标题|一级标题|二级标题|三级标题|四级标题|heading\s*\d|标题\s*\d)/i

// 解析样式表：styleId → { name, outlineLvl }
export function buildStyleMap(stylesDoc) {
  const map = new Map()
  if (!stylesDoc) return map
  const styles = stylesDoc.getElementsByTagNameNS(W, 'style')
  for (const s of styles) {
    const id = s.getAttribute('w:styleId')
    if (!id) continue
    const nameEl = child(s, 'name')
    const name = nameEl ? nameEl.getAttribute('w:val') : ''
    let outlineLvl = null
    const ol = child(s, 'outlineLvl')
    if (ol) outlineLvl = parseInt(ol.getAttribute('w:val') || '-1', 10)
    map.set(id, { name: name || '', outlineLvl })
  }
  return map
}

function levelFromStyle(pStyleId, stylesMap) {
  const st = stylesMap.get(pStyleId)
  if (!st) return 0
  if (st.outlineLvl != null && st.outlineLvl >= 0 && st.outlineLvl <= 8) return st.outlineLvl + 1
  const m = st.name.match(/(\d)/)
  if (m && m[1] >= 1 && m[1] <= 4) return Number(m[1])
  if (/一级|标题\s*一|标题\s*1|^h1$/i.test(st.name)) return 1
  if (/二级|标题\s*二|标题\s*2/i.test(st.name)) return 2
  if (/三级|标题\s*三|标题\s*3/i.test(st.name)) return 3
  if (/四级|标题\s*四|标题\s*4/i.test(st.name)) return 4
  return 0
}

// 清除段落开头的旧编号（文本编号 + 自动编号）
function stripOldNumber(p) {
  // 1) 移除自动编号（numPr）：Word 会据此在行首渲染 一、/1. 等，无法用文本正则删除
  let removed = 0
  const pPr = child(p, 'pPr')
  const numPr = pPr && child(pPr, 'numPr')
  if (numPr) {
    pPr.removeChild(numPr)
    removed++
  }

  // 2) 删除文本中直接书写的旧编号。
  //    注意：编号可能被 Word 拆到多个 run（如 "一" 和 "、总则" 分属不同 run），
  //    因此必须基于整段合并文本匹配，再按字符位置从前到后删除到各 run。
  const runs = directRuns(p)
  if (!runs.length) return removed

  // 收集所有 w:t 文本（按文档顺序）
  const texts = []
  for (const r of runs) {
    for (const c of Array.from(r.childNodes)) {
      if (isW(c, 't')) texts.push(c)
    }
  }
  if (!texts.length) return removed

  const full = texts.map((t) => t.textContent || '').join('')
  // 找要删除的前缀长度
  let stripLen = 0
  for (const re of STRIP_PATTERNS) {
    re.lastIndex = 0
    const m = full.match(re)
    if (m) {
      stripLen = m[0].length
      break
    }
  }
  if (stripLen <= 0) return removed

  // 按字符顺序，从第一个 w:t 开始依次删除 stripLen 个字符
  let remaining = stripLen
  for (let i = 0; i < texts.length && remaining > 0; i++) {
    const t = texts[i]
    const s = t.textContent || ''
    if (!s) continue
    if (s.length <= remaining) {
      // 整个 w:t 都被删掉（若变空且没有其他内容，可保留空文本；交由后续逻辑处理）
      t.textContent = ''
      remaining -= s.length
    } else {
      t.textContent = s.slice(remaining)
      remaining = 0
    }
  }
  return removed + stripLen
}

// 将正文中带编号的列举段落统一转为实心圆点 ●：
// - 仅处理未被识别为标题的段落（不在 levels 中）
// - 删除开头的编号前缀，在段首插入 "● "
// 返回转换的段落数
export function convertBodyNumbersToBullets(doc, levels, opts = {}) {
  if (!opts.enabled) return 0
  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  if (!body) return 0
  const paras = bodyParas(body)
  let count = 0
  for (const p of paras) {
    // 标题段落不处理
    if (levels && levels.has(p)) continue
    const text = (getPText(p) || '').trim()
    if (!text) continue
    if (!BODY_NUM_RE.test(text)) continue

    // 删除开头编号前缀
    const runs = directRuns(p)
    const texts = []
    for (const r of runs) {
      for (const c of Array.from(r.childNodes)) {
        if (isW(c, 't')) texts.push(c)
      }
    }
    if (!texts.length) continue
    const full = texts.map((t) => t.textContent || '').join('')
    const m = BODY_NUM_RE.exec(full)
    if (!m) continue
    const stripLen = m[0].length
    // 从第一个 w:t 依次删除 stripLen 个字符
    let remaining = stripLen
    for (let i = 0; i < texts.length && remaining > 0; i++) {
      const t = texts[i]
      const s = t.textContent || ''
      if (!s) continue
      if (s.length <= remaining) {
        t.textContent = ''
        remaining -= s.length
      } else {
        t.textContent = s.slice(remaining)
        remaining = 0
      }
    }

    // 在段首插入 "● " 文本 run
    const bulletRun = wEl('w:r')
    const bText = textEl('● ')
    bulletRun.appendChild(bText)
    const pPr = child(p, 'pPr')
    if (pPr && pPr.nextSibling) p.insertBefore(bulletRun, pPr.nextSibling)
    else p.insertBefore(bulletRun, p.firstChild)
    count++
  }
  return count
}

// 主流程：检测 + 编号，返回 { levels: Map(段落→级别), tree }
export function processHeadings(doc, stylesMap, opts = {}) {
  const enableNumbering = opts.enableNumbering !== false
  const textDetect = opts.textDetect !== false
  const stripOld = opts.stripOld !== false

  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  if (!body) return { levels: new Map(), tree: [] }

  const paras = bodyParas(body)
  const levels = new Map()
  const tree = []
  const stack = [] // 当前标题栈

  // 第一遍：检测层级
  // 优先级：标题样式 > 文本编号(textDetect) > outlineLvl
  // 说明：Word 中 outlineLvl 常被误设为 0，而文本编号更能反映真实层级，故置于其前。
  for (const p of paras) {
    let level = 0
    const pPr = child(p, 'pPr')
    const pStyle = pPr && child(pPr, 'pStyle')
    if (pStyle) {
      const id = pStyle.getAttribute('w:val')
      if (HEADING_NAME_RE.test(stylesMap.get(id)?.name || '')) {
        level = levelFromStyle(id, stylesMap)
      }
    }
    // 无样式层级时，优先用文本编号识别
    if (!level && textDetect) {
      level = textLevel(getPText(p))
    }
    // 仍无法识别时，才回退到 outlineLvl
    if (!level) {
      const ol = pPr && child(pPr, 'outlineLvl')
      if (ol) {
        const v = parseInt(ol.getAttribute('w:val') || '-1', 10)
        if (v >= 0 && v <= 3) level = v + 1
      }
    }
    if (level) levels.set(p, level)
  }

  if (!levels.size) {
    logger.info('未检测到标题（建议文档使用“标题1~4”样式，或开启“按文本编号识别”）')
    return { levels, tree }
  }

  // 第二遍：清除旧编号 + 自动编号（含依赖联动）
  const counters = [0, 0, 0, 0]
  let removedCount = 0
  const numbered = []

  for (const p of paras) {
    const level = levels.get(p)
    if (!level) continue

    // 清除旧编号
    if (stripOld) removedCount += stripOldNumber(p)

    // 更新计数器：本级 +1，低级重置（依赖关系：一级变化 → 二级三级四级随之更新）
    counters[level - 1]++
    for (let i = level; i < 4; i++) counters[i] = 0

    const number = makeNumber(counters, level)
    // 取纯文本应在插入编号之前（否则 getPText 会包含新编号）
    const text = (getPText(p) || '').trim()
    if (enableNumbering) {
      insertNumberRun(p, number)
    }
    // 标题树（父子依赖结构）
    while (stack.length && stack[stack.length - 1].level >= level) stack.pop()
    const node = {
      level,
      number: enableNumbering ? number : '',
      title: text,
      children: []
    }
    if (stack.length) stack[stack.length - 1].children.push(node)
    else tree.push(node)
    stack.push(node)
    numbered.push(node)
  }

  if (enableNumbering) {
    logger.success(`📝 标题编号完成：共 ${numbered.length} 个标题，清除旧编号 ${removedCount} 处，编号已按层级自动联动（一级变化，下级自动重排）`, { kind: 'number', count: numbered.length })
  } else {
    logger.info(`📝 标题检测完成：共 ${numbered.length} 个标题（未启用自动编号），清除旧编号 ${removedCount} 处`)
  }

  return { levels, tree, counted: numbered.length }
}

function makeNumber(counters, level) {
  const n = counters[level - 1]
  switch (level) {
    case 1: return toCn(n) + '、'
    case 2: return '（' + toCn(n) + '）'
    case 3: return n + '.'
    case 4: return '（' + n + '）'
    default: return ''
  }
}

// 在段落最前插入编号 run（紧随 pPr 之后）
function insertNumberRun(p, number) {
  const run = wEl('w:r')
  run.appendChild(textEl(number))
  const pPrEl = child(p, 'pPr')
  const anchor = pPrEl ? pPrEl.nextSibling : p.firstChild
  p.insertBefore(run, anchor)
  return run
}

// 记录标题 outlineLvl（供 Word 导航/大纲使用）
export function applyOutlineLevels(doc, levels) {
  for (const [p, level] of levels) {
    const pPrEl = ppr(p)
    let ol = child(pPrEl, 'outlineLvl')
    if (!ol) {
      ol = wEl('w:outlineLvl')
      pPrEl.appendChild(ol)
    }
    ol.setAttribute('w:val', 'level-' + level)
  }
}
