// 格式清洗：13 项功能，直接作用于 docx XML（保留原有格式与排版）

import { W, isW, child, getPText, directRuns, setRText, clearRText } from './xml.js'
import { logger } from './logger.js'

const RE_ZERO = /[\u200B\u200C\u200D\uFEFF\u2060\u180E]/g
const RE_SOFTHYPHEN = /\u00AD/g
const RE_CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F\u2028\u2029]/g
const SPACE_CHARS = ' \u3000\u00A0\t'
const isSpace = (c) => SPACE_CHARS.includes(c)

export function cleanDocument(doc, opts) {
  const stats = {
    zeroWidth: 0, softHyphen: 0, controlChars: 0,
    leadingSpace: 0, trailingSpace: 0, nbspToSpace: 0, tabToSpace: 0,
    clearHyperlinks: 0, clearColor: 0, clearHighlight: 0,
    hiddenDelete: 0, hiddenDisplay: 0, mergeSpaces: 0, emptyNumbered: 0
  }

  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  if (!body) return stats

  // ---- 1. 清除超链接（把 w:hyperlink 替换为其子内容） ----
  if (opts.clearHyperlinks) {
    const links = Array.from(body.getElementsByTagNameNS(W, 'hyperlink'))
    for (const link of links) {
      let count = 0
      while (link.firstChild) {
        link.parentNode.insertBefore(link.firstChild, link)
        count++
      }
      link.parentNode.removeChild(link)
      stats.clearHyperlinks++
      if (count === 0) stats.clearHyperlinks-- // 空链接不计
    }
  }

  // ---- 2. 收集所有段落（含表格/文本框内段落） ----
  const paras = Array.from(body.getElementsByTagNameNS(W, 'p'))

  for (const p of paras) {
    // 隐藏文字处理优先（整段级与 run 级）
    if (opts.hiddenText !== 'keep') processHidden(p, opts, stats)

    // 删除有编号但内容为空的段落（编号后无实际内容）
    // 需在 run 判断前执行，以覆盖无 run 的自动编号空段落
    if (opts.emptyNumbered && isEmptyNumberedParagraph(p)) {
      // 跳过含 sectPr 的段落（避免破坏页面设置）
      const pPrEl = child(p, 'pPr')
      if (!child(p, 'sectPr') && !(pPrEl && child(pPrEl, 'sectPr'))) {
        p.parentNode.removeChild(p)
        stats.emptyNumbered++
        continue
      }
    }

    const runs = directRuns(p)
    if (!runs.length) continue

    // ---- 3. 提取 run 文本序列（跨 run 处理行首/行尾/合并空格） ----
    const seq = [] // { run, ch }
    for (const r of runs) {
      // 处理 w:t 文本 + w:tab 元素
      for (const c of r.childNodes) {
        if (c.nodeType !== 1) continue
        if (isW(c, 't')) {
          const text = c.textContent || ''
          for (const ch of text) seq.push({ run: r, ch })
        } else if (isW(c, 'tab')) {
          // 制表符：转空格时并入文本序列并移除元素；否则保留元素
          if (opts.tabToSpace) {
            r.removeChild(c)
            seq.push({ run: r, ch: '\t' })
          }
        }
        // w:br 保留（换行）
      }
    }

    if (!seq.length) continue

    // 字符级过滤
    let i = 0
    while (i < seq.length) {
      const it = seq[i]
      let keep = true
      if (RE_ZERO.test(it.ch)) { RE_ZERO.lastIndex = 0; stats.zeroWidth++; keep = false }
      else if (RE_SOFTHYPHEN.test(it.ch)) { RE_SOFTHYPHEN.lastIndex = 0; stats.softHyphen++; keep = false }
      else if (RE_CONTROL.test(it.ch)) { RE_CONTROL.lastIndex = 0; stats.controlChars++; keep = false }
      else if (it.ch === '\u00A0') { it.ch = ' '; stats.nbspToSpace++ }
      else if (it.ch === '\t') { it.ch = ' '; stats.tabToSpace++ }
      if (!keep) seq.splice(i, 1)
      else i++
    }

    if (!seq.length) { runs.forEach((r) => clearRText(r)); continue }

    // 行首空格
    if (opts.leadingSpace) {
      while (seq.length && isSpace(seq[0].ch)) { seq.shift(); stats.leadingSpace++ }
    }
    // 行尾空格
    if (opts.trailingSpace) {
      while (seq.length && isSpace(seq[seq.length - 1].ch)) { seq.pop(); stats.trailingSpace++ }
    }
    // 合并多余空格
    if (opts.mergeSpaces) {
      const out = []
      for (const it of seq) {
        const prev = out[out.length - 1]
        if (isSpace(it.ch) && prev && isSpace(prev.ch)) {
          stats.mergeSpaces++
          continue
        }
        out.push(it)
      }
      seq.length = 0
      seq.push(...out)
    }

    // ---- 4. 回写：按 run 重组文本 ----
    const byRun = new Map()
    for (const it of seq) {
      if (!byRun.has(it.run)) byRun.set(it.run, '')
      byRun.set(it.run, byRun.get(it.run) + it.ch)
    }
    for (const [r, text] of byRun) setRText(r, text)
    for (const r of runs) {
      if (!byRun.has(r)) clearRText(r)
    }

    // 手动颜色 / 高亮（rPr 级）
    if (opts.clearColor || opts.clearHighlight) {
      for (const r of runs) {
        const rPr = child(r, 'rPr')
        if (!rPr) continue
        if (opts.clearColor) {
          const c = child(rPr, 'color')
          if (c) { rPr.removeChild(c); stats.clearColor++ }
        }
        if (opts.clearHighlight) {
          const h = child(rPr, 'highlight')
          if (h) { rPr.removeChild(h); stats.clearHighlight++ }
          const shd = child(rPr, 'shd')
          if (shd) { rPr.removeChild(shd); stats.clearHighlight++ }
        }
      }
    }
  }

  logStats(stats)
  return stats
}

function processHidden(p, opts, stats) {
  // run 级隐藏文字
  const runs = Array.from(p.getElementsByTagNameNS(W, 'r'))
  for (const r of runs) {
    const rPr = child(r, 'rPr')
    const vanish = rPr && child(rPr, 'vanish')
    if (!vanish) continue
    if (opts.hiddenText === 'delete') {
      const len = (getPText(r) || '').length
      stats.hiddenDelete += len || 1
      r.parentNode.removeChild(r)
    } else if (opts.hiddenText === 'display') {
      rPr.removeChild(vanish)
      stats.hiddenDisplay++
    }
  }
  // 段落标记隐藏
  const pPr = child(p, 'pPr')
  const rPr = pPr && child(pPr, 'rPr')
  const pvanish = rPr && child(rPr, 'vanish')
  if (pvanish) {
    if (opts.hiddenText === 'delete') { rPr.removeChild(pvanish) }
    else if (opts.hiddenText === 'display') { rPr.removeChild(pvanish); stats.hiddenDisplay++ }
  }
}

// 判断段落是否为"有编号但内容为空"：
// 1) 自动编号（w:numPr + numId>0）且无文本；
// 2) 文本仅由编号组成（如 一、/（一）/1./（1））
const NUMBER_ONLY_RE = /^(第[一二三四五六七八九十百千零〇]+[章节篇]|[一二三四五六七八九十百千零〇]+[、.．,，]|（[一二三四五六七八九十百千零〇]+）|\([一二三四五六七八九十百千零〇]+\)|（\d{1,4}）|\(\d{1,4}\)|\d{1,4}[.．、])[\s\u3000\u00A0]*$/
function isEmptyNumberedParagraph(p) {
  const text = getPText(p).trim()
  // 自动编号：段落带 numPr，且没有可见文本
  const pPrEl = child(p, 'pPr')
  const numPr = pPrEl && child(pPrEl, 'numPr')
  if (numPr) {
    const numId = child(numPr, 'numId')
    const val = numId && numId.getAttribute('w:val')
    if (val && val !== '0') {
      return text === ''
    }
  }
  // 文本编号：仅由编号组成且无实际内容
  if (text) {
    return NUMBER_ONLY_RE.test(text)
  }
  return false
}

function logStats(stats) {
  const map = {
    zeroWidth: '零宽字符', softHyphen: '软连字符', controlChars: '不可见控制字符',
    leadingSpace: '行首空格', trailingSpace: '行尾空格', nbspToSpace: '不间断空格',
    tabToSpace: '制表符', clearHyperlinks: '超链接', clearColor: '手动颜色',
    clearHighlight: '手动高亮', hiddenDelete: '隐藏文字(删除)', hiddenDisplay: '隐藏文字(显示)',
    mergeSpaces: '多余空格', emptyNumbered: '空编号行'
  }
  let total = 0
  const parts = []
  for (const [k, v] of Object.entries(stats)) {
    if (v > 0) { parts.push(`${map[k]} ${v} 处`); total += v }
  }
  if (parts.length) logger.info(`🧹 格式清洗完成：${parts.join('，')}`, { kind: 'clean', stats })
  else logger.info('🧹 格式清洗完成：未发现需要处理的内容')
}
