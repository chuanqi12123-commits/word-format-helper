// 模板格式化：页面设置、页码、段落/标题/特殊标记、表格

import { W, child, ensureChild, removeChild, wEl, ppr, rpr, bodyParas, directRuns, getRText, getPText, setRText, findSectPr } from './xml.js'
import { addPageFooter } from './docxIO.js'
import { logger } from './logger.js'

export async function applyTemplate(doc, zip, template, opts = {}) {
  opts = {
    markerColor: 'red',   // red | black
    borderSz: 8,          // 1.0pt = 8, 0.75pt = 6
    headerFill: template.table.headerFill,
    cellAlign: 'center',  // center | both
    headerBold: true,     // 表头是否加粗
    tblAlign: 'left',     // 表格整体对齐：left | center
    tblLayout: 'autofit', // 列宽：autofit 自动 | fixed 固定
    ...opts
  }

  // 1. 页面设置
  applyPageSetup(doc, template)

  // 2. 页码（底部居中）
  if (template.footer) {
    const rId = await addPageFooter(zip, doc)
    if (rId) addFooterReference(doc, rId)
  }

  // 3. 段落格式（标题 + 正文 + 特殊标记）
  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  const paras = bodyParas(body)
  let headingCount = 0
  let markerCount = 0
  for (const p of paras) {
    const role = opts.levels ? opts.levels.get(p) || 0 : 0
    formatParagraph(p, template, opts, role)
    if (role > 0) headingCount++
    if (role === 0 && opts.formatMarkers !== false) {
      if (ensureMarkerFormat(p, template, opts)) markerCount++
    }
  }

  // 4. 表格
  const tableCount = formatTables(doc, template, opts)

  logger.success(`🎨 模板格式化完成（${template.name}）：标题 ${headingCount} 个，特殊标记 ▲/★ ${markerCount} 处，表格 ${tableCount} 张`, { kind: 'format' })

  return { headingCount, markerCount, tableCount }
}

// ---------- 页面设置 ----------
function applyPageSetup(doc, template) {
  const sectPr = findSectPr(doc)
  if (!sectPr) return
  const pgSz = ensureChild(sectPr, 'pgSz')
  pgSz.setAttribute('w:w', template.page.size.w)
  pgSz.setAttribute('w:h', template.page.size.h)
  pgSz.setAttribute('w:orient', 'portrait')
  const pgMar = ensureChild(sectPr, 'pgMar')
  const m = template.page.margins
  pgMar.setAttribute('w:top', m.top)
  pgMar.setAttribute('w:bottom', m.bottom)
  pgMar.setAttribute('w:left', m.left)
  pgMar.setAttribute('w:right', m.right)
  pgMar.setAttribute('w:gutter', m.gutter)
}

function addFooterReference(doc, rId) {
  const sectPr = findSectPr(doc)
  if (!sectPr) return
  if (child(sectPr, 'footerReference')) return
  const fr = wEl('w:footerReference', { 'w:type': 'default', 'r:id': rId })
  sectPr.insertBefore(fr, sectPr.firstChild)
}

// ---------- 段落 ----------
function formatParagraph(p, template, opts, role) {
  const pPrEl = ppr(p)
  const isHeading = role > 0
  const h = isHeading ? template.headings[role - 1] : null

  // 对齐
  const jc = ensureChild(pPrEl, 'jc')
  jc.setAttribute('w:val', isHeading ? h.align : template.body.align)

  // 间距
  const sp = ensureChild(pPrEl, 'spacing')
  if (isHeading) {
    if (h.beforeLines != null) {
      sp.setAttribute('w:beforeLines', h.beforeLines)
      sp.setAttribute('w:afterLines', h.afterLines)
    }
    if (h.before != null) {
      sp.setAttribute('w:before', h.before)
      sp.setAttribute('w:after', h.after)
    }
    sp.setAttribute('w:line', template.line)
    sp.setAttribute('w:lineRule', 'auto')
  } else {
    sp.setAttribute('w:before', '0')
    sp.setAttribute('w:after', '0')
    sp.setAttribute('w:beforeLines', '0')
    sp.setAttribute('w:afterLines', '0')
    if (template.id === 'special') {
      sp.setAttribute('w:line', template.line)
      sp.setAttribute('w:lineRule', 'auto')
    } else {
      sp.removeAttribute('w:line')
      sp.removeAttribute('w:lineRule')
    }
  }

  // 缩进：正文首行缩进 2 字符；标题无缩进
  const ind = child(pPrEl, 'ind')
  if (isHeading) {
    if (ind) pPrEl.removeChild(ind)
  } else {
    let indEl = ind
    if (!indEl) { indEl = wEl('w:ind'); pPrEl.appendChild(indEl) }
    indEl.setAttribute('w:firstLineChars', template.body.indentChars)
    indEl.setAttribute('w:firstLine', template.body.indentTwip)
    indEl.removeAttribute('w:left')
    indEl.removeAttribute('w:right')
    indEl.removeAttribute('w:hanging')
  }

  // 与下段同页（标题）
  if (isHeading) ensureChild(pPrEl, 'keepNext')
  else removeChild(pPrEl, 'keepNext')

  // 大纲级别
  if (isHeading) {
    let ol = child(pPrEl, 'outlineLvl')
    if (!ol) { ol = wEl('w:outlineLvl'); pPrEl.appendChild(ol) }
    ol.setAttribute('w:val', 'level-' + role)
  } else {
    removeChild(pPrEl, 'outlineLvl')
  }

  // 段落标记 rPr（字体）
  const pRPr = ensureChild(pPrEl, 'rPr')
  applyFontsToRPr(pRPr, template, isHeading ? h.size : template.body.size, isHeading)
  if (isHeading) ensureChild(pRPr, 'b')
  else removeChild(pRPr, 'b')

  // 段落内所有 run：字体 + 字号 + 加粗 + 重置颜色
  const runs = directRuns(p)
  const size = isHeading ? h.size : template.body.size
  for (const r of runs) {
    applyRunFonts(r, template, size, isHeading)
  }
}

// 给 rPr 写入字体信息
function applyFontsToRPr(rPrEl, template, size, bold) {
  const rf = ensureChild(rPrEl, 'rFonts')
  rf.setAttribute('w:ascii', template.fonts.ascii)
  rf.setAttribute('w:hAnsi', template.fonts.hAnsi)
  rf.setAttribute('w:eastAsia', template.fonts.eastAsia)
  rf.setAttribute('w:cs', template.fonts.ascii)
  const sz = ensureChild(rPrEl, 'sz')
  sz.setAttribute('w:val', size)
  const szCs = ensureChild(rPrEl, 'szCs')
  szCs.setAttribute('w:val', size)
  const color = child(rPrEl, 'color')
  if (color) rPrEl.removeChild(color)
}

function applyRunFonts(r, template, size, bold) {
  applyFontsToRPr(rpr(r), template, size, bold)
  if (bold) ensureChild(rpr(r), 'b')
  else removeChild(rpr(r), 'b')
}

// ---------- 特殊标记 ▲ / ★ ----------
function ensureMarkerFormat(p, template, opts) {
  const runs = directRuns(p)
  if (!runs.length) return false
  const fullText = getPText(p)
  const trimmed = fullText.trimStart()
  // 仅当 ▲/★ 位于条款正文最前方时才处理（避免误伤句中符号）
  if (!/^[▲★]/.test(trimmed)) return false
  const off = fullText.length - trimmed.length // 跳过行首空格
  const markerCh = trimmed[0]

  // 定位包含该字符的 run
  let acc = 0, run = null, localOff = -1
  for (const r of runs) {
    const t = getRText(r)
    if (acc + t.length > off) { run = r; localOff = off - acc; break }
    acc += t.length
  }
  if (!run) return false
  const text = getRText(run)
  const before = text.slice(0, localOff)
  const after = text.slice(localOff + 1)

  // 标记 run（红色或黑色加粗）
  const mRun = wEl('w:r')
  setRText(mRun, markerCh)
  const mRPr = rpr(mRun)
  ensureChild(mRPr, 'b')
  const color = wEl('w:color', { 'w:val': opts.markerColor === 'black' ? '000000' : template.marker.color })
  mRPr.appendChild(color)

  // 拆分为 前置文本 + 标记 + 后置文本（前置空白不保留，保证标记置于最前方）
  const beforeClean = before.trim()
  const bRun = beforeClean ? (() => { const r = wEl('w:r'); setRText(r, beforeClean); return r })() : null
  const aRun = wEl('w:r')
  setRText(aRun, after)

  // 在原 run 位置替换为 bRun? + mRun + aRun
  const parent = run.parentNode
  parent.replaceChild(mRun, run)
  if (bRun) parent.insertBefore(bRun, mRun)
  parent.insertBefore(aRun, mRun.nextSibling)

  // 标记置于条款正文最前方（若已在前则跳过）
  const pPrEl = child(p, 'pPr')
  const insertPos = pPrEl ? pPrEl.nextSibling : p.firstChild
  if (insertPos !== mRun) p.insertBefore(mRun, insertPos)
  return true
}

// ---------- 表格 ----------
function formatTables(doc, template, opts) {
  const tbls = doc.getElementsByTagNameNS(W, 'tbl')
  for (const tbl of tbls) {
    const tblPr = child(tbl, 'tblPr')
    if (!tblPr) continue

    // 表格整体对齐（居左 / 居中）
    if (opts.tblAlign) {
      const tjc = ensureChild(tblPr, 'jc')
      tjc.setAttribute('w:val', opts.tblAlign)
    }

    // 列宽布局：autofit 自动适应内容 / fixed 固定
    if (opts.tblLayout) {
      let layout = child(tblPr, 'tblLayout')
      if (!layout) { layout = wEl('w:tblLayout'); tblPr.appendChild(layout) }
      layout.setAttribute('w:type', opts.tblLayout === 'fixed' ? 'fixed' : 'autofit')
    }

    // 黑色实线边框
    let borders = child(tblPr, 'tblBorders')
    if (!borders) { borders = wEl('w:tblBorders'); tblPr.appendChild(borders) }
    for (const edge of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
      let e = child(borders, edge)
      if (!e) { e = wEl('w:' + edge); borders.appendChild(e) }
      e.setAttribute('w:val', 'single')
      e.setAttribute('w:sz', opts.borderSz)
      e.setAttribute('w:space', '0')
      e.setAttribute('w:color', '000000')
    }

    const rows = Array.from(tbl.getElementsByTagNameNS(W, 'tr'))
    rows.forEach((tr, ri) => {
      // 行高最小值
      let trPr = child(tr, 'trPr')
      if (!trPr) { trPr = wEl('w:trPr'); tr.insertBefore(trPr, tr.firstChild) }
      let th = child(trPr, 'trHeight')
      if (!th) { th = wEl('w:trHeight'); trPr.appendChild(th) }
      th.setAttribute('w:val', template.table.rowMinHeight)
      th.setAttribute('w:hRule', 'atLeast')

      const cells = Array.from(tr.getElementsByTagNameNS(W, 'tc'))
      cells.forEach((tc) => {
        let tcPr = child(tc, 'tcPr')
        if (!tcPr) { tcPr = wEl('w:tcPr'); tc.insertBefore(tcPr, tc.firstChild) }
        const jc = ensureChild(tcPr, 'jc')
        if (ri === 0) {
          // 表头：灰色背景填充 + 居中
          jc.setAttribute('w:val', 'center')
          let shd = child(tcPr, 'shd')
          if (!shd) { shd = wEl('w:shd'); tcPr.appendChild(shd) }
          shd.setAttribute('w:val', 'clear')
          shd.setAttribute('w:color', 'auto')
          shd.setAttribute('w:fill', opts.headerFill)
        } else {
          jc.setAttribute('w:val', opts.cellAlign)
          removeChild(tcPr, 'shd')
        }
        // 单元格段落
        const cps = Array.from(tc.getElementsByTagNameNS(W, 'p'))
        for (const cp of cps) {
          for (const r of directRuns(cp)) applyRunFonts(r, template, template.body.size, false)
          const cpPr = ppr(cp)
          removeChild(cpPr, 'ind')
          const cjc = ensureChild(cpPr, 'jc')
          cjc.setAttribute('w:val', ri === 0 ? 'center' : opts.cellAlign)
          if (ri === 0 && opts.headerBold !== false) {
            const cRPr = ensureChild(cpPr, 'rPr')
            ensureChild(cRPr, 'b')
          }
        }
      })

      // 表头跨页重复
      if (ri === 0) {
        const trPr2 = child(tr, 'trPr')
        if (trPr2) {
          if (!child(trPr2, 'tblHeader')) {
            const hd = wEl('w:tblHeader')
            trPr2.appendChild(hd)
          }
        }
      }
    })
  }
  return tbls.length
}
