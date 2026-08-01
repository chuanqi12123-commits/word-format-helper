// 处理流水线：读取 → 清洗 → 标题编号 → 模板格式化 → 输出

import { openZip, readDocumentXml, readStylesXml, writeXmlPart, buildBlob } from './docxIO.js'
import { readDocText, buildDocxZipFromParagraphs } from './docIO.js'
import { cleanDocument } from './cleaner.js'
import { processHeadings, buildStyleMap, applyOutlineLevels, convertBodyNumbersToBullets } from './numbering.js'
import { applyTemplate } from './formatApply.js'
import { logger } from './logger.js'

export async function readSource(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase()
  if (ext === 'docx') {
    const zip = await openZip(file)
    const doc = await readDocumentXml(zip)
    const stylesDoc = await readStylesXml(zip)
    logger.info(`📂 已读取 .docx：${file.name}（${(file.size / 1024).toFixed(1)} KB）`)
    return { zip, doc, stylesDoc }
  }
  if (ext === 'doc') {
    const paras = await readDocText(await file.arrayBuffer())
    const zip = await buildDocxZipFromParagraphs(paras)
    const doc = await readDocumentXml(zip)
    logger.warn('⚠️ 旧版 .doc 以纯文本方式读取：原排版细节不保留，格式清洗 / 标题编号 / 模板套用均可正常执行')
    return { zip, doc, stylesDoc: null }
  }
  throw new Error('仅支持 .docx 和 .doc 格式，请重新选择文件')
}

// 仅用于预览：读取 + 标题检测（不修改内容）
export async function previewHeadings(file) {
  const src = await readSource(file)
  const stylesMap = buildStyleMap(src.stylesDoc)
  const { tree } = processHeadings(src.doc, stylesMap, { enableNumbering: false, textDetect: true, stripOld: false })
  return { tree, src }
}

// 统计清洗结果，生成质检数据
function buildQualityReport(tree, cleanStats, formatInfo) {
  const counts = [0, 0, 0, 0]
  const walk = (arr) => {
    for (const n of arr) {
      counts[n.level - 1]++
      if (n.children && n.children.length) walk(n.children)
    }
  }
  walk(tree)

  const cleanTotal = cleanStats
    ? Object.values(cleanStats).reduce((a, b) => a + (b || 0), 0)
    : 0

  return {
    headings: { total: counts.reduce((a, b) => a + b, 0), byLevel: counts },
    clean: { total: cleanTotal, items: cleanStats || {} },
    format: formatInfo || { headings: 0, markers: 0, tables: 0, template: '' }
  }
}

// 完整处理
export async function processFile(file, options) {
  const src = await readSource(file)
  const { zip, doc, stylesDoc } = src
  const stylesMap = buildStyleMap(stylesDoc)

  // 1. 格式清洗
  let cleanStats = null
  if (options.clean) {
    cleanStats = cleanDocument(doc, options.cleanOptions)
  } else {
    logger.info('跳过格式清洗（未勾选）')
  }

  // 2. 标题检测与自动编号（格式模板也需要标题层级）
  let tree = []
  let levels = new Map()
  let counted = 0
  if (options.number || options.format) {
    const r = processHeadings(doc, stylesMap, {
      enableNumbering: options.number,
      textDetect: options.numberOptions?.textDetect !== false,
      stripOld: options.number
    })
    tree = r.tree
    levels = r.levels
    counted = r.counted || 0
    applyOutlineLevels(doc, levels)
  }

  // 2.5 正文列举项编号转实心圆点（不在目录中，且非标题的带编号段落）
  const bulletCount = convertBodyNumbersToBullets(doc, levels, {
    enabled: options.numberOptions?.bulletsEnabled !== false
  })
  if (bulletCount > 0) {
    logger.info(`● 已将 ${bulletCount} 个正文编号段落统一为实心圆点项目符号`)
  }

  // 3. 模板格式化
  const formatInfo = { headings: 0, markers: 0, tables: 0, template: options.template?.name || '' }
  if (options.format && options.template) {
    const res = await applyTemplate(doc, zip, options.template, {
      ...options.formatOptions,
      levels
    })
    if (res) {
      formatInfo.headings = res.headingCount || 0
      formatInfo.markers = res.markerCount || 0
      formatInfo.tables = res.tableCount || 0
    }
  } else {
    logger.info('跳过模板格式化（未勾选）')
  }

  // 4. 写回输出
  writeXmlPart(zip, 'word/document.xml', doc)
  const blob = await buildBlob(zip)
  const outName = file.name.replace(/\.(docx|doc)$/i, '') + '-已格式化.docx'
  const report = buildQualityReport(tree, cleanStats, formatInfo)
  logger.success(`✅ 处理完成，已生成：${outName}`, { kind: 'done', report })
  return { blob, outName, tree, report, counted }
}
