// docx 读写：JSZip 打开/写回、XML 解析、页码页脚注入

import JSZip from 'jszip'
import { parseXml, serializeXml, CT, PKG_REL, child, findSectPr } from './xml.js'
import { logger } from './logger.js'

export const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

export async function openZip(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer())
  return zip
}

export async function readDocumentXml(zip) {
  const f = zip.file('word/document.xml')
  if (!f) throw new Error('不是有效的 .docx 文件：缺少 word/document.xml')
  return parseXml(await f.async('string'))
}

export async function readStylesXml(zip) {
  const f = zip.file('word/styles.xml')
  if (!f) return null
  return parseXml(await f.async('string'))
}

export function writeXmlPart(zip, path, xmlDoc) {
  zip.file(path, serializeXml(xmlDoc))
}

export async function buildBlob(zip) {
  return zip.generateAsync({
    type: 'blob',
    mimeType: DOCX_MIME,
    compression: 'DEFLATE'
  })
}

/**
 * 注入底部居中页码（域代码 PAGE）。幂等：已存在则跳过。
 * @returns {Promise<string|null>} rId
 */
export async function addPageFooter(zip, doc) {
  const sectPr = findSectPr(doc)
  if (sectPr && child(sectPr, 'footerReference')) return null

  const footerPath = 'word/footer1.xml'
  const ctFile = zip.file('[Content_Types].xml')
  const relsFile = zip.file('word/_rels/document.xml.rels')
  if (!ctFile || !relsFile) return null

  const ctDoc = parseXml(await ctFile.async('string'))
  const relsDoc = parseXml(await relsFile.async('string'))

  // 幂等
  const existing = Array.from(ctDoc.documentElement.children).find(
    (n) => n.localName === 'Override' && n.getAttribute('PartName') === '/' + footerPath
  )
  if (existing) return null

  // Content-Types
  const ov = ctDoc.createElementNS(CT, 'Override')
  ov.setAttribute('PartName', '/' + footerPath)
  ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
  ctDoc.documentElement.appendChild(ov)

  // 关系
  const used = new Set(Array.from(relsDoc.documentElement.children).map((n) => n.getAttribute('Id')))
  let i = 900
  while (used.has('rIdFooter' + i)) i++
  const rId = 'rIdFooter' + i
  const rel = relsDoc.createElementNS(PKG_REL, 'Relationship')
  rel.setAttribute('Id', rId)
  rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
  rel.setAttribute('Target', 'footer1.xml')
  relsDoc.documentElement.appendChild(rel)

  // 页脚内容：居中 "第 X 页"
  const ftr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/></w:rPr>
    </w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">第 </w:t></w:r>
    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> 页</w:t></w:r>
  </w:p>
</w:ftr>`

  zip.file(footerPath, ftr)
  writeXmlPart(zip, '[Content_Types].xml', ctDoc)
  writeXmlPart(zip, 'word/_rels/document.xml.rels', relsDoc)
  logger.info('📄 已在页脚注入页码（底部居中）')
  return rId
}

// 浏览器下载
export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 2000)
}
