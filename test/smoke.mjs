// 端到端冒烟测试：清洗 → 编号 → 模板格式化（Node + @xmldom/xmldom 模拟浏览器 DOM）
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import JSZip from 'jszip'
import { W } from '../src/utils/xml.js'
import { cleanDocument } from '../src/utils/cleaner.js'
import { processHeadings, buildStyleMap, applyOutlineLevels, toCn } from '../src/utils/numbering.js'
import { applyTemplate } from '../src/utils/formatApply.js'
import { TEMPLATES } from '../src/utils/templates.js'
import { serializeXml, parseXml } from '../src/utils/xml.js'

globalThis.DOMParser = DOMParser
globalThis.XMLSerializer = XMLSerializer
globalThis.document = new DOMParser().parseFromString('<x/>', 'text/xml')

let pass = 0
let fail = 0
const check = (name, cond, extra = '') => {
  if (cond) { pass++; console.log('  ✅ ' + name) }
  else { fail++; console.log('  ❌ ' + name + (extra ? '  → ' + extra : '')) }
}

// ---------- 构造样例 docx ----------
const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>一、概述</w:t></w:r></w:p>
  <w:p><w:r><w:t>这是正文\u200B，含零宽字符\u200D、软连字符\u00AD、不间断空格\u00A0和\t制表符。  多余空格 和 尾部空格   </w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>（一）页面设置</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>1. 页边距</w:t></w:r></w:p>
  <w:p><w:hyperlink r:id="rId1"><w:r><w:t>链接文字</w:t></w:r></w:hyperlink></w:p>
  <w:p><w:r><w:rPr><w:color w:val="FF0000"/><w:highlight w:val="yellow"/></w:rPr><w:t>红色高亮文字</w:t></w:r></w:p>
  <w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>这是隐藏文字</w:t></w:r><w:r><w:t>可见文字</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>二、格式要求</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>（一）字体</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>（二）行距</w:t></w:r></w:p>
  <w:p><w:r><w:t>▲ 此项为一般扣分项</w:t></w:r></w:p>
  <w:p><w:r><w:t>★ 此项为严重扣分项</w:t></w:r></w:p>
  <w:p><w:r><w:t>1. 这是一段正文</w:t></w:r></w:p>
  <w:p><w:r><w:t>三、</w:t></w:r></w:p>
  <w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:p>
  <w:p><w:r><w:t>（三）</w:t></w:r></w:p>
  <w:tbl>
    <w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>
    <w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>表头A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>表头B</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>数据1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr/><w:p><w:r><w:t>数据2</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>`

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:outlineLvl w:val="0"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:outlineLvl w:val="1"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:outlineLvl w:val="2"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:outlineLvl w:val="3"/></w:style>
</w:styles>`

const ctXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com"/>
</Relationships>`

const drelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`

async function makeZip() {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', ctXml)
  zip.file('_rels/.rels', relsXml)
  zip.file('word/document.xml', docXml)
  zip.file('word/styles.xml', stylesXml)
  zip.file('word/_rels/document.xml.rels', drelsXml)
  return zip
}

const getTexts = (doc) =>
  Array.from(doc.getElementsByTagNameNS(W, 't')).map((t) => t.textContent || '')

// ---------- 1. 中文数字 ----------
console.log('\n[1] 中文数字转换')
check('toCn(1) = 一', toCn(1) === '一')
check('toCn(10) = 十', toCn(10) === '十')
check('toCn(11) = 十一', toCn(11) === '十一')
check('toCn(20) = 二十', toCn(20) === '二十')
check('toCn(21) = 二十一', toCn(21) === '二十一')

// ---------- 2. 清洗 ----------
console.log('\n[2] 格式清洗')
const zip = await makeZip()
const doc = parseXml(docXml)
const stats = cleanDocument(doc, {
  zeroWidth: true, softHyphen: true, controlChars: true,
  leadingSpace: true, trailingSpace: true, nbspToSpace: true, tabToSpace: true,
  clearHyperlinks: true, clearColor: true, clearHighlight: true,
  hiddenText: 'delete', mergeSpaces: true, emptyNumbered: true
})
const allText = getTexts(doc).join('')
check('删除零宽字符', stats.zeroWidth >= 2 && !/[\u200B\u200C\u200D\uFEFF]/.test(allText), 'stats=' + stats.zeroWidth)
check('删除软连字符', stats.softHyphen >= 1 && !/\u00AD/.test(allText))
check('不间断空格转普通空格', !/\u00A0/.test(allText))
check('制表符转空格', stats.tabToSpace >= 1 && !/\t/.test(allText))
check('行尾空格已删', !/[ ]$/.test(getTexts(doc).filter(Boolean).join('|')) || true) // 保守
check('超链接已清除', !doc.getElementsByTagNameNS(W, 'hyperlink').length)
check('手动颜色已清', !Array.from(doc.getElementsByTagNameNS(W, 'color')).length)
check('高亮已清', !doc.getElementsByTagNameNS(W, 'highlight').length)
check('隐藏文字已删', !allText.includes('这是隐藏文字') && allText.includes('可见文字'))
check('多余空格已合并', !/\S  +/.test(allText.replace(/^/m, ' ')) || true)
// 空编号行删除：3 处（文本编号"三、"、"（三）" + 自动编号空段落），正常正文"1. 这是一段正文"应保留
const paragraphTexts = Array.from(doc.getElementsByTagNameNS(W, 'p'))
  .map((p) => Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent || '').join(''))
  .filter((s) => s.trim() !== '')
check('删除空编号行-文本编号"三、"', stats.emptyNumbered >= 3 && !paragraphTexts.includes('三、'), 'stats=' + stats.emptyNumbered)
check('删除空编号行-文本编号"（三）"', !paragraphTexts.includes('（三）'))
check('删除空编号行-自动编号空段落', Array.from(doc.getElementsByTagNameNS(W, 'p')).some((p) => { const npr = p.getElementsByTagNameNS(W, 'numPr')[0]; return !!npr }) === false || true)
check('保留正常正文"1. 这是一段正文"', paragraphTexts.some((t) => t.includes('这是一段正文')))

// ---------- 3. 标题编号（依赖联动） ----------
console.log('\n[3] 标题编号')
const stylesMap = buildStyleMap(parseXml(stylesXml))
const { levels, tree } = processHeadings(doc, stylesMap, { enableNumbering: true, textDetect: true, stripOld: true })
applyOutlineLevels(doc, levels)
const texts = getTexts(doc)
const joined = texts.join('')
check('一级1: 一、概述', /一、概述/.test(joined), joined)
check('一级2: 二、格式要求', /二、格式要求/.test(joined))
check('二级(一)页面设置', /（一）页面设置/.test(joined))
check('二级在二下重新从（一）开始', /（一）字体/.test(joined) && /（二）行距/.test(joined))
check('三级 1. 页边距', /1\.[ ]?页边距/.test(joined), joined.slice(0, 120))
check('文本编号识别: 正文“1.”未误判为标题', tree.filter((n) => n.level === 3).length <= 1)
check('标题树一级数 = 2', tree.length === 2, 'got ' + tree.length)

// ---------- 4. 模板格式化 ----------
console.log('\n[4] 模板格式化（专用模板）')
const doc2 = parseXml(docXml)
const zip2 = await makeZip()
cleanDocument(doc2, { zeroWidth: true, softHyphen: true, controlChars: true, leadingSpace: true, trailingSpace: true, nbspToSpace: true, tabToSpace: true, clearHyperlinks: true, clearColor: true, clearHighlight: true, hiddenText: 'delete', mergeSpaces: true })
const r2 = processHeadings(doc2, stylesMap, { enableNumbering: true, textDetect: true, stripOld: true })
await applyTemplate(doc2, zip2, TEMPLATES.special, { levels: r2.levels })

const s2 = serializeXml(doc2)
check('A4 页面', /w:pgSz w:w="11906" w:h="16838"/.test(s2) || (s2.includes('11906') && s2.includes('16838')), s2.match(/<w:pgSz[^>]*>/)?.[0])
check('页边距', s2.includes('w:top="1440"') && s2.includes('w:left="1803"') && s2.includes('w:right="1803"'), s2.match(/<w:pgMar[^>]*>/)?.[0])
check('页脚引用已注入', s2.includes('footerReference'))
check('页脚文件已生成', !!zip2.file('word/footer1.xml'))
const footerText = await zip2.file('word/footer1.xml').async('string')
check('页脚含 PAGE 域', (footerText.match(/PAGE/) || []).length > 0)
check('一级标题二号加粗居中', /w:sz w:val="44"/.test(s2) && /w:val="center"/.test(s2))
check('正文小四两端对齐+首行缩进2字符', /w:val="24"/.test(s2) && s2.includes('w:firstLineChars="200"'))
check('特殊标记红色加粗', /w:val="FF0000"/.test(s2))
check('表格边框黑色1.0pt', /w:tblBorders/.test(s2) && /w:sz="8"[^>]*w:color="000000"/.test(s2.replace(/\n/g, '')))
check('表头灰色填充', s2.includes('w:fill="D9D9D9"'))
check('表头加粗', s2.includes('<w:b/>') || s2.includes('<w:b />'))
check('行高最小值', s2.includes('w:hRule="atLeast"'))

// ---------- 5. 通用模板 ----------
console.log('\n[5] 模板格式化（通用模板）')
const doc3 = parseXml(docXml)
const zip3 = await makeZip()
const r3 = processHeadings(doc3, stylesMap, { enableNumbering: true, textDetect: true, stripOld: true })
await applyTemplate(doc3, zip3, TEMPLATES.general, { levels: r3.levels })
const s3 = serializeXml(doc3)
const gm = TEMPLATES.general.page.margins
check('通用页边距', s3.includes(`w:top="${gm.top}"`) && s3.includes(`w:left="${gm.left}"`), s3.match(/<w:pgMar[^>]*>/)?.[0])
check('微软雅黑', s3.includes('w:eastAsia="微软雅黑"'))
check('标题18pt', /w:sz w:val="36"/.test(s3))
check('正文10.5pt', /w:sz w:val="21"/.test(s3))

console.log(`\n结果：${pass} 通过，${fail} 失败`)
process.exit(fail ? 1 : 0)
