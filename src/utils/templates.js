// 两套格式化模板定义（专用模板 / 通用模板）
// 尺寸单位：磅（pt 字号用半磅 halfPt）、行距用 w:line（1/240 行基准）、间距用 twips（1/20 磅）

export const cm2twip = (cm) => Math.round(cm * 567)

const A4 = { w: 11906, h: 16838 }

export const TEMPLATES = {
  special: {
    id: 'special',
    name: '专用模板',
    desc: 'A4 纸张 · 标准页边距 · 宋体/Times New Roman · 1.5 倍行距',
    page: {
      size: A4,
      margins: { top: cm2twip(2.54), bottom: cm2twip(2.54), left: cm2twip(3.18), right: cm2twip(3.18), gutter: 0 }
    },
    fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '宋体' },
    line: 360, // 1.5 倍行距
    body: {
      size: 24, // 小四 12pt
      align: 'both',
      indentChars: 200,
      indentTwip: 420,
      beforeLines: 0,
      afterLines: 0
    },
    headings: [
      // 一级：二号 22pt 加粗居中，段前段后 0.5 行
      { size: 44, bold: true, align: 'center', beforeLines: 50, afterLines: 50 },
      // 二级：三号 16pt 加粗左对齐，段前 0.5 行 段后 0.3 行
      { size: 32, bold: true, align: 'left', beforeLines: 50, afterLines: 30 },
      // 三级：四号 14pt 加粗左对齐，段前段后 0.3 行
      { size: 28, bold: true, align: 'left', beforeLines: 30, afterLines: 30 },
      // 四级：小四 12pt 加粗左对齐，段前段后 0.2 行
      { size: 24, bold: true, align: 'left', beforeLines: 20, afterLines: 20 }
    ],
    marker: { chars: ['▲', '★'], color: 'FF0000', bold: true },
    table: {
      borderSz: 8, // 1.0pt
      headerFill: 'D9D9D9', // 灰色 15%
      headerBold: true,
      rowMinHeight: 340 // 0.6cm ≈ 340 twips
    },
    footer: true
  },
  general: {
    id: 'general',
    name: '通用模板',
    desc: 'A4 纸张 · 2.40/2.50cm 页边距 · 微软雅黑 · 标题行距 1.2',
    page: {
      size: A4,
      margins: { top: cm2twip(2.40), bottom: cm2twip(2.40), left: cm2twip(2.50), right: cm2twip(2.50), gutter: 0 }
    },
    fonts: { ascii: '微软雅黑', hAnsi: '微软雅黑', eastAsia: '微软雅黑' },
    line: 288, // 标题 1.2 倍行距
    body: {
      size: 21, // 微软雅黑 10.5pt
      align: 'both',
      indentChars: 200,
      indentTwip: 420, // 0.74cm
      beforeLines: 0,
      afterLines: 0
    },
    headings: [
      { size: 36, bold: true, align: 'left', before: 50, after: 20 }, // 18pt 段前 2.5pt 段后 1pt
      { size: 32, bold: true, align: 'left', before: 50, after: 20 }, // 16pt
      { size: 30, bold: true, align: 'left', before: 50, after: 20 }, // 15pt
      { size: 28, bold: true, align: 'left', before: 50, after: 20 }  // 14pt
    ],
    marker: { chars: ['▲', '★'], color: 'FF0000', bold: true },
    table: {
      borderSz: 8,
      headerFill: 'D9D9D9',
      headerBold: true,
      rowMinHeight: 340
    },
    footer: true
  }
}

// 默认清洗选项（13 项，全开）
export const CLEAN_DEFAULT = {
  zeroWidth: true,      // 删除零宽字符
  softHyphen: true,     // 删除软连字符
  controlChars: true,   // 删除不可见控制字符
  leadingSpace: true,   // 删除行首空格
  trailingSpace: true,  // 删除行尾空格
  nbspToSpace: true,    // 不间断空格→普通空格
  tabToSpace: true,     // 制表符→空格
  clearHyperlinks: true, // 清除超链接格式
  clearColor: true,     // 清除手动颜色
  clearHighlight: true, // 清除手动高亮
  hiddenText: 'delete', // 清除隐藏文字：delete 删除 / display 显示 / keep 保留
  mergeSpaces: true,    // 合并多余空格
  emptyNumbered: true   // 删除有编号但内容为空的段落
}

export const CLEAN_ITEMS = [
  { key: 'zeroWidth', label: '删除零宽字符', desc: '去除 \u200B \u200C \u200D 等零宽字符' },
  { key: 'softHyphen', label: '删除软连字符', desc: '去除 \u00AD（软连字符）' },
  { key: 'controlChars', label: '删除不可见控制字符', desc: '去除 \u0000-\u001F、\u007F-\u009F 控制字符' },
  { key: 'leadingSpace', label: '删除行首空格', desc: '去除每段行首空格' },
  { key: 'trailingSpace', label: '删除行尾空格', desc: '去除每段行尾空格' },
  { key: 'nbspToSpace', label: '不间断空格转普通空格', desc: '将 \u00A0 替换为普通空格' },
  { key: 'tabToSpace', label: '制表符转空格', desc: '将 Tab 制表符替换为空格' },
  { key: 'clearHyperlinks', label: '清除超链接格式', desc: '移除超链接，仅保留文字' },
  { key: 'clearColor', label: '清除手动颜色', desc: '颜色重置为自动（黑色）' },
  { key: 'clearHighlight', label: '清除手动高亮', desc: '移除文字背景高亮' },
  { key: 'hiddenText', label: '清除隐藏文字', desc: '删除或显示隐藏文字' },
  { key: 'mergeSpaces', label: '合并多余空格', desc: '连续多个空格合并为一个' },
  { key: 'emptyNumbered', label: '删除空编号行', desc: '删除只有编号、无实际内容的段落' }
]
