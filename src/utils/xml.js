// OOXML XML 操作辅助工具（基于浏览器 DOMParser / XMLSerializer，命名空间感知）

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
export const CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
export const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

export function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml')
  const err = doc.getElementsByTagName('parsererror')
  if (err && err.length) throw new Error('XML 解析失败: ' + (err[0].textContent || '').slice(0, 200))
  return doc
}

export function serializeXml(node) {
  return new XMLSerializer().serializeToString(node)
}

export function wEl(tag, attrs = {}) {
  const e = document.createElementNS(W, tag)
  for (const k in attrs) e.setAttribute(k, String(attrs[k]))
  return e
}

export function textEl(text) {
  const t = wEl('w:t')
  t.setAttribute('xml:space', 'preserve')
  t.textContent = text
  return t
}

export function isW(node, localName) {
  return node && node.nodeType === 1 && node.localName === localName && node.namespaceURI === W
}

// 在 parent 下查找第一个指定名称的子元素
export function child(parent, localName) {
  for (const c of parent.childNodes) {
    if (isW(c, localName)) return c
  }
  return null
}

export function ensureChild(parent, localName) {
  let e = child(parent, localName)
  if (!e) {
    e = wEl('w:' + localName)
    parent.appendChild(e)
  }
  return e
}

export function removeChild(parent, localName) {
  const e = child(parent, localName)
  if (e) parent.removeChild(e)
  return e
}

export function setAttr(el, name, val) {
  if (val == null || val === '') el.removeAttribute(name)
  else el.setAttribute(name, String(val))
}

// 段落 pPr（无则创建并置于段落最前）
export function ppr(p) {
  let e = child(p, 'pPr')
  if (!e) {
    e = wEl('w:pPr')
    p.insertBefore(e, p.firstChild)
  }
  return e
}

// 文本 rPr（无则创建并置于 run 最前）
export function rpr(r) {
  let e = child(r, 'rPr')
  if (!e) {
    e = wEl('w:rPr')
    r.insertBefore(e, r.firstChild)
  }
  return e
}

// 获取 body 的直接段落（不含表格/文本框内的嵌套段落）
export function bodyParas(body) {
  return Array.from(body.childNodes).filter((n) => isW(n, 'p'))
}

// 段落内"本级"的所有 run：遍历 hyperlink / ins / smartTag 等容器，跳过嵌套段落
export function directRuns(p) {
  const out = []
  const walk = (node) => {
    for (const c of node.childNodes) {
      if (c.nodeType !== 1) continue
      const ln = c.localName
      if (ln === 'p') continue
      if (ln === 'r') { out.push(c); continue }
      if (ln === 'hyperlink' || ln === 'ins' || ln === 'smartTag' ||
          ln === 'moveFrom' || ln === 'moveTo' || ln === 'sdt' || ln === 'subDoc') {
        walk(c)
      }
    }
  }
  walk(p)
  return out
}

// 获取段落/run 的纯文本
export function getPText(p) {
  return Array.from(p.getElementsByTagNameNS(W, 't')).map((t) => t.textContent || '').join('')
}

export function getRText(r) {
  return Array.from(r.getElementsByTagNameNS(W, 't')).map((t) => t.textContent || '').join('')
}

// 在 run 中写入文本（合并多个 w:t，去掉 w:tab 之外的其他 w:t）
export function setRText(r, text) {
  const ts = Array.from(r.childNodes).filter((c) => isW(c, 't'))
  const first = ts[0]
  if (first) {
    first.textContent = text
    ts.slice(1).forEach((t) => r.removeChild(t))
  } else {
    const t = textEl(text)
    const rPr = child(r, 'rPr')
    if (rPr) rPr.parentNode.insertBefore(t, rPr.nextSibling)
    else r.insertBefore(t, r.firstChild)
  }
}

// 移除 run 的文本（保留 w:tab / w:br / 图形等）
export function clearRText(r) {
  Array.from(r.childNodes).filter((c) => isW(c, 't')).forEach((t) => r.removeChild(t))
}

export function runTextIsEmpty(r) {
  return Array.from(r.childNodes).filter((c) => isW(c, 't')).every((t) => !(t.textContent || '').trim())
}

// 查找文档 sectPr：优先 body 直接子元素，其次最后一个段落的 pPr 内
export function findSectPr(doc) {
  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  if (!body) return null
  let s = child(body, 'sectPr')
  if (s) return s
  const ps = bodyParas(body)
  if (ps.length) {
    s = child(ppr(ps[ps.length - 1]), 'sectPr')
    if (s) return s
  }
  return null
}

export function findStyles(doc) {
  const root = doc.getElementsByTagNameNS(W, 'styles')[0]
  return root || null
}
