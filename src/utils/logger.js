// 日志系统：供各处理模块与 UI 共享，支持订阅实时推送
export const LEVEL_TAG = {
  info: '信息',
  success: '成功',
  warn: '警告',
  error: '错误',
  debug: '调试'
}

class Logger {
  constructor() {
    this.entries = []
    this.listeners = new Set()
    this.max = 3000
  }

  _push(level, message, meta) {
    const entry = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      time: new Date(),
      level,
      message: String(message),
      meta
    }
    this.entries.push(entry)
    if (this.entries.length > this.max) {
      this.entries.splice(0, this.entries.length - this.max)
    }
    this.listeners.forEach((fn) => {
      try { fn(entry) } catch (e) { /* ignore */ }
    })
  }

  info(msg, meta) { this._push('info', msg, meta) }
  success(msg, meta) { this._push('success', msg, meta) }
  warn(msg, meta) { this._push('warn', msg, meta) }
  error(msg, meta) { this._push('error', msg, meta) }
  debug(msg, meta) { this._push('debug', msg, meta) }

  subscribe(fn) {
    this.listeners.add(fn)
    return () => this.listeners.delete(fn)
  }

  clear() {
    this.entries = []
  }

  get count() {
    return this.entries.length
  }
}

export const logger = new Logger()
