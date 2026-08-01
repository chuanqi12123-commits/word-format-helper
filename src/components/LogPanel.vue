<script setup>
import { ref, onMounted, onBeforeUnmount, nextTick } from 'vue'
import { logger } from '../utils/logger.js'

const entries = ref([])
const boxRef = ref(null)
let unsub = null

onMounted(() => {
  entries.value = logger.entries.slice()
  unsub = logger.subscribe((e) => {
    entries.value.push(e)
    nextTick(() => scrollBottom())
  })
})
onBeforeUnmount(() => unsub && unsub())

function scrollBottom() {
  const el = boxRef.value
  if (el) el.scrollTop = el.scrollHeight
}

const time = (d) =>
  `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}`

const levelText = { info: '信息', success: '成功', warn: '警告', error: '错误', debug: '调试' }
const levelColor = { info: '#1b6bb5', success: '#1e8f54', warn: '#c26a12', error: '#c83434', debug: '#5a677a' }

function clearLog() {
  logger.clear()
  entries.value = []
}
</script>

<template>
  <div class="log">
    <div class="log-head">
      <span>运行日志</span>
      <span class="count">{{ entries.length }}</span>
      <button class="clear" @click="clearLog">清空</button>
    </div>
    <div class="log-body" ref="boxRef">
      <div v-if="!entries.length" class="empty">暂无日志，请选择文件并开始处理</div>
      <div v-for="e in entries" :key="e.id" class="row">
        <span class="dot" :style="{ background: levelColor[e.level] }"></span>
        <span class="badge" :style="{ color: levelColor[e.level], background: levelColor[e.level] + '1f' }">{{ levelText[e.level] }}</span>
        <span class="msg">{{ e.message }}</span>
        <span class="time">{{ time(e.time) }}</span>
      </div>
    </div>
  </div>
</template>

<style scoped>
.log {
  border: 1px solid var(--border);
  border-radius: 10px;
  background: #fbfdff;
  overflow: hidden;
}
.log-head {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 12px 18px;
  color: var(--text-strong);
  font-size: 18px;
  font-weight: 700;
  border-bottom: 1px solid var(--border);
  background: #f0f5fb;
}
.count {
  font-size: 15px;
  font-weight: 700;
  color: var(--primary-700);
  background: var(--primary-50);
  border-radius: 999px;
  padding: 2px 12px;
}
.clear {
  margin-left: auto;
  font-size: 16px;
  color: var(--muted);
  background: #fff;
  border: 2px solid var(--border);
  border-radius: 8px;
  padding: 6px 18px;
  cursor: pointer;
  transition: all 0.2s;
}
.clear:hover {
  color: var(--primary);
  border-color: var(--primary);
}
.log-body {
  max-height: 240px;
  overflow-y: auto;
  padding: 8px 18px 12px;
  font-size: 17px;
  line-height: 1.8;
}
.row {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 5px 0;
  border-bottom: 1px dashed #e3eaf2;
}
.row:last-child {
  border-bottom: none;
}
.dot {
  flex: none;
  width: 10px;
  height: 10px;
  border-radius: 50%;
}
.badge {
  flex: none;
  font-size: 14px;
  font-weight: 700;
  border-radius: 6px;
  padding: 2px 10px;
}
.msg {
  color: var(--text);
  word-break: break-all;
  flex: 1;
  min-width: 0;
}
.time {
  flex: none;
  color: var(--muted);
  font-family: Consolas, monospace;
  font-size: 14px;
}
.empty {
  color: var(--muted);
  text-align: center;
  padding: 24px;
  font-size: 17px;
}
.log-body::-webkit-scrollbar {
  width: 10px;
}
.log-body::-webkit-scrollbar-thumb {
  background: #a7b9cd;
  border-radius: 5px;
}
</style>
