<script setup>
import { ref } from 'vue'

const props = defineProps({
  file: { type: File, default: null },
  busy: Boolean
})
const emit = defineEmits(['select', 'clear'])
const inputRef = ref(null)
const dragging = ref(false)

function onPick() {
  inputRef.value.click()
}
function onInput(e) {
  const f = e.target.files && e.target.files[0]
  if (f) emit('select', f)
  e.target.value = ''
}
function onDrop(e) {
  dragging.value = false
  const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]
  if (f) emit('select', f)
}
const sizeText = (n) => (n > 1048576 ? (n / 1048576).toFixed(2) + ' MB' : (n / 1024).toFixed(1) + ' KB')
</script>

<template>
  <div class="upload">
    <div
      v-if="!props.file"
      class="zone"
      :class="{ dragging }"
      @click="onPick"
      @dragover.prevent="dragging = true"
      @dragleave="dragging = false"
      @drop.prevent="onDrop"
    >
      <div class="zone-icon">
        <span class="z-doc">DOC</span>
        <span class="z-plus">+</span>
      </div>
      <div class="title">点击选择或拖拽 Word 文档到此处</div>
      <div class="hint">支持 .docx 与 .doc 格式 · 全程本地处理，不上传任何服务器</div>
      <input ref="inputRef" type="file" accept=".docx,.doc" @change="onInput" />
    </div>

    <div v-else class="picked">
      <div class="file-info">
        <span class="ext">{{ props.file.name.split('.').pop() }}</span>
        <div class="meta">
          <div class="name">{{ props.file.name }}</div>
          <div class="size">{{ sizeText(props.file.size) }}</div>
        </div>
      </div>
      <div class="actions">
        <button class="ghost" :disabled="props.busy" @click="emit('clear')">更换文件</button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.zone {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 14px;
  padding: 48px 20px 40px;
  border: 3px dashed #9db3cb;
  border-radius: 12px;
  background: #fbfdff;
  cursor: pointer;
  transition: all 0.2s;
  text-align: center;
  min-height: 220px;
}
.zone:hover,
.zone.dragging {
  border-color: var(--primary);
  background: var(--primary-50);
  box-shadow: 0 0 0 5px var(--primary-100);
}
.zone input {
  display: none;
}
.zone-icon {
  position: relative;
  width: 84px;
  height: 84px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 14px;
  background: var(--primary-50);
  border: 1px solid var(--primary-100);
}
.z-doc {
  font-size: 20px;
  font-weight: 700;
  color: var(--primary-700);
  background: #fff;
  border: 1px solid var(--primary-100);
  border-radius: 8px;
  padding: 10px 14px;
  letter-spacing: 1px;
}
.z-plus {
  position: absolute;
  right: -10px;
  top: -10px;
  width: 32px;
  height: 32px;
  border-radius: 50%;
  background: var(--primary);
  color: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 22px;
  font-weight: 700;
  box-shadow: 0 2px 8px rgba(27, 107, 181, 0.4);
}
.title {
  font-size: var(--fs-lg);
  font-weight: 700;
  color: var(--text-strong);
}
.hint {
  font-size: var(--fs-sm);
  color: var(--muted);
}
.picked {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  padding: 20px 22px;
  border: 2px solid var(--border);
  border-radius: 12px;
  background: #fbfdff;
}
.file-info {
  display: flex;
  align-items: center;
  gap: 14px;
  min-width: 0;
}
.ext {
  flex: none;
  font-size: 18px;
  font-weight: 700;
  color: #fff;
  background: var(--primary);
  border-radius: 8px;
  padding: 10px 12px;
  text-transform: uppercase;
}
.meta {
  min-width: 0;
}
.name {
  font-size: 19px;
  font-weight: 700;
  color: var(--text-strong);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.size {
  font-size: 16px;
  color: var(--muted);
  margin-top: 3px;
}
.actions {
  flex: none;
}
.ghost {
  border: 2px solid var(--border);
  background: #fff;
  color: var(--muted);
  font-size: 18px;
  font-weight: 700;
  border-radius: 10px;
  padding: 12px 22px;
  cursor: pointer;
  transition: all 0.2s;
}
.ghost:hover {
  color: var(--primary);
  border-color: var(--primary);
  background: var(--primary-50);
}
</style>
