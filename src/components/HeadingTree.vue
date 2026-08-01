<script setup>
import { ref, computed } from 'vue'
import TreeNode from './TreeNode.vue'

const props = defineProps({
  tree: { type: Array, default: () => [] },
  showNumbers: { type: Boolean, default: true }
})
const levelName = ['一级', '二级', '三级', '四级']

// 展开状态：path -> boolean，默认全收起（仅显示一级）
const openMap = ref({})

// 给树节点注入 path 和 __children
function withPath(arr, parent = '') {
  return arr.map((n, i) => {
    const path = parent ? parent + '-' + i : String(i)
    n.__path = path
    n.__children = n.children && n.children.length ? withPath(n.children, path) : []
    return n
  })
}
const treeWithPath = computed(() => withPath(props.tree))

function toggle(path) {
  openMap.value = { ...openMap.value, [path]: !openMap.value[path] }
}

// 全部展开 / 全部折叠
function expandAll() {
  const m = {}
  const walk = (arr) => {
    for (const n of arr) {
      m[n.__path] = true
      walk(n.__children)
    }
  }
  walk(treeWithPath.value)
  openMap.value = m
}
function collapseAll() {
  openMap.value = {}
}
</script>

<template>
  <div class="tree" v-if="tree.length">
    <div class="head">
      <span>标题结构（自动编号 · 父子依赖联动）</span>
      <span class="head-actions">
        <button class="mini" @click="expandAll">全部展开</button>
        <button class="mini" @click="collapseAll">全部收起</button>
      </span>
    </div>

    <ul class="list">
      <TreeNode
        v-for="n in treeWithPath"
        :key="n.__path"
        :node="n"
        :show-numbers="showNumbers"
        :level-name="levelName"
        :open-map="openMap"
        @toggle="toggle"
      />
    </ul>
  </div>
  <div class="tree empty" v-else>暂未检测到标题（可开启「按文本编号识别标题」）</div>
</template>

<style scoped>
.tree {
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 16px 18px;
  background: #fbfdff;
}
.head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 8px;
  font-size: 18px;
  font-weight: 700;
  color: var(--primary);
  margin-bottom: 12px;
}
.head-actions {
  display: flex;
  gap: 8px;
}
.mini {
  border: 1px solid var(--border);
  background: #fff;
  color: var(--primary-700);
  font-size: 14px;
  font-weight: 700;
  border-radius: 6px;
  padding: 5px 12px;
  cursor: pointer;
}
.mini:hover {
  border-color: var(--primary);
  background: var(--primary-50);
}
.empty {
  color: var(--muted);
  font-size: 17px;
  text-align: center;
  padding: 24px;
}
.list {
  list-style: none;
  margin: 0;
  padding: 0;
}
</style>
