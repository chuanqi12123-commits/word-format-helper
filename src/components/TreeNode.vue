<script setup>
import TreeNode from './TreeNode.vue'

defineProps({
  node: { type: Object, required: true },
  showNumbers: { type: Boolean, default: true },
  levelName: { type: Array, default: () => ['一级', '二级', '三级', '四级'] },
  openMap: { type: Object, required: true }
})
const emit = defineEmits(['toggle'])
</script>

<template>
  <li class="lv" :class="'lv' + node.level">
    <span
      v-if="node.__children && node.__children.length"
      class="arrow"
      :class="{ open: openMap[node.__path] === true }"
      @click.stop="emit('toggle', node.__path)"
    >▶</span>
    <span v-else class="arrow placeholder"></span>
    <span class="num" v-if="showNumbers && node.number" @click.stop="emit('toggle', node.__path)">{{ node.number }}</span>
    <span class="txt" @click.stop="emit('toggle', node.__path)">{{ node.title || '（空标题）' }}</span>
    <span class="tag">{{ levelName[node.level - 1] }}</span>

    <ul class="list" v-if="node.__children && node.__children.length && openMap[node.__path] === true">
      <TreeNode
        v-for="c in node.__children"
        :key="c.__path"
        :node="c"
        :show-numbers="showNumbers"
        :level-name="levelName"
        :open-map="openMap"
        @toggle="emit('toggle', $event)"
      />
    </ul>
  </li>
</template>

<style scoped>
.lv {
  display: flex;
  align-items: baseline;
  gap: 8px;
  padding: 5px 0;
  font-size: 17px;
  color: var(--text);
  flex-wrap: wrap;
}
.list {
  list-style: none;
  margin: 0;
  padding: 0;
  width: 100%;
  padding-left: 22px;
  border-left: 1px dashed #cfd9e6;
  margin-left: 8px;
}
.arrow {
  flex: none;
  width: 18px;
  height: 18px;
  font-size: 12px;
  color: var(--primary-700);
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  transition: transform 0.2s;
  transform: translateY(-1px);
}
.arrow.open {
  transform: rotate(90deg) translateY(0);
}
.arrow.placeholder {
  visibility: hidden;
}
.num {
  flex: none;
  font-weight: 700;
  color: var(--primary-700);
  cursor: pointer;
}
.txt {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  cursor: pointer;
}
.tag {
  flex: none;
  font-size: 14px;
  font-weight: 700;
  color: #fff;
  border-radius: 6px;
  padding: 2px 12px;
  transform: translateY(-1px);
}
.lv1 > .tag { background: #1b6bb5; }
.lv2 > .tag { background: #2f8bd0; }
.lv3 > .tag { background: #4a9ad6; }
.lv4 > .tag { background: #6f9ec4; }
</style>
