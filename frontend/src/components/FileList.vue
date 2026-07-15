<template>
  <div class="file-list-panel">
    <!-- 上传区 -->
    <div
      class="upload-zone"
      :class="{ dragover: isDragover }"
      @click="fileInput.click()"
      @dragover.prevent="isDragover = true"
      @dragleave.prevent="isDragover = false"
      @drop.prevent="onDrop"
    >
      <el-icon :size="32"><UploadFilled /></el-icon>
      <p>点击或拖拽上传 .docx</p>
      <input ref="fileInput" type="file" accept=".docx" multiple hidden @change="onSelect" />
    </div>

    <!-- 文件列表 -->
    <div class="file-list">
      <div class="list-header">
        <span>文件列表 ({{ files.length }})</span>
        <el-button v-if="files.length" size="small" text type="danger" @click="$emit('clearAll')">清空</el-button>
      </div>
      <div v-if="files.length === 0" class="empty">暂无文件</div>
      <div
        v-for="f in files"
        :key="f.id"
        class="file-item"
        :class="{ active: f.id === selectedId }"
        @click="$emit('select', f.id)"
      >
        <div class="item-main">
          <el-icon :size="18">
            <Loading v-if="f.status === 'processing'" class="spin" />
            <CircleCheckFilled v-else-if="f.status === 'completed'" color="#22c55e" />
            <Document v-else />
          </el-icon>
          <div class="item-info">
            <div class="item-name" :title="f.name">{{ f.name }}</div>
            <div class="item-size">{{ fmtSize(f.size) }}</div>
          </div>
        </div>
        <el-button
          v-if="f.status === 'completed'"
          size="small" :icon="Download" circle text type="primary"
          @click.stop="$emit('download', f.id)" title="下载"
        />
        <el-button
          size="small" :icon="Close" circle text
          @click.stop="$emit('remove', f.id)" title="移除"
        />
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { UploadFilled, Document, Loading, CircleCheckFilled, Download, Close } from '@element-plus/icons-vue'

defineProps({ files: { type: Array, default: () => [] }, selectedId: String })
const emit = defineEmits(['upload', 'select', 'remove', 'process', 'download', 'clearAll'])
const fileInput = ref(null)
const isDragover = ref(false)

function fmtSize(b) {
  if (!b) return ''
  if (b < 1024) return b + ' B'
  if (b < 1048576) return (b / 1024).toFixed(1) + ' KB'
  return (b / 1048576).toFixed(1) + ' MB'
}
function onDrop(e) { isDragover.value = false; const dt = e.dataTransfer; if (dt?.files) emit('upload', dt.files) }
function onSelect(e) { if (e.target.files) emit('upload', e.target.files); e.target.value = '' }
</script>

<style scoped>
.file-list-panel { display: flex; flex-direction: column; height: 100%; }
.upload-zone { margin: 12px; padding: 20px; border: 2px dashed #c8d9ef; border-radius: 10px; text-align: center; cursor: pointer; transition: all 0.2s; color: #2d6aa0; flex-shrink: 0; }
.upload-zone:hover, .upload-zone.dragover { border-color: #2d6aa0; background: #f0f7ff; }
.upload-zone p { font-size: 13px; margin-top: 8px; }
.file-list { flex: 1; overflow-y: auto; padding: 0 12px 12px; }
.list-header { display: flex; justify-content: space-between; align-items: center; padding: 6px 4px 8px; font-size: 13px; font-weight: 600; color: #555; }
.empty { text-align: center; color: #bbb; padding: 30px 0; font-size: 13px; }
.file-item { display: flex; align-items: center; justify-content: space-between; padding: 8px 10px; border-radius: 8px; cursor: pointer; transition: background 0.15s; gap: 6px; }
.file-item:hover { background: #f0f7ff; }
.file-item.active { background: #e6f0ff; border: 1px solid #b0d0ff; }
.item-main { display: flex; align-items: center; gap: 8px; flex: 1; min-width: 0; }
.item-info { min-width: 0; }
.item-name { font-size: 13px; font-weight: 500; color: #333; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.item-size { font-size: 11px; color: #999; }
.spin { animation: spin 1s linear infinite; }
@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
</style>
