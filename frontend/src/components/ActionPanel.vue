<template>
  <div class="action-panel">
    <h3>操作</h3>

    <div class="file-info" v-if="selectedFile">
      <div class="info-name" :title="selectedFile.name">{{ selectedFile.name }}</div>
      <div class="info-size">{{ fmtSize(selectedFile.size) }}</div>
      <el-tag v-if="isCompleted" type="success" size="small" effect="light">已处理</el-tag>
      <el-tag v-else-if="isProcessing" type="warning" size="small" effect="light">处理中</el-tag>
      <el-tag v-else type="info" size="small" effect="light">待处理</el-tag>
    </div>
    <div v-else class="no-file">未选择文件</div>

    <div class="template-box">
      <label>转换模板</label>
      <el-select
        :model-value="selectedTemplateId"
        size="default"
        style="width:100%;"
        :disabled="!selectedFile || isProcessing"
        @update:model-value="value => $emit('changeTemplate', value)"
      >
        <el-option
          v-for="template in templates"
          :key="template.id"
          :label="template.label"
          :value="template.id"
        />
      </el-select>
      <p v-if="selectedTemplate" class="template-description">{{ selectedTemplate.description }}</p>
    </div>

    <el-button
      type="primary" size="large" style="width:100%;"
      :loading="batchProcessing"
      :disabled="!selectedFile || isProcessing"
      @click="$emit('process', selectedFile?.id)"
    >
      <el-icon><Operation /></el-icon>
      套用格式（当前文件）
    </el-button>
    <el-button
      size="large" style="width:100%; margin-top:8px;"
      :disabled="totalCount === 0 || pendingCount === 0 || batchProcessing"
      @click="$emit('process')"
    >
      批量处理全部
    </el-button>

    <el-button
      type="success" size="large" style="width:100%; margin-top:12px;"
      :disabled="!isCompleted"
      @click="$emit('download', selectedFile?.id)"
    >
      <el-icon><Download /></el-icon> 下载结果
    </el-button>

    <el-divider />
    <div class="stats">
      <div>总文件: {{ totalCount }} 个</div>
      <div>待处理: {{ pendingCount }} 个</div>
      <div>处理中: {{ processingCount }} 个</div>
      <div>已处理: {{ completedCount }} 个</div>
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'
import { Operation, Download } from '@element-plus/icons-vue'

const props = defineProps({
  selectedFile: Object,
  isProcessing: Boolean,
  isCompleted: Boolean,
  totalCount: { type: Number, default: 0 },
  completedCount: { type: Number, default: 0 },
  processingCount: { type: Number, default: 0 },
  pendingCount: { type: Number, default: 0 },
  templates: { type: Array, default: () => [] },
  selectedTemplateId: { type: String, default: 'generic' },
})
defineEmits(['process', 'download', 'changeTemplate'])

const batchProcessing = computed(() => props.processingCount > 0 || props.isProcessing)
const selectedTemplate = computed(() => props.templates.find(item => item.id === props.selectedTemplateId) || null)

function fmtSize(b) {
  if (!b) return ''
  if (b < 1024) return b + ' B'
  if (b < 1048576) return (b / 1024).toFixed(1) + ' KB'
  return (b / 1048576).toFixed(1) + ' MB'
}
</script>

<style scoped>
.action-panel { height: 100%; overflow-y: auto; padding: 16px; }
.action-panel h3 { font-size: 16px; color: #1a3a5c; margin-bottom: 16px; }
.file-info { margin-bottom: 16px; padding: 12px; background: #f8fafc; border-radius: 8px; }
.info-name { font-size: 13px; font-weight: 600; color: #333; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-bottom: 4px; }
.info-size { font-size: 11px; color: #999; margin-bottom: 6px; }
.no-file { text-align: center; color: #ccc; padding: 24px 0; font-size: 13px; }
.template-box { margin-bottom: 14px; padding: 10px; border: 1px solid #dce8f4; border-radius: 8px; background: #f8fbff; }
.template-box > label { display: block; margin-bottom: 6px; color: #1a3a5c; font-size: 12px; font-weight: 700; }
.template-description { margin: 7px 2px 0; color: #718096; font-size: 11px; line-height: 1.5; }
.stats { font-size: 12px; color: #888; line-height: 2; }
</style>
