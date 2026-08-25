<template>
  <div class="structure-panel" v-loading="loading">
    <div class="struct-header" v-if="data || error">
      <div class="struct-title">
        <span>识别结果 · 共 {{ paragraphs.length }} 段</span>
        <span class="low-tip" v-if="lowCount > 0">⚠ {{ lowCount }} 段低置信，请复核</span>
      </div>
      <p class="struct-hint">逐段确认类型后，点右侧「套用格式」生成公文。</p>
      <p v-if="error" class="struct-error">{{ error }}</p>
    </div>

    <div class="struct-body">
      <el-empty v-if="!selectedFile" description="选择左侧文件即可识别" :image-size="90" />
      <el-empty v-else-if="loading" description="识别中..." :image-size="60" />
      <el-empty v-else-if="error" :description="error" :image-size="60" />
      <el-empty v-else-if="!data || !paragraphs.length" description="暂无可识别段落" :image-size="60" />
      <template v-else>
        <div
          v-for="p in paragraphs"
          :key="p.index"
          class="struct-row"
          :class="{ low: p.lowConfidence }"
        >
          <div class="row-index">{{ p.index + 1 }}</div>
          <div class="row-text" :title="p.fullText">{{ p.text }}</div>
          <div class="row-conf" :class="{ low: p.lowConfidence }">{{ Math.round((p.confidence || 0) * 100) }}%</div>
          <el-select
            :model-value="p.role"
            size="small"
            class="row-select"
            @update:model-value="val => emitChange(p.index, val)"
          >
            <el-option
              v-for="opt in roleOptions"
              :key="opt.value"
              :label="opt.label"
              :value="opt.value"
            />
          </el-select>
        </div>
      </template>
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'

const props = defineProps({
  selectedFile: Object,
  data: Object,
  labels: { type: Object, default: () => ({}) },
  loading: Boolean,
  error: { type: String, default: '' },
})
const emit = defineEmits(['changeRole'])

// 校正下拉的角色顺序（仅展示有标签的）
const ROLE_ORDER = ['title', 'subtitle', 'recipient', 'body', 'h1', 'h2', 'h3', 'h4', 'attachment_head', 'attachment_other', 'sign_unit', 'sign_date', 'sign_contact', 'security', 'other']
const FALLBACK_LABELS = {
  title: '标题', subtitle: '副标题', recipient: '发文对象', body: '正文', h1: '一级标题', h2: '二级标题',
  h3: '三级标题', h4: '四级标题', attachment_head: '附件头', attachment_other: '其他附件', sign_unit: '落款·单位',
  sign_date: '落款·日期', sign_contact: '落款·联系人', security: '涉密标识', other: '其他',
}

const paragraphs = computed(() => props.data?.paragraphs || [])
const lowCount = computed(() => paragraphs.value.filter(p => p.lowConfidence).length)

const roleOptions = computed(() => {
  const labels = { ...FALLBACK_LABELS, ...props.labels }
  return ROLE_ORDER.filter(r => labels[r]).map(r => ({ value: r, label: labels[r] }))
})

function emitChange(index, role) {
  emit('changeRole', { index, role })
}
</script>

<style scoped>
.structure-panel { display: flex; flex-direction: column; height: 100%; min-height: 0; }
.struct-header { padding: 12px 18px; border-bottom: 1px solid #edf2f7; background: #f6f9fc; flex-shrink: 0; }
.struct-title { display: flex; align-items: center; gap: 14px; color: #1a3a5c; font-size: 14px; font-weight: 700; }
.low-tip { color: #c05621; font-size: 12px; font-weight: 600; }
.struct-hint { margin: 4px 0 0; color: #7a8997; font-size: 12px; }
.struct-error { margin: 6px 0 0; color: #c05621; font-size: 12px; font-weight: 600; }
.struct-body { flex: 1; overflow-y: auto; padding: 8px 12px; background: #fbfdff; }
.struct-row {
  display: grid; grid-template-columns: 32px 1fr 48px 130px; align-items: center; gap: 10px;
  padding: 7px 8px; border-bottom: 1px solid #f0f4f8;
}
.struct-row.low { background: #fff8ef; border-left: 3px solid #ed8936; }
.row-index { color: #9aa8b6; font-size: 12px; text-align: center; }
.row-text { color: #2d3748; font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.row-conf { font-size: 12px; color: #38a169; text-align: right; font-variant-numeric: tabular-nums; }
.row-conf.low { color: #dd6b20; font-weight: 700; }
.row-select { width: 130px; }
</style>
