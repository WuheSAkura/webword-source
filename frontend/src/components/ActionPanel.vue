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
    <div class="edit-panel">
      <h3>局部微调</h3>
      <div v-if="selection" class="selection-box" :title="selection.text">{{ selectionText }}</div>
      <div v-else class="selection-empty">在处理后预览中选中文字</div>

      <div class="edit-title">字体</div>
      <div class="edit-group" :class="{ disabled: !selection }">
        <label>中文字体</label>
        <el-select v-model="fontForm.font_east" size="small" :disabled="!selection">
          <el-option v-for="f in cnFonts" :key="f" :label="f" :value="f" />
        </el-select>
        <label>西文字体</label>
        <el-select v-model="fontForm.font_west" size="small" :disabled="!selection">
          <el-option label="Times New Roman" value="Times New Roman" />
        </el-select>
        <label>字号</label>
        <el-select v-model="fontForm.size" size="small" :disabled="!selection">
          <el-option v-for="s in sizeOptions" :key="s.value" :label="s.label" :value="s.value" />
        </el-select>
        <el-checkbox v-model="fontForm.bold" :disabled="!selection">加粗</el-checkbox>
      </div>

      <div class="edit-title">段落</div>
      <div class="edit-group" :class="{ disabled: !selection }">
        <label>段落类型</label>
        <el-select v-model="paragraphForm.role" size="small" :disabled="!selection" @change="applyRoleDefaults">
          <el-option v-for="opt in roleOptions" :key="opt.value" :label="opt.label" :value="opt.value" />
        </el-select>
        <label>对齐</label>
        <el-select v-model="paragraphForm.alignment" size="small" :disabled="!selection">
          <el-option label="左对齐" value="left" />
          <el-option label="居中" value="center" />
          <el-option label="右对齐" value="right" />
          <el-option label="两端对齐" value="justify" />
        </el-select>
        <label>行距规则</label>
        <el-select v-model="paragraphForm.line_rule" size="small" :disabled="!selection">
          <el-option label="固定值" value="exact" />
          <el-option label="单倍行距" value="single" />
          <el-option label="多倍行距" value="multiple" />
        </el-select>
        <template v-if="paragraphForm.line_rule === 'exact'">
          <label>固定行距</label>
          <el-input-number v-model="paragraphForm.line_spacing" size="small" :min="12" :max="60" :disabled="!selection" />
        </template>
        <template v-else-if="paragraphForm.line_rule === 'multiple'">
          <label>行距倍数</label>
          <el-input-number v-model="paragraphForm.line_multiple" size="small" :min="0.5" :max="3" :step="0.05" :precision="2" :disabled="!selection" />
        </template>
        <label>首行缩进(字)</label>
        <el-input-number v-model="paragraphForm.first_indent" size="small" :min="0" :max="6" :step="1" :disabled="!selection" />
      </div>

      <el-button type="primary" size="small" style="width:100%; margin-top:10px;" :disabled="!selection" @click="emitApply">
        应用微调
      </el-button>
      <el-button size="small" style="width:100%; margin:8px 0 0;" :disabled="!isCompleted" @click="$emit('undoEdit')">
        撤销上一步
      </el-button>
    </div>

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
import { computed, ref, watch } from 'vue'
import { Operation, Download } from '@element-plus/icons-vue'

const props = defineProps({
  selectedFile: Object,
  isProcessing: Boolean,
  isCompleted: Boolean,
  totalCount: { type: Number, default: 0 },
  completedCount: { type: Number, default: 0 },
  processingCount: { type: Number, default: 0 },
  pendingCount: { type: Number, default: 0 },
  selection: Object,
  styles: { type: Object, default: () => ({}) },
  labels: { type: Object, default: () => ({}) },
  templates: { type: Array, default: () => [] },
  selectedTemplateId: { type: String, default: 'generic' },
})
const emit = defineEmits(['process', 'download', 'applyEdit', 'undoEdit', 'changeTemplate'])

const batchProcessing = computed(() => props.processingCount > 0 || props.isProcessing)
const selectedTemplate = computed(() => props.templates.find(item => item.id === props.selectedTemplateId) || null)

const ROLE_ORDER = ['title', 'subtitle', 'recipient', 'body', 'h1', 'h2', 'h3', 'h4', 'attachment_head', 'attachment_other', 'sign_unit', 'sign_date', 'sign_contact', 'security']
const FALLBACK_LABELS = {
  title: '标题', subtitle: '副标题', recipient: '发文对象', body: '正文', h1: '一级标题', h2: '二级标题',
  h3: '三级标题', h4: '四级标题', attachment_head: '附件头', attachment_other: '其他附件', sign_unit: '落款·单位',
  sign_date: '落款·日期', sign_contact: '落款·联系人', security: '涉密标识',
}

const cnFonts = ['仿宋_GB2312', '方正小标宋简体', '黑体', '楷体_GB2312', '宋体']
const sizeOptions = [
  { label: '二号 22pt', value: 22 }, { label: '三号 16pt', value: 16 },
  { label: '小三 15pt', value: 15 }, { label: '四号 14pt', value: 14 },
  { label: '小四 12pt', value: 12 }, { label: '五号 10.5pt', value: 10.5 },
]

const roleOptions = computed(() => {
  const labels = { ...FALLBACK_LABELS, ...props.labels }
  return ROLE_ORDER.filter(r => labels[r]).map(r => ({ value: r, label: labels[r] }))
})

// 从配置（props.styles）派生每个角色的默认排版，消除前端硬编码
function styleToDefault(role) {
  const s = props.styles?.[role]
  if (!s) return null
  const alignRaw = s.align || 'justify'
  const alignment = ['center_to_date', 'right_indent_4chars', 'right_indent_2chars'].includes(alignRaw) ? 'right' : alignRaw
  return {
    font_east: s.cn,
    font_west: 'Times New Roman',
    size: s.size_pt,
    bold: !!s.bold,
    alignment,
    line_spacing: s.line_pt || (role === 'title' ? 35 : 28),
    line_rule: s.line_rule || 'exact',
    line_multiple: s.line_multiple || 1,
    first_indent: s.first_line_chars || 0,
  }
}

const fontForm = ref({ font_east: '仿宋_GB2312', font_west: 'Times New Roman', size: 16, bold: false })
const paragraphForm = ref({ role: 'body', alignment: 'justify', line_rule: 'exact', line_spacing: 28, line_multiple: 1, first_indent: 2 })

const selectionText = computed(() => {
  const text = props.selection?.text || ''
  return text.length > 42 ? text.slice(0, 42) + '...' : text
})

function applyRoleDefaults(role) {
  const d = styleToDefault(role) || { font_east: '仿宋_GB2312', font_west: 'Times New Roman', size: 16, bold: false, alignment: 'justify', line_rule: 'exact', line_spacing: 28, line_multiple: 1, first_indent: role === 'body' ? 2 : 0 }
  fontForm.value = { font_east: d.font_east, font_west: d.font_west, size: d.size, bold: d.bold }
  paragraphForm.value = {
    role,
    alignment: d.alignment,
    line_rule: d.line_rule,
    line_spacing: d.line_spacing,
    line_multiple: d.line_multiple,
    first_indent: d.first_indent,
  }
}

watch([() => props.selection, () => props.styles], ([selection]) => {
  if (selection?.role) applyRoleDefaults(selection.role)
})

function emitApply() {
  emit('applyEdit', { font: { ...fontForm.value }, paragraph: { ...paragraphForm.value } })
}

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
.edit-panel h3 { font-size: 15px; margin: 0 0 10px; color: #1a3a5c; }
.selection-box { min-height: 34px; max-height: 74px; overflow: hidden; padding: 8px; border: 1px solid #cfe0f2; border-radius: 6px; background: #f4f9ff; color: #345; font-size: 12px; line-height: 1.5; word-break: break-all; }
.selection-empty { padding: 10px 8px; border: 1px dashed #d9e2ec; border-radius: 6px; color: #9aa8b6; font-size: 12px; text-align: center; }
.edit-group { display: grid; gap: 6px; margin: 4px 0 8px; }
.edit-group.disabled { opacity: 0.58; }
.edit-group label { color: #607080; font-size: 12px; }
.edit-title { margin: 10px 0 4px; color: #1a3a5c; font-size: 13px; font-weight: 700; }
</style>
