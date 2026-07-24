<template>
  <div class="preview-panel">
    <div class="preview-summary" v-if="resultData">
      <div class="summary-item">
        <span>逻辑段落</span>
        <strong>{{ paragraphCount }}</strong>
      </div>
      <div class="summary-roles">
        <span v-for="item in roleCounts" :key="item.role" class="summary-chip" :class="'role-' + item.role">
          {{ item.label }} {{ item.count }}
        </span>
      </div>
      <div class="summary-actions">
        <el-button size="small" :loading="docxLoading" @click="renderPreview">刷新预览</el-button>
        <el-button size="small" type="primary" :disabled="!previewObjectUrl" @click="downloadPreviewDocx">下载DOCX</el-button>
      </div>
    </div>

    <div class="preview-body" v-loading="loading || docxLoading">
      <el-empty v-if="!selectedFile" description="选择左侧文件即可预览" :image-size="90" />
      <el-empty v-else-if="loading" description="加载中..." :image-size="60" />
      <el-empty v-else-if="!resultData || !resultData.paragraphs" description="尚未套用格式" :image-size="60" />
      <div v-else class="docx-stage">
        <el-alert
          v-if="previewError"
          :title="previewError"
          type="warning"
          show-icon
          :closable="false"
        />
        <iframe ref="previewFrame" class="docx-frame" title="DOCX 预览" />
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import { renderAsync } from 'docx-preview'
import { downloadFile } from '../api'

const props = defineProps({
  selectedFile: Object,
  resultData: Object,
  loading: Boolean,
})

const roleLabels = {
  title: '标题', subtitle: '副标题', recipient: '发文对象', body: '正文', h1: '一级标题', h2: '二级标题',
  h3: '三级标题', h4: '四级标题', attachment_head: '附件头', attachment_other: '其他附件', sign_unit: '落款·单位',
  sign_date: '落款·日期', sign_contact: '落款·联系人', security: '涉密标识', other: '其他',
}

const previewFrame = ref(null)
const docxLoading = ref(false)
const previewError = ref('')
const previewObjectUrl = ref('')
let renderRequestId = 0

const paragraphCount = computed(() => props.resultData?.paragraphs?.length || 0)

const roleCounts = computed(() => {
  const map = new Map()
  for (const p of props.resultData?.paragraphs || []) map.set(p.role, (map.get(p.role) || 0) + 1)
  return Array.from(map.entries()).map(([role, count]) => ({ role, count, label: roleLabels[role] || role }))
})

watch(
  () => [props.selectedFile?.id, props.resultData?.paragraphs],
  () => {
    if (props.selectedFile?.id && props.resultData?.paragraphs) renderPreview()
    else clearPreview()
  },
  { immediate: true },
)

async function renderPreview() {
  if (!props.selectedFile?.id || !props.resultData?.paragraphs) return
  const requestId = ++renderRequestId
  docxLoading.value = true
  previewError.value = ''
  await nextTick()
  const root = preparePreviewFrame()

  try {
    const res = await downloadFile(props.selectedFile.id)
    if (requestId !== renderRequestId) return
    const blob = new Blob([res.data], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    })
    setPreviewObjectUrl(URL.createObjectURL(blob))
    await renderAsync(blob, root, root, {
      className: 'docx',
      inWrapper: true,
      breakPages: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      renderComments: false,
      experimental: true,
    })
    adjustPreviewFrameHeight()
    setTimeout(adjustPreviewFrameHeight, 200)
    setTimeout(adjustPreviewFrameHeight, 800)
  } catch (error) {
    if (requestId !== renderRequestId) return
    clearPreview(false, false)
    previewError.value = error?.response?.data?.detail || error?.message || '预览生成失败，请下载文件查看'
  } finally {
    if (requestId === renderRequestId) docxLoading.value = false
  }
}

function downloadPreviewDocx() {
  if (!previewObjectUrl.value) return
  const link = document.createElement('a')
  link.href = previewObjectUrl.value
  link.download = props.selectedFile?.name?.replace(/\.docx$/i, '_公文格式.docx') || '公文格式.docx'
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
}

function setPreviewObjectUrl(url) {
  if (previewObjectUrl.value?.startsWith('blob:')) URL.revokeObjectURL(previewObjectUrl.value)
  previewObjectUrl.value = url
}

function clearPreview(clearError = true, invalidateRequest = true) {
  if (invalidateRequest) renderRequestId += 1
  resetPreviewFrame()
  setPreviewObjectUrl('')
  if (clearError) previewError.value = ''
}

function preparePreviewFrame() {
  const frame = previewFrame.value
  const doc = frame?.contentDocument
  if (!doc) throw new Error('预览容器初始化失败')
  doc.open()
  doc.write(`<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <style>
      html, body { margin: 0; min-height: 100%; background: #e9eef5; }
      body {
        font-family: "Times New Roman", "仿宋_GB2312", "FangSong_GB2312", "FangSong", "SimSun", serif;
        -webkit-font-smoothing: antialiased;
        text-rendering: geometricPrecision;
      }
      .preview-root { min-height: 100%; padding: 18px; box-sizing: border-box; }
      .docx-wrapper {
        background: transparent !important;
        padding: 0 !important;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        gap: 18px !important;
      }
      .docx {
        margin: 0 !important;
        box-shadow: 0 16px 38px rgba(31, 41, 55, 0.16) !important;
      }
      @media (max-width: 860px) {
        .preview-root { padding: 10px; overflow-x: auto; }
      }
    </style>
  </head>
  <body><div id="docx-root" class="preview-root"></div></body>
</html>`)
  doc.close()
  const root = doc.getElementById('docx-root')
  if (!root) throw new Error('预览容器初始化失败')
  return root
}

function resetPreviewFrame() {
  const frame = previewFrame.value
  if (!frame) return
  frame.removeAttribute('src')
  frame.style.height = '920px'
  const doc = frame.contentDocument
  if (doc?.body) doc.body.innerHTML = ''
}

function adjustPreviewFrameHeight() {
  const frame = previewFrame.value
  const body = frame?.contentDocument?.body
  if (!frame || !body) return
  frame.style.height = `${Math.max(body.scrollHeight, 920)}px`
}

onBeforeUnmount(() => clearPreview())
</script>

<style scoped>
.preview-panel { display: flex; flex-direction: column; height: 100%; min-height: 0; }
.preview-summary {
  display: flex; align-items: center; gap: 14px; padding: 10px 18px;
  border-bottom: 1px solid #edf2f7; background: #f6f9fc; flex-shrink: 0;
}
.summary-item { display: flex; align-items: baseline; gap: 8px; min-width: 86px; }
.summary-item span { color: #7a8997; font-size: 12px; }
.summary-item strong { color: #1a3a5c; font-size: 20px; line-height: 1; }
.summary-roles { display: flex; flex-wrap: wrap; gap: 6px; }
.summary-actions { display: flex; gap: 8px; margin-left: auto; flex-shrink: 0; }
.summary-chip { padding: 2px 8px; border-radius: 999px; background: #edf2f7; color: #475569; font-size: 12px; line-height: 20px; }
.preview-body { flex: 1; overflow-y: auto; background: #e9eef5; }
.docx-stage { min-height: 100%; padding: 18px; }
.docx-frame { width: 100%; min-height: 920px; border: 0; display: block; background: #e9eef5; }
.summary-chip.role-title { background: #fff0f0; color: #c53030; }
.summary-chip.role-subtitle { background: #fdf0ff; color: #97266d; }
.summary-chip.role-recipient { background: #eef7f0; color: #2f6b4f; }
.summary-chip.role-security { background: #edf2f7; color: #2d3748; }
.summary-chip.role-h1 { background: #fff4e8; color: #c05621; }
.summary-chip.role-h2 { background: #fff8db; color: #a36b00; }
.summary-chip.role-h3 { background: #e9f8ef; color: #25855a; }
.summary-chip.role-h4 { background: #e6fffa; color: #2c7a7b; }
.summary-chip.role-body { background: #eaf3ff; color: #2b6cb0; }
.summary-chip.role-attachment_head,
.summary-chip.role-attachment_other { background: #f4edff; color: #6b46c1; }
.summary-chip.role-sign_unit,
.summary-chip.role-sign_date,
.summary-chip.role-sign_contact { background: #edf2f7; color: #4a5568; }

@media (max-width: 860px) {
  .preview-summary { align-items: flex-start; flex-direction: column; }
  .summary-actions { margin-left: 0; width: 100%; }
  .docx-stage { padding: 10px; overflow-x: auto; }
  .docx-frame { min-width: 760px; }
}
</style>
