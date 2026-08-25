<template>
  <div class="preview-panel">
    <div class="preview-summary" v-if="resultData || error">
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
        <el-radio-group v-model="previewMode" size="small" @change="switchPreviewMode">
          <el-radio-button value="pdf">PDF 真渲染</el-radio-button>
          <el-radio-button value="docx">浏览器预览</el-radio-button>
        </el-radio-group>
        <el-button size="small" :loading="docxLoading || renderChecking" @click="reloadPreview">刷新</el-button>
        <el-button size="small" type="primary" :disabled="!previewObjectUrl" @click="downloadPreviewDocx">下载DOCX</el-button>
        <el-button size="small" type="success" :disabled="!pdfObjectUrl" @click="downloadPdf">下载PDF</el-button>
      </div>
    </div>

    <div v-if="resultData?.paragraphs?.length" class="edit-toolbar">
      <label>段落</label>
      <el-select v-model="editParagraphIndex" size="small" style="width: 220px;">
        <el-option
          v-for="p in resultData.paragraphs"
          :key="p.index"
          :label="`${p.index + 1}. ${p.text}`"
          :value="p.index"
        />
      </el-select>
      <el-checkbox v-model="editBold">加粗</el-checkbox>
      <el-button size="small" :loading="editLoading" @click="applyLocalEdit">应用字体</el-button>
      <el-button size="small" :loading="undoLoading" @click="undoLocalEdit">撤销</el-button>
    </div>

    <div
      class="preview-body"
      v-loading="previewBusy"
      element-loading-text="预览加载中..."
      element-loading-background="rgba(255, 255, 255, 0.72)"
    >
      <el-empty v-if="!selectedFile" description="选择左侧文件即可预览" :image-size="90" />
      <el-empty v-else-if="loading" description="加载中..." :image-size="60" />
      <el-empty v-else-if="error && !resultData" :description="error" :image-size="60" />
      <el-empty v-else-if="!resultData || !resultData.paragraphs" description="尚未套用格式" :image-size="60" />
      <div v-else class="docx-stage">
        <el-alert
          v-if="previewMode === 'docx'"
          title="浏览器预览仅供参考；默认已优先展示 LibreOffice 真实渲染 PDF。"
          type="info"
          show-icon
          :closable="false"
          class="preview-tip"
        />
        <el-alert
          v-if="error || previewError || renderMessage"
          :title="error || previewError || renderMessage"
          :type="previewError || error ? 'warning' : 'success'"
          show-icon
          :closable="false"
        />
        <iframe
          ref="previewFrame"
          class="docx-frame"
          :title="previewMode === 'pdf' ? 'PDF 真实渲染预览' : 'DOCX 预览'"
        />
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue'
import { ElMessage } from 'element-plus'
import { renderAsync } from 'docx-preview'
import {
  downloadFile, downloadRenderPdf, editFile, readBlobError, renderCheck, undoFile,
} from '../api'

const props = defineProps({
  selectedFile: Object,
  resultData: Object,
  loading: Boolean,
  error: { type: String, default: '' },
})

const emit = defineEmits(['updated'])

const roleLabels = {
  title: '标题', subtitle: '副标题', recipient: '发文对象', body: '正文', h1: '一级标题', h2: '二级标题',
  h3: '三级标题', h4: '四级标题', attachment_head: '附件头', attachment_other: '其他附件', sign_unit: '落款·单位',
  sign_date: '落款·日期', sign_contact: '落款·联系人', security: '涉密标识', other: '其他',
}

const previewFrame = ref(null)
const previewMode = ref('pdf')
const docxLoading = ref(false)
const renderChecking = ref(false)
const previewBusy = computed(() => props.loading || docxLoading.value || renderChecking.value)
const editLoading = ref(false)
const undoLoading = ref(false)
const previewError = ref('')
const renderMessage = ref('')
const previewObjectUrl = ref('')
const pdfObjectUrl = ref('')
const editParagraphIndex = ref(0)
const editBold = ref(false)
let renderRequestId = 0
let heightObserver = null

const paragraphCount = computed(() => props.resultData?.paragraphs?.length || 0)

const roleCounts = computed(() => {
  const map = new Map()
  for (const p of props.resultData?.paragraphs || []) map.set(p.role, (map.get(p.role) || 0) + 1)
  return Array.from(map.entries()).map(([role, count]) => ({ role, count, label: roleLabels[role] || role }))
})

watch(
  () => [props.selectedFile?.id, props.resultData?.paragraphs, props.loading],
  async () => {
    if (props.loading) return
    if (props.selectedFile?.id && props.resultData?.paragraphs) {
      editParagraphIndex.value = props.resultData.paragraphs[0]?.index ?? 0
      await nextTick()
      await reloadPreview()
    } else {
      clearPreview()
    }
  },
  { immediate: true },
)

function withTimeout(promise, timeoutMs, message) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), timeoutMs)
    promise.then(
      (value) => {
        clearTimeout(timer)
        resolve(value)
      },
      (error) => {
        clearTimeout(timer)
        reject(error)
      },
    )
  })
}

async function waitForPreviewFrame(maxAttempts = 24) {
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    await nextTick()
    const frame = previewFrame.value
    if (frame?.contentDocument) return frame
    await new Promise(resolve => setTimeout(resolve, 50))
  }
  throw new Error('预览容器初始化失败，请切换到“浏览器预览”后重试')
}

async function reloadPreview() {
  if (previewMode.value === 'pdf') await loadPdfPreview()
  else await renderDocxPreview()
}

function switchPreviewMode() {
  reloadPreview()
}

async function loadPdfPreview() {
  if (!props.selectedFile?.id) return
  const requestId = ++renderRequestId
  renderChecking.value = true
  previewError.value = ''
  renderMessage.value = ''
  try {
    const res = await withTimeout(
      renderCheck(props.selectedFile.id),
      45000,
      '真实渲染超时，已切换浏览器预览',
    )
    if (requestId !== renderRequestId) return
    if (!res.data?.success) {
      previewError.value = res.data?.message || res.data?.detail || '真实渲染不可用，已回退浏览器预览'
      previewMode.value = 'docx'
      await renderDocxPreview({ parentRequestId: requestId })
      return
    }
    const pdfRes = await withTimeout(
      downloadRenderPdf(props.selectedFile.id),
      45000,
      'PDF 下载超时，已切换浏览器预览',
    )
    if (requestId !== renderRequestId) return
    setPdfObjectUrl(URL.createObjectURL(new Blob([pdfRes.data], { type: 'application/pdf' })))
    await nextTick()
    showPdfInFrame(pdfObjectUrl.value)
    renderMessage.value = res.data.cached ? '真实渲染完成（缓存）' : '真实渲染完成'
    await ensureDocxBlob()
  } catch (error) {
    if (requestId !== renderRequestId) return
    previewError.value = await readBlobError(error, '真实渲染失败，已回退浏览器预览')
    previewMode.value = 'docx'
    await renderDocxPreview({ parentRequestId: requestId })
  } finally {
    renderChecking.value = false
  }
}

async function ensureDocxBlob() {
  if (previewObjectUrl.value) return
  try {
    const res = await downloadFile(props.selectedFile.id)
    const blob = new Blob([res.data], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    })
    setPreviewObjectUrl(URL.createObjectURL(blob))
  } catch {
    // 下载按钮仍可走服务端
  }
}

function showPdfInFrame(url) {
  const frame = previewFrame.value
  if (!frame) return
  heightObserver?.disconnect()
  heightObserver = null
  frame.removeAttribute('srcdoc')
  frame.src = url
  frame.style.height = '920px'
}

async function renderDocxPreview(options = {}) {
  if (!props.selectedFile?.id || !props.resultData?.paragraphs) return
  const ownsRequest = !options.parentRequestId
  const requestId = ownsRequest ? ++renderRequestId : options.parentRequestId
  if (ownsRequest) docxLoading.value = true
  previewError.value = ''
  try {
    await waitForPreviewFrame()
    if (requestId !== renderRequestId) return
    const root = prepareDocxFrame()
    const res = await withTimeout(
      downloadFile(props.selectedFile.id),
      45000,
      'DOCX 下载超时，请稍后重试',
    )
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
    bindPreviewHeightObserver()
    adjustPreviewFrameHeight()
  } catch (error) {
    if (requestId !== renderRequestId) return
    previewError.value = await readBlobError(
      error,
      error?.response?.status === 404
        ? '文件已过期，请重新处理或重新生成后再预览。'
        : '预览生成失败，请下载文件查看',
    )
  } finally {
    if (ownsRequest) docxLoading.value = false
  }
}

async function applyLocalEdit() {
  if (!props.selectedFile?.id) return
  const paragraph = props.resultData?.paragraphs?.find(p => p.index === editParagraphIndex.value)
  if (!paragraph) return
  editLoading.value = true
  try {
    const res = await editFile(props.selectedFile.id, {
      selection: {
        startParagraph: paragraph.index,
        endParagraph: paragraph.index,
        startOffset: 0,
        endOffset: (paragraph.fullText || paragraph.text || '').length,
      },
      font: { bold: editBold.value },
    })
    emit('updated', res.data)
    ElMessage.success('已应用局部修改')
    await reloadPreview()
  } catch (error) {
    ElMessage.error(await readBlobError(error, '局部修改失败'))
  } finally {
    editLoading.value = false
  }
}

async function undoLocalEdit() {
  if (!props.selectedFile?.id) return
  undoLoading.value = true
  try {
    const res = await undoFile(props.selectedFile.id)
    emit('updated', res.data)
    ElMessage.success('已撤销最近一次修改')
    await reloadPreview()
  } catch (error) {
    ElMessage.error(await readBlobError(error, '撤销失败'))
  } finally {
    undoLoading.value = false
  }
}

function downloadPreviewDocx() {
  if (!previewObjectUrl.value) {
    ElMessage.warning('请先刷新预览后再下载，或从右侧操作栏下载')
    return
  }
  triggerDownload(previewObjectUrl.value, buildDocxName())
}

function downloadPdf() {
  if (!pdfObjectUrl.value) return
  const stem = (props.selectedFile?.name || '公文').replace(/\.docx$/i, '')
  triggerDownload(pdfObjectUrl.value, `${stem}_真实渲染预览.pdf`)
}

function buildDocxName() {
  const raw = props.selectedFile?.name || '公文.docx'
  const stem = raw.replace(/\.docx$/i, '')
  return (stem.endsWith('_AI公文') || stem.endsWith('_公文格式'))
    ? `${stem}.docx`
    : `${stem}_公文格式.docx`
}

function triggerDownload(url, filename) {
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
}

function setPreviewObjectUrl(url) {
  if (previewObjectUrl.value?.startsWith('blob:')) URL.revokeObjectURL(previewObjectUrl.value)
  previewObjectUrl.value = url
}

function setPdfObjectUrl(url) {
  if (pdfObjectUrl.value?.startsWith('blob:')) URL.revokeObjectURL(pdfObjectUrl.value)
  pdfObjectUrl.value = url
}

function clearPreview(clearError = true, invalidateRequest = true) {
  if (invalidateRequest) renderRequestId += 1
  heightObserver?.disconnect()
  heightObserver = null
  renderChecking.value = false
  docxLoading.value = false
  resetPreviewFrame()
  setPreviewObjectUrl('')
  setPdfObjectUrl('')
  if (clearError) {
    previewError.value = ''
    renderMessage.value = ''
  }
}

function prepareDocxFrame() {
  const frame = previewFrame.value
  const doc = frame?.contentDocument
  if (!doc) throw new Error('预览容器初始化失败')
  frame.removeAttribute('src')
  frame.src = 'about:blank'
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
      .docx { margin: 0 !important; box-shadow: 0 16px 38px rgba(31, 41, 55, 0.16) !important; }
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
  frame.style.height = `${Math.max(body.scrollHeight + 24, 920)}px`
}

function bindPreviewHeightObserver() {
  heightObserver?.disconnect()
  const body = previewFrame.value?.contentDocument?.body
  if (!body || typeof ResizeObserver === 'undefined') return
  heightObserver = new ResizeObserver(() => adjustPreviewFrameHeight())
  heightObserver.observe(body)
}

onBeforeUnmount(() => clearPreview())
</script>

<style scoped>
.preview-panel { display: flex; flex-direction: column; height: 100%; min-height: 0; }
.preview-summary {
  display: flex; align-items: center; gap: 14px; padding: 10px 18px;
  border-bottom: 1px solid #edf2f7; background: #f6f9fc; flex-shrink: 0; flex-wrap: wrap;
}
.summary-item { display: flex; align-items: baseline; gap: 8px; min-width: 86px; }
.summary-item span { color: #7a8997; font-size: 12px; }
.summary-item strong { color: #1a3a5c; font-size: 20px; line-height: 1; }
.summary-roles { display: flex; flex-wrap: wrap; gap: 6px; }
.summary-actions { display: flex; gap: 8px; margin-left: auto; flex-shrink: 0; flex-wrap: wrap; align-items: center; }
.summary-chip { padding: 2px 8px; border-radius: 999px; background: #edf2f7; color: #475569; font-size: 12px; line-height: 20px; }
.edit-toolbar {
  display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
  padding: 8px 18px; border-bottom: 1px solid #edf2f7; background: #fff;
}
.edit-toolbar label { font-size: 12px; color: #64748b; font-weight: 700; }
.preview-body { flex: 1; overflow-y: auto; background: #e9eef5; position: relative; min-height: 0; }
.docx-stage { min-height: 100%; padding: 18px; }
.preview-tip { margin-bottom: 10px; }
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
