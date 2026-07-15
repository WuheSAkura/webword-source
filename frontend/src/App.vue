<template>
  <div class="app-wrapper">
    <header class="page-header">
      <h1>Word 公文格式转换工具</h1>
      <p>上传文档 → 识别结构 → 人工校正类型 → 套用格式 → 下载</p>
    </header>

    <div class="main-layout">
      <div class="panel panel-left">
        <FileList
          :files="fileList"
          :selectedId="selectedId"
          @upload="handleUpload"
          @select="handleSelect"
          @remove="handleRemove"
          @process="handleProcess"
          @download="handleDownload"
          @clearAll="handleClearAll"
        />
      </div>

      <div class="panel panel-center">
        <div class="center-tabs">
          <button class="tab" :class="{ active: activeTab === 'structure' }" @click="setActiveTab('structure')">
            识别校正
          </button>
          <button
            class="tab"
            :class="{ active: activeTab === 'result', disabled: !resultAvailable }"
            :disabled="!resultAvailable"
            @click="setActiveTab('result')"
          >
            处理后
          </button>
        </div>

        <StructurePanel
          v-show="activeTab === 'structure'"
          :selectedFile="selectedFile"
          :data="structureData"
          :labels="roleLabels"
          :loading="previewLoading"
          @changeRole="handleChangeRole"
        />
        <PreviewPanel
          v-show="activeTab === 'result'"
          :selectedFile="selectedFile"
          :resultData="resultData"
          :loading="previewLoading"
          @selectRange="handleSelectRange"
        />
      </div>

      <div class="panel panel-right">
        <ActionPanel
          :selectedFile="selectedFile"
          :isProcessing="isSelectedProcessing"
          :isCompleted="isSelectedCompleted"
          :totalCount="fileList.length"
          :completedCount="completedCount"
          :processingCount="processingCount"
          :pendingCount="pendingCount"
          :selection="selectionRange"
          :styles="roleStyles"
          :labels="roleLabels"
          @process="handleProcess"
          @download="handleDownload"
          @applyEdit="handleApplyEdit"
          @undoEdit="handleUndoEdit"
        />
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import FileList from './components/FileList.vue'
import StructurePanel from './components/StructurePanel.vue'
import PreviewPanel from './components/PreviewPanel.vue'
import ActionPanel from './components/ActionPanel.vue'
import {
  uploadFiles, getStructure, getFormatConfig, previewResult,
  convertFile, downloadFile, editFile, undoFile, clearAll,
} from './api/index.js'

const fileList = ref([])
const selectedId = ref(null)
const activeTab = ref('structure')
const structureData = ref(null)
const resultData = ref(null)
const previewLoading = ref(false)
const selectionRange = ref(null)
const roleStyles = ref({})
const roleLabels = ref({})

// 缓存每个文件的结构清单与结果预览 { fid: { structure, result } }
const cache = ref({})

const completedCount = computed(() => fileList.value.filter(f => f.status === 'completed').length)
const processingCount = computed(() => fileList.value.filter(f => f.status === 'processing').length)
const pendingCount = computed(() => fileList.value.filter(f => f.status === 'pending').length)

const selectedFile = computed(() => fileList.value.find(f => f.id === selectedId.value) || null)
const isSelectedCompleted = computed(() => selectedFile.value?.status === 'completed')
const isSelectedProcessing = computed(() => selectedFile.value?.status === 'processing')
const resultAvailable = computed(() => !!resultData.value || selectedFile.value?.status === 'completed')

onMounted(async () => {
  try {
    const res = await getFormatConfig()
    if (res.data.success) {
      roleStyles.value = res.data.styles || {}
      roleLabels.value = res.data.labels || {}
    }
  } catch { /* 用前端兜底标签 */ }
})

function getCache(fid) {
  if (!cache.value[fid]) cache.value[fid] = {}
  return cache.value[fid]
}

function refreshDisplay(fid) {
  const c = getCache(fid)
  structureData.value = c.structure || null
  resultData.value = c.result || null
}

async function loadStructure(fid) {
  previewLoading.value = true
  try {
    const res = await getStructure(fid)
    getCache(fid).structure = res.data
  } catch {
    getCache(fid).structure = null
  } finally {
    previewLoading.value = false
    if (selectedId.value === fid) refreshDisplay(fid)
  }
}

async function loadResultData(fid) {
  previewLoading.value = true
  try {
    const res = await previewResult(fid)
    getCache(fid).result = res.data
  } catch {
    getCache(fid).result = null
  } finally {
    previewLoading.value = false
    if (selectedId.value === fid) refreshDisplay(fid)
  }
}

function setActiveTab(tab) {
  if (tab === 'result' && !resultAvailable.value) return
  activeTab.value = tab
}

function setFileStatus(fid, status) {
  const f = fileList.value.find(x => x.id === fid)
  if (f) f.status = status
}

// ── handlers ──
async function handleUpload(files) {
  const docxFiles = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.docx'))
  if (docxFiles.length === 0) { ElMessage.warning('请选择 .docx 文件'); return }
  try {
    const res = await uploadFiles(docxFiles)
    if (res.data.success) {
      for (const item of res.data.files) fileList.value.push({ ...item, status: 'pending' })
      if (!selectedId.value && fileList.value.length > 0) await selectFile(fileList.value[0].id)
      ElMessage.success(`已添加 ${res.data.files.length} 个文件`)
    }
  } catch (err) {
    ElMessage.error('上传失败: ' + (err.response?.data?.detail || err.message))
  }
}

async function selectFile(fid) {
  selectedId.value = fid
  selectionRange.value = null
  activeTab.value = 'structure'
  if (!getCache(fid).structure) await loadStructure(fid)
  refreshDisplay(fid)
}

async function handleSelect(fid) { await selectFile(fid) }

async function handleRemove(fid) {
  fileList.value = fileList.value.filter(f => f.id !== fid)
  delete cache.value[fid]
  if (selectedId.value === fid) {
    selectedId.value = fileList.value[0]?.id || null
    structureData.value = null
    resultData.value = null
    selectionRange.value = null
    if (selectedId.value) await selectFile(selectedId.value)
  }
}

function rolesFromStructure(fid) {
  const struct = getCache(fid).structure
  if (!struct?.paragraphs) return null
  const roles = {}
  for (const p of struct.paragraphs) roles[String(p.index)] = p.role
  return roles
}

function handleChangeRole({ index, role }) {
  const fid = selectedId.value
  const struct = getCache(fid)?.structure
  if (!struct?.paragraphs) return
  const item = struct.paragraphs.find(p => p.index === index)
  if (item) { item.role = role; item.lowConfidence = false }
}

async function processOneFile(targetId, showMessage = true) {
  setFileStatus(targetId, 'processing')
  try {
    const roles = rolesFromStructure(targetId)
    const res = await convertFile(targetId, roles)
    if (res.data.success) {
      setFileStatus(targetId, 'completed')
      const pr = await previewResult(targetId)
      getCache(targetId).result = pr.data
      if (selectedId.value === targetId) {
        selectionRange.value = null
        activeTab.value = 'result'
        refreshDisplay(targetId)
      }
      if (showMessage) ElMessage.success('处理完成')
      return true
    }
  } catch (err) {
    setFileStatus(targetId, 'pending')
    if (showMessage) ElMessage.error('处理失败: ' + (err.response?.data?.detail || err.message))
    return false
  }
  return false
}

async function handleProcess(fid) {
  if (fid) { await processOneFile(fid); return }
  const targets = fileList.value.filter(f => f.status !== 'completed' && f.status !== 'processing')
  if (targets.length === 0) {
    ElMessage.warning(fileList.value.length ? '没有待处理文件' : '请先上传文件')
    return
  }
  let ok = 0, fail = 0
  for (const file of targets) (await processOneFile(file.id, false)) ? ok++ : fail++
  if (selectedId.value && getCache(selectedId.value).result) {
    activeTab.value = 'result'
    refreshDisplay(selectedId.value)
  }
  if (fail > 0) ElMessage.warning(`批量处理完成，成功 ${ok} 个，失败 ${fail} 个`)
  else ElMessage.success(`批量处理完成，共处理 ${ok} 个文件`)
}

async function handleDownload(fid) {
  const targetId = fid || selectedId.value
  if (!targetId) return
  try {
    const res = await downloadFile(targetId)
    const file = fileList.value.find(f => f.id === targetId)
    const name = file ? file.name.replace('.docx', '_公文格式.docx') : '公文格式.docx'
    const url = window.URL.createObjectURL(new Blob([res.data]))
    const link = document.createElement('a')
    link.href = url; link.download = name
    document.body.appendChild(link); link.click(); document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
    ElMessage.success('下载完成')
  } catch { ElMessage.error('下载失败') }
}

function handleSelectRange(selection) { selectionRange.value = selection }

async function handleApplyEdit(payload) {
  const targetId = selectedId.value
  if (!targetId || !selectionRange.value) {
    ElMessage.warning('请先在处理后预览中选中文字')
    return
  }
  try {
    const res = await editFile(targetId, {
      selection: selectionRange.value,
      font: payload.font,
      paragraph: payload.paragraph,
    })
    const c = getCache(targetId)
    c.result = { ...(c.result || { success: true, name: selectedFile.value?.name, type: 'processed' }), paragraphs: res.data.paragraphs }
    activeTab.value = 'result'
    refreshDisplay(targetId)
    selectionRange.value = null
    window.getSelection()?.removeAllRanges()
    ElMessage.success('修改已应用')
  } catch (err) {
    const detail = err.response?.data?.detail || err.message
    ElMessage.error(String(detail).includes('文件不存在')
      ? '应用修改失败：后端文件已过期，请重新上传并处理'
      : '应用修改失败: ' + detail)
  }
}

async function handleUndoEdit() {
  const targetId = selectedId.value
  if (!targetId) return
  try {
    const res = await undoFile(targetId)
    const c = getCache(targetId)
    c.result = { ...(c.result || { success: true, name: selectedFile.value?.name, type: 'processed' }), paragraphs: res.data.paragraphs }
    activeTab.value = 'result'
    refreshDisplay(targetId)
    selectionRange.value = null
    ElMessage.success('已撤销上一步修改')
  } catch (err) {
    ElMessage.error('撤销失败: ' + (err.response?.data?.detail || err.message))
  }
}

async function handleClearAll() {
  try { await clearAll() } catch {}
  fileList.value = []
  selectedId.value = null
  structureData.value = null
  resultData.value = null
  selectionRange.value = null
  cache.value = {}
}
</script>

<style scoped>
.app-wrapper { min-height: 100vh; }
.page-header { background: linear-gradient(135deg, #1a3a5c 0%, #2d6aa0 100%); padding: 16px 24px; text-align: center; box-shadow: 0 2px 12px rgba(0,0,0,0.12); }
.page-header h1 { color: #fff; font-size: 22px; font-weight: 700; }
.page-header p { color: rgba(255,255,255,0.7); font-size: 13px; margin-top: 4px; }
.main-layout { display: flex; gap: 0; height: calc(100vh - 80px); max-width: 1500px; margin: 0 auto; padding: 16px; }
.panel { background: #fff; border-radius: 12px; box-shadow: 0 1px 8px rgba(0,0,0,0.06); overflow: hidden; }
.panel-left { width: 280px; min-width: 280px; margin-right: 12px; display: flex; flex-direction: column; }
.panel-center { flex: 1; margin-right: 12px; display: flex; flex-direction: column; }
.panel-right { width: 280px; min-width: 280px; display: flex; flex-direction: column; }
.center-tabs { display: flex; gap: 4px; padding: 10px 14px; border-bottom: 1px solid #e6edf5; background: #fbfdff; flex-shrink: 0; }
.center-tabs .tab {
  min-width: 88px; min-height: 32px; padding: 0 14px; border: 1px solid #d8e5f2; border-radius: 6px;
  background: #fff; color: #738294; font-size: 13px; font-weight: 700; cursor: pointer; transition: all 0.2s;
}
.center-tabs .tab.active { color: #fff; background: #2d6aa0; border-color: #2d6aa0; }
.center-tabs .tab.disabled { color: #ccc; cursor: not-allowed; }
</style>
