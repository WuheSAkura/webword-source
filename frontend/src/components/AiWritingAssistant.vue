<template>
  <div class="ai-assistant-shell" :style="assistantStyle">
    <button
      v-if="!open"
      class="assistant-launcher"
      type="button"
      @click="handleLauncherClick"
      @pointerdown="startLauncherDrag"
    >
      公文写作
    </button>

    <div v-else class="ai-container-wrapper">
      <div class="assistant-topbar">
        <div class="main-heading">公文写作助手</div>
        <button class="assistant-close" type="button" title="收起" @click="open = false">收起</button>
      </div>

      <div class="ai-container" :class="{ expanded: messages.length || generating || result }">
        <div v-if="messages.length || generating || result" class="chat-panel">
          <div v-for="message in messages" :key="message.id" class="chat-message" :class="message.role">
            <div class="message-label">{{ message.role === 'user' ? '用户要求' : '处理结果' }}</div>
            <div class="message-body">{{ message.text }}</div>
          </div>

          <div v-if="generating" class="chat-message assistant">
            <div class="message-label">处理状态</div>
            <div class="message-body">正在理解意图、匹配公文模板并生成正式文件...</div>
          </div>

          <div v-if="result" class="result-card">
            <div class="result-header">
              <div>
                <div class="result-title">已生成文件</div>
                <div class="result-subtitle">
                  共 {{ result.documents?.length || 1 }} 份 · 模型：{{ modelName || '默认模型' }}
                </div>
              </div>
            </div>
            <div v-for="document in result.documents || [result]" :key="document.documentId" class="file-result-row">
              <div>
                <span>{{ document.filename }}</span>
                <small>模板：{{ document.templateName || document.templateId || '自动判断' }}</small>
              </div>
              <button class="download-button" type="button" @click="handleDownload(document)">下载</button>
            </div>
            <pre class="result-preview">{{ result.preview }}</pre>
          </div>
        </div>

        <textarea
          v-model="prompt"
          class="ai-input"
          rows="2"
          placeholder="请输入写作要求，例如：根据上传材料起草一份关于加强食品安全检查工作的通知"
          @keydown.ctrl.enter.prevent="handleGenerate"
        />

        <div v-if="files.length" class="selected-files">
          <span v-for="file in files" :key="file.name + file.size" class="selected-file">
            {{ file.name }}
          </span>
        </div>

        <div class="button-group">
          <div class="suggest-btn-wrapper">
            <button type="button" class="ai-button suggest" title="上传公文模板" @click="triggerTemplateUpload">
              <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" height="24" width="24">
                <path
                  fill="currentColor"
                  d="M2 12C2 6.47715 6.47715 2 12 2C17.5228 2 22 6.47715 22 12C22 17.5228 17.5228 22 12 22C6.47715 22 2 17.5228 2 12ZM11.9851 4.00291C11.9933 4.00046 11.9982 4.00006 11.9996 4C12.001 4.00006 12.0067 4.00046 12.0149 4.00291C12.0256 4.00615 12.047 4.01416 12.079 4.03356C12.2092 4.11248 12.4258 4.32444 12.675 4.77696C12.9161 5.21453 13.1479 5.8046 13.3486 6.53263C13.6852 7.75315 13.9156 9.29169 13.981 11H10.019C10.0844 9.29169 10.3148 7.75315 10.6514 6.53263C10.8521 5.8046 11.0839 5.21453 11.325 4.77696C11.5742 4.32444 11.7908 4.11248 11.921 4.03356C11.953 4.01416 11.9744 4.00615 11.9851 4.00291ZM8.01766 11C8.08396 9.13314 8.33431 7.41167 8.72334 6.00094C8.87366 5.45584 9.04762 4.94639 9.24523 4.48694C6.48462 5.49946 4.43722 7.9901 4.06189 11H8.01766ZM4.06189 13H8.01766C8.09487 15.1737 8.42177 17.1555 8.93 18.6802C9.02641 18.9694 9.13134 19.2483 9.24522 19.5131C6.48461 18.5005 4.43722 16.0099 4.06189 13ZM10.019 13H13.981C13.9045 14.9972 13.6027 16.7574 13.1726 18.0477C12.9206 18.8038 12.6425 19.3436 12.3823 19.6737C12.2545 19.8359 12.1506 19.9225 12.0814 19.9649C12.0485 19.9852 12.0264 19.9935 12.0153 19.9969C12.0049 20.0001 11.9999 20 11.9999 20C11.9999 20 11.9948 20 11.9847 19.9969C11.9736 19.9935 11.9515 19.9852 11.9186 19.9649C11.8494 19.9225 11.7455 19.8359 11.6177 19.6737C11.3575 19.3436 11.0794 18.8038 10.8274 18.0477C10.3973 16.7574 10.0955 14.9972 10.019 13ZM15.9823 13C15.9051 15.1737 15.5782 17.1555 15.07 18.6802C14.9736 18.9694 14.8687 19.2483 14.7548 19.5131C17.5154 18.5005 19.5628 16.0099 19.9381 13H15.9823ZM19.9381 11C19.5628 7.99009 17.5154 5.49946 14.7548 4.48694C14.9524 4.94639 15.1263 5.45584 15.2767 6.00094C15.6657 7.41167 15.916 9.13314 15.9823 11H19.9381Z"
                  clip-rule="evenodd"
                  fill-rule="evenodd"
                />
              </svg>
                <span class="search-text">上传模板</span>
            </button>
            <input
              ref="templateUploadInput"
              type="file"
              multiple
              accept=".docx,.txt,.wps"
              hidden
              @change="handleTemplateFiles"
            />
          </div>

          <div class="option-group">
            <label for="ai-file-input" class="ai-button voice" title="上传文件">
              <input id="ai-file-input" type="file" multiple accept=".docx,.txt" @change="handleFiles" />
              <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="ai-icon">
                <path d="M9.5 4C8.67157 4 8 4.67157 8 5.5V18.5C8 19.3284 8.67157 20 9.5 20C10.3284 20 11 19.3284 11 18.5V5.5C11 4.67157 10.3284 4 9.5 4Z" fill="currentColor" class="bar bar-1" />
                <path d="M13 8.5C13 7.67157 13.6716 7 14.5 7C15.3284 7 16 7.67157 16 8.5V15.5C16 16.3284 15.3284 17 14.5 17C13.6716 17 13 16.3284 13 15.5V8.5Z" fill="currentColor" class="bar bar-2" />
                <path d="M4.5 9C3.67157 9 3 9.67157 3 10.5V13.5C3 14.3284 3.67157 15 4.5 15C5.32843 15 6 14.3284 6 13.5V10.5C6 9.67157 5.32843 9 4.5 9Z" fill="currentColor" class="bar bar-3" />
                <path d="M19.5 9C18.6716 9 18 9.67157 18 10.5V13.5C18 14.3284 18.6716 15 19.5 15C20.3284 15 21 14.3284 21 13.5V10.5C21 9.67157 20.3284 9 19.5 9Z" fill="currentColor" class="bar bar-4" />
              </svg>
            </label>
            <button type="button" class="ai-button image" title="模型服务配置" @click="showSettings = !showSettings">
              <span class="tool-text">参</span>
            </button>
            <button type="button" class="ai-button camera" title="定位到模板选择" @click="focusTemplateSelect">
              <span class="tool-text">模</span>
            </button>
            <button type="button" class="ai-button submit" :disabled="generating" @click="handleGenerate">
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true" class="ai-icon small">
                <path
                  fill-rule="evenodd"
                  d="M5 4a.75.75 0 0 1 .738.616l.252 1.388A1.25 1.25 0 0 0 6.996 7.01l1.388.252a.75.75 0 0 1 0 1.476l-1.388.252A1.25 1.25 0 0 0 5.99 9.996l-.252 1.388a.75.75 0 0 1-1.476 0L4.01 9.996A1.25 1.25 0 0 0 3.004 8.99l-1.388-.252a.75.75 0 0 1 0-1.476l1.388-.252A1.25 1.25 0 0 0 4.01 6.004l.252-1.388A.75.75 0 0 1 5 4ZM12 1a.75.75 0 0 1 .721.544l.195.682c.118.415.443.74.858.858l.682.195a.75.75 0 0 1 0 1.442l-.682.195a1.25 1.25 0 0 0-.858.858l-.195.682a.75.75 0 0 1-1.442 0l-.195-.682a1.25 1.25 0 0 0-.858-.858l-.682-.195a.75.75 0 0 1 0-1.442l.682-.195a1.25 1.25 0 0 0 .858-.858l.195-.682A.75.75 0 0 1 12 1ZM10 11a.75.75 0 0 1 .728.568.968.968 0 0 0 .704.704.75.75 0 0 1 0 1.456.968.968 0 0 0-.704.704.75.75 0 0 1-1.456 0 .968.968 0 0 0-.704-.704.75.75 0 0 1 0-1.456.968.968 0 0 0 .704-.704A.75.75 0 0 1 10 11Z"
                  clip-rule="evenodd"
                />
              </svg>
              {{ generating ? '生成中' : '发送' }}
            </button>
          </div>
        </div>

        <div ref="templatePanelRef" class="config-panel template-panel">
          <label>公文模板</label>
          <div class="template-radio-list">
            <label class="template-radio" :class="{ active: templateId === '' }">
              <input v-model="templateId" type="radio" value="" />
              <span>自动判断</span>
            </label>
            <label
              v-for="template in templates"
              :key="template.id + template.name"
              class="template-radio"
              :class="{ active: templateId === template.id }"
            >
              <input v-model="templateId" type="radio" :value="template.id" />
              <span>{{ template.name }}</span>
              <em>{{ template.label }} / {{ template.fileCount }}份</em>
            </label>
          </div>
          <small v-if="selectedTemplate" class="template-meta">
            已收录：{{ selectedTemplate.sampleFile || '无样例文件' }}
          </small>
        </div>

        <div v-if="showSettings" class="config-panel">
          <label>模型服务地址</label>
          <input v-model="requestUrl" class="ai-setting-input" placeholder="例如：https://api.deepseek.com 或内网 /v1 地址" />
          <label>API Key</label>
          <input v-model="apiKey" class="ai-setting-input" type="password" placeholder="留空则使用后端 .env 配置" />
          <label>模型选择</label>
          <div class="model-row">
            <select v-model="modelName" class="ai-select">
              <option v-for="model in modelOptions" :key="model" :value="model">{{ model }}</option>
            </select>
            <button type="button" class="save-config-button" @click="saveModelSettings">保存</button>
          </div>
          <input
            v-model="customModelName"
            class="ai-setting-input"
            placeholder="新增模型名称，回车加入列表"
            @keydown.enter.prevent="addCustomModel"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed, onBeforeUnmount, onMounted, ref } from 'vue'
import { ElMessage } from 'element-plus'
import {
  downloadAiDocument,
  generateAiDocument,
  getAiModelConfig,
  getAiTemplates,
  saveAiModelConfig,
  syncAiTemplates,
  uploadAiTemplates,
} from '../api/index.js'

const emit = defineEmits(['generated'])
const open = ref(false)
const prompt = ref('')
const files = ref([])
const templates = ref([])
const templateId = ref('')
const requestUrl = ref('')
const apiKey = ref('')
const modelName = ref('')
const customModelName = ref('')
const modelOptions = ref([])
const showSettings = ref(false)
const generating = ref(false)
const result = ref(null)
const messages = ref([])
const templatePanelRef = ref(null)
const templateUploadInput = ref(null)
const selectedTemplate = computed(() => templates.value.find(item => item.id === templateId.value) || null)
const launcherPosition = ref(null)
const dragState = ref(null)
const suppressLauncherClick = ref(false)
const assistantStyle = computed(() => {
  if (open.value || !launcherPosition.value) return {}
  return {
    left: `${launcherPosition.value.left}px`,
    top: `${launcherPosition.value.top}px`,
    right: 'auto',
    bottom: 'auto',
  }
})

onMounted(async () => {
  await Promise.all([refreshTemplates(false), loadModelSettings()])
})

onBeforeUnmount(() => {
  stopLauncherDrag()
})

function handleLauncherClick() {
  if (suppressLauncherClick.value) {
    suppressLauncherClick.value = false
    return
  }
  open.value = true
}

function startLauncherDrag(event) {
  if (event.button !== 0) return
  const rect = event.currentTarget.getBoundingClientRect()
  dragState.value = {
    pointerId: event.pointerId,
    startX: event.clientX,
    startY: event.clientY,
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height,
    moved: false,
  }
  event.currentTarget.setPointerCapture?.(event.pointerId)
  window.addEventListener('pointermove', moveLauncher)
  window.addEventListener('pointerup', stopLauncherDrag)
}

function moveLauncher(event) {
  const state = dragState.value
  if (!state) return
  const dx = event.clientX - state.startX
  const dy = event.clientY - state.startY
  if (Math.abs(dx) + Math.abs(dy) > 4) state.moved = true
  const margin = 8
  launcherPosition.value = {
    left: Math.min(Math.max(margin, state.left + dx), window.innerWidth - state.width - margin),
    top: Math.min(Math.max(margin, state.top + dy), window.innerHeight - state.height - margin),
  }
}

function stopLauncherDrag() {
  if (dragState.value?.moved) suppressLauncherClick.value = true
  dragState.value = null
  window.removeEventListener('pointermove', moveLauncher)
  window.removeEventListener('pointerup', stopLauncherDrag)
}

async function refreshTemplates(showMessage = true) {
  try {
    const res = showMessage ? await syncAiTemplates() : await getAiTemplates()
    templates.value = res.data.templates || []
    if (showMessage) ElMessage.success(`已收录 ${templates.value.length} 类模板`)
  } catch (err) {
    ElMessage.warning('模板索引读取失败：' + (err.response?.data?.detail || err.message))
  }
}

function focusTemplateSelect() {
  templatePanelRef.value?.scrollIntoView({ behavior: 'smooth', block: 'nearest' })
}

function triggerTemplateUpload() {
  if (!templateId.value) {
    ElMessage.warning('请先选择要接入的公文模板类别')
    focusTemplateSelect()
    return
  }
  templateUploadInput.value?.click()
}

async function handleTemplateFiles(event) {
  const selectedFiles = Array.from(event.target.files || [])
  event.target.value = ''
  if (!selectedFiles.length) return
  if (!templateId.value) {
    ElMessage.warning('请先选择要接入的公文模板类别')
    return
  }
  try {
    const res = await uploadAiTemplates({
      files: selectedFiles,
      templateId: templateId.value,
    })
    templates.value = res.data.templates || templates.value
    ElMessage.success(`已上传并收录 ${res.data.count || 0} 个模板文件`)
  } catch (err) {
    ElMessage.error('模板上传失败：' + (err.response?.data?.detail || err.message))
  }
}

async function loadModelSettings() {
  try {
    const res = await getAiModelConfig()
    requestUrl.value = res.data.requestUrl || ''
    apiKey.value = res.data.apiKey || ''
    modelName.value = res.data.modelName || ''
    modelOptions.value = res.data.models || []
  } catch (err) {
    ElMessage.warning('模型配置读取失败：' + (err.response?.data?.detail || err.message))
  }
}

function handleFiles(event) {
  files.value = Array.from(event.target.files || [])
  event.target.value = ''
}

function addCustomModel() {
  const name = customModelName.value.trim()
  if (!name) return
  if (!modelOptions.value.includes(name)) modelOptions.value.push(name)
  modelName.value = name
  customModelName.value = ''
}

async function saveModelSettings() {
  try {
    const res = await saveAiModelConfig({
      requestUrl: requestUrl.value,
      apiKey: apiKey.value,
      modelName: modelName.value,
      models: modelOptions.value,
    })
    modelOptions.value = res.data.models || modelOptions.value
    modelName.value = res.data.modelName || modelName.value
    ElMessage.success('模型服务配置已保存')
  } catch (err) {
    ElMessage.error('模型配置保存失败：' + (err.response?.data?.detail || err.message))
  }
}

async function handleGenerate() {
  if (!prompt.value.trim() && files.value.length === 0) {
    ElMessage.warning('请上传文件或填写写作要求')
    return
  }

  const userText = [
    prompt.value.trim() || '根据上传文件生成正式公文',
    files.value.length ? `已上传 ${files.value.length} 个文件` : '',
  ].filter(Boolean).join('\n')

  messages.value = [
    ...messages.value,
    { id: Date.now(), role: 'user', text: userText },
  ]
  generating.value = true
  result.value = null

  try {
    const res = await generateAiDocument({
      files: files.value,
      prompt: prompt.value,
      requestUrl: requestUrl.value,
      apiKey: apiKey.value,
      modelName: modelName.value,
      templateId: templateId.value,
    })
    result.value = res.data
    emit('generated', res.data)
    messages.value.push({
      id: Date.now() + 1,
      role: 'assistant',
      text: `已完成模板匹配和文件生成：${res.data.filename}`,
    })
    if (res.data.warnings?.length) ElMessage.warning(res.data.warnings[0])
    else ElMessage.success('公文文件已生成')
  } catch (err) {
    const detail = err.response?.data?.detail || err.message
    messages.value.push({ id: Date.now() + 2, role: 'assistant', text: `生成失败：${detail}` })
    ElMessage.error('生成失败：' + detail)
  } finally {
    generating.value = false
  }
}

async function handleDownload(generatedDoc = null) {
  const target = generatedDoc || result.value
  if (!target?.documentId) return
  try {
    const res = await downloadAiDocument(target.documentId)
    const url = window.URL.createObjectURL(new Blob([res.data]))
    const link = document.createElement('a')
    link.href = url
    link.download = target.filename || 'AI公文.docx'
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
  } catch (err) {
    ElMessage.error('下载失败：' + (err.response?.data?.detail || err.message))
  }
}
</script>

<style scoped>
.ai-assistant-shell {
  position: fixed;
  right: 22px;
  bottom: 28px;
  z-index: 80;
}

.assistant-launcher {
  border: none;
  border-radius: 999px;
  padding: 10px 18px;
  background: #000;
  color: #fff;
  font-size: 13px;
  font-weight: 800;
  cursor: grab;
  touch-action: none;
  user-select: none;
  box-shadow: 0 10px 24px rgba(15, 23, 42, 0.22);
}

.assistant-launcher:active {
  cursor: grabbing;
}

.ai-container-wrapper {
  display: flex;
  flex-direction: column;
  row-gap: 10px;
  width: min(640px, calc(100vw - 30px));
  max-height: calc(100vh - 56px);
  font-family: Poppins, "Microsoft YaHei", sans-serif;
}

.assistant-topbar {
  display: grid;
  grid-template-columns: 1fr auto;
  align-items: center;
  padding: 0 4px;
}

.main-heading {
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 600;
  font-size: 24px;
  color: #000000;
}

.assistant-close {
  min-width: 54px;
  height: 30px;
  border: none;
  border-radius: 999px;
  background: #e9eef5;
  color: #333;
  font-size: 12px;
  font-weight: 700;
  cursor: pointer;
}

.ai-container {
  display: flex;
  flex-direction: column;
  row-gap: 8px;
  background: #fff;
  padding: 16px;
  border-radius: 18px;
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
  text-align: center;
  max-width: 100%;
  max-height: calc(100vh - 100px);
  overflow: auto;
  flex-shrink: 0;
}

.ai-container.expanded {
  row-gap: 12px;
}

.chat-panel {
  display: flex;
  flex-direction: column;
  row-gap: 10px;
  max-height: min(300px, 36vh);
  overflow: auto;
  padding: 14px;
  background: #e9ecef;
  border-radius: 14px;
  text-align: left;
}

.chat-message {
  max-width: 86%;
  padding: 10px 12px;
  border-radius: 16px;
  background: #fff;
  color: #1f2937;
}

.chat-message.user {
  margin-left: auto;
  background: #f4f8ff;
}

.chat-message.assistant {
  margin-right: auto;
}

.message-label {
  margin-bottom: 4px;
  color: #6b7280;
  font-size: 12px;
  font-weight: 700;
}

.message-body {
  font-size: 13px;
  line-height: 1.6;
  white-space: pre-wrap;
}

.result-card {
  padding: 12px;
  border-radius: 16px;
  background: #fff;
}

.result-header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 12px;
}

.result-title {
  color: #111827;
  font-size: 14px;
  font-weight: 800;
}

.result-subtitle {
  margin-top: 3px;
  color: #6b7280;
  font-size: 12px;
}

.file-result-row {
  display: flex;
  justify-content: space-between;
  gap: 10px;
  margin-top: 10px;
  padding: 8px 10px;
  border-radius: 12px;
  background: #f8fafc;
  color: #334155;
  font-size: 13px;
}

.result-preview {
  max-height: 150px;
  overflow: auto;
  margin-top: 10px;
  color: #334155;
  font: 13px/1.6 "Microsoft YaHei", sans-serif;
  white-space: pre-wrap;
}

.download-button,
.save-config-button {
  border: none;
  border-radius: 14px;
  background: #000;
  color: #fff;
  padding: 6px 12px;
  cursor: pointer;
}

.ai-input {
  width: 100%;
  min-height: 44px;
  padding: 10px 0;
  border: 0;
  font-size: 14px;
  outline: none;
  resize: vertical;
  max-height: 120px;
  transition: border-color 0.3s ease;
}

.selected-files {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.selected-file {
  max-width: 220px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  padding: 6px 10px;
  border-radius: 14px;
  background: #e9ecef;
  color: #475569;
  font-size: 12px;
}

.button-group,
.option-group {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
}

.option-group {
  column-gap: 8px;
}

.button-group {
  justify-content: space-between;
  gap: 10px;
}

.ai-button {
  display: flex;
  align-items: center;
  justify-content: center;
  column-gap: 4px;
  background: #000000;
  color: #fff;
  border: none;
  border-radius: 20px;
  padding: 10px;
  font-size: 12px;
  cursor: pointer;
  transition: background-color 0.3s ease, transform 0.3s ease;
  text-align: center;
  text-decoration: none;
}

.ai-button:hover {
  background: #46484b;
  transform: scale(1.05);
}

.ai-button:focus {
  background: #232426;
  transform: scale(0.9);
}

.ai-button:active {
  transform: scale(0.95);
}

.search-text {
  width: 0;
  overflow: hidden;
  transition: width 0.2s ease-in-out;
}

.ai-button.suggest {
  position: relative;
  z-index: 2;
  min-height: 38px;
  padding: 7px 10px;
  column-gap: 0;
  border-radius: 999px;
  background: #0084ff;
  overflow: hidden;
}

.ai-button.suggest:hover {
  background: hsl(209, 100%, 31%);
}

.ai-button.suggest .search-text {
  width: auto;
  padding-left: 4px;
}

.suggest-btn-wrapper {
  display: inline-block;
  position: relative;
  overflow: hidden;
  z-index: 1;
  padding: 0;
  border-radius: 999px;
}

.suggest-btn-wrapper:before {
  content: none;
  position: absolute;
  top: -1px;
  left: -1px;
  right: -1px;
  bottom: -1px;
  background: conic-gradient(from 4.67deg at 50% 50%, hsla(0, 0%, 100%, 0) -28.51deg, #0084ff 0.63deg, hsla(0, 0%, 100%, 0.22) 17.41deg, hsla(0, 0%, 100%, 0) 38.09deg, hsla(0, 0%, 100%, 0.24) 105.14deg, hsla(0, 0%, 100%, 0) 170.34deg, #0084ff 189.44deg, hsla(0, 0%, 100%, 0.08) 208.21deg, hsla(0, 0%, 100%, 0.43) 267.69deg, hsla(0, 0%, 100%, 0) 331.49deg, #0084ff 360.63deg);
  transform: scaleX(4);
  z-index: 1;
  animation: Button_rotate-gradient 5s linear infinite;
}

@keyframes Button_rotate-gradient {
  0% { transform: scale(9, 3) rotate(0); }
  50% { transform: scale(9, 3) rotate(180deg); }
  99.9999% { transform: scale(9, 3) rotate(1turn); }
  to { transform: scale(9, 3) rotate(0); }
}

.ai-button.voice,
.ai-button.image,
.ai-button.camera {
  justify-content: center;
  border-radius: 50%;
  height: 36px;
  width: 36px;
  padding: 0;
}

.ai-icon {
  flex-shrink: 0;
  height: 20px;
  width: 20px;
}

.ai-icon.small {
  height: 16px;
  width: 16px;
}

.tool-text {
  font-size: 13px;
  font-weight: 800;
}

.ai-button.voice {
  background: #6f42c1;
}

.ai-button.voice:hover {
  background: #5936a2;
}

.bar {
  transform-origin: center;
}

@keyframes wave {
  0%, 100% { transform: scaleY(1); }
  50% { transform: scaleY(1.8); }
}

@keyframes wave2 {
  0%, 100% { transform: scaleY(1); }
  50% { transform: scaleY(1.4); }
}

@keyframes wave3 {
  0%, 100% { transform: scaleY(1); }
  50% { transform: scaleY(1.6); }
}

@keyframes wave4 {
  0%, 100% { transform: scaleY(1); }
  50% { transform: scaleY(1.3); }
}

.ai-icon:hover .bar-1,
.ai-icon:focus .bar-1 {
  animation: wave 0.6s infinite ease-in-out alternate;
}

.ai-icon:hover .bar-2,
.ai-icon:focus .bar-2 {
  animation: wave2 0.6s infinite ease-in-out alternate;
}

.ai-icon:hover .bar-3,
.ai-icon:focus .bar-3 {
  animation: wave3 0.6s infinite ease-in-out alternate;
}

.ai-icon:hover .bar-4,
.ai-icon:focus .bar-4 {
  animation: wave4 0.6s infinite ease-in-out alternate;
}

.ai-button.image {
  background: #fd7e14;
}

.ai-button.image:hover {
  background: #e06d0f;
}

.ai-button.camera {
  background: rgb(32, 201, 151);
}

.ai-button.camera:hover {
  background: #1aa383;
}

#ai-file-input {
  display: none;
}

.config-panel {
  display: grid;
  row-gap: 8px;
  padding: 12px;
  background: #eef2f6;
  border-radius: 14px;
  text-align: left;
}

.config-panel label {
  color: #475569;
  font-size: 12px;
  font-weight: 800;
}

.template-panel {
  row-gap: 9px;
}

.template-radio-list {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 7px;
}

.template-radio {
  display: grid;
  grid-template-columns: 14px minmax(0, 1fr);
  grid-template-rows: auto auto;
  align-items: center;
  column-gap: 6px;
  min-height: 44px;
  padding: 7px 9px;
  border: 1px solid #d8e0e8;
  border-radius: 8px;
  background: #fff;
  color: #26384d;
  cursor: pointer;
}

.template-radio.active {
  border-color: #0084ff;
  background: #eef6ff;
}

.template-radio input {
  grid-row: 1 / span 2;
  margin: 0;
}

.template-radio span {
  overflow: hidden;
  font-size: 12px;
  font-weight: 700;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.template-radio em {
  overflow: hidden;
  color: #64748b;
  font-size: 10px;
  font-style: normal;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.template-meta {
  overflow: hidden;
  color: #64748b;
  font-size: 12px;
  line-height: 1.45;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.model-row {
  display: grid;
  grid-template-columns: 1fr auto;
  gap: 8px;
}

.ai-select,
.ai-setting-input {
  width: 100%;
  border: 0;
  border-radius: 16px;
  padding: 8px 10px;
  outline: none;
}

@media (max-width: 640px) {
  .ai-assistant-shell {
    right: 15px;
    left: 15px;
  }

  .ai-container-wrapper {
    width: 100%;
  }

  .template-radio-list {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }

  .button-group {
    row-gap: 10px;
  }
}
</style>
