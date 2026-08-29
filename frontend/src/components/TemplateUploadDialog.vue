<template>
  <el-dialog
    v-model="visibleProxy"
    title="模板管理"
    width="1260px"
    top="4vh"
    append-to-body
    destroy-on-close
    class="template-manager-dialog"
    @open="handleOpen"
  >
    <div v-loading="loadingTemplates" class="upload-dialog-body">
      <div class="mode-row">
        <label class="mode-option" :class="{ active: categoryMode === 'existing' }">
          <input v-model="categoryMode" type="radio" value="existing" />
          <span>归入已有</span>
        </label>
        <label class="mode-option" :class="{ active: categoryMode === 'new' }">
          <input v-model="categoryMode" type="radio" value="new" />
          <span>新建分类</span>
        </label>
        <label class="mode-option" :class="{ active: categoryMode === 'manage' }">
          <input v-model="categoryMode" type="radio" value="manage" />
          <span>管理删除</span>
        </label>
      </div>

      <template v-if="categoryMode === 'existing'">
        <label>选择模板分类</label>
        <select v-model="selectedTemplateKey" class="field-select">
          <option value="">请选择已有分类</option>
          <option
            v-for="item in localTemplates"
            :key="item.templateKey"
            :value="item.templateKey"
          >
            {{ item.name }}（{{ item.label }} / {{ item.fileCount }} 份）
          </option>
        </select>
      </template>

      <template v-else-if="categoryMode === 'new'">
        <label>模板分类名称</label>
        <input
          v-model="categoryName"
          class="field-input"
          placeholder="例如：纪要、函、自定义分类名"
          maxlength="80"
        />
        <p class="field-hint">填写分类名并上传范文后，将在「公文模板」列表中新增一项供选择。</p>
      </template>

      <template v-else>
        <label>选择要管理的模板分类</label>
        <select v-model="manageTemplateKey" class="field-select">
          <option value="">请选择分类</option>
          <option
            v-for="item in localTemplates"
            :key="'manage-' + item.templateKey"
            :value="item.templateKey"
          >
            {{ item.name }}（{{ item.fileCount }} 份）
          </option>
        </select>
        <div v-if="manageCategory" class="manage-panel">
          <div class="manage-head">
            <span>分类内文件（{{ manageFiles.length }}）</span>
            <el-button
              type="danger"
              size="small"
              plain
              :disabled="!manageTemplateKey || !!pendingDelete"
              :loading="deletingCategory"
              @click="requestDeleteCategory"
            >
              删除整个分类
            </el-button>
          </div>
          <ul v-if="manageFiles.length" class="manage-file-list">
            <li v-for="name in manageFiles" :key="name">
              <span class="file-name" :title="name">{{ name }}</span>
              <button
                type="button"
                class="delete-file-btn"
                :disabled="!!pendingDelete || deletingFile === name"
                @click="requestDeleteFile(name)"
              >
                {{ deletingFile === name ? '删除中' : '删除' }}
              </button>
            </li>
          </ul>
          <p v-else class="field-hint">该分类下暂无模板文件</p>

          <div v-if="pendingDelete" class="delete-confirm-layer">
            <div class="delete-confirm-card" role="alertdialog" aria-modal="true">
              <p class="delete-confirm-text">{{ pendingDelete.message }}</p>
              <div class="delete-confirm-actions">
                <el-button @click="pendingDelete = null">取消</el-button>
                <el-button
                  type="danger"
                  :loading="deletingCategory || !!deletingFile"
                  @click="executePendingDelete"
                >
                  删除
                </el-button>
              </div>
            </div>
          </div>
        </div>
      </template>

      <template v-if="categoryMode !== 'manage'">
        <label>模板文件</label>
        <div
          class="file-drop"
          :class="{ dragover: isDragover }"
          @click="openFilePicker"
          @dragover.prevent="isDragover = true"
          @dragleave.prevent="isDragover = false"
          @drop.prevent="onDrop"
        >
          <p v-if="!pendingFiles.length">点击或拖拽上传 DOCX / PDF / TXT 等模板文件</p>
          <ul v-else class="file-list">
            <li v-for="(file, index) in pendingFiles" :key="file.name + index">
              {{ file.name }}
              <button type="button" @click.stop="removePendingFile(index)">移除</button>
            </li>
          </ul>
        </div>
        <input
          ref="fileInput"
          type="file"
          class="hidden-file-input"
          multiple
          accept=".docx,.txt,.pdf,.png,.jpg,.jpeg,.bmp,.tiff,.wps,.ofd"
          @change="onSelect"
        />
        <p class="hint">上传后将自动解析范文结构并纳入模板库，下次生成公文时可在「公文模板」中选择使用。</p>
      </template>
    </div>

    <template v-if="categoryMode !== 'manage'" #footer>
      <el-button @click="visibleProxy = false">取消</el-button>
      <el-button type="primary" :loading="uploading" @click="submitUpload">上传并解析</el-button>
    </template>
    <template v-else #footer>
      <el-button @click="visibleProxy = false">关闭</el-button>
    </template>
  </el-dialog>
</template>

<script setup>
import { computed, ref } from 'vue'
import { ElMessage } from 'element-plus'
import {
  deleteAiTemplateCategory,
  deleteAiTemplateFile,
  getAiTemplates,
  uploadAiTemplates,
} from '../api/index.js'

const props = defineProps({
  visible: { type: Boolean, default: false },
})
const emit = defineEmits(['update:visible', 'updated'])

const visibleProxy = computed({
  get: () => props.visible,
  set: (value) => emit('update:visible', value),
})

const categoryMode = ref('new')
const selectedTemplateKey = ref('')
const manageTemplateKey = ref('')
const categoryName = ref('')
const localTemplates = ref([])
const pendingFiles = ref([])
const fileInput = ref(null)
const isDragover = ref(false)
const uploading = ref(false)
const loadingTemplates = ref(false)
const deletingFile = ref('')
const deletingCategory = ref(false)
const pendingDelete = ref(null)

function normalizeTemplateList(list) {
  return (list || []).map((item) => ({
    ...item,
    templateKey: item.templateKey || `${item.id}::${item.sourceDir || item.name}`,
  }))
}

const selectedTemplate = computed(() =>
  localTemplates.value.find(item => item.templateKey === selectedTemplateKey.value) || null,
)
const manageCategory = computed(() =>
  localTemplates.value.find(item => item.templateKey === manageTemplateKey.value) || null,
)
const manageFiles = computed(() => manageCategory.value?.files || [])

async function loadTemplates() {
  loadingTemplates.value = true
  try {
    const res = await getAiTemplates()
    localTemplates.value = normalizeTemplateList(res.data.templates)
  } catch (err) {
    ElMessage.warning('模板列表加载失败：' + (err.response?.data?.detail || err.message))
  } finally {
    loadingTemplates.value = false
  }
}

async function handleOpen() {
  categoryMode.value = 'new'
  categoryName.value = ''
  pendingFiles.value = []
  pendingDelete.value = null
  await loadTemplates()
  selectedTemplateKey.value = localTemplates.value[0]?.templateKey || ''
  manageTemplateKey.value = localTemplates.value[0]?.templateKey || ''
}

function openFilePicker() {
  fileInput.value?.click()
}

function addFiles(fileList) {
  const next = [...pendingFiles.value]
  for (const file of Array.from(fileList || [])) {
    if (!next.some(item => item.name === file.name && item.size === file.size)) {
      next.push(file)
    }
  }
  pendingFiles.value = next
}

function onSelect(event) {
  addFiles(event.target.files)
  event.target.value = ''
}

function onDrop(event) {
  isDragover.value = false
  addFiles(event.dataTransfer?.files)
}

function removePendingFile(index) {
  pendingFiles.value = pendingFiles.value.filter((_, i) => i !== index)
}

function emitTemplates(payload) {
  if (payload?.templates) {
    localTemplates.value = normalizeTemplateList(payload.templates)
  }
  emit('updated', payload)
}

async function submitUpload() {
  if (!pendingFiles.value.length) {
    ElMessage.warning('请先选择模板文件')
    return
  }
  if (categoryMode.value === 'existing' && !selectedTemplate.value) {
    ElMessage.warning('请选择已有模板分类')
    return
  }
  if (categoryMode.value === 'new' && !categoryName.value.trim()) {
    ElMessage.warning('请填写模板分类名称')
    return
  }

  uploading.value = true
  try {
    const res = await uploadAiTemplates({
      files: pendingFiles.value,
      templateId: categoryMode.value === 'existing' ? (selectedTemplate.value?.id || '') : '',
      sourceDir: categoryMode.value === 'existing' ? (selectedTemplate.value?.sourceDir || '') : '',
      categoryName: categoryMode.value === 'new' ? categoryName.value.trim() : '',
      categoryMode: categoryMode.value,
    })
    const saved = res.data.count || 0
    const skipped = res.data.skipped || []
    if (!saved) {
      const detail = skipped.map(item => `${item.filename}：${item.reason}`).join('；')
      ElMessage.warning(detail || '没有文件被收录')
      return
    }
    if (skipped.length) {
      ElMessage.warning(`已收录 ${saved} 个，跳过 ${skipped.length} 个：${skipped.map(i => i.reason).join('；')}`)
    } else {
      ElMessage.success(`模板已入库并完成预解析，共收录 ${saved} 个文件`)
    }
    emitTemplates(res.data)
    visibleProxy.value = false
  } catch (err) {
    ElMessage.error('模板入库失败：' + (err.response?.data?.detail || err.message))
  } finally {
    uploading.value = false
  }
}

function requestDeleteFile(filename) {
  pendingDelete.value = {
    kind: 'file',
    filename,
    message: `确定删除文件「${filename}」？`,
  }
}

function requestDeleteCategory() {
  const category = manageCategory.value
  if (!category) return
  pendingDelete.value = {
    kind: 'category',
    message: `确定删除整个分类「${category.name}」及其全部 ${manageFiles.value.length} 个文件？此操作不可恢复。`,
  }
}

async function executePendingDelete() {
  const pending = pendingDelete.value
  const category = manageCategory.value
  if (!pending || !category?.sourceDir) {
    ElMessage.warning('无法删除：分类信息不完整，请关闭弹窗后重新打开')
    return
  }
  if (pending.kind === 'file') {
    deletingFile.value = pending.filename
    try {
      const res = await deleteAiTemplateFile(category.sourceDir, pending.filename)
      ElMessage.success('模板文件已删除')
      emitTemplates(res.data)
      pendingDelete.value = null
    } catch (err) {
      ElMessage.error('删除失败：' + (err.response?.data?.detail || err.message))
    } finally {
      deletingFile.value = ''
    }
    return
  }
  deletingCategory.value = true
  try {
    const res = await deleteAiTemplateCategory(category.sourceDir)
    ElMessage.success('模板分类已删除')
    manageTemplateKey.value = ''
    pendingDelete.value = null
    emitTemplates(res.data)
  } catch (err) {
    ElMessage.error('删除失败：' + (err.response?.data?.detail || err.message))
  } finally {
    deletingCategory.value = false
  }
}
</script>

<style scoped>
.upload-dialog-body {
  display: flex;
  flex-direction: column;
  gap: 14px;
  min-height: 280px;
}
.upload-dialog-body label {
  font-size: 15px;
  font-weight: 600;
  color: #334155;
}
.mode-row {
  display: flex;
  gap: 12px;
  margin-bottom: 4px;
}
.mode-option {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  padding: 14px 10px;
  border: 1px solid #d8e5f2;
  border-radius: 10px;
  cursor: pointer;
  font-size: 15px;
  color: #475569;
}
.mode-option.active {
  border-color: #2d6aa0;
  background: #f0f7ff;
  color: #2d6aa0;
  font-weight: 600;
}
.mode-option input {
  accent-color: #2d6aa0;
}
.field-input,
.field-select {
  width: 100%;
  box-sizing: border-box;
  border: 1px solid #d8e5f2;
  border-radius: 10px;
  padding: 12px 14px;
  font-size: 15px;
}
.hidden-file-input {
  display: none;
}
.file-drop {
  min-height: 140px;
  border: 2px dashed #c8d9ef;
  border-radius: 12px;
  padding: 24px;
  text-align: center;
  cursor: pointer;
  color: #2d6aa0;
  background: #fbfdff;
}
.file-drop.dragover {
  border-color: #2d6aa0;
  background: #f0f7ff;
}
.file-drop p {
  margin: 0;
  font-size: 15px;
}
.file-list,
.manage-file-list {
  margin: 0;
  padding: 0;
  list-style: none;
  text-align: left;
}
.file-list li,
.manage-file-list li {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 12px;
  padding: 10px 0;
  font-size: 15px;
  color: #334155;
  border-bottom: 1px solid #eef2f7;
}
.file-list button,
.delete-file-btn {
  border: none;
  background: none;
  color: #ef4444;
  cursor: pointer;
  font-size: 14px;
  flex-shrink: 0;
}
.manage-panel {
  position: relative;
  border: 1px solid #e2e8f0;
  border-radius: 12px;
  padding: 14px 16px;
  background: #fafcfe;
}
.delete-confirm-layer {
  position: absolute;
  inset: 0;
  z-index: 10;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 16px;
  border-radius: 12px;
  background: rgba(15, 36, 56, 0.48);
}
.delete-confirm-card {
  width: min(100%, 420px);
  padding: 20px 22px;
  border-radius: 12px;
  background: #fff;
  box-shadow: 0 12px 32px rgba(15, 36, 56, 0.18);
}
.delete-confirm-text {
  margin: 0 0 18px;
  color: #334155;
  font-size: 15px;
  line-height: 1.7;
}
.delete-confirm-actions {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
}
.manage-head {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 10px;
  font-size: 14px;
  font-weight: 600;
  color: #475569;
}
.file-name {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  min-width: 0;
}
.hint,
.field-hint {
  margin: 0;
  color: #64748b;
  font-size: 14px;
  line-height: 1.6;
}
</style>

<style>
.template-manager-dialog .el-dialog__body {
  padding-top: 12px;
  padding-bottom: 16px;
}
</style>
