<template>
  <el-dialog
    v-model="visibleProxy"
    title="模型配置"
    width="1170px"
    top="6vh"
    append-to-body
    destroy-on-close
    class="model-config-dialog"
    @open="loadModelSettings"
  >
    <div class="model-config-body">
      <label>模型服务地址</label>
      <input
        v-model="requestUrl"
        class="setting-input"
        placeholder="例如：https://api.deepseek.com 或内网 /v1 地址"
      />
      <label>API Key</label>
      <input
        v-model="apiKey"
        class="setting-input"
        type="password"
        placeholder="已配置则显示掩码；留空或保持掩码则不改动"
      />
      <label>模型选择</label>
      <div class="model-row">
        <select v-model="modelName" class="setting-select">
          <option v-for="model in modelOptions" :key="model" :value="model">{{ model }}</option>
        </select>
        <el-button type="primary" @click="saveModelSettings">保存</el-button>
      </div>
      <input
        v-model="customModelName"
        class="setting-input"
        placeholder="新增模型名称，回车加入列表"
        @keydown.enter.prevent="addCustomModel"
      />
      <p class="hint">配置保存后，公文写作与公文纠错将共用此模型服务。</p>
    </div>
  </el-dialog>
</template>

<script setup>
import { computed, ref } from 'vue'
import { ElMessage } from 'element-plus'
import { getAiModelConfig, saveAiModelConfig } from '../api/index.js'

const props = defineProps({
  visible: { type: Boolean, default: false },
})
const emit = defineEmits(['update:visible'])

const visibleProxy = computed({
  get: () => props.visible,
  set: (value) => emit('update:visible', value),
})

const requestUrl = ref('')
const apiKey = ref('')
const modelName = ref('')
const customModelName = ref('')
const modelOptions = ref([])

async function loadModelSettings() {
  try {
    const res = await getAiModelConfig()
    requestUrl.value = res.data.requestUrl || ''
    apiKey.value = res.data.apiKeyConfigured ? (res.data.apiKey || '********') : ''
    modelName.value = res.data.modelName || ''
    modelOptions.value = res.data.models || []
  } catch (err) {
    ElMessage.warning('模型配置读取失败：' + (err.response?.data?.detail || err.message))
  }
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
</script>

<style scoped>
.model-config-body {
  display: flex;
  flex-direction: column;
  gap: 14px;
}
.model-config-body label {
  font-size: 15px;
  font-weight: 600;
  color: #334155;
}
.setting-input,
.setting-select {
  width: 100%;
  box-sizing: border-box;
  border: 1px solid #d8e5f2;
  border-radius: 10px;
  padding: 12px 14px;
  font-size: 15px;
}
.model-row {
  display: grid;
  grid-template-columns: 1fr auto;
  gap: 10px;
  align-items: center;
}
.hint {
  margin: 4px 0 0;
  color: #64748b;
  font-size: 14px;
  line-height: 1.6;
}
</style>

<style>
.model-config-dialog .el-dialog__body {
  padding-top: 12px;
  padding-bottom: 16px;
}
</style>
