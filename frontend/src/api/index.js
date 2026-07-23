import axios from 'axios'

const api = axios.create({ baseURL: '/api', timeout: 120000 })

/** 批量上传文件 */
export function uploadFiles(files) {
  const formData = new FormData()
  files.forEach(f => formData.append('files', f))
  return api.post('/upload', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
  })
}

/** 获取格式配置与角色标签 */
export function getFormatConfig() {
  return api.get('/config')
}

/** 识别结构清单（带置信度），供人工校正 */
export function getStructure(fileId, templateId) {
  return api.get(`/structure/${fileId}`, {
    params: { template_id: templateId || undefined },
  })
}

/** 按文种模板与人工校正后的类型套格式。roles: { 逻辑段索引: 角色 } */
export function convertFile(fileId, roles, templateId) {
  return api.post(`/convert/${fileId}`, {
    roles: roles || null,
    template_id: templateId || null,
  })
}

/** 预览处理结果 */
export function previewResult(fileId) {
  return api.get(`/preview-result/${fileId}`)
}

/** 下载处理后的文件 */
export function downloadFile(fileId) {
  return api.get(`/download/${fileId}`, { responseType: 'blob' })
}

/** 对处理后文档应用局部修改 */
export function editFile(fileId, payload) {
  return api.post(`/edit/${fileId}`, payload)
}

/** 撤销当前文档最近一次局部修改 */
export function undoFile(fileId) {
  return api.post(`/undo/${fileId}`)
}

/** 清空所有文件 */
export function clearAll() {
  return api.delete('/files')
}

/** AI 公文写作：读取已收录模板索引 */
export function getAiTemplates() {
  return api.get('/ai/templates')
}

/** AI 公文写作：重新收录模板目录到本地数据库 */
export function syncAiTemplates() {
  return api.post('/ai/templates/sync')
}

/** AI 公文写作：上传模板文件并接入模板库 */
export function uploadAiTemplates(payload) {
  const formData = new FormData()
  ;(payload.files || []).forEach(f => formData.append('files', f))
  formData.append('template_id', payload.templateId || '')
  return api.post('/ai/templates/upload', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
    timeout: 180000,
  })
}

/** AI 公文写作：读取模型服务配置和可选模型 */
export function getAiModelConfig() {
  return api.get('/ai/model-config')
}

/** AI 公文写作：保存模型服务配置和可选模型 */
export function saveAiModelConfig(payload) {
  return api.post('/ai/model-config', {
    request_url: payload.requestUrl || '',
    api_key: payload.apiKey || '',
    model_name: payload.modelName || '',
    models: payload.models || [],
  })
}

/** AI 公文写作：提交材料、提示词和内网模型参数 */
export function generateAiDocument(payload) {
  const formData = new FormData()
  ;(payload.files || []).forEach(f => formData.append('files', f))
  formData.append('prompt', payload.prompt || '')
  formData.append('request_url', payload.requestUrl || '')
  formData.append('api_key', payload.apiKey || '')
  formData.append('model_name', payload.modelName || '')
  formData.append('template_id', payload.templateId || '')
  formData.append('temperature', String(payload.temperature ?? 0.2))
  return api.post('/ai/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
    timeout: 600000,
  })
}

/** AI 公文写作：下载生成后的正式 docx */
export function downloadAiDocument(documentId) {
  return api.get(`/ai/download/${documentId}`, { responseType: 'blob' })
}
