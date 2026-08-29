import axios from 'axios'

const api = axios.create({ baseURL: '/api', timeout: 180000 })

/** 将 blob 错误响应解析为可读文案（避免把 JSON 错误体当 docx 下载） */
export async function readBlobError(error, fallback = '请求失败') {
  const data = error?.response?.data
  if (data instanceof Blob) {
    try {
      const text = await data.text()
      const parsed = JSON.parse(text)
      if (parsed?.detail) {
        return typeof parsed.detail === 'string' ? parsed.detail : JSON.stringify(parsed.detail)
      }
      if (parsed?.message) return parsed.message
      return text || fallback
    } catch {
      return fallback
    }
  }
  if (typeof data?.detail === 'string') return data.detail
  return error?.message || fallback
}

function ensureBlobOk(response, fallback) {
  const type = String(response.headers?.['content-type'] || '')
  if (type.includes('application/json')) {
    return response.data.text().then((text) => {
      let detail = fallback
      try {
        const parsed = JSON.parse(text)
        detail = parsed?.detail || parsed?.message || text || fallback
      } catch {
        detail = text || fallback
      }
      const err = new Error(detail)
      err.response = { status: response.status, data: { detail } }
      throw err
    })
  }
  return response
}

/** 批量上传文件 */
export function uploadFiles(files) {
  const formData = new FormData()
  files.forEach(f => formData.append('files', f))
  return api.post('/upload', formData)
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
export function downloadFile(fileId, options = {}) {
  return api.get(`/download/${fileId}`, {
    responseType: 'blob',
    onDownloadProgress: options.onProgress,
  }).then((res) => ensureBlobOk(res, '下载失败，文件可能已过期，请重新处理或重新生成'))
}

/** 对处理后文档应用局部修改 */
export function editFile(fileId, payload) {
  return api.post(`/edit/${fileId}`, payload)
}

/** 撤销当前文档最近一次局部修改 */
export function undoFile(fileId) {
  return api.post(`/undo/${fileId}`)
}

/** 删除单个文件 */
export function deleteFile(fileId) {
  return api.delete(`/files/${fileId}`)
}

/** 清空所有文件 */
export function clearAll() {
  return api.delete('/files')
}

/** AI 公文写作：读取已收录模板索引 */
export function getAiTemplates() {
  return api.get('/ai/templates', { timeout: 30000 })
}

/** AI 公文写作：重新收录模板目录到本地数据库 */
export function syncAiTemplates() {
  return api.post('/ai/templates/sync', null, { timeout: 300000 })
}

/** AI 公文写作：上传模板文件并接入模板库 */
export function uploadAiTemplates(payload) {
  const formData = new FormData()
  ;(payload.files || []).forEach(f => formData.append('files', f))
  formData.append('template_id', payload.templateId || '')
  formData.append('source_dir', payload.sourceDir || '')
  formData.append('category_name', payload.categoryName || '')
  formData.append('category_mode', payload.categoryMode || 'existing')
  return api.post('/ai/templates/upload', formData, {
    timeout: 300000,
  })
}

/** AI 公文写作：删除模板分类下的单个文件 */
export function deleteAiTemplateFile(sourceDir, filename) {
  return api.delete('/ai/templates/file', {
    params: { source_dir: sourceDir, filename },
    timeout: 60000,
  })
}

/** AI 公文写作：删除整个模板分类 */
export function deleteAiTemplateCategory(sourceDir) {
  return api.delete('/ai/templates/category', {
    params: { source_dir: sourceDir },
    timeout: 60000,
  })
}

/** AI 公文写作：文种目录（新建模板分类时选择） */
export function getAiTemplateCatalog() {
  return api.get('/ai/templates/catalog', { timeout: 30000 })
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
  ;(payload.materialFiles || []).forEach(f => formData.append('material_files', f))
  ;(payload.templateFiles || []).forEach(f => formData.append('template_files', f))
  formData.append('prompt', payload.prompt || '')
  // 无上传文件时，对话框文本可作为材料正文（免文件提取）
  if (payload.materialText) {
    formData.append('material_text', payload.materialText)
  }
  formData.append('request_url', payload.requestUrl || '')
  formData.append('api_key', payload.apiKey || '')
  formData.append('model_name', payload.modelName || '')
  formData.append('template_id', payload.templateId || '')
  if (payload.templateKey) {
    formData.append('template_key', payload.templateKey)
  }
  formData.append('temperature', String(payload.temperature ?? 0.2))
  formData.append('speed_mode', payload.speedMode || 'standard')
  formData.append('strict_reference_isolation', payload.strictReferenceIsolation ? 'true' : 'false')
  formData.append('allow_degradation', payload.allowDegradation ? 'true' : 'false')
  return api.post('/ai/generate', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
    // 多材料 + 多轮模型调用；高并发排队时整次生成可能超过 10 分钟
    timeout: 900000,
    signal: payload.signal,
  })
}

/** 查询公文调度任务状态与节点日志 */
export function getAiTask(taskId) {
  return api.get(`/ai/tasks/${taskId}`, { timeout: 60000 })
}

/** AI 公文写作：下载生成后的正式 docx（统一走工作台 fileId 通道时优先用 downloadFile） */
export function downloadAiDocument(documentId) {
  return api.get(`/ai/download/${documentId}`, { responseType: 'blob' }).then((res) => ensureBlobOk(res, 'AI 生成文件不存在或已过期，请重新生成'))
}

/** LibreOffice 真实渲染校验 */
export function renderCheck(fileId) {
  return api.get(`/render-check/${fileId}`, { timeout: 120000 })
}

/** 下载真实渲染 PDF */
export function downloadRenderPdf(fileId) {
  return api.get(`/render-pdf/${fileId}`, { responseType: 'blob' }).then((res) => ensureBlobOk(res, '渲染 PDF 不存在，请先执行真实渲染校验'))
}

/** 公文纠错：从平台文件创建会话 */
export function createProofreadFromPlatform(fileId) {
  return api.post('/proofread/from-platform', { file_id: fileId }, { timeout: 120000 })
}

/** 公文纠错：本地上传创建会话 */
export function createProofreadFromUpload(file) {
  const formData = new FormData()
  formData.append('file', file)
  return api.post('/proofread/upload', formData, {
    headers: { 'Content-Type': 'multipart/form-data' },
    timeout: 180000,
  })
}

/** 公文纠错：读取会话 */
export function getProofreadSession(sessionId) {
  return api.get(`/proofread/session/${sessionId}`, { timeout: 60000 })
}

/** 公文纠错：开始识别（DeepSeek） */
export function runProofread(sessionId) {
  return api.post(`/proofread/session/${sessionId}/run`, null, { timeout: 900000 })
}

/** 公文纠错：导出当前正文 */
export function exportProofread(sessionId, paragraphs) {
  return api.post(
    `/proofread/session/${sessionId}/export`,
    { paragraphs: paragraphs || null },
    { responseType: 'blob', timeout: 180000 },
  ).then((res) => ensureBlobOk(res, '导出失败'))
}
