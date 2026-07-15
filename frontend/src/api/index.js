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
export function getStructure(fileId) {
  return api.get(`/structure/${fileId}`)
}

/** 按人工校正后的类型套格式。roles: { 逻辑段索引: 角色 } */
export function convertFile(fileId, roles) {
  return api.post(`/convert/${fileId}`, { roles: roles || null })
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
