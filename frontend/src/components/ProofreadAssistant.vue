<template>
  <div v-if="visible" class="proofread-overlay">
    <div class="proofread-shell">
      <header class="proofread-header">
        <div class="title-wrap">
          <h2>公文纠错</h2>
          <p v-if="step === 'pick'">选择本地文件或平台列表中的文件</p>
          <p v-else-if="step === 'preview'">预览原文后开始纠错</p>
          <p v-else>核对批注并修改，完成后导出</p>
        </div>
        <button class="close-btn" type="button" @click="close">关闭</button>
      </header>

      <!-- 选文件 -->
      <div v-if="step === 'pick'" class="pick-layout">
        <div v-if="loading" class="running-mask pick-mask">
          <div class="running-card">正在打开文件…</div>
        </div>
        <section class="pick-pane">
          <h3>本地上传</h3>
          <p class="hint">支持 docx / txt / pdf / 图片等（优先 docx）</p>
          <div
            class="upload-box"
            :class="{ dragover: localDrag }"
            @click="localInput?.click()"
            @dragover.prevent="localDrag = true"
            @dragleave.prevent="localDrag = false"
            @drop.prevent="onLocalDrop"
          >
            <el-icon :size="28"><UploadFilled /></el-icon>
            <span>点击或拖拽上传单个文件</span>
            <input
              ref="localInput"
              type="file"
              hidden
              accept=".docx,.doc,.txt,.pdf,.png,.jpg,.jpeg,.bmp,.tiff,.wps,.ofd"
              @change="onLocalSelect"
            />
          </div>
        </section>
        <section class="pick-pane">
          <h3>平台列表</h3>
          <p class="hint">来自左侧已上传文件（单选）</p>
          <div v-if="!platformFiles.length" class="empty-list">暂无平台文件</div>
          <div v-else class="platform-list">
            <button
              v-for="f in platformFiles"
              :key="f.id"
              type="button"
              class="platform-item"
              :class="{ active: pickedPlatformId === f.id }"
              @click="pickedPlatformId = f.id"
            >
              <el-icon><Document /></el-icon>
              <span class="name" :title="f.name">{{ f.name }}</span>
            </button>
          </div>
          <el-button
            type="primary"
            class="confirm-platform"
            :disabled="!pickedPlatformId"
            :loading="loading"
            @click="openFromPlatform"
          >
            使用所选文件
          </el-button>
        </section>
      </div>

      <!-- 预览 / 纠错工作区 -->
      <div v-else class="work-layout">
        <div class="toolbar">
          <div class="file-meta">
            <strong>{{ session?.sourceName || '未命名' }}</strong>
            <span v-if="running" class="status running">纠错中…</span>
            <span v-else-if="step === 'work' && !running" class="status done">纠错完成</span>
          </div>
          <div class="toolbar-actions">
            <el-button v-if="step === 'preview'" @click="backToPick">重选文件</el-button>
            <el-button
              v-if="step === 'preview'"
              type="primary"
              :loading="running"
              :disabled="session && session.hasContent === false"
              @click="startProofread"
            >
              开始纠错
            </el-button>
            <el-button v-if="step === 'work'" @click="backToPreview">返回预览</el-button>
            <el-button
              v-if="step === 'work'"
              :loading="running"
              :disabled="session && session.hasContent === false"
              @click="startProofread"
            >
              再次纠错
            </el-button>
            <el-button
              v-if="step === 'work'"
              type="primary"
              :loading="exporting"
              @click="doExport"
            >
              导出文件
            </el-button>
          </div>
        </div>

        <div v-if="(session?.warnings || []).length" class="warn-bar">
          {{ session.warnings.join('；') }}
        </div>
        <div v-if="step === 'preview' && session && !session.hasContent" class="warn-bar danger">
          未能抽取到可读正文，无法预览纠错。请确认是未加密的 .docx（不要用旧版 .doc），或先将扫描件转为可编辑文字。
        </div>

        <div class="work-body" :class="{ 'with-notes': step === 'work' }">
          <div v-if="running" class="running-mask">
            <div class="running-card">
              <div>纠错中，已等待 {{ runElapsed }} 秒…</div>
              <p class="running-hint">小文档通常约 30 秒～2 分钟</p>
            </div>
          </div>

          <!-- 预览：单独滚动 -->
          <div v-if="step === 'preview'" ref="docPane" class="doc-pane preview-only">
            <article class="doc-paper">
              <div
                v-for="para in paragraphs"
                :key="para.index"
                class="doc-para"
                :style="{ textAlign: para.align || 'left' }"
              >
                <span class="para-text">{{ para.text || '\u00A0' }}</span>
              </div>
            </article>
          </div>

          <!-- 纠错：正文+批注同一滚动容器，水平引线随内容滚动 -->
          <div v-else ref="scrollPane" class="annotate-scroll">
            <div ref="annotateInner" class="annotate-inner">
              <div ref="docPane" class="doc-pane">
                <article class="doc-paper">
                  <div
                    v-for="para in paragraphs"
                    :key="para.index"
                    class="doc-para"
                    :data-para-index="para.index"
                    :style="{ textAlign: para.align || 'left' }"
                  >
                    <span
                      class="para-editable"
                      contenteditable="true"
                      :data-para-index="para.index"
                      @input="onParaInput(para.index, $event)"
                      @blur="onParaBlur(para.index, $event)"
                    />
                  </div>
                </article>
              </div>

              <div class="note-track">
                <div class="note-title sticky-title">批注</div>
                <div class="note-cards-area" :style="cardsAreaStyle">
                  <div v-if="!activeIssues.length" class="empty-notes">未发现需修改的问题</div>
                  <div
                    v-for="issue in activeIssues"
                    :id="`note-${issue.id}`"
                    :key="issue.id"
                    class="note-card"
                    :class="issue.level"
                    :data-issue-id="issue.id"
                    :style="noteStyle(issue.id)"
                  >
                    <div class="note-level">{{ issue.level === 'error' ? '明显错误' : '建议修改' }}</div>
                    <div class="note-body" :class="issue.level">
                      <div class="row"><span class="label">原文</span>{{ issue.original }}</div>
                      <div class="row"><span class="label">建议</span>{{ issue.suggestion }}</div>
                      <div v-if="issue.comment" class="row comment">{{ issue.comment }}</div>
                    </div>
                    <el-button size="small" type="primary" plain @click="applyIssue(issue)">一键替换</el-button>
                  </div>
                </div>

                <div class="note-summary">
                  <div class="sum-title">汇总</div>
                  <div v-if="elapsedText" class="sum-line elapsed">{{ elapsedText }}</div>
                  <div class="sum-line suggest">建议修改（黄）：{{ pendingSuggestCount }}</div>
                  <div class="sum-line error">明显错误（红）：{{ pendingErrorCount }}</div>
                  <div class="sum-line">已替换：{{ appliedCount }}</div>
                  <el-button
                    class="apply-all-btn"
                    type="primary"
                    size="small"
                    :disabled="!activeIssues.length"
                    @click="applyAllIssues"
                  >
                    全部替换
                  </el-button>
                </div>
              </div>

              <!-- 水平引线：锚定在内容坐标系，滚动时不重算斜线 -->
              <div
                v-for="line in leaderLines"
                :key="line.id"
                class="h-leader"
                :class="line.level"
                :style="{
                  top: line.top + 'px',
                  left: line.left + 'px',
                  width: line.width + 'px',
                }"
              />
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { ElMessage } from 'element-plus'
import { Document, UploadFilled } from '@element-plus/icons-vue'
import {
  createProofreadFromPlatform,
  createProofreadFromUpload,
  exportProofread,
  readBlobError,
  runProofread,
} from '../api/index.js'

const props = defineProps({
  visible: { type: Boolean, default: false },
  platformFiles: { type: Array, default: () => [] },
})
const emit = defineEmits(['close'])

const step = ref('pick') // pick | preview | work
const loading = ref(false)
const running = ref(false)
const exporting = ref(false)
const localDrag = ref(false)
const localInput = ref(null)
const pickedPlatformId = ref(null)
const session = ref(null)
const paragraphs = ref([])
const issues = ref([])
const docPane = ref(null)
const scrollPane = ref(null)
const annotateInner = ref(null)
const leaderLines = ref([])
const noteTops = ref({})
const cardsAreaHeight = ref(120)
const runElapsed = ref(0)
let lineTimer = null
let runTimer = null

const activeIssues = computed(() => issues.value.filter(i => !i.applied))
const pendingSuggestCount = computed(() => activeIssues.value.filter(i => i.level === 'suggest').length)
const pendingErrorCount = computed(() => activeIssues.value.filter(i => i.level === 'error').length)
const appliedCount = computed(() => issues.value.filter(i => i.applied).length)

function formatElapsed(sec) {
  if (sec == null || Number.isNaN(sec)) return ''
  const total = Math.max(0, Math.round(Number(sec)))
  if (total < 60) return `${total} 秒`
  const minutes = Math.floor(total / 60)
  const seconds = total % 60
  return seconds ? `${minutes} 分 ${seconds} 秒` : `${minutes} 分`
}

const elapsedText = computed(() => {
  const sec = session.value?.meta?.elapsedSeconds
  if (sec == null) return ''
  return `总耗时：${formatElapsed(sec)}`
})

const issueMap = computed(() => {
  const map = {}
  for (const issue of activeIssues.value) {
    const idx = issue.paragraphIndex
    if (!map[idx]) map[idx] = []
    map[idx].push(issue)
  }
  return map
})

const cardsAreaStyle = computed(() => ({
  minHeight: `${Math.max(120, cardsAreaHeight.value)}px`,
}))

function noteStyle(id) {
  const top = noteTops.value[id]
  if (top == null) return { visibility: 'hidden' }
  return { top: `${top}px` }
}

watch(() => props.visible, (open) => {
  if (open) resetToPick()
})

onMounted(() => {
  window.addEventListener('resize', scheduleLayout)
})
onBeforeUnmount(() => {
  window.removeEventListener('resize', scheduleLayout)
  if (lineTimer) clearTimeout(lineTimer)
  stopRunTimer()
})

function startRunTimer() {
  stopRunTimer()
  runElapsed.value = 0
  runTimer = setInterval(() => {
    runElapsed.value += 1
  }, 1000)
}

function stopRunTimer() {
  if (runTimer) {
    clearInterval(runTimer)
    runTimer = null
  }
}

function resetToPick() {
  step.value = 'pick'
  session.value = null
  paragraphs.value = []
  issues.value = []
  pickedPlatformId.value = null
  leaderLines.value = []
  noteTops.value = {}
  cardsAreaHeight.value = 120
  running.value = false
  stopRunTimer()
  runElapsed.value = 0
}

function close() {
  emit('close')
}

function backToPick() {
  resetToPick()
}

function backToPreview() {
  step.value = 'preview'
  leaderLines.value = []
  noteTops.value = {}
}

function applySession(data) {
  session.value = data
  paragraphs.value = (data.paragraphs || []).map(p => ({ ...p }))
  issues.value = (data.issues || []).map(i => ({ ...i, applied: !!i.applied }))
}

async function openFromPlatform() {
  if (!pickedPlatformId.value) return
  loading.value = true
  try {
    const res = await createProofreadFromPlatform(pickedPlatformId.value)
    if (!res.data?.success && !res.data?.sessionId) throw new Error(res.data?.detail || '创建失败')
    applySession(res.data)
    step.value = 'preview'
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || err.message || '打开失败')
  } finally {
    loading.value = false
  }
}

async function openFromFile(file) {
  if (!file) return
  loading.value = true
  try {
    const res = await createProofreadFromUpload(file)
    applySession(res.data)
    step.value = 'preview'
  } catch (err) {
    ElMessage.error(err.response?.data?.detail || err.message || '上传失败')
  } finally {
    loading.value = false
  }
}

function onLocalSelect(e) {
  const file = e.target.files?.[0]
  e.target.value = ''
  if (file) openFromFile(file)
}

function onLocalDrop(e) {
  localDrag.value = false
  const file = e.dataTransfer?.files?.[0]
  if (file) openFromFile(file)
}

async function startProofread() {
  if (!session.value?.sessionId) return
  if (session.value.hasContent === false) {
    ElMessage.warning('文档无正文，无法纠错')
    return
  }
  running.value = true
  startRunTimer()
  try {
    const res = await runProofread(session.value.sessionId)
    applySession(res.data)
    step.value = 'work'
    await nextTick()
    refreshAllParaHtml()
    scheduleLayout()
    const warn = res.data?.meta?.warning || (res.data?.warnings || []).slice(-1)[0]
    if (warn) ElMessage.warning(String(warn))
    const elapsed = res.data?.meta?.elapsedSeconds
    const elapsedHint = elapsed != null ? `，耗时 ${formatElapsed(elapsed)}` : ''
    const total = pendingSuggestCount.value + pendingErrorCount.value
    const summary = total
      ? `建议 ${pendingSuggestCount.value}，明显错误 ${pendingErrorCount.value}`
      : '未发现需修改的问题'
    ElMessage.success(`纠错完成：${summary}${elapsedHint}`)
  } catch (err) {
    const isTimeout = err.code === 'ECONNABORTED' || /timeout/i.test(String(err.message || ''))
    ElMessage.error(
      isTimeout
        ? '纠错超时（前端等待上限 15 分钟），可稍后重试或缩短文档'
        : (err.response?.data?.detail || err.message || '纠错失败'),
    )
  } finally {
    running.value = false
    stopRunTimer()
  }
}

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function renderParaHtml(para) {
  const text = String(para.text || '')
  const list = [...(issueMap.value[para.index] || [])].sort(
    (a, b) => text.indexOf(a.original) - text.indexOf(b.original),
  )
  if (!list.length) return escapeHtml(text) || '&nbsp;'

  const marks = []
  const used = []
  for (const issue of list) {
    const start = text.indexOf(issue.original)
    if (start < 0) continue
    const end = start + issue.original.length
    if (used.some(([s, e]) => !(end <= s || start >= e))) continue
    used.push([start, end])
    marks.push({ start, end, issue })
  }
  marks.sort((a, b) => a.start - b.start)
  let html = ''
  let cursor = 0
  for (const mark of marks) {
    html += escapeHtml(text.slice(cursor, mark.start))
    html += `<mark class="issue-mark ${mark.issue.level}" data-issue-id="${mark.issue.id}">${escapeHtml(text.slice(mark.start, mark.end))}</mark>`
    cursor = mark.end
  }
  html += escapeHtml(text.slice(cursor))
  return html || '&nbsp;'
}

function refreshParaHtml(index) {
  const para = paragraphs.value.find(p => p.index === index)
  const el = docPane.value?.querySelector(`.para-editable[data-para-index="${index}"]`)
  if (!para || !el) return
  el.innerHTML = renderParaHtml(para)
}

function refreshAllParaHtml() {
  for (const para of paragraphs.value) refreshParaHtml(para.index)
}

function onParaInput(index, event) {
  const el = event.target
  const text = el.innerText.replace(/\u00A0/g, ' ')
  const para = paragraphs.value.find(p => p.index === index)
  if (para) para.text = text
}

function onParaBlur(index, event) {
  onParaInput(index, event)
  const para = paragraphs.value.find(p => p.index === index)
  if (!para) return
  for (const issue of issues.value) {
    if (issue.paragraphIndex !== index || issue.applied) continue
    if (!para.text.includes(issue.original)) issue.applied = true
  }
  refreshParaHtml(index)
  nextTick(() => scheduleLayout())
}

function applyIssue(issue) {
  const para = paragraphs.value.find(p => p.index === issue.paragraphIndex)
  if (!para) return
  if (!para.text.includes(issue.original)) {
    ElMessage.warning('原文中已找不到该片段，可能已被修改')
    issue.applied = true
    refreshParaHtml(para.index)
    scheduleLayout()
    return
  }
  para.text = para.text.replace(issue.original, issue.suggestion)
  issue.applied = true
  nextTick(() => {
    refreshParaHtml(para.index)
    scheduleLayout()
  })
}

function applyAllIssues() {
  const pending = [...activeIssues.value]
  if (!pending.length) {
    ElMessage.info('没有可替换的批注')
    return
  }
  const byPara = new Map()
  for (const issue of pending) {
    const idx = issue.paragraphIndex
    if (!byPara.has(idx)) byPara.set(idx, [])
    byPara.get(idx).push(issue)
  }
  let applied = 0
  let missing = 0
  for (const [paraIndex, paraIssues] of byPara) {
    const para = paragraphs.value.find(p => p.index === paraIndex)
    if (!para) continue
    const ordered = [...paraIssues].sort((a, b) => {
      const ia = para.text.indexOf(a.original)
      const ib = para.text.indexOf(b.original)
      if (ia < 0 && ib < 0) return 0
      if (ia < 0) return 1
      if (ib < 0) return -1
      return ib - ia
    })
    for (const issue of ordered) {
      if (!para.text.includes(issue.original)) {
        issue.applied = true
        missing += 1
        continue
      }
      para.text = para.text.replace(issue.original, issue.suggestion)
      issue.applied = true
      applied += 1
    }
    refreshParaHtml(paraIndex)
  }
  nextTick(() => scheduleLayout())
  if (applied > 0) {
    const tail = missing ? `，${missing} 处原文已不存在` : ''
    ElMessage.success(`已全部替换 ${applied} 处${tail}`)
  } else {
    ElMessage.warning('未能替换任何批注')
  }
}

function scheduleLayout() {
  if (lineTimer) clearTimeout(lineTimer)
  lineTimer = setTimeout(layoutAnnotations, 50)
}

function layoutAnnotations() {
  if (step.value !== 'work' || !annotateInner.value || !docPane.value) {
    leaderLines.value = []
    return
  }
  const inner = annotateInner.value
  const innerBox = inner.getBoundingClientRect()
  const noteTrack = inner.querySelector('.note-track')
  const cardsArea = inner.querySelector('.note-cards-area')
  if (!noteTrack || !cardsArea) return

  const cardGap = 10
  const titleOffset = 8
  const placements = []
  for (const issue of activeIssues.value) {
    const mark = docPane.value.querySelector(`mark[data-issue-id="${issue.id}"]`)
    if (!mark) continue
    const markBox = mark.getBoundingClientRect()
    const markTop = markBox.top - innerBox.top
    const markMid = markTop + markBox.height / 2
    const markRight = markBox.right - innerBox.left
    placements.push({
      id: issue.id,
      level: issue.level,
      markTop,
      markMid,
      markRight,
    })
  }
  placements.sort((a, b) => a.markTop - b.markTop)

  // 批注 top 相对 note-cards-area（其顶部约等于标题之下）
  const cardsBox = cardsArea.getBoundingClientRect()
  const cardsOffset = cardsBox.top - innerBox.top
  const tops = {}
  let cursor = titleOffset
  for (const item of placements) {
    const localTop = Math.max(item.markTop - cardsOffset, titleOffset)
    const top = Math.max(localTop, cursor)
    tops[item.id] = top
    cursor = top + 118 + cardGap
  }
  noteTops.value = tops

  nextTick(() => {
    const refined = { ...tops }
    let nextY = titleOffset
    const lines = []
    for (const item of placements) {
      const note = inner.querySelector(`#note-${item.id}`)
      const height = note ? note.offsetHeight : 110
      const localTop = Math.max(item.markTop - cardsOffset, titleOffset)
      const top = Math.max(localTop, nextY)
      refined[item.id] = top
      nextY = top + height + cardGap

      const noteBox = note?.getBoundingClientRect()
      const lineY = item.markMid
      const x1 = item.markRight + 2
      const x2 = noteBox
        ? (noteBox.left - innerBox.left - 2)
        : (noteTrack.getBoundingClientRect().left - innerBox.left + 8)
      lines.push({
        id: item.id,
        level: item.level,
        top: lineY,
        left: x1,
        width: Math.max(12, x2 - x1),
      })
    }
    noteTops.value = refined
    leaderLines.value = lines
    cardsAreaHeight.value = Math.max(120, nextY + 24)

    const paper = docPane.value.querySelector('.doc-paper')
    const paperBottom = paper
      ? paper.getBoundingClientRect().bottom - innerBox.top
      : cardsAreaHeight.value
    if (inner.offsetHeight < paperBottom + 40) {
      inner.style.minHeight = `${Math.ceil(paperBottom + 40)}px`
    }
  })
}

async function doExport() {
  if (!session.value?.sessionId) return
  exporting.value = true
  try {
    const res = await exportProofread(session.value.sessionId, paragraphs.value)
    const disposition = res.headers?.['content-disposition'] || ''
    let filename = `${(session.value.sourceName || '公文').replace(/\.[^.]+$/, '')}_纠错.docx`
    const match = /filename\*?=(?:UTF-8''|")?([^";]+)/i.exec(disposition)
    if (match) {
      try { filename = decodeURIComponent(match[1].replace(/"/g, '')) } catch { /* keep */ }
    }
    const url = window.URL.createObjectURL(new Blob([res.data]))
    const link = document.createElement('a')
    link.href = url
    link.download = filename
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
    ElMessage.success('已导出修改后的文件')
  } catch (err) {
    ElMessage.error(await readBlobError(err, '导出失败'))
  } finally {
    exporting.value = false
  }
}

watch(activeIssues, () => nextTick(() => scheduleLayout()), { deep: true })
</script>

<style scoped>
.proofread-overlay {
  position: fixed; inset: 0; z-index: 4000;
  background: #0f2438;
  display: flex; align-items: stretch; justify-content: stretch;
  padding: 0;
}
.proofread-shell {
  width: 100%;
  height: 100%;
  max-width: none;
  background: #f5f8fc;
  border-radius: 0;
  box-shadow: none;
  display: flex; flex-direction: column; overflow: hidden;
}
.proofread-header {
  display: flex; align-items: center; justify-content: space-between;
  padding: 20px 28px; background: linear-gradient(135deg, #1a3a5c, #2d6aa0); color: #fff;
  flex-shrink: 0;
}
.proofread-header h2 { font-size: 24px; font-weight: 700; }
.proofread-header p { margin-top: 6px; font-size: 15px; opacity: 0.85; }
.close-btn {
  border: 1px solid rgba(255,255,255,0.35); background: transparent; color: #fff;
  border-radius: 10px; padding: 10px 18px; cursor: pointer; font-size: 15px;
}
.close-btn:hover { background: rgba(255,255,255,0.12); }

.pick-layout {
  position: relative;
  flex: 1; display: grid; grid-template-columns: 1fr 1fr; gap: 24px; padding: 28px 32px; min-height: 0;
}
.pick-mask { inset: 28px 32px; }
.pick-pane {
  background: #fff; border-radius: 14px; padding: 24px; border: 1px solid #e3ebf5;
  display: flex; flex-direction: column; min-height: 0;
}
.pick-pane h3 { font-size: 19px; color: #1f3d5c; }
.hint { color: #7a8b9c; font-size: 14px; margin: 8px 0 16px; }
.upload-box {
  flex: 1; border: 2px dashed #c8d9ef; border-radius: 14px; min-height: 280px;
  display: flex; flex-direction: column; align-items: center; justify-content: center;
  gap: 12px; color: #2d6aa0; cursor: pointer; font-size: 15px;
}
.upload-box.dragover, .upload-box:hover { background: #f0f7ff; border-color: #2d6aa0; }
.platform-list { flex: 1; overflow: auto; display: flex; flex-direction: column; gap: 8px; }
.platform-item {
  display: flex; align-items: center; gap: 10px; text-align: left;
  border: 1px solid #e6edf5; background: #fff; border-radius: 10px; padding: 12px 14px; cursor: pointer;
}
.platform-item.active { border-color: #2d6aa0; background: #eef5ff; }
.platform-item .name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 15px; }
.empty-list { color: #aaa; text-align: center; padding: 48px 0; font-size: 15px; }
.confirm-platform { margin-top: 16px; align-self: flex-start; }

.work-layout { flex: 1; display: flex; flex-direction: column; min-height: 0; }
.toolbar {
  display: flex; justify-content: space-between; align-items: center; gap: 14px;
  padding: 14px 28px; background: #fff; border-bottom: 1px solid #e6edf5;
  flex-shrink: 0;
}
.file-meta { display: flex; align-items: center; gap: 12px; font-size: 16px; color: #24384f; }
.status { font-size: 13px; padding: 3px 10px; border-radius: 999px; }
.status.running { background: #fff4e5; color: #c27803; }
.status.done { background: #e8f8ef; color: #1f9d57; }
.toolbar-actions { display: flex; gap: 10px; }
.warn-bar {
  margin: 0 28px; margin-top: 12px; padding: 10px 14px; border-radius: 10px;
  background: #fff8e8; color: #9a6b16; font-size: 14px;
}
.warn-bar.danger {
  background: #fff1f0;
  color: #b42318;
}
.work-body {
  position: relative; flex: 1; display: flex; flex-direction: column; min-height: 0;
  padding: 20px 28px 24px;
}
.running-mask {
  position: absolute; inset: 20px 28px 24px; z-index: 5;
  background: rgba(245, 248, 252, 0.72);
  display: flex; align-items: center; justify-content: center;
  border-radius: 14px;
}
.running-card {
  background: #fff; border: 1px solid #d8e5f2; border-radius: 12px;
  padding: 22px 36px; color: #2d6aa0; font-weight: 700; font-size: 18px;
  box-shadow: 0 8px 24px rgba(30, 60, 100, 0.12);
  text-align: center; max-width: 420px;
}
.running-hint {
  margin-top: 10px; font-size: 13px; font-weight: 400; color: #6b7c8f; line-height: 1.5;
}
.doc-pane.preview-only { overflow: auto; flex: 1; min-height: 0; }
.annotate-scroll {
  flex: 1; min-height: 0; overflow: auto;
  background: #eef2f7; border: 1px solid #e1e8f0; border-radius: 14px;
}
.annotate-inner {
  position: relative;
  display: grid;
  grid-template-columns: minmax(0, 1fr) 360px;
  gap: 0;
  align-items: start;
  min-height: 100%;
  padding: 24px 20px 48px;
}
.doc-pane { min-width: 0; }
.doc-paper {
  max-width: 920px; margin: 0 auto; background: #fff; min-height: 560px;
  padding: 58px 68px; box-shadow: 0 2px 16px rgba(20, 45, 80, 0.08);
  border: 1px solid #e8eef5; font-family: "Songti SC", "SimSun", "Noto Serif SC", serif;
  font-size: 18px; line-height: 1.95; color: #222;
}
.doc-para { min-height: 1.95em; margin: 0 0 0.25em; white-space: pre-wrap; word-break: break-word; }
.para-editable { outline: none; display: inline; min-width: 1em; }
.para-editable:focus { background: rgba(45, 106, 160, 0.04); }

.note-track {
  position: relative;
  min-height: 100%;
  display: flex;
  flex-direction: column;
  padding: 6px 12px 0 22px;
  background: #f7f8fa;
  border-left: 1px solid #e5ebf2;
}
.sticky-title {
  flex-shrink: 0;
  position: sticky;
  top: 0;
  z-index: 4;
  background: #f7f8fa;
  padding: 6px 0 12px;
  font-weight: 700; color: #1f3d5c; font-size: 16px;
}
.note-cards-area {
  position: relative;
  flex: 1 0 auto;
  min-height: 120px;
  padding-bottom: 12px;
}
.empty-notes { color: #aaa; font-size: 15px; padding: 24px 0; text-align: center; }
.note-card {
  position: absolute; left: 0; right: 8px;
  border: 1px solid #edf1f6; border-radius: 10px; padding: 12px;
  background: #fff; box-shadow: 0 1px 4px rgba(20, 40, 70, 0.04);
  z-index: 1;
}
.note-card.error { border-left: 3px solid #d64545; }
.note-card.suggest { border-left: 3px solid #d4a017; }
.note-level { font-size: 13px; font-weight: 700; margin-bottom: 8px; color: #667788; }
.note-body.error { color: #c62828; }
.note-body.suggest { color: #b8860b; }
.note-body .row { font-size: 14px; margin-bottom: 5px; line-height: 1.55; }
.note-body .label {
  display: inline-block; min-width: 36px; margin-right: 6px; color: #8a97a6; font-size: 13px;
}
.note-body .comment { opacity: 0.9; }
.note-summary {
  flex-shrink: 0;
  position: sticky;
  bottom: 0;
  z-index: 5;
  margin: 0 -12px 0 -22px;
  padding: 14px 22px 16px;
  border-radius: 0;
  background: #fff;
  border: none;
  border-top: 1px dashed #dbe4ef;
  font-size: 14px;
  box-shadow: 0 -6px 12px rgba(20, 40, 70, 0.06);
}
.sum-title { font-weight: 700; margin-bottom: 8px; color: #1f3d5c; font-size: 15px; }
.sum-line { margin: 5px 0; }
.sum-line.elapsed { color: #4a6078; }
.sum-line.suggest { color: #b8860b; }
.sum-line.error { color: #c62828; }
.apply-all-btn { width: 100%; margin-top: 12px; }

/* 水平引线：锚定内容坐标，随页面一起滚动 */
.h-leader {
  position: absolute;
  height: 0;
  border-top: 1px dashed #c9ced6;
  pointer-events: none;
  z-index: 0;
  transform: translateY(-0.5px);
}
.h-leader.error { border-top-color: #e07070; }
.h-leader.suggest { border-top-color: #d4a017; }

:deep(mark.issue-mark) {
  background: rgba(230, 120, 120, 0.28);
  color: inherit;
  padding: 1px 2px;
  border-radius: 2px;
  box-decoration-break: clone;
  -webkit-box-decoration-break: clone;
  border-left: 2px solid rgba(210, 70, 70, 0.55);
  text-decoration: underline;
  text-decoration-thickness: 1px;
  text-underline-offset: 3px;
}
:deep(mark.issue-mark.error) {
  background: rgba(220, 80, 80, 0.32);
  border-left-color: #c62828;
  text-decoration-color: #c62828;
}
:deep(mark.issue-mark.suggest) {
  background: rgba(230, 120, 120, 0.22);
  border-left-color: #d4a017;
  text-decoration-color: #d4a017;
}

@media (max-width: 960px) {
  .pick-layout { grid-template-columns: 1fr; padding: 16px; gap: 16px; }
  .annotate-inner { grid-template-columns: 1fr; padding: 16px 12px 32px; }
  .doc-paper { max-width: none; padding: 36px 28px; font-size: 16px; }
  .h-leader { display: none; }
  .note-track { border-left: none; border-top: 1px solid #e5ebf2; margin-top: 12px; }
  .note-card { position: relative; left: auto; right: auto; top: auto !important; margin-bottom: 10px; }
  .note-summary { position: relative; bottom: auto; margin: 0; }
}
</style>
