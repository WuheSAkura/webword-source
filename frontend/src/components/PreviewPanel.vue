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
    </div>

    <div class="preview-body" v-loading="loading">
      <el-empty v-if="!selectedFile" description="选择左侧文件即可预览" :image-size="90" />
      <el-empty v-else-if="loading" description="加载中..." :image-size="60" />
      <el-empty v-else-if="!resultData || !resultData.paragraphs" description="尚未套用格式" :image-size="60" />
      <template v-else>
        <div class="document-canvas">
          <div class="document-page result-page" @mouseup="captureSelection" @keyup="captureSelection">
            <div
              v-for="p in resultData.paragraphs"
              :key="p.index"
              class="doc-paragraph"
              :class="'role-' + p.role"
              :data-order="p.index"
              :data-para-index="p.sourceIndex ?? p.index"
              :data-role="p.role"
              :style="paragraphStyle(p)"
              :title="`${p.roleLabel} / 源段落 ${(p.sourceIndex ?? p.index) + 1}`"
            >
              <span class="doc-text">
                <span
                  v-for="(run, runIndex) in paragraphRuns(p)"
                  :key="runIndex"
                  :style="runStyle(run)"
                >{{ run.text }}</span>
              </span>
            </div>
          </div>
        </div>
      </template>
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'

const props = defineProps({
  selectedFile: Object,
  resultData: Object,
  loading: Boolean,
})
const emit = defineEmits(['selectRange'])

const roleLabels = {
  title: '标题', subtitle: '副标题', recipient: '发文对象', body: '正文', h1: '一级标题', h2: '二级标题',
  h3: '三级标题', h4: '四级标题', attachment_head: '附件头', attachment_other: '其他附件', sign_unit: '落款·单位',
  sign_date: '落款·日期', sign_contact: '落款·联系人', security: '涉密标识', other: '其他',
}

const paragraphCount = computed(() => props.resultData?.paragraphs?.length || 0)

const roleCounts = computed(() => {
  const map = new Map()
  for (const p of props.resultData?.paragraphs || []) map.set(p.role, (map.get(p.role) || 0) + 1)
  return Array.from(map.entries()).map(([role, count]) => ({ role, count, label: roleLabels[role] || role }))
})

function paragraphRuns(paragraph) {
  return paragraph.runs?.length ? paragraph.runs : [{ text: paragraph.fullText }]
}

function runStyle(run) {
  const style = {}
  const fonts = fontFamilyList(run.fontEast, run.fontWest)
  if (fonts.length) style.fontFamily = fonts.join(', ')
  if (run.size) style.fontSize = `${run.size}pt`
  if (run.bold !== null && run.bold !== undefined) style.fontWeight = run.bold ? '700' : '400'
  return style
}

function paragraphStyle(paragraph) {
  const layout = paragraph.layout || {}
  const style = {}
  const alignMap = { left: 'left', center: 'center', right: 'right', justify: 'justify', distribute: 'justify' }
  if (layout.align) style.textAlign = alignMap[layout.align] || layout.align
  if (layout.lineRule === 'exact' && layout.linePt) {
    style.lineHeight = `${layout.linePt}pt`
  } else if (layout.lineRule === 'multiple' && layout.lineMultiple) {
    style.lineHeight = String(layout.lineMultiple)
  }
  if (layout.firstLineChars) style.textIndent = `${layout.firstLineChars}em`
  else style.textIndent = '0'
  if (layout.leftChars) style.marginLeft = `${layout.leftChars}em`
  if (layout.rightChars) style.marginRight = `${layout.rightChars}em`
  return style
}

function quoteFont(name) { return `"${String(name).replace(/"/g, '')}"` }

function fontFamilyList(east, west) {
  const eastName = east || ''
  const fallbackMap = [
    { test: '小标宋', fonts: ['方正小标宋简体', 'FZXiaoBiaoSong-B05S', 'SimSun', '宋体'] },
    { test: '黑体', fonts: ['黑体', 'SimHei', 'Microsoft YaHei'] },
    { test: '楷体', fonts: ['楷体_GB2312', '楷体', 'KaiTi_GB2312', 'KaiTi', 'STKaiti'] },
    { test: '仿宋', fonts: ['仿宋_GB2312', '仿宋', 'FangSong_GB2312', 'FangSong', 'STFangsong'] },
    { test: '宋体', fonts: ['宋体', 'SimSun'] },
  ]
  const matched = fallbackMap.find(item => eastName.includes(item.test))
  const fonts = matched ? matched.fonts : (eastName ? [eastName] : [])
  if (west) fonts.push(west)
  return Array.from(new Set(fonts)).map(quoteFont)
}

function findTextSpan(node) {
  let cur = node?.nodeType === Node.ELEMENT_NODE ? node : node?.parentElement
  while (cur && !cur.classList?.contains('doc-text')) cur = cur.parentElement
  return cur
}

function offsetInSpan(container, offset, span) {
  const range = document.createRange()
  range.selectNodeContents(span)
  try {
    range.setEnd(container, offset)
    return range.toString().length
  } catch {
    return 0
  }
}

function captureSelection() {
  const selection = window.getSelection()
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return
  const anchorSpan = findTextSpan(selection.anchorNode)
  const focusSpan = findTextSpan(selection.focusNode)
  if (!anchorSpan || !focusSpan) return
  const anchorPara = anchorSpan.closest('.doc-paragraph')
  const focusPara = focusSpan.closest('.doc-paragraph')
  if (!anchorPara || !focusPara) return

  const anchor = {
    order: Number(anchorPara.dataset.order),
    paragraph: Number(anchorPara.dataset.paraIndex),
    role: anchorPara.dataset.role,
    offset: offsetInSpan(selection.anchorNode, selection.anchorOffset, anchorSpan),
  }
  const focus = {
    order: Number(focusPara.dataset.order),
    paragraph: Number(focusPara.dataset.paraIndex),
    role: focusPara.dataset.role,
    offset: offsetInSpan(selection.focusNode, selection.focusOffset, focusSpan),
  }
  let start = anchor, end = focus
  if (anchor.order > focus.order || (anchor.order === focus.order && anchor.offset > focus.offset)) {
    start = focus; end = anchor
  }
  if (start.paragraph === end.paragraph && start.offset === end.offset) return
  emit('selectRange', {
    startParagraph: start.paragraph,
    startOffset: start.offset,
    endParagraph: end.paragraph,
    endOffset: end.offset,
    role: start.role,
    text: selection.toString(),
  })
}
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
.summary-chip { padding: 2px 8px; border-radius: 999px; background: #edf2f7; color: #475569; font-size: 12px; line-height: 20px; }
.preview-body { flex: 1; overflow-y: auto; background: #e9eef5; }
.document-canvas { min-height: 100%; padding: 28px 18px 36px; overflow-x: auto; }
.document-page {
  width: min(794px, 100%); min-height: 1123px; margin: 0 auto; padding: 86px 96px;
  background: #fff; color: #111827; border: 1px solid #d7dee8;
  box-shadow: 0 16px 38px rgba(31, 41, 55, 0.16);
  font-family: FangSong, 仿宋_GB2312, 仿宋, "Times New Roman", serif;
}
.doc-paragraph {
  position: relative; margin: 0; color: #111; font-size: 21px; line-height: 1.75;
  white-space: pre-wrap; word-break: break-word;
}
.doc-paragraph + .doc-paragraph { margin-top: 2px; }
.result-page .doc-paragraph.role-title {
  margin-bottom: 18px; text-align: center; font-family: "方正小标宋简体", SimSun, 宋体, serif;
  font-size: 29px; font-weight: 400; line-height: 1.45;
}
.result-page .doc-paragraph.role-subtitle {
  text-align: center; font-family: KaiTi, 楷体_GB2312, 楷体, serif; font-size: 21px;
}
.result-page .doc-paragraph.role-h1 {
  font-family: SimHei, 黑体, sans-serif; font-size: 21px; font-weight: 400; text-indent: 0; text-align: justify;
}
.result-page .doc-paragraph.role-security {
  font-family: SimHei, 黑体, sans-serif; font-size: 21px; text-align: left; text-indent: 0;
}
.result-page .doc-paragraph.role-h2 {
  font-family: KaiTi, 楷体_GB2312, 楷体, serif; font-size: 21px; font-weight: 400; text-indent: 0; text-align: justify;
}
.result-page .doc-paragraph.role-h3,
.result-page .doc-paragraph.role-h4 {
  font-family: FangSong, 仿宋_GB2312, 仿宋, serif; font-size: 21px; text-indent: 0; text-align: justify;
}
.result-page .doc-paragraph.role-recipient { text-indent: 0; text-align: justify; }
.result-page .doc-paragraph.role-body { text-indent: 2em; text-align: justify; }
.result-page .doc-paragraph.role-attachment_head { text-indent: 2em; text-align: justify; }
.result-page .doc-paragraph.role-attachment_other { text-indent: 5em; text-align: justify; }
.result-page .doc-paragraph.role-sign_unit { text-align: right; text-indent: 0; }
.result-page .doc-paragraph.role-sign_date { text-align: right; text-indent: 0; }
.result-page .doc-paragraph.role-sign_contact { text-align: justify; text-indent: 2em; }
.result-page .doc-paragraph[class*="role-"]::after {
  content: attr(title); position: absolute; left: calc(100% + 14px); top: 8px;
  max-width: 110px; padding: 3px 7px; border-radius: 6px; background: #f7fafc;
  color: #718096; border: 1px solid #e2e8f0; font-family: "Microsoft YaHei", sans-serif;
  font-size: 11px; line-height: 1.4; text-indent: 0; opacity: 0; pointer-events: none; transition: opacity 0.15s;
}
.result-page .doc-paragraph.role-body::after { display: none; }
.doc-paragraph:hover::after { opacity: 1; }
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
  .document-canvas { padding: 14px 10px 24px; }
  .document-page { width: 100%; min-height: 780px; padding: 42px 34px; }
  .doc-paragraph { font-size: 17px; line-height: 1.85; }
  .result-page .doc-paragraph.role-title { font-size: 23px; }
  .doc-paragraph::after { display: none; }
}
</style>
