<script setup lang="ts">
import { ref, computed } from 'vue'
import { useData, withBase } from 'vitepress'
import { copyToClipboard } from '../clipboard'

export interface FormulaRow {
  cells?: string[]
  input?: string
  formula?: string
  desc?: string
  out?: string
}

const props = defineProps<{
  title?: string
  columns?: string[]
  startRow?: number
  rows: FormulaRow[]
}>()

const { lang } = useData()
const isKo = computed(() => lang.value.startsWith('ko'))

const copiedAll = ref(false)
const copiedIndex = ref<number | null>(null)
let allTimer: ReturnType<typeof setTimeout> | undefined
let rowTimer: ReturnType<typeof setTimeout> | undefined

const baseRow = computed(() => props.startRow ?? 2)

const colCount = computed(() => {
  if (props.columns?.length) return props.columns.length
  let maxCells = 0
  let hasInput = false
  for (const r of props.rows) {
    if (r.cells?.length) {
      maxCells = Math.max(maxCells, r.cells.length)
    }
    if (r.input !== undefined) {
      hasInput = true
    }
  }
  if (maxCells > 0) return hasInput ? Math.max(2, maxCells) : maxCells
  return hasInput ? 2 : 1
})

const colLetters = computed(() =>
  ['A', 'B', 'C', 'D', 'E', 'F'].slice(0, colCount.value)
)

const excelHeaders = computed(() => {
  if (props.columns?.length === colCount.value) return props.columns
  if (colCount.value === 1) return [isKo.value ? '수식' : 'Formula']
  if (colCount.value === 2) {
    return isKo.value ? ['입력값', '수식'] : ['Input', 'Formula']
  }
  return colLetters.value.map((letter, idx) =>
    idx === colCount.value - 1
      ? (isKo.value ? '수식' : 'Formula')
      : (isKo.value ? `입력 ${letter}` : `Input ${letter}`)
  )
})

const hasDesc = computed(() => props.rows.some((r) => Boolean(r.desc)))
const hasOut = computed(() => props.rows.some((r) => Boolean(r.out)))

const normalizedRows = computed(() => {
  const count = colCount.value
  const letters = colLetters.value
  return props.rows.map((row, idx) => {
    const rowNum = baseRow.value + idx
    let excelCells: string[]
    if (row.cells) {
      if (row.cells.length === count) {
        excelCells = [...row.cells]
      } else if (row.cells.length < count) {
        excelCells = [...Array(count - row.cells.length).fill(''), ...row.cells]
      } else {
        excelCells = row.cells.slice(0, count)
      }
    } else if (count >= 2) {
      excelCells = [
        ...Array(count - 2).fill(''),
        row.input ?? '',
        row.formula ?? ''
      ]
    } else {
      excelCells = [row.formula ?? '']
    }

    const foundFirst = excelCells.findIndex((c) => c !== '')
    const firstActiveCol = foundFirst >= 0 ? foundFirst : 0
    const lastColIdx = count - 1
    const rangeLabel =
      firstActiveCol >= lastColIdx
        ? `${letters[lastColIdx]}${rowNum}`
        : `${letters[firstActiveCol]}${rowNum}:${letters[lastColIdx]}${rowNum}`
    const copyPayload = excelCells.slice(firstActiveCol).join('\t')

    return {
      rowNum,
      excelCells,
      firstActiveCol,
      rangeLabel,
      copyPayload,
      desc: row.desc ?? '',
      out: row.out ?? ''
    }
  })
})

const headerRowNum = computed(() => baseRow.value - 1)

const tableRangeLabel = computed(() => {
  if (!normalizedRows.value.length) return `A${headerRowNum.value}`
  const lastCol = colLetters.value[colCount.value - 1] ?? 'A'
  const lastRow = baseRow.value + normalizedRows.value.length - 1
  return `A${headerRowNum.value}:${lastCol}${lastRow}`
})

const headerTitle = computed(() =>
  props.title || (isKo.value ? '엑셀 수식 예제' : 'Excel Formula Examples')
)

const copyAllLabel = computed(() =>
  copiedAll.value
    ? (isKo.value ? '표 전체 복사됨' : 'Copied Table')
    : (isKo.value ? '전체 복사' : 'Copy All')
)

const copyAllTitle = computed(() =>
  isKo.value
    ? `시트 전체 (${tableRangeLabel.value})를 TSV로 복사합니다. Excel A${headerRowNum.value} 셀에 붙여넣으면 입력값과 수식이 열/행에 맞춰 바로 계산됩니다.`
    : `Copy entire grid (${tableRangeLabel.value}) as TSV. Paste into Excel at cell A${headerRowNum.value} to populate inputs and live formulas across columns and rows.`
)

async function copyAll() {
  if (!normalizedRows.value.length) return
  const payload = [
    excelHeaders.value.join('\t'),
    ...normalizedRows.value.map((r) => r.excelCells.join('\t'))
  ].join('\r\n')
  if (await copyToClipboard(payload)) {
    copiedIndex.value = null
    copiedAll.value = true
    if (allTimer) clearTimeout(allTimer)
    allTimer = setTimeout(() => {
      copiedAll.value = false
    }, 2000)
  }
}

async function copyRow(index: number) {
  const row = normalizedRows.value[index]
  if (!row || !row.copyPayload) return
  if (await copyToClipboard(row.copyPayload)) {
    copiedAll.value = false
    copiedIndex.value = index
    if (rowTimer) clearTimeout(rowTimer)
    rowTimer = setTimeout(() => {
      copiedIndex.value = null
    }, 2000)
  }
}
</script>

<template>
  <div class="s2i-ft-wrap vp-raw">
    <div class="s2i-ft-header">
      <div class="s2i-ft-title-group">
        <img
          :src="withBase('/icons/excel.svg')"
          alt=""
          width="16"
          height="16"
          class="s2i-ft-icon"
          aria-hidden="true"
        />
        <span class="s2i-ft-title">{{ headerTitle }}</span>
        <span class="s2i-ft-range">{{ tableRangeLabel }}</span>
      </div>
      <div class="s2i-ft-header-actions">
        <span class="s2i-ft-paste-hint">
          {{ isKo ? `A${headerRowNum} 셀에 붙여넣기` : `Paste at A${headerRowNum}` }}
        </span>
        <button
          type="button"
          class="s2i-ft-all-btn"
          :class="{ success: copiedAll }"
          :title="copyAllTitle"
          :aria-label="copyAllLabel"
          @click="copyAll"
        >
          <Icon :icon="copiedAll ? 'lucide:check' : 'lucide:copy'" width="14" height="14" aria-hidden="true" />
          <span>{{ copyAllLabel }}</span>
        </button>
      </div>
    </div>

    <div class="s2i-ft-scroll">
      <table class="s2i-ft-table">
        <thead>
          <tr class="s2i-ft-thead-row">
            <th class="s2i-ft-th-rownum" :title="isKo ? '행 번호' : 'Row number'">
              {{ headerRowNum }}
            </th>
            <th
              v-for="(hdr, cIdx) in excelHeaders"
              :key="cIdx"
              class="s2i-ft-th-excel"
            >
              <span class="s2i-ft-col-badge">{{ colLetters[cIdx] }}</span>
              <span>{{ hdr }}</span>
            </th>
            <th v-if="hasDesc" class="s2i-ft-th-desc">
              {{ isKo ? '설명' : 'Description' }}
            </th>
            <th v-if="hasOut" class="s2i-ft-th-out">
              {{ isKo ? '결과' : 'Result' }}
            </th>
            <th class="s2i-ft-th-action">
              {{ isKo ? '복사' : 'Copy' }}
            </th>
          </tr>
        </thead>
        <tbody>
          <tr
            v-for="(row, idx) in normalizedRows"
            :key="row.rowNum"
            class="s2i-ft-row"
            :class="{ 'is-copied': copiedIndex === idx || copiedAll }"
          >
            <td class="s2i-ft-cell-rownum">{{ row.rowNum }}</td>
            <td
              v-for="(cellVal, cIdx) in row.excelCells"
              :key="cIdx"
              class="s2i-ft-cell-excel"
              :class="{
                'is-empty': !cellVal,
                'is-formula': cellVal.startsWith('='),
                'is-active-copy': cIdx >= row.firstActiveCol
              }"
              :title="
                cellVal
                  ? `${colLetters[cIdx]}${row.rowNum} — ${isKo ? '클릭하여 행 복사' : 'Click to copy row'} (${row.rangeLabel})`
                  : `${colLetters[cIdx]}${row.rowNum} (${isKo ? '빈 셀' : 'Empty cell'})`
              "
              @click="copyRow(idx)"
            >
              <code v-if="cellVal" class="s2i-ft-code">{{ cellVal }}</code>
              <span v-else class="s2i-ft-empty-dash" aria-hidden="true">—</span>
            </td>
            <td v-if="hasDesc" class="s2i-ft-cell-desc">{{ row.desc }}</td>
            <td v-if="hasOut" class="s2i-ft-cell-out">
              <span v-if="row.out" class="s2i-ft-out">{{ row.out }}</span>
            </td>
            <td class="s2i-ft-cell-action">
              <button
                type="button"
                class="s2i-ft-row-btn"
                :class="{ success: copiedIndex === idx || copiedAll }"
                :aria-label="
                  copiedIndex === idx
                    ? (isKo ? '복사됨' : 'Copied')
                    : `${isKo ? '복사' : 'Copy'} ${row.rangeLabel}`
                "
                :title="
                  copiedIndex === idx
                    ? (isKo ? '복사됨' : 'Copied')
                    : `${isKo ? '복사' : 'Copy'} ${row.rangeLabel}`
                "
                @click="copyRow(idx)"
              >
                <Icon
                  :icon="copiedIndex === idx || copiedAll ? 'lucide:check' : 'lucide:copy'"
                  width="14"
                  height="14"
                  aria-hidden="true"
                />
              </button>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
</template>

<style scoped>
.s2i-ft-wrap {
  margin: 16px 0 20px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 12px;
  background: var(--vp-c-bg);
  overflow: hidden;
}

.s2i-ft-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 10px 12px 10px 14px;
  border-bottom: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg-soft);
}

.s2i-ft-title-group {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
}

.s2i-ft-icon {
  display: block;
  flex-shrink: 0;
}

.s2i-ft-title {
  font-size: 0.88rem;
  font-weight: 700;
  color: var(--vp-c-text-1);
  line-height: 1.3;
}

.s2i-ft-range {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  height: 20px;
  padding: 0 7px;
  border-radius: 5px;
  background: var(--s2i-excel-bg);
  color: var(--s2i-excel);
  font-family: var(--vp-font-family-mono);
  font-size: 0.72rem;
  font-weight: 700;
  white-space: nowrap;
}

.s2i-ft-header-actions {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  flex-shrink: 0;
}

.s2i-ft-paste-hint {
  font-size: 0.75rem;
  font-weight: 500;
  color: var(--vp-c-text-2);
  white-space: nowrap;
}

.s2i-ft-all-btn {
  flex-shrink: 0;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  height: 30px;
  padding: 0 11px;
  border-radius: 7px;
  border: 1px solid var(--s2i-excel-border);
  background: var(--s2i-excel);
  color: #ffffff;
  font-size: 0.78rem;
  font-weight: 600;
  cursor: pointer;
  transition: opacity 0.15s ease, background-color 0.15s ease, border-color 0.15s ease;
  user-select: none;
}

.s2i-ft-all-btn:hover {
  opacity: 0.92;
}

.dark .s2i-ft-all-btn {
  color: #0b1912;
}

.s2i-ft-all-btn.success {
  background: #10b981;
  border-color: #10b981;
  color: #ffffff;
}

.s2i-ft-scroll {
  overflow-x: auto;
}

.s2i-ft-table {
  display: table;
  width: 100%;
  margin: 0;
  border-collapse: collapse;
}

.s2i-ft-table th,
.s2i-ft-table td {
  border-bottom: 1px solid var(--vp-c-divider);
  border-right: 1px solid var(--vp-c-divider);
}

.s2i-ft-table th:last-child,
.s2i-ft-table td:last-child {
  border-right: none;
}

.s2i-ft-table tbody tr:last-child td {
  border-bottom: none;
}

.s2i-ft-thead-row th {
  background: var(--vp-c-bg-alt);
  padding: 7px 12px;
  font-size: 0.77rem;
  font-weight: 600;
  color: var(--vp-c-text-2);
  text-align: left;
  white-space: nowrap;
}

.s2i-ft-th-rownum,
.s2i-ft-cell-rownum {
  width: 34px;
  min-width: 34px;
  padding: 8px 6px;
  text-align: center;
  background: var(--vp-c-bg-alt);
  font-family: var(--vp-font-family-mono);
  font-size: 0.75rem;
  font-weight: 600;
  color: var(--vp-c-text-2);
  user-select: none;
}

.s2i-ft-thead-row .s2i-ft-th-rownum {
  text-align: center;
  padding: 7px 6px;
}

.s2i-ft-col-badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 19px;
  height: 19px;
  margin-right: 6px;
  padding: 0 5px;
  border-radius: 4px;
  background: var(--s2i-excel-bg);
  color: var(--s2i-excel);
  font-family: var(--vp-font-family-mono);
  font-size: 0.71rem;
  font-weight: 700;
}

.s2i-ft-th-action,
.s2i-ft-cell-action {
  width: 48px;
  min-width: 48px;
  padding: 8px 10px;
  text-align: center;
  white-space: nowrap;
  vertical-align: middle;
}

.s2i-ft-thead-row .s2i-ft-th-action {
  text-align: center;
  padding: 7px 10px;
}

.s2i-ft-row {
  transition: background-color 0.12s ease;
}

.s2i-ft-cell-excel {
  background: var(--vp-c-bg);
  padding: 8px 12px;
  vertical-align: middle;
  white-space: nowrap;
  cursor: pointer;
  transition: background-color 0.14s ease;
}

.s2i-ft-cell-excel.is-empty {
  text-align: center;
  color: var(--vp-c-text-3);
}

.s2i-ft-empty-dash {
  font-family: var(--vp-font-family-mono);
  font-size: 0.78rem;
  color: var(--vp-c-text-3);
  user-select: none;
}

.s2i-ft-code {
  padding: 0;
  background: transparent;
  color: var(--vp-c-text-1);
  font-family: var(--vp-font-family-mono);
  font-size: 0.83rem;
  font-weight: 500;
  white-space: nowrap;
}

.s2i-ft-cell-excel.is-formula .s2i-ft-code {
  color: var(--s2i-excel);
  font-weight: 600;
}

.s2i-ft-row:hover .s2i-ft-cell-excel.is-active-copy,
.s2i-ft-row.is-copied .s2i-ft-cell-excel.is-active-copy {
  background: var(--s2i-excel-bg);
}

.s2i-ft-row:hover .s2i-ft-cell-rownum,
.s2i-ft-row.is-copied .s2i-ft-cell-rownum {
  color: var(--s2i-excel);
}

.s2i-ft-cell-desc {
  min-width: 170px;
  width: 100%;
  padding: 8px 12px;
  font-size: 0.84rem;
  color: var(--vp-c-text-2);
  line-height: 1.45;
  vertical-align: middle;
  background: var(--vp-c-bg-soft);
}

.s2i-ft-cell-out {
  padding: 8px 12px;
  vertical-align: middle;
  white-space: nowrap;
  background: var(--vp-c-bg-soft);
}

.s2i-ft-out {
  display: inline-block;
  padding: 2px 7px;
  border-radius: 5px;
  background: var(--vp-c-default-soft);
  color: var(--vp-c-text-1);
  font-family: var(--vp-font-family-mono);
  font-size: 0.76rem;
  white-space: nowrap;
}

.s2i-ft-cell-action {
  background: var(--vp-c-bg-soft);
}

.s2i-ft-row-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 28px;
  height: 28px;
  padding: 0;
  border-radius: 6px;
  border: 1px solid var(--vp-c-divider);
  background: var(--vp-c-bg);
  color: var(--vp-c-text-2);
  cursor: pointer;
  transition: color 0.15s ease, border-color 0.15s ease, background-color 0.15s ease;
  user-select: none;
}

.s2i-ft-row-btn:hover {
  color: var(--s2i-excel);
  border-color: var(--s2i-excel-border);
}

.s2i-ft-row-btn.success {
  color: #10b981;
  border-color: #10b981;
  background: rgba(16, 185, 129, 0.1);
}

@media (max-width: 640px) {
  .s2i-ft-paste-hint {
    display: none;
  }
}
</style>
