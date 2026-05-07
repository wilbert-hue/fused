import * as XLSX from 'xlsx'
import { applyDemoBodyRows, buildColumnHints } from './cmi-demo-data'

const CMI_FILENAME = 'Demo Customer Intelligence Database_Fused Magnesia Buyers_CMI.xlsx'

/** Excel rows 4–6 (0-indexed) — hierarchical header block */
const HEADER_TOP = 4
const HEADER_BOTTOM = 6
const DATA_START = 7

/** Max data rows per proposition sheet (after parsing / demo fill). */
const MAX_CMI_BODY_ROWS = 40

type Merge = { s: { r: number; c: number }; e: { r: number; c: number } }

export type CmiBanner = {
  title: string
  subtitle: string
}

export type CmiHeaderCell = {
  text: string
  rowSpan: number
  colSpan: number
  variant: 'sno' | 'group' | 'leaf'
}

export type CmiHeaderRow = { cells: CmiHeaderCell[] }

export type CmiSheetModel = {
  sheetName: string
  displayTitle: string
  banner: CmiBanner
  /** Excel row 4 right-rail titles (e.g. Company & Facility Identification) */
  headerStripTitle: string
  headerRows: CmiHeaderRow[]
  columnCount: number
  bodyRows: (string | number)[][]
}

function cellText(sh: XLSX.WorkSheet, r: number, c: number): string {
  const addr = XLSX.utils.encode_cell({ r, c })
  const cell = sh[addr]
  if (!cell) return ''
  if (cell.w != null) return String(cell.w).trim()
  const v = cell.v
  if (v == null || v === '') return ''
  return String(v).trim()
}

function parseBanner(sh: XLSX.WorkSheet): CmiBanner {
  const raw = cellText(sh, 0, 0)
  const lines = raw
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean)
  return {
    title: lines[0] ?? '',
    subtitle: lines.slice(1).join(' ') ?? '',
  }
}

function getMaxCol(sh: XLSX.WorkSheet, grid: unknown[][]): number {
  let max = 0
  const ref = sh['!ref']
  if (ref) {
    const d = XLSX.utils.decode_range(ref)
    max = Math.max(max, d.e.c + 1)
  }
  for (const row of grid) {
    if (Array.isArray(row)) max = Math.max(max, row.length)
  }
  return max
}

/** Rightmost column index covered by any merge that intersects the header block (rows 4–6). */
function maxHeaderMergeRightColumn(merges: Merge[]): number {
  let max = -1
  for (const m of merges) {
    if (m.e.r < HEADER_TOP || m.s.r > HEADER_BOTTOM) continue
    max = Math.max(max, m.e.c)
  }
  return max
}

function columnUsedInBody(
  grid: unknown[][],
  maxCol: number,
  startRow: number
): boolean[] {
  const used = Array(maxCol).fill(false)
  for (let r = startRow; r < grid.length; r++) {
    const row = (grid[r] as unknown[]) || []
    for (let c = 0; c < maxCol; c++) {
      const v = row[c]
      if (v !== '' && v != null && String(v).trim() !== '') used[c] = true
    }
  }
  return used
}

/** Columns present in body but with no real header merges / text — match Excel column N behavior */
function syntheticHeaderMerges(
  sh: XLSX.WorkSheet,
  merges: Merge[],
  maxCol: number,
  bodyUsed: boolean[]
): Merge[] {
  const covered = new Set<number>()
  for (const m of merges) {
    if (m.e.r < HEADER_TOP || m.s.r > HEADER_BOTTOM) continue
    for (let c = m.s.c; c <= m.e.c; c++) {
      if (c >= 0 && c < maxCol) covered.add(c)
    }
  }

  const extra: Merge[] = []
  for (let c = 0; c < maxCol; c++) {
    if (covered.has(c)) continue
    const hasHeaderText =
      !!cellText(sh, HEADER_TOP, c) ||
      !!cellText(sh, HEADER_TOP + 1, c) ||
      !!cellText(sh, HEADER_BOTTOM, c)
    if (!hasHeaderText && bodyUsed[c]) {
      extra.push({
        s: { r: HEADER_TOP, c },
        e: { r: HEADER_BOTTOM, c },
      })
      covered.add(c)
    }
  }
  return extra
}

type CoverCell =
  | { kind: 'start'; merge: Merge }
  | { kind: 'skip' }

function buildHeaderRows(
  sh: XLSX.WorkSheet,
  merges: Merge[],
  maxCol: number
): { headerRows: CmiHeaderRow[]; row4StripParts: string[] } {
  const nHeaderRows = HEADER_BOTTOM - HEADER_TOP + 1
  const cover: (CoverCell | null)[][] = Array.from({ length: nHeaderRows }, () =>
    Array<CoverCell | null>(maxCol).fill(null)
  )

  const relevant = merges.filter(
    (m) => m.s.r >= HEADER_TOP && m.s.r <= HEADER_BOTTOM && m.e.r >= HEADER_TOP
  )

  for (const m of relevant) {
    const r0 = Math.max(m.s.r, HEADER_TOP)
    const r1 = Math.min(m.e.r, HEADER_BOTTOM)
    const c0 = m.s.c
    const c1 = Math.min(m.e.c, maxCol - 1)
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const ir = r - HEADER_TOP
        if (ir < 0 || ir >= nHeaderRows) continue
        if (c < 0 || c >= maxCol) continue
        if (r === m.s.r && c === m.s.c) cover[ir][c] = { kind: 'start', merge: m }
        else cover[ir][c] = { kind: 'skip' }
      }
    }
  }

  const headerRows: CmiHeaderRow[] = []
  const row4StripParts: string[] = []

  for (let ir = 0; ir < nHeaderRows; ir++) {
    const cells: CmiHeaderCell[] = []
    for (let c = 0; c < maxCol; c++) {
      const slot = cover[ir][c]
      if (slot?.kind === 'skip') continue

      if (slot?.kind === 'start') {
        const m = slot.merge
        const text = cellText(sh, m.s.r, m.s.c)
        const rs = Math.min(m.e.r, HEADER_BOTTOM) - m.s.r + 1
        const cs = m.e.c - m.s.c + 1
        const isSno =
          m.s.c === 0 &&
          (text.toLowerCase().includes('s.no') ||
            text.replace(/\s/g, '').toLowerCase().includes('s.no.'))
        /** Top Excel row wide bands → green strip above grid (matches reference UI). */
        const isRow4FullWidthBand =
          ir === 0 &&
          m.s.r === HEADER_TOP &&
          m.e.r === HEADER_TOP &&
          m.s.c > 0 &&
          cs > 1
        if (isRow4FullWidthBand) {
          if (text) row4StripParts.push(text)
          continue
        }
        cells.push({
          text,
          rowSpan: rs,
          colSpan: cs,
          variant: isSno ? 'sno' : 'group',
        })
        continue
      }

      const excelR = HEADER_TOP + ir
      const t = cellText(sh, excelR, c)
      if (!t) continue

      cells.push({
        text: t,
        rowSpan: 1,
        colSpan: 1,
        variant: 'leaf',
      })
    }
    headerRows.push({ cells })
  }

  return { headerRows, row4StripParts }
}

function normalizeBody(
  grid: unknown[][],
  maxCol: number,
  startRow: number
): (string | number)[][] {
  const out: (string | number)[][] = []
  for (let r = startRow; r < grid.length; r++) {
    const row = (grid[r] as unknown[]) || []
    const slice: (string | number)[] = []
    for (let c = 0; c < maxCol; c++) {
      const cell = row[c]
      if (typeof cell === 'number') slice.push(cell)
      else if (cell === '' || cell == null) slice.push('')
      else slice.push(String(cell).trim())
    }
    if (slice.every((v) => v === '' || v === null)) continue
    out.push(slice)
  }
  return out
}

function isSerialNumberHint(hint: string): boolean {
  const t = hint.toLowerCase().replace(/\s/g, '')
  return t.includes('s.no') || t.includes('serialno')
}

/** After slicing rows, make S.No. column 1…n so gaps from the workbook do not show. */
function renumberSnoColumn(
  rows: (string | number)[][],
  columnHints: string[]
): (string | number)[][] {
  if (rows.length === 0 || !isSerialNumberHint(columnHints[0] ?? '')) {
    return rows
  }
  return rows.map((row, ri) => {
    const next = row.slice()
    if (next.length > 0) next[0] = ri + 1
    return next
  })
}

function parseOneSheet(sheetName: string, sh: XLSX.WorkSheet): CmiSheetModel {
  const grid = XLSX.utils.sheet_to_json<unknown[]>(sh, {
    header: 1,
    defval: '',
  })

  const baseMerges = (sh['!merges'] || []) as Merge[]
  let maxCol = getMaxCol(sh, grid as unknown[][])
  const mergeRight = maxHeaderMergeRightColumn(baseMerges)
  if (mergeRight >= 0) {
    maxCol = Math.min(maxCol, mergeRight + 1)
  }
  const bodyUsed = columnUsedInBody(grid as unknown[][], maxCol, DATA_START)
  const synthetic = syntheticHeaderMerges(sh, baseMerges, maxCol, bodyUsed)
  const allMerges = [...baseMerges, ...synthetic]

  const { headerRows, row4StripParts } = buildHeaderRows(sh, allMerges, maxCol)
  const headerStripTitle =
    row4StripParts.length > 0
      ? row4StripParts.join(' · ')
      : cellText(sh, HEADER_TOP, 1)
  const columnHints = buildColumnHints(grid as unknown[][], maxCol)
  const bodyRows = renumberSnoColumn(
    applyDemoBodyRows(
      normalizeBody(grid as unknown[][], maxCol, DATA_START),
      columnHints
    ).slice(0, MAX_CMI_BODY_ROWS),
    columnHints
  )

  return {
    sheetName,
    displayTitle: sheetName.trim(),
    banner: parseBanner(sh),
    headerStripTitle,
    headerRows,
    columnCount: maxCol,
    bodyRows,
  }
}

export function parseCmiWorkbookFromBuffer(buf: Buffer): CmiSheetModel[] {
  const wb = XLSX.read(buf, { type: 'buffer' })
  const names = wb.SheetNames.filter(
    (n) => n.trim().toLowerCase() !== 'home'
  )
  return names.map((name) => {
    const sh = wb.Sheets[name]
    if (!sh) {
      return {
        sheetName: name,
        displayTitle: name.trim(),
        banner: { title: '', subtitle: '' },
        headerStripTitle: '',
        headerRows: [],
        columnCount: 0,
        bodyRows: [],
      }
    }
    return parseOneSheet(name, sh)
  })
}

export function getCmiExcelFilename(): string {
  return CMI_FILENAME
}
