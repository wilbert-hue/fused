import fs from 'fs/promises'
import path from 'path'
import Image from 'next/image'
import {
  filterCmiSheetsProposition3Only,
  getCmiExcelFilename,
  parseCmiWorkbookFromBuffer,
  type CmiHeaderCell,
  type CmiSheetModel,
} from '@/lib/cmi-excel'

function dashboardSubtitleFromSheets(sheets: CmiSheetModel[]): string {
  const title = sheets[0]?.banner.title ?? ''
  if (/dead\s*burned/i.test(title)) return 'Dead burned Magnesia Buyers'
  if (/fused magnesia/i.test(title)) return 'Fused Magnesia Buyers'
  return 'Fused Magnesia Buyers'
}

function PageTitleAndDemoNote() {
  return (
    <div className="mb-5 flex flex-col items-stretch gap-2">
      <h2 className="text-xl font-bold text-gray-900">
        Customer Intelligence Database
      </h2>
      <p className="w-full text-sm font-normal normal-case leading-relaxed text-amber-950 bg-amber-50 border border-orange-200 rounded-md px-3 py-2.5 text-left">
        <span className="font-semibold">NOTE:</span> All the data in the
        dashboard is demo data. No real world data is related to this.
      </p>
    </div>
  )
}

function headerCellClass(cell: CmiHeaderCell): string {
  const wrap =
    'border border-black px-3 py-2 text-center align-middle text-gray-900 leading-snug whitespace-normal break-words [overflow-wrap:anywhere]'
  if (cell.variant === 'sno')
    return `${wrap} bg-[#f9e79f] font-semibold whitespace-nowrap min-w-[3.25rem] w-14 px-2`
  if (cell.variant === 'leaf')
    return `${wrap} bg-[#e8f5e9] text-xs font-semibold min-w-[12rem]`
  return `${wrap} bg-[#e8f5e9] text-xs font-semibold min-w-[12rem]`
}

function bodyCellClass(columnIndex: number): string {
  const base =
    'border border-black px-3 py-2 align-top bg-white text-gray-900 break-words [overflow-wrap:anywhere]'
  if (columnIndex === 0) {
    return `${base} w-14 min-w-[3.25rem] max-w-[4rem] text-center tabular-nums whitespace-nowrap`
  }
  return `${base} min-w-[12rem] whitespace-normal`
}

function CmiSingleTable({ sheet }: { sheet: CmiSheetModel }) {
  const hasTable =
    sheet.headerRows.length > 0 && sheet.columnCount > 0

  return (
    <div className="rounded-lg border border-gray-200 bg-white shadow-sm">
      <h2 className="text-base font-semibold text-gray-900 px-4 py-3 border-b border-gray-200 bg-white">
        {sheet.displayTitle}
      </h2>
      <div className="p-3 sm:p-4 bg-gray-100">
        {!hasTable ? (
          <p className="text-sm text-gray-600">No table structure in this sheet.</p>
        ) : (
          <div className="rounded-md border border-gray-300 bg-white">
            <div className="shrink-0 bg-[#2c3e50] px-4 py-3 text-right text-white w-full">
              <div className="text-sm font-semibold leading-tight whitespace-normal [overflow-wrap:anywhere]">
                {sheet.banner.title}
              </div>
              {sheet.banner.subtitle ? (
                <div className="mt-1 text-xs leading-snug text-white/90 whitespace-normal [overflow-wrap:anywhere]">
                  {sheet.banner.subtitle}
                </div>
              ) : null}
            </div>
            {sheet.headerStripTitle ? (
              <div className="shrink-0 w-full border-x border-b border-black bg-[#e8f5e9] px-4 py-2 text-right text-sm font-semibold text-gray-900 whitespace-normal [overflow-wrap:anywhere]">
                {sheet.headerStripTitle}
              </div>
            ) : null}
            <div className="w-full overflow-x-auto overscroll-x-contain [scrollbar-gutter:stable]">
              <table
                className={`w-max border-collapse border-l border-r border-b border-black text-sm text-gray-900 ${
                  sheet.headerStripTitle ? '' : 'border-t border-black'
                }`}
              >
                <thead>
                  {sheet.headerRows.map((row, ri) => (
                    <tr key={ri}>
                      {row.cells.map((cell, ci) => (
                        <th
                          key={`${ri}-${ci}`}
                          scope="col"
                          rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                          colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                          className={headerCellClass(cell)}
                        >
                          {cell.text || '\u00a0'}
                        </th>
                      ))}
                    </tr>
                  ))}
                </thead>
                <tbody>
                  {sheet.bodyRows.map((row, ri) => (
                    <tr key={ri}>
                      {Array.from({ length: sheet.columnCount }, (_, ci) => (
                        <td key={ci} className={bodyCellClass(ci)}>
                          {row[ci] === '' || row[ci] == null ? (
                            <span className="text-gray-500">—</span>
                          ) : (
                            String(row[ci])
                          )}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

export default async function DashboardPage() {
  const filePath = path.join(process.cwd(), getCmiExcelFilename())
  let sheets: CmiSheetModel[] = []
  let loadError: string | null = null

  try {
    const buf = await fs.readFile(filePath)
    sheets = filterCmiSheetsProposition3Only(parseCmiWorkbookFromBuffer(buf))
  } catch (e) {
    loadError =
      e instanceof Error ? e.message : 'Could not read the CMI Excel file.'
  }

  const dashboardSubtitle =
    sheets.length > 0
      ? dashboardSubtitleFromSheets(sheets)
      : 'Fused Magnesia Buyers'

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Top bar — logo left, titles centered (Coherent-style) */}
      <header className="bg-white border-b border-gray-200">
        <div className="mx-auto flex w-full max-w-[2000px] items-center gap-4 px-4 sm:px-6 py-5">
          <div className="flex w-[clamp(140px,26vw,200px)] shrink-0 justify-start">
            <Image
              src="/logo.png"
              alt="Coherent Market Insights"
              width={180}
              height={72}
              className="h-auto w-auto max-w-[180px]"
              priority
            />
          </div>
          <div className="min-w-0 flex-1 px-2 text-center">
            <h1 className="text-2xl font-bold text-gray-900 md:text-3xl">
              Coherent Dashboard
            </h1>
            <p className="mt-1 text-sm text-gray-500 md:text-base">
              {dashboardSubtitle}
            </p>
          </div>
          <div
            className="hidden w-[clamp(140px,26vw,200px)] shrink-0 sm:block"
            aria-hidden
          />
        </div>
      </header>

      <div className="mx-auto w-full max-w-[2000px] px-4 sm:px-6 py-6">
        {loadError ? (
          <div>
            <PageTitleAndDemoNote />
          <div
            className="rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-800"
            role="alert"
          >
            <p className="font-medium">Unable to load workbook</p>
            <p className="mt-1">{loadError}</p>
            <p className="mt-2 text-red-700">
              Place{' '}
              <code className="rounded bg-red-100 px-1 py-0.5 text-xs">
                {getCmiExcelFilename()}
              </code>{' '}
              in the project root.
            </p>
          </div>
          </div>
        ) : (
          <div className="flex flex-col gap-6">
            <div className="rounded-lg border border-gray-200 bg-white p-4 shadow-sm">
              <h2 className="text-xs font-bold uppercase tracking-wider text-gray-500 mb-3">
                Chart view
              </h2>
              <div className="rounded-lg border border-sky-200 border-l-[6px] border-l-teal-600 bg-sky-50/90 p-3 shadow-sm max-w-xl">
                <div className="flex items-start gap-2">
                  <span className="text-lg" aria-hidden>
                    👤
                  </span>
                  <div>
                    <div className="text-sm font-semibold text-sky-950">
                      Customer Intelligence
                    </div>
                    <p className="mt-1 text-xs text-sky-900/85 leading-snug">
                      Proposition 3 buyer database (single table, full width).
                    </p>
                  </div>
                </div>
              </div>
            </div>

            <main className="space-y-6 min-w-0 w-full">
              <PageTitleAndDemoNote />
              {sheets.length === 0 ? (
                <p className="text-gray-600">
                  No Preposition 3 / Proposition 3 sheet found. Add a worksheet
                  whose name includes that tab (e.g. &quot;Preposition 3&quot;)
                  in the CMI Excel file.
                </p>
              ) : (
                sheets.map((sheet) => (
                  <CmiSingleTable key={sheet.sheetName} sheet={sheet} />
                ))
              )}
            </main>
          </div>
        )}
      </div>
    </div>
  )
}
