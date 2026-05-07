/** Replace placeholder / template cells with plausible demo values (not real entities). */

const COMPANIES = [
  'Nordic Refractory Group AS',
  'Pacific Slag & Steel Co.',
  'Baltic Furnace Materials Ltd.',
  'Sunrise Glass Industries',
  'Continental EAF Services',
  'Highland Ceramics GmbH',
  'Verde Industrial Minerais',
  'Sterling Heatwork Solutions',
]

const CITIES = [
  'Oslo, Norway',
  'Busan, South Korea',
  'Mumbai, India',
  'Düsseldorf, Germany',
  'Monterrey, Mexico',
  'Atlanta, USA',
  'Lyon, France',
  'Jeddah, Saudi Arabia',
]

const CONTACT_NAMES = [
  'Jordan Lee',
  'Sam Rivera',
  'Priya Nair',
  'Marco Hoffmann',
  'Taylor Brooks',
  'Elena Vogt',
  'James Okonkwo',
  'Rina Tanaka',
]

const TITLES = [
  'VP Procurement',
  'Plant Manager',
  'Chief Engineer',
  'Category Manager',
  'Supply Chain Director',
]

const INDUSTRIES = [
  'Iron & Steel',
  'Glass',
  'Electrical & Electronics',
  'Cement & Lime',
]

const SEGMENTS = [
  'EAF Operations',
  'Heating Elements',
  'Glass Furnace',
  'Ladle Refractory',
]

const STATUS = ['Operating', 'Expansion', 'New Project', 'Pilot line']

const PRODUCTS = [
  'Standard Fused Magnesia',
  'White Fused Magnesia',
  'Electrical Grade Fused Magnesia',
]

const EMAIL_DOMAIN = 'example.com'

function isPlaceholder(s: string): boolean {
  const t = s.trim().toLowerCase()
  if (!t) return true
  if (/^customer\s*xx\s*$/i.test(s.trim())) return true
  return (
    t === 'x' ||
    t === 'xx' ||
    t === 'xxx' ||
    t === 'n/a' ||
    t === 'tbd' ||
    t === '—' ||
    t === '-'
  )
}

function substitute(
  value: string | number,
  rowIdx: number,
  colIdx: number,
  hintRaw: string
): string | number {
  if (typeof value === 'number') return value

  const s = String(value).trim()
  const hint = (hintRaw || '').toLowerCase()

  /** S.No. is always column 0; handle before `return s` so literals like 60 / "Demo 34-1" are not kept */
  if (colIdx === 0) {
    if (/^\d+$/.test(s)) return parseInt(s, 10)
    return rowIdx + 1
  }

  const customerMatch = /^customer\s*(\d+)\s*$/i.exec(s)
  if (customerMatch) {
    const n = Math.max(0, parseInt(customerMatch[1], 10) - 1)
    return COMPANIES[(n + rowIdx) % COMPANIES.length]
  }

  if (!isPlaceholder(s) && s.length > 0) return s

  if (hint.includes('company name') || hint.includes('parent company')) {
    return COMPANIES[rowIdx % COMPANIES.length]
  }
  if (hint.includes('year of establishment') || hint.includes('establishment')) {
    return String(1982 + ((rowIdx * 5 + colIdx * 3) % 38))
  }
  if (hint.includes('headquarters') || hint.includes('hq')) {
    return CITIES[(rowIdx + colIdx) % CITIES.length]
  }
  if (
    hint.includes('revenue') ||
    hint.includes('us$') ||
    hint.includes('usd')
  ) {
    return String(95 + rowIdx * 42 + colIdx * 7)
  }
  if (hint.includes('end-use industry') || hint.includes('end use')) {
    return INDUSTRIES[rowIdx % INDUSTRIES.length]
  }
  if (hint.includes('sub-segment') || hint.includes('sub segment')) {
    return SEGMENTS[(rowIdx + colIdx) % SEGMENTS.length]
  }
  if (hint.includes('facility status')) {
    return STATUS[rowIdx % STATUS.length]
  }
  if (hint.includes('linkedin')) {
    return `https://linkedin.com/in/demo-procurement-${rowIdx + 1}`
  }
  if (hint.includes('email')) {
    return `buyer.demo${rowIdx + 1}@${EMAIL_DOMAIN}`
  }
  if (hint.includes('phone') || hint.includes('mobile')) {
    return `+1-555-010${String(rowIdx + 1).padStart(2, '0')}`
  }
  if (hint.includes('designation') || hint.includes('title')) {
    return TITLES[(rowIdx + colIdx) % TITLES.length]
  }
  if (
    hint === 'name' ||
    (hint.includes('name') &&
      (hint.includes('contact') ||
        hint.includes('decision') ||
        hint.includes('maker')))
  ) {
    return CONTACT_NAMES[rowIdx % CONTACT_NAMES.length]
  }
  if (hint.includes('required product') || hint.includes('product')) {
    return PRODUCTS[(rowIdx + colIdx) % PRODUCTS.length]
  }
  if (hint.includes('grade')) {
    return ['Low Silicon–High Calcium', 'Large-Crystal', 'Double-Step'][
      rowIdx % 3
    ]
  }
  if (hint.includes('mgo content') || hint.includes('content required')) {
    return ['96%–97.9%', '98%–98.9%', '99% and above'][rowIdx % 3]
  }
  if (hint.includes('form required') || hint.includes('lumps')) {
    return ['Lumps', 'Granules', 'Fine powder'][rowIdx % 3]
  }
  if (hint.includes('electrical')) {
    return ['Insulation strength focus', 'Thermal conductivity focus'][
      rowIdx % 2
    ]
  }
  if (hint.includes('certification')) {
    return 'COA / TDS (demo)'
  }
  if (hint.includes('buying model') || hint.includes('purchase')) {
    return ['Direct purchase', 'Distributor-led', 'Annual agreement'][
      rowIdx % 3
    ]
  }
  if (hint.includes('frequency')) {
    return ['Monthly', 'Quarterly', 'Annual'][rowIdx % 3]
  }
  if (hint.includes('supplier') || hint.includes('competitor')) {
    return `Incumbent supplier ${String.fromCharCode(65 + (rowIdx % 26))} (demo)`
  }
  if (hint.includes('customer size')) {
    return ['Small', 'Medium', 'Large', 'Strategic account'][rowIdx % 4]
  }
  if (hint.includes('annual purchase') || hint.includes('volume')) {
    return ['Medium (demo)', '800 t/yr (demo)', 'High (demo)'][rowIdx % 3]
  }
  if (hint.includes('switching')) {
    return ['Low', 'Medium', 'High'][rowIdx % 3]
  }

  if (colIdx === 1) return COMPANIES[rowIdx % COMPANIES.length]

  return `Demo ${rowIdx + 1}-${colIdx + 1}`
}

export function buildColumnHints(
  grid: unknown[][],
  maxCol: number
): string[] {
  const r6 = (grid[6] as unknown[]) || []
  const r5 = (grid[5] as unknown[]) || []
  const out: string[] = []
  for (let c = 0; c < maxCol; c++) {
    const a =
      r6[c] !== '' && r6[c] != null ? String(r6[c]).trim() : ''
    const b =
      r5[c] !== '' && r5[c] != null ? String(r5[c]).trim() : ''
    out.push((a || b).toLowerCase())
  }
  return out
}

export function applyDemoBodyRows(
  bodyRows: (string | number)[][],
  columnHints: string[]
): (string | number)[][] {
  return bodyRows.map((row, ri) =>
    row.map((cell, ci) => substitute(cell, ri, ci, columnHints[ci] ?? ''))
  )
}
