/* eslint-disable no-alert */

const STORAGE_KEY = 'flavp_web_state_v1'

// Opções existentes no XLSX atual (B9 e B12 nas abas de compra)
const DEFAULT_STATUS_OPTIONS = ['Carregando', 'Viajando', 'Concluído']
const DEFAULT_TIPO_OPTIONS = ['Moradas', 'VALCATORCE INCA']

function uid(prefix = 'id') {
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now()}`
}

function toIsoDate(value) {
  if (!value) return ''
  const d = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(d.getTime())) return ''
  return d.toISOString().slice(0, 10)
}

function fmtMoney(n) {
  const num = Number(n || 0)
  return num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
}

function fmtNum(n) {
  const num = Number(n || 0)
  return num.toLocaleString('pt-BR', { maximumFractionDigits: 6 })
}

function safeNumber(v) {
  if (v === null || v === undefined || v === '') return null
  const n = Number(v)
  return Number.isFinite(n) ? n : null
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return { purchases: [], ledger: [], meta: { statusOptions: [], tipoOptions: [] } }
    const parsed = JSON.parse(raw)
    if (!parsed || typeof parsed !== 'object') return { purchases: [], ledger: [], meta: { statusOptions: [], tipoOptions: [] } }
    return {
      purchases: Array.isArray(parsed.purchases) ? parsed.purchases : [],
      ledger: Array.isArray(parsed.ledger) ? parsed.ledger : [],
      meta:
        parsed.meta && typeof parsed.meta === 'object'
          ? {
              statusOptions: Array.isArray(parsed.meta.statusOptions) ? parsed.meta.statusOptions : [],
              tipoOptions: Array.isArray(parsed.meta.tipoOptions) ? parsed.meta.tipoOptions : [],
            }
          : { statusOptions: [], tipoOptions: [] },
    }
  } catch {
    return { purchases: [], ledger: [], meta: { statusOptions: [], tipoOptions: [] } }
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state))
}

function calcPurchase(purchase) {
  const txComercial = Number(purchase.txComercial || 0)
  const txParalelo = Number(purchase.txParalelo || 0)
  const taxaPresumido = Number(purchase.taxaPresumido ?? 0.0291)

  const custoFreteTotal = Number(purchase.custoFreteTotal || 0) // I20
  const despachoTotal = Number(purchase.despachoTotal || 0) // J20
  const taxaCCambio = Number(purchase.taxaCCambio || 0) // D21

  const lucroOperacional = Number(purchase.lucroOperacional || 0) // L24
  const despesasOperacionais = Number(purchase.despesasOperacionais || 0) // K32 (ou fórmula no modelo)

  const items = (purchase.items || []).map((it) => ({ ...it }))

  const totalQty = items.reduce((s, it) => s + Number(it.quantidade || 0), 0)
  const freteUnit = totalQty > 0 ? custoFreteTotal / totalQty : 0
  const despachoUnit = totalQty > 0 ? despachoTotal / totalQty : 0

  // Nota de fidelidade:
  // - A planilha calcula D e F por unidade (sem multiplicar pela quantidade)
  // - Totais na linha 20 são ponderados pela quantidade
  // - K20 no Excel = G20 + I20 + J20 + D21 (não inclui H20 explicitamente)

  for (const it of items) {
    const dollar = Number(it.dollar || 0)
    const dollarPf = Number(it.dollarPf || 0)
    const venda = Number(it.venda || 0)

    it.faturaEm = dollar * txComercial // D
    it.pfEm = dollarPf * txParalelo // F
    it.custo = it.faturaEm + it.pfEm // G

    // H, I, J unitários: na planilha, I e J dependem apenas do total/quantidade.
    it.custoFrete = freteUnit
    it.despacho = despachoUnit

    it.totalFronteira = it.custo + it.custoFrete + it.despacho // (H não modelado aqui)

    it.presumido = venda * taxaPresumido // M
    it.valorFinal = venda + it.presumido // O
  }

  const totals = {
    totalQty,
    totalDollar: items.reduce((s, it) => s + Number(it.quantidade || 0) * Number(it.dollar || 0), 0),
    totalDollarPf: items.reduce((s, it) => s + Number(it.quantidade || 0) * Number(it.dollarPf || 0), 0),
    totalFaturaEm: items.reduce((s, it) => s + Number(it.quantidade || 0) * Number(it.faturaEm || 0), 0),
    totalPfEm: items.reduce((s, it) => s + Number(it.quantidade || 0) * Number(it.pfEm || 0), 0),
    totalCusto: items.reduce((s, it) => s + Number(it.quantidade || 0) * Number(it.custo || 0), 0),
    totalVenda: items.reduce((s, it) => s + Number(it.quantidade || 0) * Number(it.venda || 0), 0),
    totalPresumido: items.reduce((s, it) => s + Number(it.quantidade || 0) * Number(it.presumido || 0), 0),
    totalValorFinal: items.reduce((s, it) => s + Number(it.quantidade || 0) * Number(it.valorFinal || 0), 0),
  }

  // K20 (Total Fronteira) conforme fórmula do Excel
  const K20 = totals.totalCusto + custoFreteTotal + despachoTotal + taxaCCambio

  // Bloco final (23-33)
  const K23 = totals.totalValorFinal
  const N23 = Number(purchase.taxaPresumido ?? 0.0291)
  const L23 = K23 * N23
  const K24 = K23 - L23
  const K25 = K24 - lucroOperacional

  const K29 = K20
  const K30 = L23
  const K31 = lucroOperacional

  const K32 = purchase.modeloCebola
    ? totals.totalValorFinal - K29 - K30 - K31
    : despesasOperacionais

  const K33 = K29 + K30 + K31 + K32

  return {
    ...purchase,
    items,
    computed: {
      ...totals,
      K20,
      K23,
      L23,
      K24,
      K25,
      K29,
      K30,
      K31,
      K32,
      K33,
    },
  }
}

function recalcAll() {
  state.purchases = state.purchases.map(calcPurchase)

  // Recalcula saldos do caixa sequencialmente
  let saldo = 0
  state.ledger = state.ledger.map((row, idx) => {
    const deb = Number(row.debito || 0)
    const cred = Number(row.credito || 0)
    saldo = idx === 0 ? deb - cred : saldo + deb - cred
    return { ...row, saldo }
  })

  saveState()
  render()
}

function setTab(name) {
  document.querySelectorAll('.tab').forEach((b) => b.classList.toggle('is-active', b.dataset.tab === name))
  document.querySelectorAll('[data-panel]').forEach((p) => {
    p.hidden = p.dataset.panel !== name
  })
}

function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag)
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v)
    else if (v === true) node.setAttribute(k, '')
    else if (v !== false && v !== null && v !== undefined) node.setAttribute(k, String(v))
  }
  for (const c of children) {
    node.append(c)
  }
  return node
}

function renderKpis() {
  const purchases = state.purchases
  const ledger = state.ledger

  const totalCompras = purchases.length
  const totalQty = purchases.reduce((s, p) => s + Number(p.computed?.totalQty || 0), 0)
  const totalCusto = purchases.reduce((s, p) => s + Number(p.computed?.K20 || 0), 0)
  const totalLucro = purchases.reduce((s, p) => s + Number(p.lucroOperacional || 0), 0)
  const totalValorFinal = purchases.reduce((s, p) => s + Number(p.computed?.totalValorFinal || 0), 0)
  const saldoAtual = ledger.length ? Number(ledger[ledger.length - 1].saldo || 0) : 0

  const kpis = [
    { label: 'Compras', value: totalCompras },
    { label: 'Quantidade Total', value: fmtNum(totalQty) },
    { label: 'Custo Fronteira Total', value: fmtMoney(totalCusto) },
    { label: 'Lucro Operacional', value: fmtMoney(totalLucro) },
    { label: 'Valor Final Total', value: fmtMoney(totalValorFinal) },
    { label: 'Saldo Atual (Caixa)', value: fmtMoney(saldoAtual) },
  ]

  const root = document.getElementById('kpis')
  if (!root) return
  root.innerHTML = ''
  for (const kpi of kpis) {
    root.append(
      el('div', { class: 'kpi' }, [
        el('div', { class: 'label' }, [document.createTextNode(kpi.label)]),
        el('div', { class: 'value' }, [document.createTextNode(String(kpi.value))]),
      ]),
    )
  }
}

function sortPurchasesForList(purchases) {
  return [...purchases].sort((a, b) => {
    const da = new Date(a.data || 0).getTime() || 0
    const db = new Date(b.data || 0).getTime() || 0
    return db - da
  })
}

function renderPurchaseTable() {
  const table = document.getElementById('purchaseTable')
  if (!table) return

  let rows = sortPurchasesForList(state.purchases)
  
  // Apply status filter
  if (state.ui.statusFilter && state.ui.statusFilter !== 'Todos') {
    rows = rows.filter(p => p.status === state.ui.statusFilter)
  }
  
  table.innerHTML = ''

  table.append(
    el('thead', {}, [
      el('tr', {}, [
        el('th', {}, [document.createTextNode('Data')]),
        el('th', {}, [document.createTextNode('Referência')]),
        el('th', {}, [document.createTextNode('Status')]),
        el('th', {}, [document.createTextNode('Produto/Tipo')]),
        el('th', {}, [document.createTextNode('Qtd Total')]),
        el('th', {}, [document.createTextNode('Custo Fronteira')]),
        el('th', {}, [document.createTextNode('Custo Conta Corrente')]),
        el('th', {}, [document.createTextNode('Ações')]),
      ]),
    ]),
  )

  table.append(
    el(
      'tbody',
      {},
      rows.map((p) => {
        const row = el('tr', {}, [
          el('td', {}, [document.createTextNode(toIsoDate(p.data))]),
          el('td', {}, [document.createTextNode(p.referencia || '—')]),
          el('td', {}, [document.createTextNode(p.status || '—')]),
          el('td', {}, [document.createTextNode(p.tipoCebola || '—')]),
          el('td', {}, [document.createTextNode(fmtNum(p.computed?.totalQty || 0))]),
          el('td', {}, [document.createTextNode(fmtMoney(p.computed?.K20 || 0))]),
          el('td', {}, [document.createTextNode(fmtMoney(p.computed?.K33 || 0))]),
          el('td', { onclick: (e) => e.stopPropagation() }, [
            el('div', { class: 'table-actions' }, [
              el(
                'button',
                {
                  class: 'btn btn-secondary',
                  onclick: () => {
                    // Store current tab before navigating
                    const currentPanel = document.querySelector('[data-panel]:not([hidden])')
                    if (currentPanel) state.ui.previousTab = currentPanel.dataset.panel
                    state.ui.selectedPurchaseId = p.id
                    setTab('purchases')
                    render()
                  },
                },
                [document.createTextNode('Editar')],
              ),
              el(
                'button',
                {
                  class: 'btn btn-danger',
                  onclick: () => {
                    if (!confirm('Excluir esta compra?')) return
                    state.purchases = state.purchases.filter((x) => x.id !== p.id)
                    if (state.ui.selectedPurchaseId === p.id) state.ui.selectedPurchaseId = state.purchases[0]?.id ?? null
                    recalcAll()
                  },
                },
                [document.createTextNode('Excluir')],
              ),
            ]),
          ]),
        ])
        row.addEventListener('click', () => {
          // Store current tab before navigating
          const currentPanel = document.querySelector('[data-panel]:not([hidden])')
          if (currentPanel) state.ui.previousTab = currentPanel.dataset.panel
          state.ui.selectedPurchaseId = p.id
          setTab('purchases')
          render()
        })
        return row
      }),
    ),
  )
}

function inputRow(label, value, type, onChange) {
  const input = el('input', {
    type,
    value: value ?? '',
    oninput: (e) => onChange(e.target.value),
  })
  return el('div', { class: 'field-row' }, [
    el('div', { class: 'label' }, [document.createTextNode(label)]),
    input,
  ])
}

function uniqueNonEmptyStrings(values) {
  const set = new Set()
  for (const v of values) {
    const s = String(v ?? '').trim()
    if (!s) continue
    set.add(s)
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'pt-BR'))
}

function uniqueUnionStrings(a, b) {
  return uniqueNonEmptyStrings([...(a || []), ...(b || [])])
}

function selectRow(label, value, options, onChange) {
  const current = String(value ?? '').trim()
  const opts = Array.isArray(options) ? [...options] : []
  if (current && !opts.includes(current)) opts.unshift(current)

  const select = el(
    'select',
    {
      onchange: (e) => onChange(e.target.value),
    },
    [
      el('option', { value: '' }, [document.createTextNode('—')]),
      ...opts.map((o) => el('option', { value: o, selected: o === current }, [document.createTextNode(o)])),
    ],
  )

  return el('div', { class: 'field-row' }, [
    el('div', { class: 'label' }, [document.createTextNode(label)]),
    select,
  ])
}

function renderPurchaseEditor() {
  const root = document.getElementById('purchaseEditor')
  const btnNewPurchase = document.getElementById('btnNewPurchase')
  const purchaseHint = document.getElementById('purchaseHint')
  const purchaseTitle = document.getElementById('purchaseTitle')
  const purchaseRow = document.getElementById('purchaseRow')
  const purchaseSection = document.getElementById('purchaseSection')
  const editorTitle = document.getElementById('editorTitle')
  
  // Check if editing temp (new) purchase or existing purchase
  const originalPurchase = state.purchases.find((p) => p.id === state.ui.selectedPurchaseId)
  const isNewPurchase = !originalPurchase && state.ui.tempPurchase && state.ui.tempPurchase.id === state.ui.selectedPurchaseId
  
  let purchase = null
  if (isNewPurchase) {
    purchase = state.ui.tempPurchase
  } else if (originalPurchase) {
    // Create editing copy if not exists
    if (!state.ui.editingPurchase || state.ui.editingPurchase.id !== originalPurchase.id) {
      state.ui.editingPurchase = JSON.parse(JSON.stringify(originalPurchase))
    }
    purchase = state.ui.editingPurchase
  }

  if (!purchase) {
    root.className = 'muted'
    root.textContent = 'Nenhuma compra selecionada.'
    if (btnNewPurchase) btnNewPurchase.style.display = ''
    if (purchaseHint) purchaseHint.style.display = ''
    if (purchaseTitle) purchaseTitle.style.display = ''
    if (purchaseRow) purchaseRow.style.display = ''
    if (purchaseSection) purchaseSection.style.marginTop = ''
    if (purchaseSection) purchaseSection.style.paddingTop = ''
    if (purchaseSection) purchaseSection.style.borderTop = ''
    if (editorTitle) editorTitle.style.display = ''
    return
  }

  // Hide title, row, section spacing and editor title when editing
  if (purchaseTitle) purchaseTitle.style.display = 'none'
  if (purchaseRow) purchaseRow.style.display = 'none'
  if (purchaseSection) purchaseSection.style.marginTop = '0'
  if (purchaseSection) purchaseSection.style.paddingTop = '0'
  if (purchaseSection) purchaseSection.style.borderTop = 'none'
  if (editorTitle) editorTitle.style.display = 'none'
  if (btnNewPurchase) btnNewPurchase.style.display = 'none'
  if (purchaseHint) purchaseHint.style.display = 'none'

  root.className = ''
  root.innerHTML = ''

  const header = el('div', { style: 'margin-bottom: 24px;' }, [
    el('div', { class: 'form-row' }, [
      el('button', {
        class: 'btn btn-secondary',
        onclick: () => {
          state.ui.selectedPurchaseId = null
          state.ui.tempPurchase = null
          state.ui.editingPurchase = null
          setTab(state.ui.previousTab || 'list')
          render()
        },
      }, [document.createTextNode('Voltar')]),
      el('button', {
        class: 'btn',
        onclick: () => {
          if (isNewPurchase) {
            // Add new purchase to list
            const calculated = calcPurchase(purchase)
            state.purchases.unshift(calculated)
            state.ui.tempPurchase = null
          } else {
            // Update existing purchase
            const idx = state.purchases.findIndex(p => p.id === purchase.id)
            if (idx >= 0) {
              state.purchases[idx] = calcPurchase(purchase)
            }
            state.ui.editingPurchase = null
          }
          recalcAll()
          alert('Salvo com sucesso!')
        },
      }, [document.createTextNode('Salvar')]),
      isNewPurchase ? el('button', {
        class: 'btn btn-secondary',
        onclick: () => {
          state.ui.selectedPurchaseId = null
          state.ui.tempPurchase = null
          state.ui.editingPurchase = null
          render()
        },
      }, [document.createTextNode('Cancelar')]) : el('button', {
        class: 'btn btn-danger',
        onclick: () => {
          if (!confirm('Excluir esta compra?')) return
          state.purchases = state.purchases.filter((p) => p.id !== purchase.id)
          state.ui.selectedPurchaseId = null
          state.ui.editingPurchase = null
          recalcAll()
        },
      }, [document.createTextNode('Excluir')]),
    ]),
  ])

  const fields = el('div', { class: 'excel-grid' }, [
    el('div', { class: 'excel-section' }, [
      el('div', { class: 'excel-section-title' }, [document.createTextNode('Dados Gerais')]),
      inputRow('Número Compra', purchase.numeroCompra, 'number', (v) => (purchase.numeroCompra = safeNumber(v))),
      inputRow('Referência', purchase.referencia, 'text', (v) => (purchase.referencia = v)),
      inputRow('Data', toIsoDate(purchase.data), 'date', (v) => (purchase.data = v)),
      selectRow('Status', purchase.status, state.meta?.statusOptions || [], (v) => (purchase.status = v)),
      selectRow('Produto/Tipo', purchase.tipoCebola, state.meta?.tipoOptions || [], (v) => (purchase.tipoCebola = v)),
    ]),
    el('div', { class: 'excel-section' }, [
      el('div', { class: 'excel-section-title' }, [document.createTextNode('Fornecedor & Transporte')]),
      inputRow('Fornecedor', purchase.fornecedor, 'text', (v) => (purchase.fornecedor = v)),
      inputRow('Exportador', purchase.exportador, 'text', (v) => (purchase.exportador = v)),
      inputRow('Importador', purchase.importador, 'text', (v) => (purchase.importador = v)),
      inputRow('Transportadora', purchase.transportadora, 'text', (v) => (purchase.transportadora = v)),
    ]),
    el('div', { class: 'excel-section' }, [
      el('div', { class: 'excel-section-title' }, [document.createTextNode('Taxas de Câmbio')]),
      inputRow('Tx Comercial (B13)', purchase.txComercial, 'number', (v) => (purchase.txComercial = safeNumber(v))),
      inputRow('Tx Paralelo (B14)', purchase.txParalelo, 'number', (v) => (purchase.txParalelo = safeNumber(v))),
      inputRow('Taxa C-Câmbio (D21)', purchase.taxaCCambio, 'number', (v) => (purchase.taxaCCambio = safeNumber(v))),
    ]),
    el('div', { class: 'excel-section' }, [
      el('div', { class: 'excel-section-title' }, [document.createTextNode('Custos & Operação')]),
      inputRow('Frete Total (I20)', purchase.custoFreteTotal, 'number', (v) => (purchase.custoFreteTotal = safeNumber(v))),
      inputRow('Despacho Total (J20)', purchase.despachoTotal, 'number', (v) => (purchase.despachoTotal = safeNumber(v))),
      inputRow('Lucro Operacional (L24)', purchase.lucroOperacional, 'number', (v) => (purchase.lucroOperacional = safeNumber(v))),
      inputRow('Despesas Operacionais (K32)', purchase.despesasOperacionais, 'number', (v) => (purchase.despesasOperacionais = safeNumber(v))),
    ]),
  ])

  const itemsTable = el('div', { class: 'table-wrap' }, [
    el('table', { class: 'table' }, [
      el('thead', {}, [
        el('tr', {}, [
          el('th', {}, [document.createTextNode('Caixa')]),
          el('th', {}, [document.createTextNode('Qtd (B)')]),
          el('th', {}, [document.createTextNode('Dollar (C)')]),
          el('th', {}, [document.createTextNode('Dollar PF (E)')]),
          el('th', {}, [document.createTextNode('Venda (L)')]),
          el('th', {}, [document.createTextNode('Total Fronteira (K un.)')]),
          el('th', {}, [document.createTextNode('Valor Final (O un.)')]),
        ]),
      ]),
      el(
        'tbody',
        {},
        purchase.items.map((it) => {
          const makeCell = (type, val, setter) =>
            el('td', {}, [
              el('input', {
                type,
                value: val ?? '',
                oninput: (e) => {
                  setter(e.target.value)
                },
              }),
            ])

          return el('tr', {}, [
            el('td', {}, [
              el('input', {
                type: 'text',
                value: it.caixa ?? '',
                oninput: (e) => (it.caixa = e.target.value),
              }),
            ]),
            makeCell('number', it.quantidade ?? '', (v) => (it.quantidade = safeNumber(v))),
            makeCell('number', it.dollar ?? '', (v) => (it.dollar = safeNumber(v))),
            makeCell('number', it.dollarPf ?? '', (v) => (it.dollarPf = safeNumber(v))),
            makeCell('number', it.venda ?? '', (v) => (it.venda = safeNumber(v))),
            el('td', {}, [document.createTextNode(fmtMoney(it.totalFronteira || 0))]),
            el('td', {}, [document.createTextNode(fmtMoney(it.valorFinal || 0))]),
          ])
        }),
      ),
    ]),
  ])

  const computed = calcPurchase(purchase).computed
  const computedBox = el('div', { class: 'panel-note' }, [
    document.createTextNode(
      `Totais: Qtd=${fmtNum(computed.totalQty)} | Valor Final=${fmtMoney(computed.totalValorFinal)} | Custo Fronteira=${fmtMoney(
        computed.K20,
      )} | Custo Conta Corrente=${fmtMoney(computed.K33)}`,
    ),
  ])

  root.append(header, fields, itemsTable, computedBox)
}

function renderLedger() {
  const table = document.getElementById('ledgerTable')
  if (!table) return

  const monthSel = document.getElementById('ledgerMonth')
  const yearSel = document.getElementById('ledgerYear')
  const kpiRoot = document.getElementById('ledgerKpis')

  const now = new Date()
  const selectedMonth = Number(state.ui.ledgerMonth ?? now.getMonth() + 1)
  const selectedYear = Number(state.ui.ledgerYear ?? now.getFullYear())

  // init selects
  if (monthSel && !monthSel.childElementCount) {
    const months = [
      'Janeiro',
      'Fevereiro',
      'Março',
      'Abril',
      'Maio',
      'Junho',
      'Julho',
      'Agosto',
      'Setembro',
      'Outubro',
      'Novembro',
      'Dezembro',
    ]
    monthSel.append(
      ...months.map((m, idx) =>
        el('option', { value: String(idx + 1), selected: idx + 1 === selectedMonth }, [document.createTextNode(m)]),
      ),
    )
    monthSel.addEventListener('change', (e) => {
      state.ui.ledgerMonth = Number(e.target.value)
      render()
    })
  }

  const allDates = state.ledger
    .map((r) => new Date(r.data || 0))
    .filter((d) => !Number.isNaN(d.getTime()))
  const years = Array.from(new Set(allDates.map((d) => d.getFullYear()))).sort((a, b) => b - a)
  const yearsToShow = years.length ? years : [now.getFullYear()]

  if (yearSel && !yearSel.childElementCount) {
    yearSel.append(
      ...yearsToShow.map((y) => el('option', { value: String(y), selected: y === selectedYear }, [document.createTextNode(String(y))])),
    )
    yearSel.addEventListener('change', (e) => {
      state.ui.ledgerYear = Number(e.target.value)
      render()
    })
  }

  if (monthSel) monthSel.value = String(selectedMonth)
  if (yearSel) {
    const y = yearsToShow.includes(selectedYear) ? selectedYear : yearsToShow[0]
    state.ui.ledgerYear = y
    if (String(yearSel.value) !== String(y) || yearSel.options.length !== yearsToShow.length) {
      yearSel.innerHTML = ''
      yearSel.append(...yearsToShow.map((yy) => el('option', { value: String(yy), selected: yy === y }, [document.createTextNode(String(yy))])))
    }
  }

  const start = new Date(Date.UTC(Number(state.ui.ledgerYear), selectedMonth - 1, 1, 0, 0, 0))
  const end = new Date(Date.UTC(Number(state.ui.ledgerYear), selectedMonth, 0, 23, 59, 59))

  const rows = state.ledger.filter((r) => {
    const d = new Date(r.data || 0)
    if (Number.isNaN(d.getTime())) return false
    return d.getTime() >= start.getTime() && d.getTime() <= end.getTime()
  })

  const totalDeb = rows.reduce((s, r) => s + Number(r.debito || 0), 0)
  const totalCred = rows.reduce((s, r) => s + Number(r.credito || 0), 0)
  const saldoAteFimDoMes = (() => {
    const upto = state.ledger
      .filter((r) => {
        const d = new Date(r.data || 0)
        if (Number.isNaN(d.getTime())) return false
        return d.getTime() <= end.getTime()
      })
      .at(-1)
    return Number(upto?.saldo || 0)
  })()

  if (kpiRoot) {
    const items = [
      { label: 'Total Débito (mês)', value: fmtMoney(totalDeb) },
      { label: 'Total Crédito (mês)', value: fmtMoney(totalCred) },
      { label: 'Saldo (fim do mês)', value: fmtMoney(saldoAteFimDoMes) },
      { label: 'Lançamentos (mês)', value: String(rows.length) },
    ]
    kpiRoot.innerHTML = ''
    for (const k of items) {
      kpiRoot.append(
        el('div', { class: 'kpi' }, [
          el('div', { class: 'label' }, [document.createTextNode(k.label)]),
          el('div', { class: 'value' }, [document.createTextNode(String(k.value))]),
        ]),
      )
    }
  }

  table.innerHTML = ''
  table.append(
    el('thead', {}, [
      el('tr', {}, [
        el('th', {}, [document.createTextNode('Data')]),
        el('th', {}, [document.createTextNode('Histórico')]),
        el('th', {}, [document.createTextNode('Débito')]),
        el('th', {}, [document.createTextNode('Crédito')]),
        el('th', {}, [document.createTextNode('Saldo')]),
      ]),
    ]),
  )

  table.append(
    el(
      'tbody',
      {},
      rows.map((r) =>
        el('tr', {}, [
          el('td', {}, [document.createTextNode(toIsoDate(r.data))]),
          el('td', {}, [document.createTextNode(r.historico || '')]),
          el('td', {}, [document.createTextNode(fmtMoney(r.debito || 0))]),
          el('td', {}, [document.createTextNode(fmtMoney(r.credito || 0))]),
          el('td', {}, [document.createTextNode(fmtMoney(r.saldo || 0))]),
        ]),
      ),
    ),
  )
}

function renderLedgerPurchases() {
  const table = document.getElementById('ledgerPurchasesTable')
  if (!table) return

  const selectedMonth = Number(state.ui.ledgerMonth ?? new Date().getMonth() + 1)
  const selectedYear = Number(state.ui.ledgerYear ?? new Date().getFullYear())
  const start = new Date(Date.UTC(selectedYear, selectedMonth - 1, 1, 0, 0, 0))
  const end = new Date(Date.UTC(selectedYear, selectedMonth, 0, 23, 59, 59))

  const filtered = state.purchases.filter((p) => {
    const d = new Date(p.data || 0)
    if (Number.isNaN(d.getTime())) return false
    return d.getTime() >= start.getTime() && d.getTime() <= end.getTime()
  })

  table.innerHTML = ''
  table.append(
    el('thead', {}, [
      el('tr', {}, [
        el('th', {}, [document.createTextNode('Data')]),
        el('th', {}, [document.createTextNode('Referência')]),
        el('th', {}, [document.createTextNode('Status')]),
        el('th', {}, [document.createTextNode('Custo Fronteira')]),
        el('th', {}, [document.createTextNode('Custo C. Corrente')]),
      ]),
    ]),
  )

  table.append(
    el(
      'tbody',
      {},
      filtered.map((p) => {
        const row = el('tr', { style: 'cursor: pointer;' }, [
          el('td', {}, [document.createTextNode(toIsoDate(p.data))]),
          el('td', {}, [document.createTextNode(p.referencia || '—')]),
          el('td', {}, [document.createTextNode(p.status || '—')]),
          el('td', {}, [document.createTextNode(fmtMoney(p.computed?.K20 || 0))]),
          el('td', {}, [document.createTextNode(fmtMoney(p.computed?.K33 || 0))]),
        ])
        row.addEventListener('click', () => {
          state.ui.selectedPurchaseId = p.id
          setTab('purchases')
          render()
        })
        return row
      }),
    ),
  )
}

function withCanvasScale(canvas, draw) {
  if (!canvas) return
  const rect = canvas.getBoundingClientRect()
  if (!rect.width || !rect.height) return
  const dpr = Math.min(window.devicePixelRatio || 1, 2)
  const w = Math.max(10, Math.floor(rect.width))
  const h = Math.max(10, Math.floor(rect.height))
  if (canvas.width !== Math.floor(w * dpr) || canvas.height !== Math.floor(h * dpr)) {
    canvas.width = Math.floor(w * dpr)
    canvas.height = Math.floor(h * dpr)
  }
  const ctx = canvas.getContext('2d')
  if (!ctx) return
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
  draw(ctx, w, h)
}

function drawDonut(canvas, segments) {
  const colors = ['rgba(125,211,252,0.85)', 'rgba(168,85,247,0.70)', 'rgba(52,211,153,0.70)', 'rgba(251,191,36,0.70)']
  const safe = segments.filter((s) => Number(s.value || 0) > 0)
  withCanvasScale(canvas, (ctx, w, h) => {
    ctx.clearRect(0, 0, w, h)
    const cx = Math.floor(w / 2)
    const cy = Math.floor(h / 2)
    const r = Math.min(w, h) * 0.36
    const inner = r * 0.62
    const total = safe.reduce((s, x) => s + Number(x.value || 0), 0) || 1
    let a = -Math.PI / 2

    ctx.beginPath()
    ctx.arc(cx, cy, r, 0, Math.PI * 2)
    ctx.arc(cx, cy, inner, Math.PI * 2, 0, true)
    ctx.closePath()
    ctx.fillStyle = 'rgba(255,255,255,0.05)'
    ctx.fill()

    safe.forEach((seg, idx) => {
      const v = Number(seg.value || 0)
      const da = (v / total) * Math.PI * 2
      ctx.beginPath()
      ctx.arc(cx, cy, r, a, a + da)
      ctx.arc(cx, cy, inner, a + da, a, true)
      ctx.closePath()
      ctx.fillStyle = colors[idx % colors.length]
      ctx.fill()
      a += da
    })

    ctx.fillStyle = 'rgba(232,240,255,0.92)'
    ctx.font = '700 16px system-ui'
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.fillText('Total', cx, cy - 8)
    ctx.font = '700 15px system-ui'
    ctx.fillText(fmtMoney(total), cx, cy + 14)
  })
}

function drawBars(canvas, items) {
  withCanvasScale(canvas, (ctx, w, h) => {
    ctx.clearRect(0, 0, w, h)
    const top = 10
    const left = 10
    const right = 10
    const bottom = 26
    const chartW = w - left - right
    const chartH = h - top - bottom

    const values = items.map((x) => Number(x.value || 0))
    const max = Math.max(1, ...values)
    const n = Math.max(1, items.length)
    const gap = 8
    const barW = Math.max(10, Math.floor((chartW - gap * (n - 1)) / n))

    ctx.strokeStyle = 'rgba(255,255,255,0.08)'
    ctx.lineWidth = 1
    for (let i = 0; i <= 4; i++) {
      const y = top + (chartH * i) / 4
      ctx.beginPath()
      ctx.moveTo(left, y)
      ctx.lineTo(left + chartW, y)
      ctx.stroke()
    }

    items.forEach((it, i) => {
      const v = Number(it.value || 0)
      const bh = Math.max(0, (v / max) * chartH)
      const x = left + i * (barW + gap)
      const y = top + chartH - bh

      ctx.fillStyle = 'rgba(125,211,252,0.75)'
      ctx.fillRect(x, y, barW, bh)
      ctx.fillStyle = 'rgba(232,240,255,0.75)'
      ctx.font = '12px system-ui'
      ctx.textAlign = 'center'
      ctx.textBaseline = 'top'
      ctx.fillText(String(it.label || ''), x + barW / 2, top + chartH + 6)
    })
  })
}

function renderDonutLegend(segments) {
  const container = document.getElementById('donutLegend')
  if (!container) return
  const colors = ['rgba(125,211,252,0.85)', 'rgba(168,85,247,0.70)', 'rgba(52,211,153,0.70)', 'rgba(251,191,36,0.70)']
  container.innerHTML = ''
  segments.forEach((seg, idx) => {
    const item = el('div', { class: 'legend-item' }, [
      el('div', { class: 'legend-color', style: `background: ${colors[idx % colors.length]}` }, []),
      el('div', { class: 'legend-label' }, [document.createTextNode(`${seg.label}: ${fmtMoney(seg.value)}`)]),
    ])
    const isActive = !state.ui.chartFilters.donut.includes(seg.label)
    if (!isActive) item.classList.add('inactive')
    item.addEventListener('click', () => {
      const idx = state.ui.chartFilters.donut.indexOf(seg.label)
      if (idx >= 0) {
        state.ui.chartFilters.donut.splice(idx, 1)
      } else {
        state.ui.chartFilters.donut.push(seg.label)
      }
      render()
    })
    container.append(item)
  })
}

function renderBarsLegend(items) {
  const container = document.getElementById('barsLegend')
  if (!container) return
  container.innerHTML = ''
  const total = items.reduce((s, i) => s + Number(i.value || 0), 0)
  const legendItem = el('div', { class: 'legend-item' }, [
    el('div', { class: 'legend-color', style: 'background: rgba(125,211,252,0.85)' }, []),
    el('div', { class: 'legend-label' }, [document.createTextNode(`Quantidade Total: ${total.toLocaleString('pt-BR')}`)]),
  ])
  container.append(legendItem)
}

function renderCharts() {
  const purchases = sortPurchasesForList(state.purchases)

  const totalK20 = purchases.reduce((s, p) => s + Number(p.computed?.K20 || 0), 0)
  const totalPresumido = purchases.reduce((s, p) => s + Number(p.computed?.L23 || 0), 0)
  const totalLucro = purchases.reduce((s, p) => s + Number(p.lucroOperacional || 0), 0)
  const totalDespesas = purchases.reduce((s, p) => s + Number(p.computed?.K32 || 0), 0)

  const allSegments = [
    { label: 'Custo Fronteira', value: totalK20 },
    { label: 'Presumido', value: totalPresumido },
    { label: 'Lucro Operacional', value: totalLucro },
    { label: 'Despesas Operacionais', value: totalDespesas },
  ]
  const activeSegments = allSegments.filter(s => !state.ui.chartFilters.donut.includes(s.label))
  drawDonut(document.getElementById('chartDonut'), activeSegments)
  renderDonutLegend(allSegments)

  // Agregar quantidade por caixa
  const caixaMap = new Map()
  purchases.forEach((p) => {
    p.items?.forEach((it) => {
      const key = it.caixa || 'Sem Caixa'
      caixaMap.set(key, (caixaMap.get(key) || 0) + Number(it.quantidade || 0))
    })
  })
  const items = Array.from(caixaMap.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([caixa, qty]) => ({ label: caixa, value: qty }))
  drawBars(document.getElementById('chartBars'), items)
  renderBarsLegend(items)
}

function render() {
  renderKpis()
  renderPurchaseEditor()
  renderPurchaseTable()
  renderLedger()
  renderLedgerPurchases()
  renderCharts()
}

function getCell(sheet, addr) {
  const c = sheet[addr]
  if (!c) return null
  return c.v ?? null
}

function getDate(sheet, addr) {
  const cell = sheet[addr]
  if (!cell) return ''
  const v = cell.v
  if (v == null) return ''
  if (v instanceof Date) return toIsoDate(v)
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v)
    if (!d) return ''
    const js = new Date(Date.UTC(d.y, (d.m || 1) - 1, d.d || 1, d.H || 0, d.M || 0, d.S || 0))
    return toIsoDate(js)
  }
  const parsed = new Date(v)
  return toIsoDate(parsed)
}

function parsePurchaseSheet(workbook, sheetName, { modeloCebola }) {
  const sheet = workbook.Sheets[sheetName]
  if (!sheet) return null

  const purchase = {
    id: uid('purchase'),
    sheetName,
    modeloCebola: Boolean(modeloCebola),

    numeroCompra: safeNumber(getCell(sheet, 'B2')),
    referencia: String(getCell(sheet, 'B3') ?? ''),
    data: getDate(sheet, 'B4'),
    fornecedor: String(getCell(sheet, 'B5') ?? ''),
    exportador: String(getCell(sheet, 'B6') ?? ''),
    proformaNum: String(getCell(sheet, 'B7') ?? ''),
    faturaNum: String(getCell(sheet, 'B8') ?? ''),
    status: String(getCell(sheet, 'B9') ?? ''),
    importador: String(getCell(sheet, 'B10') ?? ''),
    transportadora: String(getCell(sheet, 'B11') ?? ''),
    tipoCebola: String(getCell(sheet, 'B12') ?? ''),

    txComercial: safeNumber(getCell(sheet, 'B13')),
    txParalelo: safeNumber(getCell(sheet, 'B14')),

    custoFreteTotal: safeNumber(getCell(sheet, 'I20')),
    despachoTotal: safeNumber(getCell(sheet, 'J20')),
    taxaCCambio: safeNumber(getCell(sheet, 'D21')),

    taxaPresumido: safeNumber(getCell(sheet, 'N23')) ?? 0.0291,
    lucroOperacional: safeNumber(getCell(sheet, 'L24')),
    despesasOperacionais: safeNumber(getCell(sheet, 'K32')),

    nfvf: String(getCell(sheet, 'P20') ?? ''),

    items: [],
  }

  // Itens fixos (Caixa 2..5) nas linhas 16-19
  const itemRows = [16, 17, 18, 19]
  for (const r of itemRows) {
    const caixa = String(getCell(sheet, `A${r}`) ?? '')
    const quantidade = safeNumber(getCell(sheet, `B${r}`))
    const dollar = safeNumber(getCell(sheet, `C${r}`))
    const dollarPf = safeNumber(getCell(sheet, `E${r}`))
    const venda = safeNumber(getCell(sheet, `L${r}`))

    purchase.items.push({
      id: uid('item'),
      caixa,
      quantidade,
      dollar,
      dollarPf,
      venda,
    })
  }

  return calcPurchase(purchase)
}

function parseLedgerSheet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName]
  if (!sheet) return []

  // A partir da linha 6
  const rows = []
  for (let r = 6; r <= 200; r++) {
    const data = getCell(sheet, `A${r}`)
    const historico = getCell(sheet, `B${r}`)
    const debito = getCell(sheet, `C${r}`)
    const credito = getCell(sheet, `D${r}`)

    if (data == null && historico == null && debito == null && credito == null) continue

    rows.push({
      id: uid('ledger'),
      data: getDate(sheet, `A${r}`),
      historico: historico ? String(historico) : '',
      debito: safeNumber(debito),
      credito: safeNumber(credito),
      saldo: 0,
    })
  }

  return rows
}

async function importXlsx(file) {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array', cellFormula: true })

  const purchases = []
  const statusRaw = []
  const tipoRaw = []
  for (const name of wb.SheetNames) {
    if (name.startsWith('FL-AV-P-')) {
      // Coleta opções diretamente do XLSX (mesmo se a estrutura mudar no parse)
      const sh = wb.Sheets[name]
      if (sh) {
        statusRaw.push(getCell(sh, 'B9'))
        tipoRaw.push(getCell(sh, 'B12'))
      }
      purchases.push(parsePurchaseSheet(wb, name, { modeloCebola: false }))
    }
    if (name === 'Modelo Cebola') {
      const sh = wb.Sheets[name]
      if (sh) {
        statusRaw.push(getCell(sh, 'B9'))
        tipoRaw.push(getCell(sh, 'B12'))
      }
      purchases.push(parsePurchaseSheet(wb, name, { modeloCebola: true }))
    }
  }

  const ledger = parseLedgerSheet(wb, 'CAIXA FL-AV-P')

  // Tenta substituir créditos do caixa com K33 das compras (mesmo padrão do Excel)
  const byRef = new Map(purchases.map((p) => [p.referencia, p]))
  const mergedLedger = ledger.map((row) => {
    const refMatch = /FL-AV-P-\d{3}/.exec(row.historico || '')
    if (!refMatch) return row
    const p = byRef.get(refMatch[0])
    if (!p) return row
    return { ...row, credito: p.computed?.K33 ?? row.credito }
  })

  state.purchases = purchases.filter(Boolean)
  state.ledger = mergedLedger

  // Opções dos dropdowns (somente o que existe no XLSX importado)
  state.meta = {
    statusOptions: uniqueNonEmptyStrings(statusRaw),
    tipoOptions: uniqueNonEmptyStrings(tipoRaw),
  }

  // Garante que o dropdown sempre tenha as opções conhecidas do XLSX atual
  state.meta.statusOptions = uniqueUnionStrings(DEFAULT_STATUS_OPTIONS, state.meta.statusOptions)
  state.meta.tipoOptions = uniqueUnionStrings(DEFAULT_TIPO_OPTIONS, state.meta.tipoOptions)

  state.ui.selectedPurchaseId = state.purchases[0]?.id ?? null
  recalcAll()

  return { sheetNames: wb.SheetNames, purchases: state.purchases.length, ledger: state.ledger.length }
}

const state = {
  ...loadState(),
  ui: {
    selectedPurchaseId: null,
    tempPurchase: null,
    editingPurchase: null,
    previousTab: 'list',
    statusFilter: 'Todos',
    ledgerMonth: new Date().getMonth() + 1,
    ledgerYear: new Date().getFullYear(),
    chartFilters: { donut: [], bars: [] },
  },
}

if (!state.meta || typeof state.meta !== 'object') {
  state.meta = { statusOptions: [], tipoOptions: [] }
}

// Compat com localStorage antigo: se não houver meta ainda, usa opções do XLSX atual
state.meta.statusOptions = uniqueUnionStrings(DEFAULT_STATUS_OPTIONS, state.meta.statusOptions)
state.meta.tipoOptions = uniqueUnionStrings(DEFAULT_TIPO_OPTIONS, state.meta.tipoOptions)

// UI wiring

document.getElementById('tabs').addEventListener('click', (e) => {
  const btn = e.target.closest('[data-tab]')
  if (!btn) return
  setTab(btn.dataset.tab)
})

const importLog = document.getElementById('importLog')

document.getElementById('btnImport').addEventListener('click', async () => {
  const file = document.getElementById('fileInput').files?.[0]
  if (!file) {
    alert('Selecione um .xlsx')
    return
  }

  importLog.textContent = 'Importando...\n'
  try {
    const res = await importXlsx(file)
    importLog.textContent += `OK. Planilhas: ${res.sheetNames.join(', ')}\nCompras importadas: ${res.purchases}\nLançamentos no caixa: ${res.ledger}\n`
  } catch (err) {
    importLog.textContent += `ERRO: ${String(err?.message || err)}\n`
    console.error(err)
  }
})

document.getElementById('btnClear').addEventListener('click', () => {
  if (!confirm('Apagar todos os dados salvos (localStorage)?')) return
  localStorage.removeItem(STORAGE_KEY)
  location.reload()
})

document.getElementById('btnNewPurchase').addEventListener('click', () => {
  const p = {
    id: uid('purchase'),
    sheetName: '',
    modeloCebola: false,
    numeroCompra: null,
    referencia: '',
    data: '',
    fornecedor: '',
    exportador: '',
    proformaNum: '',
    faturaNum: '',
    status: '',
    importador: '',
    transportadora: '',
    tipoCebola: '',
    txComercial: null,
    txParalelo: null,
    custoFreteTotal: null,
    despachoTotal: null,
    taxaCCambio: null,
    taxaPresumido: 0.0291,
    lucroOperacional: null,
    despesasOperacionais: null,
    nfvf: '',
    items: [
      { id: uid('item'), caixa: 'Caixa 2', quantidade: null, dollar: null, dollarPf: null, venda: null },
      { id: uid('item'), caixa: 'Caixa 3', quantidade: null, dollar: null, dollarPf: null, venda: null },
      { id: uid('item'), caixa: 'Caixa 4', quantidade: null, dollar: null, dollarPf: null, venda: null },
      { id: uid('item'), caixa: 'Caixa 5', quantidade: null, dollar: null, dollarPf: null, venda: null },
    ],
  }
  // Store as temp purchase for editing
  state.ui.tempPurchase = p
  state.ui.selectedPurchaseId = p.id
  setTab('purchases')
  render()
})

document.getElementById('statusFilter').addEventListener('change', (e) => {
  state.ui.statusFilter = e.target.value
  render()
})

document.getElementById('btnAddLedger').addEventListener('click', () => {
  const historico = prompt('Histórico:')
  if (historico === null) return
  const debito = safeNumber(prompt('Débito (vazio = 0):') || '') || 0
  const credito = safeNumber(prompt('Crédito (vazio = 0):') || '') || 0

  state.ledger.push({
    id: uid('ledger'),
    data: toIsoDate(new Date()),
    historico,
    debito,
    credito,
    saldo: 0,
  })
  recalcAll()
})

const dataLog = document.getElementById('dataLog')

document.getElementById('btnExportJson').addEventListener('click', () => {
  const blob = new Blob([JSON.stringify({ purchases: state.purchases, ledger: state.ledger, meta: state.meta }, null, 2)], {
    type: 'application/json',
  })
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = 'flavp-data.json'
  a.click()
  URL.revokeObjectURL(a.href)
  dataLog.textContent = 'Exportado como flavp-data.json'
})

document.getElementById('btnImportJson').addEventListener('click', async () => {
  const file = document.getElementById('jsonInput').files?.[0]
  if (!file) {
    alert('Selecione um JSON')
    return
  }
  try {
    const txt = await file.text()
    const parsed = JSON.parse(txt)
    state.purchases = Array.isArray(parsed.purchases) ? parsed.purchases : []
    state.ledger = Array.isArray(parsed.ledger) ? parsed.ledger : []
    state.meta =
      parsed.meta && typeof parsed.meta === 'object'
        ? {
            statusOptions: Array.isArray(parsed.meta.statusOptions) ? parsed.meta.statusOptions : [],
            tipoOptions: Array.isArray(parsed.meta.tipoOptions) ? parsed.meta.tipoOptions : [],
          }
        : { statusOptions: [], tipoOptions: [] }

    state.meta.statusOptions = uniqueUnionStrings(DEFAULT_STATUS_OPTIONS, state.meta.statusOptions)
    state.meta.tipoOptions = uniqueUnionStrings(DEFAULT_TIPO_OPTIONS, state.meta.tipoOptions)

    state.ui.selectedPurchaseId = state.purchases[0]?.id ?? null
    recalcAll()
    dataLog.textContent = 'JSON importado com sucesso.'
  } catch (e) {
    dataLog.textContent = `Erro: ${String(e?.message || e)}`
  }
})

// Initial render
state.ui.selectedPurchaseId = state.purchases[0]?.id ?? null
recalcAll()
setTab('dashboard')
