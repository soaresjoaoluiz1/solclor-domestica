/**
 * ENTRADA DE LEADS — LP REVENDEDOR SOLCLOR (produtos.solclor.com.br)
 *
 * Recebe POST do form, salva TUDO na planilha, e SE qualificado
 * (tem R$ 1500 pra investir) encaminha pro CRM Dros conta 'solclor'
 * com apenas 2 tags oficiais: "Revendedor Linha Domestica" e "Revendedor Piscinas".
 * Todo o restante (perfil, ja_revende, como_pretende, UTMs, etc) entra em
 * observacoes/notes do lead, nao em tags.
 *
 * NAO enviar CRM se qualificado === false (marcou "Ainda nao tenho esse valor").
 */

var SHEET_NAME = 'Pagina1' // ou 'Sheet1' se estiver em ingles
var CRM_WEBHOOK_URL = 'https://drosagencia.com.br/crm/api/webhooks/sheets/solclor'

// Tag oficial escolhida pela LP de origem (identificada pela URL do form)
// - piscinas.solclor.com.br -> "Revendedor Piscinas"
// - produtos.solclor.com.br (ou qualquer outra) -> "Revendedor Linha Domestica"
function resolveTagsFromUrl(url) {
  var u = String(url || '').toLowerCase()
  if (u.indexOf('piscinas.solclor') >= 0) return ['Revendedor Piscinas']
  if (u.indexOf('produtos.solclor') >= 0) return ['Revendedor Linha Domestica']
  // Fallback: sem URL reconhecivel, marca as duas pra nao perder o lead
  return ['Revendedor Linha Domestica', 'Revendedor Piscinas']
}

var HEADERS = [
  'timestamp',           // 1
  'nome',                // 2
  'telefone',            // 3
  'cidade_estado',       // 4
  'cpf_cnpj',            // 5
  'tem_1500',            // 6
  'como_pretende',       // 7
  'qualificado',         // 8  SIM/NAO
  'origem',              // 9
  'url',                 // 10
  'referrer',            // 11
  'user_agent',          // 12
  'utm_source',          // 13
  'utm_medium',          // 14
  'utm_campaign',        // 15
  'utm_content',         // 16
  'utm_term',            // 17
  'fbp',                 // 18
  'fbc',                 // 19
  'fbclid',              // 20
  'gclid',               // 21
  'ctwa_clid',           // 22
  'event_id',            // 23
  'crm_status',          // 24
  'crm_lead_id',         // 25
  'crm_response',        // 26
]

function doPost(e) {
  try {
    var data = {}
    try { data = JSON.parse(e.postData.contents) } catch (err) {
      data = e.parameter || {}
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0]

    // Garante cabecalhos na primeira linha
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS)
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f0f0f0')
      sh.setFrozenRows(1)
    }

    // ─── Interpreta qualificacao ─────────────────────────
    // O front manda `qualificado: true/false` explicitamente.
    // Fallback: se nao vier, deduz do texto do `tem_1500`.
    var tem1500Text = String(data.tem_1500 || '')
    var qualificado
    if (data.qualificado === true || data.qualificado === 'true') {
      qualificado = true
    } else if (data.qualificado === false || data.qualificado === 'false') {
      qualificado = false
    } else {
      // Fallback pra compatibilidade: se contem "nao tenho" ou "ainda nao"
      var lower = tem1500Text.toLowerCase()
      qualificado = !(lower.indexOf('ainda nao') >= 0 || lower.indexOf('ainda não') >= 0 || lower.indexOf('nao tenho') >= 0 || lower.indexOf('não tenho') >= 0)
    }

    var eventId = data.event_id || _uuid_()

    // ─── Escreve linha na planilha (SEMPRE, qualificado ou nao) ─
    var row = [
      new Date().toISOString(),
      data.nome || '',
      data.telefone || '',
      data.cidade_estado || '',
      data.cpf_cnpj || '',
      tem1500Text,
      data.como_pretende || '',
      qualificado ? 'SIM' : 'NAO',
      data.origem || 'lp-revendedor-solclor',
      data.url || data.page_url || '',
      data.referrer || '',
      data.user_agent || '',
      data.utm_source || '',
      data.utm_medium || '',
      data.utm_campaign || '',
      data.utm_content || '',
      data.utm_term || '',
      data.fbp || '',
      data.fbc || '',
      data.fbclid || '',
      data.gclid || '',
      data.ctwa_clid || '',
      eventId,
      '', // crm_status
      '', // crm_lead_id
      ''  // crm_response
    ]
    sh.appendRow(row)
    var rowIndex = sh.getLastRow()
    var COL = { crm_status: 24, crm_lead_id: 25, crm_response: 26 }

    // ─── Se NAO qualificado, para aqui: NAO envia CRM ────
    if (!qualificado) {
      sh.getRange(rowIndex, COL.crm_status).setValue('pulado_sem_1500')
      return _json({ ok: true, qualificado: false, sent_to_crm: false })
    }

    // ─── Monta observacoes ricas (tudo que NAO eh tag) ───
    var notesLines = []
    notesLines.push('Cadastro pela LP /revendedor em ' + _brDate_())
    notesLines.push('R$ 1.500 pra investir: ' + tem1500Text)
    if (data.como_pretende) notesLines.push('Como pretende trabalhar: ' + data.como_pretende)
    if (data.cpf_cnpj) notesLines.push('CPF/CNPJ: ' + data.cpf_cnpj)
    if (data.cidade_estado) notesLines.push('Cidade/UF: ' + data.cidade_estado)
    if (data.utm_source || data.utm_medium || data.utm_campaign) {
      notesLines.push('Origem trafego: ' +
        [data.utm_source, data.utm_medium, data.utm_campaign].filter(function(x){return x}).join(' / '))
    }
    if (data.utm_content) notesLines.push('Ad content: ' + data.utm_content)
    if (data.fbclid) notesLines.push('Meta Ads (fbclid presente)')
    if (data.gclid) notesLines.push('Google Ads (gclid presente)')
    if (data.referrer) notesLines.push('Referrer: ' + data.referrer)

    // ─── Source detail curto pro card do CRM ─────────────
    var detailBits = ['LP: /revendedor']
    if (data.utm_campaign) detailBits.push('camp=' + data.utm_campaign)
    if (data.utm_content) detailBits.push('ad=' + data.utm_content)
    if (data.cidade_estado) detailBits.push(data.cidade_estado)

    // ─── Payload CRM (apenas as 2 tags oficiais) ─────────
    var payload = {
      name: data.nome || '',
      phone: data.telefone || '',
      city: data.cidade_estado || '',
      cpf_cnpj: data.cpf_cnpj || '',
      source: 'lp-revendedor-solclor',
      source_detail: detailBits.join(' | ').substring(0, 500),
      tags: resolveTagsFromUrl(data.url || data.page_url || ''),
      notes: notesLines.join('\n'),
      observacoes: notesLines.join('\n'),
      interesse: 'Revendedor autorizado — kit inicial R$ 1.500',

      // UTMs completos (o CRM pode usar pra atribuicao)
      utm_source: data.utm_source || '',
      utm_medium: data.utm_medium || '',
      utm_campaign: data.utm_campaign || '',
      utm_content: data.utm_content || '',
      utm_term: data.utm_term || '',

      // Click IDs
      fbp: data.fbp || '',
      fbc: data.fbc || '',
      fbclid: data.fbclid || '',
      gclid: data.gclid || '',
      ctwa_clid: data.ctwa_clid || '',

      user_agent: data.user_agent || '',
      page_url: data.url || data.page_url || '',
      event_id: eventId,
    }

    var crmResp
    try {
      var httpResp = UrlFetchApp.fetch(CRM_WEBHOOK_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        followRedirects: true
      })
      var code = httpResp.getResponseCode()
      var text = httpResp.getContentText()
      var respJson = null
      try { respJson = JSON.parse(text) } catch (err) {}
      var leadId = (respJson && (respJson.leadId || respJson.lead_id)) || ''

      if (code >= 200 && code < 300) {
        sh.getRange(rowIndex, COL.crm_status).setValue('ok_' + code)
        sh.getRange(rowIndex, COL.crm_lead_id).setValue(leadId)
        sh.getRange(rowIndex, COL.crm_response).setValue(text.substring(0, 400))
        crmResp = { ok: true, code: code, leadId: leadId }
      } else {
        sh.getRange(rowIndex, COL.crm_status).setValue('erro_' + code)
        sh.getRange(rowIndex, COL.crm_response).setValue(text.substring(0, 400))
        crmResp = { ok: false, code: code, body: text.substring(0, 200) }
      }
    } catch (err) {
      sh.getRange(rowIndex, COL.crm_status).setValue('excecao')
      sh.getRange(rowIndex, COL.crm_response).setValue(String(err).substring(0, 400))
      crmResp = { ok: false, error: String(err) }
    }

    return _json({ ok: true, qualificado: true, sent_to_crm: true, crm: crmResp })
  } catch (err) {
    return _json({ ok: false, error: String(err) })
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}

function _uuid_() {
  var hex = '0123456789abcdef'
  var s = ''
  for (var i = 0; i < 32; i++) {
    var r = Math.floor(Math.random() * 16)
    if (i === 12) r = 4
    if (i === 16) r = (r & 0x3) | 0x8
    s += hex[r]
    if (i === 7 || i === 11 || i === 15 || i === 19) s += '-'
  }
  return s
}

function _brDate_() {
  var d = new Date()
  return Utilities.formatDate(d, 'America/Sao_Paulo', "dd/MM/yyyy 'as' HH:mm")
}

// Util: rodar 1x manualmente pra criar cabecalhos (Executar > initSheet)
function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0]
  sh.clear()
  sh.appendRow(HEADERS)
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f0f0f0')
  sh.setFrozenRows(1)
}

// Testes rapidos
function testQualificadoProdutos() {
  // Simula lead da LP produtos.solclor.com.br -> deve receber tag "Revendedor Linha Domestica"
  var e = { postData: { contents: JSON.stringify({
    nome: 'TESTE Qualificado Produtos',
    telefone: '48999998888',
    cidade_estado: 'Sao Paulo / SP',
    cpf_cnpj: '12345678909',
    tem_1500: 'Sim, tenho e quero começar',
    como_pretende: 'Revenda para amigos, familiares e vizinhos',
    qualificado: true,
    origem: 'lp-revendedor-solclor',
    url: 'https://produtos.solclor.com.br/?utm_source=facebook',
    utm_source: 'facebook', utm_medium: 'cpc', utm_campaign: 'teste',
    user_agent: 'Mozilla Test',
  })}}
  Logger.log(doPost(e).getContent())
}
function testQualificadoPiscinas() {
  // Simula lead da LP piscinas.solclor.com.br -> deve receber tag "Revendedor Piscinas"
  var e = { postData: { contents: JSON.stringify({
    nome: 'TESTE Qualificado Piscinas',
    telefone: '48999996666',
    cidade_estado: 'Curitiba / PR',
    cpf_cnpj: '11122233344',
    tem_1500: 'Sim, tenho e quero começar',
    como_pretende: 'Vendas para clientes da minha regiao',
    qualificado: true,
    origem: 'lp-piscinas-solclor',
    url: 'https://piscinas.solclor.com.br/?utm_source=google',
    utm_source: 'google', utm_medium: 'cpc', utm_campaign: 'piscinas-poa',
    user_agent: 'Mozilla Test',
  })}}
  Logger.log(doPost(e).getContent())
}
function testNaoQualificado() {
  var e = { postData: { contents: JSON.stringify({
    nome: 'TESTE Nao Qualificado',
    telefone: '48999997777',
    cidade_estado: 'Rio de Janeiro / RJ',
    cpf_cnpj: '98765432100',
    tem_1500: 'Ainda não tenho esse valor disponível',
    como_pretende: 'Quero avaliar a melhor forma de começar',
    qualificado: false,
    origem: 'lp-revendedor-solclor',
  })}}
  Logger.log(doPost(e).getContent())
}
