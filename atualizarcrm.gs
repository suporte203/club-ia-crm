// ════════════════════════════════════════════════════════════
// CLUB IA CRM — Script de Atualização Automática
// Versão: 24/08/2026
//
// COMO CONFIGURAR (uma vez só):
// 1. Abra a planilha no Google Sheets
// 2. Clique em Extensões → Apps Script
// 3. Cole este código e salve (Ctrl+S)
// 4. Clique em "Executar" → "configurarScript"
//    (autorize o script quando pedido)
// 5. Em Gatilhos (ícone de relógio), crie:
//    - Função: atualizarCRM
//    - Evento: Baseado em tempo → Semanalmente → Segunda → Entre 9h e 10h
// ════════════════════════════════════════════════════════════

// ─── CONFIGURAÇÕES ─────────────────────────────────────────
const GITHUB_TOKEN = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
const GITHUB_REPO  = 'suporte203/club-ia-crm';
const GITHUB_FILE  = 'data.json';
const SHEET_ID     = '1clS-yQgRe5iZ_IE-CuwKzntMTDJ2HuhMwc3xFS7Mujc';

// ─── FUNÇÃO PRINCIPAL ──────────────────────────────────────
function atualizarCRM() {
  try {
    Logger.log('Iniciando atualização do CRM...');

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = ss.getSheets();

    Logger.log('Abas disponíveis: ' + sheets.map(s => s.getName()).join(', '));

    // Encontra as abas dos Blocos
    let bloco1Sheet = null, bloco2Sheet = null, bloco3Sheet = null;
    for (const sheet of sheets) {
      const name = sheet.getName().toLowerCase();
      if (name.includes('bloco') && name.includes('01') || name.includes('bloco 1')) bloco1Sheet = sheet;
      else if (name.includes('bloco') && name.includes('02') || name.includes('bloco 2')) bloco2Sheet = sheet;
      else if (name.includes('bloco') && name.includes('03') || name.includes('bloco 3')) bloco3Sheet = sheet;
    }

    // Lê os dados de presença de cada Bloco
    const bloco1Data = bloco1Sheet ? lerPresenca(bloco1Sheet, 1) : [];
    const bloco2Data = bloco2Sheet ? lerPresenca(bloco2Sheet, 2) : [];
    const bloco3Data = bloco3Sheet ? lerPresenca(bloco3Sheet, 3) : [];

    // Combina todos os alunos
    const allStudents = [...bloco1Data, ...bloco2Data, ...bloco3Data];

    // Adiciona timestamp
    const now = new Date();
    const updated = [
      String(now.getDate()).padStart(2,'0'),
      String(now.getMonth()+1).padStart(2,'0'),
      now.getFullYear()
    ].join('/');

    const dataJson = JSON.stringify({
      updated: updated,
      students: allStudents
    }, null, 2);

    // Atualiza o arquivo data.json no GitHub
    atualizarGitHub(dataJson);

    Logger.log('CRM atualizado com sucesso! ' + allStudents.length + ' alunos.');

  } catch(e) {
    Logger.log('ERRO: ' + e.message);
    MailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      'Erro na atualização do CRM Club IA',
      'Erro: ' + e.message + '\n\nStack: ' + e.stack
    );
  }
}

// ─── LER PRESENÇA DE UMA ABA ───────────────────────────────
function lerPresenca(sheet, blocoNum) {
  const data = sheet.getDataRange().getValues();
  const students = [];

  let currentCat = 'SÓCIO';
  let idCounter = blocoNum * 1000; // IDs únicos por bloco
  let presColIdx = -1;
  let nomeColIdx = -1;
  let emailColIdx = -1;
  let telefoneColIdx = -1;

  // Detecta colunas pelo cabeçalho
  for (let i = 0; i < data.length; i++) {
    const row = data[i].map(c => String(c).trim().toUpperCase());
    if (row.includes('PRESENÇA') || row.includes('PRESENCA')) {
      presColIdx    = row.findIndex(c => c === 'PRESENÇA' || c === 'PRESENCA');
      nomeColIdx    = row.findIndex(c => c === 'NOME' || c === 'NOME ');
      emailColIdx   = row.findIndex(c => c === 'E-MAIL' || c === 'EMAIL' || c === 'E-MAIL ');
      telefoneColIdx= row.findIndex(c => c === 'TELEFONE' || c === 'TELEFONE ');

      // Lê dados a partir da próxima linha
      for (let j = i + 1; j < data.length; j++) {
        const r = data[j];

        // Detecta seção BLACK
        const anyCell = r.map(c => String(c).trim().toUpperCase());
        if (anyCell.some(c => c === 'BLACK')) {
          currentCat = 'BLACK';
          continue;
        }

        // Linha de dados
        const nome = nomeColIdx >= 0 ? String(r[nomeColIdx]).trim() : '';
        const email = emailColIdx >= 0 ? String(r[emailColIdx]).trim().toLowerCase() : '';
        const tel = telefoneColIdx >= 0 ? String(r[telefoneColIdx]).replace(/\D/g,'').substring(0,15) : '';
        const pres = presColIdx >= 0 ? String(r[presColIdx]).trim().toUpperCase() : '';

        if (!nome || nome === '' || nome.toUpperCase() === 'NOME') continue;

        // Converte presença
        let presValue = null;
        if (pres === 'PRESENTE') presValue = 'present';
        else if (pres === 'FALTOU') presValue = 'absent';
        else if (blocoNum === 3) presValue = 'tbd'; // Bloco 3 ainda não aconteceu

        idCounter++;
        const student = {
          id: idCounter,
          name: nome,
          email: email,
          phone: tel,
          bloco: blocoNum,
          cat: currentCat,
          pres1: blocoNum === 1 ? presValue : null,
          pres2: blocoNum === 2 ? presValue : null,
          pres3: blocoNum === 3 ? presValue : null,
          notes: ''
        };

        students.push(student);
      }
      break; // Parou de processar o cabeçalho
    }
  }

  return students;
}

// ─── ATUALIZAR ARQUIVO NO GITHUB ───────────────────────────
function atualizarGitHub(conteudo) {
  if (!GITHUB_TOKEN) {
    throw new Error('Token do GitHub não configurado! Execute configurarScript() primeiro.');
  }

  const apiUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_FILE}`;

  // Busca o SHA atual do arquivo (necessário para atualizar)
  const getResp = UrlFetchApp.fetch(apiUrl, {
    method: 'GET',
    headers: {
      'Authorization': 'token ' + GITHUB_TOKEN,
      'Accept': 'application/vnd.github.v3+json'
    },
    muteHttpExceptions: true
  });

  let sha = null;
  if (getResp.getResponseCode() === 200) {
    const fileData = JSON.parse(getResp.getContentText());
    sha = fileData.sha;
    Logger.log('SHA atual: ' + sha);
  } else {
    Logger.log('Arquivo não existe ainda, será criado.');
  }

  // Codifica conteúdo em Base64
  const encoded = Utilities.base64Encode(conteudo, Utilities.Charset.UTF_8);

  const payload = {
    message: 'Atualização automática CRM Club IA - ' + new Date().toLocaleDateString('pt-BR'),
    content: encoded
  };
  if (sha) payload.sha = sha;

  // Atualiza (ou cria) o arquivo
  const putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'PUT',
    headers: {
      'Authorization': 'token ' + GITHUB_TOKEN,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (putResp.getResponseCode() === 200 || putResp.getResponseCode() === 201) {
    Logger.log('data.json atualizado no GitHub com sucesso!');
  } else {
    throw new Error('Erro ao atualizar GitHub: ' + putResp.getResponseCode() + ' - ' + putResp.getContentText());
  }
}

// ─── CONFIGURAÇÃO INICIAL (executar uma vez) ────────────────
function configurarScript() {
  // Salva o token do GitHub como propriedade do script
  const token = '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';
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);

  Logger.log('Token configurado! Agora configure o gatilho semanal.');
  Logger.log('1. Clique no ícone de relógio (Gatilhos)');
  Logger.log('2. + Adicionar gatilho');
  Logger.log('3. Função: atualizarCRM');
  Logger.log('4. Evento: Baseado em tempo → Semanalmente → Segunda → 9h-10h');

  // Testa a conexão
  try {
    atualizarCRM();
    Logger.log('✅ Teste concluído! Tudo funcionando.');
  } catch(e) {
    Logger.log('❌ Erro no teste: ' + e.message);
  }
}
