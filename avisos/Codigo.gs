/**
 * Instituto Diney Moura — Lista de avisos ("Quero ser avisado")
 *
 * Este script fica dentro de uma Planilha Google da conta
 * institutodineymoura@gmail.com e faz três coisas:
 *   1. Recebe os cadastros feitos no site e guarda na aba "Interessados".
 *   2. Entrega ao site as datas publicadas na aba "Agenda".
 *   3. Quando uma data nova é publicada, manda e-mail para quem estava esperando.
 *
 * Passo a passo de instalação: veja o arquivo COMO-CONFIGURAR.md.
 */

// ===== Ajustes =====
const SERVICOS = ['Exame de vista'];          // serviços com lista de avisos
const NOME_REMETENTE = 'Instituto Diney Moura';
const WHATSAPP_INSTITUTO = '553193443778';    // (31) 9344-3778
const SITE = 'https://institutodineymoura.github.io';
const ENDERECO_PADRAO = 'R. José Ovídio Guerra, 260 – Cardoso, Belo Horizonte – MG';
const FUSO = 'America/Sao_Paulo';

// ===== Abas e colunas =====
const ABA_INTERESSADOS = 'Interessados';
const COL_I = ['Cadastrado em', 'Nome', 'E-mail', 'WhatsApp', 'Interesse', 'Situação', 'Avisado em', 'Avisar no WhatsApp'];
const ABA_AGENDA = 'Agenda';
const COL_A = ['Serviço', 'Data', 'Horário', 'Local', 'Vagas', 'Observações', 'Publicar e avisar', 'Avisos enviados em', 'Pessoas avisadas'];

const AGUARDANDO = 'Aguardando';
const AVISADO = 'Avisado';
const CANCELADO = 'Cancelado';

// =====================================================================
// Menu da planilha
// =====================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Instituto')
    .addItem('Enviar avisos agora', 'enviarAvisosMenu')
    .addSeparator()
    .addItem('Configurar planilha (1ª vez)', 'configurar')
    .addToUi();
}

/** Cria as abas, formata e liga o envio automático a cada hora. Rode uma vez. */
function configurar() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(FUSO);

  const shI = ss.getSheetByName(ABA_INTERESSADOS) || ss.insertSheet(ABA_INTERESSADOS);
  shI.getRange(1, 1, 1, COL_I.length).setValues([COL_I]).setFontWeight('bold').setBackground('#f6f3ec');
  shI.setFrozenRows(1);
  shI.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm');
  shI.getRange('G2:G').setNumberFormat('dd/MM/yyyy HH:mm');
  shI.getRange('D2:D').setNumberFormat('@');
  shI.getRange('F2:F').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([AGUARDANDO, AVISADO, CANCELADO], true).build());
  shI.setColumnWidths(1, COL_I.length, 150);
  shI.setColumnWidth(3, 230);

  const shA = ss.getSheetByName(ABA_AGENDA) || ss.insertSheet(ABA_AGENDA);
  shA.getRange(1, 1, 1, COL_A.length).setValues([COL_A]).setFontWeight('bold').setBackground('#f6f3ec');
  shA.setFrozenRows(1);
  shA.getRange('A2:A200').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(SERVICOS, true).build());
  shA.getRange('B2:B200').setNumberFormat('dd/MM/yyyy').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('Digite a data, por exemplo 13/12/2026').build());
  shA.getRange('C2:C200').setNumberFormat('@');
  shA.getRange('G2:G200').insertCheckboxes();
  shA.getRange('H2:H200').setNumberFormat('dd/MM/yyyy HH:mm');
  shA.setColumnWidths(1, COL_A.length, 140);
  shA.setColumnWidth(4, 260);
  shA.setColumnWidth(6, 260);

  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 2 && padrao.getLastRow() === 0) ss.deleteSheet(padrao);

  // Envio automático a cada hora (sem duplicar o gatilho)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'enviarAvisos')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarAvisos').timeBased().everyHours(1).create();

  avisar_('Pronto! Abas criadas e envio automático ligado (a cada hora).');
}

// =====================================================================
// Site → planilha
// =====================================================================

/** GET: o site pede as datas publicadas. */
function doGet(e) {
  const acao = (e && e.parameter && e.parameter.acao) || 'agenda';
  if (acao === 'agenda') return json_({ ok: true, agenda: agendaPublica_() });
  return json_({ ok: false, erro: 'acao' });
}

/** POST: cadastro de "Quero ser avisado". */
function doPost(e) {
  let d;
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (_) { return json_({ ok: false, erro: 'formato' }); }

  // Robôs preenchem o campo escondido "site"; pessoas não.
  if (d.site) return json_({ ok: true });

  const nome = limpa_(d.nome, 80);
  const email = limpa_(d.email, 120).toLowerCase();
  const whats = soDigitos_(d.whatsapp).slice(0, 13);
  const interesse = limpa_(d.interesse, 60);

  if (nome.length < 2) return json_({ ok: false, erro: 'nome' });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json_({ ok: false, erro: 'email' });
  if (whats && (whats.length < 10)) return json_({ ok: false, erro: 'whatsapp' });
  if (SERVICOS.indexOf(interesse) < 0) return json_({ ok: false, erro: 'interesse' });
  if (d.consentimento !== true) return json_({ ok: false, erro: 'consentimento' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let novo = false;
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(ABA_INTERESSADOS);
    const n = sh.getLastRow() - 1;
    const linhas = n > 0 ? sh.getRange(2, 1, n, COL_I.length).getValues() : [];
    const i = linhas.findIndex(r => String(r[2]).toLowerCase() === email && r[4] === interesse && r[5] === AGUARDANDO);
    if (i >= 0) {
      // Já estava na lista: só atualiza nome e WhatsApp.
      sh.getRange(i + 2, 2).setValue(nome);
      if (whats) sh.getRange(i + 2, 4).setValue(whats);
    } else {
      sh.appendRow([new Date(), nome, email, whats, interesse, AGUARDANDO, '', '']);
      novo = true;
    }
  } finally {
    lock.releaseLock();
  }

  if (novo) {
    try { emailConfirmacao_(nome, email, interesse); } catch (err) { console.warn(err); }
  }
  return json_({ ok: true, jaCadastrado: !novo });
}

// =====================================================================
// Avisos
// =====================================================================
function enviarAvisosMenu() {
  const r = enviarAvisos();
  avisar_(r.mensagem);
}

/** Procura datas publicadas que ainda não foram avisadas e manda os e-mails. */
function enviarAvisos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { mensagem: 'Já existe um envio em andamento. Tente de novo em 1 minuto.' };
  try {
    const ss = SpreadsheetApp.getActive();
    const shA = ss.getSheetByName(ABA_AGENDA);
    const shI = ss.getSheetByName(ABA_INTERESSADOS);
    const hoje = inicioDoDia_(new Date());

    const nA = shA.getLastRow() - 1;
    const agenda = nA > 0 ? shA.getRange(2, 1, nA, COL_A.length).getValues() : [];
    const nI = shI.getLastRow() - 1;
    const pessoas = nI > 0 ? shI.getRange(2, 1, nI, COL_I.length).getValues() : [];

    let totalEnviados = 0, datas = 0, faltou = false;

    for (let a = 0; a < agenda.length; a++) {
      const [servico, data, horario, local, vagas, obs, publicar, enviadoEm, jaAvisados] = agenda[a];
      if (publicar !== true || enviadoEm || !servico || !(data instanceof Date)) continue;
      if (inicioDoDia_(data) < hoje) continue;

      const ev = { servico, data, horario: String(horario || ''), local: String(local || ''), vagas, obs: String(obs || '') };
      let enviadosAqui = Number(jaAvisados) || 0;
      datas++;

      for (let p = 0; p < pessoas.length; p++) {
        const r = pessoas[p];
        if (r[4] !== servico || r[5] !== AGUARDANDO) continue;
        if (MailApp.getRemainingDailyQuota() < 1) { faltou = true; break; }
        try {
          emailAviso_(r[1], r[2], ev);
        } catch (err) {
          console.warn('Falha ao enviar para ' + r[2] + ': ' + err);
          continue;
        }
        const linha = p + 2;
        const agora = new Date();
        shI.getRange(linha, 6, 1, 3).setValues([[AVISADO, agora, linkWhatsApp_(r[1], r[3], ev)]]);
        r[5] = AVISADO;
        enviadosAqui++;
        totalEnviados++;
      }

      shA.getRange(a + 2, 9).setValue(enviadosAqui);
      if (!faltou) shA.getRange(a + 2, 8).setValue(new Date());
      SpreadsheetApp.flush();
      if (faltou) break;
    }

    let msg;
    if (!datas) msg = 'Nenhuma data nova para avisar. (Marque "Publicar e avisar" na aba Agenda.)';
    else msg = totalEnviados + ' e-mail(s) enviado(s).';
    if (faltou) msg += ' O limite diário de e-mails acabou; o restante será enviado automaticamente amanhã.';
    return { mensagem: msg, enviados: totalEnviados };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// E-mails
// =====================================================================
function emailConfirmacao_(nome, email, interesse) {
  const primeiro = esc_(primeiroNome_(nome));
  const html = moldura_(
    '<p>Olá, ' + primeiro + '!</p>' +
    '<p>Seu nome está na lista de avisos de <strong>' + esc_(interesse.toLowerCase()) + '</strong> do Instituto Diney Moura.</p>' +
    '<p>Assim que tivermos uma data marcada, mandamos um e-mail para você com o dia, o horário e como garantir sua vaga.</p>');
  MailApp.sendEmail({
    to: email, name: NOME_REMETENTE,
    subject: 'Você está na lista de avisos — ' + interesse,
    htmlBody: html,
    body: 'Olá, ' + primeiroNome_(nome) + '! Seu nome está na lista de avisos de ' + interesse.toLowerCase() +
      ' do Instituto Diney Moura. Assim que tivermos uma data marcada, avisamos por e-mail.'
  });
}

function emailAviso_(nome, email, ev) {
  const quando = dataPorExtenso_(ev.data) + (ev.horario ? ', às ' + ev.horario : '');
  const local = ev.local || ENDERECO_PADRAO;
  const zap = 'https://wa.me/' + WHATSAPP_INSTITUTO + '?text=' +
    encodeURIComponent('Olá! Recebi o aviso e quero garantir minha vaga no ' + ev.servico.toLowerCase() + ' do dia ' + dataCurta_(ev.data) + '.');
  const html = moldura_(
    '<p>Olá, ' + esc_(primeiroNome_(nome)) + '!</p>' +
    '<p>Saiu a data que você estava esperando:</p>' +
    '<table style="border-collapse:collapse;margin:8px 0 18px;font-size:16px">' +
    linhaTabela_('Serviço', ev.servico) +
    linhaTabela_('Quando', quando) +
    linhaTabela_('Onde', local) +
    (ev.vagas ? linhaTabela_('Vagas', String(ev.vagas)) : '') +
    (ev.obs ? linhaTabela_('Importante', ev.obs) : '') +
    '</table>' +
    '<p><a href="' + zap + '" style="display:inline-block;background:#3fb39a;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:999px">Garantir minha vaga pelo WhatsApp</a></p>' +
    '<p style="color:#4d4d4d;font-size:14px">As vagas são limitadas. Se não puder ir, não precisa fazer nada.</p>');
  MailApp.sendEmail({
    to: email, name: NOME_REMETENTE,
    subject: ev.servico + ': nova data em ' + dataCurta_(ev.data),
    htmlBody: html,
    body: 'Olá, ' + primeiroNome_(nome) + '! Saiu a data de ' + ev.servico.toLowerCase() + ': ' + quando +
      '. Local: ' + local + '. Garanta sua vaga pelo WhatsApp: ' + zap
  });
}

function moldura_(conteudo) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#141414;line-height:1.5;max-width:560px">' +
    '<div style="height:6px;background:linear-gradient(90deg,#e3342f 25%,#f2c500 25% 50%,#3fb39a 50% 75%,#2f78c8 75%)"></div>' +
    '<div style="padding:20px 4px">' + conteudo +
    '<p style="margin-top:28px">Um abraço,<br><strong>Instituto Diney Moura</strong><br>' +
    'WhatsApp (31) 9344-3778 · <a href="' + SITE + '" style="color:#2f78c8">' + SITE.replace('https://', '') + '</a></p>' +
    '<p style="color:#888;font-size:12px;margin-top:24px">Você recebeu este e-mail porque pediu para ser avisado pelo site do Instituto. ' +
    'Para sair da lista, é só responder este e-mail com a palavra SAIR.</p>' +
    '</div></div>';
}

function linhaTabela_(rotulo, valor) {
  return '<tr><td style="padding:4px 16px 4px 0;color:#4d4d4d;vertical-align:top">' + esc_(rotulo) +
    '</td><td style="padding:4px 0;font-weight:bold">' + esc_(valor) + '</td></tr>';
}

/** Link pronto para o Instituto mandar o aviso pelo WhatsApp com um clique. */
function linkWhatsApp_(nome, whats, ev) {
  const d = soDigitos_(whats);
  if (d.length < 10) return '';
  const numero = d.length <= 11 ? '55' + d : d;
  const texto = 'Olá, ' + primeiroNome_(nome) + '! Aqui é do Instituto Diney Moura. Saiu a data de ' +
    ev.servico.toLowerCase() + ': ' + dataPorExtenso_(ev.data) + (ev.horario ? ', às ' + ev.horario : '') +
    '. Quer garantir sua vaga?';
  const url = 'https://wa.me/' + numero + '?text=' + encodeURIComponent(texto);
  return '=HYPERLINK("' + url.replace(/"/g, '""') + '","Enviar no WhatsApp")';
}

// =====================================================================
// Agenda pública (o que o site mostra)
// =====================================================================
function agendaPublica_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ABA_AGENDA);
  if (!sh) return [];
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const hoje = inicioDoDia_(new Date());
  return sh.getRange(2, 1, n, COL_A.length).getValues()
    .filter(r => r[0] && r[1] instanceof Date && r[6] === true && inicioDoDia_(r[1]) >= hoje)
    .sort((x, y) => x[1] - y[1])
    .map(r => ({
      servico: String(r[0]),
      data: Utilities.formatDate(r[1], FUSO, 'yyyy-MM-dd'),
      dataTexto: dataPorExtenso_(r[1]),
      horario: String(r[2] || ''),
      local: String(r[3] || ''),
      vagas: r[4] === '' ? '' : String(r[4]),
      obs: String(r[5] || '')
    }));
}

// =====================================================================
// Utilidades
// =====================================================================
const DIAS = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
const MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

function dataPorExtenso_(d) {
  const dia = Number(Utilities.formatDate(d, FUSO, 'd'));
  const mes = Number(Utilities.formatDate(d, FUSO, 'M')) - 1;
  const sem = Number(Utilities.formatDate(d, FUSO, 'u')) % 7; // 1=segunda … 7=domingo
  return DIAS[sem] + ', ' + dia + ' de ' + MESES[mes];
}
function dataCurta_(d) { return Utilities.formatDate(d, FUSO, 'dd/MM'); }
function inicioDoDia_(d) { return Number(Utilities.formatDate(d, FUSO, 'yyyyMMdd')); }
function primeiroNome_(n) { return String(n || '').trim().split(/\s+/)[0] || ''; }
function soDigitos_(s) { return String(s || '').replace(/\D/g, ''); }
function limpa_(s, max) {
  // Tira caracteres de controle e impede que um texto vire fórmula na planilha.
  return String(s || '').replace(/[\u0000-\u001f]/g, ' ').trim().replace(/^[=+\-@]+/, '').trim().slice(0, max);
}
function esc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function avisar_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) { console.log(msg); }
}
