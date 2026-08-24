/*************************************************************
 * PAUTA — backend Google Apps Script
 * Guarda usuários, colunas, ações, posts, pedidos de senha
 * na planilha e os anexos (PNG, JPG, PDF, SVG) no Drive.
 *
 * COMO INSTALAR
 * 1. Crie uma planilha no Google Sheets e copie o ID da URL
 *    (.../spreadsheets/d/AQUI_O_ID/edit).
 * 2. Extensões > Apps Script, apague o conteúdo e cole este arquivo.
 * 3. Preencha PLANILHA_ID e PASTA_ID abaixo (a pasta do Drive onde
 *    os anexos serão salvos; deixe vazio para criar uma automática).
 * 4. Execute a função instalar() uma vez e autorize os acessos.
 * 5. Implantar > Nova implantação > Aplicativo da Web
 *    - Executar como: eu
 *    - Quem pode acessar: qualquer pessoa
 * 6. Copie a URL /exec e cole em CONFIG.API_URL no index.html.
 *
 * OBSERVAÇÕES
 * - O app aceita anexos de até 60 MB. O Apps Script tem limite
 *   próprio de payload (~50 MB): arquivos muito grandes podem
 *   falhar no envio e ficam disponíveis apenas na sessão do
 *   navegador. Para uso pesado, suba o arquivo direto no Drive
 *   e cole o link no card.
 * - As senhas ficam em texto na planilha. Mantenha a planilha
 *   restrita a quem administra o sistema.
 *************************************************************/

var PLANILHA_ID = '';                 // ID da planilha (deixe vazio para o script criar uma sozinho)
var PASTA_ID    = '';                 // ID da pasta do Drive para anexos (opcional)
var PASTA_NOME  = 'PAUTA - anexos';   // usada quando PASTA_ID está vazio

var ABAS = {
  USUARIOS: ['id','nome','usuario','senha','papel'],
  COLUNAS : ['id','titulo','cor'],
  CARTOES : ['id','colId','titulo','desc','prioridade','inicio','prazo','hora','resp','tipo','criado','anexos'],
  POSTS   : ['id','titulo','texto','canal','autor','data','hora','status','parecer','avaliador','anexos'],
  PEDIDOS : ['id','usuario','quando','status'],
  LOG     : ['quando','acao','detalhe']
};

/* ---------------- infraestrutura ---------------- */

/* Usa PLANILHA_ID quando preenchido. Se estiver vazio, cria uma
   planilha "PAUTA - base" na primeira execução e guarda o ID nas
   propriedades do script (não precisa editar o código). */
function planilha_() {
  if (PLANILHA_ID) return SpreadsheetApp.openById(PLANILHA_ID);
  var props = PropertiesService.getScriptProperties();
  var salvo = props.getProperty('PLANILHA_ID');
  if (salvo) {
    try { return SpreadsheetApp.openById(salvo); } catch (e) { props.deleteProperty('PLANILHA_ID'); }
  }
  var nova = SpreadsheetApp.create('PAUTA - base');
  props.setProperty('PLANILHA_ID', nova.getId());
  return nova;
}

/* Mostra onde os dados estão gravados */
function ondeEstaAPlanilha() {
  var ss = planilha_();
  Logger.log('Planilha: ' + ss.getName() + '\nID: ' + ss.getId() + '\nURL: ' + ss.getUrl());
  return ss.getUrl();
}

function aba_(nome) {
  var ss = planilha_();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(ABAS[nome]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function pasta_() {
  if (PASTA_ID) return DriveApp.getFolderById(PASTA_ID);
  var it = DriveApp.getFoldersByName(PASTA_NOME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_NOME);
}

function instalar() {
  Object.keys(ABAS).forEach(function (nome) { aba_(nome); });
  var sh = aba_('USUARIOS');
  if (sh.getLastRow() < 2) {
    sh.appendRow([Utilities.getUuid().slice(0, 8), 'Administrador', 'admin', 'pauta2026', 'admin']);
  }
  var f = pasta_();
  var ss = planilha_();
  var msg = 'Instalado.\nPlanilha: ' + ss.getUrl() +
            '\nID da planilha: ' + ss.getId() +
            '\nPasta de anexos: ' + f.getName() + ' (' + f.getId() + ')';
  Logger.log(msg);
  return msg;
}

/* ---------------- leitura e gravação ---------------- */

function lerAba_(nome) {
  var sh = aba_(nome);
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var cab = vals[0];
  return vals.slice(1).filter(function (l) { return String(l[0]).trim() !== ''; })
    .map(function (l) {
      var o = {};
      cab.forEach(function (c, i) {
        var v = l[i];
        if (c === 'anexos') { try { o[c] = v ? JSON.parse(v) : []; } catch (e) { o[c] = []; } }
        else if (Object.prototype.toString.call(v) === '[object Date]') {
          o[c] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else { o[c] = v === null ? '' : String(v); }
      });
      return o;
    });
}

function gravarAba_(nome, lista) {
  var sh = aba_(nome);
  var cab = ABAS[nome];
  sh.clear();
  var linhas = [cab];
  (lista || []).forEach(function (o) {
    linhas.push(cab.map(function (c) {
      var v = o[c];
      if (c === 'anexos') return JSON.stringify(v || []);
      return v === undefined || v === null ? '' : v;
    }));
  });
  sh.getRange(1, 1, linhas.length, cab.length).setValues(linhas);
  sh.setFrozenRows(1);
}

function log_(acao, detalhe) {
  try {
    aba_('LOG').appendRow([new Date(), acao, detalhe || '']);
  } catch (e) {}
}

/* ---------------- ações ---------------- */

function carregarTudo_() {
  return {
    usuarios: lerAba_('USUARIOS'),
    colunas : lerAba_('COLUNAS'),
    cartoes : lerAba_('CARTOES'),
    posts   : lerAba_('POSTS'),
    pedidos : lerAba_('PEDIDOS')
  };
}

function salvarTudo_(d) {
  var trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    if (d.usuarios) gravarAba_('USUARIOS', d.usuarios);
    if (d.colunas)  gravarAba_('COLUNAS',  d.colunas);
    if (d.cartoes)  gravarAba_('CARTOES',  d.cartoes);
    if (d.posts)    gravarAba_('POSTS',    d.posts);
    if (d.pedidos)  gravarAba_('PEDIDOS',  d.pedidos);
    log_('salvarTudo', (d.cartoes || []).length + ' ações / ' + (d.posts || []).length + ' posts');
    return { salvo: true, quando: new Date().toISOString() };
  } finally {
    trava.releaseLock();
  }
}

/* Recebe {nome, dados} onde dados é um dataURL base64 e devolve {id, url} */
function anexar_(d) {
  var nome = d.nome || ('arquivo-' + Date.now());
  var partes = String(d.dados || '').split(',');
  var mime = (partes[0].match(/:(.*?);/) || [])[1] || 'application/octet-stream';
  var permitidos = ['image/png', 'image/jpeg', 'application/pdf', 'image/svg+xml'];
  var ext = nome.toLowerCase().split('.').pop();
  if (permitidos.indexOf(mime) < 0 && ['png','jpg','jpeg','pdf','svg'].indexOf(ext) < 0) {
    throw new Error('Formato não aceito: ' + nome);
  }
  var bytes = Utilities.base64Decode(partes[1] || '');
  if (bytes.length > 60 * 1024 * 1024) throw new Error('Arquivo acima de 60 MB: ' + nome);
  var blob = Utilities.newBlob(bytes, mime, nome);
  var arq = pasta_().createFile(blob);
  arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  log_('anexar', nome + ' (' + Math.round(bytes.length / 1024) + ' KB)');
  return {
    id: arq.getId(),
    nome: nome,
    tamanho: bytes.length,
    url: 'https://drive.google.com/uc?export=view&id=' + arq.getId(),
    link: arq.getUrl()
  };
}

function login_(d) {
  var u = String(d.usuario || '').trim().toLowerCase();
  var s = String(d.senha || '');
  var achou = lerAba_('USUARIOS').filter(function (x) {
    return String(x.usuario).toLowerCase() === u && String(x.senha) === s;
  })[0];
  if (!achou) throw new Error('Usuário ou senha não confere.');
  log_('login', u);
  return { id: achou.id, nome: achou.nome, usuario: achou.usuario, papel: achou.papel };
}

function pedirSenha_(d) {
  var sh = aba_('PEDIDOS');
  sh.appendRow([Utilities.getUuid().slice(0, 8), String(d.usuario || '').toLowerCase(),
                Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), 'aberto']);
  log_('pedirSenha', d.usuario);
  return { enviado: true };
}

function definirSenha_(d) {
  var sh = aba_('USUARIOS');
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(d.id) ||
        String(vals[i][2]).toLowerCase() === String(d.usuario || '').toLowerCase()) {
      sh.getRange(i + 1, 4).setValue(String(d.senha));
      log_('definirSenha', vals[i][2]);
      return { alterado: true };
    }
  }
  throw new Error('Usuário não encontrado.');
}

/* ---------------- entrada web ---------------- */

function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return responder_(req.acao, req.dados || {});
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  return responder_(p.acao || 'carregarTudo', p);
}

function responder_(acao, dados) {
  var saida;
  try {
    switch (acao) {
      case 'carregarTudo': saida = { ok: true, dados: carregarTudo_() }; break;
      case 'salvarTudo':   saida = { ok: true, dados: salvarTudo_(dados) }; break;
      case 'anexar':       saida = { ok: true, dados: anexar_(dados) }; break;
      case 'login':        saida = { ok: true, dados: login_(dados) }; break;
      case 'pedirSenha':   saida = { ok: true, dados: pedirSenha_(dados) }; break;
      case 'definirSenha': saida = { ok: true, dados: definirSenha_(dados) }; break;
      case 'ping':         saida = { ok: true, dados: { versao: 'pauta-1.0', quando: new Date().toISOString() } }; break;
      default:             saida = { ok: false, erro: 'Ação desconhecida: ' + acao };
    }
  } catch (err) {
    saida = { ok: false, erro: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida))
    .setMimeType(ContentService.MimeType.JSON);
}
