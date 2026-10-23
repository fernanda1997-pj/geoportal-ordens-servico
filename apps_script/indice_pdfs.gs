/**
 * Índice dos PDFs do SISTEMA_AGETO (nome -> ID do arquivo no Google Drive).
 *
 * O Drive instalado no PC (G:\) não mostra o ID dos arquivos, então o
 * geoportal das O.S.P. só conseguia abrir uma BUSCA pelo nome do PDF. Este
 * script roda na conta Google da usuária, lista os PDFs das pastas
 * BOLETIM_PDF, INVENTARIOS_PDF e INSP_PDF_CONSOLIDADO de cada "LOTE XX" e
 * grava um JSON no Meu Drive. O converter_os.py lê esse JSON por G:\Meu Drive
 * e o site passa a abrir o PDF direto (drive.google.com/file/d/<ID>/view).
 *
 * Só LÊ a pasta do órgão — nada é escrito no SISTEMA_AGETO.
 *
 * Instalação (uma vez): script.google.com > Novo projeto > colar este arquivo
 * > rodar gerarIndicePdfs (autorizar) > rodar instalarGatilhoDiario.
 */

// Mesma pasta que aparece em G:\.shortcut-targets-by-id\<ID>\SISTEMA_AGETO
var ID_PASTA_SISTEMA = '16Cw6zdJvWIidBLYdaIQIh6ITuwcbe6d8';
var NOME_MANUTENCAO = 'MANUTENÇÃO RODOVIÁRIA';
var SUBPASTAS = ['BOLETIM_PDF', 'INVENTARIOS_PDF', 'INSP_PDF_CONSOLIDADO'];
var NOME_INDICE = 'indice_pdfs_sistema_ageto.json';
// Apps Script corta a execução em 6 min — para antes e avisa no log.
var LIMITE_MS = 5 * 60 * 1000;

function gerarIndicePdfs() {
  var inicio = Date.now();
  var raiz = pastaManutencao_();
  var arquivos = {};
  var total = 0;
  var lotes = raiz.getFolders();
  while (lotes.hasNext()) {
    var lote = lotes.next();
    var nomeLote = lote.getName();
    if (!/^LOTE /i.test(nomeLote)) continue;
    SUBPASTAS.forEach(function (sub) {
      var pastas = lote.getFoldersByName(sub);
      while (pastas.hasNext()) {
        var it = pastas.next().getFiles();
        while (it.hasNext()) {
          if (Date.now() - inicio > LIMITE_MS) {
            throw new Error('Passou de 5 min no ' + nomeLote + '/' + sub + ' — índice NÃO foi gravado.');
          }
          var f = it.next();
          if (!/\.pdf$/i.test(f.getName())) continue;
          arquivos[nomeLote + '/' + sub + '/' + f.getName()] = f.getId();
          total++;
        }
      }
    });
  }
  var conteudo = JSON.stringify({
    gerado_em: new Date().toISOString(),
    total: total,
    arquivos: arquivos
  });
  var existentes = DriveApp.getRootFolder().getFilesByName(NOME_INDICE);
  if (existentes.hasNext()) {
    existentes.next().setContent(conteudo);
  } else {
    DriveApp.getRootFolder().createFile(NOME_INDICE, conteudo, 'application/json');
  }
  Logger.log(total + ' PDF(s) indexados em ' + Math.round((Date.now() - inicio) / 1000) + ' s');
}

// A pasta do atalho pode ser o próprio SISTEMA_AGETO ou a pasta que o contém.
function pastaManutencao_() {
  var pasta = DriveApp.getFolderById(ID_PASTA_SISTEMA);
  if (pasta.getName() !== 'SISTEMA_AGETO') {
    var sis = pasta.getFoldersByName('SISTEMA_AGETO');
    if (!sis.hasNext()) throw new Error('SISTEMA_AGETO não encontrado em ' + pasta.getName());
    pasta = sis.next();
  }
  var man = pasta.getFoldersByName(NOME_MANUTENCAO);
  if (!man.hasNext()) throw new Error(NOME_MANUTENCAO + ' não encontrada em SISTEMA_AGETO');
  return man.next();
}

// Roda todo dia entre 8h e 9h — antes da rotina das 10:00 que roda o converter.
function instalarGatilhoDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'gerarIndicePdfs') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('gerarIndicePdfs').timeBased().everyDays(1).atHour(8).create();
}
