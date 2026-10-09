/* Impulso Empatia 360 | Serviço de relatório e entrega por e-mail.
 * Projeto independente. Não altera scripts das outras avaliações.
 * Histórico e perguntas são preservados no frontend original.
 */
const SERVICE = Object.freeze({
  researchId: 'empatia-360',
  label: 'Pesquisa Impulso Empatia 360',
  spreadsheetId: '1duwnwums_1eb-spvzvfaoujQc6DoZAbdWHzKem8ebS4',
  spreadsheetTitle: 'Impulso - Resultados Empatia 360',
  sheetName: 'Avaliações',
  driveFolderName: 'Impulso_Empatia_360_Relatorios',
  adminEmail: 'impulsoflow@gmail.com',
  maxPdfBase64: 8500000,
  maxDailyPerEmail: 5,
  columns: ["Registrado em", "ID da aplicação", "Nome completo", "E-mail", "WhatsApp", "Idade", "Profissão", "Escolaridade", "TEQ (0–64)", "Tomada de perspectiva (0–16)", "Reconhecimento emocional (0–16)", "Cuidado compassivo (0–16)", "Escuta e validação (0–16)", "Resposta empática e limites (0–16)", "Arquivo PDF", "Link para PDF", "E-mail enviado?", "Falha/observações", "Versão da avaliação", "Consentimento", "Respondido em", "TEQ 01", "TEQ 02", "TEQ 03", "TEQ 04", "TEQ 05", "TEQ 06", "TEQ 07", "TEQ 08", "TEQ 09", "TEQ 10", "TEQ 11", "TEQ 12", "TEQ 13", "TEQ 14", "TEQ 15", "TEQ 16", "Complementar 01", "Complementar 02", "Complementar 03", "Complementar 04", "Complementar 05", "Complementar 06", "Complementar 07", "Complementar 08", "Complementar 09", "Complementar 10", "Complementar 11", "Complementar 12", "Complementar 13", "Complementar 14", "Complementar 15", "Complementar 16", "Complementar 17", "Complementar 18", "Complementar 19", "Complementar 20"]
});
function fmt(v) { return String(v == null ? '' : v).trim(); }
function safeCell_(v) {const s=fmt(v);return /^[=+@-]/.test(s)?"'"+s:s;}
function output(obj, callback) {
  const raw = JSON.stringify(obj);
  if(callback) {
    const valid = /^__empatiaStatus_[A-Za-z0-9_]{6,80}$/.test(callback);
    if(!valid) return ContentService.createTextOutput('Invalid callback').setMimeType(ContentService.MimeType.TEXT);
    return ContentService.createTextOutput(callback+'('+raw+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(raw).setMimeType(ContentService.MimeType.JSON);
}
function ensureStorage_() {
  // Nunca cria outra planilha nem altera cabeçalhos existentes.
  const sheet = SpreadsheetApp.openById(SERVICE.spreadsheetId).getSheetByName(SERVICE.sheetName);
  if(!sheet)throw Error('A aba Avaliações não está acessível.');
  const headers = sheet.getRange(1,1,1,SERVICE.columns.length).getValues()[0];
  if(headers.join('|')!==SERVICE.columns.join('|'))
    throw Error('Estrutura da planilha diferente do padrão Empatia 360. Gravação bloqueada.');
  const props=PropertiesService.getScriptProperties();
  const key='FOLDER_ID';
  const folderId=props.getProperty(key);
  let folder=null;
  if(folderId)try{folder=DriveApp.getFolderById(folderId);folder.getName();}catch(e){folder=null;}
  if(!folder) {
    folder=DriveApp.createFolder(SERVICE.driveFolderName);
    props.setProperty(key,folder.getId());
  }
  return {sheet,folder};
}
function findSubmission_(sheet,id) {
  if(sheet.getLastRow()<2)return 0;
  const found=sheet.getRange(2,2,sheet.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();
  return found?found.getRow():0;
}
function validate_(p) {
  if(!p||p.researchId!==SERVICE.researchId||p.action!=='saveEmpatia360')throw Error('Tipo de pesquisa não reconhecido.');
  if(!/^[a-zA-Z0-9_-]{14,90}$/.test(fmt(p.submissionId)))throw Error('Identificador de envio inválido.');
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(fmt(p.email))||fmt(p.nome).length<3)throw Error('Nome ou e-mail inválido.');
  if(!Array.isArray(p.teqRespostas)||p.teqRespostas.length!==16||
     !Array.isArray(p.respostasComplementares)||p.respostasComplementares.length!==20)
    throw Error('As 36 respostas são obrigatórias.');
  const items=p.teqRespostas.concat(p.respostasComplementares);
  if(items.some(x=>!Number.isInteger(x)||x<0||x>4))throw Error('Escala de respostas inválida.');
  const reverse=[1,3,6,9,10,11,13,14];
  const correct=items.slice(0,16).reduce((sum,v,i)=>sum+(reverse.indexOf(i)>=0?4-v:v),0);
  if(correct!==Number(p.teqScore))throw Error('Pontuação TEQ não confere.');
  if(!Array.isArray(p.dimensoes)||p.dimensoes.length!==5 ||
    p.dimensoes.some(x=>!Number.isInteger(x.pontos)||x.pontos<0||x.pontos>16))throw Error('Dimensões inválidas.');
  const b64=fmt(p.pdfBase64);
  if(!b64||b64.length>SERVICE.maxPdfBase64||!/^[A-Za-z0-9+/]+={0,2}$/.test(b64))throw Error('PDF ausente, inválido ou acima do tamanho permitido.');
  if(p.sendEmail!==true||p.consentimento!==true)throw Error('Confirmação de envio e consentimento obrigatória.');
}
function parse_(e) {
  if(e&&e.parameter&&e.parameter.payload) return JSON.parse(e.parameter.payload);
  return JSON.parse(e.postData.contents);
}
function safeFile_(name) {
  return fmt(name).replace(/[^A-Za-z0-9_.-]+/g,'_').slice(0,110)||'Impulso_Empatia_360.pdf';
}
function status_(id) {
  const spreadsheetId=SERVICE.spreadsheetId;
  const sheet=SpreadsheetApp.openById(spreadsheetId).getSheetByName(SERVICE.sheetName);
  const row=findSubmission_(sheet,id);
  if(!row)return {ok:false,pending:true,submissionId:id};
  const cells=sheet.getRange(row,16,1,3).getValues()[0];
  const sent=String(cells[1]).toUpperCase()==='SIM';
  const err=fmt(cells[2]);
  return {ok:sent,saved:true,emailSent:sent,pending:!sent&&!err,submissionId:id,error:err||undefined};
}
function doGet(e) {
  const p=(e&&e.parameter)||{};
  try {
    if(p.action!=='status'||!/^[a-zA-Z0-9_-]{14,90}$/.test(fmt(p.submissionId)))
      return output({ok:true,service:'Impulso Empatia 360',version:'2026-10-09'});
    return output(status_(fmt(p.submissionId)),fmt(p.callback));
  } catch(error){return output({ok:false,pending:true,submissionId:fmt(p.submissionId),message:fmt(error)},fmt(p.callback));}
}
function doPost(e) {
  const lock=LockService.getScriptLock();
  let p=null, row=0;
  try {
    lock.waitLock(30000);
    p=parse_(e);
    validate_(p);
    const db=ensureStorage_(),sheet=db.sheet;
    const id=fmt(p.submissionId);
    row=findSubmission_(sheet,id);
    if(row&&String(sheet.getRange(row,17).getValue()).toUpperCase()==='SIM')
       return output({ok:true,saved:true,emailSent:true,submissionId:id,duplicate:true});
    if(!row) {
      const email=fmt(p.email).toLowerCase();
      const previous=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues().filter(x=>{const dt=x[0] instanceof Date?x[0]:new Date(x[0]);return String(x[3]).toLowerCase()===email && !isNaN(dt.getTime()) && Date.now()-dt.getTime()<86400000;}).length:0;
      if(previous>=SERVICE.maxDailyPerEmail)throw Error('Limite de aplicações para este e-mail; solicite suporte ao Instituto Impulso.');
      sheet.appendRow([
  new Date(),id,safeCell_(p.nome),email,fmt(p.whatsapp),fmt(p.idade),
  safeCell_(p.profissao),safeCell_(p.escolaridade),Number(p.teqScore),
  ...p.dimensoes.map(x=>Number(x.pontos)),
  safeFile_(p.pdfFileName),'','PENDENTE','','teq16_autoral20_v1','SIM',
  fmt(p.respondidoEm),...p.teqRespostas,...p.respostasComplementares
]);
      row=sheet.getLastRow();
      SpreadsheetApp.flush();
    }
    let pdfLink=fmt(sheet.getRange(row,16).getValue());
    let pdfFile=null;
    const base64=fmt(p.pdfBase64);
    const blob=Utilities.newBlob(Utilities.base64Decode(base64),'application/pdf',safeFile_(p.pdfFileName));
    if(!pdfLink) {
      pdfFile=db.folder.createFile(blob);
      pdfFile.setDescription('Empatia 360 | '+id+' | '+fmt(p.nome));
      pdfLink=pdfFile.getUrl();
      sheet.getRange(row,16).setValue(pdfLink);
      SpreadsheetApp.flush();
    }
    try{
      MailApp.sendEmail(fmt(p.email), 'Seu relatório | Impulso Empatia 360',
        'Olá, '+fmt(p.nome)+'! Seu relatório Impulso Empatia 360 está em anexo. Equipe Instituto Impulso.',
        {name:'Instituto Impulso IE de Liderança',attachments:[blob],
         htmlBody:'<p>Olá, <b>'+escape_(p.nome)+'</b>!</p><p>Sua avaliação <b>Impulso Empatia 360</b> foi concluída. O relatório completo em PDF segue anexado a esta mensagem.</p><p>O relatório tem finalidade educativa e foi desenvolvido para apoiar seu autoconhecimento e desenvolvimento.</p><p>Equipe Instituto Impulso IE™ de Liderança</p>'});
      sheet.getRange(row,17,1,2).setValues([['SIM','']]);
      SpreadsheetApp.flush();
      if(p.notifyAdmin===true&&fmt(p.email).toLowerCase()!==SERVICE.adminEmail) {
        try { MailApp.sendEmail(SERVICE.adminEmail,'Nova pesquisa Empatia 360 concluída',
        'Nova pesquisa concluída. Nome: '+fmt(p.nome)+'\nE-mail: '+fmt(p.email)+'\nPDF: '+pdfLink); }catch(adminError){}
      }
      return output({ok:true,saved:true,emailSent:true,submissionId:id,reportLink:pdfLink});
    }catch(emailError){
      sheet.getRange(row,17,1,2).setValues([['NAO',fmt(emailError)]]);
      SpreadsheetApp.flush();
      return output({ok:false,saved:true,emailSent:false,submissionId:id,error:'Falha no envio do e-mail; o PDF foi salvo.'});
    }
  }catch(error){
    if(row&&p){ try {const sheet=SpreadsheetApp.openById(SERVICE.spreadsheetId).getSheetByName(SERVICE.sheetName);sheet.getRange(row,18).setValue(fmt(error));} catch(ignore) {}}
    return output({ok:false,saved:!!row,emailSent:false,submissionId:p&&p.submissionId||'',error:fmt(error)});
  }finally{try{lock.releaseLock();}catch(ignore){}}
}
function escape_(value) {return fmt(value).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
function authorizeServices_(){
  const db=ensureStorage_();
  MailApp.getRemainingDailyQuota();
  return {sheet:db.sheet.getParent().getUrl(),folder:db.folder.getUrl(),gmail:true};
}function authorizeServices(){return authorizeServices_();}

// Configuração de autorização validada para a planilha Empatia 360.
