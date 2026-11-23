/** Parciais de Venda – Servidor (Apps Script) · Postos Pichilau
 *  Conta apoio.escadabr · App: https://apoio-escada.github.io/parciais-venda/
 *  1ª vez: rodar configurar() no editor e autorizar. Depois Implantar › App da Web (Executar como: Eu · Acesso: Qualquer pessoa).
 */
const VERSAO='1.0.0';
const ABAS={Config:['Posto','JSON','PinHash','AtualizadoEm'],Parciais:['Id','Posto','Data','Hora','Tipo','Lider','JSON','Em'],Vendedores:['Id','Posto','Data','Hora','Tipo','Lider','JSON','Em'],Marcas:['Id','Posto','Data','JSON','Em'],Fotos:['Id','Img','Em'],Log:['Quando','Acao','Posto','Detalhe']};
const PR=()=>PropertiesService.getScriptProperties();
function ss_(){ const id=PR().getProperty('PLANILHA_ID'); if(id) return SpreadsheetApp.openById(id);
  const ss=SpreadsheetApp.create('Parciais de Venda – Controle'); PR().setProperty('PLANILHA_ID',ss.getId());
  Object.keys(ABAS).forEach((n,i)=>{ const sh=i===0?ss.getSheets()[0].setName(n):ss.insertSheet(n); sh.getRange('A:Z').setNumberFormat('@'); sh.getRange(1,1,1,ABAS[n].length).setValues([ABAS[n]]).setFontWeight('bold'); sh.setFrozenRows(1); });
  return ss; }
function aba_(n){ return ss_().getSheetByName(n); }
function agora_(){ return Utilities.formatDate(new Date(),'America/Recife','yyyy-MM-dd HH:mm:ss'); }
function linhas_(n){ const v=aba_(n).getDataRange().getValues(); return v.slice(1); }
function achar_(n,id){ const sh=aba_(n), last=sh.getLastRow(); if(last<2) return 0; const ids=sh.getRange(2,1,last-1,1).getValues(); for(let i=0;i<ids.length;i++) if(String(ids[i][0])===id) return i+2; return 0; }
function upsert_(n,id,vals){ const sh=aba_(n), row=achar_(n,id); if(row) sh.getRange(row,1,1,vals.length).setValues([vals]); else sh.appendRow(vals); }
function apagar_(n,id){ const row=achar_(n,id); if(row) aba_(n).deleteRow(row); }
function hash_(s){ return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,'parciais|'+s,Utilities.Charset.UTF_8)); }
function erro_(m,code){ const e=new Error(m); e.code=code||''; return e; }
function log_(a,p,d){ try{ aba_('Log').appendRow([agora_(),a,String(p||''),String(d||'').slice(0,500)]); }catch(e){} }
function cfgs_(){ const o={}; linhas_('Config').forEach(r=>{ try{ o[String(r[0])]={cfg:JSON.parse(r[1]),pin:String(r[2]||'')}; }catch(e){} }); return o; }
function auth_(b){
  if(b.senha){ const h=PR().getProperty('ADMIN_HASH');
    if(!h){ if(String(b.senha).length<6) throw erro_('A senha precisa ter 6 caracteres ou mais.','auth'); PR().setProperty('ADMIN_HASH',hash_(b.senha)); log_('senha_admin_criada','',''); return {admin:true}; }
    if(hash_(b.senha)!==h) throw erro_('Senha do administrador errada.','auth'); return {admin:true}; }
  if(b.authPosto&&b.pin){ const c=cfgs_()[String(b.authPosto)]; if(!c||!c.pin||hash_(String(b.pin))!==c.pin) throw erro_('Posto ou PIN errado.','auth'); return {posto:String(b.authPosto)}; }
  throw erro_('Faça o login.','auth'); }
function pode_(a,posto){ if(!a.admin&&a.posto!==String(posto)) throw erro_('Sem acesso a este posto.','perm'); }
function admin_(a){ if(!a.admin) throw erro_('Só o administrador pode fazer isso.','perm'); }
function montar_(aba,postos,de,ate){ const o={};
  linhas_(aba).forEach(r=>{ const p=String(r[1]), d=String(r[2]); if(postos.indexOf(p)<0||d<de||d>ate) return; let j; try{ j=JSON.parse(r[6]); }catch(e){ return; }
    const k=p+'_'+d, doc=o[k]=o[k]||{posto:p,data:d,parciais:{},nomes:{},setores:{}}; doc.parciais[String(r[3]).replace(':','')]=j.p; Object.assign(doc.nomes,j.nomes||{}); Object.assign(doc.setores,j.setores||{}); });
  return Object.keys(o).map(k=>o[k]); }
function doGet(){ return out_({ok:true,app:'Parciais de Venda – Servidor',versao:VERSAO}); }
function doPost(e){ try{ const b=JSON.parse(e.postData.contents); return out_({ok:true,r:acao_(b)}); }catch(err){ return out_({ok:false,erro:String(err.message||err),code:err.code||''}); } }
function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function acao_(b){ const A=b.action;
  if(A==='postos'){ const all=cfgs_(); return Object.keys(all).filter(k=>all[k].pin).map(k=>({codigo:k,nome:all[k].cfg.nome||''})); }
  if(A==='primeiroAcesso') return !PR().getProperty('ADMIN_HASH');
  const a=auth_(b);
  if(A==='login') return a;
  if(A==='cfgs'){ const all=cfgs_(), o={}; Object.keys(all).forEach(k=>{ if(a.admin||k===a.posto){ const c=all[k].cfg; c.temPin=!!all[k].pin; c.temTeams=!!PR().getProperty('TEAMS_'+k); o[k]=c; } }); return o; }
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    if(A==='cfgSave'){ admin_(a); const c=b.cfg, k=String(c.codigo); delete c.temPin; delete c.temTeams; const old=cfgs_()[k];
      const pin=b.pin?hash_(String(b.pin)):(old?old.pin:''); if(b.teams) PR().setProperty('TEAMS_'+k,String(b.teams));
      upsert_('Config',k,[k,JSON.stringify(c),pin,agora_()]); log_('cfgSave',k,b.pin?'PIN alterado':''); return true; }
    if(A==='dias'||A==='vends'){ const ps=(b.postos||[]).map(String).filter(p=>a.admin||p===a.posto); return montar_(A==='dias'?'Parciais':'Vendedores',ps,String(b.de),String(b.ate)); }
    if(A==='parcialSave'||A==='vendSave'){ pode_(a,b.posto); const aba=A==='parcialSave'?'Parciais':'Vendedores', p=b.p, id=b.posto+'_'+b.data+'_'+String(p.h).replace(':','');
      upsert_(aba,id,[id,String(b.posto),String(b.data),String(p.h),String(p.tipo),(p.lider||{}).nome||'',JSON.stringify({p:p,nomes:b.nomes||{},setores:b.setores||{}}),agora_()]);
      return montar_(aba,[String(b.posto)],String(b.data),String(b.data))[0]; }
    if(A==='parcialDel'){ admin_(a); apagar_('Parciais',b.posto+'_'+b.data+'_'+b.key); log_('parcialDel',b.posto,b.data+' '+b.key); return montar_('Parciais',[String(b.posto)],String(b.data),String(b.data))[0]||null; }
    if(A==='marcas'){ pode_(a,b.posto); const row=achar_('Marcas',b.posto+'_'+b.data); if(!row) return {}; try{ return JSON.parse(aba_('Marcas').getRange(row,4).getValue()); }catch(e){ return {}; } }
    if(A==='marcar'){ pode_(a,b.posto); const id=b.posto+'_'+b.data, row=achar_('Marcas',id); let m={}; if(row){ try{ m=JSON.parse(aba_('Marcas').getRange(row,4).getValue()); }catch(e){} }
      if(b.v) m[b.k]=b.v; else delete m[b.k]; upsert_('Marcas',id,[id,String(b.posto),String(b.data),JSON.stringify(m),agora_()]); return true; }
    if(A==='foto'){ pode_(a,b.posto); const row=achar_('Fotos',b.id); return row?String(aba_('Fotos').getRange(row,2).getValue()):null; }
    if(A==='fotoSave'){ admin_(a); if(b.img) upsert_('Fotos',b.id,[b.id,b.img,agora_()]); else apagar_('Fotos',b.id); return true; }
    if(A==='enviar'){ pode_(a,b.posto); return enviar_(b); }
  } finally { lock.releaseLock(); }
  throw erro_('Ação desconhecida: '+A); }
function enviar_(b){ const c=cfgs_()[String(b.posto)], res={email:false,teams:false}; const para=c&&c.cfg.emailPista;
  if(para){ const blob=Utilities.newBlob(Utilities.base64Decode(b.png),'image/png',b.nome||'parcial.png');
    MailApp.sendEmail({to:para,subject:b.assunto||'Parcial de venda',body:(b.texto||'')+'\n\nEnviado pelo app Parciais de Venda.',attachments:[blob]}); res.email=true; }
  const wh=PR().getProperty('TEAMS_'+b.posto);
  if(wh){ const card={type:'message',attachments:[{contentType:'application/vnd.microsoft.card.adaptive',content:{type:'AdaptiveCard',version:'1.4',$schema:'http://adaptivecards.io/schemas/adaptive-card.json',body:String(b.texto||'').split('\n').filter(String).map((t,i)=>({type:'TextBlock',text:t.replace(/\*/g,''),wrap:true,weight:i===0?'Bolder':'Default',spacing:'None'}))}}]};
    const r=UrlFetchApp.fetch(wh,{method:'post',contentType:'application/json',payload:JSON.stringify(card),muteHttpExceptions:true}); res.teams=r.getResponseCode()<300; }
  log_('enviar',b.posto,JSON.stringify(res)); return res; }
/** Rodar 1x no editor: cria a planilha e pede as autorizações (planilha, e-mail, internet). */
function configurar(){ const ss=ss_(); MailApp.getRemainingDailyQuota(); try{ UrlFetchApp.fetch('https://www.google.com',{muteHttpExceptions:true}); }catch(e){} Logger.log('Planilha: '+ss.getUrl()); }
/** Emergência: apaga a senha do administrador; o próximo login de administrador cria a nova. */
function redefinirSenhaAdmin(){ PR().deleteProperty('ADMIN_HASH'); Logger.log('Senha do administrador apagada.'); }
