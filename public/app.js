const $ = s => document.querySelector(s);
const names = { atti:'Andrea Atti', cocchianella:'Massimo Cocchianella', florini:'Ivan Florini' };
let state = { records:[] };
function flash(message,error=false){ const el=$('#message'); el.textContent=message; el.classList.toggle('error',error); el.hidden=false; clearTimeout(flash.timer); flash.timer=setTimeout(()=>el.hidden=true,4500); }
async function api(method='GET',data,exportFile=false){
 const res=await fetch('/api/partecipanti'+(exportFile?'?export=xlsx':''),{ method,headers:data?{'Content-Type':'application/json'}:{},body:data?JSON.stringify(data):undefined,cache:'no-store' });
 if(!res.ok){let problem;try{problem=(await res.json()).error}catch{};throw new Error(problem||`Errore ${res.status}`)}
 return exportFile?res.blob():res.json();
}
async function refresh(){state=await api();render();}
function escapeHTML(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function render(){
 const list=state.records, q=$('#search').value.trim().toLocaleLowerCase('it'), sam=$('#filterSam').value;
 $('#total').textContent=list.length;$('#high').textContent=list.filter(x=>x.priorita==='Alta').length;
 $('#dealers').textContent=new Set(list.map(x=>x.codice.trim().toLowerCase())).size;
 $('#mine').textContent=new Set(list.map(x=>x.sam)).size;
 const shown=list.filter(r=>(!sam||r.sam===sam)&&(!q||[r.dealer,r.codice,r.nome,r.cognome,r.email].some(v=>String(v||'').toLocaleLowerCase('it').includes(q))));
 $('#records').innerHTML=shown.length?shown.map(r=>`<article class="record"><div class="record-head"><div><h3>${escapeHTML(r.nome)} ${escapeHTML(r.cognome)}</h3><p class="dealer">${escapeHTML(r.dealer)} · ${escapeHTML(r.codice)}</p></div><span class="badge ${escapeHTML(r.priorita)}">${escapeHTML(r.priorita)}</span></div><p>${escapeHTML(r.ruolo)} · ${escapeHTML(names[r.sam])}</p><p>${escapeHTML(r.email)} · Precedenti: ${escapeHTML(r.edizioni||'—')}</p>${r.motivazione?`<p class="reason">${escapeHTML(r.motivazione)}</p>`:''}<div class="record-actions"><button data-action="edit" data-id="${r.id}">Modifica</button><button data-action="delete" data-id="${r.id}">Elimina</button></div></article>`).join(''):'<div class="empty">Nessun nominativo trovato.</div>';
}
function resetForm(){ $('#entryForm').reset();$('#entryForm').elements.id.value='';$('#entryForm').elements.updatedAt.value='';$('#formTitle').textContent='Aggiungi un accettatore';$('#save').firstChild.textContent='Salva nominativo ';$('#cancel').hidden=true; }
$('#entryForm').addEventListener('submit',async e=>{e.preventDefault();const form=e.currentTarget;const data=Object.fromEntries(new FormData(form));$('#save').disabled=true;try{await api(data.id?'PUT':'POST',data);resetForm();await refresh();flash('Nominativo salvato online.')}catch(err){flash(err.message,true)}finally{$('#save').disabled=false}});
$('#cancel').addEventListener('click',resetForm);
$('#records').addEventListener('click',async e=>{const b=e.target.closest('button[data-action]');if(!b)return;const r=state.records.find(x=>x.id===b.dataset.id);if(!r)return;if(b.dataset.action==='edit'){for(const key of ['id','updatedAt','sam','dealer','codice','nome','cognome','ruolo','email','edizioni','priorita','motivazione']){if($('#entryForm').elements[key])$('#entryForm').elements[key].value=r[key]||''}$('#formTitle').textContent='Modifica accettatore';$('#save').firstChild.textContent='Salva modifiche ';$('#cancel').hidden=false;window.scrollTo({top:0,behavior:'smooth'});return}if(confirm(`Eliminare ${r.nome} ${r.cognome} dall’elenco?`)){try{await api('DELETE',{id:r.id});await refresh();flash('Nominativo eliminato.')}catch(err){flash(err.message,true)}}});
$('#refresh').addEventListener('click',()=>refresh().then(()=>flash('Elenco aggiornato.')).catch(err=>flash(err.message,true)));
$('#search').addEventListener('input',render);$('#filterSam').addEventListener('change',render);
$('#export').addEventListener('click',async()=>{try{const blob=await api('GET',null,true);const url=URL.createObjectURL(blob);const link=document.createElement('a');link.href=url;link.download='Corso_Accettatori_2026_Partecipanti.xlsx';link.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}catch(err){flash(err.message,true)}});
refresh().catch(err=>flash(err.message,true));
