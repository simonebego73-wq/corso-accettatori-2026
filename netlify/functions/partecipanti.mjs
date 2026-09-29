import { getStore } from '@netlify/blobs';
import ExcelJS from 'exceljs';
import { randomUUID } from 'node:crypto';

const names = { atti: 'Andrea Atti', cocchianella: 'Massimo Cocchianella', florini: 'Ivan Florini', admin: 'Simone Beghelli' };
const fields = ['dealer', 'codice', 'nome', 'cognome', 'ruolo', 'email', 'edizioni', 'motivazione', 'priorita'];
const clean = value => String(value ?? '').trim().slice(0, 1000);
const response = (statusCode, body, headers = {}) => ({ statusCode, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', ...headers }, body: JSON.stringify(body) });
async function all(store) {
  const { blobs } = await store.list({ prefix: 'partecipanti/' });
  const records = await Promise.all(blobs.map(x => store.get(x.key, { type: 'json', consistency: 'strong' })));
  return records.filter(Boolean).sort((a, b) => a.dealer.localeCompare(b.dealer, 'it') || a.cognome.localeCompare(b.cognome, 'it'));
}
export async function handler(event) {
  try {
    const store = getStore('corso-accettatori-2026');
    const method = event.httpMethod;
    if (method === 'GET') {
      const records = await all(store);
      if (event.queryStringParameters?.export === 'xlsx') {
        const book = new ExcelJS.Workbook();
        const sheet = book.addWorksheet('Partecipanti');
        sheet.columns = [
          ['SAM','sam',22], ['Concessionaria','dealer',32], ['Codice dealer','codice',17],
          ['Nome','nome',20], ['Cognome','cognome',22], ['Ruolo','ruolo',22],
          ['Email Web Academy','email',35], ['Edizioni precedenti','edizioni',25],
          ['Motivo della proposta','motivazione',55], ['Priorità','priorita',16],
          ['Inserito il','createdAt',23], ['Aggiornato il','updatedAt',23]
        ].map(([header,key,width]) => ({header,key,width}));
        sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB5121B' } };
        sheet.views = [{ state:'frozen', ySplit:1 }];
        sheet.autoFilter = { from: 'A1', to: 'L1' };
        for (const r of records) sheet.addRow({ ...r, sam: names[r.sam] || r.sam });
        const buffer = await book.xlsx.writeBuffer();
        return { statusCode: 200, isBase64Encoded: true, headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="Corso_Accettatori_2026_Partecipanti.xlsx"', 'Cache-Control': 'no-store' }, body: Buffer.from(buffer).toString('base64') };
      }
      return response(200, { records });
    }
    if (!['POST','PUT','DELETE'].includes(method)) return response(405, { error:'Metodo non consentito.' }, { Allow:'GET, POST, PUT, DELETE' });
    let input;
    try { input = JSON.parse(event.body || '{}'); } catch { return response(400, { error:'Dati non validi.' }); }
    if (method === 'POST') {
      const data = Object.fromEntries(fields.map(k => [k, clean(input[k])]));
      if (!data.dealer || !data.codice || !data.nome || !data.cognome || !data.ruolo || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email) || !data.priorita) return response(400, { error:'Compila tutti i campi obbligatori e verifica l’email.' });
      const now = new Date().toISOString();
      if (!names[input.sam] || input.sam === 'admin') return response(400, { error:'Seleziona il SAM responsabile.' });
      const record = { id:randomUUID(), sam:input.sam, ...data, createdAt:now, updatedAt:now };
      await store.setJSON(`partecipanti/${record.id}`, record, { onlyIfNew: true });
      return response(201, { record });
    }
    const id = clean(input.id);
    if (!/^[0-9a-f-]{36}$/i.test(id)) return response(400, { error:'Identificativo non valido.' });
    const key = `partecipanti/${id}`;
    const current = await store.getWithMetadata(key, { type:'json', consistency:'strong' });
    if (!current) return response(404, { error:'Nominativo non trovato.' });
    if (method === 'DELETE') { await store.delete(key); return response(200, { ok:true }); }
    if (input.updatedAt !== current.data.updatedAt) return response(409, { error:'Il nominativo è stato modificato da un altro utente. Aggiorna l’elenco.' });
    const data = Object.fromEntries(fields.map(k => [k, clean(input[k])]));
    if (!data.dealer || !data.codice || !data.nome || !data.cognome || !data.ruolo || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email) || !data.priorita) return response(400, { error:'Compila tutti i campi obbligatori e verifica l’email.' });
    if (!names[input.sam] || input.sam === 'admin') return response(400, { error:'Seleziona il SAM responsabile.' });
    const record = { ...current.data, ...data, sam:input.sam, updatedAt:new Date().toISOString() };
    const saved = await store.setJSON(key, record, { onlyIfMatch: current.etag });
    if (!saved.modified) return response(409, { error:'Modifica contemporanea: aggiorna l’elenco e riprova.' });
    return response(200, { record });
  } catch (err) {
    console.error('Errore archivio partecipanti:', err);
    const kind = String(err?.name || 'UnknownError').replace(/[^A-Za-z0-9]/g, '').slice(0, 60);
    const details = {
      MissingBlobsEnvironmentError: 'Netlify Blobs non è configurato nel deploy. Pubblica il progetto sorgente tramite Git o Netlify CLI con build delle Functions.',
      BlobsStoreNotFoundError: 'Archivio Netlify Blobs non trovato. Verifica il deploy delle Functions.'
    };
    return response(500, { error:details[kind] || `Archivio non disponibile (${kind}). Controlla i log della Function partecipanti.`, code:kind });
  }
}
