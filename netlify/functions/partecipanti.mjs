import { getStore } from '@netlify/blobs';
import ExcelJS from 'exceljs';
import { randomUUID } from 'node:crypto';

const names = {
  atti: 'Andrea Atti',
  cocchianella: 'Massimo Cocchianella',
  florini: 'Ivan Florini'
};

const fields = [
  'dealer', 'codice', 'nome', 'cognome', 'ruolo',
  'email', 'edizioni', 'motivazione', 'priorita'
];

const clean = value => String(value ?? '').trim().slice(0, 1000);

function json(status, data, headers = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store',
      ...headers
    }
  });
}

function valid(data) {
  return data.dealer && data.codice && data.nome &&
    data.cognome && data.ruolo && data.priorita &&
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email);
}

async function all(store) {
  const { blobs } = await store.list({ prefix: 'partecipanti/' });
  const records = await Promise.all(
    blobs.map(blob => store.get(blob.key, {
      type: 'json',
      consistency: 'strong'
    }))
  );

  return records.filter(Boolean).sort((a, b) =>
    a.dealer.localeCompare(b.dealer, 'it') ||
    a.cognome.localeCompare(b.cognome, 'it')
  );
}

export default async function (request) {
  try {
    const store = getStore({
      name: 'corso-accettatori-2026',
      consistency: 'strong'
    });

    const method = request.method;
    const url = new URL(request.url);

    if (method === 'GET') {
      const records = await all(store);

      if (url.searchParams.get('export') === 'xlsx') {
        const book = new ExcelJS.Workbook();
        const sheet = book.addWorksheet('Partecipanti');

        sheet.columns = [
          ['SAM', 'sam', 22],
          ['Concessionaria', 'dealer', 32],
          ['Codice dealer', 'codice', 17],
          ['Nome', 'nome', 20],
          ['Cognome', 'cognome', 22],
          ['Ruolo', 'ruolo', 22],
          ['Email Web Academy', 'email', 35],
          ['Edizioni precedenti', 'edizioni', 25],
          ['Motivo della proposta', 'motivazione', 55],
          ['Priorità', 'priorita', 16],
          ['Inserito il', 'createdAt', 23],
          ['Aggiornato il', 'updatedAt', 23]
        ].map(([header, key, width]) => ({ header, key, width }));

        sheet.getRow(1).font = {
          bold: true,
          color: { argb: 'FFFFFFFF' }
        };

        sheet.getRow(1).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFB5121B' }
        };

        sheet.views = [{ state: 'frozen', ySplit: 1 }];
        sheet.autoFilter = { from: 'A1', to: 'L1' };

        for (const record of records) {
          sheet.addRow({
            ...record,
            sam: names[record.sam] || record.sam
          });
        }

        const buffer = await book.xlsx.writeBuffer();

        return new Response(Buffer.from(buffer), {
          status: 200,
          headers: {
            'Content-Type':
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition':
              'attachment; filename="Corso_Accettatori_2026_Partecipanti.xlsx"',
            'Cache-Control': 'no-store'
          }
        });
      }

      return json(200, { records });
    }

    if (!['POST', 'PUT', 'DELETE'].includes(method)) {
      return json(405, {
        error: 'Metodo non consentito.'
      }, {
        Allow: 'GET, POST, PUT, DELETE'
      });
    }

    let input;
    try {
      input = await request.json();
      if (!input || typeof input !== 'object' || Array.isArray(input)) {
        return json(400, { error: 'Dati non validi.' });
      }
    } catch {
      return json(400, { error: 'Dati non validi.' });
    }

    if (method === 'POST') {
      const data = Object.fromEntries(
        fields.map(key => [key, clean(input[key])])
      );

      if (!valid(data)) {
        return json(400, {
          error: 'Compila tutti i campi obbligatori e verifica l’email.'
        });
      }

      if (!Object.hasOwn(names, input.sam)) {
        return json(400, { error: 'Seleziona il SAM responsabile.' });
      }

      const now = new Date().toISOString();
      const record = {
        id: randomUUID(),
        sam: input.sam,
        ...data,
        createdAt: now,
        updatedAt: now
      };

      const saved = await store.setJSON(
        `partecipanti/${record.id}`,
        record,
        { onlyIfNew: true }
      );

      if (!saved.modified) {
        return json(409, { error: 'Riprova il salvataggio.' });
      }

      return json(201, { record });
    }

    const id = clean(input.id);

    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) {
      return json(400, { error: 'Identificativo non valido.' });
    }

    const key = `partecipanti/${id}`;
    const current = await store.getWithMetadata(key, {
      type: 'json',
      consistency: 'strong'
    });

    if (!current) {
      return json(404, { error: 'Nominativo non trovato.' });
    }

    if (method === 'DELETE') {
      await store.delete(key);
      return json(200, { ok: true });
    }

    if (input.updatedAt !== current.data.updatedAt) {
      return json(409, {
        error: 'Il nominativo è stato modificato da un altro utente. Aggiorna l’elenco.'
      });
    }

    const data = Object.fromEntries(
      fields.map(field => [field, clean(input[field])])
    );

    if (!valid(data)) {
      return json(400, {
        error: 'Compila tutti i campi obbligatori e verifica l’email.'
      });
    }

    if (!Object.hasOwn(names, input.sam)) {
      return json(400, { error: 'Seleziona il SAM responsabile.' });
    }

    const record = {
      ...current.data,
      ...data,
      sam: input.sam,
      updatedAt: new Date(
        Math.max(Date.now(), Date.parse(current.data.updatedAt) + 1)
      ).toISOString()
    };

    const saved = await store.setJSON(key, record, {
      onlyIfMatch: current.etag
    });

    if (!saved.modified) {
      return json(409, {
        error: 'Modifica contemporanea: aggiorna l’elenco e riprova.'
      });
    }

    return json(200, { record });
  } catch (error) {
    console.error('Errore archivio partecipanti:', error);

    const code = String(error?.name || 'UnknownError')
      .replace(/[^A-Za-z0-9]/g, '')
      .slice(0, 60);

    return json(500, {
      error: `Archivio non disponibile (${code}). Controlla i log della Function partecipanti.`,
      code
    });
  }
}
