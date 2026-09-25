const express = require('express');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');
const { parse } = require('csv-parse/sync');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_BASE = 'https://api.g1.datacrazy.io/api/v1';

app.use(express.json({ limit: '8mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const SOURCES = [
  {
    name: 'Google ADS Start I - Form Padrão',
    channel: 'Google Ads',
    regional: 'Start I / Form Padrão',
    landingPage: 'https://unifateciestart.com.br/',
    sheetId: '11muw-TF0tSOHQxmg9JmKSifrBSzesqRSrB53r9IqcPA',
    sheetUrl: 'https://docs.google.com/spreadsheets/d/11muw-TF0tSOHQxmg9JmKSifrBSzesqRSrB53r9IqcPA/edit?usp=sharing'
  },
  {
    name: 'Google ADS Start II - Form WhatsApp',
    channel: 'Google Ads',
    regional: 'Start II / WhatsApp',
    landingPage: 'https://unifateciestart.com.br/fale-pelo-whatsapp/',
    sheetId: '1qjtrTpE2wnpziFQQK8ADejs_qOWtPXG_SigjywsJkQU',
    sheetUrl: 'https://docs.google.com/spreadsheets/d/1qjtrTpE2wnpziFQQK8ADejs_qOWtPXG_SigjywsJkQU/edit?usp=sharing'
  },
  {
    name: 'Meta ADS Start - SP-INTERIOR',
    channel: 'Meta Ads',
    regional: 'SP-INTERIOR',
    landingPage: 'https://unifateciestart.com.br/fale-pelo-whatsapp-spinterior',
    sheetId: '1cRpAzUkQB8VBWKbMvBK-WrxTMNph34bo40NomC95ey0',
    sheetUrl: 'https://docs.google.com/spreadsheets/d/1cRpAzUkQB8VBWKbMvBK-WrxTMNph34bo40NomC95ey0/edit?usp=sharing'
  },
  {
    name: 'Meta ADS Start - SAMPA',
    channel: 'Meta Ads',
    regional: 'SAMPA',
    landingPage: 'https://unifateciestart.com.br/fale-pelo-whatsapp-sampa',
    sheetId: '1fimvMBO_Mk8Ye7m7cOFOy6aWN_jXSWrxCtQYIwAASLw',
    sheetUrl: 'https://docs.google.com/spreadsheets/d/1fimvMBO_Mk8Ye7m7cOFOy6aWN_jXSWrxCtQYIwAASLw/edit?usp=sharing'
  },
  {
    name: 'Meta ADS Start - MT-GO',
    channel: 'Meta Ads',
    regional: 'MT-GO',
    landingPage: 'https://unifateciestart.com.br/fale-pelo-whatsapp-mtgo',
    sheetId: '1wATr-nTWIDi2-Fz0HKbQfrFcv7bZkqTQqPtMF0jTIXY',
    sheetUrl: 'https://docs.google.com/spreadsheets/d/1wATr-nTWIDi2-Fz0HKbQfrFcv7bZkqTQqPtMF0jTIXY/edit?usp=sharing'
  }
];

const CONSULTANTS = [
  'Aline Fatecie',
  'Ana Regina',
  'Bianca',
  'Bruna',
  'grazielly',
  'KETHYLEEN CRISTINE DE ARAGAO ANDRADE',
  'Maria Vitoria Matrículas',
  'Tamires',
  'Tatiana'
];

const jobs = new Map();
let originCache = { expiresAt: 0, lookup: null, status: [] };

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase();
}

function normalizePhone(value) {
  let digits = String(value || '').replace(/\D/g, '');
  if (!digits) return '';
  if (digits.length >= 12 && digits.startsWith('55')) digits = digits.slice(2);
  if (digits.length > 11) digits = digits.slice(-11);
  return digits.length >= 10 ? digits : '';
}

function getToken(req) {
  return (process.env.DATACRAZY_TOKEN || req.headers['x-datacrazy-token'] || '').trim();
}

function authGuard(req, res, next) {
  const configured = String(process.env.APP_PASSWORD || '').trim();
  if (!configured) return next();
  const supplied = String(req.headers['x-app-password'] || '');
  if (supplied !== configured) return res.status(401).json({ error: 'Senha do painel inválida.' });
  next();
}

app.use('/api', (req, res, next) => {
  if (req.path === '/config') return next();
  return authGuard(req, res, next);
});

async function dcFetch(url, token, state) {
  let attempt = 0;
  while (true) {
    attempt++;
    if (state) {
      const elapsed = Date.now() - state.lastRequestAt;
      const remain = state.delayMs - elapsed;
      if (remain > 0) await sleep(remain + Math.floor(Math.random() * 10));
      state.lastRequestAt = Date.now();
      state.requests++;
    }

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }
    });

    if (response.status === 429) {
      if (state) {
        state.rateLimits++;
        state.delayMs = Math.min(5000, Math.max(350, Math.ceil(state.delayMs * 1.75)));
      }
      if (attempt >= 15) throw new Error('HTTP 429 persistente na API DataCrazy.');
      const retryAfter = Number(response.headers.get('retry-after') || 0);
      await sleep((retryAfter > 0 ? retryAfter + 1 : Math.min(30, 2 + attempt)) * 1000);
      continue;
    }

    const text = await response.text();
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}${text ? ` - ${text.slice(0, 500)}` : ''}`);
    }

    if (state && state.delayMs > 50) state.delayMs = Math.max(50, state.delayMs - 10);
    return text ? JSON.parse(text) : null;
  }
}

function dataArray(payload) {
  if (!payload) return [];
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload.data)) return payload.data;
  if (payload.data && Array.isArray(payload.data.items)) return payload.data.items;
  if (Array.isArray(payload.items)) return payload.items;
  return [];
}

function currentConversation(c) {
  if (!c) return false;
  if (c.deletedAt || c.archivedAt || c.hidden === true || c.finished === true) return false;
  return true;
}

function hasAttendant(c) {
  if (Array.isArray(c?.attendants)) return c.attendants.length > 0;
  if (Array.isArray(c?.currentThread?.attendants)) return c.currentThread.attendants.length > 0;
  return false;
}

function attendantIds(c) {
  const src = Array.isArray(c?.attendants) ? c.attendants : (Array.isArray(c?.currentThread?.attendants) ? c.currentThread.attendants : []);
  return [...new Set(src.map(a => typeof a === 'string' ? a : a?.id).filter(Boolean))];
}

function conversationRow(c, status, situacao, attendantMap = {}) {
  const phone = c?.contact?.contactId || c?.contact?.id || '';
  const ids = attendantIds(c);
  return {
    Nome: c?.name || c?.contact?.name || '',
    Telefone: phone,
    Plataforma: c?.contact?.platform || '',
    ConversationId: c?.id || '',
    Status: status,
    Situacao: situacao,
    Departamento: c?.currentDepartment?.name || '',
    Atendentes: ids.map(id => attendantMap[id] || id).join(', '),
    AtualizadoEm: c?.updatedAt || ''
  };
}

async function selectTake(token, state, job) {
  const candidates = [5000, 3000, 2000, 1500, 1000, 750, 500, 250, 100];
  for (const take of candidates) {
    try {
      if (job) job.message = `Testando paginação: take ${take}`;
      const url = `${DATA_BASE}/conversations?skip=0&take=${take}&filter%5BopenWindow%5D=all&filter%5Bstatus%5D=unstarted`;
      const page = dataArray(await dcFetch(url, token, state));
      if (page.length > 0) {
        if (page.length < take && page.length >= 50) return page.length;
        return take;
      }
    } catch (_) {
      // tenta um tamanho menor
    }
  }
  return 100;
}

async function loadAttendants(token, state) {
  try {
    const payload = await dcFetch(`${DATA_BASE}/attendants/multi`, token, state);
    const map = {};
    for (const a of dataArray(payload)) {
      if (a?.id) map[String(a.id)] = a?.name || '';
      if (a?.userId) map[String(a.userId)] = a?.name || '';
    }
    return { map, raw: dataArray(payload) };
  } catch (_) {
    return { map: {}, raw: [] };
  }
}

async function collectStatus(token, take, status, state, job, label) {
  let skip = 0;
  let pageNum = 0;
  let noNew = 0;
  const map = new Map();

  while (true) {
    pageNum++;
    if (job) job.message = `${label}: página ${pageNum} • ${map.size} únicos`;
    const url = `${DATA_BASE}/conversations?skip=${skip}&take=${take}&filter%5BopenWindow%5D=all&filter%5Bstatus%5D=${encodeURIComponent(status)}`;
    const page = dataArray(await dcFetch(url, token, state));
    if (!page.length) break;

    const before = map.size;
    for (const item of page) if (item?.id) map.set(String(item.id), item);
    noNew = map.size === before ? noNew + 1 : 0;
    if (noNew >= 3) throw new Error(`API repetiu 3 páginas sem registros novos em ${label}.`);

    skip += take;
    if (skip > 1000000) throw new Error(`Limite de segurança atingido em ${label}.`);
  }
  return map;
}

async function getOriginLookup(force = false, job = null) {
  if (!force && originCache.lookup && Date.now() < originCache.expiresAt) return originCache;

  const lookup = new Map();
  const status = [];

  for (const src of SOURCES) {
    try {
      if (job) job.message = `Origem: carregando ${src.name}`;
      const url = `https://docs.google.com/spreadsheets/d/${src.sheetId}/export?format=csv`;
      const response = await fetch(url, { redirect: 'follow' });
      const text = await response.text();
      if (!response.ok || !text || text.trimStart().startsWith('<')) throw new Error('planilha não retornou CSV');

      const rows = parse(text, { columns: true, skip_empty_lines: true, relax_column_count: true, bom: true });
      const headers = rows.length ? Object.keys(rows[0]) : [];
      let phoneHeaders = headers.filter(h => /TELEF|CELULAR|WHATS|PHONE|FONE|CONTATO/i.test(normalizeText(h)));
      if (!phoneHeaders.length) phoneHeaders = headers;
      let added = 0;

      for (const row of rows) {
        for (const h of phoneHeaders) {
          const phone = normalizePhone(row[h]);
          if (!phone) continue;
          if (!lookup.has(phone)) lookup.set(phone, []);
          const arr = lookup.get(phone);
          if (!arr.some(x => x.name === src.name)) {
            arr.push(src);
            added++;
          }
        }
      }
      status.push({ fonte: src.name, status: 'OK', telefones: added });
    } catch (err) {
      status.push({ fonte: src.name, status: 'ERRO', erro: err.message, telefones: 0 });
    }
  }

  originCache = { expiresAt: Date.now() + 5 * 60 * 1000, lookup, status };
  return originCache;
}

function resolveOrigin(phone, lookup) {
  const p = normalizePhone(phone);
  if (!p) return {
    Origem: 'NÃO LOCALIZADO', OrigemDetalhada: 'Telefone inválido/ausente', Regional: '',
    ValidacaoOrigem: 'SEM TELEFONE VÁLIDO', FonteOrigem: '', LandingPage: ''
  };
  if (!lookup.has(p)) return {
    Origem: 'NÃO LOCALIZADO', OrigemDetalhada: 'Não encontrado nas planilhas de captação', Regional: '',
    ValidacaoOrigem: 'NÃO ENCONTRADO', FonteOrigem: '', LandingPage: ''
  };
  const matches = lookup.get(p);
  if (matches.length === 1) {
    const m = matches[0];
    return {
      Origem: m.channel, OrigemDetalhada: m.name, Regional: m.regional,
      ValidacaoOrigem: 'CONFIRMADA', FonteOrigem: m.sheetUrl, LandingPage: m.landingPage
    };
  }
  return {
    Origem: 'CONFLITO',
    OrigemDetalhada: matches.map(m => m.name).join(' | '),
    Regional: matches.map(m => m.regional).join(' | '),
    ValidacaoOrigem: 'REVISAR - TELEFONE EM MAIS DE UMA FONTE',
    FonteOrigem: matches.map(m => m.sheetUrl).join(' | '),
    LandingPage: matches.map(m => m.landingPage).join(' | ')
  };
}

async function runCollection(job, token) {
  const state = { delayMs: 50, lastRequestAt: 0, rateLimits: 0, requests: 0 };
  try {
    job.status = 'running';
    job.phase = 'auth';
    job.message = 'Testando API DataCrazy...';
    await dcFetch(`${DATA_BASE}/conversations?skip=0&take=1&filter%5BopenWindow%5D=all`, token, state);

    job.phase = 'pagination';
    const take = await selectTake(token, state, job);

    job.phase = 'attendants';
    job.message = 'Carregando atendentes...';
    const attendants = await loadAttendants(token, state);

    const normalMap = new Map();
    const statusById = new Map();
    const statuses = [
      ['unstarted', 'SEM ATENDIMENTO INICIADO'],
      ['waiting', 'AGUARDANDO'],
      ['opened', 'ABERTO'],
      ['automation', 'AUTOMAÇÃO']
    ];

    for (let i = 0; i < statuses.length; i++) {
      const [status, label] = statuses[i];
      job.phase = `status-${status}`;
      job.progress = 10 + i * 15;
      const map = await collectStatus(token, take, status, state, job, label);
      for (const [id, c] of map) {
        if (currentConversation(c)) {
          normalMap.set(id, c);
          statusById.set(id, status);
        }
      }
    }

    job.phase = 'status-error';
    job.progress = 70;
    const failedRaw = await collectStatus(token, take, 'error', state, job, 'FAILED');
    const failedMap = new Map();
    for (const [id, c] of failedRaw) {
      if (!currentConversation(c)) continue;
      failedMap.set(id, c);
      normalMap.delete(id);
      statusById.delete(id);
    }

    const sem = [];
    const failedSem = [];
    const failedCom = [];

    for (const [id, c] of normalMap) {
      if (!hasAttendant(c)) sem.push(conversationRow(c, statusById.get(id), 'sem atendente', attendants.map));
    }
    for (const [, c] of failedMap) {
      if (hasAttendant(c)) failedCom.push(conversationRow(c, 'error', 'failed + com atendente', attendants.map));
      else failedSem.push(conversationRow(c, 'error', 'failed + sem atendente', attendants.map));
    }

    const allNoAttendant = [...sem, ...failedSem];

    job.phase = 'origin';
    job.progress = 85;
    const origins = await getOriginLookup(true, job);
    const enriched = allNoAttendant.map(row => ({ ...row, ...resolveOrigin(row.Telefone, origins.lookup) }));

    enriched.sort((a, b) => String(a.Nome || '').localeCompare(String(b.Nome || ''), 'pt-BR'));

    job.result = {
      rows: enriched,
      failedWithAttendant: failedCom,
      summary: {
        semAtendenteNaoFailed: sem.length,
        failedSemAtendente: failedSem.length,
        totalSemAtendente: allNoAttendant.length,
        failedComAtendente: failedCom.length,
        totalFailed: failedSem.length + failedCom.length,
        origemConfirmada: enriched.filter(r => r.ValidacaoOrigem === 'CONFIRMADA').length,
        origemConflito: enriched.filter(r => r.Origem === 'CONFLITO').length,
        origemNaoLocalizada: enriched.filter(r => r.Origem === 'NÃO LOCALIZADO').length,
        take,
        requests: state.requests,
        rateLimits: state.rateLimits,
        delayFinalMs: state.delayMs
      },
      sourceStatus: origins.status,
      collectedAt: new Date().toISOString()
    };
    job.status = 'done';
    job.phase = 'done';
    job.progress = 100;
    job.message = `Concluído: ${enriched.length} conversas sem atendente.`;
  } catch (err) {
    job.status = 'error';
    job.error = err.message;
    job.message = err.message;
  }
}

app.get('/health', (req, res) => res.json({ ok: true, service: 'lopex-datacrazy', now: new Date().toISOString() }));

app.get('/api/config', (req, res) => {
  res.json({
    appPasswordConfigured: Boolean(String(process.env.APP_PASSWORD || '').trim()),
    dataCrazyTokenConfigured: Boolean(String(process.env.DATACRAZY_TOKEN || '').trim()),
    sources: SOURCES.map(({ sheetId, ...s }) => s),
    consultants: CONSULTANTS
  });
});

app.post('/api/collect/start', async (req, res) => {
  const token = getToken(req);
  if (!token) return res.status(400).json({ error: 'Informe a chave API do DataCrazy.' });

  const id = crypto.randomUUID();
  const job = {
    id,
    status: 'queued',
    phase: 'queued',
    progress: 0,
    message: 'Na fila...',
    startedAt: new Date().toISOString(),
    result: null,
    error: null
  };
  jobs.set(id, job);
  res.status(202).json({ jobId: id });
  setImmediate(() => runCollection(job, token));
});

app.get('/api/jobs/:id', (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: 'Coleta não encontrada.' });
  res.json(job);
});

app.get('/api/history/:conversationId', async (req, res) => {
  try {
    const token = getToken(req);
    if (!token) return res.status(400).json({ error: 'Informe a chave API do DataCrazy.' });
    const id = encodeURIComponent(req.params.conversationId);
    const payload = await dcFetch(`${DATA_BASE}/conversations/${id}/messages`, token, null);
    const messages = dataArray(payload).map(m => ({
      createdAt: m?.createdAt || '',
      direction: typeof m?.received === 'boolean' ? (m.received ? 'RECEBIDA' : 'ENVIADA') : '',
      status: m?.status || '',
      body: m?.body || (Array.isArray(m?.attachments) && m.attachments.length ? `[ANEXO x${m.attachments.length}]` : ''),
      errorCode: m?.erroCode || m?.errorCode || '',
      errorMessage: m?.errorMessage || '',
      attendants: Array.isArray(m?.attendants) ? m.attendants.map(a => typeof a === 'string' ? a : (a?.name || a?.id || '')).filter(Boolean).join(', ') : ''
    })).sort((a, b) => new Date(a.createdAt || 0) - new Date(b.createdAt || 0));
    res.json({ messages });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/round-robin', async (req, res) => {
  try {
    const token = getToken(req);
    if (!token) return res.status(400).json({ error: 'Informe a chave API do DataCrazy.' });
    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];
    let startIndex = Number(req.body?.startIndex || 0);
    if (!Number.isInteger(startIndex) || startIndex < 0 || startIndex >= CONSULTANTS.length) startIndex = 0;
    if (!rows.length) return res.status(400).json({ error: 'Nenhuma conversa validada foi enviada.' });

    const attendants = await loadAttendants(token, null);
    const byName = new Map();
    for (const a of attendants.raw) {
      const name = String(a?.name || '');
      const id = String(a?.id || '');
      if (!name || !id) continue;
      const key = normalizeText(name);
      if (!byName.has(key)) byName.set(key, []);
      byName.get(key).push({ name, id });
    }

    const consultantObjects = [];
    const issues = [];
    for (const name of CONSULTANTS) {
      const matches = byName.get(normalizeText(name)) || [];
      if (matches.length !== 1) {
        issues.push(`${name}: ${matches.length ? `${matches.length} correspondências` : 'não encontrado'}`);
      } else consultantObjects.push({ name, id: matches[0].id });
    }
    if (issues.length) return res.status(409).json({ error: 'Não foi possível validar os 9 consultores.', issues });

    const ordered = [...rows].sort((a, b) => {
      const da = Date.parse(a.AtualizadoEm || '') || Number.MAX_SAFE_INTEGER;
      const db = Date.parse(b.AtualizadoEm || '') || Number.MAX_SAFE_INTEGER;
      return da - db || String(a.Nome || '').localeCompare(String(b.Nome || ''), 'pt-BR');
    });

    const counts = Object.fromEntries(CONSULTANTS.map(c => [c, 0]));
    const plan = ordered.map((row, i) => {
      const consultant = consultantObjects[(startIndex + i) % consultantObjects.length];
      counts[consultant.name]++;
      return {
        Ordem: i + 1,
        ...row,
        ConsultorDestino: consultant.name,
        AtendenteId: consultant.id
      };
    });

    const nextIndex = (startIndex + plan.length) % CONSULTANTS.length;
    const summary = consultantObjects.map(c => ({ Consultor: c.name, AtendenteId: c.id, Quantidade: counts[c.name] }));
    res.json({ plan, summary, nextIndex, nextConsultant: consultantObjects[nextIndex].name, transferExecuted: false });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/export/xlsx', (req, res) => {
  try {
    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];
    const title = String(req.body?.sheetName || 'LEADS').slice(0, 31);
    const filename = String(req.body?.filename || 'datacrazy.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    const headers = rows.length ? Object.keys(rows[0]) : [];
    ws['!cols'] = headers.map(h => ({ wch: Math.min(42, Math.max(12, String(h).length + 2)) }));
    XLSX.utils.book_append_sheet(wb, ws, title || 'LEADS');
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Lopex DataCrazy ouvindo na porta ${PORT}`);
});
