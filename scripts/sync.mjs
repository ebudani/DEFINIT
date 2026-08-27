// Sincroniza datos de ventas/otorgamiento desde la API EVUP y publica
// data/dicts.json + data/rows.json + data/meta.json para el tablero estático.
//
// Corre exclusivamente server-side (GitHub Actions). Las credenciales viven
// en secrets del repo (EVUP_LOGIN / EVUP_PASSWORD) y NUNCA se escriben a disco
// ni se incluyen en los archivos publicados. Los archivos de salida son
// agregados/anonimizados: nunca se persiste nombre, teléfono, e-mail, CPF/CNPJ,
// dirección ni fecha de nacimiento de clientes.
//
// Uso local (para probar antes de confiar en el workflow):
//   EVUP_LOGIN=... EVUP_PASSWORD=... node scripts/sync.mjs

import { writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT_DIR = path.join(__dirname, '..', 'data');
const API_BASE = process.env.EVUP_API_BASE || 'https://definit.api.evup.dev';
const CONCURRENCY = Number(process.env.SYNC_CONCURRENCY || 2);
const MAX_RETRIES = 7;

const LOGIN = process.env.EVUP_LOGIN;
const PASSWORD = process.env.EVUP_PASSWORD;

if (!LOGIN || !PASSWORD) {
  console.error('Faltan EVUP_LOGIN / EVUP_PASSWORD en el entorno. Configuralos como secrets del repo.');
  process.exit(1);
}

// Mapa de estados de la API (EBudgetStatus) a las etiquetas usadas en el tablero.
const STATUS_LABELS = {
  underNegotiation: 'Em Negociação',
  waitingApproval: 'Aguardando Aprovação',
  readyForSale: 'Pronto para Venda',
  finalized: 'Finalizado',
  canceled: 'Cancelado',
  expired: 'Expirado',
  paid: 'Pago',
  awaitingSignature: 'Aguardando Assinatura',
  inTransfer: 'Em Transferência',
  transfered: 'Transferido',
  awaitingEndingTerm: 'Aguardando Termo de Encerramento',
  awaitingTransferTerm: 'Aguardando Termo Transferencia',
  awaitingCancelationTerm: 'Aguardando Termo de Cancelamento',
  awaitingFinePayment: 'Aguardando Pagamento de Multa',
  waitingCancelationApproval: 'Aguardando Aprovação de Cancelamento',
  awaitingPreCancelationTerm: 'Aguardando Termo Pré-Cancelamento',
  partiallyCanceled: 'Parcialmente Cancelado',
  awaitingCardCancelation: 'Aguardando Cancelamento de Cartão',
  awaitingAcceptTerm: 'Aguardando Aceite de Termo',
  awaitingPaymentConfirmation: 'Aguardando Confirmação de Pagamento',
  awaitingCancel: 'Aguardando Cancelamento',
  overdue: 'Em Atraso',
  waitingCancellationConfirmation: 'Aguardando Confirmação de Cancelamento',
  suspectedContract: 'Contrato Suspeito',
  awaitingRenegotiationTerm: 'Aguardando Termo de Renegociação',
};

const GENDER_LABELS = { nothing: 'NÃO INFORMADO', female: 'FEMININO', male: 'MASCULINO', others: 'OUTROS' };

// ---------------- HTTP helpers ----------------
let token = null;

async function login() {
  const res = await fetch(`${API_BASE}/Auth/Login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ Login: LOGIN, Password: PASSWORD }),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Login falló: HTTP ${res.status} ${body}`);
  }
  let text = await res.text();
  // El endpoint puede devolver el token como string JSON-quoteado o texto plano.
  try {
    const parsed = JSON.parse(text);
    if (typeof parsed === 'string') text = parsed;
    else if (parsed && typeof parsed === 'object') {
      text = parsed.token || parsed.Token || parsed.accessToken || parsed.AccessToken || text;
    }
  } catch { /* texto plano, se usa tal cual */ }
  token = text.trim().replace(/^"|"$/g, '');
  if (!token) throw new Error('Login OK pero no se pudo extraer el token de la respuesta.');
}

async function apiPost(pathName, body, { retry = true } = {}) {
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const res = await fetch(`${API_BASE}${pathName}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify(body),
    });
    if (res.status === 401 && retry && attempt === 0) {
      console.warn(`401 en ${pathName}, reintentando login...`);
      await login();
      continue;
    }
    if (res.status === 429 || res.status >= 500) {
      const retryAfterHeader = res.headers.get('retry-after');
      const rateHeaders = [...res.headers.entries()].filter(([k]) => /rate|retry/i.test(k));
      const errBody = await res.text().catch(() => '');
      const retryAfterMs = retryAfterHeader ? Number(retryAfterHeader) * 1000 : null;
      const backoff = retryAfterMs && !Number.isNaN(retryAfterMs) ? retryAfterMs : Math.min(60000, 2000 * 2 ** attempt);
      console.warn(`HTTP ${res.status} en ${pathName} (intento ${attempt + 1}/${MAX_RETRIES + 1}). Headers: ${JSON.stringify(rateHeaders)}. Body: ${errBody.slice(0, 400)}. Reintento en ${backoff}ms.`);
      await sleep(backoff);
      continue;
    }
    if (!res.ok) {
      const errBody = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status} en ${pathName}: ${errBody.slice(0, 300)}`);
    }
    return res.json();
  }
  throw new Error(`Se agotaron los reintentos para ${pathName}`);
}

function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

// Extrae un array de items de una respuesta paginada, tolerando distintas
// convenciones de envoltorio ({items:[]}, {Items:[]}, array plano, etc.)
// ya que el swagger no documenta el schema exacto de las respuestas List.
function extractItems(json) {
  if (Array.isArray(json)) return { items: json, hasMore: null };
  if (json && typeof json === 'object') {
    const items = json.items || json.Items || json.data || json.Data || [];
    const hasMore =
      json.hasNextPage ?? json.HasNextPage ?? json.hasMore ?? json.HasMore ?? null;
    return { items, hasMore };
  }
  return { items: [], hasMore: false };
}

async function fetchAllPages(pathName, baseBody) {
  const all = [];
  let pageNumber = 1;
  while (true) {
    const json = await apiPost(pathName, { ...baseBody, PageNumber: pageNumber });
    const { items, hasMore } = extractItems(json);
    all.push(...items);
    if (items.length === 0) break;
    if (hasMore === false) break;
    if (hasMore === null && items.length < 1) break;
    pageNumber++;
    if (pageNumber > 2000) { console.warn(`Corte de seguridad de paginación en ${pathName}`); break; }
  }
  return all;
}

// Pool de concurrencia simple.
async function mapPool(items, limit, fn) {
  const results = new Array(items.length);
  let i = 0;
  let active = 0;
  return new Promise((resolve, reject) => {
    let doneCount = 0;
    function kick() {
      if (i >= items.length && active === 0) { resolve(results); return; }
      while (active < limit && i < items.length) {
        const idx = i++;
        active++;
        fn(items[idx], idx)
          .then((r) => { results[idx] = r; })
          .catch(reject)
          .finally(() => { active--; doneCount++; kick(); });
      }
    }
    kick();
  });
}

// ---------------- dictionaries ----------------
function makeDict() {
  const map = new Map();
  return {
    idx(key) {
      if (!map.has(key)) map.set(key, map.size);
      return map.get(key);
    },
    toArray() {
      const arr = new Array(map.size);
      for (const [k, v] of map.entries()) arr[v] = k;
      return arr;
    },
    get size() { return map.size; },
  };
}

// ---------------- main ----------------
async function main() {
  console.log(`Sincronizando desde ${API_BASE} ...`);
  await login();
  console.log('Login OK.');

  console.log('Listando clientes...');
  const rawClients = await fetchAllPages('/Client/List', { Filters: [], Sort: [{ Field: 'Id', Operation: 'asc' }] });
  console.log(`Clientes encontrados: ${rawClients.length}`);

  // Se anonimiza en el momento: solo se conserva id, género y canal de captación.
  const clients = rawClients.map((c) => ({
    id: c.Id ?? c.id,
    gender: GENDER_LABELS[c.Gender ?? c.gender] || 'NÃO INFORMADO',
    mediaCliente: (c.Media && (c.Media.Name ?? c.Media.name)) || 'SEM ORIGEM',
  })).filter((c) => c.id != null);

  const dictItem = makeDict();
  const dictVendedor = makeDict();
  const dictGenero = makeDict();
  const dictMediaCliente = makeDict();
  const dictMediaContrato = makeDict();
  const dictStatus = makeDict();
  const dictMes = makeDict();
  const dictEstab = makeDict();
  const dictContrato = makeDict();
  const dictCliente = makeDict();

  const rows = [];
  let processed = 0;
  let budgetsSeen = 0;

  await mapPool(clients, CONCURRENCY, async (client) => {
    const clIdx = dictCliente.idx(client.id);
    const gIdx = dictGenero.idx(client.gender);
    const mcIdx = dictMediaCliente.idx(client.mediaCliente);

    let budgets;
    try {
      budgets = await fetchAllPages(`/Budget/Client/List/${client.id}`, { Filters: [], Sort: [] });
    } catch (err) {
      console.warn(`Cliente ${client.id}: error listando presupuestos — ${err.message}`);
      budgets = [];
    }

    for (const b of budgets) {
      budgetsSeen++;
      const items = b.Items || b.items || [];
      if (!items.length) continue;

      const contratoKey = b.Id ?? b.Code ?? `${client.id}-${b.CreateDate}`;
      const cIdx = dictContrato.idx(contratoKey);

      const createDate = b.CreateDate ? new Date(b.CreateDate) : null;
      const mesKey = createDate
        ? `${createDate.getUTCFullYear()}-${String(createDate.getUTCMonth() + 1).padStart(2, '0')}`
        : 'SIN-FECHA';
      const mIdx = dictMes.idx(mesKey);

      const statusRaw = b.Status;
      const statusLabel = STATUS_LABELS[statusRaw] || statusRaw || 'Desconhecido';
      const stIdx = dictStatus.idx(statusLabel);
      const pVal = (statusRaw === 'paid' || statusRaw === 'transfered') ? 1 : 0;

      const estabLabel = (b.OwnerOrgStruct && (b.OwnerOrgStruct.Description ?? b.OwnerOrgStruct.description)) || 'SEM UNIDADE';
      const eIdx = dictEstab.idx(estabLabel);

      const mediaContratoLabel = (b.Media && (b.Media.Name ?? b.Media.name)) || 'SEM ORIGEM';
      const mtIdx = dictMediaContrato.idx(mediaContratoLabel);

      const qi = items.length;

      for (const it of items) {
        const svc = it.Item || it.item;
        const pkg = it.PackageItem || it.packageItem;
        const itemName = (svc && (svc.Name ?? svc.name)) || (pkg && (pkg.Description ?? pkg.description)) || 'ITEM SEM NOME';
        const itIdx = dictItem.idx(itemName.toUpperCase());
        const axVal = itemName.toUpperCase().includes('AXILA') ? 1 : 0;

        const salesman = it.Salesman || it.salesman;
        const vendedorName = (salesman && (salesman.Name ?? salesman.name)) || 'SEM VENDEDOR';
        const vIdx = dictVendedor.idx(vendedorName);

        const grossValue = it.TotalGrossValue ?? it.GrossValue ?? it.totalGrossValue ?? it.grossValue ?? 0;
        const netValue = it.NetValue ?? it.netValue ?? 0;

        rows.push([cIdx, clIdx, itIdx, axVal, vIdx, gIdx, mcIdx, mtIdx, qi, stIdx, grossValue, netValue, pVal, mIdx, eIdx]);
      }
    }

    processed++;
    if (processed % 100 === 0) console.log(`  ... ${processed}/${clients.length} clientes procesados`);
  });

  console.log(`Listo. Presupuestos vistos: ${budgetsSeen}. Líneas de ítem: ${rows.length}.`);

  const dicts = {
    contrato: dictContrato.size,
    cliente: dictCliente.size,
    item: dictItem.toArray(),
    vendedor: dictVendedor.toArray(),
    genero: dictGenero.toArray(),
    mediaCliente: dictMediaCliente.toArray(),
    mediaContrato: dictMediaContrato.toArray(),
    status: dictStatus.toArray(),
    mes: dictMes.toArray(),
    estabelecimento: dictEstab.toArray(),
  };

  const meta = {
    generatedAt: new Date().toISOString(),
    source: 'EVUP API',
    totalContratos: dictContrato.size,
    totalClientes: dictCliente.size,
  };

  await mkdir(OUT_DIR, { recursive: true });
  await writeFile(path.join(OUT_DIR, 'dicts.json'), JSON.stringify(dicts));
  await writeFile(path.join(OUT_DIR, 'rows.json'), JSON.stringify(rows));
  await writeFile(path.join(OUT_DIR, 'meta.json'), JSON.stringify(meta, null, 2));

  console.log(`Escrito: ${dictContrato.size} contratos, ${dictCliente.size} clientes, ${rows.length} líneas.`);
}

main().catch((err) => {
  console.error('Sync falló:', err);
  process.exit(1);
});
