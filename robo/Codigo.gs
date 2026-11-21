// Robô do Mapa de Buenos Aires
// Lê a lista "Buenos Aires 2025 (V2)" do Google Maps e atualiza a aba "Lugares":
// lugar novo entra com o emoji ❓; nota mudada no Maps é atualizada;
// lugar que saiu da lista fica marcado na coluna "Na lista do Maps".
// Também puxa o quadro público do Yaay (ver atualizarDoYaay).
// Roda sozinho uma vez por semana (segunda, por volta das 9h).

const LISTA_MAPS = 'exgS7L6v6yaHQAPKbFAMfdUPdYuDuA';

function atualizarDoMaps() {
  const url = 'https://www.google.com/maps/preview/entitylist/getlist?authuser=0&hl=pt-BR&gl=br&pb=!1m4!1s' +
    LISTA_MAPS + '!2e1!3m1!1e1!2e2!3e2!4i500!16b1';
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('O Maps respondeu ' + resp.getResponseCode());
  const dados = JSON.parse(resp.getContentText().replace(/^\)\]\}'\n/, ''));
  const itens = (dados[0] && dados[0][8]) || [];
  if (!itens.length) throw new Error('A lista do Maps veio vazia; nada foi alterado.');

  const aba = SpreadsheetApp.getActive().getSheetByName('Lugares');
  const valores = aba.getDataRange().getValues();
  const cab = valores[0].map(String);
  let colStatus = cab.indexOf('Na lista do Maps');
  if (colStatus < 0) {
    colStatus = cab.length;
    aba.getRange(1, colStatus + 1).setValue('Na lista do Maps');
  }
  const col = nome => cab.indexOf(nome);
  const cId = col('ID'), cNota = col('Nota no Maps');

  const linhaPorId = {};
  for (let i = 1; i < valores.length; i++) {
    const id = String(valores[i][cId]).replace(/^'/, '');
    if (id) linhaPorId[id] = i + 1;
  }

  const vistos = {};
  const novas = [];
  itens.forEach(e => {
    const info = e[1] || [];
    const cid = BigInt.asUintN(64, BigInt(info[6][1])).toString();
    vistos[cid] = true;
    const nota = String(e[3] || '').trim();
    const linha = linhaPorId[cid];
    if (linha) {
      if (cNota >= 0 && String(valores[linha - 1][cNota]).trim() !== nota) aba.getRange(linha, cNota + 1).setValue(nota);
      return;
    }
    const coords = info[5] || [];
    const linhaNova = new Array(cab.length > colStatus ? cab.length : colStatus + 1).fill('');
    const poe = (nome, v) => { const c = col(nome); if (c >= 0) linhaNova[c] = v; };
    poe('Emoji', '❓');
    poe('Lugar', e[2] || '');
    poe('Nota no Maps', nota);
    poe('Endereço', String(info[4] || '').replace(/, Argentina$/, ''));
    poe('Latitude', coords[2]);
    poe('Longitude', coords[3]);
    poe('Link no Maps', 'https://maps.google.com/?cid=' + cid);
    poe('ID', "'" + cid);
    let t = Number(e[9] && e[9][0]) || 0;
    while (t > 1e13) t /= 1000;            // aceita segundos, ms ou µs
    const quando = t ? new Date(t < 1e11 ? t * 1000 : t) : new Date();
    poe('Salvo em', Utilities.formatDate(quando, 'America/Sao_Paulo', 'dd/MM/yyyy'));
    linhaNova[colStatus] = 'sim';
    novas.push(linhaNova);
  });

  // marca quem saiu da lista (não apaga)
  if (valores.length > 1) {
    const status = valores.slice(1).map(l => {
      const id = String(l[cId]).replace(/^'/, '');
      if (!/^\d+$/.test(id)) return [l[colStatus] === undefined ? '' : l[colStatus]]; // lugares do Yaay ficam como estão
      return [vistos[id] ? 'sim' : 'não'];
    });
    aba.getRange(2, colStatus + 1, status.length, 1).setValues(status);
  }

  if (novas.length) aba.getRange(ultimaLinhaComDados(aba) + 1, 1, novas.length, novas[0].length).setValues(novas);
  Logger.log(novas.length + ' lugar(es) novo(s); ' + itens.length + ' na lista do Maps.');
}

// ───────────────────────── Yaay ─────────────────────────
// Lê o quadro público "Argentina" do Yaay (sem login) e acrescenta os lugares
// que ainda não estão na planilha. Não mexe em nada que já existe.

const YAAY_PERFIL = '019d5d7e-6e98-7453-b05a-0bd427c36d21';
const YAAY_QUADRO = '019f23cf-e7d3-7b00-a9b8-42eaff63a7ac';

const YAAY_EMOJI = {pizza:'🍕',bbq:'🥩',steakhouse:'🥩',argentinian_restaurant:'🍽️',restaurant:'🍽️',fine_dining:'🍽️',
  brunch:'☕',buffet:'🍽️',vegetarian:'🍽️',cafe:'☕',bakery:'🥐',dessert_and_sweets:'🥐',ice_cream:'🍦',burger:'🥪',
  hot_dog:'🥪',fast_food:'🥪',deli:'🥪',kebab:'🌍',sushi:'🌍',italian_restaurant:'🍝',pasta:'🍝',spanish_restaurant:'🌍',
  pub:'🍸',bar:'🍸',wine_bar:'🍸',plaza:'🏛️',park:'🏛️',river:'🏛️',attraction:'🏛️',landmark:'🏛️',historical_site:'🏛️',
  stadium:'🏛️',tour:'🏛️',destination:'🏛️',experience:'🏛️',shopping_mall:'🛒'};
const YAAY_TIPO = {pizza:'Pizza',bbq:'Churrascaria',steakhouse:'Churrascaria',argentinian_restaurant:'Argentina',
  restaurant:'Restaurante',fine_dining:'Jantar sofisticado',brunch:'Brunch',buffet:'Bufê',vegetarian:'Vegetariana',
  cafe:'Café',bakery:'Padaria',dessert_and_sweets:'Doces',ice_cream:'Sorvete',burger:'Hambúrguer',hot_dog:'Cachorro-quente',
  fast_food:'Fast Food',deli:'Rotisseria',kebab:'Shawarma',sushi:'Japonesa',italian_restaurant:'Italiana',pasta:'Massas',
  spanish_restaurant:'Espanhola',pub:'Cervejaria',bar:'Bar',wine_bar:'Bar de vinhos',plaza:'Praça',park:'Parque',
  river:'Orla',attraction:'Atração',landmark:'Marco',historical_site:'Sítio histórico',stadium:'Estádio',tour:'Passeio',
  destination:'Passeio',experience:'Experiência',shopping_mall:'Shopping'};
const YAAY_IGNORAR = {hotel:1, car_rental:1};
const YAAY_GENERICO = {destination:1, itinerary:1, hot_spot:1, other:1, argentinian_restaurant:1, restaurant:1};

function atualizarDoYaay() {
  const q = 'query Q($by: PublicUserBy!) { user { public(by: $by) { boards { id groups { items { savedId createdAt ' +
    'savedData { item { externalUrl createdAt } externalInfo { authorNickname } matches { title subCategory ' +
    'description hidden latitude longitude googleMapsId address } } } } } } } }';
  const resp = UrlFetchApp.fetch('https://yaaytravel.com/graphql', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ query: q, variables: { by: { id: YAAY_PERFIL } } })
  });
  if (resp.getResponseCode() !== 200) throw new Error('O Yaay respondeu ' + resp.getResponseCode());
  const quadro = (JSON.parse(resp.getContentText()).data.user.public.boards || []).filter(b => b.id === YAAY_QUADRO)[0];
  if (!quadro) throw new Error('Não achei o quadro do Yaay (ele continua público?).');

  const aba = SpreadsheetApp.getActive().getSheetByName('Lugares');
  const valores = aba.getDataRange().getValues();
  const cab = valores[0].map(String);
  const col = nome => cab.indexOf(nome);
  const num = v => Number(String(v).replace(',', '.'));
  const perto = (a, b, c, d, m) => Math.hypot((a - c) * 111000, (b - d) * 92000) < m;

  const existentes = valores.slice(1).map(l => ({ id: String(l[col('ID')]).replace(/^'/, ''), lat: num(l[col('Latitude')]), lng: num(l[col('Longitude')]) }))
    .filter(e => e.id || e.lat);
  const abaIgnorar = SpreadsheetApp.getActive().getSheetByName('Ignorar');   // IDs que você apagou e não quer de volta
  const ignorar = {};
  if (abaIgnorar) abaIgnorar.getDataRange().getValues().slice(1).forEach(l => { if (l[0]) ignorar[String(l[0]).trim()] = 1; });
  const novas = [];
  const data = s => { const m = String(s || '').match(/(\d{4})-(\d\d)-(\d\d)/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : ''; };
  const limpa = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

  quadro.groups.forEach(g => (g.items || []).forEach(it => {
    const sd = it.savedData; if (!sd) return;
    (sd.matches || []).forEach(m => {
      const lat = Number(m.latitude), lng = Number(m.longitude), sc = m.subCategory || '';
      if (m.hidden || !m.googleMapsId || !lat || ignorar[m.googleMapsId]) return;
      if (!(lat > -35.0 && lat < -34.3 && lng > -58.95 && lng < -58.15)) return;      // fora da Grande Buenos Aires
      if (YAAY_IGNORAR[sc]) return;
      const rua = limpa(m.address).split(',')[0].replace(/^[A-Z]\d{4}[A-Z]{0,3}\b/, '');
      if (YAAY_GENERICO[sc] && !/\d/.test(rua) && !/Parque|Plaza|Costanera|Feria|Reserva|Puerto/i.test(m.title)) return; // bairro/cidade genérico
      if (existentes.some(e => e.id === m.googleMapsId || (e.lat && perto(lat, lng, e.lat, e.lng, 60)))) return;
      existentes.push({ id: m.googleMapsId, lat: lat, lng: lng });

      const linha = new Array(cab.length).fill('');
      const poe = (nome, v) => { const c = col(nome); if (c >= 0) linha[c] = v; };
      poe('Emoji', YAAY_EMOJI[sc] || '❓');
      poe('Lugar', limpa(m.title));
      poe('Tipo (Google)', YAAY_TIPO[sc] || sc);
      poe('Endereço', limpa(m.address).replace(/, Argentina$/, '').replace('Cdad. Autónoma de Buenos Aires', 'CABA'));
      poe('Latitude', lat);
      poe('Longitude', lng);
      poe('Link no Maps', 'https://www.google.com/maps/place/?q=place_id:' + m.googleMapsId);
      poe('ID', m.googleMapsId);
      poe('Salvo em', data((sd.item && sd.item.createdAt) || it.createdAt));
      poe('Origem', 'Yaay');
      poe('Post do Instagram', String((sd.item && sd.item.externalUrl) || '').split('?')[0]);
      const autor = sd.externalInfo && sd.externalInfo.authorNickname;
      poe('Resumo do Yaay', limpa(m.description) + (autor ? ' (via @' + autor + ')' : ''));
      novas.push(linha);
    });
  }));

  if (novas.length) {
    const primeira = ultimaLinhaComDados(aba) + 1;
    aba.getRange(primeira, col('ID') + 1, novas.length, 1).setNumberFormat('@');   // ID sempre como texto
    aba.getRange(primeira, 1, novas.length, cab.length).setValues(novas);
  }
  normalizarNumeros(aba);
  Logger.log(novas.length + ' lugar(es) novo(s) vindos do Yaay.');
}

// Latitude/Longitude precisam ser números (o site lê a planilha e ignora coordenadas gravadas como texto).
function normalizarNumeros(aba) {
  const valores = aba.getDataRange().getValues();
  const cab = valores[0].map(String);
  const cols = [cab.indexOf('Latitude'), cab.indexOf('Longitude')].filter(c => c >= 0);
  cols.forEach(c => {
    let mudou = false;
    const coluna = valores.slice(1).map(l => {
      const v = l[c];
      if (typeof v === 'string' && v.trim() && !isNaN(Number(v.replace(',', '.')))) { mudou = true; return [Number(v.replace(',', '.'))]; }
      return [v];
    });
    if (mudou) {
      const r = aba.getRange(2, c + 1, coluna.length, 1);
      r.setNumberFormat('0.000000');
      r.setValues(coluna);
    }
  });
  const cData = cab.indexOf('Salvo em');
  if (cData >= 0) {
    let mudou = false;
    const coluna = valores.slice(1).map(l => {
      const m = typeof l[cData] === 'string' && l[cData].match(/^(\d\d)\/(\d\d)\/(\d{4})$/);
      if (m) { mudou = true; return [new Date(+m[3], +m[2] - 1, +m[1])]; }
      return [l[cData]];
    });
    if (mudou) {
      const r = aba.getRange(2, cData + 1, coluna.length, 1);
      r.setNumberFormat('dd/MM/yyyy');
      r.setValues(coluna);
    }
  }
}

function ultimaLinhaComDados(aba) {
  const b = aba.getRange(1, 2, aba.getMaxRows(), 1).getValues();   // coluna "Lugar"
  for (let i = b.length - 1; i >= 0; i--) if (String(b[i][0]).trim()) return i + 1;
  return 1;
}

// O que roda toda semana: Maps e depois Yaay (um erro em um não impede o outro).
function atualizarTudo() {
  const erros = [];
  try { atualizarDoMaps(); } catch (e) { erros.push('Maps: ' + e.message); }
  try { atualizarDoYaay(); } catch (e) { erros.push('Yaay: ' + e.message); }
  if (erros.length) throw new Error(erros.join(' | '));
}

// Rode esta uma única vez: cria o agendamento semanal (segunda, ~9h) e já faz a primeira rodada.
function instalarAgendamento() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('atualizarTudo').timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  atualizarTudo();
}
