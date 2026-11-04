#!/usr/bin/env node
/* Actualiza tasa.json con la tasa oficial del BCV.

   La tienda muestra los bolívares a partir de ese archivo. El BCV publica la
   tasa de cada día la tarde anterior y entra en vigor a las 12:00 am (hora de
   Venezuela). DolarAPI guarda en su histórico cada tasa con su «fecha valor»,
   así que aquí se toma la más reciente cuya fecha ya llegó: a las 12:30 am la
   tienda ya tiene la del día, sin depender de a qué hora la publique el BCV.

   En cada ejecución:
   · Consulta el histórico de DolarAPI y, de respaldo, su tasa vigente, con
     reintentos y tiempo límite.
   · Cambia tasa y fecha solo si hay una tasa más nueva y válida.
   · Anota en «verificado» cuándo se comprobó por última vez, una vez al día
     aunque la tasa no cambie, para que se vea que el proceso sigue vivo.
   · Si ninguna fuente responde, conserva la tasa anterior. Si lleva más de
     36 horas sin poder verificarla, termina con error: así GitHub avisa por
     correo al dueño del repositorio.

   Uso:  node scripts/actualizar-tasa.js
   Solo para pruebas (todas opcionales): TASA_ARCHIVO, TASA_URL_HISTORICO,
   TASA_URL_ACTUAL, TASA_AHORA (fecha ISO), TASA_FORZAR=1, TASA_ESPERA_MS,
   TASA_TIMEOUT_MS.                                                          */

const fs   = require('fs');
const path = require('path');

const ARCHIVO  = process.env.TASA_ARCHIVO || path.join(__dirname, '..', 'tasa.json');
const URL_HIST = process.env.TASA_URL_HISTORICO || 'https://ve.dolarapi.com/v1/historicos/dolares/oficial';
const URL_ACT  = process.env.TASA_URL_ACTUAL    || 'https://ve.dolarapi.com/v1/dolares/oficial';
const AHORA    = process.env.TASA_AHORA ? new Date(process.env.TASA_AHORA) : new Date();
const FORZAR   = process.env.TASA_FORZAR === '1';
const TIMEOUT  = Number(process.env.TASA_TIMEOUT_MS || 20000);
const ESPERA   = Number(process.env.TASA_ESPERA_MS  || 4000);   // pausa entre reintentos (se duplica)
const MAX_HORAS_SIN_VERIFICAR = 36;

// Venezuela está en UTC−4 todo el año (no tiene horario de verano).
const enVenezuela = (d) => new Date(d.getTime() - 4 * 3600 * 1000).toISOString();
const HOY        = enVenezuela(AHORA).slice(0, 10);                 // AAAA-MM-DD
const SELLO_AHORA = enVenezuela(AHORA).slice(0, 19) + '-04:00';
const FECHA_OK   = /^\d{4}-\d{2}-\d{2}$/;

const dormir = (ms) => new Promise((ok) => setTimeout(ok, ms));
const aviso  = (m) => console.log('::warning::' + m);

/* ── Consulta con reintentos ─────────────────────────────────────── */
async function pedirJSON(url) {
  let ultimo;
  for (let i = 0; i < 3; i++) {
    if (i) await dormir(ESPERA * i);
    try {
      const r = await fetch(url, {
        signal: AbortSignal.timeout(TIMEOUT),
        headers: { 'User-Agent': 'svr-venezuela-tasa/1.0 (+https://svr-venezuela.com)', 'Accept': 'application/json' }
      });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return await r.json();
    } catch (e) { ultimo = e; }
  }
  throw ultimo;
}

/* ── Cada fuente entrega { tasa, fecha } o lanza un error ─────────── */
function lectura(tasa, fecha) {
  tasa = Number(tasa);
  if (!isFinite(tasa) || tasa <= 0) throw new Error('tasa inválida (' + tasa + ')');
  if (!FECHA_OK.test(fecha))        throw new Error('fecha inválida (' + fecha + ')');
  return { tasa, fecha };
}

// La tasa que ya rige: la de fecha más reciente que no sea de mañana.
function desdeHistorico(lista) {
  if (!Array.isArray(lista)) throw new Error('el histórico no es una lista');
  let mejor = null;
  for (const e of lista) {
    if (!e || !FECHA_OK.test(String(e.fecha)) || e.fecha > HOY) continue;
    if (!mejor || e.fecha > mejor.fecha) mejor = e;
  }
  if (!mejor) throw new Error('el histórico no trae ninguna tasa vigente');
  return lectura(mejor.promedio, mejor.fecha);
}

function desdeActual(obj) {
  return lectura(obj && obj.promedio, String((obj && obj.fechaActualizacion) || '').slice(0, 10));
}

/* ── tasa.json actual (si falta o está dañado, se vuelve a crear) ──── */
function leerPrevio() {
  try {
    const p = JSON.parse(fs.readFileSync(ARCHIVO, 'utf8'));
    if (Number(p.tasa) > 0 && FECHA_OK.test(String(p.fecha))) return p;
    aviso('tasa.json no tiene una tasa válida: se vuelve a crear.');
  } catch (e) {
    if (e.code !== 'ENOENT') aviso('tasa.json no se pudo leer (' + e.message + '): se vuelve a crear.');
  }
  return null;
}

async function main() {
  const previo = leerPrevio();
  console.log(`Hoy en Venezuela: ${HOY} · tasa guardada: ${previo ? previo.tasa + ' (' + previo.fecha + ')' : 'ninguna'}`);

  const fuentes = [
    ['DolarAPI histórico', URL_HIST, desdeHistorico],
    ['DolarAPI actual',    URL_ACT,  desdeActual]
  ];
  const lecturas = [];
  for (const [nombre, url, leer] of fuentes) {
    try {
      const l = leer(await pedirJSON(url));
      lecturas.push({ nombre, ...l });
      console.log(`✓ ${nombre}: ${l.tasa} (fecha valor ${l.fecha})`);
    } catch (e) {
      aviso(`${nombre} no respondió bien: ${e.message}`);
    }
  }

  // Solo sirven las que ya rigen y son creíbles frente a la guardada (entre la mitad y el doble).
  const vigentes = lecturas.filter((l) => {
    if (l.fecha > HOY) return false;
    if (previo && (l.tasa < previo.tasa / 2 || l.tasa > previo.tasa * 2)) {
      aviso(`${l.nombre} dio ${l.tasa}, muy lejos de ${previo.tasa}: se descarta.`);
      return false;
    }
    return true;
  }).sort((a, b) => (a.fecha < b.fecha) - (a.fecha > b.fecha));     // la más reciente primero

  /* Ninguna fuente sirvió: se conserva lo anterior, salvo que ya sea muy viejo. */
  if (!vigentes.length) {
    const ref = previo && (previo.verificado || previo.actualizado);
    const horas = ref ? (AHORA - new Date(ref)) / 3600000 : Infinity;
    if (horas > MAX_HORAS_SIN_VERIFICAR) {
      console.log(`::error::No se pudo comprobar la tasa BCV en ${isFinite(horas) ? Math.floor(horas) + ' horas' : 'ninguna ocasión'}. ` +
                  'Revisa DolarAPI (https://ve.dolarapi.com) y la pestaña Actions del repositorio.');
      return 1;
    }
    aviso(`Ninguna fuente respondió. Se conserva ${previo.tasa} (${previo.fecha}); última comprobación hace ${horas.toFixed(1)} h.`);
    return 0;
  }

  const mejor = vigentes[0];
  const otra  = vigentes.find((l) => l !== mejor && l.fecha === mejor.fecha);
  if (otra && Math.abs(otra.tasa - mejor.tasa) / mejor.tasa > 0.005) {
    aviso(`Las fuentes no coinciden para ${mejor.fecha}: ${mejor.nombre} ${mejor.tasa} y ${otra.nombre} ${otra.tasa}. Se usa la primera.`);
  }

  // ¿Es más nueva que la guardada? (o la misma fecha con otro valor)
  const nueva = !previo || mejor.fecha > previo.fecha || (mejor.fecha === previo.fecha && mejor.tasa !== previo.tasa);
  const latido = FORZAR || !previo || !previo.verificado || String(previo.verificado).slice(0, 10) !== HOY;

  if (!nueva && !latido) {
    console.log(`Sin cambios: la tasa vigente sigue siendo ${previo.tasa} (${previo.fecha}) y ya se verificó hoy.`);
    return 0;
  }

  const salida = {
    moneda: 'VES',
    fuente: 'BCV',
    tasa:        nueva ? mejor.tasa : previo.tasa,
    fecha:       nueva ? mejor.fecha : previo.fecha,
    actualizado: nueva ? mejor.fecha + 'T00:00:00-04:00' : (previo.actualizado || previo.fecha + 'T00:00:00-04:00'),
    verificado:  SELLO_AHORA
  };
  fs.writeFileSync(ARCHIVO, JSON.stringify(salida, null, 2) + '\n');
  console.log(nueva
    ? `✓ Tasa nueva: ${salida.tasa} Bs/USD, fecha valor ${salida.fecha}.`
    : `✓ La tasa sigue en ${salida.tasa} (${salida.fecha}); se anota la verificación de hoy.`);
  return 0;
}

main().then((codigo) => process.exit(codigo)).catch((e) => { console.log('::error::' + (e && e.stack || e)); process.exit(1); });
