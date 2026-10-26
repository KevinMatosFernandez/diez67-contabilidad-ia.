// Pruebas del bot contable. No usan Telegram, Gemini, elTOQUE ni Supabase reales:
// fetch se sustituye por una base de datos en memoria y respuestas simuladas.
// Ejecutar: deno test --allow-env --allow-read index.test.ts

import { assert, assertEquals, assertMatch, assertStringIncludes } from "jsr:@std/assert@1";
import * as XLSX from "npm:xlsx@0.18.5";

// deno-lint-ignore no-explicit-any
type Row = Record<string, any>;

const OWNER = 42;
const ADMIN = 77;
const INACTIVE = 55;
const SECRET = "secreto-webhook";
const CRON_TOKEN = "T".repeat(64);

for (const [k, v] of Object.entries({
  TELEGRAM_BOT_TOKEN: "t",
  GEMINI_API_KEY: "g",
  TELEGRAM_OWNER_ID: String(OWNER),
  PROJECT_URL: "https://db.test",
  DATABASE_API_KEY: "k",
  ELTOQUE_API_KEY: "e",
  TELEGRAM_WEBHOOK_SECRET: SECRET,
})) Deno.env.set(k, v);

/* ----- Base de datos en memoria ----- */

const ACCOUNT_CODES: [number, string, string][] = [
  [1, "1101", "activo"], [2, "1102", "activo"], [19, "1103", "activo"], [4, "1301", "activo"],
  [6, "2101", "pasivo"], [7, "2102", "pasivo"], [9, "3101", "patrimonio"], [10, "3201", "patrimonio"],
  [11, "4101", "ingreso"], [13, "5101", "gasto"], [14, "6101", "gasto"], [15, "6102", "gasto"],
  [16, "6103", "gasto"], [17, "6104", "gasto"], [18, "6105", "gasto"],
  [3, "1201", "activo"], [5, "1501", "activo"], [12, "4201", "ingreso"], [20, "1190", "activo"], [21, "1502", "activo"],
  [31, "6113", "gasto"], [8, "2201", "pasivo"], [35, "1104", "activo"], [36, "1202", "activo"], [37, "3202", "patrimonio"], [38, "2202", "pasivo"], [32, "6114", "gasto"], [33, "4204", "ingreso"], [34, "6115", "gasto"], [22, "4202", "ingreso"], [23, "4203", "ingreso"], [24, "6106", "gasto"], [25, "6107", "gasto"], [26, "6108", "gasto"],
  [27, "6109", "gasto"], [28, "6110", "gasto"], [29, "6111", "gasto"], [30, "6112", "gasto"], [39, "6116", "gasto"],
];

const db: Record<string, Row[]> = {};
let nextId = 1;

function resetDb() {
  for (const t of [
    "operaciones_pendientes", "asientos_contables", "lineas_asiento", "transacciones", "cache_respuestas_ia",
    "historial_tasas", "periodos_cerrados", "auditoria", "obligaciones", "registros_excel_diario",
    "pagos_obligaciones", "activos_fijos", "actividad_bot", "tasas_proveedor", "cierres_recaudacion", "gastos_fijos", "provisiones",
  ]) db[t] = [];
  db.usuarios = [
    { id: 1, telegram_id: OWNER, nombre: "Christian", rol: "dueno", activo: true },
    { id: 2, telegram_id: ADMIN, nombre: "Diana", rol: "admin", activo: true },
    { id: 3, telegram_id: INACTIVE, nombre: "Antiguo", rol: "admin", activo: false },
  ];
  db.plan_cuentas = ACCOUNT_CODES.map(([id, codigo, tipo]) => ({ id, codigo, tipo, nombre: `Cuenta ${codigo}`, activa: true }));
  nextId = 100;
}

function matches(row: Row, key: string, filter: string): boolean {
  const value = key.startsWith("propuesta->>") ? row.propuesta?.[key.slice(12)] : row[key];
  const dot = filter.indexOf(".");
  const op = filter.slice(0, dot);
  const arg = filter.slice(dot + 1);
  const v = String(value);
  switch (op) {
    case "eq": return v === arg;
    case "neq": return v !== arg;
    case "lt": return v < arg;
    case "lte": return v <= arg;
    case "gte": return v >= arg;
    case "in": return arg.slice(1, -1).split(",").includes(v);
    default: throw new Error(`Filtro no simulado: ${op}`);
  }
}

/* ----- Servicios externos simulados ----- */

const sent: Row[] = [];
let geminiCalls = 0;
let geminiReply = "";
// Respuesta de revaluar_cuentas_cup (reajuste cambiario); null = sin ajuste.
let revaluation: Row | null = null;
// Resultado del mes en USD que devuelve estado_resultados (para /utilidad).
let estadoResultado = 0;
// deno-lint-ignore no-explicit-any
let lastGeminiBody: any = null;
// deno-lint-ignore no-explicit-any
const geminiBodies: any[] = [];
// true = Gemini responde 400 a cualquier petición con responseSchema (para probar el reintento sin esquema).
let geminiRejectsSchema = false;

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}

function balances(moneda: string): Row[] {
  const out = new Map<number, Row>();
  for (const l of db.lineas_asiento) {
    const a = db.asientos_contables.find((x) => x.id === l.asiento_id);
    if (!a || a.estado === "anulado" || (moneda !== "FUNCIONAL" && l.moneda !== moneda)) continue;
    const acc = db.plan_cuentas.find((c) => c.id === l.cuenta_id)!;
    const row = out.get(l.cuenta_id) ?? { cuenta_id: acc.id, codigo: acc.codigo, nombre: acc.nombre, tipo: acc.tipo, debe: 0, haber: 0 };
    // Moneda funcional: el contravalor en USD de cada línea (las de CUP, a la TRMI simulada de 480).
    const usd = (native: number, stored: unknown) => stored != null ? Number(stored) : l.moneda === "USD" ? native : native / 480;
    row.debe += moneda === "FUNCIONAL" ? usd(Number(l.debe), l.debe_usd) : Number(l.debe);
    row.haber += moneda === "FUNCIONAL" ? usd(Number(l.haber), l.haber_usd) : Number(l.haber);
    out.set(l.cuenta_id, row);
  }
  return [...out.values()];
}

function rpc(fn: string, b: Row): Response {
  switch (fn) {
    case "confirmar_operacion_pendiente_con_tasa": {
      const op = db.operaciones_pendientes.find((o) => o.id === b.p_operacion_id)!;
      if (db.periodos_cerrados.some((p) => p.periodo === `${op.fecha.slice(0, 7)}-01`)) {
        return json({ code: "P0001", message: `El mes ${op.fecha.slice(0, 7)} está cerrado.` }, 400);
      }
      const asientoId = nextId++;
      db.asientos_contables.push({ id: asientoId, fecha: op.fecha, estado: "borrador", descripcion: op.propuesta.descripcion, creado_en: new Date().toISOString() });
      for (const l of op.propuesta.lineas) {
        db.lineas_asiento.push({ id: nextId++, asiento_id: asientoId, cuenta_id: l.cuenta_id, debe: l.debe, haber: l.haber, moneda: l.moneda });
      }
      const obl = op.propuesta.obligacion;
      if (obl?.accion === "crear") {
        db.obligaciones.push({ id: nextId++, fecha: op.fecha, acreedor: obl.contraparte, moneda: obl.moneda, importe_original: Number(obl.importe), estado: "pendiente", sentido: obl.sentido, asiento_id: asientoId });
      }
      if (obl?.accion === "pagar") {
        for (const p of obl.pagos) db.pagos_obligaciones.push({ id: nextId++, obligacion_id: p.obligacion_id, importe: Number(p.importe), anulado: false, asiento_id: asientoId });
      }
      const act = op.propuesta.activo;
      if (act) db.activos_fijos.push({ id: nextId++, nombre: act.nombre, fecha_compra: op.fecha, moneda: act.moneda, costo: Number(act.costo), vida_util_meses: act.vida_util_meses, estado: "activo", asiento_id: asientoId });
      // Igual que el trigger aplicar_provisiones_operacion de la migración.
      for (const a of op.propuesta.provisiones ?? []) {
        const p = db.provisiones.find((x) => x.id === a.id)!;
        if (a.accion === "activar" && p.estado === "pendiente") p.estado = "activa";
        if (a.accion === "reservar") p.reservado_cup = Number(p.reservado_cup ?? 0) + a.importe;
        if (a.accion === "pagar") {
          p.pagado_cup = Number(p.pagado_cup ?? 0) + a.importe;
          p.reservado_cup = Math.max(Number(p.reservado_cup ?? 0) - (a.de_reserva ?? 0), 0);
          if (p.pagado_cup >= p.importe_cup - 0.01) p.estado = "pagada";
        }
      }
      op.estado = "confirmada";
      op.asiento_id = asientoId;
      op.tasa_usada = b.p_tasa;
      return json({ ok: true, asiento_id: asientoId });
    }
    case "saldos_obligaciones": return json(
      db.obligaciones.filter((o) => (o.sentido ?? "pagar") === b.p_sentido && ["pendiente", "parcial"].includes(o.estado)).map((o) => {
        const pagado = db.pagos_obligaciones.filter((p) => p.obligacion_id === o.id && !p.anulado).reduce((s, p) => s + p.importe, 0);
        return { id: o.id, fecha: o.fecha, contraparte: o.acreedor, moneda: o.moneda, importe_original: o.importe_original, pagado, saldo: o.importe_original - pagado, estado: o.estado };
      }),
    );
    case "saldos_cuentas": return json(balances(b.p_moneda));
    case "resumen_asientos": return json({ borrador: db.asientos_contables.filter((a) => a.estado === "borrador").length, validado: 0 });
    case "indicadores_financieros": return json({ hay_datos: false });
    case "verificar_token_cron": return json(b.p_token === CRON_TOKEN);
    case "cerrar_periodo": {
      const ym = String(b.p_mes).slice(0, 7);
      const drafts = db.asientos_contables.filter((a) => a.estado === "borrador" && a.fecha.startsWith(ym)).length;
      if (drafts) return json({ code: "P0001", message: `No se puede cerrar ${ym}: hay ${drafts} asiento(s) en borrador.` }, 400);
      db.periodos_cerrados.push({ periodo: b.p_mes, cerrado_en: new Date().toISOString() });
      return json({ ok: true, periodo: ym, asientos_validados: 0 });
    }
    case "inventario_preparar_conteo": {
      const stock: Record<string, [number, number, number]> = { moyoc: [7, 801, 0.48], "jabon kare": [8, 14, 0.27], ajo: [2, 59, 0.15] };
      const items: Row[] = [];
      const missing: string[] = [];
      let faltante = 0;
      let sobrante = 0;
      for (const it of b.p_items) {
        const key = String(it.producto).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        const s = stock[key];
        if (!s) { missing.push(it.producto); continue; }
        const dif = it.cantidad - s[1];
        if (dif < 0) faltante += -dif * s[2];
        if (dif > 0) sobrante += dif * s[2];
        items.push({ producto_id: s[0], producto: it.producto, sistema: s[1], contado: it.cantidad, diferencia: dif, costo_unitario_usd: s[2], valor_usd: dif * s[2] });
      }
      return json({ items, no_encontrados: missing, faltante_usd: faltante, sobrante_usd: sobrante });
    }
    case "inventario_conciliacion": return json({ fecha: b.p_fecha, contabilidad_usd: 858.37, productos_usd: 861.61, diferencia_usd: -3.24, cuadra: false });
    case "inventario_existencias": return json([
      { producto_id: 7, producto: "Moyoc", ubicacion_id: 1, ubicacion: "Principal", cantidad: 798, costo_promedio_usd: 0.48, valor_usd: 383.04 },
      { producto_id: 2, producto: "Ajo", ubicacion_id: 1, ubicacion: "Principal", cantidad: -1, costo_promedio_usd: null, valor_usd: -0.15 },
    ]);
    case "inventario_revisar_excel": return json({ filas: 1, movimientos: 3, sin_costo: [], diferencias_inicio: [{ producto: "Cerveza", sistema: 71, excel: 70 }], diferencias_final: [] });
    case "inventario_costo_excel": return json({ costo_ventas_usd: 7.74, mermas_usd: 0.38, consumo_usd: 0.2, movimientos: 3, filas: 1 });
    case "estado_resultados": return json({ resultado: estadoResultado });
    case "revaluar_cuentas_cup": return json(revaluation ?? { asiento_id: null, ajuste_usd: 0, ya_existia: false, cuentas: [] });
    default: throw new Error(`RPC no simulada: ${fn}`);
  }
}

let telegramFile = new Uint8Array([1, 2, 3, 4]);

globalThis.fetch = async (input: string | URL | Request, init: RequestInit = {}): Promise<Response> => {
  const url = String(input instanceof Request ? input.url : input);
  const method = init.method ?? "GET";

  if (url.startsWith("https://api.telegram.org/file/")) return new Response(telegramFile);
  if (url.startsWith("https://api.telegram.org")) {
    if (url.endsWith("/getFile")) return json({ ok: true, result: { file_path: "documentos/archivo" } });
    // sendDocument llega como multipart (FormData); el resto, como JSON.
    const body = init.body instanceof FormData ? Object.fromEntries(init.body.entries()) : JSON.parse(String(init.body ?? "{}"));
    sent.push({ method: url.split("/").pop(), ...body });
    return json({ ok: true, result: {} });
  }
  if (url.startsWith("https://generativelanguage.googleapis.com")) {
    geminiCalls++;
    lastGeminiBody = JSON.parse(String(init.body ?? "{}"));
    geminiBodies.push(lastGeminiBody);
    if (geminiRejectsSchema && lastGeminiBody.generationConfig?.responseSchema) return json({ error: { message: "Invalid schema" } }, 400);
    return json({ candidates: [{ content: { parts: [{ text: geminiReply }] } }] });
  }
  if (url.startsWith("https://tasas.eltoque.com")) return json({ tasas: { USD: { mediana: 480 } } });

  const u = new URL(url);
  const path = u.pathname.replace("/rest/v1/", "");
  if (path.startsWith("rpc/")) return rpc(path.slice(4), JSON.parse(String(init.body ?? "{}")));

  const table = db[path];
  if (!table) throw new Error(`Tabla no simulada: ${path}`);
  const filters = [...u.searchParams].filter(([k]) => !["select", "order", "limit", "and", "on_conflict"].includes(k));
  const and = u.searchParams.get("and");
  if (and) {
    for (const part of and.slice(1, -1).split(",")) {
      const [col, op, ...rest] = part.split(".");
      filters.push([col, `${op}.${rest.join(".")}`]);
    }
  }
  const rows = table.filter((r) => filters.every(([k, v]) => matches(r, k, v)));

  if (method === "GET") return json(rows);
  if (method === "POST") {
    const body = JSON.parse(String(init.body));
    const list = (Array.isArray(body) ? body : [body]).map((r: Row) => ({ id: nextId++, estado: "pendiente", creado_en: new Date().toISOString(), ...(path === "provisiones" ? { reservado_cup: 0, pagado_cup: 0 } : {}), ...r }));
    table.push(...list);
    return json(list, 201);
  }
  if (method === "PATCH") {
    const body = JSON.parse(String(init.body));
    rows.forEach((r) => Object.assign(r, body));
    return json(rows);
  }
  throw new Error(`Método no simulado: ${method}`);
};

/* ----- Carga del bot sin abrir un servidor ----- */

let handler: (req: Request) => Promise<Response>;
Object.defineProperty(Deno, "serve", {
  configurable: true,
  writable: true,
  value: (h: typeof handler) => {
    handler = h;
    return {};
  },
});
await import("./index.ts");

let updateId = 1;

async function say(text: string, from = OWNER): Promise<string> {
  sent.length = 0;
  await handler(new Request("https://bot.test", {
    method: "POST",
    headers: { "x-telegram-bot-api-secret-token": SECRET },
    body: JSON.stringify({ update_id: updateId++, message: { date: Math.floor(Date.now() / 1000), chat: { id: from, type: "private" }, from: { id: from }, text } }),
  }));
  return sent.map((m) => m.text ?? "").join("\n");
}

async function press(data: string, from = OWNER): Promise<Row[]> {
  sent.length = 0;
  await handler(new Request("https://bot.test", {
    method: "POST",
    headers: { "x-telegram-bot-api-secret-token": SECRET },
    body: JSON.stringify({ update_id: updateId++, callback_query: { id: "cb", from: { id: from }, data, message: { message_id: 7, chat: { id: from, type: "private" } } } }),
  }));
  return [...sent];
}

async function cron(token: string, tarea: string): Promise<Response> {
  sent.length = 0;
  return await handler(new Request("https://bot.test", {
    method: "POST",
    headers: { "x-diez67-cron": token },
    body: JSON.stringify({ tarea }),
  }));
}

function todayHavana(): string {
  return new Intl.DateTimeFormat("en-CA", { timeZone: "America/Havana" }).format(new Date());
}

function test(name: string, fn: () => Promise<void>) {
  Deno.test(name, async () => {
    resetDb();
    geminiCalls = 0;
    geminiReply = "";
    revaluation = null;
    estadoResultado = 0;
    lastGeminiBody = null;
    geminiBodies.length = 0;
    geminiRejectsSchema = false;
    await fn();
  });
}

/* ----- Seguridad ----- */

test("rechaza peticiones sin el secreto del webhook", async () => {
  const res = await handler(new Request("https://bot.test", { method: "POST", body: "{}" }));
  assertEquals(res.status, 401);
});

test("ignora mensajes de otros usuarios", async () => {
  assertEquals(await say("Pagué 500 CUP de electricidad en efectivo", 99), "");
  assertEquals(db.operaciones_pendientes.length, 0);
});

test("un botón pulsado por otra persona no hace nada", async () => {
  const msgs = await press("confirmar:1", 99);
  assertEquals(msgs.map((m) => m.text), ["No autorizado."]);
});

test("las tareas programadas exigen el token de Vault", async () => {
  assertEquals((await cron("token-falso", "alertas_diarias")).status, 401);
  assertEquals(sent.length, 0);
});

/* ----- Varios usuarios ----- */

test("un usuario desactivado no puede usar el bot", async () => {
  assertEquals(await say("/pendientes", INACTIVE), "");
});

test("la admin propone, el dueño ve quién fue y puede confirmarla", async () => {
  await say("Pagué 500 CUP de electricidad en efectivo", ADMIN);
  const op = db.operaciones_pendientes[0];
  assertEquals(op.usuario_telegram, String(ADMIN));
  assertStringIncludes(await say("/pendientes"), "de Diana");
  assertStringIncludes(await say(`/confirmar ${op.id}`), "borrador");
  assertEquals(op.estado, "confirmada");
});

test("la admin puede confirmar y cancelar operaciones del dueño con botones", async () => {
  await say("Pagué 500 CUP de electricidad en efectivo");
  await say("Pagué 300 CUP de agua en efectivo");
  const [a, b] = db.operaciones_pendientes;
  await press(`confirmar:${a.id}`, ADMIN);
  await press(`cancelar:${b.id}`, ADMIN);
  assertEquals([a.estado, b.estado], ["confirmada", "cancelada"]);
});

test("se registra lo que hace la admin y no lo que hace el dueño", async () => {
  await say("/pendientes");
  await say("Pagué 500 CUP de electricidad en efectivo", ADMIN);
  await press(`confirmar:${db.operaciones_pendientes[0].id}`, ADMIN);
  assertEquals(db.actividad_bot.map((r) => [r.telegram_id, r.tipo]), [[ADMIN, "mensaje"], [ADMIN, "boton"]]);
  assertEquals(db.actividad_bot[0].nombre, "Diana");
});

test("el reporte diario le cuenta al dueño lo que hizo la admin", async () => {
  await say("Pagué 500 CUP de electricidad en efectivo", ADMIN);
  const res = await cron(CRON_TOKEN, "reporte_diario_usuarios");
  assertEquals(res.status, 200);
  assertEquals(sent.length, 1);
  assertEquals(sent[0].chat_id, OWNER);
  assertStringIncludes(sent[0].text, "Diana");
  assertStringIncludes(sent[0].text, "electricidad");
});

test("sin actividad de otros usuarios no se envía reporte diario", async () => {
  await say("/pendientes");
  const res = await cron(CRON_TOKEN, "reporte_diario_usuarios");
  assertEquals(res.status, 200);
  assertEquals(sent.length, 0);
});

/* ----- Propuestas de asiento ----- */

test("un gasto claro se propone por reglas, sin gastar Gemini", async () => {
  const reply = await say("Pagué 500 CUP de electricidad en efectivo");
  assertEquals(geminiCalls, 0);
  assertStringIncludes(reply, "DEBE  | 6102");
  assertStringIncludes(reply, "HABER | 1101");
  assertStringIncludes(reply, "500.00 CUP");
  const op = db.operaciones_pendientes[0];
  assertEquals(op.propuesta.lineas.length, 2);
  assertEquals(op.propuesta.lineas[0].debe, op.propuesta.lineas[1].haber);
});

test("un curso pagado va a 6116 Gastos de capacitación, por reglas", async () => {
  const reply = await say("Pagué 5000 CUP de cursos de contabilidad por Transfermóvil");
  assertEquals(geminiCalls, 0);
  assertStringIncludes(reply, "DEBE  | 6116");
  assertStringIncludes(reply, "HABER | 1103");
});

test("Gemini puede clasificar un gasto como capacitación", async () => {
  geminiReply = JSON.stringify({ intencion: "operacion", tipo: "gasto", categoria_gasto: "capacitacion", monto_texto: "3000", moneda: "CUP", descripcion: "taller de ventas", fecha_operacion: todayHavana(), menciona_otra_fecha: false, varias_operaciones: false, pregunta_aclaratoria: null });
  const reply = await say("Le pagué 3000 CUP a Diana por el taller de ventas en efectivo");
  assertStringIncludes(reply, "DEBE  | 6116");
});

test("el clasificador pide a Gemini el esquema JSON con los tipos y categorías del bot", async () => {
  geminiReply = JSON.stringify({ intencion: "operacion", tipo: "gasto", categoria_gasto: "capacitacion", monto_texto: "3000", moneda: "CUP", descripcion: "taller de ventas", fecha_operacion: todayHavana(), menciona_otra_fecha: false, varias_operaciones: false, pregunta_aclaratoria: null });
  await say("Le pagué 3000 CUP a Diana por el taller de ventas en efectivo");
  const config = lastGeminiBody.generationConfig;
  assertEquals(config.responseMimeType, "application/json");
  assertEquals(config.responseSchema.properties.tipo.enum.includes("compra_activo"), true);
  assertEquals(config.responseSchema.properties.categoria_gasto.enum.includes("capacitacion"), true);
  assertEquals(config.responseSchema.required.includes("intencion"), true);
});

test("si Gemini rechaza el esquema, se repite sin él y la operación sale igual", async () => {
  geminiRejectsSchema = true;
  geminiReply = JSON.stringify({ intencion: "operacion", tipo: "gasto", categoria_gasto: "capacitacion", monto_texto: "3000", moneda: "CUP", descripcion: "taller de ventas", fecha_operacion: todayHavana(), menciona_otra_fecha: false, varias_operaciones: false, pregunta_aclaratoria: null });
  const reply = await say("Le pagué 3000 CUP a Diana por el taller de ventas en efectivo");
  assertEquals(geminiCalls, 2);
  assertEquals(geminiBodies[1].generationConfig.responseSchema, undefined);
  assertEquals(geminiBodies[1].generationConfig.responseMimeType, "application/json");
  assertStringIncludes(reply, "DEBE  | 6116");
});

test("lo que menciona la planta o los pallets lo decide Gemini, no las reglas", async () => {
  geminiReply = JSON.stringify({ intencion: "operacion", tipo: "gasto", categoria_gasto: "electricidad_servicios", monto_texto: "27000", moneda: "CUP", descripcion: "aceite de la planta", fecha_operacion: todayHavana(), menciona_otra_fecha: false, varias_operaciones: false, pregunta_aclaratoria: null });
  const reply = await say("Pagué 27000 CUP de aceite de la planta en efectivo");
  assertEquals(geminiCalls, 1);
  assertStringIncludes(reply, "DEBE  | 6102");
});

test("la propuesta trae botones de confirmar y cancelar", async () => {
  await say("Pagué 500 CUP de electricidad en efectivo");
  const id = db.operaciones_pendientes[0].id;
  const buttons = sent.at(-1)!.reply_markup.inline_keyboard[0].map((b: Row) => b.callback_data);
  assertEquals(buttons, [`confirmar:${id}`, `cancelar:${id}`]);
});

test("no mezcla monedas en una misma operación", async () => {
  geminiReply = JSON.stringify({ intencion: "operacion", tipo: "gasto", categoria_gasto: "electricidad_servicios", monto_texto: "500", moneda: "CUP", descripcion: "luz", fecha_operacion: todayHavana(), menciona_otra_fecha: false, varias_operaciones: false, pregunta_aclaratoria: null });
  const reply = await say("Pagué 500 CUP y 20 USD de luz en efectivo");
  assertStringIncludes(reply, "No generé ninguna propuesta");
  assertEquals(db.operaciones_pendientes.length, 0);
});

test("rechaza operaciones con fecha futura", async () => {
  const [y, m, d] = todayHavana().split("-").map(Number);
  const tomorrow = new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10);
  geminiReply = JSON.stringify({ intencion: "operacion", tipo: "gasto", categoria_gasto: "electricidad_servicios", monto_texto: "500", moneda: "CUP", descripcion: "luz", fecha_operacion: tomorrow, menciona_otra_fecha: true, varias_operaciones: false, pregunta_aclaratoria: null });
  const reply = await say("Pagué 500 CUP de electricidad en efectivo mañana");
  assertStringIncludes(reply, "posterior a hoy");
  assertEquals(db.operaciones_pendientes.length, 0);
});

/* ----- Ciclo del asiento ----- */

test("confirmar crea un asiento en borrador y ofrece validarlo", async () => {
  await say("Pagué 500 CUP de electricidad en efectivo");
  const id = db.operaciones_pendientes[0].id;
  const msgs = await press(`confirmar:${id}`);
  const asiento = db.asientos_contables[0];
  assertEquals(asiento.estado, "borrador");
  const last = msgs.at(-1)!;
  assertStringIncludes(last.text, `Operación #${id} confirmada`);
  assertEquals(last.reply_markup.inline_keyboard[0][0].callback_data, `validar:${asiento.id}`);
});

test("pulsar confirmar dos veces no duplica el asiento", async () => {
  await say("Pagué 500 CUP de electricidad en efectivo");
  const id = db.operaciones_pendientes[0].id;
  await press(`confirmar:${id}`);
  const msgs = await press(`confirmar:${id}`);
  assertEquals(db.asientos_contables.length, 1);
  assertStringIncludes(msgs.at(-1)!.text, "ya figura como confirmada");
});

test("no valida un asiento descuadrado", async () => {
  db.asientos_contables.push({ id: 1, fecha: todayHavana(), estado: "borrador", descripcion: "x" });
  db.lineas_asiento.push(
    { id: 2, asiento_id: 1, cuenta_id: 15, debe: 500, haber: 0, moneda: "CUP" },
    { id: 3, asiento_id: 1, cuenta_id: 1, debe: 0, haber: 400, moneda: "CUP" },
  );
  const reply = await say("/validar 1");
  assertStringIncludes(reply, "descuadrado");
  assertEquals(db.asientos_contables[0].estado, "borrador");
});

/* ----- Cierre mensual ----- */

test("no acepta operaciones con fecha de un mes cerrado", async () => {
  db.periodos_cerrados.push({ periodo: `${todayHavana().slice(0, 7)}-01` });
  const reply = await say("Pagué 500 CUP de electricidad en efectivo");
  assertMatch(reply, /El mes \d{4}-\d{2} está cerrado/);
  assertEquals(db.operaciones_pendientes.length, 0);
});

test("no cierra un mes con borradores", async () => {
  db.asientos_contables.push({ id: 1, fecha: "2026-09-10", estado: "borrador", descripcion: "x" });
  const reply = await say("/cerrar 2026-09");
  assertStringIncludes(reply, "No se cerró el mes 2026-09");
  assertStringIncludes(reply, "borrador");
  assertEquals(db.periodos_cerrados.length, 0);
});

/* ----- Alertas y tareas programadas ----- */

test("sin alertas, la tarea diaria no escribe al usuario", async () => {
  const res = await cron(CRON_TOKEN, "alertas_diarias");
  assertEquals(res.status, 200);
  assertEquals(sent.length, 0);
});

test("avisa de borradores atrasados", async () => {
  db.asientos_contables.push({ id: 5, fecha: "2026-01-02", estado: "borrador", creado_en: "2026-01-02T10:00:00Z" });
  const res = await cron(CRON_TOKEN, "alertas_diarias");
  assertEquals(res.status, 200);
  assertEquals(sent.length, 1);
  assertStringIncludes(sent[0].text, "#5");
});

test("rechaza tareas programadas desconocidas", async () => {
  assertEquals((await cron(CRON_TOKEN, "borrar_todo")).status, 400);
});

/* ----- Fotos y PDF ----- */

async function sendFile(kind: "photo" | "pdf", caption?: string): Promise<string> {
  sent.length = 0;
  const media = kind === "photo"
    ? { photo: [{ file_id: "chica", file_size: 100 }, { file_id: "grande", file_size: 4000 }] }
    : { document: { file_id: "pdf1", file_name: "factura.pdf", mime_type: "application/pdf", file_size: 4000 } };
  await handler(new Request("https://bot.test", {
    method: "POST",
    headers: { "x-telegram-bot-api-secret-token": SECRET },
    body: JSON.stringify({ update_id: updateId++, message: { date: Math.floor(Date.now() / 1000), chat: { id: OWNER, type: "private" }, from: { id: OWNER }, ...media, ...(caption ? { caption } : {}) } }),
  }));
  return sent.map((m) => m.text ?? "").join("\n");
}

function reading(overrides: Row = {}): string {
  return JSON.stringify({
    es_comprobante: true, tipo: "gasto", categoria_gasto: "electricidad_servicios", concepto: "Factura de electricidad",
    proveedor: "UNE", fecha: null, moneda: "CUP", total: 500, medio_pago: "efectivo",
    renglones: [{ descripcion: "Consumo octubre", importe: 500 }], dudas: null, ...overrides,
  });
}

test("una foto de factura se convierte en propuesta con botones, con una sola petición a Gemini", async () => {
  geminiReply = reading();
  const reply = await sendFile("photo");
  assertStringIncludes(reply, "Esto leí en la foto");
  assertStringIncludes(reply, "Pagué 500 CUP de electricidad en efectivo");
  assertStringIncludes(reply, "HABER | 1101");
  assertEquals(db.operaciones_pendientes.length, 1);
  assertEquals(geminiCalls, 1);
  assertEquals(lastGeminiBody.contents[0].parts[0].inline_data.mime_type, "image/jpeg");
  assertEquals(lastGeminiBody.generationConfig.responseSchema.properties.renglones.type, "ARRAY");
  const id = db.operaciones_pendientes[0].id;
  assertEquals(sent.at(-1)!.reply_markup.inline_keyboard[0][0].callback_data, `confirmar:${id}`);
});

test("la misma foto reenviada no gasta otra petición de Gemini", async () => {
  geminiReply = reading();
  await sendFile("photo");
  await sendFile("photo");
  assertEquals(geminiCalls, 1);
});

test("el pie de foto completa el medio de pago que no trae el documento", async () => {
  geminiReply = reading({ medio_pago: null });
  const reply = await sendFile("pdf", "pagado por transfermóvil");
  assertStringIncludes(reply, "Esto leí en el PDF");
  assertStringIncludes(reply, "HABER | 1103");
  assertEquals(lastGeminiBody.contents[0].parts[0].inline_data.mime_type, "application/pdf");
});

test("un PDF sin moneda clara no genera propuesta", async () => {
  geminiReply = reading({ moneda: null });
  const reply = await sendFile("pdf");
  assertStringIncludes(reply, "la moneda (CUP o USD)");
  assertEquals(db.operaciones_pendientes.length, 0);
});

test("una foto que no es un comprobante no genera propuesta", async () => {
  geminiReply = reading({ es_comprobante: false, tipo: "otro" });
  const reply = await sendFile("photo");
  assertStringIncludes(reply, "No parece una factura");
  assertEquals(db.operaciones_pendientes.length, 0);
});

test("si Gemini no responde, avisa y no registra nada", async () => {
  geminiReply = "esto no es json";
  const reply = await sendFile("photo");
  assertStringIncludes(reply, "No pude leer la foto");
  assertEquals(db.operaciones_pendientes.length, 0);
});

/* ----- Informes en Excel ----- */

async function readWorkbook(doc: Row) {
  assertEquals(doc.method, "sendDocument");
  assert(doc.document instanceof File);
  return XLSX.read(new Uint8Array(await doc.document.arrayBuffer()), { type: "array" });
}

function seedSaleThisMonth() {
  const fecha = `${todayHavana().slice(0, 7)}-01`;
  db.asientos_contables.push({ id: 900, fecha, estado: "validado", descripcion: "Venta del día", creado_en: new Date().toISOString() });
  db.lineas_asiento.push(
    { id: 901, asiento_id: 900, cuenta_id: 1, debe: 1500, haber: 0, moneda: "CUP", debe_usd: 3.125, haber_usd: 0 },
    { id: 902, asiento_id: 900, cuenta_id: 11, debe: 0, haber: 1500, moneda: "CUP", debe_usd: 0, haber_usd: 3.125 },
  );
}

test("/excel envía el informe del mes en texto y en Excel, sin usar la IA", async () => {
  seedSaleThisMonth();
  await say("/excel");
  const doc = sent.find((m) => m.method === "sendDocument")!;
  assert(sent.findIndex((m) => m.method === "sendMessage") < sent.indexOf(doc), "el texto llega antes que el archivo");
  assertStringIncludes(sent.map((m) => m.text ?? "").join("\n"), "ESTADO DE RESULTADOS");
  assertMatch(doc.document.name, /^DIEZ67 informe \d{4}-\d{2}.*\.xlsx$/);
  const wb = await readWorkbook(doc);
  assertEquals(wb.SheetNames, ["Explicación", "Resultados", "Balance", "Saldos", "Asientos"]);
  const resultados = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets["Resultados"], { header: 1 });
  const total = resultados.find((r) => r[1] === "RESULTADO DEL PERÍODO")!;
  assertEquals([total[0], total[4]], ["USD", 3.13]);
  const asientos = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets["Asientos"], { header: 1 });
  assertEquals(asientos.filter((r) => r[1] === 900).length, 2);
  assertEquals(geminiCalls, 0);
});

test("/excel con un mes mal escrito no envía archivo", async () => {
  const reply = await say("/excel octubre");
  assertStringIncludes(reply, "AAAA-MM");
  assertEquals(sent.filter((m) => m.method === "sendDocument").length, 0);
});

test("el resumen semanal programado llega con su Excel", async () => {
  const res = await cron(CRON_TOKEN, "resumen_semanal");
  assertEquals(res.status, 200);
  assertEquals((await res.json()).excel, true);
  assertStringIncludes(sent[0].text, "RESUMEN SEMANAL");
  const wb = await readWorkbook(sent[sent.length - 1]);
  const explicacion = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets["Explicación"], { header: 1 });
  assertStringIncludes(String(explicacion[0][0]), "Resumen semanal DIEZ67");
});

test("el informe mensual programado es del mes anterior", async () => {
  const res = await cron(CRON_TOKEN, "informe_mensual");
  assertEquals(res.status, 200);
  const [y, m] = todayHavana().split("-").map(Number);
  const prev = m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, "0")}`;
  assertStringIncludes(sent[0].text, `INFORME MENSUAL DIEZ67 ${prev}`);
  assertEquals(sent[sent.length - 1].document.name, `DIEZ67 informe ${prev}.xlsx`);
});

/* ----- Caché de la IA ----- */

test("una pregunta repetida se responde desde la caché", async () => {
  db.registros_excel_diario.push({ fecha_inicio: "2026-10-01", fecha_fin: "2026-10-01", producto: "Cerveza", venta: 40, importe: 12000, importe_costo: 6000, costo: 150 });
  geminiReply = "Análisis de prueba";
  const first = await say("analiza los productos del 1 de octubre");
  const second = await say("analiza los productos del 1 de octubre");
  assertEquals(geminiCalls, 1);
  assertEquals(first, second);
  assert(db.cache_respuestas_ia.length === 1);
});

test("las respuestas de la IA llegan sin asteriscos, almohadillas ni tablas", async () => {
  db.registros_excel_diario.push({ fecha_inicio: "2026-10-01", fecha_fin: "2026-10-01", producto: "Cerveza", venta: 40, importe: 12000, importe_costo: 6000, costo: 150 });
  geminiReply = "## Resumen\n\n**Cerveza** dejó *6.000 CUP*.\n\n| Producto | Ganancia |\n|---|---|\n| Cerveza | 6000 |\n\n* Revisar `costo`";
  const answer = await say("analiza los productos del 1 de octubre");
  assert(!/[*#`|]/.test(answer), answer);
  assertStringIncludes(answer, "Resumen\n\nCerveza dejó 6.000 CUP.");
  assertStringIncludes(answer, "• Cerveza · 6000");
  assertStringIncludes(answer, "• Revisar costo");
});

/* ----- Operaciones nuevas (v17) ----- */

function intent(extra: Row): string {
  return JSON.stringify({
    intencion: "operacion", categoria_gasto: null, moneda: "CUP", descripcion: null, fecha_operacion: todayHavana(),
    menciona_otra_fecha: false, varias_operaciones: false, pregunta_aclaratoria: null, contraparte: null, a_credito: false,
    costo_texto: null, vida_util_anios: null, nombre_activo: null, subtipo_saldo: null, ...extra,
  });
}

function lines(op: Row): string[] {
  return op.propuesta.lineas.map((l: Row) => `${l.cuenta_id}:${l.debe}/${l.haber} ${l.moneda}`);
}

test("entiende montos abreviados como 590k", async () => {
  const reply = await say("Compré mercancía por 590k CUP en efectivo");
  assertEquals(geminiCalls, 0);
  assertStringIncludes(reply, "Interpreté «590k» como 590000");
  assertEquals(lines(db.operaciones_pendientes[0]), ["4:590000/0 CUP", "1:0/590000 CUP"]);
});

test("varias operaciones en líneas separadas dan varias propuestas con sus botones", async () => {
  await say("Pagué 500 CUP de electricidad en efectivo\nPagué 45000 CUP de salarios en efectivo");
  assertEquals(geminiCalls, 0);
  assertEquals(db.operaciones_pendientes.length, 2);
  const rows = sent.at(-1)!.reply_markup.inline_keyboard;
  assertEquals(rows.length, 2);
  assertEquals(rows[1][0].callback_data, `confirmar:${db.operaciones_pendientes[1].id}`);
});

test("varias operaciones en una frase se separan con la IA", async () => {
  geminiReply = JSON.stringify({ ...JSON.parse(intent({ tipo: "gasto", monto_texto: "5000" })), varias_operaciones: true, operaciones: ["Pagué 5000 CUP de luz en efectivo", "Pagué 2000 CUP de agua en efectivo"] });
  const reply = await say("Pagué 5000 CUP de luz y 2000 CUP de agua en efectivo");
  assertStringIncludes(reply, "Separé tu mensaje en 2 operaciones");
  assertEquals(geminiBodies.at(-1).generationConfig.responseSchema.properties.operaciones.type, "ARRAY");
  assertEquals(db.operaciones_pendientes.length, 2);
});

test("una compra a crédito crea la deuda y el pago la descuenta", async () => {
  db.tasas_proveedor.push({ proveedor: "Pérez", proveedor_norm: "perez", sobretasa_cup: 0 });
  geminiReply = intent({ tipo: "compra_inventario", monto_texto: "300000", contraparte: "Pérez", a_credito: true });
  await say("Compré mercancía a crédito a Pérez por 300000 CUP");
  const compra = db.operaciones_pendientes[0];
  assertEquals(lines(compra), ["4:300000/0 CUP", "6:0/300000 CUP"]);
  await press(`confirmar:${compra.id}`);
  assertEquals(db.obligaciones.length, 1);

  geminiReply = intent({ tipo: "pago_proveedor", monto_texto: "100000", contraparte: "Pérez" });
  const reply = await say("Le pagué 100000 CUP a Pérez en efectivo");
  assertStringIncludes(reply, "Saldo pendiente después de este pago: 200000 CUP");
  const pago = db.operaciones_pendientes[1];
  assertEquals(pago.propuesta.obligacion.pagos, [{ obligacion_id: db.obligaciones[0].id, importe: "100000" }]);
  await press(`confirmar:${pago.id}`);
  assertStringIncludes(await say("/deudas"), "Pérez: 200.000,00 CUP");
});

test("no deja pagar a un proveedor sin deuda registrada", async () => {
  geminiReply = intent({ tipo: "pago_proveedor", monto_texto: "150000", contraparte: "Pérez" });
  const reply = await say("Pagué 150000 CUP al proveedor Pérez en efectivo");
  assertStringIncludes(reply, "No tengo deudas abiertas");
  assertStringIncludes(reply, "Saldo inicial: debo 150000 CUP a Pérez");
  assertEquals(db.operaciones_pendientes.length, 0);
});

test("una venta con costo descuenta el inventario", async () => {
  geminiReply = intent({ tipo: "venta", monto_texto: "60000", costo_texto: "40000" });
  await say("Vendí 10 cajas de cerveza por 60000 CUP en efectivo, costo 40000 CUP");
  assertEquals(lines(db.operaciones_pendientes[0]), ["1:60000/0 CUP", "11:0/60000 CUP", "13:40000/0 CUP", "4:0/40000 CUP"]);
});

test("una venta a crédito crea la cuenta por cobrar y el cobro la salda", async () => {
  geminiReply = intent({ tipo: "venta", monto_texto: "50000", contraparte: "Juan", a_credito: true });
  await say("Vendí a crédito a Juan 50000 CUP");
  const venta = db.operaciones_pendientes[0];
  assertEquals(lines(venta).slice(0, 2), ["3:50000/0 CUP", "11:0/50000 CUP"]);
  await press(`confirmar:${venta.id}`);
  geminiReply = intent({ tipo: "cobro_cliente", monto_texto: "50000", contraparte: "Juan" });
  await say("Cobré 50000 CUP a Juan por Transfermóvil");
  assertEquals(lines(db.operaciones_pendientes[1]), ["19:50000/0 CUP", "3:0/50000 CUP"]);
});

test("un cambio de moneda cuadra cada moneda y calcula la diferencia cambiaria sin IA", async () => {
  const reply = await say("Cambié 100 USD por 52000 CUP en efectivo");
  assertEquals(geminiCalls, 0);
  const op = db.operaciones_pendientes[0];
  assertEquals(op.propuesta.moneda, "MIXTA");
  assertEquals(lines(op), ["20:100/0 USD", "2:0/100 USD", "1:52000/0 CUP", "20:0/48000 CUP", "22:0/4000 CUP"]);
  assertStringIncludes(reply, "Ganancia cambiaria: 4000 CUP");
  await press(`confirmar:${op.id}`);
  assertEquals(op.tasa_usada, 480);
});

test("vender dólares a una tasa calcula los pesos", async () => {
  await say("Vendí 100 USD a 470");
  assertEquals(lines(db.operaciones_pendientes[0]).slice(2), ["1:47000/0 CUP", "20:0/48000 CUP", "25:1000/0 CUP"]);
});

test("saldos iniciales de deudas y equipos por reglas", async () => {
  await say("Saldo inicial: debo 200000 CUP a Pérez");
  const deuda = db.operaciones_pendientes[0];
  assertEquals(lines(deuda), ["9:200000/0 CUP", "6:0/200000 CUP"]);
  assertEquals(deuda.propuesta.obligacion.contraparte, "Pérez");
  await say("Saldo inicial: nevera 1500 USD");
  const equipo = db.operaciones_pendientes[1];
  assertEquals(lines(equipo), ["5:1500/0 USD", "9:0/1500 USD"]);
  assertEquals(equipo.propuesta.activo.vida_util_meses, 60);
  assertEquals(geminiCalls, 0);
});

test("la depreciación mensual empieza el mes siguiente a la compra", async () => {
  db.activos_fijos.push({ id: 1, nombre: "Nevera", fecha_compra: "2026-01-15", moneda: "USD", costo: 1500, vida_util_meses: 60, estado: "activo" });
  assertStringIncludes(await say("/depreciar 2026-01"), "No hay activos que se deprecien");
  await say("/depreciar 2026-02");
  assertEquals(lines(db.operaciones_pendientes[0]), ["24:25/0 USD", "21:0/25 USD"]);
  assertStringIncludes(await say("/depreciar 2026-02"), "ya tiene la operación");
});

test("el conteo físico ajusta el inventario", async () => {
  await say("Compré mercancía por 590000 CUP en efectivo");
  await press(`confirmar:${db.operaciones_pendientes[0].id}`);
  await say("/conteo 500000");
  assertEquals(lines(db.operaciones_pendientes[1]), ["26:90000/0 CUP", "4:0/90000 CUP"]);
});

test("mermas, nuevas categorías de gasto y retiros por Transfermóvil por reglas", async () => {
  await say("Se rompieron mercancías por 4500 CUP");
  await say("Pagué 2000 CUP de internet en efectivo");
  await say("Retiré 3000 CUP por Transfermóvil para uso personal");
  assertEquals(geminiCalls, 0);
  assertEquals(db.operaciones_pendientes.map(lines), [
    ["26:4500/0 CUP", "4:0/4500 CUP"],
    ["28:2000/0 CUP", "1:0/2000 CUP"],
    ["10:3000/0 CUP", "19:0/3000 CUP"],
  ]);
});

/* ----- Inventario por producto (v19) ----- */

test("/contar propone el ajuste por producto con sus movimientos", async () => {
  const reply = await say("/contar Moyoc 798; Jabón Kare 15\nAjo 58");
  assertStringIncludes(reply, "Moyoc: sistema 801, contado 798 (faltan 3, 1,44 USD)");
  assertStringIncludes(reply, "Jabón Kare: sistema 14, contado 15 (sobran 1, 0,27 USD)");
  const op = db.operaciones_pendientes[0];
  assertEquals(lines(op), ["26:1.59/0 USD", "4:0/1.59 USD", "4:0.27/0 USD", "23:0/0.27 USD"]);
  assertEquals(op.propuesta.inventario.map((m: Row) => `${m.producto_id}:${m.tipo}:${m.cantidad}`), ["7:ajuste_conteo:-3", "8:ajuste_conteo:1", "2:ajuste_conteo:-1"]);
});

test("/contar no inventa productos que no están en el catálogo", async () => {
  const reply = await say("/contar Cerveza Imaginaria 10");
  assertStringIncludes(reply, "No encontré en el catálogo: Cerveza Imaginaria");
  assertEquals(db.operaciones_pendientes.length, 0);
});

test("/contar sin cantidad explica el formato", async () => {
  const reply = await say("/contar Moyoc");
  assertStringIncludes(reply, "Formato: /contar");
});

test("/ventas contabiliza también roturas y consumo del Excel", async () => {
  db.registros_excel_diario.push({ fecha_inicio: "2026-10-01", fecha_fin: "2026-10-01", producto: "Cerveza", venta: 40, importe: 12000, importe_costo: 6000, costo: 150, rotura: 2, consumo: 1 });
  const reply = await say("/ventas 2026-10-01");
  // Venta en CUP; costo y mermas en USD desde el inventario por producto.
  assertEquals(lines(db.operaciones_pendientes[0]), ["1:12000/0 CUP", "11:0/12000 CUP", "13:7.74/0 USD", "26:0.38/0 USD", "36:0.20/0 USD", "4:0/8.32 USD"]);
  assertEquals(db.operaciones_pendientes[0].propuesta.moneda, "MIXTA");
  assertStringIncludes(reply, "0,38 USD");
  assertStringIncludes(reply, "0,20 USD");
});

test("/existencias avisa de existencias negativas y /conciliar explica la diferencia", async () => {
  const stock = await say("/existencias");
  assertStringIncludes(stock, "Moyoc: 798 u × 0,48 = 383,04 USD");
  assertStringIncludes(stock, "Con existencia negativa");
  const conc = await say("/conciliar");
  assertStringIncludes(conc, "Los productos valen 3,24 USD más que la contabilidad");
});

/* ----- Moneda funcional USD ----- */

test("/conciliar ajustar pasa una diferencia pequeña al costo de ventas", async () => {
  const reply = await say("/conciliar ajustar");
  assertStringIncludes(reply, "Diferencia de precio de compra");
  assertEquals(lines(db.operaciones_pendientes[0]), ["4:3.24/0 USD", "13:0/3.24 USD"]);
});

test("/reajuste informa la pérdida cambiaria de las cuentas en CUP", async () => {
  revaluation = { asiento_id: 77, ajuste_usd: -2.95, ya_existia: false, cuentas: [{ codigo: "1101", nombre: "Caja CUP", ajuste_usd: -2.59 }, { codigo: "1103", nombre: "Transfermóvil", ajuste_usd: -0.36 }] };
  const reply = await say("/reajuste 2026-10-02");
  assertStringIncludes(reply, "Reajuste del 2026-10-02 a la TRMI 480");
  assertStringIncludes(reply, "Pérdida por devaluación del CUP: 2,95 USD");
  assertStringIncludes(await say("/reajuste 2030-01-01"), "Formato: /reajuste");
});

test("la tarea diaria avisa del reajuste solo si hubo ajuste", async () => {
  revaluation = { asiento_id: 78, ajuste_usd: -1.2, ya_existia: false, cuentas: [] };
  await cron(CRON_TOKEN, "alertas_diarias");
  assertEquals(sent.length, 1);
  assertStringIncludes(sent[0].text, "Pérdida por devaluación del CUP: 1,20 USD");
});

test("los informes salen por defecto en USD, la moneda funcional", async () => {
  seedSaleThisMonth();
  const reply = await say("/resultados");
  assertStringIncludes(reply, "USD (moneda funcional)");
  assertStringIncludes(reply, "RESULTADO DEL PERÍODO: 3,13 USD");
  assertStringIncludes(await say("/resultados CUP"), "RESULTADO DEL PERÍODO: 1500,00 CUP");
});

/* ----- Compras en CUP con precio en USD (sobreprecio aparte) ----- */

test("una compra en CUP con su precio en USD separa el sobreprecio", async () => {
  const reply = await say("Compré mercancía por 54000 CUP (100 USD) en efectivo");
  const op = db.operaciones_pendientes[0];
  assertEquals(op.propuesta.moneda, "MIXTA");
  assertEquals(lines(op), ["4:100/0 USD", "20:0/100 USD", "20:48000/0 CUP", "31:6000/0 CUP", "1:0/54000 CUP"]);
  assertStringIncludes(reply, "Sobreprecio por pagar en CUP: 6000");
});

test("a crédito, la deuda queda en CUP y se avisa de la ganancia al pagar más tarde", async () => {
  geminiReply = intent({ tipo: "compra_inventario", monto_texto: "54000", contraparte: "Pérez", a_credito: true });
  const reply = await say("Compré mercancía a crédito a Pérez por 54000 CUP (equivalente a 100 USD)");
  const op = db.operaciones_pendientes[0];
  assertEquals(lines(op), ["4:100/0 USD", "20:0/100 USD", "20:48000/0 CUP", "31:6000/0 CUP", "6:0/54000 CUP"]);
  assertEquals([op.propuesta.obligacion.moneda, op.propuesta.obligacion.importe], ["CUP", "54000"]);
  assertStringIncludes(reply, "reajuste diario");
});

test("si el precio en CUP queda por debajo de la TRMI, el ahorro se abona a 6113", async () => {
  await say("Compré mercancía por 47000 CUP (100 USD) por Transfermóvil");
  assertEquals(lines(db.operaciones_pendientes[0]), ["4:100/0 USD", "20:0/100 USD", "20:48000/0 CUP", "31:0/1000 CUP", "19:0/47000 CUP"]);
});

test("el precio en USD solo vale para compras de mercancía", async () => {
  const reply = await say("Pagué 5000 CUP de electricidad en efectivo (10 USD)");
  assertStringIncludes(reply, "solo los uso en compras de mercancía en CUP");
  assertEquals(db.operaciones_pendientes.length, 0);
});

test("la regla del proveedor (TRMI + N) calcula el precio en USD sin escribirlo", async () => {
  db.tasas_proveedor.push({ proveedor: "Pérez", proveedor_norm: "perez", sobretasa_cup: 40 });
  geminiReply = intent({ tipo: "compra_inventario", monto_texto: "52000", contraparte: "Pérez", a_credito: true });
  const reply = await say("Compré mercancía a crédito a Pérez por 52000 CUP");
  assertEquals(lines(db.operaciones_pendientes[0]), ["4:100/0 USD", "20:0/100 USD", "20:48000/0 CUP", "31:4000/0 CUP", "6:0/52000 CUP"]);
  assertStringIncludes(reply, "Tasa de Pérez: TRMI 480 + 40 = 520");
});

test("un proveedor sin tasa guardada: el bot la pide y /tasaproveedor la guarda", async () => {
  geminiReply = intent({ tipo: "compra_inventario", monto_texto: "52000", contraparte: "Gómez", a_credito: true });
  assertStringIncludes(await say("Compré mercancía a crédito a Gómez por 52000 CUP"), "No tengo la tasa de Gómez");
  assertEquals(db.operaciones_pendientes.length, 0);
  assertStringIncludes(await say("/tasaproveedor Gómez +40"), "Gómez cobra en CUP a la TRMI + 40");
  assertEquals([db.tasas_proveedor[0].proveedor_norm, db.tasas_proveedor[0].sobretasa_cup], ["gomez", 40]);
  assertStringIncludes(await say("/tasaproveedor"), "Gómez: TRMI + 40");
});

test("la tasa escrita en el mensaje manda sobre la regla", async () => {
  await say("Compré mercancía por 54000 CUP a tasa 540 en efectivo");
  assertEquals(lines(db.operaciones_pendientes[0]), ["4:100/0 USD", "20:0/100 USD", "20:48000/0 CUP", "31:6000/0 CUP", "1:0/54000 CUP"]);
});

test("una compra sin proveedor entra a la TRMI y lo avisa", async () => {
  const reply = await say("Compré mercancía por 5000 CUP en efectivo");
  assertEquals(lines(db.operaciones_pendientes[0]), ["4:5000/0 CUP", "1:0/5000 CUP"]);
  assertStringIncludes(reply, "Sin proveedor");
});

/* ----- Cierre del turno ----- */

function seedShift() {
  db.registros_excel_diario.push({ fecha_inicio: "2026-10-01", fecha_fin: "2026-10-01", producto: "Cerveza", venta: 40, importe: 12000, importe_costo: 6000, costo: 150, tasa_ref: 500, archivo_origen: "Ventas 1 octubre 2026.xlsx" });
}

test("/cierre reparte la venta por medio de cobro, descuenta el salario y guarda la caja", async () => {
  seedShift();
  const reply = await say("/cierre 2026-10-01 efectivo 6000 transferencia 4000 usd 4");
  const op = db.operaciones_pendientes[0];
  assertEquals(op.propuesta.moneda, "MIXTA");
  assertEquals(lines(op), [
    "1:6000/0 CUP", "19:4000/0 CUP", "11:0/10000 CUP", "2:4/0 USD", "11:0/4 USD",
    "13:7.74/0 USD", "26:0.38/0 USD", "36:0.20/0 USD", "4:0/8.32 USD",
  ]);
  assertEquals(lines(db.operaciones_pendientes[1]), ["34:300/0 CUP", "1:0/240 CUP", "19:0/60 CUP"]);
  assertStringIncludes(reply, "✅ Cuadra con el Excel.");
  assertStringIncludes(reply, "Comisión del vendedor: 300,00 CUP (240,00 en efectivo y 60,00 por Transfermóvil)");
  const c = db.cierres_recaudacion[0];
  assertEquals([c.fecha, c.efectivo_cup, c.transfermovil_cup, c.efectivo_usd, c.tasa_turno, c.venta_total_cup, c.diferencia_cup], ["2026-10-01", 6000, 4000, 4, 500, 12000, 0]);
  assertStringIncludes(await say("/ventas 2026-10-01"), "ya tienen la operación");
});

test("/cierre lleva la diferencia con el Excel a faltante de caja", async () => {
  seedShift();
  const reply = await say("/cierre 2026-10-01 efectivo 7000 transferencia 4000");
  assertEquals(lines(db.operaciones_pendientes[0]).slice(0, 4), ["1:7000/0 CUP", "19:4000/0 CUP", "32:1000/0 CUP", "11:0/12000 CUP"]);
  assertStringIncludes(reply, "Faltan 1000,00 CUP");
});

test("responder al mensaje de cierre del turno hace el cierre", async () => {
  seedShift();
  sent.length = 0;
  await handler(new Request("https://bot.test", {
    method: "POST",
    headers: { "x-telegram-bot-api-secret-token": SECRET },
    body: JSON.stringify({ update_id: updateId++, message: {
      date: Math.floor(Date.now() / 1000), chat: { id: OWNER, type: "private" }, from: { id: OWNER },
      text: "efectivo 8000 transferencia 4000 usd 0",
      reply_to_message: { from: { id: 1, is_bot: true }, text: "🧾 Cierre del turno 2026-10-01 a 2026-10-01\nVenta total del Excel: 12000,00 CUP" },
    } }),
  }));
  assertStringIncludes(sent.map((m) => m.text).join("\n"), "Cierre del turno 2026-10-01 a 2026-10-01");
  assertEquals(db.cierres_recaudacion.length, 1);
  assertEquals(geminiCalls, 0);
});

async function replyTo(promptText: string, text: string) {
  sent.length = 0;
  await handler(new Request("https://bot.test", {
    method: "POST",
    headers: { "x-telegram-bot-api-secret-token": SECRET },
    body: JSON.stringify({ update_id: updateId++, message: {
      date: Math.floor(Date.now() / 1000), chat: { id: OWNER, type: "private" }, from: { id: OWNER },
      text, reply_to_message: { from: { id: 1, is_bot: true }, text: promptText },
    } }),
  }));
  return sent[sent.length - 1].text as string;
}

test("cuestionario del cierre: una pregunta por medio de cobro y luego la propuesta", async () => {
  seedShift();
  const q1 = "🧾 Cierre del turno 2026-10-01 a 2026-10-01\nVenta total del Excel: 12000,00 CUP.\nTexto.\n\n❓ Pregunta 1 de 3: ¿Cuánto EFECTIVO en CUP se cobró en el turno?\nResponde a este mensaje solo con el número (0 si no hubo).";
  const q2 = await replyTo(q1, "8000");
  assertStringIncludes(q2, "✓ Efectivo CUP: 8000");
  assertStringIncludes(q2, "Pregunta 2 de 3");
  assertMatch(q2, /^🧾 Cierre del turno 2026-10-01 a 2026-10-01/);
  const again = await replyTo(q2, "cuatro mil");
  assertStringIncludes(again, "Pregunta 2 de 3");
  assertEquals(db.cierres_recaudacion.length, 0);
  const bad = await replyTo(q2, "4.000,5,5");
  assertStringIncludes(bad, "No entendí");
  assertMatch(bad, /^🧾 Cierre del turno/);
  const q3 = await replyTo(bad, "4000");
  assertStringIncludes(q3, "✓ Transfermóvil CUP: 4000");
  assertStringIncludes(q3, "Pregunta 3 de 3");
  const done = await replyTo(q3, "0");
  assertStringIncludes(done, "Respuestas del cierre: efectivo 8000 transferencia 4000 usd 0.");
  const c = db.cierres_recaudacion[0];
  assertEquals([c.efectivo_cup, c.transfermovil_cup, c.efectivo_usd], [8000, 4000, 0]);
  assertEquals(geminiCalls, 0);
});

test("/cierre con el número delante de cada palabra, en varias líneas", async () => {
  seedShift();
  await say("/cierre 2026-10-01 6000 efectivo \n4000 transferencia \n4 Usd");
  const c = db.cierres_recaudacion[0];
  assertEquals([c.efectivo_cup, c.transfermovil_cup, c.efectivo_usd], [6000, 4000, 4]);
});

test("/cierre sin importes explica el formato", async () => {
  seedShift();
  assertStringIncludes(await say("/cierre 2026-10-01 hola"), "No encontré importes");
});

/* ----- Gastos fijos y utilidad del dueño ----- */

function seedProvision(extra: Row = {}): Row {
  const p = {
    id: nextId++, tipo: "gasto_fijo", gasto_fijo_id: 1, beneficiario: "Diana Fernández", concepto: "Salario quincenal", cuenta_pasivo_id: 8,
    importe_origen: 70, moneda_origen: "USD", tasa_fijada: 500, importe_cup: 35000, fecha_inicio: "2026-10-01", fecha_pago: "2026-10-15",
    dias: 15, cuota_diaria_cup: 2333.33, reservado_cup: 0, pagado_cup: 0, estado: "activa", ...extra,
  };
  db.provisiones.push(p);
  return p;
}

test("el día que empieza la quincena se propone el salario en CUP con la TRMI de ese día", async () => {
  const today = todayHavana();
  db.gastos_fijos.push({ id: 1, concepto: "Salario quincenal", beneficiario: "Diana Fernández", cuenta_id: 16, moneda: "USD", importe: 70, periodicidad: "quincenal", proximo_inicio: today, activo: true });
  await cron(CRON_TOKEN, "alertas_diarias");
  const op = db.operaciones_pendientes[0];
  assertEquals(lines(op), ["16:33600/0 CUP", "8:0/33600 CUP"]);
  assertEquals(op.propuesta.tasa_referencia.tasa, 480);
  const p = db.provisiones[0];
  assertEquals([p.importe_cup, p.tasa_fijada, p.estado, p.fecha_pago], [33600, 480, "pendiente", Number(today.slice(8)) <= 15 ? `${today.slice(0, 7)}-15` : p.fecha_pago]);
  assert(db.gastos_fijos[0].proximo_inicio > today);
  assertStringIncludes(sent.map((m) => m.text).join("\n"), "Gasto fijo que empieza hoy");
  await say(`/confirmar ${op.id}`);
  assertEquals(p.estado, "activa");
  // Al día siguiente no se repite.
  await cron(CRON_TOKEN, "alertas_diarias");
  assertEquals(db.operaciones_pendientes.length, 1);
});

test("el cierre del turno aparta la cuota diaria a Caja reservada y recupera los días sin cierre", async () => {
  seedShift();
  const p = seedProvision({ fecha_inicio: "2026-09-30", importe_cup: 1500, dias: 15, cuota_diaria_cup: 100 });
  const reply = await say("/cierre 2026-10-01 efectivo 8000 transferencia 4000");
  const reserva = db.operaciones_pendientes.find((o) => o.propuesta.tipo_operacion === "reserva_obligaciones")!;
  assertEquals(lines(reserva), ["35:200/0 CUP", "1:0/200 CUP"]);
  assertStringIncludes(reply, "Reserva para obligaciones");
  await say(`/confirmar ${reserva.id}`);
  assertEquals(p.reservado_cup, 200);
});

test("/pagar usa primero lo reservado y el resto sale de la caja", async () => {
  const p = seedProvision({ reservado_cup: 30000 });
  const reply = await say("/pagar diana");
  const op = db.operaciones_pendientes[0];
  assertEquals(lines(op), ["8:35000/0 CUP", "35:0/30000 CUP", "1:0/5000 CUP"]);
  assertStringIncludes(reply, "70 USD a la tasa fijada 500");
  await say(`/confirmar ${op.id}`);
  assertEquals([p.estado, p.pagado_cup, p.reservado_cup], ["pagada", 35000, 0]);
  assertStringIncludes(await say("/pagar diana"), "No hay salarios");
});

test("/pagar con transfermovil saca el resto de Transfermóvil", async () => {
  seedProvision({ beneficiario: "Renta", concepto: "Alquiler del local", importe_origen: 6000, moneda_origen: "CUP", tasa_fijada: null, importe_cup: 6000, reservado_cup: 0 });
  await say("/pagar renta transfermovil");
  assertEquals(lines(db.operaciones_pendientes[0]), ["8:6000/0 CUP", "19:0/6000 CUP"]);
});

test("/fijos lista los gastos y cambia la fecha de inicio", async () => {
  db.gastos_fijos.push({ id: 1, concepto: "Salario quincenal", beneficiario: "Gerardo (mensajero)", cuenta_id: 16, moneda: "USD", importe: 50, periodicidad: "quincenal", proximo_inicio: "2026-11-01", activo: true });
  seedProvision();
  const list = await say("/fijos");
  assertStringIncludes(list, "Gerardo (mensajero) — Salario quincenal: 50,00 USD quincenal");
  assertStringIncludes(list, "35.000,00 CUP (70 USD a 500)");
  assertStringIncludes(await say("/fijos inicio gerardo 2026-10-16"), "empieza el 2026-10-16");
  assertEquals(db.gastos_fijos[0].proximo_inicio, "2026-10-16");
  assertStringIncludes(await say("/fijos importe gerardo 60 USD"), "60,00 USD");
  assertEquals(db.gastos_fijos[0].importe, 60);
});

test("/utilidad reparte el 30 % menos los consumos del dueño y lo fija en CUP", async () => {
  estadoResultado = 1000;
  db.asientos_contables.push({ id: 1, fecha: "2026-09-10", estado: "validado" });
  db.lineas_asiento.push({ id: 2, asiento_id: 1, cuenta_id: 36, debe: 20, haber: 0, moneda: "USD" }, { id: 3, asiento_id: 1, cuenta_id: 4, debe: 0, haber: 20, moneda: "USD" });
  const reply = await say("/utilidad 2026-09");
  assertStringIncludes(reply, "30 % para Christian: 300,00 USD");
  assertStringIncludes(reply, "Neto a pagar: 280,00 USD");
  const op = db.operaciones_pendientes[0];
  assertEquals(lines(op), ["37:300/0 USD", "36:0/20 USD", "20:0/280 USD", "20:134400/0 CUP", "38:0/134400 CUP"]);
  const p = db.provisiones[0];
  assertEquals([p.tipo, p.importe_cup, p.fecha_inicio, p.fecha_pago, p.dias], ["utilidad", 134400, "2026-10-01", "2026-11-01", 31]);
  assertStringIncludes(await say("/utilidad 2026-09"), "ya tiene la operación");
  assertStringIncludes(await say("/utilidad 2099-01"), "todavía no terminó");
});

test("/utilidad no propone nada si los consumos superan la parte del dueño", async () => {
  estadoResultado = 50;
  db.asientos_contables.push({ id: 1, fecha: "2026-09-10", estado: "validado" });
  db.lineas_asiento.push({ id: 2, asiento_id: 1, cuenta_id: 36, debe: 20, haber: 0, moneda: "USD" });
  assertStringIncludes(await say("/utilidad 2026-09"), "Los consumos superan su parte");
  assertEquals(db.operaciones_pendientes.length, 0);
});

function shiftDay(iso: string, days: number): string {
  const d = new Date(`${iso}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

test("cada obligación se aparta en su propia propuesta; si no se apartó, no suma y el siguiente cierre lo recupera", async () => {
  seedShift();
  db.registros_excel_diario.push({ fecha_inicio: "2026-10-02", fecha_fin: "2026-10-02", producto: "Cerveza", venta: 40, importe: 12000, importe_costo: 6000, costo: 150, tasa_ref: 500 });
  const diana = seedProvision({ fecha_inicio: "2026-10-01", importe_cup: 1500, dias: 15, cuota_diaria_cup: 100 });
  const renta = seedProvision({ beneficiario: "Renta", concepto: "Alquiler del local", fecha_inicio: "2026-10-01", importe_cup: 700, dias: 7, cuota_diaria_cup: 100, fecha_pago: "2026-10-07" });
  const reply = await say("/cierre 2026-10-01 efectivo 8000 transferencia 4000");
  assertStringIncludes(reply, "¿Se apartaron 100,00 CUP para Diana Fernández (Salario quincenal)?");
  assertStringIncludes(reply, "¿Se apartaron 100,00 CUP para Renta (Alquiler del local)?");
  const reservas = () => db.operaciones_pendientes.filter((o) => o.propuesta.tipo_operacion === "reserva_obligaciones");
  const [rd, rr] = reservas();
  await say(`/confirmar ${rd.id}`);
  await say(`/cancelar ${rr.id}`);
  assertEquals([diana.reservado_cup, renta.reservado_cup], [100, 0]);
  const reply2 = await say("/cierre 2026-10-02 efectivo 8000 transferencia 4000");
  assertStringIncludes(reply2, "¿Se apartaron 200,00 CUP para Renta");
  assertStringIncludes(reply2, "Incluye 100,00 CUP de días en que no se apartó");
  assertEquals(reservas().slice(2).map((o) => lines(o)[0]), ["35:100/0 CUP", "35:200/0 CUP"]);
});

test("una reserva sin confirmar no se vuelve a proponer en el siguiente cierre", async () => {
  seedShift();
  db.registros_excel_diario.push({ fecha_inicio: "2026-10-02", fecha_fin: "2026-10-02", producto: "Cerveza", venta: 40, importe: 12000, importe_costo: 6000, costo: 150, tasa_ref: 500 });
  seedProvision({ fecha_inicio: "2026-10-01", importe_cup: 1500, dias: 15, cuota_diaria_cup: 100 });
  await say("/cierre 2026-10-01 efectivo 8000 transferencia 4000");
  await say("/cierre 2026-10-02 efectivo 8000 transferencia 4000");
  const r = db.operaciones_pendientes.filter((o) => o.propuesta.tipo_operacion === "reserva_obligaciones").map((o) => lines(o)[0]);
  assertEquals(r, ["35:100/0 CUP", "35:100/0 CUP"]);
});

test("avisa el día antes y el día del pago pregunta si se pagó; si no, vuelve a preguntar al día siguiente", async () => {
  const today = todayHavana();
  seedProvision({ beneficiario: "Diana Fernández", fecha_pago: shiftDay(today, 1), reservado_cup: 30000 });
  const renta = seedProvision({ beneficiario: "Renta", concepto: "Alquiler del local", importe_origen: 6000, moneda_origen: "CUP", tasa_fijada: null, importe_cup: 6000, reservado_cup: 6000, fecha_pago: today });
  await cron(CRON_TOKEN, "alertas_diarias");
  const text = sent.map((m) => m.text).join("\n");
  assertStringIncludes(text, `⏰ Mañana (${shiftDay(today, 1)}) toca pagar a Diana Fernández (Salario quincenal): 35.000,00 CUP. Apartado en Caja reservada: 30.000,00 CUP; faltan 5000,00 CUP por apartar.`);
  assertStringIncludes(text, "📅 Hoy toca pagar Renta (Alquiler del local): 6000,00 CUP.\n¿Se pagó?");
  assert(sent.some((m) => JSON.stringify(m.reply_markup ?? "").includes("confirmar:")));
  const pago = db.operaciones_pendientes.find((o) => o.propuesta.tipo_operacion === "pago_obligacion")!;
  assertEquals(lines(pago), ["8:6000/0 CUP", "35:0/6000 CUP"]);
  // Sin respuesta todavía: no se duplica.
  await cron(CRON_TOKEN, "alertas_diarias");
  assertEquals(db.operaciones_pendientes.filter((o) => o.propuesta.tipo_operacion === "pago_obligacion").length, 1);
  assertStringIncludes(await say("/pagar renta"), "ya espera confirmación");
  // «No se pagó»: se cancela y al día siguiente vuelve a preguntar.
  await say(`/cancelar ${pago.id}`);
  assertEquals(renta.estado, "activa");
  await cron(CRON_TOKEN, "alertas_diarias");
  assertEquals(db.operaciones_pendientes.filter((o) => o.propuesta.tipo_operacion === "pago_obligacion").length, 2);
  const otra = db.operaciones_pendientes.filter((o) => o.propuesta.tipo_operacion === "pago_obligacion")[1];
  await say(`/confirmar ${otra.id}`);
  assertEquals([renta.estado, renta.pagado_cup, renta.reservado_cup], ["pagada", 6000, 0]);
});

/* ----- Excel con un producto en varias filas ----- */

async function sendExcel(fileName: string, rows: unknown[][]): Promise<string> {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Hoja1");
  telegramFile = new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
  sent.length = 0;
  await handler(new Request("https://bot.test", {
    method: "POST",
    headers: { "x-telegram-bot-api-secret-token": SECRET },
    body: JSON.stringify({ update_id: updateId++, message: { date: Math.floor(Date.now() / 1000), chat: { id: OWNER, type: "private" }, from: { id: OWNER }, document: { file_id: "xl1", file_name: fileName, file_size: 5000 } } }),
  }));
  telegramFile = new Uint8Array([1, 2, 3, 4]);
  return sent.map((m) => m.text ?? "").join("\n");
}

test("una fila repetida vacía se omite y el resto se importa", async () => {
  db.registros_excel_diario.length = 0;
  const reply = await sendExcel("26  octubre 2026.xlsx", [
    ["Producto", "Inicio", "Entrada", "Rotura", "Consumo", "Disponible", "Final", "Precio", "Venta", "Importe"],
    ["Mortadella", 7, null, null, null, 7, 1, 1300, 6, 7800],
    ["Mortadella", null, null, null, null, 0, null, 1500, 0, 0],
    ["Arroz", 5, null, null, null, 5, 5, 200, 0, 0],
  ]);
  assertStringIncludes(reply, "Importación confirmada");
  assertStringIncludes(reply, "Productos guardados: 2");
  assertStringIncludes(reply, "Mortadella: 1 fila(s) repetida(s) vacía(s) no se guardaron");
  const mort = db.registros_excel_diario.find((r) => r.producto === "Mortadella")!;
  assertEquals([mort.inicio, mort.final, mort.venta, mort.precio], [7, 1, 6, 1300]);
  db.registros_excel_diario.length = 0;
});

test("dos productos con el mismo nombre y datos no se mezclan: pide nombres distintos", async () => {
  db.registros_excel_diario.length = 0;
  const reply = await sendExcel("26  octubre 2026.xlsx", [
    ["Producto", "Inicio", "Entrada", "Disponible", "Final", "Precio", "Venta", "Importe"],
    ["Sopa instant", 2, null, 2, 0, 350, 2, 700],
    ["Sopa instant", 40, null, 40, 40, 350, 0, 0],
    ["Café", 3, null, 3, 1, 500, 2, 1000],
    ["Cafe ", null, 10, 10, 10, 500, 0, 0],
    ["Arroz", 5, null, 5, 5, 200, 0, 0],
  ]);
  assertStringIncludes(reply, "No importé nada");
  assertStringIncludes(reply, "nombres distintos");
  assertStringIncludes(reply, "• Sopa instant:\n   1) precio 350, inicio 2, venta 2\n   2) precio 350, inicio 40, final 40");
  assertStringIncludes(reply, "• Café:");
  assertEquals(db.registros_excel_diario.length, 0);
});

test("/ayuda registra el menú de comandos de Telegram y cada comando existe", async () => {
  const reply = await say("/ayuda");
  assertStringIncludes(reply, "Comandos:");
  const call = sent.find((m) => m.method === "setMyCommands");
  assert(call, "se llamó a setMyCommands");
  const commands = call.commands as { command: string; description: string }[];
  assert(commands.length > 0 && commands.length <= 100);
  const source = await Deno.readTextFile(new URL("./index.ts", import.meta.url));
  for (const c of commands) {
    assertMatch(c.command, /^[a-z0-9_]{1,32}$/);
    assert(c.description.length >= 1 && c.description.length <= 256);
    assertStringIncludes(source, `case "${c.command}":`);
  }
  assertEquals(new Set(commands.map((c) => c.command)).size, commands.length);
});

test("si Telegram rechaza el menú, /start responde igual", async () => {
  const realFetch = globalThis.fetch;
  globalThis.fetch = (input: RequestInfo | URL, init?: RequestInit) =>
    String(input instanceof Request ? input.url : input).endsWith("/setMyCommands")
      ? Promise.resolve(new Response(JSON.stringify({ ok: false }), { status: 400 }))
      : realFetch(input, init);
  try {
    assertStringIncludes(await say("/start"), "Comandos:");
  } finally {
    globalThis.fetch = realFetch;
  }
});
