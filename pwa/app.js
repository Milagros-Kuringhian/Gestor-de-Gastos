const STORAGE_KEY = "mi-plata-v1";

const CATEGORIAS = [
  { id: "cobre", nombre: "Cobré", tipo: "ingreso", tag: "Ingreso" },
  { id: "padel", nombre: "Pádel", tipo: "gasto", tag: "Gasto" },
  { id: "facultad", nombre: "Facultad", tipo: "gasto", tag: "Gasto" },
  { id: "transporte", nombre: "Transporte", tipo: "gasto", tag: "Gasto" },
  { id: "hormiga", nombre: "Gastos hormiga", tipo: "gasto", tag: "Gasto" },
  { id: "otros", nombre: "Otros", tipo: "gasto", tag: "Gasto" },
];

const $ = (sel) => document.querySelector(sel);

const els = {
  fechaHoy: $("#fecha-hoy"),
  plata: $("#plata-disponible"),
  cobrado: $("#total-cobrado"),
  gastado: $("#total-gastado"),
  grid: $("#cat-grid"),
  lista: $("#lista-movimientos"),
  empty: $("#empty-state"),
  modalMonto: $("#modal-monto"),
  modalTitulo: $("#modal-titulo"),
  modalSub: $("#modal-sub"),
  inputMonto: $("#input-monto"),
  btnGuardar: $("#btn-guardar"),
  modalAjustes: $("#modal-ajustes"),
  inputSaldo: $("#input-saldo"),
  btnAjustes: $("#btn-ajustes"),
  btnGuardarSaldo: $("#btn-guardar-saldo"),
  btnExportar: $("#btn-exportar"),
  installHint: $("#install-hint"),
  toast: $("#toast"),
};

let state = loadState();
let categoriaActiva = null;

function defaultState() {
  return {
    saldoInicial: 0,
    movimientos: [],
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    return {
      saldoInicial: Number(parsed.saldoInicial) || 0,
      movimientos: Array.isArray(parsed.movimientos) ? parsed.movimientos : [],
    };
  } catch {
    return defaultState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function formatARS(n) {
  return new Intl.NumberFormat("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: 0,
  }).format(n || 0);
}

function hoyLabel() {
  return new Intl.DateTimeFormat("es-AR", {
    weekday: "long",
    day: "numeric",
    month: "long",
  }).format(new Date());
}

function hoyISO() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function totales() {
  let cobrado = 0;
  let gastado = 0;
  for (const mov of state.movimientos) {
    if (mov.tipo === "ingreso") cobrado += mov.monto;
    else gastado += mov.monto;
  }
  return {
    cobrado,
    gastado,
    disponible: state.saldoInicial + cobrado - gastado,
  };
}

function renderCategorias() {
  els.grid.innerHTML = CATEGORIAS.map(
    (cat) => `
    <button type="button" class="cat-btn ${cat.tipo}" data-id="${cat.id}">
      <span class="tag">${cat.tag}</span>
      <span class="name">${cat.nombre}</span>
    </button>`
  ).join("");
}

function render() {
  els.fechaHoy.textContent = hoyLabel();
  const t = totales();
  els.plata.textContent = formatARS(t.disponible);
  els.cobrado.textContent = formatARS(t.cobrado);
  els.gastado.textContent = formatARS(t.gastado);

  const ordenados = [...state.movimientos].sort((a, b) =>
    a.fechaISO < b.fechaISO ? 1 : a.fechaISO > b.fechaISO ? -1 : b.createdAt - a.createdAt
  );

  els.lista.innerHTML = ordenados
    .slice(0, 40)
    .map((mov) => {
      const signo = mov.tipo === "ingreso" ? "+" : "−";
      return `
      <li>
        <div class="meta">
          <strong>${mov.nombre}</strong>
          <span>${formatFechaCorta(mov.fechaISO)}</span>
        </div>
        <div class="monto ${mov.tipo}">${signo}${formatARS(mov.monto)}</div>
        <button type="button" class="btn-borrar" data-del="${mov.id}" aria-label="Borrar">
          Borrar
        </button>
      </li>`;
    })
    .join("");

  els.empty.hidden = ordenados.length > 0;
}

function formatFechaCorta(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  const date = new Date(y, m - 1, d);
  return new Intl.DateTimeFormat("es-AR", {
    day: "numeric",
    month: "short",
  }).format(date);
}

function abrirModalMonto(cat) {
  categoriaActiva = cat;
  els.modalTitulo.textContent = cat.nombre;
  els.modalSub.textContent =
    cat.tipo === "ingreso"
      ? "¿Cuánto cobraste hoy?"
      : "¿Cuánto gastaste?";
  els.inputMonto.value = "";
  els.modalMonto.hidden = false;
  setTimeout(() => els.inputMonto.focus(), 50);
}

function cerrarModales() {
  els.modalMonto.hidden = true;
  els.modalAjustes.hidden = true;
  categoriaActiva = null;
}

function guardarMovimiento() {
  if (!categoriaActiva) return;
  const monto = Number(String(els.inputMonto.value).replace(",", "."));
  if (!Number.isFinite(monto) || monto <= 0) {
    showToast("Poné un monto válido");
    els.inputMonto.focus();
    return;
  }

  const tipo = categoriaActiva.tipo;
  state.movimientos.push({
    id:
      (crypto.randomUUID && crypto.randomUUID()) ||
      `m-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    categoriaId: categoriaActiva.id,
    nombre: categoriaActiva.nombre,
    tipo,
    monto,
    fechaISO: hoyISO(),
    createdAt: Date.now(),
  });
  saveState();
  cerrarModales();
  render();
  showToast(tipo === "ingreso" ? "Cobro anotado" : "Gasto anotado");
}

function borrarMovimiento(id) {
  state.movimientos = state.movimientos.filter((m) => m.id !== id);
  saveState();
  render();
  showToast("Borrado");
}

function exportarCSV() {
  const header = "fecha,tipo,categoria,monto\n";
  const rows = state.movimientos
    .map(
      (m) =>
        `${m.fechaISO},${m.tipo},${escapeCsv(m.nombre)},${m.monto}`
    )
    .join("\n");
  const blob = new Blob([header + rows], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `mi-plata-${hoyISO()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast("CSV descargado");
}

function escapeCsv(value) {
  if (value.includes(",") || value.includes('"')) {
    return `"${value.replaceAll('"', '""')}"`;
  }
  return value;
}

let toastTimer;
function showToast(msg) {
  els.toast.textContent = msg;
  els.toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    els.toast.hidden = true;
  }, 1800);
}

function wireEvents() {
  els.grid.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-id]");
    if (!btn) return;
    const cat = CATEGORIAS.find((c) => c.id === btn.dataset.id);
    if (cat) abrirModalMonto(cat);
  });

  els.btnGuardar.addEventListener("click", guardarMovimiento);
  els.inputMonto.addEventListener("keydown", (e) => {
    if (e.key === "Enter") guardarMovimiento();
  });

  els.lista.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-del]");
    if (!btn) return;
    borrarMovimiento(btn.dataset.del);
  });

  els.btnAjustes.addEventListener("click", () => {
    els.inputSaldo.value = state.saldoInicial || "";
    els.modalAjustes.hidden = false;
    setTimeout(() => els.inputSaldo.focus(), 50);
  });

  els.btnGuardarSaldo.addEventListener("click", () => {
    const saldo = Number(String(els.inputSaldo.value).replace(",", "."));
    if (!Number.isFinite(saldo) || saldo < 0) {
      showToast("Saldo inválido");
      return;
    }
    state.saldoInicial = saldo;
    saveState();
    cerrarModales();
    render();
    showToast("Saldo guardado");
  });

  els.btnExportar.addEventListener("click", exportarCSV);

  document.querySelectorAll("[data-close]").forEach((el) => {
    el.addEventListener("click", cerrarModales);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") cerrarModales();
  });
}

function setupInstallHint() {
  const isStandalone =
    window.matchMedia("(display-mode: standalone)").matches ||
    window.navigator.standalone === true;
  if (!isStandalone) {
    els.installHint.hidden = false;
  }
}

function registerSW() {
  if (!("serviceWorker" in navigator)) return;
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch(() => {
      /* ignore in file:// */
    });
  });
}

renderCategorias();
wireEvents();
render();
setupInstallHint();
registerSW();
