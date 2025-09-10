/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const CIMA_BASE = "https://cima.aemps.es";

// Arranque y enganche del único botón
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const sideload = document.getElementById("sideload-msg");
    const body = document.getElementById("app-body");
    if (sideload) sideload.style.display = "none";
    if (body) body.style.display = "block";

    const attach = () => {
      const buscar = document.getElementById("btnBuscar");
      if (buscar && !buscar._h) {
        buscar.addEventListener("click", onBuscarPrincipioActivo);
        buscar._h = true;
      }
    };
    attach();
    document.addEventListener("DOMContentLoaded", attach);
  }
});

/* ========= UI ========= */
function setStatus(msg) {
  const el = document.getElementById("status");
  if (el) el.textContent = msg || "";
}
function setIndicaciones(text) {
  const el = document.getElementById("indicaciones");
  if (el) el.textContent = text && text.trim() ? text.trim() : "—";
}

/* ======== Word ======== */
async function getSelectedTextTrimmed() {
  let text = "";
  await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load("text");
    await context.sync();
    text = (sel.text || "").replace(/[\s\u00A0]+$/g, "").trim();
  });
  return text;
}
async function replaceSelection(text) {
  await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.insertText(text, Word.InsertLocation.replace);
    await context.sync();
  });
}

/* ======== Red/CIMA ======== */
// fetch JSON con fallback CORS (solo dev)
async function fetchJSON(url, timeoutMs = 12000) {
  const ctrl = new AbortController();
  const to = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const r = await fetch(url, { signal: ctrl.signal, headers: { Accept: "application/json" } });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return await r.json();
  } catch (err) {
    const proxied = `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`;
    const r2 = await fetch(proxied, { headers: { Accept: "application/json" } });
    if (!r2.ok) throw err;
    return await r2.json();
  } finally { clearTimeout(to); }
}
// fetch TEXT (HTML) con fallback CORS (solo dev)
async function fetchText(url, timeoutMs = 15000) {
  const ctrl = new AbortController();
  const to = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const r = await fetch(url, { signal: ctrl.signal, headers: { Accept: "text/html, text/plain" } });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return await r.text();
  } catch (err) {
    const proxied = `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`;
    const r2 = await fetch(proxied, { headers: { Accept: "text/html, text/plain" } });
    if (!r2.ok) throw err;
    return await r2.text();
  } finally { clearTimeout(to); }
}

// Normaliza diferentes formatos de respuesta a lista
function toList(resp) {
  if (Array.isArray(resp)) return resp;
  if (resp && Array.isArray(resp.resultados)) return resp.resultados; // paginado
  if (resp && Array.isArray(resp.lista)) return resp.lista;
  if (resp && Array.isArray(resp.datos)) return resp.datos;
  if (resp && resp.nombre) return [resp];
  return [];
}

// Búsqueda por nombre
async function searchMedsByName(name) {
  const url = `${CIMA_BASE}/cima/rest/medicamentos?nombre=${encodeURIComponent(name)}`;
  const resp = await fetchJSON(url);
  return toList(resp);
}

// Detalle por nregistro (trae pactivos/principiosActivos + docs)
async function getMedByNregistro(nreg) {
  const url = `${CIMA_BASE}/cima/rest/medicamento?nregistro=${encodeURIComponent(nreg)}`;
  const resp = await fetchJSON(url);
  const list = toList(resp);
  return list.length ? list[0] : null;
}

/* ======== Formato de salida ======== */
function toTitleCaseWord(w) { return w ? w[0].toUpperCase() + w.slice(1).toLowerCase() : w; }
function formatBrandFromSelection(sel) {
  const trimmed = (sel || "").trim();
  if (!trimmed) return "";
  return trimmed.split(/\s+/).map((t) => toTitleCaseWord(t)).join(" ");
}
function normalizeActiveName(s) {
  if (!s) return s;
  let out = s.toLowerCase().trim().replace(/\s+/g, " ");
  out = out.replace(/^(.+)\s+acido$/i, "ácido $1");
  out = out.replace(/\bacido\b/g, "ácido");
  out = out.replace(/\bsodico\b/g, "sódico").replace(/\bpotasico\b/g, "potásico");
  out = out.replace(/\bclorhidrico\b/g, "clorhídrico").replace(/\bhidroxido\b/g, "hidróxido");
  out = out.replace(/\bacetilsalicilico\b/g, "acetilsalicílico");
  return out;
}
function formatActives(actives) { return (actives || []).map(normalizeActiveName).join(", "); }
// Extrae principios activos
function getActivesFromMed(med) {
  if (Array.isArray(med?.principiosActivos) && med.principiosActivos.length) {
    const xs = med.principiosActivos.map((p) => p?.nombre || p?.principioActivo || p?.principio || "").filter(Boolean);
    if (xs.length) return xs;
  }
  if (typeof med?.pactivos === "string" && med.pactivos.trim()) return [med.pactivos.trim()];
  if (med?.pactivos && typeof med.pactivos === "object" && med.pactivos.nombre) return [String(med.pactivos.nombre).trim()];
  return [];
}

// Corrige la marca desde el nombre oficial (quita dosis/forma/EFG)
function deriveBrandFromMedName(med) {
  const raw = (med?.nombre || "").trim();
  if (!raw) return "";
  const tokens = raw.split(/\s+/);
  const chosen = [];
  const STOP = new Set([
    "COMPRIMIDOS","COMPRIMIDO","CAPSULAS","CÁPSULAS","CAPSULA","CÁPSULA","EFG","TURBUHALER",
    "AEROSOL","INHALADOR","INHALACIÓN","INHALACION","SOLUCIÓN","SOLUCION","SUSPENSIÓN","SUSPENSION",
    "JARABE","TABLETAS","TABLETA","TAB","VIAL","VIALES","RETARD","SR","ER","XR","MR","LIBERACIÓN","LIBERACION",
    "RECUBIERTOS","RECUBIERTO","GASTRORRESISTENTES","DISPERSABLE","DISPERSABLES","COLIRIO","PARCHES","PARCHE",
    "SPRAY","CREMA","GOTAS","POLVO","SOBRES","UNIDOSIS","SOLUBLE","ORAL","NASAL","OFTÁLMICA","OFTALMICA",
  ]);
  for (const t of tokens) {
    const T = t.toUpperCase();
    if (/^\d/.test(T)) break;
    if (/^\d+([.,]\d+)?(\/\d+([.,]\d+)?)?$/.test(T)) break;
    if (/^(MG|MCG|ΜG|UG|µG|G|ML|%|IU|UI)$/.test(T)) break;
    if (STOP.has(T)) break;
    chosen.push(t);
  }
  const base = chosen.length ? chosen : tokens.slice(0, 1);
  return base.map(toTitleCaseWord).join(" ");
}

/* ======== Erratas: variantes de búsqueda ======== */
async function trySearchVariants(query) {
  const tried = new Set();
  const tryVariant = async (v) => {
    const key = v.toLowerCase();
    if (tried.has(key)) return null;
    tried.add(key);
    const list = await searchMedsByName(v);
    if (list.length) return { list, usedVariant: v };
    return null;
  };
  const rules = [
    (s) => s.replace(/y/gi, "i"),
    (s) => s.replace(/i/gi, "y"),
    (s) => s.replace(/z/gi, "s"),
    (s) => s.replace(/s/gi, "z"),
    (s) => s.replace(/v/gi, "b"),
    (s) => s.replace(/b/gi, "v"),
    (s) => s.replace(/h/gi, ""),
  ];
  for (const fn of rules) {
    const v = fn(query);
    if (v !== query) {
      const res = await tryVariant(v);
      if (res) return res;
    }
  }
  for (let i = 0; i < query.length; i++) {
    const v = query.slice(0, i) + query.slice(i + 1);
    const res = await tryVariant(v);
    if (res) return res;
  }
  return null;
}

/* ======== Indicaciones (FT/Prospecto) ======== */
function stripSpaces(s) { return s.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim(); }
function normalizeNoAccents(s) { return s.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }

function extractIndicationsFromHTML(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const text = doc.body ? doc.body.innerText : html;
  const plain = stripSpaces(text);

  // 1) Ficha Técnica: sección 4.1 hasta 4.2/4.3/5.
  let reFT = /4\.\s*1[\s\S]*?(?=4\.\s*2|4\.\s*3|5\.)/i;
  let m = plain.match(reFT);
  if (m) {
    let out = m[0];
    const idx = out.toLowerCase().indexOf("indicaciones");
    if (idx > -1) out = out.slice(idx);
    return stripSpaces(out);
  }

  // 2) Prospecto: "para qué se utiliza"
  const noAcc = normalizeNoAccents(plain.toLowerCase());
  const idxPQ = noAcc.indexOf("para que se utiliza");
  if (idxPQ > -1) {
    const rest = plain.slice(idxPQ);
    const m2 = rest.match(/^[\s\S]*?(?=\n\s*\d+\s*\.)/m);
    return stripSpaces(m2 ? m2[0] : rest);
  }

  // 3) Fallback: buscar "indicaciones"
  const idxInd = noAcc.indexOf("indicaciones");
  if (idxInd > -1) {
    const rest = plain.slice(idxInd);
    const m3 = rest.match(/^[\s\S]*?(?=\n\s*\d+\s*\.)/m);
    return stripSpaces(m3 ? m3[0] : rest);
  }

  return null;
}

function pickDocHtml(med, tipoPreferido = 1) {
  if (!Array.isArray(med?.docs)) return null;
  const preferred = med.docs.filter(d => d.tipo === tipoPreferido && d.urlHtml);
  if (preferred.length) return preferred[0];
  const anyHtml = med.docs.find(d => d.urlHtml);
  return anyHtml || null;
}

async function getIndicaciones(med) {
  try {
    const docFT = pickDocHtml(med, 1); // Ficha Técnica
    const docP = pickDocHtml(med, 2);  // Prospecto
    const chosen = docFT || docP;
    if (!chosen || !chosen.urlHtml) return null;

    const html = await fetchText(chosen.urlHtml);
    const ind = extractIndicationsFromHTML(html);
    if (!ind) return null;

    // Limita a 1200 caracteres sin cortar palabra
    const max = 1200;
    if (ind.length <= max) return ind;
    const cut = ind.lastIndexOf(" ", max - 10);
    return ind.slice(0, cut > 0 ? cut : max).trim() + "…";
  } catch {
    return null;
  }
}

/* ======== Acción principal ======== */
async function onBuscarPrincipioActivo() {
  try {
    setStatus("Buscando...");
    setIndicaciones("—");

    const selected = await getSelectedTextTrimmed();
    if (!selected) {
      setStatus("Selecciona un nombre comercial.");
      return;
    }

    // Paso A: buscar candidatos
    let list = await searchMedsByName(selected);

    // Paso B: si no hay resultados, probar variantes por erratas
    if (!list.length) {
      const alt = await trySearchVariants(selected);
      if (alt && alt.list && alt.list.length) {
        list = alt.list;
      }
    }
    if (!list.length) {
      setStatus("No encontrado en CIMA.");
      return;
    }

    const candidato = list[0];

    // Paso C: detalle por nregistro
    const nreg = candidato.nregistro || candidato.nregistroId || candidato.id;
    const medDetalle = nreg ? await getMedByNregistro(nreg) : null;
    const med = medDetalle || candidato;

    // Paso D: activos y reemplazo en el documento
    const actives = getActivesFromMed(med);
    if (!actives.length) {
      setStatus("Sin principios activos en la respuesta.");
      return;
    }
    const brandCorrected = deriveBrandFromMedName(med);
    const brandFallback = formatBrandFromSelection(selected);
    const brandDisplay = brandCorrected || brandFallback;

    await replaceSelection(`${brandDisplay} (${formatActives(actives)})`);
    setStatus("Listo.");

    // Paso E: Indicaciones (solo panel)
    setStatus("Buscando indicaciones…");
    const indic = await getIndicaciones(med);
    if (indic) {
      setIndicaciones(indic);
      setStatus("Listo.");
    } else {
      setIndicaciones("No se pudieron extraer las indicaciones de la Ficha Técnica/Prospecto.");
      setStatus("Listo.");
    }
  } catch (e) {
    console.error(e);
    setStatus("Error consultando CIMA. Revisa tu conexión.");
  }
}

