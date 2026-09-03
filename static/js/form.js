// Memoria Descriptiva - Form JS
var STEPS = [
  {num:1, label:"Datos Gen."},
  {num:2, label:"Croquis"},
  {num:3, label:"Desc. Gen."},
  {num:4, label:"Exterior"},
  {num:5, label:"Interiores"},
  {num:6, label:"Giros"},
  {num:7, label:"Trabajos"},
  {num:8, label:"Estructura"},
  {num:9, label:"Instalaciones"},
  {num:10, label:"Seguridad"}
];
var currentStep = 1, totalSteps = 10;
var photoState = {}, croquis = null, listItems = [], saveTimer = null;

// Check if running on localhost/server vs file://
function checkEnvironment() {
  if (window.location.protocol === "file:" || !window.FORM_ID || window.FORM_ID.indexOf("{{") !== -1) {
    var warning = document.createElement("div");
    warning.style.cssText = "position:fixed;top:0;left:0;right:0;background:#dc3545;color:#fff;padding:14px;text-align:center;font-weight:bold;z-index:99999;box-shadow:0 4px 12px rgba(0,0,0,.3)";
    warning.innerHTML = "⚠️ ATENCIÓN: No estás corriendo el servidor. Para subir fotos y generar el Word, debes abrir desde <u>http://localhost:5000</u> ejecutando antes <b>iniciar.bat</b>";
    document.body.prepend(warning);
  }
}

function buildProgress() {
  var prog = document.getElementById("step-progress");
  if (!prog) return;
  prog.innerHTML = STEPS.map(function(s) {
    var cls = "step-item";
    if (s.num === currentStep) cls += " active";
    else if (s.num < currentStep) cls += " done";
    return '<div class="' + cls + '" onclick="goStep(' + s.num + ')"><span class="step-num">' + s.num + '</span>' + s.label + '</div>';
  }).join("");
}

function showStep(n) {
  document.querySelectorAll(".step-panel").forEach(function(p) { p.style.display = "none"; });
  var panel = document.querySelector('[data-step="' + n + '"]');
  if (panel) panel.style.display = "block";
  var prevBtn = document.getElementById("btn-prev");
  var nextBtn = document.getElementById("btn-next");
  if (prevBtn) prevBtn.style.display = n === 1 ? "none" : "";
  if (nextBtn) nextBtn.style.display = n === totalSteps ? "none" : "";
  buildProgress();
}

function goStep(n) { saveData(); currentStep = n; showStep(n); window.scrollTo(0,0); }
function nextStep() { if (currentStep < totalSteps) { saveData(); currentStep++; showStep(currentStep); window.scrollTo(0,0); } }
function prevStep() { if (currentStep > 1) { saveData(); currentStep--; showStep(currentStep); window.scrollTo(0,0); } }

// ---- PHOTOS ----
function triggerFile(zone) {
  var fileInput = zone.querySelector("input[type=file]");
  if (fileInput) fileInput.click();
}
function ev(e) { e.preventDefault(); }
function drop(e, zone) {
  e.preventDefault();
  if (e.dataTransfer && e.dataTransfer.files) {
    uploadFiles(e.dataTransfer.files, zone.dataset.section);
  }
}
function addPhotos(input) {
  var zone = input.closest(".photo-zone");
  if (zone && input.files) {
    uploadFiles(input.files, zone.dataset.section);
  }
  input.value = "";
}

async function uploadFiles(files, section) {
  if (!files || !files.length) return;
  showSaving("Subiendo fotos...");
  for (var i = 0; i < files.length; i++) {
    try {
      var fd = new FormData();
      fd.append("file", files[i]);
      var res = await fetch("/api/upload/" + FORM_ID, { method: "POST", body: fd });
      var data = await res.json();
      if (data.url) {
        if (!photoState[section]) photoState[section] = [];
        photoState[section].push({ url: data.url, filename: data.filename, descripcion: "" });
        renderPhotoGrid(section);
        scheduleSave();
      } else {
        alert("Error al subir imagen: " + (data.error || "Formato no permitido"));
      }
    } catch (err) {
      console.error("Upload error:", err);
      alert("No se pudo conectar al servidor para subir la imagen. Verifica que el servidor esté activo.");
    }
  }
  hideSaving();
}

function renderPhotoGrid(section) {
  var grid = document.querySelector('.photo-grid[data-section="' + section + '"]');
  if (!grid) return;
  var photos = photoState[section] || [];
  var html = "";
  for (var i = 0; i < photos.length; i++) {
    var p = photos[i];
    html += '<div class="photo-item">';
    html += '<img src="' + p.url + '" alt="foto">';
    html += '<div class="photo-desc">';
    html += '<textarea placeholder="Descripción de la foto..." onchange="updateDesc(\'' + section + '\',' + i + ',this.value)">' + (p.descripcion || "") + '</textarea>';
    html += '</div>';
    html += '<button type="button" class="photo-remove" onclick="removePhoto(\'' + section + '\',' + i + ')">✕ Eliminar foto</button>';
    html += '</div>';
  }
  grid.innerHTML = html;
}

function updateDesc(section, idx, val) {
  if (photoState[section] && photoState[section][idx]) {
    photoState[section][idx].descripcion = val;
    scheduleSave();
  }
}

async function removePhoto(section, idx) {
  var photo = (photoState[section] || [])[idx];
  if (photo) {
    try {
      await fetch("/api/upload/" + FORM_ID + "/" + photo.filename, { method: "DELETE" });
    } catch(e) {}
    photoState[section].splice(idx, 1);
    renderPhotoGrid(section);
    scheduleSave();
  }
}

// ---- CROQUIS ----
async function uploadCroquis(input) {
  var file = input.files[0];
  if (!file) return;
  showSaving("Subiendo croquis...");
  try {
    var fd = new FormData();
    fd.append("file", file);
    var res = await fetch("/api/upload/" + FORM_ID, { method: "POST", body: fd });
    var data = await res.json();
    if (data.url) {
      croquis = { url: data.url, filename: data.filename };
      var ph = document.getElementById("croquis-placeholder");
      if (ph) ph.style.display = "none";
      var prev = document.getElementById("croquis-preview");
      if (prev) prev.innerHTML = '<img src="' + data.url + '" style="max-width:100%;max-height:300px;border-radius:8px">';
      scheduleSave();
    } else {
      alert("Error al subir croquis: " + (data.error || ""));
    }
  } catch(e) {
    alert("Error de conexión al subir el croquis.");
  }
  input.value = "";
  hideSaving();
}

// ---- LIST ITEMS ----
function addListItem(val) { listItems.push(val || ""); renderListItems(); }
function renderListItems() {
  var area = document.getElementById("lista-items");
  if (!area) return;
  var html = "";
  for (var i = 0; i < listItems.length; i++) {
    html += '<div class="list-item-row">';
    html += '<input type="text" class="form-control" value="' + (listItems[i] || "").replace(/"/g, "&quot;") + '" placeholder="Elemento de remodelación..." oninput="listItems[' + i + ']=this.value;scheduleSave()">';
    html += '<button type="button" class="btn-remove" onclick="removeItem(' + i + ')">✕</button></div>';
  }
  area.innerHTML = html;
}
function removeItem(i) { listItems.splice(i, 1); renderListItems(); scheduleSave(); }

// ---- COLLECT DATA ----
function collectData() {
  var dg = {}, ext = {}, intr = {}, inst = {}, seg = {}, giros = {}, trabajos = {}, estructura = {}, dgen_texto = "";
  document.querySelectorAll(".form-control").forEach(function(el) {
    if (!el.name) return;
    var parts = el.name.split(".");
    var val = el.value;
    if (parts[0] === "dg") dg[parts[1]] = val;
    else if (parts[0] === "ext") { if (!ext[parts[1]]) ext[parts[1]] = {}; ext[parts[1]][parts[2]] = val; }
    else if (parts[0] === "int") { if (!intr[parts[1]]) intr[parts[1]] = {}; intr[parts[1]][parts[2]] = val; }
    else if (parts[0] === "inst") inst[parts[1]] = val;
    else if (parts[0] === "seg") seg[parts[1]] = val;
    else if (parts[0] === "giros") giros[parts[1]] = val;
    else if (parts[0] === "trabajos") trabajos[parts[1]] = val;
    else if (parts[0] === "estructura") estructura[parts[1]] = val;
    else if (parts[0] === "desc_general") dgen_texto = val;
  });
  Object.keys(photoState).forEach(function(sec) {
    var parts = sec.split(".");
    if (parts[0] === "ext") { if (!ext[parts[1]]) ext[parts[1]] = {}; ext[parts[1]].fotos = photoState[sec]; }
    else if (parts[0] === "int") { if (!intr[parts[1]]) intr[parts[1]] = {}; intr[parts[1]].fotos = photoState[sec]; }
  });
  return {
    datos_generales: dg,
    croquis: croquis,
    descripcion_general: { texto: dgen_texto, lista_items: listItems },
    remodelacion_exterior: ext,
    remodelacion_interiores: intr,
    giros_negocio: giros,
    descripcion_trabajos: trabajos,
    estructura: estructura,
    instalaciones: inst,
    medidas_seguridad: seg
  };
}

// ---- SAVE / LOAD ----
function scheduleSave() { clearTimeout(saveTimer); saveTimer = setTimeout(saveData, 1000); }
async function saveData() {
  var data = collectData();
  showSaving("Guardando...");
  try {
    await fetch("/api/form/" + FORM_ID, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data)
    });
  } catch (err) {
    console.error("Save error:", err);
  }
  hideSaving();
}

function showSaving(msg) {
  var el = document.getElementById("saving-ind");
  if (el) { el.textContent = msg || "Guardando..."; el.classList.add("show"); }
}
function hideSaving() {
  setTimeout(function() {
    var el = document.getElementById("saving-ind");
    if (el) el.classList.remove("show");
  }, 1000);
}

async function loadData() {
  try {
    var res = await fetch("/api/form/" + FORM_ID);
    var data = await res.json();
    if (data.error) return;

    var dg = data.datos_generales || {};
    Object.keys(dg).forEach(function(k) {
      var el = document.querySelector('[name="dg.' + k + '"]');
      if (el) el.value = dg[k] || "";
    });

    if (data.croquis && data.croquis.url) {
      croquis = data.croquis;
      var ph = document.getElementById("croquis-placeholder");
      if (ph) ph.style.display = "none";
      var prev = document.getElementById("croquis-preview");
      if (prev) prev.innerHTML = '<img src="' + data.croquis.url + '" style="max-width:100%;max-height:300px;border-radius:8px">';
    }

    var dgen = data.descripcion_general || {};
    var tEl = document.querySelector('[name="desc_general.texto"]');
    if (tEl) tEl.value = dgen.texto || "";
    listItems = dgen.lista_items || [];
    renderListItems();

    var ext = data.remodelacion_exterior || {};
    ["cubierta","estacionamiento","anuncio_espectacular","fachadas","anden","area_servicio"].forEach(function(k) {
      var s = ext[k] || {};
      var el = document.querySelector('[name="ext.' + k + '.texto"]');
      if (el) el.value = s.texto || "";
      if (s.fotos && s.fotos.length) { photoState["ext." + k] = s.fotos; renderPhotoGrid("ext." + k); }
    });

    var intr = data.remodelacion_interiores || {};
    ["portico_acceso","oficinas_frontales","sanitarios_clientes","piso_ventas","area_perecederos","comedor_asociados","oficinas_posteriores","sanitarios_asociados","acceso_personal","facturacion_sistemas","trastienda"].forEach(function(k) {
      var s = intr[k] || {};
      var el = document.querySelector('[name="int.' + k + '.texto"]');
      if (el) el.value = s.texto || "";
      if (s.fotos && s.fotos.length) { photoState["int." + k] = s.fotos; renderPhotoGrid("int." + k); }
    });

    var simples = [
      ["giros.consultorio_medico", (data.giros_negocio || {}).consultorio_medico],
      ["trabajos.texto", (data.descripcion_trabajos || {}).texto],
      ["trabajos.area_trabajo", (data.descripcion_trabajos || {}).area_trabajo],
      ["trabajos.afectacion_estructural", (data.descripcion_trabajos || {}).afectacion_estructural],
      ["estructura.texto", (data.estructura || {}).texto],
      ["seg.texto", (data.medidas_seguridad || {}).texto],
      ["seg.medidas_clientes", (data.medidas_seguridad || {}).medidas_clientes],
      ["seg.consideraciones", (data.medidas_seguridad || {}).consideraciones]
    ];
    simples.forEach(function(pair) {
      var el = document.querySelector('[name="' + pair[0] + '"]');
      if (el && pair[1]) el.value = pair[1];
    });

    var inst = data.instalaciones || {};
    ["refrigeracion","aire","electrica","hidro_sanitaria","gas","filtrado","contra_incendio"].forEach(function(k) {
      var el = document.querySelector('[name="inst.' + k + '"]');
      if (el && inst[k]) el.value = inst[k];
    });

    var nombre = (data.datos_generales || {}).nombre_tienda || "";
    if (nombre) document.getElementById("nav-title").textContent = "Editando: " + nombre;
  } catch (err) {
    console.error("Load error:", err);
  }
}

// ---- GENERATE WORD ----
async function generateWord() {
  var status = document.getElementById("gen-status");
  status.innerHTML = '<div class="alert" style="background:#e8f4ff;border:1px solid #b6d4fe;color:#084298">⏳ Guardando y generando documento Word, por favor espera...</div>';
  await saveData();

  try {
    var res = await fetch("/api/generar/" + FORM_ID, { method: "POST" });
    var data = await res.json();
    if (data.ok) {
      status.innerHTML = '<div class="alert alert-success">✅ ¡Documento generado con éxito! <a href="' + data.download + '" class="btn btn-success" style="margin-left:1rem">⬇ Descargar Word (.docx)</a></div>';
    } else {
      status.innerHTML = '<div class="alert alert-danger">❌ Error al generar: ' + (data.error || "Desconocido") + '</div>';
    }
  } catch (err) {
    console.error("Generate error:", err);
    status.innerHTML = '<div class="alert alert-danger">❌ Error de conexión con el servidor al generar Word.</div>';
  }
}

// ---- INIT ----
document.addEventListener("DOMContentLoaded", async function() {
  checkEnvironment();
  buildProgress();
  showStep(1);
  if (listItems.length === 0) addListItem();
  await loadData();
  var form = document.getElementById("main-form");
  if (form) form.addEventListener("input", scheduleSave);
});
