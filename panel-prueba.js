/* Prueba del Panel PRAXIS en el borrador Baires de Wess. Se carga a mano desde la pestaña: no es parte de ninguna web. */
(function(){
var PX=window.PX;PX.enEditor=true;window.__pxpGuardados=[];
var nodo=document.querySelector('[data-px="faq"]');
var C={archivo:'prueba',edicion_url:'',faq:{titulo:nodo.querySelector('.pxf-tit').textContent,items:[].map.call(nodo.querySelectorAll('.pxf-it'),function(d){return{p:d.querySelector('.pxf-qt').textContent,r:d.querySelector('.pxf-a').textContent.trim()}})},
'panel-praxis':{bloques:{faq:{nombre:'Preguntas frecuentes',excel:true,
 esquema:{nombre:'Preguntas frecuentes',campos:[{k:'titulo',t:'Título',tipo:'texto',ideal:60,max:100},{k:'items',t:'Preguntas',singular:'Pregunta',tipo:'lista',max:12,min:1,nuevo:'Agregar pregunta',item:[{k:'p',t:'Pregunta',tipo:'texto',ideal:90,max:150},{k:'r',t:'Respuesta',tipo:'area',ideal:500,max:1200}]}]},
 variantes:{pc:[{v:'lista',n:'Lista'},{v:'columnas',n:'Dos columnas',css:'& .pxf-lista{display:grid;grid-template-columns:1fr 1fr;column-gap:40px}'}],cel:[{v:'lista',n:'Lista'},{v:'tarjetas',n:'Tarjetas',css:'& .pxf-it{margin-bottom:10px;padding:0 14px;border:1px solid #333;border-radius:10px}'}]}}}}};
PX.rehacer=function(id){if(id!=='faq')return;var F=C.faq;nodo.querySelector('.pxf-tit').textContent=F.titulo||'';var L=nodo.querySelector('.pxf-lista');var m=L.querySelector('.pxf-it').cloneNode(true);L.textContent='';(F.items||[]).forEach(function(it){var d=m.cloneNode(true);d.querySelector('.pxf-qt').textContent=it.p||'';var a=d.querySelector('.pxf-a');a.textContent='';var p=document.createElement('p');p.textContent=it.r||'';a.appendChild(p);L.appendChild(d)})};
PX.panelGuardar=function(b,c){(window.parent.__pxpG=window.parent.__pxpG||[]).push({bloque:b,campos:c});return Promise.resolve()};
window.__pxpC=C;
/* ═════════════════════════════════════════════════════════════
   DATOS POR EXCEL · bajar → cambiar → subir (02-11-2026)
   Nace de lo que hizo el cotizador de Aurel (23-10), ahora para
   cualquier bloque: las hojas salen de su esquema (PX.esquemas).

     · texto / área  → hoja «Textos», una fila por campo
     · lista         → una hoja por lista, una fila por ítem
     · lo que no es texto (fotos, videos) no viaja por el Excel

   Lo que vuelve se valida antes de tocar nada: si una fila está mal,
   se avisa cuál y no se aplica el archivo. Sólo corre en el editor.
   ═════════════════════════════════════════════════════════════ */
(function () {
  var XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  var HOJA_MAX = 31;

  var conXLSX = function (fn) {
    if (window.XLSX) { fn(window.XLSX); return; }
    var sc = document.createElement('script');
    sc.src = XLSX_URL;
    sc.onload = function () { fn(window.XLSX); };
    sc.onerror = function () { fn(null); };
    document.head.appendChild(sc);
  };
  /* Excel no acepta estos caracteres en el nombre de una hoja */
  var nombreHoja = function (t) { return String(t).replace(/[\\\/?*\[\]:]/g, ' ').trim().slice(0, HOJA_MAX); };
  var esTexto = function (c) { return c.tipo === 'texto' || c.tipo === 'area' || c.tipo === 'numero'; };

  var hojasDe = function (esq, datos) {
    var H = {};
    var textos = [['Campo', 'Texto']];
    (esq.campos || []).forEach(function (c) {
      if (esTexto(c)) { textos.push([c.t, datos[c.k] == null ? '' : datos[c.k]]); }
    });
    H['Cómo usar'] = [
      ['Cómo cambiar «' + esq.nombre + '» con este Excel'], [''],
      ['1. Cambiá los datos en las otras hojas. Cada fila es un ítem: podés sumar, sacar o reordenar filas.'],
      ['2. No cambies los títulos de las columnas ni los nombres de las hojas.'],
      ['3. Guardalo como .xlsx y subilo con «Subir Excel». Vas a ver el cambio antes de guardar.']
    ];
    if (textos.length > 1) { H.Textos = textos; }
    (esq.campos || []).forEach(function (c) {
      if (c.tipo !== 'lista') { return; }
      var cols = (c.item || []).filter(esTexto);
      if (!cols.length) { return; }
      var filas = [cols.map(function (x) { return x.t; })];
      (datos[c.k] || []).forEach(function (it) {
        filas.push(cols.map(function (x) { return it[x.k] == null ? '' : it[x.k]; }));
      });
      H[nombreHoja(c.t)] = filas;
    });
    return H;
  };

  var valor = function (c, v, donde, err) {
    v = typeof v === 'string' ? v.trim() : v;
    if (c.tipo === 'numero') {
      if (v === '') { err.push(donde + ': «' + c.t + '» no puede quedar vacío.'); return ''; }
      var n = parseFloat(String(v).replace(',', '.'));
      if (!isFinite(n)) { err.push(donde + ': «' + c.t + '» tiene que ser un número.'); return ''; }
      return n;
    }
    v = String(v == null ? '' : v);
    if (c.max && v.length > c.max) { err.push(donde + ': «' + c.t + '» pasa los ' + c.max + ' caracteres.'); }
    return v;
  };

  var leerHojas = function (X, wb, esq, datos) {
    var err = [], out = {};
    var filas = function (n) { return X.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' }); };
    var campos = esq.campos || [];
    if (campos.some(esTexto)) {
      if (!wb.Sheets.Textos) { err.push('Falta la hoja «Textos».'); }
      else {
        var tx = {};
        filas('Textos').slice(1).forEach(function (f) { tx[String(f[0]).trim()] = f[1]; });
        campos.filter(esTexto).forEach(function (c) {
          if (c.t in tx) { out[c.k] = valor(c, tx[c.t], 'Textos', err); }
        });
      }
    }
    campos.forEach(function (c) {
      if (c.tipo !== 'lista') { return; }
      var cols = (c.item || []).filter(esTexto);
      if (!cols.length) { return; }
      var n = nombreHoja(c.t);
      if (!wb.Sheets[n]) { err.push('Falta la hoja «' + n + '».'); return; }
      var todo = filas(n), cab = (todo[0] || []).map(function (x) { return String(x).trim(); });
      var pos = cols.map(function (x) { return cab.indexOf(x.t); });
      cols.forEach(function (x, i) { if (pos[i] < 0) { err.push('En «' + n + '» falta la columna «' + x.t + '».'); } });
      if (pos.indexOf(-1) > -1) { return; }
      var lista = [];
      todo.slice(1).forEach(function (f, i) {
        if (!f.join('').trim()) { return; }
        var it = {};
        cols.forEach(function (x, j) { it[x.k] = valor(x, f[pos[j]], n + ', fila ' + (i + 2), err); });
        lista.push(it);
      });
      if (c.max && lista.length > c.max) { err.push('«' + n + '» admite hasta ' + c.max + ' filas.'); }
      if (c.min && lista.length < c.min) { err.push('«' + n + '» necesita al menos ' + c.min + ' filas.'); }
      if (c.validar) { c.validar(lista, err); }
      /* lo que no viaja por el Excel (fotos de cada ítem) sigue en su fila */
      var antes = datos[c.k] || [];
      out[c.k] = lista.map(function (it, i) {
        var o = {}, a = antes[i] || {};
        Object.keys(a).forEach(function (k) { if (k !== '_o') { o[k] = a[k]; } });
        Object.keys(it).forEach(function (k) { o[k] = it[k]; });
        return o;
      });
    });
    return { campos: out, err: err };
  };

  PX.excel = {
    hojasDe: hojasDe,
    leerHojas: leerHojas,
    conXLSX: conXLSX,
    bajar: function (esq, datos, listo) {
      conXLSX(function (X) {
        if (!X) { listo('No se pudo preparar el Excel. Probá de nuevo.'); return; }
        var wb = X.utils.book_new(), H = hojasDe(esq, datos);
        Object.keys(H).forEach(function (n) {
          var ws = X.utils.aoa_to_sheet(H[n]);
          ws['!cols'] = (H[n][0] || []).map(function (c) { return { wch: n === 'Cómo usar' ? 100 : Math.max(16, Math.min(60, String(c).length + 6)) }; });
          X.utils.book_append_sheet(wb, ws, n);
        });
        X.writeFile(wb, esq.nombre + ' - datos.xlsx');
        listo('');
      });
    },
    /* buffer: ArrayBuffer del archivo → listo(error o null, campos) */
    leer: function (buffer, esq, datos, listo) {
      conXLSX(function (X) {
        if (!X) { listo('No se pudo leer el Excel. Probá de nuevo.'); return; }
        var r;
        try { r = leerHojas(X, X.read(new Uint8Array(buffer), { type: 'array' }), esq, datos); }
        catch (e) { listo('Ese archivo no se pudo abrir como Excel (.xlsx).'); return; }
        if (r.err.length) { listo(r.err.slice(0, 8).join('\n')); return; }
        listo(null, r.campos);
      });
    }
  };
})();

/* ═════════════════════════════════════════════════════════════
   PANEL PRAXIS · nuestros bloques dentro de «Página de inicio»
   (02-11-2026, probado en el borrador Baires de Wess)

   Tiendanube guarda cualquier id en su lista de orden: una fila
   `px_<bloque>` entra en `home_order_position_N`, responde 200 y al
   recargar la dibuja sola, con su ojo. Entonces:

     · el orden y el ojo los guarda TN, con su borrador y su publicar;
     · acá sólo se le pone el nombre a la fila, se abre el formulario
       al tocarla y se mueve el bloque en la vista previa;
     · al guardar en TN se guardan también nuestros datos, así hay un
       solo botón. El visitante no recibe los avisos del editor: por
       eso el orden y las variantes viajan en `datos_json`.

   Piel: "panel-praxis": { "bloques": { "<id>": {
     "nombre": "...",
     "variantes": { "pc": [{ "v", "n", "css" }], "cel": [ ... ] },
     "excel": true } } }
   ═════════════════════════════════════════════════════════════ */
(function () {
  var K = C['panel-praxis'];
  if (!K || !K.bloques) { return; }
  var IDS = Object.keys(K.bloques);
  var fila = function (id) { return 'px_' + id.replace(/-/g, '_'); };
  var deFila = function (f) { return IDS.filter(function (id) { return fila(id) === f; })[0]; };

  /* id de la lista de TN → sección de la tienda */
  var SECCION = {
    slider: 'home-slider', informatives: 'banner-services', products: 'home-products-featured',
    sale: 'home-products-sale', 'new': 'home-products-new', categories: 'home-banner-categories',
    institutional: 'home-institutional-message', announcement: 'home-announcement-message',
    welcome: 'home-welcome-message', brands: 'home-brands', main_product: 'home-product-main',
    instafeed: 'home-instagram-feed', main_categories: 'home-categories-featured',
    newsletter: 'home-newsletter', video: 'home-video', promotional: 'home-banner-promotional',
    news_banners: 'home-banner-news', testimonials: 'home-testimonials', modules: 'home-image-text-module'
  };

  var guardado = {};
  try { guardado = K.datos_json ? JSON.parse(K.datos_json) : {}; } catch (e) { guardado = {}; }
  var VARS = guardado.vars || {};

  /* ─── variantes por dispositivo: un atributo por bloque y su CSS ─── */
  var cssVar = '';
  IDS.forEach(function (id) {
    var v = K.bloques[id].variantes || {};
    ['pc', 'cel'].forEach(function (d) {
      (v[d] || []).forEach(function (o) {
        if (!o.css) { return; }
        cssVar += '@media ' + (d === 'pc' ? '(min-width:768px)' : '(max-width:767px)') + '{' +
          o.css.replace(/&/g, '[data-px="' + id + '"][data-pxv-' + d + '="' + o.v + '"]') + '}';
      });
    });
  });
  if (cssVar) { PX.estilo('panel-praxis-var', cssVar); }
  var nodoDe = function (id) { return document.querySelector('[data-px="' + id + '"]'); };
  var pintarVar = function (id) {
    var n = nodoDe(id), v = VARS[id] || {};
    if (!n) { return; }
    ['pc', 'cel'].forEach(function (d) {
      if (v[d]) { n.setAttribute('data-pxv-' + d, v[d]); } else { n.removeAttribute('data-pxv-' + d); }
    });
  };

  /* ─── el lugar del bloque: debajo de la sección visible anterior ─── */
  var ubicar = function (secciones) {
    secciones.forEach(function (s, i) {
      var id = deFila(s.sectionId), n = id && nodoDe(id);
      if (!n) { return; }
      n.style.display = s.show ? '' : 'none';
      if (!s.show) { return; }
      for (var j = i - 1; j >= 0; j--) {
        if (!secciones[j].show) { continue; }
        var ant = null, previo = deFila(secciones[j].sectionId);
        if (previo) { ant = nodoDe(previo); }
        else if (SECCION[secciones[j].sectionId]) {
          ant = Array.prototype.filter.call(document.querySelectorAll('[data-store="' + SECCION[secciones[j].sectionId] + '"]'),
            function (e) { return e.offsetHeight > 0; })[0];
        }
        if (ant && ant.parentNode) {
          if (ant.nextSibling !== n) { ant.parentNode.insertBefore(n, ant.nextSibling); }
          return;
        }
      }
    });
  };

  /* ─── en la tienda: se aplica lo guardado y listo ─── */
  var enTienda = function () {
    IDS.forEach(pintarVar);
    if (guardado.orden && /template-home/.test(document.body.className)) { ubicar(guardado.orden); }
  };
  if (!PX.enEditor) {
    enTienda();
    [600, 1500, 3000].forEach(function (t) { setTimeout(enTienda, t); });
    return;
  }

  /* ═══════════ de acá para abajo, sólo dentro del editor ═══════════ */
  var P, PW;
  try { PW = window.parent; P = PW.document; } catch (e) { return; }
  var LISTA = '.sortable-sections';
  if (PW.__pxPanelFin) { PW.__pxPanelFin(); }  /* la vista previa se recargó */

  var lista = function () { return P.querySelector(LISTA); };
  var sucios = {};
  var hayCambios = false;
  var avisarTN = function () {
    /* TN habilita «Guardar» con su propio aviso de cambio */
    var L = lista(), k = L && Object.keys(L).filter(function (x) { return /^Sortable/.test(x); })[0];
    var r = L && L.querySelector('.section-order-row');
    if (k && r) { L[k].options.onUpdate({ item: r }); }
  };
  var esquemaDe = function (id) { return (PX.esquemas && PX.esquemas[id]) || K.bloques[id].esquema || { nombre: K.bloques[id].nombre, campos: [] }; };
  var redibujar = function (id) {
    if (PX.rehacer) { PX.rehacer(id); }
    pintarVar(id);
  };
  var leerOrden = function () {
    var L = lista();
    if (!L) { return []; }
    return Array.prototype.map.call(L.children, function (r) { return { sectionId: r.getAttribute('data-id'), show: r.getAttribute('data-show') === 'true' }; });
  };

  /* ─── estilos: los toma de TN (Geist, .form-control) ─── */
  var css = P.createElement('style');
  css.id = 'pxp-css';
  css.textContent =
    '#pxp{position:absolute;inset:0;z-index:30;display:flex;flex-direction:column;background:#fff;color:#333;font-size:14px}' +
    '#pxp-cab{display:flex;align-items:center;gap:8px;min-height:44px;padding:0 12px;border-bottom:1px solid #eee}' +
    '#pxp-cab button{width:32px;height:32px;border:0;background:none;cursor:pointer;font-size:18px;color:#333}' +
    '#pxp-cuerpo{flex:1;min-height:0;overflow-y:auto;padding:14px 16px 24px}' +
    '#pxp .pxp-c{margin-bottom:14px}#pxp label.l{display:flex;justify-content:space-between;gap:8px;margin-bottom:4px;font-size:12px;font-weight:400;color:#5d5d5d}' +
    '#pxp .pxp-n{color:#8a8a8a;font-variant-numeric:tabular-nums}#pxp .pxp-n.pasa{color:#b45309}' +
    '#pxp textarea.form-control{height:auto;min-height:84px;resize:vertical}' +
    '#pxp h4{margin:18px 0 8px;font-size:13px;font-weight:600}' +
    '#pxp .pxp-g{margin-bottom:10px;padding:10px;border:1px solid #e5e5e5;border-radius:8px}' +
    '#pxp .pxp-gt{display:flex;align-items:center;gap:4px;margin-bottom:8px;font-size:12px;font-weight:600}' +
    '#pxp .pxp-gt span{margin-right:auto}' +
    '#pxp .pxp-b{min-height:30px;min-width:30px;padding:0 10px;border:1px solid #ccc;border-radius:8px;background:#fff;color:#333;cursor:pointer;font-size:13px}' +
    '#pxp .pxp-b:disabled{opacity:.35;cursor:not-allowed}' +
    '#pxp .pxp-var{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px}' +
    '#pxp .pxp-var button[aria-pressed=true]{border-color:#006bc8;background:#e8f2fb;color:#004f94}' +
    '#pxp .pxp-acc{display:flex;flex-wrap:wrap;gap:8px}' +
    '#pxp .pxp-msg{margin:8px 0 0;font-size:12px;white-space:pre-line;color:#5d5d5d}#pxp .pxp-msg.mal{color:#b91c1c}' +
    '.section-order-row[data-id^="px_"] .section-order-name{cursor:pointer}';
  P.head.appendChild(css);

  /* ─── las filas ─── */
  var nombrar = function () {
    var L = lista();
    if (!L) { return; }
    IDS.forEach(function (id) {
      var r = L.querySelector('[data-id="' + fila(id) + '"] .section-order-name');
      if (r && r.textContent !== K.bloques[id].nombre) { r.textContent = K.bloques[id].nombre; }
    });
  };
  var ojoDe = function (r) { return r.lastElementChild; };
  /* una fila recién sumada no tiene el ojo conectado hasta que TN la dibuja */
  var cambiarOjo = function (r) {
    var prender = r.getAttribute('data-show') !== 'true';
    var modelo = lista().querySelector('.section-order-row:not([data-pxp-nueva])[data-show="' + prender + '"]');
    r.setAttribute('data-show', String(prender));
    if (modelo) { r.className = modelo.className; ojoDe(r).innerHTML = ojoDe(modelo).innerHTML; ojoDe(r).className = ojoDe(modelo).className; }
    hayCambios = true;
    avisarTN();
  };
  var alCliquear = function (e) {
    var r = e.target.closest && e.target.closest('.section-order-row[data-id^="px_"]');
    if (!r) { return; }
    if (e.target.closest('.section-order-name')) { abrir(deFila(r.getAttribute('data-id'))); return; }
    if (r.hasAttribute('data-pxp-nueva') && !e.target.closest('.section-order-handle')) { e.stopPropagation(); cambiarOjo(r); }
  };
  var asegurarFilas = function () {
    var L = lista();
    if (!L) { return false; }
    var modelo = L.querySelector('.section-order-row[data-show="true"]');
    if (!modelo) { return false; }
    IDS.forEach(function (id) {
      if (L.querySelector('[data-id="' + fila(id) + '"]')) { return; }
      var r = modelo.cloneNode(true);
      r.setAttribute('data-id', fila(id));
      r.setAttribute('data-pxp-nueva', '1');
      /* va después de la última sección prendida */
      var prendidas = L.querySelectorAll('.section-order-row[data-show="true"]');
      L.insertBefore(r, prendidas[prendidas.length - 1].nextSibling);
      hayCambios = true;
    });
    nombrar();
    return true;
  };

  /* ─── el formulario, en el mismo lugar del panel ─── */
  var cont = null;
  var cerrar = function () { if (cont) { cont.remove(); cont = null; } };
  var el = function (tag, cls, txt) {
    var n = P.createElement(tag);
    if (cls) { n.className = cls; }
    if (txt != null) { n.textContent = txt; }
    return n;
  };
  var cambio = function (id) { sucios[id] = true; hayCambios = true; avisarTN(); };
  var campoTexto = function (c, obj, id) {
    var w = el('div', 'pxp-c'), lab = el('label', 'l', c.t);
    var v = obj[c.k] == null ? '' : String(obj[c.k]);
    var cuenta = el('span', 'pxp-n');
    var ctl = el(c.tipo === 'area' ? 'textarea' : 'input', 'form-control');
    if (c.tipo !== 'area') { ctl.type = 'text'; }
    if (c.max) { ctl.maxLength = c.max; }
    ctl.value = v;
    var contar = function () {
      var tope = c.ideal || c.max;
      cuenta.textContent = tope ? ctl.value.length + '/' + tope : '';
      cuenta.classList.toggle('pasa', !!c.ideal && ctl.value.length > c.ideal);
    };
    contar();
    ctl.addEventListener('input', function () {
      obj[c.k] = c.tipo === 'numero' ? ctl.value.replace(',', '.') : ctl.value;
      contar(); cambio(id); redibujar(id);
    });
    lab.appendChild(cuenta);
    w.appendChild(lab); w.appendChild(ctl);
    return w;
  };
  var pintarLista = function (c, datos, id, caja) {
    caja.textContent = '';
    var items = datos[c.k] = datos[c.k] || [];
    items.forEach(function (it, i) {
      var g = el('div', 'pxp-g'), t = el('div', 'pxp-gt');
      t.appendChild(el('span', '', (c.singular || 'Ítem') + ' ' + (i + 1)));
      [['↑', -1], ['↓', 1], ['×', 0]].forEach(function (b) {
        var x = el('button', 'pxp-b', b[0]);
        x.type = 'button';
        x.setAttribute('aria-label', b[1] === -1 ? 'Subir' : b[1] === 1 ? 'Bajar' : 'Sacar');
        if ((b[1] === -1 && i === 0) || (b[1] === 1 && i === items.length - 1)) { x.disabled = true; }
        x.addEventListener('click', function () {
          if (b[1]) { items.splice(i + b[1], 0, items.splice(i, 1)[0]); } else { items.splice(i, 1); }
          cambio(id); redibujar(id); pintarLista(c, datos, id, caja);
        });
        t.appendChild(x);
      });
      g.appendChild(t);
      (c.item || []).forEach(function (s) { if (/^(texto|area|numero)$/.test(s.tipo)) { g.appendChild(campoTexto(s, it, id)); } });
      caja.appendChild(g);
    });
    if (!c.max || items.length < c.max) {
      var mas = el('button', 'pxp-b', c.nuevo || 'Agregar');
      mas.type = 'button';
      mas.addEventListener('click', function () {
        var n = {};
        (c.item || []).forEach(function (s) { n[s.k] = ''; });
        items.push(n);
        cambio(id); redibujar(id); pintarLista(c, datos, id, caja);
      });
      caja.appendChild(mas);
    }
  };
  var variantes = function (id, cuerpo) {
    var V = K.bloques[id].variantes;
    if (!V) { return; }
    [['pc', 'Diseño en computadora'], ['cel', 'Diseño en celular']].forEach(function (d) {
      if (!V[d[0]]) { return; }
      cuerpo.appendChild(el('h4', '', d[1]));
      var g = el('div', 'pxp-var');
      var marcar = function () {
        var hoy = (VARS[id] || {})[d[0]] || V[d[0]][0].v;
        Array.prototype.forEach.call(g.children, function (b) { b.setAttribute('aria-pressed', String(b.getAttribute('data-v') === hoy)); });
      };
      V[d[0]].forEach(function (o) {
        var b = el('button', 'pxp-b', o.n);
        b.type = 'button';
        b.setAttribute('data-v', o.v);
        b.addEventListener('click', function () {
          VARS[id] = VARS[id] || {};
          VARS[id][d[0]] = o.v;
          marcar(); pintarVar(id); hayCambios = true; avisarTN();
        });
        g.appendChild(b);
      });
      marcar();
      cuerpo.appendChild(g);
    });
  };
  var seccionExcel = function (id, cuerpo, datos, rehacerForm) {
    if (!K.bloques[id].excel || !PX.excel) { return; }
    cuerpo.appendChild(el('h4', '', 'Datos por Excel'));
    var acc = el('div', 'pxp-acc'), msg = el('p', 'pxp-msg');
    var bajar = el('button', 'pxp-b', 'Descargar Excel');
    bajar.type = 'button';
    bajar.addEventListener('click', function () {
      msg.className = 'pxp-msg'; msg.textContent = 'Preparando el Excel…';
      PX.excel.bajar(esquemaDe(id), datos, function (e) { msg.className = 'pxp-msg' + (e ? ' mal' : ''); msg.textContent = e || 'Listo: el Excel se descargó.'; });
    });
    var lab = el('label', 'pxp-b', 'Subir Excel');
    lab.style.display = 'inline-flex'; lab.style.alignItems = 'center';
    var inp = el('input');
    inp.type = 'file'; inp.accept = '.xlsx,.xls'; inp.hidden = true;
    lab.appendChild(inp);
    inp.addEventListener('change', function () {
      var a = inp.files && inp.files[0];
      inp.value = '';
      if (!a) { return; }
      var fr = new FileReader();
      fr.onload = function () { PX.panelImportar(id, fr.result, msg, rehacerForm); };
      fr.readAsArrayBuffer(a);
    });
    acc.appendChild(bajar); acc.appendChild(lab);
    cuerpo.appendChild(acc); cuerpo.appendChild(msg);
  };
  PX.panelImportar = function (id, buffer, msg, listo) {
    var datos = C[id] = C[id] || {};
    PX.excel.leer(buffer, esquemaDe(id), datos, function (e, campos) {
      if (e) { if (msg) { msg.className = 'pxp-msg mal'; msg.textContent = 'Hay que corregir esto en el Excel:\n' + e; } return; }
      Object.keys(campos).forEach(function (k) { datos[k] = campos[k]; });
      cambio(id); redibujar(id);
      if (listo) { listo(); }
    });
  };
  var abrir = function (id) {
    if (!id) { return; }
    cerrar();
    var col = lista().closest('.sidebar') || P.body;
    if (getComputedStyle(col).position === 'static') { col.style.position = 'relative'; }
    cont = el('div'); cont.id = 'pxp';
    var cab = el('div'); cab.id = 'pxp-cab';
    var atras = el('button', '', '‹');
    atras.type = 'button'; atras.setAttribute('aria-label', 'Volver');
    atras.addEventListener('click', cerrar);
    cab.appendChild(atras);
    cab.appendChild(el('span', '', K.bloques[id].nombre));
    var cuerpo = el('div'); cuerpo.id = 'pxp-cuerpo';
    cont.appendChild(cab); cont.appendChild(cuerpo);
    /* el botón «Guardar borrador» de TN tiene que seguir a la vista */
    var pie = P.querySelector('.js-save-draft-btn');
    if (pie) { var alto = col.getBoundingClientRect().bottom - pie.getBoundingClientRect().top + 16; cont.style.bottom = Math.max(alto, 0) + 'px'; }
    col.appendChild(cont);

    var datos = C[id] = C[id] || {};
    var rehacerForm = function () { abrir(id); };
    variantes(id, cuerpo);
    (esquemaDe(id).campos || []).forEach(function (c) {
      if (c.tipo === 'lista') {
        cuerpo.appendChild(el('h4', '', c.t));
        var caja = el('div');
        cuerpo.appendChild(caja);
        pintarLista(c, datos, id, caja);
      } else if (/^(texto|area|numero)$/.test(c.tipo)) {
        cuerpo.appendChild(campoTexto(c, datos, id));
      }
    });
    seccionExcel(id, cuerpo, datos, rehacerForm);
    var n = nodoDe(id);
    if (n) { window.scrollTo({ top: n.getBoundingClientRect().top + window.scrollY - 120, behavior: 'auto' }); }
  };
  PX.panelAbrir = abrir;

  /* ─── guardar: el botón de TN guarda también lo nuestro ─── */
  var enviar = PX.panelGuardar || function (bloque, campos) {
    return fetch(C.edicion_url + '/e/' + C.archivo, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ bloque: bloque, campos: campos }) });
  };
  var alGuardar = function () {
    Object.keys(sucios).forEach(function (id) {
      var campos = {};
      (esquemaDe(id).campos || []).forEach(function (c) { if (C[id] && c.k in C[id]) { campos[c.k] = C[id][c.k]; } });
      enviar(id, campos);
    });
    sucios = {};
    K.datos_json = JSON.stringify({ orden: leerOrden(), vars: VARS });
    enviar('panel-praxis', { datos_json: K.datos_json });
    hayCambios = false;
  };
  if (!PW.__pxPanelSend) {
    PW.__pxPanelSend = PW.XMLHttpRequest.prototype.send;
    PW.XMLHttpRequest.prototype.send = function (b) {
      var x = this;
      if (b && b.forEach) {
        var tiene = false;
        b.forEach(function (v, k) { if (/home_order_position_/.test(k)) { tiene = true; } });
        if (tiene) { x.addEventListener('load', function () { if (x.status === 200 && PW.__pxPanelAlGuardar) { PW.__pxPanelAlGuardar(); } }); }
      }
      return PW.__pxPanelSend.call(x, b);
    };
  }
  PW.__pxPanelAlGuardar = alGuardar;

  /* ─── TN avisa cada cambio del orden: se mueve el bloque en vivo ─── */
  var alMensaje = function (e) {
    var d = e.data && e.data.data;
    if (d && d.data && d.data.method === 'updateSections') { nombrar(); ubicar(d.data.args.sections); }
  };
  window.addEventListener('message', alMensaje);

  /* ─── arranque y limpieza ─── */
  var obs = null;
  var arrancar = function () {
    if (!asegurarFilas()) { setTimeout(arrancar, 500); return; }
    obs = new MutationObserver(nombrar);
    obs.observe(lista(), { childList: true, subtree: true, characterData: true });
    lista().addEventListener('click', alCliquear, true);
    IDS.forEach(pintarVar);
    ubicar(leerOrden());
    if (hayCambios) { avisarTN(); }
  };
  PW.__pxPanelFin = function () {
    if (obs) { obs.disconnect(); }
    var L = lista();
    if (L) { L.removeEventListener('click', alCliquear, true); }
    window.removeEventListener('message', alMensaje);
    cerrar();
    var c0 = P.getElementById('pxp-css');
    if (c0) { c0.remove(); }
  };
  arrancar();
})();

})();
