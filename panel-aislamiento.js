/* PRUEBA · NINGÚN BLOQUE ALTERA A OTRO (Panel PRAXIS, 02-11-2026)
   Se pega en la ventana del editor real de TN, en el borrador. No guarda nada: cambia cada
   campo de texto de cada bloque, mira que TODOS los demás sigan idénticos y lo vuelve atrás.
   Mira de cada vecino: su contenido, su lugar en la página y su ancho.
   Resultado: window.__pxpAisla = { ok, cambiados, fallas: [...] } */
(async function () {
  var w = document.querySelector('iframe').contentWindow, d = w.document, PX = w.PX;
  var espera = function (ms) { return new Promise(function (a) { setTimeout(a, ms); }); };
  var cont = d.querySelector('.js-home-sections-container');
  var todos = function () {
    return Array.prototype.filter.call(d.querySelectorAll('[data-px],[data-store]'), function (e) {
      return e.parentElement && e.parentElement.closest('.js-home-sections-container') === cont && e.parentElement === cont;
    });
  };
  var huella = function (e) {
    var cs = w.getComputedStyle(e), r = e.getBoundingClientRect();
    return (e.getAttribute('data-px') || e.getAttribute('data-store')) + '|' + e.textContent.replace(/\s+/g, ' ').trim() + '|' +
      Math.round(r.width) + '|' + cs.display + '|' + cs.backgroundColor + '|' + cs.color;
  };
  /* la clave es el nombre, no la posición: si el bloque probado cambia de largo no corre a los demás */
  var foto = function (fuera) {
    var o = {}, orden = [], cuenta = {};
    todos().forEach(function (e) {
      var k = e.getAttribute('data-px') || e.getAttribute('data-store');
      cuenta[k] = (cuenta[k] || 0) + 1;
      orden.push(k);
      if (k !== fuera) { o[k + '#' + cuenta[k]] = huella(e); }
    });
    return { o: o, orden: orden.filter(function (k) { return k !== fuera; }).join(',') };
  };
  /* Lo que cambia solo (textos que rotan, carruseles) no cuenta como alteración: se mira
     antes de tocar nada y de esos bloques sólo se compara ancho y orden. */
  var ruido = {};
  var marcarRuido = function (a, b) { Object.keys(a.o).forEach(function (k) { if (a.o[k] !== b.o[k]) { ruido[k] = true; } }); };
  var sinTexto = function (h) { var p = h.split('|'); return p[0] + '|' + p.slice(2).join('|'); };
  var R = { ok: true, cambiados: 0, fallas: [], ruido: [] };
  var f0 = foto();
  for (var n = 0; n < 4; n++) { await espera(1500); marcarRuido(f0, foto()); }
  R.ruido = Object.keys(ruido);
  var filas = Array.prototype.filter.call(document.querySelectorAll('.section-order-row.pxp-nuestra'), function () { return true; });
  var ids = filas.map(function (f) {
    var v = f.getAttribute('data-id');
    return v === 'slider' ? 'banner' : v.replace(/^px_/, '').replace(/_/g, '-');
  }).filter(function (id) { return d.querySelector('[data-px="' + id + '"]'); });
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    PX.panelAbrir(id);
    await espera(100);
    if (document.querySelector('#pxp input[type=password]')) { throw new Error('falta la clave de la tienda'); }
    var campos = Array.prototype.filter.call(document.querySelectorAll('#pxp input.form-control[type=text],#pxp textarea.form-control'), function (c) {
      return !/^https?:|^\//.test(c.value) && c.value.trim();
    });
    for (var j = 0; j < campos.length; j++) {
      var c = campos[j], antes = foto(id), ori = c.value;
      c.value = 'Prueba ' + (i + 1) + '-' + (j + 1);
      c.dispatchEvent(new Event('input', { bubbles: true }));
      await espera(450);
      R.cambiados++;
      var luego = foto(id);
      if (luego.orden !== antes.orden) { R.ok = false; R.fallas.push(id + ' campo ' + (j + 1) + ' cambió el orden de los demás'); }
      Object.keys(antes.o).forEach(function (k) {
        var a = antes.o[k], b = luego.o[k];
        if (ruido[k]) { a = sinTexto(a || ''); b = sinTexto(b || ''); }
        if (a !== b) { R.ok = false; R.fallas.push(id + ' campo ' + (j + 1) + ' alteró ' + k.split('#')[0]); }
      });
      c.value = ori;
      c.dispatchEvent(new Event('input', { bubbles: true }));
      await espera(350);
    }
    var cab = document.querySelector('#pxp-cab button');
    if (cab) { cab.click(); }
  }
  R.bloques = ids;
  /* sin campos cambiados no se probó nada: eso nunca es un OK */
  if (!R.cambiados) { R.ok = false; R.fallas.push('no se cambió ningún campo: la prueba no probó nada'); }
  window.__pxpAisla = R;
})().catch(function (e) { window.__pxpAisla = { ok: false, error: e.message }; });
