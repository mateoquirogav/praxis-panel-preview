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
  var foto = function (fuera) {
    var o = {}, orden = [];
    todos().forEach(function (e) {
      var k = e.getAttribute('data-px') || e.getAttribute('data-store');
      orden.push(k);
      if (k !== fuera) { o[k + '#' + orden.length] = huella(e); }
    });
    return { o: o, orden: orden.filter(function (k) { return k !== fuera; }).join(',') };
  };
  var R = { ok: true, cambiados: 0, fallas: [] };
  var filas = Array.prototype.filter.call(document.querySelectorAll('.section-order-row.pxp-nuestra'), function () { return true; });
  var ids = filas.map(function (f) {
    var v = f.getAttribute('data-id');
    return v === 'slider' ? 'banner' : v.replace(/^px_/, '').replace(/_/g, '-');
  }).filter(function (id) { return d.querySelector('[data-px="' + id + '"]'); });
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    PX.panelAbrir(id);
    await espera(100);
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
      if (luego.orden !== antes.orden) { R.ok = false; R.fallas.push(id + ' cambió el orden de los demás'); }
      Object.keys(antes.o).forEach(function (k) {
        if (antes.o[k] !== luego.o[k]) { R.ok = false; R.fallas.push(id + ' campo ' + (j + 1) + ' alteró ' + k.split('#')[0]); }
      });
      c.value = ori;
      c.dispatchEvent(new Event('input', { bubbles: true }));
      await espera(350);
    }
    var cab = document.querySelector('#pxp-cab button');
    if (cab) { cab.click(); }
  }
  R.bloques = ids;
  window.__pxpAisla = R;
})().catch(function (e) { window.__pxpAisla = { ok: false, error: e.message }; });
