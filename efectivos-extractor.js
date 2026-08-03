/* Efectivos Junta de Andalucía – Extractor bookmarklet v2.0
 * Todo por fetch, sin recargas. Compatible bookmarklet.
 * Novedades v2.0:
 *   - Reintentos automáticos con backoff y timeout (no se bloquea).
 *   - Pausa/Reanudación: continúa por donde se quedó.
 *   - Bloque Centro robusto ante 4 o 5 líneas: la provincia es SIEMPRE la última
 *     línea y la localidad la 3ª (la Junta pasó de 4 a 5 líneas y volvió a 4). */
(function() {
  'use strict';

  if (window.__extractorEfectivos) { alert('El extractor ya está cargado.'); return; }
  window.__extractorEfectivos = true;

  /*** ESTADO ***/
  let cancelado = false;
  let pausado = false;
  let resumeResolve = null;   // resuelve la promesa que mantiene la ejecución en pausa
  let enEjecucion = false;

  const TIMEOUT_MS = 30000;   // aborta una petición colgada a los 30 s
  const MAX_INTENTOS = 4;     // reintentos por petición (backoff 1s, 2s, 4s, 8s)

  /*** UTILIDADES ***/
  // Espera aquí mientras el proceso esté en pausa (permite reanudar por donde se quedó).
  async function esperarSiPausado() {
    if (pausado) {
      await new Promise(r => { resumeResolve = r; });
    }
  }

  // fetch de un documento con timeout y reintentos con backoff exponencial.
  async function fetchDoc(url, options = {}) {
    let ultimoError;
    for (let intento = 1; intento <= MAX_INTENTOS; intento++) {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
      try {
        const res = await fetch(url, { credentials: 'include', signal: ctrl.signal, ...options });
        clearTimeout(timer);
        if (!res.ok) throw new Error('HTTP ' + res.status);
        const buf = await res.arrayBuffer();
        const html = new TextDecoder('iso-8859-1').decode(buf);
        return new DOMParser().parseFromString(html, 'text/html');
      } catch (e) {
        clearTimeout(timer);
        ultimoError = e;
        console.warn(`⚠️ Intento ${intento}/${MAX_INTENTOS} falló para ${url}:`, e.message || e);
        if (intento < MAX_INTENTOS) {
          const espera = 1000 * Math.pow(2, intento - 1);   // 1s, 2s, 4s, 8s
          await new Promise(r => setTimeout(r, espera));
        }
      }
    }
    throw ultimoError;
  }

  function fetchDocGET(url) {
    return fetchDoc(url);
  }

  function fetchDocPOST(url, params) {
    return fetchDoc(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(params).toString()
    });
  }

  function getField(doc, starts) {
    for (const th of doc.querySelectorAll('th')) {
      const txt = th.textContent.replace(/\u00a0/g, ' ').trim();
      if (txt.startsWith(starts)) {
        let value = th.nextElementSibling?.textContent.replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim() || '';
        if (starts.includes('Tiempo efectivo') || starts.includes('Tiempo de servicio')) {
          value = value.replace(/\s*\([^)]*\)\s*/g, '');
          if (value && /^\d/.test(value)) value = `'${value}`;
        }
        return value;
      }
    }
    return '';
  }

  function getTiempoServicio(doc) {
    return getField(doc, 'Tiempo efectivo de servicios como funcionario:')
        || getField(doc, 'Tiempo de servicio:');
  }

  function getCentroParts(doc) {
    for (const th of doc.querySelectorAll('th')) {
      if (th.textContent.replace(/\u00a0/g, ' ').trim() === 'Centro') {
        const td = th.nextElementSibling;
        if (!td) return ['', '', '', ''];
        const parts = td.innerHTML.split(/<br\s*\/?>?/i)
          .map(line => line.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ')
            .replace(/&gt;|&lt;|>|</g, '').replace(/\s+/g, ' ').trim())
          .filter(Boolean);

        // Estructura del bloque Centro (compatible con ambos formatos, por si la
        // Consejería corrige el error en cualquier momento):
        //   4 líneas (correcto):    [código, nombre, localidad, provincia]
        //   5 líneas (error 2026):  [código, nombre, localidad, <línea errónea>, provincia]
        // Regla robusta: la provincia es SIEMPRE la última línea y la localidad la 3ª;
        // cualquier línea intermedia entre ambas se descarta. Así el mapeo es correcto
        // tanto si hay 4 como 5 líneas, sin necesidad de tocar el código.
        const codigo = parts[0] || '';
        const nombre = parts[1] || '';
        const localidad = parts[2] || '';
        const provincia = parts.length > 3 ? parts[parts.length - 1] : '';

        if (parts.length > 4) {
          console.log('ℹ️ Centro con líneas extra descartadas:', parts.slice(3, parts.length - 1));
        }
        return [codigo, nombre, localidad, provincia];
      }
    }
    return ['', '', '', ''];
  }

  function getAnio(doc) {
    for (const row of doc.querySelectorAll('tr')) {
      const th = row.querySelector('th');
      const td = row.querySelector('td');
      if (!th || !td) continue;
      if (th.textContent.replace(/\u00a0/g, ' ').trim().includes('Año de ingreso en el cuerpo')) {
        return td.textContent.replace(/\u00a0/g, ' ').trim();
      }
    }
    return '';
  }

  function getEspecialidad(doc) {
    const h3 = doc.querySelector('h3.text-center');
    if (h3) {
      const m = h3.textContent.trim().match(/Puesto:\s*\(([^)]+)\)\s*(.+)/);
      if (m) return { codigo: m[1].trim(), nombre: m[2].trim() };
    }
    return { codigo: 'N/A', nombre: '' };
  }

  function getPaginacion(doc) {
    // "Pág. 1 / 4."
    const m = doc.body.textContent.match(/Pág\.\s*(\d+)\s*\/\s*(\d+)/i);
    const current = m ? parseInt(m[1], 10) : 1;
    const total = m ? parseInt(m[2], 10) : 1;
    // Token PUESTO necesario para el POST de paginación
    const puestoInput = doc.querySelector('input[name="PUESTO"]');
    const puesto = puestoInput ? puestoInput.value : '';
    return { current, total, puesto };
  }

  function extractBasicData(tr, orden, especialidadCode) {
    const cells = tr.querySelectorAll('td');
    if (cells.length < 4) return null;
    const colectivo = (cells[1]?.textContent?.trim() || '')
      .replace(/\(\s*/, '(').replace(/\)\s*/, ') ').replace(/\s+/g, ' ').trim();
    let nif = '', nombre = '';
    const t = cells[2]?.textContent.trim() || '';
    const m = t.match(/\(([^)]+)\)\s*(.+)/);
    if (m) { nif = m[1]; nombre = m[2]; }
    return {
      especialidad: especialidadCode, orden, nif, nombre, colectivo,
      tiempo_servicio: '', anio_ingreso: '', nota: '',
      centro_codigo: cells[3]?.textContent?.trim() || '',
      prov: '',                         // (J) provincia
      centro_localidad: ''              // (K) "Nombre del centro, Localidad"
    };
  }

  function recogerEnlaces(doc, especialidadCode) {
    return Array.from(doc.querySelectorAll('#example tr'))
      .filter(tr => tr.querySelector('a[href*="idemp="]'))
      .map(tr => {
        const orden = tr.querySelector('td')?.textContent.trim() || '';
        const href = new URL(tr.querySelector('a[href*="idemp="]').href, location).href;
        return { orden, href, basicData: extractBasicData(tr, orden, especialidadCode) };
      })
      .filter(x => x.basicData !== null);
  }

  async function procesarFicha(href, basicData) {
    try {
      const doc = await fetchDocGET(href);   // ya reintenta internamente
      const colectivo = getField(doc, 'Colectivo:')
        .replace(/\(\s*/, '(').replace(/\)\s*/, ') ').replace(/\s+/g, ' ').trim();
      // centro = [código, nombre, localidad, provincia]
      const centro = getCentroParts(doc);
      const centroYLocalidad = [centro[1], centro[2]].filter(Boolean).join(', ');
      return {
        especialidad: basicData.especialidad,
        orden: basicData.orden,
        nif: getField(doc, 'N.I.F.:') || basicData.nif,
        nombre: getField(doc, 'Apellidos y nombre:') || basicData.nombre,
        colectivo: colectivo || basicData.colectivo,
        tiempo_servicio: getTiempoServicio(doc),
        anio_ingreso: getAnio(doc),
        nota: getField(doc, 'Nota ingreso cuerpo:'),
        centro_codigo: centro[0] || basicData.centro_codigo,   // (I)
        prov: centro[3],                              // (J) provincia (última línea)
        centro_localidad: centroYLocalidad            // (K) "Nombre, Localidad"
      };
    } catch (e) {
      // Tras agotar los reintentos, no perdemos la fila: devolvemos los datos básicos.
      console.error('❌ Error en ficha (se conservan datos básicos)', href, e);
      return basicData;
    }
  }

  function descargarCSV(data, esp) {
    let nombreArchivo = 'adjudicaciones_completas';
    if (esp.codigo !== 'N/A') {
      nombreArchivo = `${esp.codigo}_${esp.nombre}`
        .replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
    }
    // Columnas A–K del pipeline de la hoja de efectivos (espejo del CGT):
    // A Especialidad B Orden C NIF D Nombre E Colectivo
    // F Tiempo servicio G Año ingreso H Nota I Centro código J PROV K CentroyLocalidad
    const headers = ['Especialidad','Orden','NIF','Nombre','Colectivo','Tiempo servicio','Año ingreso','Nota','Centro código','PROV','CentroyLocalidad'];
    const rows = [headers.join(';')];
    data.forEach(d => {
      rows.push([d.especialidad,d.orden,d.nif,d.nombre,d.colectivo,
        d.tiempo_servicio,d.anio_ingreso,d.nota,d.centro_codigo,d.prov,d.centro_localidad]
        .map(s => `"${String(s).replace(/"/g,'""')}"`).join(';'));
    });
    // El constructor Blob SIEMPRE serializa el texto en UTF-8, así que exportamos
    // como UTF-8 con BOM para que ñ y tildes lleguen bien tanto al
    // importador de la hoja como a Excel al abrir el CSV directamente.
    const blob = new Blob(['\ufeff' + rows.join('\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${nombreArchivo}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    console.log(`✅ CSV descargado: ${nombreArchivo}.csv con ${data.length} registros`);
  }

  /*** PROCESO PRINCIPAL: todo en memoria, sin recargas ***/
  // Estado persistente entre pausas para poder reanudar por donde se quedó.
  const estado = {
    resultados: [],
    esp: null,
    pag: null,
    pagina: 1,        // página en curso
    indiceFicha: 0,   // índice de ficha dentro de la página en curso
    enlaces: null     // enlaces de la página en curso (se cachean para reanudar)
  };

  async function extraerTodo() {
    cancelado = false;
    enEjecucion = true;

    if (!estado.esp) {
      estado.resultados = [];
      estado.esp = getEspecialidad(document);
      estado.pag = getPaginacion(document);
      estado.pagina = 1;
      estado.indiceFicha = 0;
      estado.enlaces = null;
      console.log(`📚 Especialidad: ${estado.esp.codigo} | Páginas: ${estado.pag.total}`);
    } else {
      console.log('▶️ Reanudando extracción...');
    }

    const { esp, pag } = estado;

    try {
      for (; estado.pagina <= pag.total; estado.pagina++) {
        if (cancelado) break;
        await esperarSiPausado();
        if (cancelado) break;

        const pagina = estado.pagina;

        // Obtener los enlaces de la página (cacheados para no re-pedirla al reanudar).
        if (!estado.enlaces) {
          let doc;
          if (pagina === 1) {
            doc = document;
          } else {
            btn.textContent = `⏳ ${esp.codigo} - Cargando página ${pagina}/${pag.total}...`;
            doc = await fetchDocPOST(location.href, {
              P: String(pagina), APA: 'SI', PUESTO: pag.puesto
            });
          }
          estado.enlaces = recogerEnlaces(doc, esp.codigo);
          estado.indiceFicha = 0;
          console.log(`📄 Página ${pagina}/${pag.total} - ${estado.enlaces.length} enlaces`);
        }

        if (estado.enlaces.length === 0) {
          console.log('📭 Página sin enlaces - deteniendo');
          break;
        }

        for (; estado.indiceFicha < estado.enlaces.length; estado.indiceFicha++) {
          if (cancelado) break;
          await esperarSiPausado();
          if (cancelado) break;

          const i = estado.indiceFicha;
          const { href, basicData } = estado.enlaces[i];
          btn.textContent = `⏳ ${esp.codigo} - Pág ${pagina}/${pag.total} - Ficha ${i+1}/${estado.enlaces.length} (Total: ${estado.resultados.length + 1})`;
          const ficha = await procesarFicha(href, basicData);
          estado.resultados.push(ficha);
          await new Promise(r => setTimeout(r, 10));
        }

        if (cancelado) break;
        // Preparar la siguiente página.
        estado.enlaces = null;
        estado.indiceFicha = 0;
      }
    } catch (e) {
      console.error('❌ Error irrecuperable en el proceso:', e);
      btn.textContent = `⚠️ Error (${estado.resultados.length} registros). Puedes reanudar.`;
      btn.disabled = false;
      btn.style.background = '#fd7e14';
      enEjecucion = false;
      return;
    }

    // Fin (normal o por cancelación)
    finalizar();
  }

  function finalizar() {
    enEjecucion = false;
    const { resultados, esp } = estado;
    if (resultados.length > 0) {
      if (cancelado) {
        if (confirm(`Cancelado. ¿Descargar los ${resultados.length} registros procesados?`)) {
          descargarCSV(resultados, esp);
        }
        btn.textContent = `🛑 Cancelado (${resultados.length} registros)`;
      } else {
        descargarCSV(resultados, esp);
        btn.textContent = `✅ Completado (${resultados.length} registros)`;
      }
    } else {
      btn.textContent = '⚠️ Sin datos';
    }
    btn.disabled = false;
    btn.style.background = '#28a745';
    // Permitir una nueva extracción desde cero.
    estado.esp = null;
  }

  /*** BOTONES ***/
  const btn = document.createElement('button');
  btn.textContent = '▶ Extraer todas las adjudicaciones';
  Object.assign(btn.style, {
    position:'fixed', top:'20px', right:'20px', zIndex:9999, padding:'10px',
    background:'#28a745', color:'#fff', border:'none', borderRadius:'6px',
    cursor:'pointer', fontWeight:'bold', maxWidth:'340px'
  });
  document.body.appendChild(btn);

  btn.addEventListener('click', () => {
    if (enEjecucion) return;
    btn.disabled = true;
    btn.style.background = '#ffc107';
    btn.textContent = estado.esp ? '▶️ Reanudando...' : '⏳ Iniciando...';
    pauseBtn.disabled = false;
    pauseBtn.textContent = '⏸ Pausar';
    extraerTodo();
  });

  // Botón Pausar / Reanudar
  const pauseBtn = document.createElement('button');
  pauseBtn.textContent = '⏸ Pausar';
  pauseBtn.disabled = true;
  Object.assign(pauseBtn.style, {
    position:'fixed', top:'62px', right:'20px', zIndex:9999, padding:'8px',
    background:'#0d6efd', color:'#fff', border:'none', borderRadius:'4px',
    cursor:'pointer', fontSize:'12px', maxWidth:'340px'
  });
  document.body.appendChild(pauseBtn);

  pauseBtn.addEventListener('click', () => {
    if (!enEjecucion && !pausado) return;
    if (!pausado) {
      // Pausar
      pausado = true;
      pauseBtn.textContent = '▶ Reanudar';
      pauseBtn.style.background = '#198754';
      btn.textContent = `⏸ Pausado (${estado.resultados.length} registros)`;
      console.log('⏸ Pausado');
    } else {
      // Reanudar por donde se quedó
      pausado = false;
      pauseBtn.textContent = '⏸ Pausar';
      pauseBtn.style.background = '#0d6efd';
      console.log('▶ Reanudando');
      if (resumeResolve) { resumeResolve(); resumeResolve = null; }
    }
  });

  // Botón Cancelar
  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = '🛑 Cancelar';
  Object.assign(cancelBtn.style, {
    position:'fixed', top:'104px', right:'20px', zIndex:9999, padding:'8px',
    background:'#dc3545', color:'#fff', border:'none', borderRadius:'4px',
    cursor:'pointer', fontSize:'12px'
  });
  document.body.appendChild(cancelBtn);

  cancelBtn.addEventListener('click', () => {
    cancelado = true;
    // Si estaba en pausa, liberamos la espera para que el proceso finalice.
    if (pausado) {
      pausado = false;
      pauseBtn.textContent = '⏸ Pausar';
      pauseBtn.style.background = '#0d6efd';
      if (resumeResolve) { resumeResolve(); resumeResolve = null; }
    }
    pauseBtn.disabled = true;
    console.log('🛑 Cancelación solicitada');
  });

})();
