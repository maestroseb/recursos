/* Efectivos Junta de Andalucía – Extractor bookmarklet v1.0
 * Todo por fetch, sin recargas. Compatible bookmarklet. */
(function() {
  'use strict';

  if (window.__extractorEfectivos) { alert('El extractor ya está cargado.'); return; }
  window.__extractorEfectivos = true;

  let cancelado = false;

  /*** UTILIDADES ***/
  async function fetchDocGET(url) {
    const res = await fetch(url, { credentials: 'include' });
    const buf = await res.arrayBuffer();
    const html = new TextDecoder('iso-8859-1').decode(buf);
    return new DOMParser().parseFromString(html, 'text/html');
  }

  async function fetchDocPOST(url, params) {
    const body = new URLSearchParams(params).toString();
    const res = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body
    });
    const buf = await res.arrayBuffer();
    const html = new TextDecoder('iso-8859-1').decode(buf);
    return new DOMParser().parseFromString(html, 'text/html');
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
        const parts = td.innerHTML.split(/<br\s*\/?/i)
          .map(line => line.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ')
            .replace(/&gt;|&lt;|>|</g, '').replace(/\s+/g, ' ').trim())
          .filter(Boolean);
        while (parts.length < 4) parts.push('');
        return parts.slice(0, 4);
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
      centro_nombre: '', centro_localidad: '', centro_provincia: ''
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
      const doc = await fetchDocGET(href);
      const colectivo = getField(doc, 'Colectivo:')
        .replace(/\(\s*/, '(').replace(/\)\s*/, ') ').replace(/\s+/g, ' ').trim();
      const centro = getCentroParts(doc);
      return {
        especialidad: basicData.especialidad,
        orden: basicData.orden,
        nif: getField(doc, 'N.I.F.:') || basicData.nif,
        nombre: getField(doc, 'Apellidos y nombre:') || basicData.nombre,
        colectivo: colectivo || basicData.colectivo,
        tiempo_servicio: getTiempoServicio(doc),
        anio_ingreso: getAnio(doc),
        nota: getField(doc, 'Nota ingreso cuerpo:'),
        centro_codigo: centro[0] || basicData.centro_codigo,
        centro_nombre: centro[1],
        centro_localidad: centro[2],
        centro_provincia: centro[3]
      };
    } catch (e) {
      console.error('❌ Error en ficha', href, e);
      return basicData;
    }
  }

  function descargarCSV(data, esp) {
    let nombreArchivo = 'adjudicaciones_completas';
    if (esp.codigo !== 'N/A') {
      nombreArchivo = `${esp.codigo}_${esp.nombre}`
        .replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
    }
    const headers = ['Especialidad','Orden','NIF','Nombre','Colectivo','Tiempo servicio','Año ingreso','Nota','Centro código','Centro nombre','Centro localidad','Centro provincia'];
    const rows = [headers.join(';')];
    data.forEach(d => {
      rows.push([d.especialidad,d.orden,d.nif,d.nombre,d.colectivo,d.tiempo_servicio,
        d.anio_ingreso,d.nota,d.centro_codigo,d.centro_nombre,d.centro_localidad,d.centro_provincia]
        .map(s => `"${String(s).replace(/"/g,'""')}"`).join(';'));
    });
    const blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=iso-8859-1' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${nombreArchivo}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    console.log(`✅ CSV descargado: ${nombreArchivo}.csv con ${data.length} registros`);
  }

  /*** PROCESO PRINCIPAL: todo en memoria, sin recargas ***/
  async function extraerTodo() {
    cancelado = false;
    const resultados = [];
    const esp = getEspecialidad(document);
    const pag = getPaginacion(document);
    console.log(`📚 Especialidad: ${esp.codigo} | Páginas: ${pag.total}`);

    for (let pagina = 1; pagina <= pag.total; pagina++) {
      if (cancelado) break;

      // Página 1 = documento actual; las demás se piden por POST
      let doc;
      if (pagina === 1) {
        doc = document;
      } else {
        btn.textContent = `⏳ ${esp.codigo} - Cargando página ${pagina}/${pag.total}...`;
        doc = await fetchDocPOST(location.href, {
          P: String(pagina), APA: 'SI', PUESTO: pag.puesto
        });
      }

      const enlaces = recogerEnlaces(doc, esp.codigo);
      console.log(`📄 Página ${pagina}/${pag.total} - ${enlaces.length} enlaces`);

      if (enlaces.length === 0) {
        console.log('📭 Página sin enlaces - deteniendo');
        break;
      }

      for (let i = 0; i < enlaces.length; i++) {
        if (cancelado) break;
        const { href, basicData } = enlaces[i];
        btn.textContent = `⏳ ${esp.codigo} - Pág ${pagina}/${pag.total} - Ficha ${i+1}/${enlaces.length} (Total: ${resultados.length + 1})`;
        const ficha = await procesarFicha(href, basicData);
        resultados.push(ficha);
        await new Promise(r => setTimeout(r, 10));
      }
    }

    // Fin (normal o por cancelación)
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
    btn.disabled = true;
    btn.style.background = '#ffc107';
    btn.textContent = '⏳ Iniciando...';
    extraerTodo();
  });

  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = '🛑 Cancelar';
  Object.assign(cancelBtn.style, {
    position:'fixed', top:'75px', right:'20px', zIndex:9999, padding:'8px',
    background:'#dc3545', color:'#fff', border:'none', borderRadius:'4px',
    cursor:'pointer', fontSize:'12px'
  });
  document.body.appendChild(cancelBtn);

  cancelBtn.addEventListener('click', () => {
    cancelado = true;
    console.log('🛑 Cancelación solicitada');
  });

})();
