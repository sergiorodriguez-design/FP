// ============================================================
// Code.gs — Backend principal de la App de Análisis Editorial
// ============================================================

const CONFIG = {
  OPENAI_MODEL:    'gpt-4.1',
  OPENAI_API_URL:  'https://api.openai.com/v1/chat/completions',
  MAX_CHUNK_CHARS: 14000,
};

// ── Web App entry point ───────────────────────────────────────
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Analizador Editorial FPE')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── API Key ───────────────────────────────────────────────────
function guardarApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('OPENAI_API_KEY', apiKey);
  return { ok: true };
}
function obtenerApiKey() {
  const k = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
  return k ? k.substring(0, 8) + '...' : null;
}
function apiKeyConfigurada() {
  return !!PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
}

// ── Extraer texto de un PDF/DOCX en base64 ───────────────────
function extraerTextoDeArchivo(base64Data, mimeType, nombreArchivo) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data), mimeType, nombreArchivo
    );
    const texto = extraerTextoConConversion(blob);
    if (!texto || texto.length < 50)
      throw new Error('No se pudo extraer texto. Verifica que el PDF tenga texto seleccionable.');
    return { ok: true, texto, longitud: texto.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function extraerTextoConConversion(blob) {
  const nombreTemp = 'fpe_temp_' + Date.now();

  const archivo   = DriveApp.createFile(blob.setName(nombreTemp));
  const archivoId = archivo.getId();
  const token     = ScriptApp.getOAuthToken();

  Utilities.sleep(2500);

  let resp, intentos = 0;
  do {
    if (intentos > 0) Utilities.sleep(5000 * intentos);
    resp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + archivoId + '/copy',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({
          name: nombreTemp + '_gdoc',
          mimeType: 'application/vnd.google-apps.document'
        }),
        muteHttpExceptions: true
      }
    );
    intentos++;
  } while ((resp.getResponseCode() === 429 || resp.getResponseCode() === 403) && intentos < 4);

  try { archivo.setTrashed(true); } catch(_) {}

  if (resp.getResponseCode() !== 200)
    throw new Error('Error convirtiendo a Google Doc (intento ' + intentos + '): ' + resp.getContentText());

  const gdocId = JSON.parse(resp.getContentText()).id;
  Utilities.sleep(3000);

  let texto = '';
  try {
    texto = DocumentApp.openById(gdocId).getBody().getText();
  } finally {
    try { DriveApp.getFileById(gdocId).setTrashed(true); } catch(_) {}
  }
  return texto;
}

// ── Construir prompt para análisis de unidad ─────────────────
function buildPromptUnidad(textoChunk, infoUnidad, numChunk, totalChunks) {
  return [
    {
      role: 'system',
      content: [
        'Eres un experto en análisis de materiales didácticos de Formación Profesional para el Empleo (FPE) en España.',
        'Tu tarea es analizar el texto de una unidad didáctica y extraer de forma estructurada:',
        '1. Los objetivos específicos y criterios de evaluación (capacidades C1, C2, etc.).',
        '2. El mapa de contenidos (conceptuales, procedimentales, actitudinales).',
        '3. Todas las actividades formativas con su identificador, tipo, texto completo y estrategias metodológicas.',
        '',
        'Tipos de actividad válidos:',
        '  - Actividad de Aprendizaje',
        '  - Actividad Colaborativa',
        '  - Tarea de Evaluación',
        '  - Ejercicio de Autoevaluación',
        '',
        'Responde EXCLUSIVAMENTE con un bloque ```json``` con esta estructura:',
        '{',
        '  "objetivos_y_criterios": "texto completo con capacidades y CEs",',
        '  "contenidos": "mapa de contenidos completo",',
        '  "actividades": [',
        '    {',
        '      "identificador": "Ej: Actividad 1 / Tarea 2",',
        '      "tipo": "uno de los 4 tipos válidos",',
        '      "texto_completo": "enunciado completo de la actividad",',
        '      "estrategias_actividades_recursos": "metodología, recursos, temporalización",',
        '      "observaciones": "notas adicionales si las hay"',
        '    }',
        '  ]',
        '}'
      ].join('\n')
    },
    {
      role: 'user',
      content: [
        'UNIDAD: ' + infoUnidad.etiqueta,
        'Archivo: ' + infoUnidad.nombreArchivo,
        'Fragmento ' + numChunk + ' de ' + totalChunks,
        '',
        '--- TEXTO ---',
        textoChunk
      ].join('\n')
    }
  ];
}

// ── Construir prompt para análisis narrativo ──────────────────
function buildPromptNarrativo(actividades, infoUnidad, objetivos_y_criterios, contenidos) {
  const resumenActividades = actividades.map(function(a, i) {
    return (i + 1) + '. [' + (a.tipo || 'Sin tipo') + '] ' + (a.identificador || '') + ': ' + (a.texto_completo || a.estrategias_actividades_recursos || '').substring(0, 200);
  }).join('\n');

  return [
    {
      role: 'system',
      content: [
        'Eres inspector de Formación Profesional para el Empleo (FPE) en España.',
        'Redacta un análisis técnico-pedagógico de la unidad didáctica para un informe de inspección.',
        'El análisis debe incluir:',
        '  1. Valoración de la coherencia entre capacidades, contenidos y actividades.',
        '  2. Adecuación metodológica y variedad de tipología de actividades.',
        '  3. Presencia de actividades colaborativas y de autoevaluación.',
        '  4. Observaciones sobre posibles mejoras o puntos fuertes.',
        'Redacta en tercera persona, tono formal, máximo 400 palabras.'
      ].join('\n')
    },
    {
      role: 'user',
      content: [
        'UNIDAD: ' + infoUnidad.etiqueta,
        '',
        'CAPACIDADES Y CRITERIOS DE EVALUACIÓN:',
        objetivos_y_criterios || '(no extraídos)',
        '',
        'CONTENIDOS:',
        contenidos || '(no extraídos)',
        '',
        'ACTIVIDADES DETECTADAS (' + actividades.length + '):',
        resumenActividades || '(ninguna)'
      ].join('\n')
    }
  ];
}

// ── Analizar una unidad completa ──────────────────────────────
function analizarUnidad(textoUnidad, infoUnidad) {
  try {
    const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) throw new Error('API Key de OpenAI no configurada.');

    const chunks  = dividirEnChunks(textoUnidad, CONFIG.MAX_CHUNK_CHARS);
    let actividades = [];
    let objetivos_y_criterios = '';
    let contenidos = '';

    for (let i = 0; i < chunks.length; i++) {
      const prompt    = buildPromptUnidad(chunks[i], infoUnidad, i + 1, chunks.length);
      const respuesta = llamarOpenAI(apiKey, prompt, 4500);
      const json      = extraerJSON(respuesta);

      if (i === 0) {
        objetivos_y_criterios = json.objetivos_y_criterios || '';
        contenidos            = json.contenidos || '';
      }
      if (json.actividades && json.actividades.length)
        actividades = actividades.concat(json.actividades);
    }

    const ejercicios = actividades.map(function(a) {
      return {
        identificador:                    a.identificador || '',
        tipo:                             a.tipo || '',
        texto_completo:                   a.texto_completo || '',
        objetivos_y_criterios:            objetivos_y_criterios,
        contenidos:                       contenidos,
        estrategias_actividades_recursos: a.estrategias_actividades_recursos || '',
        observaciones:                    a.observaciones || ''
      };
    });

    const analisis = llamarOpenAI(
      apiKey,
      buildPromptNarrativo(actividades, infoUnidad, objetivos_y_criterios, contenidos),
      4000
    );

    return {
      ok: true,
      infoUnidad,
      objetivos_y_criterios,
      contenidos,
      ejercicios,
      analisis,
      total: ejercicios.length
    };
  } catch (e) {
    return { ok: false, error: e.message, infoUnidad };
  }
}

// ── Exportar Google Sheets ────────────────────────────────────
function exportarASheets(nombreCurso, resultados) {
  try {
    const ss   = SpreadsheetApp.create('FPE Informe — ' + nombreCurso + ' — ' + fmtFecha());
    const hRes = ss.getActiveSheet().setName('Resumen');

    hRes.getRange(1,1,1,7).setValues([[
      'Unidad','Archivo','Total actividades',
      'Act. Aprendizaje','Colaborativas','Tareas Evaluación','Autoevaluación'
    ]]).setFontWeight('bold').setBackground('#000f94').setFontColor('#ffffff');

    resultados.forEach(function(r, i) {
      const e = r.ejercicios;
      hRes.getRange(i+2,1,1,7).setValues([[
        r.infoUnidad.etiqueta, r.infoUnidad.nombreArchivo, e.length,
        e.filter(function(x){ return x.tipo==='Actividad de Aprendizaje'; }).length,
        e.filter(function(x){ return x.tipo==='Actividad Colaborativa'; }).length,
        e.filter(function(x){ return x.tipo==='Tarea de Evaluación'; }).length,
        e.filter(function(x){ return x.tipo==='Ejercicio de Autoevaluación'; }).length,
      ]]);
    });
    hRes.autoResizeColumns(1,7);

    resultados.forEach(function(r) {
      const label = r.infoUnidad.etiqueta.replace(/[\/\\?*[\]:]/g,'').substring(0,30);
      const sh    = ss.insertSheet(label);

      sh.getRange(1,1,1,4).merge()
        .setValue(r.infoUnidad.etiqueta + ' — ' + r.infoUnidad.nombreArchivo)
        .setFontSize(13).setFontWeight('bold').setFontColor('#000f94');

      sh.getRange(3,1,1,4).setValues([['CRITERIOS DE EVALUACIÓN Y CAPACIDADES','','','']])
        .setFontWeight('bold').setBackground('#000f94').setFontColor('#13c9f2');
      sh.getRange(4,1,1,4).merge()
        .setValue(r.objetivos_y_criterios || '')
        .setWrap(true).setBackground('#f0f4ff').setFontColor('#000000');

      sh.getRange(6,1,1,4).setValues([['MAPA DE CONTENIDOS','','','']])
        .setFontWeight('bold').setBackground('#000f94').setFontColor('#13c9f2');
      sh.getRange(7,1,1,4).merge()
        .setValue(r.contenidos || '')
        .setWrap(true).setBackground('#f0f4ff').setFontColor('#000000');

      sh.getRange(9,1,1,4).setValues([[
        'Identificador actividad',
        'Objetivos específicos / Criterios de evaluación',
        'Contenidos',
        'Estrategias metodológicas, actividades de aprendizaje y recursos didácticos'
      ]]).setFontWeight('bold').setBackground('#13c9f2').setFontColor('#000f94');

      r.ejercicios.forEach(function(ej, i) {
        sh.getRange(i+10,1,1,4).setValues([[
          ej.identificador || '',
          r.objetivos_y_criterios || '',
          r.contenidos || '',
          ej.estrategias_actividades_recursos || ''
        ]]).setFontColor('#000000');
        if (i % 2 === 0) sh.getRange(i+10,1,1,4).setBackground('#f8fbff');
      });

      const fila = r.ejercicios.length + 12;
      sh.getRange(fila,1,1,4).merge()
        .setValue('ANÁLISIS PARA INSPECCIÓN')
        .setFontWeight('bold').setFontColor('#000f94').setBackground('#e8f8fd');
      sh.getRange(fila+1,1,1,4).merge()
        .setValue(r.analisis)
        .setWrap(true).setFontColor('#000000');

      sh.setColumnWidth(1,200); sh.setColumnWidth(2,280);
      sh.setColumnWidth(3,260); sh.setColumnWidth(4,340);
    });

    return { ok: true, url: ss.getUrl(), nombre: ss.getName() };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ── Exportar Google Doc ───────────────────────────────────────
function exportarADoc(nombreCurso, resultados) {
  try {
    const doc  = DocumentApp.create('FPE Informe — ' + nombreCurso + ' — ' + fmtFecha());
    const body = doc.getBody();
    const A    = DocumentApp.Attribute;

    body.appendParagraph('Informe de Análisis de Actividades FPE')
        .setAttributes({[A.FONT_SIZE]:20,[A.BOLD]:true,[A.FOREGROUND_COLOR]:'#000f94'});
    body.appendParagraph('Curso: ' + nombreCurso + '  |  ' + new Date().toLocaleString('es-ES'))
        .setAttributes({[A.FONT_SIZE]:10,[A.ITALIC]:true,[A.FOREGROUND_COLOR]:'#475569'});
    body.appendHorizontalRule();

    resultados.forEach(function(r) {
      body.appendParagraph(r.infoUnidad.etiqueta)
          .setAttributes({[A.FONT_SIZE]:15,[A.BOLD]:true,[A.FOREGROUND_COLOR]:'#000f94'});
      body.appendParagraph('Archivo: ' + r.infoUnidad.nombreArchivo)
          .setAttributes({[A.FONT_SIZE]:10,[A.ITALIC]:true,[A.FOREGROUND_COLOR]:'#475569'});

      body.appendParagraph('Criterios de evaluación y capacidades')
          .setAttributes({[A.FONT_SIZE]:12,[A.BOLD]:true,[A.FOREGROUND_COLOR]:'#000f94'});
      body.appendParagraph(r.objetivos_y_criterios || '—')
          .setAttributes({[A.FONT_SIZE]:10});

      body.appendParagraph('Mapa de contenidos')
          .setAttributes({[A.FONT_SIZE]:12,[A.BOLD]:true,[A.FOREGROUND_COLOR]:'#000f94'});
      body.appendParagraph(r.contenidos || '—')
          .setAttributes({[A.FONT_SIZE]:10});

      if (r.ejercicios.length) {
        body.appendParagraph('Actividades formativas (' + r.ejercicios.length + ')')
            .setAttributes({[A.FONT_SIZE]:12,[A.BOLD]:true,[A.FOREGROUND_COLOR]:'#000f94'});
        const tabla = body.appendTable();
        const cab   = tabla.appendTableRow();
        ['Identificador / Tipo', 'Estrategias, metodología y recursos']
          .forEach(function(h) {
            cab.appendTableCell(h).setAttributes({
              [A.BOLD]:true,
              [A.BACKGROUND_COLOR]:'#13c9f2',
              [A.FOREGROUND_COLOR]:'#000f94'
            });
          });
        r.ejercicios.forEach(function(ej) {
          const row = tabla.appendTableRow();
          row.appendTableCell((ej.identificador||'') + '\n[' + (ej.tipo||'') + ']')
             .setAttributes({[A.FOREGROUND_COLOR]:'#000000'});
          row.appendTableCell(ej.estrategias_actividades_recursos || '')
             .setAttributes({[A.FOREGROUND_COLOR]:'#000000'});
        });
      }

      body.appendParagraph('Análisis para inspección')
          .setAttributes({[A.FONT_SIZE]:12,[A.BOLD]:true,[A.FOREGROUND_COLOR]:'#000f94'});
      body.appendParagraph(r.analisis || '')
          .setAttributes({[A.FONT_SIZE]:10});
      body.appendHorizontalRule();
    });

    doc.saveAndClose();
    return { ok: true, url: doc.getUrl(), nombre: doc.getName() };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ── Helpers ───────────────────────────────────────────────────
function dividirEnChunks(texto, max) {
  if (texto.length <= max) return [texto];
  const chunks = []; let ini = 0;
  while (ini < texto.length) {
    let fin = Math.min(ini + max, texto.length);
    if (fin < texto.length) {
      const s = texto.lastIndexOf('\n', fin);
      if (s > ini + max * 0.6) fin = s;
    }
    chunks.push(texto.substring(ini, fin));
    ini = fin;
  }
  return chunks;
}

function llamarOpenAI(apiKey, messages, maxTokens) {
  const resp = UrlFetchApp.fetch(CONFIG.OPENAI_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: CONFIG.OPENAI_MODEL,
      messages: messages,
      max_tokens: maxTokens,
      temperature: 0.1
    }),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const data = JSON.parse(resp.getContentText());
  if (code !== 200)
    throw new Error('OpenAI: ' + (data.error ? data.error.message : 'HTTP ' + code));
  return data.choices[0].message.content;
}

function extraerJSON(texto) {
  const m = texto.match(/```json\s*([\s\S]*?)\s*```/);
  const s = m ? m[1] : texto;
  try { return JSON.parse(s); }
  catch(_) { return JSON.parse(s.replace(/[\x00-\x1F\x7F]/g,' ').trim()); }
}

function fmtFecha() { return new Date().toLocaleDateString('es-ES'); }

// ============================================================
// EXPORTAR PLANTILLA PF (Anexo IV oficial)
// ============================================================
function exportarAPF(nombreCurso, resultados) {
  try {
    const plantillaId = PropertiesService.getUserProperties().getProperty('PF_PLANTILLA_ID');
    if (!plantillaId)
      throw new Error('No hay ID de plantilla configurado. Pega el ID de Drive de PF_plantilla.docx en el panel izquierdo.');

    const token    = ScriptApp.getOAuthToken();
    const copyUrl  = 'https://www.googleapis.com/drive/v3/files/' + plantillaId + '/copy';
    const copyOpts = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({
        name: 'PF — ' + nombreCurso + ' — ' + fmtFecha(),
        mimeType: 'application/vnd.google-apps.document'
      }),
      muteHttpExceptions: true
    };

    let copyResp, intentos = 0;
    do {
      if (intentos > 0) Utilities.sleep(4000 * intentos);
      copyResp = UrlFetchApp.fetch(copyUrl, copyOpts);
      intentos++;
    } while ((copyResp.getResponseCode() === 429 || copyResp.getResponseCode() === 403) && intentos < 5);

    if (copyResp.getResponseCode() !== 200)
      throw new Error('Error copiando plantilla (código ' + copyResp.getResponseCode() + '). Verifica el ID y que tienes acceso al archivo.');

    const docId = JSON.parse(copyResp.getContentText()).id;
    Utilities.sleep(3000);

    const doc  = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Localizar tablas del Anexo IV
    const tablasPF = [];
    for (let i = 0; i < body.getNumChildren(); i++) {
      const elem = body.getChild(i);
      if (elem.getType() !== DocumentApp.ElementType.TABLE) continue;
      const tbl   = elem.asTable();
      if (tbl.getNumRows() < 2) continue;
      const layout = pfDetectarTablaAnexoIV(tbl);
      if (layout) tablasPF.push({ tabla: tbl, layout: layout });
    }

    if (tablasPF.length === 0)
      throw new Error('No se encontraron tablas del Anexo IV. Asegúrate de que el ID es el de PF_plantilla.docx.');

    // Rellenar cada tabla con su unidad
    for (let idx = 0; idx < resultados.length; idx++) {
      const tablaInfo = tablasPF[Math.min(idx, tablasPF.length - 1)];
      const tbl = tablaInfo.tabla;
      const layout = tablaInfo.layout;
      const r   = resultados[idx];

      const caps          = pfAgruparCapacidades(r.ejercicios, r.objetivos_y_criterios, r.contenidos);
      const numFilasDatos = caps.length > 0 ? caps.length : 1;
      const filaInicio    = layout.filaInicio;

      pfPrepararFilasPF(tbl, filaInicio, numFilasDatos, layout.numColumnas);

      if (caps.length === 0) {
        const fila = tbl.getRow(filaInicio);
        pfRellenarColumnaPF(fila, layout.columnas.objetivos, r.objetivos_y_criterios || '');
        pfRellenarColumnaPF(fila, layout.columnas.logro, '');
        pfRellenarColumnaPF(fila, layout.columnas.contenidos, r.contenidos || '');
        const actsTexto = (r.ejercicios || []).map(function(e) {
          return pfTextoEstrategiaPF(e);
        }).join('\n\n');
        pfRellenarColumnaPF(fila, layout.columnas.estrategias, actsTexto);
        pfRellenarColumnaPF(fila, layout.columnas.equipos, '');

      } else {
        caps.forEach(function(cap, ci) {
          const filaIdx = filaInicio + ci;
          if (filaIdx >= tbl.getNumRows()) return;
          const fila = tbl.getRow(filaIdx);

          pfRellenarColumnaPF(fila, layout.columnas.objetivos, cap.criterios || '');
          pfRellenarColumnaPF(fila, layout.columnas.logro, '');
          pfRellenarColumnaPF(fila, layout.columnas.contenidos, cap.contenidos || r.contenidos || '');

          const actsTexto = (cap.actividades || []).map(function(e) {
            return pfTextoEstrategiaPF(e);
          }).join('\n\n');

          pfRellenarColumnaPF(fila, layout.columnas.estrategias, actsTexto);
          pfRellenarColumnaPF(fila, layout.columnas.equipos, '');
        });
      }

      // Eliminar filas sobrantes
      const totalEsperado = filaInicio + numFilasDatos;
      for (let ri = tbl.getNumRows() - 1; ri >= totalEsperado; ri--) {
        tbl.removeRow(ri);
      }
    }

    doc.saveAndClose();
    return { ok: true, url: doc.getUrl(), nombre: doc.getName() };

  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── Rellena una celda línea a línea ──────────────────────────
function pfRellenarCelda(celda, texto) {
  if (!celda) return;
  celda.clear();
  const lineas = (texto || '').split('\n');
  lineas.forEach(function(linea) { celda.appendParagraph(linea); });
}

function pfDetectarTablaAnexoIV(tbl) {
  const maxFilas = Math.min(tbl.getNumRows(), 5);
  let mejor = null;

  for (let ri = 0; ri < maxFilas; ri++) {
    const fila = tbl.getRow(ri);
    const encontrados = {};
    let score = 0;

    for (let ci = 0; ci < fila.getNumCells(); ci++) {
      const txt = pfNormalizarTextoPF(fila.getCell(ci).getText());

      if (txt.indexOf('objetivo') !== -1 && encontrados.objetivos === undefined) {
        encontrados.objetivos = ci;
        score++;
      }
      if ((txt.indexOf('logro') !== -1 ||
           (txt.indexOf('resultado') !== -1 && txt.indexOf('aprendizaje') !== -1) ||
           txt.indexOf('criterio') !== -1) &&
          encontrados.logro === undefined) {
        encontrados.logro = ci;
        score++;
      }
      if (txt.indexOf('contenido') !== -1 && encontrados.contenidos === undefined) {
        encontrados.contenidos = ci;
        score++;
      }
      if ((txt.indexOf('estrategia') !== -1 ||
           txt.indexOf('actividad') !== -1 ||
           txt.indexOf('recurso') !== -1) &&
          encontrados.estrategias === undefined) {
        encontrados.estrategias = ci;
        score++;
      }
      if ((txt.indexOf('equipo') !== -1 ||
           txt.indexOf('instalacion') !== -1 ||
           txt.indexOf('equipamiento') !== -1) &&
          encontrados.equipos === undefined) {
        encontrados.equipos = ci;
        score++;
      }
    }

    const tieneColumnaObjetivos = encontrados.objetivos !== undefined || encontrados.logro !== undefined;
    if (score >= 3 && tieneColumnaObjetivos &&
        encontrados.contenidos !== undefined && encontrados.estrategias !== undefined) {
      if (!mejor || score > mejor.score) {
        mejor = { filaCabecera: ri, encontrados: encontrados, score: score, numColumnas: fila.getNumCells() };
      }
    }
  }

  if (!mejor) {
    const fila0 = tbl.getRow(0);
    const txt0 = fila0.getNumCells() ? pfNormalizarTextoPF(fila0.getCell(0).getText()) : '';
    if (txt0.indexOf('objetivo') === -1 && txt0.indexOf('logro') === -1) return null;
    const colsFallback = fila0.getNumCells() >= 5
      ? { objetivos: 0, logro: 1, contenidos: 2, estrategias: 3, equipos: 4 }
      : { objetivos: 0, logro: null, contenidos: 1, estrategias: 2, equipos: 3 };
    return {
      filaInicio: Math.min(2, tbl.getNumRows()),
      columnas: colsFallback,
      numColumnas: Math.max(fila0.getNumCells(), 4)
    };
  }

  const encontrados = mejor.encontrados;
  const objetivosCol = encontrados.objetivos !== undefined ? encontrados.objetivos : encontrados.logro;
  const logroCol = (encontrados.objetivos !== undefined &&
                    encontrados.logro !== undefined &&
                    encontrados.logro !== encontrados.objetivos)
    ? encontrados.logro
    : null;

  return {
    filaInicio: mejor.filaCabecera + 1,
    columnas: {
      objetivos: objetivosCol,
      logro: logroCol,
      contenidos: encontrados.contenidos,
      estrategias: encontrados.estrategias,
      equipos: encontrados.equipos !== undefined ? encontrados.equipos : null
    },
    numColumnas: mejor.numColumnas
  };
}

function pfPrepararFilasPF(tbl, filaInicio, numFilasDatos, numColumnas) {
  while (tbl.getNumRows() < filaInicio + numFilasDatos) {
    const nuevaFila = tbl.appendTableRow();
    while (nuevaFila.getNumCells() < numColumnas) nuevaFila.appendTableCell('');
  }

  for (let ri = tbl.getNumRows() - 1; ri >= filaInicio + numFilasDatos; ri--) {
    tbl.removeRow(ri);
  }

  for (let ri = filaInicio; ri < filaInicio + numFilasDatos; ri++) {
    const fila = tbl.getRow(ri);
    while (fila.getNumCells() < numColumnas) fila.appendTableCell('');
    for (let ci = 0; ci < fila.getNumCells(); ci++) {
      pfRellenarCelda(fila.getCell(ci), '');
    }
  }
}

function pfRellenarColumnaPF(fila, colIdx, texto) {
  if (colIdx === null || colIdx === undefined || colIdx < 0) return;
  while (fila.getNumCells() <= colIdx) fila.appendTableCell('');
  pfRellenarCelda(fila.getCell(colIdx), pfLimpiarTextoTablaPF(texto));
}

function pfTextoEstrategiaPF(ej) {
  const identificador = pfLimpiarTextoTablaPF(ej.identificador || '');
  const estrategia = pfLimpiarTextoTablaPF(ej.estrategias_actividades_recursos || '');
  if (identificador && estrategia) return identificador + ': ' + estrategia;
  return estrategia || identificador;
}

function pfLimpiarTextoTablaPF(texto) {
  let limpio = String(texto || '');
  limpio = limpio.replace(/transcripci.{0,4}n literal\s*:?\s*[\s\S]*$/i, '');
  return limpio.trim();
}

function pfNormalizarTextoPF(texto) {
  return String(texto || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

// ── Agrupa ejercicios por capacidad (C1, C2...) ──────────────
function pfAgruparCapacidades(ejercicios, objetivos_y_criterios, contenidos) {
  if (!objetivos_y_criterios) return [];

  const capOrder = [];
  const capMap   = {};
  const lineas   = objetivos_y_criterios.split('\n');
  let capActual  = null;
  let bloque     = [];

  lineas.forEach(function(linea) {
    const mCap = linea.match(/^(C\d+)\s*:/);
    if (mCap) {
      if (capActual) capMap[capActual].criterios = bloque.join('\n');
      capActual = mCap[1];
      if (!capMap[capActual]) {
        capMap[capActual] = {
          id: capActual,
          criterios: '',
          contenidos: contenidos || '',
          actividades: []
        };
        capOrder.push(capActual);
      }
      bloque = [linea];
    } else if (capActual) {
      bloque.push(linea);
    }
  });
  if (capActual && bloque.length) capMap[capActual].criterios = bloque.join('\n');

  if (!capOrder.length) return [];

  (ejercicios || []).forEach(function(ej) {
    const texto = (ej.estrategias_actividades_recursos || '') + ' ' +
                  (ej.objetivos_y_criterios || '');
    let asignada = false;
    capOrder.forEach(function(capId) {
      if (new RegExp('\\b' + capId + '\\b').test(texto)) {
        capMap[capId].actividades.push(ej);
        asignada = true;
      }
    });
    if (!asignada) capMap[capOrder[0]].actividades.push(ej);
  });

  return capOrder.map(function(id) { return capMap[id]; });
}

// ── Guardar / recuperar ID de plantilla PF ───────────────────
function guardarPlantillaId(fileId) {
  PropertiesService.getUserProperties().setProperty('PF_PLANTILLA_ID', fileId.trim());
  return { ok: true };
}

function plantillaConfigurada() {
  return !!PropertiesService.getUserProperties().getProperty('PF_PLANTILLA_ID');
}
