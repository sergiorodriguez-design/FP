// ============================================================
// Prompts.gs — Prompts para OpenAI
// ============================================================

// ── PROMPT PRINCIPAL: extrae TODO de una unidad en una sola llamada ──
// Devuelve: objetivos_y_criterios (1 vez), contenidos (1 vez), lista de actividades
function buildPromptUnidad(textoChunk, infoUnidad, numChunk, totalChunks) {
  const esPrimero = numChunk === 1;
  return [
    {
      role: 'system',
      content: `Eres un experto en programación didáctica de Formación Profesional para el Empleo (FPE) en España.
Analizas el texto de una unidad didáctica para extraer:
1. Los criterios de evaluación (capacidades C1, C2... y sus criterios CE1.1, CE1.2...) tal como aparecen literalmente en el documento.
2. El mapa/esquema de contenidos de la unidad tal como aparece en el documento (bullet points, epígrafes, bloques temáticos).
3. Todas las actividades formativas con su detalle completo y su relación explícita con los criterios de evaluación.
Responde SIEMPRE con JSON válido, sin texto adicional.`
    },
    {
      role: 'user',
      content: `Analiza el siguiente texto de la unidad: "${infoUnidad.etiqueta}"
${totalChunks > 1 ? `(Parte ${numChunk} de ${totalChunks})` : ''}

═══════════════════════════════════════════════
TEXTO:
${textoChunk}
═══════════════════════════════════════════════

${esPrimero ? `INSTRUCCIONES PARA CAMPOS DE UNIDAD (solo en la parte 1):
- "objetivos_y_criterios": Copia LITERALMENTE del texto las capacidades (C1, C2...) y todos sus criterios de evaluación (CE1.1, CE1.2...) con su descripción completa. Si hay varias capacidades, inclúyelas todas. Formato: "C1: [descripción completa]\nCE1.1 [descripción]\nCE1.2 [descripción]\n...". Si no aparecen explícitamente, deja cadena vacía.
- "contenidos": Copia LITERALMENTE el esquema, mapa conceptual o índice de contenidos de la unidad tal como aparece (epígrafes, bullets, bloques). Si no aparece un esquema claro, resume los grandes bloques temáticos.` : `NOTA: Esta es la parte ${numChunk}. Los campos "objetivos_y_criterios" y "contenidos" déjalos como cadena vacía "" ya que se extrajeron en la parte 1. Solo extrae las actividades nuevas.`}

INSTRUCCIONES PARA ACTIVIDADES:
- Identifica TODAS las actividades, ejercicios, casos prácticos, tareas y preguntas.
- "identificador": nombre EXACTO como aparece en el texto (ej: "Actividad de Aprendizaje 3", "Tarea de Evaluación 2", "Ejercicio de autoevaluación 5").
- "tipo": clasifica como: Actividad de Aprendizaje | Actividad Colaborativa | Tarea de Evaluación | Caso Práctico | Ejercicio de Autoevaluación | Supuesto Práctico | Pregunta de Desarrollo | Ejercicio de Repaso | Otro
- "texto_completo": enunciado íntegro de la actividad, copiado literalmente del texto.
- "estrategias_actividades_recursos": descripción detallada de CÓMO se realiza esta actividad específica, indicando:
    * A qué capacidad/es (C1, C2...) y criterios de evaluación (CE1.1, CE1.2...) está vinculada EXPLÍCITAMENTE.
    * La metodología (supuesto práctico, trabajo colaborativo, búsqueda documental, etc.).
    * El recurso o soporte requerido (vídeo, foro, normativa, supuesto facilitado por tutor, etc.).
    * La forma de corrección y entrega (autocorrección con feedback, envío al tutor, puesta en común en foro, etc.).
    * Si es individual, grupal o colaborativa.
- "observaciones": notas para inspección (dependencias, recursos no especificados, inconsistencias).

Responde con este JSON exacto:
{
  "objetivos_y_criterios": "Texto literal de capacidades y criterios tal como aparecen en el documento",
  "contenidos": "Esquema o mapa de contenidos de la unidad tal como aparece en el documento",
  "actividades": [
    {
      "identificador": "Nombre exacto del ejercicio",
      "tipo": "Tipo de actividad",
      "texto_completo": "Enunciado completo literal del ejercicio",
      "estrategias_actividades_recursos": "Descripción detallada incluyendo vinculación a CEs, metodología, recursos, corrección y modalidad",
      "observaciones": "Notas para inspección o cadena vacía"
    }
  ]
}

Si no hay actividades en este fragmento, devuelve "actividades": [].
No inventes nada que no esté en el texto.`
    }
  ];
}

// ── PROMPT NARRATIVO ──────────────────────────────────────────
function buildPromptNarrativo(actividades, infoUnidad, objetivos, contenidos) {
  const lista = actividades.map((e, i) =>
    `${i+1}. [${e.tipo}] ${e.identificador}`
  ).join('\n');

  return [
    {
      role: 'system',
      content: `Eres un experto en programación didáctica de FPE en España, especializado en preparar documentación para inspecciones del SEPE y comunidades autónomas. Tu análisis debe ser riguroso, técnico y orientado a demostrar la coherencia didáctica y el cumplimiento normativo.`
    },
    {
      role: 'user',
      content: `Genera un análisis narrativo completo para la unidad "${infoUnidad.etiqueta}" con ${actividades.length} actividades, orientado a inspecciones de calidad FPE.

CRITERIOS Y CAPACIDADES DE LA UNIDAD:
${objetivos || '(no especificados en el texto)'}

CONTENIDOS DE LA UNIDAD:
${contenidos || '(no especificados en el texto)'}

ACTIVIDADES DETECTADAS (${actividades.length}):
${lista}

El análisis debe incluir estos apartados marcados en negrita:

**1. Visión general de las actividades**
Número total, distribución por tipos (actividades de aprendizaje, colaborativas, tareas de evaluación, autoevaluación, etc.) y equilibrio metodológico general de la unidad.

**2. Coherencia didáctica y progresión pedagógica**
¿Las actividades guardan progresión lógica desde lo conceptual a lo aplicado? ¿Hay variedad metodológica adecuada al nivel FPE? ¿Se construyen sobre conocimientos previos?

**3. Cobertura de capacidades y criterios de evaluación**
¿Qué capacidades (C1, C2...) y criterios de evaluación (CEs) tienen actividades asociadas? ¿Existe algún CE sin actividad asignada? ¿La distribución de actividades por CE es equilibrada?

**4. Estrategias metodológicas destacadas**
Señala las estrategias más relevantes presentes: corrección automática con feedback, aprendizaje colaborativo por foro, supuestos prácticos contextualizados, tareas enviadas al tutor, trabajo en grupo, búsqueda documental, etc.

**5. Valoración para inspección**
Puntos fuertes del diseño didáctico. Aspectos que podrían requerir justificación adicional o que presentan alguna laguna: recursos no especificados, dependencias entre actividades, inconsistencias entre CEs declarados y actividades diseñadas, etc.

Redacta en prosa formal, en español. Entre 400 y 600 palabras.`
    }
  ];
}
