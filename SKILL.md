# Slide Generator · Niñas Pro

Genera presentaciones `.pptx` para todas las clases siguiendo el Brandbook Niñas Pro 2022 y los marcos pedagógicos 5E + Bloom + DUA.

> **Tono no negociable:** este generador produce slides para Niñas Pro. Las slides deben sentirse **empoderadoras, con conector entre clases, con analogías humanas, con calor**. Antes de generar, lee `.agents/skills/skill-create-course/references/tone-and-narrative-guide.md` y abre `Clase-09-Funciones.pptx` (en la raíz del repo) como benchmark visual. Slides técnicamente correctas pero frías = slides falladas.

## Prerrequisitos

```bash
npm install pptxgenjs react react-dom sharp react-icons
```

## API

```js
const {
  generateSlides,
  validateSlideSequence,
  buildClassSkeleton,
  C,
  CODE_TOKENS,
  BRAND_ICONS,
} = require(".agents/skills/skill-slide-generator/generate.js");

await generateSlides({
  className: "1",
  classTitle: "Título de la clase",
  courseName: "Programación Competitiva Básico",
  outputPath: "class-N/slides.pptx",
  logoPath: "Logo.png",      // opcional — por defecto busca Logo.png en el root
  validate: true,            // opcional, default true — emite warnings pedagógicos
  slides: [ /* ver tipos abajo */ ]
});
```

`generateSlides` corre `validateSlideSequence` automáticamente (warnings en consola, no bloquea). Pasa `validate: false` para silenciarlo.

## Esqueleto pedagógico (recomendado)

Empieza con `buildClassSkeleton` y agrega tus slides de contenido en medio. Devuelve una secuencia 5E-compliant con `title → engage → objectives → kahoot → steam → reflection → closing`.

```js
const base = buildClassSkeleton({
  classNumber: 1,
  title: "Algoritmos y\nlógica secuencial",
  subtitle: "Programación Competitiva Básico",
  objectives: [
    { verb: "Identificar", content: "los pasos de un algoritmo en pseudocódigo", bloom: 2 },
    { verb: "Diseñar", content: "un diagrama de flujo simple", bloom: 6 },
    { verb: "Comparar", content: "dos algoritmos para la misma tarea", bloom: 4 },
  ],
  steamWoman: {
    name: "Grace Hopper",
    years: "1906–1992",
    facts: ["Inventó el primer compilador (1952)", "Contraalmirante de la Marina EE.UU."],
    quote: "Lo más peligroso que puedes decir es: siempre lo hemos hecho así.",
  },
  hoyAprendimos: ["Qué es un algoritmo", "Pseudocódigo", "Diagramas de flujo"],
  proximaClase: ["Nuestro primer programa en C++", "Variables y tipos de datos"],
  tarea: ["Algoritmo de tu rutina en pseudocódigo"],
});

// Inserta slides de contenido entre objectives y kahoot
base.splice(3, 0,
  { type: "activity", /* ... */ },
  { type: "definition", /* ... */ },
  { type: "code", /* ... */ },
  { type: "manual-trace", /* ... */ },
);

await generateSlides({ className: "1", classTitle: "...", slides: base });
```

## Mapeo 5E → tipos de slide

| Fase 5E | Tipos recomendados |
|---------|-------------------|
| **Engage** | `engage` (slide 02 — pregunta gancho, nunca cátedra fría) |
| **Explore** | `activity`, `error-code`, `manual-trace` (intentar antes de explicar) |
| **Explain** | `content`, `definition`, `pillars`, `code`, `flowchart`, `reference` |
| **Elaborate** | `activity`, `steps`, `code` aplicados a contexto nuevo |
| **Evaluate** | `kahoot`, `steam`, `reflection` (Ejercicio 10), `rubric`, `closing` |

Slides estructurales (no 5E): `title`, `section`, `objectives`, `plan-b`.

## Slides obligatorios en cada clase

- `title` (slide 0).
- `engage` (slide 1 — primera de contenido).
- `objectives` con verbos observables Bloom.
- Al menos un `manual-trace` o `activity` antes del primer `code` (Explore antes que Explain).
- `steam` antes del cierre.
- `reflection` (Ejercicio 10 — no-negociable).
- `closing` con tres bloques: aprendido / próxima clase / tarea.

## Tipos de slide

### `title` — Portada (con tagline poético obligatorio)
```js
{ type: "title",
  classNumber: "9",
  title: "Funciones y\nmodularidad",
  tagline: "Dividir para conquistar. Bloques de código con nombre propio.",
  subtitle: "Programación Competitiva Básico",
  notes: "..." }
```
El `tagline` es **no-negociable** (per `tone-and-narrative-guide.md`): comprime la idea central de la clase en una frase memorable. Si no lo pasas, el validador emite warning. Renderiza el marco pixel-art retro y squiggle amarillo.

### `bridge` — Conector entre clases (slide 02, no-negociable excepto C01)
```js
{ type: "bridge",
  sectionLabel: "00 · APERTURA",
  direction: "backward",  // o "forward" para cierre con preview de próxima clase
  title: "Funciones y modularidad",
  subtitle: "Lo que ya sabes y lo que viene hoy.",
  leftLabel: "LO QUE YA SABES",
  leftSubtitle: "De clases anteriores:",
  left: ["Variables, if, ciclos.", "Arreglos y vectores.", "Búsqueda lineal (Clase 08).", "Y todo dentro de main()."],
  rightLabel: "LO QUE VIENE HOY",
  right: ["Sacamos código de main", "y lo empaquetamos en bloques con nombre propio.", "Una vez creada, la usas tantas veces como quieras sin reescribirla."],
  notes: "Conecta explícitamente con la clase anterior." }
```
Dos tarjetas lado a lado: izquierda púrpura (lo que ya saben), derecha amarilla (lo nuevo). Para el `bridge` forward (cerca del final de la clase), usa `direction: "forward"` — el right card pasa a ser un preview de la próxima clase.

### `pull-quote` — Frase poética sola
```js
{ type: "pull-quote",
  quote: "Una función hace UNA cosa, y la hace bien.",
  subtitle: "Dividir para conquistar.",
  bgColor: "LAVENDER_LIGHT",  // EAE0FE — opción cálida
  notes: "Pausa visual antes de entrar a la anatomía." }
```
Fondo lavanda claro, frase grande y centrada. Útil entre bloques densos.

### `analogy` — VIDA REAL ↔ TÉRMINO TÉCNICO
```js
{ type: "analogy",
  sectionLabel: "01 · IDEA",
  title: "Una función es como tu amiga",
  subtitle: "Tiene nombre, hace algo, y a veces te devuelve algo.",
  leftLabel: "VIDA REAL",
  rightLabel: "EN PROGRAMACIÓN",
  rows: [
    { left: "Tiene un nombre — \"Camila\"",          rightLabel: "Nombre de la función", right: "traerCafe" },
    { left: "Sabe hacer algo — preparar café",       rightLabel: "Cuerpo de la función", right: "el código entre { }" },
    { left: "Le pides con info — \"Cami, con leche\"", rightLabel: "Argumentos",         right: "traerCafe(\"leche\")" },
    { left: "Te devuelve algo — una taza de café",   rightLabel: "return",                right: "return cafe;" },
    { left: "O solo hace algo — \"Cami, avisa\"",    rightLabel: "función void",          right: "void avisar()" },
  ],
  notes: "Aquí la clase deja de sentirse fría. Usar nombres chilenos reales (Camila, Valentina)." }
```
Donde el slide deja de ser un manual técnico y empieza a ser una explicación humana. **Al menos una `analogy` por clase es obligatoria.**

### `comparison` — Antes / Después (sin X / con X)
```js
{ type: "comparison",
  sectionLabel: "01 · IDEA",
  title: "El problema de copiar y pegar",
  subtitle: "¿Qué pasa si hago lo mismo en 3 partes del programa?",
  leftLabel: "ANTES", leftFile: "sin-funciones.cpp",
  left: [
    [{ text: "// Calcular área de rectángulo 1", token: "comment", breakLine: true }],
    [{ text: "int largo1 = 5, ancho1 = 3;", breakLine: true }],
    [{ text: "int area1 = largo1 * ancho1;", breakLine: true }],
    [{ text: "// Calcular área de rectángulo 2  ← ¡repetido!", token: "comment", breakLine: true }],
    // ...
  ],
  rightLabel: "DESPUÉS", rightFile: "con-funcion.cpp",
  right: [
    [{ text: "// La función vive una sola vez", token: "comment", breakLine: true }],
    [{ text: "int area(int largo, int ancho) {", breakLine: true }],
    [{ text: "  return largo * ancho;", breakLine: true }],
    [{ text: "}", breakLine: true }],
  ],
  takeaway: "Escribimos la fórmula UNA vez. La llamamos muchas.",
  notes: "Muestra el VALOR de la nueva técnica, no solo la sintaxis." }
```
Dos tarjetas navy lado a lado, con etiquetas color-coded (rojo "ANTES", verde "DESPUÉS"). Acepta arrays de código (líneas con tokens) o de texto plano.

### `anatomy` — Descomposición en partes
```js
{ type: "anatomy",
  sectionLabel: "02 · ANATOMÍA",
  title: "Anatomía de una función",
  subtitle: "Cuatro partes. Cada una con su rol.",
  example: "int sumar(int a, int b) { return a + b; }",
  parts: [
    { label: "TIPO RETORNO", code: "int",            desc: "Qué tipo devuelve.",         color: "PURPLE" },
    { label: "NOMBRE",       code: "sumar",          desc: "Cómo la llamamos.",          color: "BLUE" },
    { label: "PARÁMETROS",   code: "(int a, int b)", desc: "Datos que recibe.",          color: "GREEN" },
    { label: "RETORNO",      code: "return a + b;",  desc: "Lo que devuelve.",           color: "YELLOW" },
  ],
  notes: "Cuadrantes 2x2 con código en Space Mono. Cada parte con color." }
```
Hasta 4 partes en grilla 2×2. Si hay 1-2 parts, se renderiza en fila. Útil para sintaxis nuevas o conceptos compuestos.

### `try-it` — Invitación cálida antes de práctica
```js
{ type: "try-it",
  title: "Antes de seguir, prueba tú.",
  prompt: "Escribe una función `bool esPar(int n)` que devuelva true si n es par, false si no. Después úsala desde main para imprimir si 7 es par.",
  duration: "5 min",
  mode: "Individual",
  hint: "Tip: el operador % te dice el resto de una división.",
  notes: "Slide de transición a práctica. Tono cálido, no lanza el ejercicio en seco." }
```
Fondo amarillo, label "INTÉNTALO TÚ", el desafío en ventana blanca con borde navy. La diferencia entre "ejercicio" y "invitación".

### `section` — Divisor de capítulo (estilo brandbook)
```js
{ type: "section", number: 3, label: "PRÁCTICA AUTÓNOMA",
  description: "Las siguientes 45 min son tuyas. Avanza en Dodona a tu ritmo.",
  notes: "..." }
```
Fondo navy + marco pixel-art con smiley + cursor + letra "A" con squiggle (composición brandbook §00). Útil para transicionar entre bloques largos.

### `welcome` — Sesión de bienvenida (sesión especial)
```js
{ type: "welcome",
  greeting: "¡BIENVENIDAS!",
  courseName: "PROGRAMACIÓN COMPETITIVA BÁSICO",
  agenda: [
    { text: "Conocernos", duration: "20 min" },
    { text: "Pensamiento computacional", duration: "30 min" },
    { text: "Mujeres en la tecnología", duration: "20 min" },
    { text: "Cierre y próximos pasos", duration: "10 min" },
  ],
  notes: "Sesión institucional de apertura. No requiere objectives/reflection/etc."
}
```
Fondo púrpura + smiley pixel-art + cursor + agenda numerada con cuadrados rosa. Sesiones `welcome` están exentas del validador 5E estándar.

### `block-transition` — Transición entre bloques 5E
```js
{ type: "block-transition",
  phase: "Elaborate",
  phaseColor: "TEAL",
  block: "PRÁCTICA AUTÓNOMA",
  duration: "45 min",
  items: [
    "Resuelve los ejercicios 3 a 7 en Dodona",
    "Si te trabas, intenta un trazado manual antes de pedir ayuda",
    "Mentora circula y acompaña — no toma el problema",
  ],
  readyCheck: "Antes de empezar: ¿tienes tu laptop y abierto Dodona?",
  notes: "Transición clara entre bloques 5E del time-distribution-guide."
}
```
Útil entre Apertura→Cátedra→Kahoot→Práctica→Cierre. Refuerza claridad del bloque y duración (DUA: instrucciones duales).

### `engage` — Gancho inicial (5E: Engage)
```js
{ type: "engage", title: "¿Pregunta?", icon: "lightbulb",
  body: ["Línea 1", { text: "Línea destacada", bold: true }], notes: "..." }
```
Máximo 5 líneas en `body` (warning si excedes). Fuente base 22pt.

### `objectives` — Objetivos de aprendizaje (Bloom + mensaje cálido obligatorio)
```js
{ type: "objectives",
  sectionLabel: "00 · APERTURA",
  title: "Objetivos de hoy",
  subtitle: "Al final de hoy podrás…",
  objectives: [
    { verb: "Declarar",  content: "y llamar funciones con parámetros y valor de retorno", bloom: 3 },
    { verb: "Distinguir", content: "funciones con retorno de funciones void",              bloom: 4 },
    { verb: "Comprender", content: "el alcance (scope) de las variables",                  bloom: 2 },
    { verb: "Descomponer", content: "un programa grande en funciones con responsabilidades claras", bloom: 4 },
  ],
  warmClosing: "Si al final puedes hacer estas cuatro cosas, la clase fue exitosa.",
  notes: "..." }
```
Fondo lavanda cálido (`#FBF7FF`). El `warmClosing` es **no-negociable** (per `tone-and-narrative-guide.md`) — sin él, los objetivos quedan secos. Si lo omites, el generator usa un default genérico.

### `steps` — Lista numerada (patrón brandbook)
```js
{ type: "steps", title: "Cómo escribir un algoritmo",
  steps: [
    { keyword: "Identifica", desc: "el problema a resolver" },
    { keyword: "Descompón",  desc: "en pasos pequeños y ordenados" },
    { keyword: "Verifica",   desc: "con un ejemplo concreto" },
  ], notes: "..." }
```

### `activity` — Tarjeta de actividad (fondo rosa)
```js
{ type: "activity", title: "Título", accentColor: "BLUE",
  activityLabel: "Actividad en parejas", icon: "tools",
  body: ["Paso 1", { text: "Paso 2", bold: true }],
  cardY: 1.2, cardH: 3.8, notes: "..." }
```

### `content` — Slide de texto simple
```js
{ type: "content", title: "Título", accentColor: "GREEN",
  body: ["Texto 1", "Texto 2"], notes: "..." }
```
Máximo 5 líneas en `body`.

### `definition` — Definición destacada (fondo navy)
```js
{ type: "definition", title: "¿Qué es?", accentColor: "PURPLE",
  definition: "Definición central",
  cards: [{ label: "Concepto", desc: "Descripción", color: "GREEN" }, ...],
  notes: "..." }
```

### `pillars` — Pilares / lista vertical con barras
```js
{ type: "pillars", title: "Título", accentColor: "BLUE",
  pillars: [{ label: "Pilar", desc: "Descripción", color: "PURPLE" }, ...],
  notes: "..." }
```

### `code` — Bloque de código (fondo navy)
```js
{ type: "code", title: "Título", accentColor: "GREEN",
  description: "Texto explicativo",
  code: [
    [{ text: "for ",    token: "keyword" }, { text: "(int i = 0; i < n; i++) {", breakLine: true }],
    [{ text: "  std::cout << ", token: "type" }, { text: "\"Hola\"", token: "string" }, { text: ";", breakLine: true }],
    [{ text: "}", breakLine: true }],
  ], notes: "..." }
```
**Tokens nombrados** (recomendado sobre hex literales):
- `keyword` → amarillo `#FDCA36`
- `type` → azul `#2B88F7`
- `string` → verde `#05A175`
- `number` → teal `#6FD0D8`
- `comment` → gris `#666666`
- `error` → rojo `#FC3535`

### `manual-trace` — Trazado manual (DUA + estándar pedagógico)
```js
{ type: "manual-trace", title: "Trazado manual del bucle",
  accentColor: "TEAL",
  code: [
    [{ text: "int suma = 0;", breakLine: true }],
    [{ text: "for (int i = 1; i <= 3; i++)", token: "keyword", breakLine: true }],
    [{ text: "  suma += i;", breakLine: true }],
  ],
  headers: ["Paso", "i", "suma"],
  rows: [
    [1, 1, 1],
    [2, 2, 3],
    [3, 3, 6],
  ],
  footer: "Comprobar manualmente antes de codificar.",
  notes: "EXPLORE/UDL. Trazado manual antes de escribir código." }
```

### `flowchart` — Símbolos de diagrama de flujo
```js
{ type: "flowchart", title: "Título", accentColor: "PURPLE",
  symbols: [{ shape: "Óvalo", meaning: "Inicio / Fin", color: "PURPLE" }, ...],
  subtitle: "Texto opcional", example: "Ejemplo visual", notes: "..." }
```

### `error-code` — Código con error
```js
{ type: "error-code", title: "Título", icon: "exclamation",
  question: "¿Dónde está el error?",
  code: [ /* igual que type code */ ],
  errorMessage: "Mensaje del compilador o de lógica",
  protocol: "Antes de explicar, pregúntale a la alumna: ¿qué te está diciendo el sistema?",
  notes: "..." }
```
El campo opcional `protocol` renderiza un callout amarillo "Mentora: …" para reforzar el protocolo de `facilitation-guide.md` (leer el error CON la alumna, no para ella).

### `reference` — Tabla de referencia
```js
{ type: "reference", title: "Título", accentColor: "BLUE",
  headers: ["Col1", "Col2", "Col3"],
  rows: [["⬭", "Óvalo", "Inicio / Fin"], ...],
  // O sin headers:
  rows: [{ sym: "⬭", name: "Óvalo", desc: "Inicio / Fin" }, ...],
  cardH: 3.8, footer: "Texto al pie", notes: "..." }
```

### `glossary` — Vocabulario clave (2-4 términos destacados)
```js
{ type: "glossary", title: "Vocabulario clave",
  subtitle: "El glosario completo está en 04-glossary.md.",
  accentColor: "TEAL",
  terms: [
    { term: "Algoritmo",  definition: "Secuencia finita de pasos ordenados para resolver un problema.",
      example: "Receta de cocina" },
    { term: "`for`",      definition: "Ciclo que repite N veces con un contador.",
      example: "for (int i = 0; i < 5; i++)" },
    { term: "Pseudocódigo", definition: "Algoritmo escrito en lenguaje humano, no en un lenguaje real.",
      example: "INICIO / hervir agua / FIN" },
  ], footer: "Estos términos también están en tu cuaderno.",
  notes: "EVALUATE. Refuerza vocabulario antes del cierre." }
```
Hasta 4 términos en grilla 2×2 con cuadrados rosa numerados. Términos en backticks se renderizan en Space Mono.

### `kahoot` — Transición Kahoot
```js
{ type: "kahoot", notes: "..." }
```

### `steam` — Mujer STEAM (alineado con `steam-woman-guide.md`)
```js
{ type: "steam",
  name: "Karen Spärck Jones", years: "1935–2007",
  area: "Procesamiento de lenguaje natural",
  country: "Reino Unido",
  connectionToClass: "Su trabajo en TF-IDF es la versión a gran escala del problema de búsqueda lineal que vimos hoy.",
  photoPath: "class-08/karen-sparck-jones.jpg",  // opcional, brandbook §04: foto B&N
  facts: [
    "Pionera del procesamiento de lenguaje natural",
    "Inventó TF-IDF (1972), base de motores de búsqueda actuales",
    "Defendía: 'la informática es demasiado importante para dejársela a los hombres'",
  ],
  quote: "La informática es demasiado importante para dejársela a los hombres.",
  quoteAttribution: "Karen Spärck Jones, entrevista BCS 2007",
  reflectionQuestion: "¿En qué otros lugares de la tecnología falta esa misma diversidad de voces?",
  source: "Wikipedia EN + obituario The Guardian 2007",
  bgColor: "PURPLE",
  notes: "EVALUATE. Conectar TF-IDF con búsqueda lineal — mismo problema a otra escala." }
```
Campos requeridos por el guide pedagógico: `name`, `years`, `area`, `country`, `connectionToClass`, `facts`, `quote`+`quoteAttribution`, `reflectionQuestion`, `source`. La foto B&N opcional se renderiza con star/squiggle de la paleta encima (brandbook §04). El validador emite warning si falta `connectionToClass`.

### `reflection` — Ejercicio 10 / cierre Evaluate (no-negociable)
```js
{ type: "reflection", title: "Reflexión final",
  prompts: [
    "¿Qué fue lo más difícil de hoy?",
    "¿Qué fue lo más sorprendente?",
    "¿Qué pregunta te queda sin responder?",
  ],
  duration: "8 min",
  notes: "EVALUATE — Ejercicio 10. Reflexión escrita individual." }
```

### `rubric` — Rúbrica formativa 1–10
```js
{ type: "rubric", title: "Rúbrica formativa de la clase",
  accentColor: "GREEN",
  // Si omites levels, usa la rúbrica estándar Niñas Pro (1-2 .. 9-10)
  levels: [
    { range: "1–2", desc: "No accede a la tarea básica", color: "RED" },
    { range: "9–10", desc: "Todos los ejercicios + casos borde + justifica", color: "GREEN" },
  ],
  footer: "La rúbrica es formativa, no una nota final." }
```

### `plan-b` — Plan B charla inspiracional (Clases 3, 5, 9)
```js
{ type: "plan-b", title: "Si la invitada no puede asistir",
  subtitle: "Activamos la actividad de respaldo:",
  steps: [
    "Formar grupos de 4",
    "Trazar manualmente el algoritmo del Ejercicio 7",
    "Cada grupo presenta su trazado en 3 min",
  ],
  notes: "Solo proyectar si se confirma la cancelación." }
```
Render como modal Windows-95 amarillo (brandbook §05).

### `closing` — Cierre
```js
{ type: "closing", label: "CIERRE Y PUENTE",
  sections: [
    { title: "Hoy aprendimos", items: ["Ítem 1", "Ítem 2", "Ítem 3"], color: "PURPLE" },
    { title: "Próxima clase",  items: ["Ítem 1", "Ítem 2", "Ítem 3"], color: "BLUE" },
    { title: "Tarea",          items: ["Ítem 1", "Ítem 2", "Ítem 3"], color: "GREEN" }
  ], notes: "..." }
```

### `closing-ceremony` — Ceremonia de cierre / Demo Day (sesión especial)
```js
{ type: "closing-ceremony",
  title: "¡Lo lograron!",
  subtitle: "10 semanas. 30 alumnas. Una promoción de programadoras nuevas.",
  milestones: [
    { label: "HOY PUEDEN", text: "Escribir programas completos en C++ con ciclos, condicionales y funciones." },
    { label: "PUEDEN ANALIZAR", text: "Trazar un algoritmo manualmente y detectar errores de lógica." },
    { label: "PUEDEN APRENDER MÁS", text: "Inscribirse en Programación Competitiva Intermedio." },
  ],
  message: "El siguiente paso no es saberlo todo. Es no parar.",
  notes: "Ceremonia de cierre. Sesión especial — exenta del validador 5E." }
```
Fondo navy con 3 tarjetas color-coded. Exenta del validador 5E.

## Colores disponibles (`C`)

`PURPLE` `#6B32ED`, `NAVY` `#171929`, `YELLOW` `#FDCA36`, `BLUE` `#2B88F7`, `GREEN` `#05A175`, `WHITE` `#FFFFFF`, `ROSE` `#FBD1E9`, `GRAY_MED` `#666666`, `TEAL` `#6FD0D8`, `RED` `#FC3535`.

**Lavandas cálidas** (extraídas del benchmark `Clase-09-Funciones.pptx`):
- `LAVENDER_BG` `#FBF7FF` — blanco con tinte lavanda. Fondo preferido para slides de contenido (en lugar de blanco puro `#FFFFFF`).
- `LAVENDER_LIGHT` `#EAE0FE` — fondo de `pull-quote` y de las tarjetas de `analogy` / `bridge`.

Pasa el nombre en mayúsculas (`"PURPLE"`) o un hex literal (`"6B32ED"`).

## Iconografía (`BRAND_ICONS`)

Reemplaza a Font Awesome. Iconos kawaii brand-compliant (trazo 2pt, puntas redondeadas):

**Iconos pedagógicos:** `lightbulb` (teal), `calculator` (rojo), `flask` (rosa), `globe` (verde), `star` (amarilla), `check` (verde), `exclamation` (rojo), `quote` (amarillo), `target` (púrpura), `tools` (púrpura).

**Elementos pixel-art brandbook** (usados internamente por `title`, `section`, `welcome`):
- `pixel-smiley` — cara sonriente blanca con sombra rosa
- `pixel-cursor` — cursor estilo Windows-95
- `letter-a-frame` — letra "A" azul con marco rectangular y squiggle amarillo

> **Deprecated:** los iconos Font Awesome (`code`, `sitemap`, `arrow`, `puzzle`, `search`, `file`, `diagram`) siguen funcionando como fallback pero no son brand-compliant. No usar en slides nuevas.

## Reglas brandbook que el generador hace cumplir

- **Fuentes:** Space Grotesk (texto/títulos), Space Mono (código). Logo usa Ciutadella Rounded / Boodle pero esas fuentes son solo del logo, no del contenido.
- **Tamaños:** cuerpo mínimo 18pt (proyección), título 32–40pt. El generador usa estos valores por defecto.
- **Máximo 5 líneas visibles** en `engage`, `content`, `activity` (warning si excedes).
- **Logo:** esquina superior derecha. Versión blanca sobre fondos oscuros, versión de color sobre fondos blancos. El generador lo carga desde `Logo.png` o `logoPath`.
- **Combinaciones legibles** (brandbook tabla `[X,Y,Z]`): el generador emite warning si detecta combinación poco legible. Combos seguras `[1,1,1]`: blanco/púrpura, blanco/navy, blanco/azul, blanco/verde, blanco/rojo, navy/blanco, navy/amarillo, navy/rosa.
- **Estética retro/Windows-95:** marco pixel-art presente en `title`, `section` y `plan-b`. Squiggle amarillo en `title` y `section`.

## Validación pedagógica (`validateSlideSequence`)

Se ejecuta automáticamente. Emite warnings (no errores) cuando:
- Slide 0 no es `title`.
- `title` no tiene `tagline` poético (regla `tone-and-narrative-guide.md`).
- Falta `bridge` en las primeras 4 slides (excepto C01) — conector con clase anterior.
- Falta `bridge` con `direction: "forward"` antes del `closing` — preview de próxima clase.
- Falta `analogy` — toda clase requiere al menos una analogía humana.
- Primera slide de contenido no es `engage` (rompe la regla "no cold lecture").
- Falta `objectives` en la secuencia.
- Falta `steam` o `reflection` o `closing`.
- `steam` aparece después de `closing`.
- Slide `steam` sin `connectionToClass`.
- Cualquier `engage`/`content`/`activity` excede 5 líneas visibles.

Para silenciar (no recomendado): `generateSlides({ ..., validate: false })`.
