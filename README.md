# skill-slide-generator · Niñas Pro

Generador programático de presentaciones `.pptx` para Niñas Pro. Produce slides alineados con el **Brandbook Niñas Pro 2022** y los marcos pedagógicos **5E + Bloom + DUA**, con validación automática.

---

## Qué hace esta skill

Convierte una descripción declarativa de slides (JavaScript) en un archivo `.pptx` real, listo para proyectar:

```js
const { generateSlides, buildClassSkeleton } = require(".agents/skills/skill-slide-generator/generate.js");

await generateSlides({
  className: "1",
  classTitle: "Algoritmos y lógica secuencial",
  outputPath: "class-01/slides.pptx",
  slides: [ /* array de slides */ ]
});
```

Aplica automáticamente:

- **Identidad visual brandbook** — paleta de colores, fuentes Space Grotesk / Space Mono, marcos retro Windows-95, iconografía kawaii, pixel-smiley, squiggle amarillo, estrellas decorativas.
- **Estructura pedagógica** — valida que la secuencia tenga `title`, `engage`, `objectives`, `steam` con conexión a la clase, `bridge` backward y forward, `farewell` al cierre.
- **Reglas de legibilidad** — combinaciones de color, fuente mínima 18pt, máximo 5 líneas visibles por slide.
- **Logo en cada slide**, breadcrumb consistente `[NN]▾ / SECCIÓN`, fondo lavanda cálido en lugar de blanco puro.

---

## Cuándo usar esta skill

Invoca esta skill cuando alguien diga:

- "Generar el `.pptx` de esta clase"
- "Regenerar las slides de la Clase NN"
- "Crear las slides del Demo Day"
- "Convertir este `01-slides.md` a `.pptx`"

**No la uses** para:
- Crear slides desde cero sin un brief pedagógico — eso lo hace `skill-create-course` primero.
- Slides que no sean Niñas Pro — la identidad visual está hardcodeada.

---

## Quick start

```bash
# Instala dependencias una sola vez en el repo del curso
npm install pptxgenjs react react-dom sharp react-icons

# Genera un .pptx ejecutando un create-slides.js
cd "Programación Competitiva Básico"
node class-01/create-slides.js
# → escribe class-01/slides.pptx
```

Un `create-slides.js` típico:

```js
const path = require("path");
const { generateSlides, buildClassSkeleton } = require("../../.agents/skills/skill-slide-generator/generate.js");

const base = buildClassSkeleton({
  classNumber: 1,
  title: "Algoritmos y lógica secuencial",
  tagline: "Pensar como una programadora es dar instrucciones tan claras que cualquiera pueda seguirlas.",
  objectives: [
    { verb: "Definir",     content: "qué es un algoritmo", bloom: 2 },
    { verb: "Representar", content: "un algoritmo en pseudocódigo",  bloom: 3 },
  ],
  steamWoman: {
    name: "Grace Hopper",
    years: "1906 – 1992",
    area: "Ciencias de la computación",
    country: "Estados Unidos",
    connectionToClass: "Grace inventó el primer compilador en 1952…",
    facts: ["Doctora en matemáticas de Yale (1934)", "Inventó el primer compilador en 1952"],
    quote: "Lo más peligroso que puedes decir es: \"siempre lo hemos hecho así.\"",
    quoteAttribution: "Grace Hopper, discursos públicos documentados",
    reflectionQuestion: "¿En qué parte de tu vida estás haciendo algo \"porque siempre se ha hecho así\"?",
    source: "Wikipedia + Computer History Museum",
  },
  nextClassTitle: "Introducción al lenguaje C++",
  nextClassWhy: "Tu pseudocódigo se convierte en código real",
  hoyAprendimos: ["Qué es un algoritmo", "Pseudocódigo", "Diagramas de flujo"],
});

// Insertar slides de contenido entre objectives/steam y bridge-forward
// base.splice(5, 0, { type: "analogy", ... }, { type: "code", ... });

await generateSlides({
  className: "1",
  classTitle: "Algoritmos y lógica secuencial",
  outputPath: path.join(__dirname, "slides.pptx"),
  slides: base,
});
```

---

## Tipos de slide disponibles (25+)

Detalle en `SKILL.md`. Resumen:

**Estructurales:**
- `title` — portada con tagline poético
- `bridge` — conector entre clases (backward / forward) en dos columnas
- `section` — divisor de capítulo estilo brandbook
- `welcome` — sesión de bienvenida con agenda numerada
- `closing-ceremony` — Demo Day con hitos color-coded
- `block-transition` — transición entre bloques 5E

**Pedagógicos:**
- `engage` — pregunta gancho con icono
- `objectives` — objetivos Bloom con cuadrados rosa numerados
- `analogy` — VIDA REAL ↔ TÉRMINO TÉCNICO (la slide donde la clase deja de sentirse fría)
- `pull-quote` — frase poética sola
- `definition` — definición + 3 tarjetas de características
- `anatomy` — descomposición en 4 cuadrantes color-coded
- `pillars` — lista vertical con barras de color
- `steps` — lista numerada con cuadrados rosa
- `comparison` — antes / después lado a lado (con código)
- `try-it` — invitación cálida antes de la práctica
- `glossary` — 2-4 términos clave
- `rubric` — rúbrica formativa 1-10

**Código:**
- `code` — bloque de código con filename pill y auto-tamaño según líneas
- `error-code` — código con error + protocolo para mentora
- `manual-trace` — código + tabla de estado de variables paso a paso
- `flowchart` — leyenda de símbolos de diagrama de flujo
- `flowchart-diagram` — flowchart real con nodos + flechas

**Especiales:**
- `steam` — Mujer STEAM al inicio con conexión a la clase
- `reference` — tarjeta de referencia imprimible
- `kahoot` — slide de transición (uso desaconsejado — mentora anuncia verbalmente)
- `plan-b` — Plan B charla inspiracional (ventana modal Windows-95)
- `farewell` — cierre rico con pixel-smiley + 3 logros + bridge a próxima clase

---

## Validador pedagógico

`generateSlides` ejecuta automáticamente `validateSlideSequence` y emite warnings en consola si:

- Slide 0 no es `title`.
- `title` no tiene `tagline` poético.
- Falta `bridge` backward (excepto C01).
- Falta `bridge` forward antes del cierre.
- Falta `analogy` (toda clase requiere al menos una).
- Primera slide de contenido no es `engage` ("no cold lecture").
- Falta `objectives` o `steam`.
- Slide `steam` sin `connectionToClass`.
- `engage`/`content`/`activity` excede 5 líneas visibles.

Los warnings **no bloquean** la generación. Sesiones especiales (`welcome`, `closing-ceremony`) están exentas.

Para silenciar: `generateSlides({ ..., validate: false })`.

---

## Helpers exportados

```js
const {
  generateSlides,         // función principal — genera el .pptx
  buildClassSkeleton,     // helper — devuelve un array de slides 5E-válido
  validateSlideSequence,  // validador — corre solo, emite warnings
  C,                      // constantes de color del brandbook
  CODE_TOKENS,            // tokens nombrados de syntax highlighting
  BRAND_ICONS,            // lista de iconos kawaii disponibles
} = require(".agents/skills/skill-slide-generator/generate.js");
```

---

## Paleta de colores

```js
C.PURPLE        // #6B32ED — color primario
C.NAVY          // #171929 — fondo oscuro
C.YELLOW        // #FDCA36 — acento energético
C.BLUE          // #2B88F7 — secundario
C.GREEN         // #05A175 — acento positivo
C.WHITE         // #FFFFFF
C.ROSE          // #FBD1E9 — tarjetas, elementos rosa
C.GRAY_MED      // #666666 — texto secundario
C.TEAL          // #6FD0D8 — acento fresco
C.RED           // #FC3535 — alertas
C.LAVENDER_BG   // #FBF7FF — fondo cálido (preferido sobre blanco puro)
C.LAVENDER_LIGHT // #EAE0FE — pull-quote, analogía, bridge
```

Pasa el nombre en mayúsculas (`"PURPLE"`) o el hex literal (`"6B32ED"`).

---

## Tokens de código

```js
{ text: "for ", token: "keyword" }   // → amarillo #FDCA36
{ text: "int",  token: "type" }      // → azul #2B88F7
{ text: "\"hola\"", token: "string" } // → verde #05A175
{ text: "// comentario", token: "comment" } // → gris #666666
{ text: "1000", token: "number" }    // → teal #6FD0D8
{ text: "error", token: "error" }    // → rojo #FC3535
```

---

## Archivos de la skill

```
skill-slide-generator/
├── README.md       ← este archivo
├── SKILL.md        ← documentación completa de tipos y parámetros (instrucciones Claude)
├── generate.js     ← el generador (~2400 líneas, 25+ builders)
└── example.js      ← demo ejecutable de todos los tipos
```

```bash
# Ejecutar el demo
node .agents/skills/skill-slide-generator/example.js
# → genera example-clase.pptx, example-bienvenida.pptx, example-cierre.pptx
```

---

## Reglas brandbook que el generador hace cumplir

- **Fuentes:** Space Grotesk (texto), Space Mono (código). Ciutadella Rounded y Boodle son solo del logo.
- **Tamaños:** cuerpo mínimo 18pt en proyector. Auto-scaling para que el código siempre quepa.
- **Logo:** esquina superior derecha en TODAS las slides (aplicado globalmente al final del builder).
- **Máximo 5 líneas visibles** en `engage`, `content`, `activity` (warning si excedes).
- **Combinaciones legibles:** tabla `[X, Y, Z]` del brandbook validada automáticamente.
- **Estética retro/Windows-95:** marco pixel-art con smiley + cursor en `title`, `section`, `welcome`, `farewell`.
- **Breadcrumb consistente:** `[NN]▾` arriba + `SECCIÓN` debajo en TODA slide con `sectionLabel`.

---

## Integración con otras skills

| Skill | Rol |
|-------|-----|
| `skill-create-course` | Produce los `create-slides.js` que esta skill ejecuta. |

`skill-create-course` apunta a esta skill en su workflow paso 5: tras generar el `01-slides.md`, crea el `create-slides.js` y lo ejecuta para producir el `.pptx`.

---

## Dependencias

- `pptxgenjs` — librería principal de generación `.pptx`
- `react` + `react-dom` + `react-icons` — fallback de iconos (deprecated, ver SKILL.md)
- `sharp` — conversión SVG → PNG para iconos kawaii inline

```bash
npm install pptxgenjs react react-dom sharp react-icons
```

---

## Documentación detallada

- **`SKILL.md`** — referencia completa de cada tipo de slide, parámetros, ejemplos, reglas. Es la fuente de verdad técnica.
- **`example.js`** — demo ejecutable con 3 escenarios (clase, bienvenida, Demo Day).

---

## Benchmark visual

El archivo `Clase-09-Funciones.pptx` (en la raíz del repo del curso) es el norte de cómo deben verse las slides. Si una clase generada con esta skill no llega al mismo nivel pedagógico-visual, revisar `references/tone-and-narrative-guide.md` de `skill-create-course`.

---

## Fuente primaria

- **Brandbook Niñas Pro 2022** — `Brandbook Niñas Pro 2022.pdf`.
- **pptxgenjs** — https://gitbrent.github.io/PptxGenJS/
# skill-create-slide
