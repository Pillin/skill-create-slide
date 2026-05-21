const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaLightbulb, FaCode, FaSitemap, FaPuzzlePiece, FaTools,
  FaArrowRight, FaCheckCircle, FaExclamationTriangle, FaTarget,
  FaQuoteLeft, FaStar, FaSearch, FaFileAlt, FaProjectDiagram
} = require("react-icons/fa");

// ─── Brand constants (Niñas Pro Brandbook 2022) ───
const C = {
  PURPLE: "6B32ED",
  NAVY: "171929",
  YELLOW: "FDCA36",
  BLUE: "2B88F7",
  GREEN: "05A175",
  WHITE: "FFFFFF",
  ROSE: "FBD1E9",
  GRAY_MED: "666666",
  TEAL: "6FD0D8",
  RED: "FC3535",
  // Warm lavender tints (extracted from Clase-09-Funciones.pptx — the reference example)
  LAVENDER_BG: "FBF7FF",     // off-white with warm lavender — preferred over pure white for content slides
  LAVENDER_LIGHT: "EAE0FE",  // pull-quote / analogy card background
  CODE_DEFAULT: "D4D4D4",
  FONT: "Space Grotesk",
  FONT_CODE: "Space Mono",
};

// Code token colors (per brandbook: keyword=yellow|blue, string=green, comment=gray)
const CODE_TOKENS = {
  keyword: C.YELLOW,
  type: C.BLUE,
  string: C.GREEN,
  number: C.TEAL,
  comment: C.GRAY_MED,
  error: C.RED,
};

// Legibility matrix — brandbook page on color combinations
// Notation: [text≤17pt, text≥18pt, icons/graphics]. 1 = legible, 0 = not.
const LEGIBLE_COMBOS = {
  // White-on-color: all safe combos
  [`FFFFFF/${C.PURPLE}`]: [1, 1, 1],
  [`FFFFFF/${C.NAVY}`]: [1, 1, 1],
  [`FFFFFF/${C.BLUE}`]: [1, 1, 1],
  [`FFFFFF/${C.GREEN}`]: [1, 1, 1],
  [`FFFFFF/${C.RED}`]: [1, 1, 1],
  [`FFFFFF/${C.GRAY_MED}`]: [0, 1, 1],
  // Navy-on-light: safe for body text
  [`${C.NAVY}/FFFFFF`]: [1, 1, 1],
  [`${C.NAVY}/${C.YELLOW}`]: [1, 1, 1],
  [`${C.NAVY}/${C.ROSE}`]: [1, 1, 1],
  [`${C.NAVY}/${C.TEAL}`]: [1, 1, 1],
  // Low-contrast combos (large text and icons only)
  [`${C.PURPLE}/FFFFFF`]: [0, 1, 1],
  [`${C.GREEN}/FFFFFF`]: [0, 1, 1],
  [`${C.BLUE}/FFFFFF`]: [0, 1, 1],
};

function checkLegibility(textHex, bgHex, fontSizePt, label) {
  const key = `${(textHex || "").toUpperCase()}/${(bgHex || "").toUpperCase()}`;
  const combo = LEGIBLE_COMBOS[key];
  if (!combo) return; // unknown combo — skip silently
  const idx = fontSizePt < 18 ? 0 : 1;
  if (!combo[idx]) {
    console.warn(
      `[brandbook] Combinación poco legible: ${textHex} sobre ${bgHex} ` +
      `con texto ${fontSizePt}pt (${label || "slide"}). ` +
      `Brandbook clasifica esta combinación como ${combo.join(",")} (texto≤17, texto≥18, íconos).`
    );
  }
}

// ─── Brand icon SVGs (kawaii style, 2pt stroke, rounded ends) ───
// Replaces Font Awesome to match Brandbook 2022 illustration system.
function brandIconSvg(name, sizePx) {
  const sz = sizePx || 256;
  const sw = Math.max(2, Math.round(sz / 64) * 2); // approximate 2pt at scale
  // Stroke attributes only — fill is set per shape to avoid duplicate-attr errors in strict XML parsers.
  const stroke = `stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round"`;
  switch (name) {
    case "lightbulb": // bombilla — teal
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.TEAL}" ${stroke}>
          <path fill="#${C.TEAL}" fill-opacity="0.2" d="M32 8c-9 0-15 7-15 15 0 6 3 10 6 13 1 1 2 3 2 5v3h14v-3c0-2 1-4 2-5 3-3 6-7 6-13 0-8-6-15-15-15z"/>
          <line fill="none" x1="23" y1="52" x2="41" y2="52"/>
          <line fill="none" x1="25" y1="56" x2="39" y2="56"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="27" cy="22" r="2"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="37" cy="22" r="2"/>
        </g>
      </svg>`;
    case "calculator": // calculadora — rojo
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.NAVY}" ${stroke}>
          <rect fill="#${C.RED}" x="14" y="8" width="36" height="48" rx="4"/>
          <rect fill="#FFFFFF" stroke="none" x="20" y="14" width="24" height="10" rx="1"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="25" cy="33" r="2.5"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="32" cy="33" r="2.5"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="39" cy="33" r="2.5"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="25" cy="42" r="2.5"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="32" cy="42" r="2.5"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="39" cy="42" r="2.5"/>
        </g>
      </svg>`;
    case "flask": // matraz — rosa
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.NAVY}" ${stroke}>
          <path fill="#${C.ROSE}" d="M26 8h12v14l10 24c2 4-1 10-6 10H20c-5 0-8-6-6-10l12-24V8z"/>
          <line fill="none" x1="26" y1="14" x2="38" y2="14"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="28" cy="38" r="2"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="36" cy="42" r="2"/>
          <circle fill="#${C.NAVY}" stroke="none" cx="32" cy="48" r="2"/>
        </g>
      </svg>`;
    case "globe": // globo terráqueo — verde
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.NAVY}" ${stroke}>
          <circle fill="#${C.GREEN}" cx="32" cy="32" r="22"/>
          <ellipse fill="none" cx="32" cy="32" rx="22" ry="9"/>
          <ellipse fill="none" cx="32" cy="32" rx="9" ry="22"/>
          <line fill="none" x1="10" y1="32" x2="54" y2="32"/>
          <line fill="none" x1="32" y1="10" x2="32" y2="54"/>
        </g>
      </svg>`;
    case "star": // estrella de cuatro puntas — amarilla
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.NAVY}" ${stroke}>
          <path fill="#${C.YELLOW}" d="M32 6 L36 28 L58 32 L36 36 L32 58 L28 36 L6 32 L28 28 Z"/>
        </g>
      </svg>`;
    case "check": // ícono check — verde
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.GREEN}" ${stroke}>
          <circle fill="#${C.GREEN}" fill-opacity="0.15" cx="32" cy="32" r="24"/>
          <polyline fill="none" points="20,33 29,42 45,24"/>
        </g>
      </svg>`;
    case "exclamation": // alerta — rojo
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.RED}" ${stroke}>
          <path fill="#${C.RED}" fill-opacity="0.15" d="M32 6 L58 54 L6 54 Z"/>
          <line fill="none" x1="32" y1="22" x2="32" y2="40"/>
          <circle fill="#${C.RED}" stroke="none" cx="32" cy="47" r="2.5"/>
        </g>
      </svg>`;
    case "quote": // cita — amarillo
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.YELLOW}" ${stroke}>
          <path fill="#${C.YELLOW}" d="M14 22 c0-6 4-10 10-10 v6 c-2 0-4 2-4 4 v2 h6 v14 H14 z"/>
          <path fill="#${C.YELLOW}" d="M34 22 c0-6 4-10 10-10 v6 c-2 0-4 2-4 4 v2 h6 v14 H34 z"/>
        </g>
      </svg>`;
    case "target": // objetivo — púrpura
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.PURPLE}" ${stroke}>
          <circle fill="#${C.PURPLE}" fill-opacity="0.15" cx="32" cy="32" r="22"/>
          <circle fill="none" cx="32" cy="32" r="14"/>
          <circle fill="#${C.PURPLE}" stroke="none" cx="32" cy="32" r="6"/>
        </g>
      </svg>`;
    case "tools": // herramientas — púrpura
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <g stroke="#${C.PURPLE}" ${stroke}>
          <rect fill="#${C.PURPLE}" fill-opacity="0.2" x="22" y="8" width="6" height="32" rx="2"/>
          <rect fill="#${C.PURPLE}" fill-opacity="0.2" x="36" y="8" width="6" height="32" rx="2"/>
          <rect fill="#${C.PURPLE}" fill-opacity="0.2" x="18" y="42" width="28" height="14" rx="2"/>
        </g>
      </svg>`;
    case "pixel-smiley": // cara sonriente pixel-art (brandbook signature)
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <rect fill="#${C.ROSE}" x="44" y="14" width="16" height="40"/>
        <rect fill="#${C.ROSE}" x="48" y="50" width="6" height="6"/>
        <circle fill="#FFFFFF" stroke="#${C.NAVY}" stroke-width="4" cx="32" cy="32" r="22"/>
        <rect fill="#${C.NAVY}" x="22" y="22" width="6" height="8"/>
        <rect fill="#${C.NAVY}" x="36" y="22" width="6" height="8"/>
        <path fill="none" stroke="#${C.NAVY}" stroke-width="4" stroke-linecap="round" d="M 20 38 Q 32 50 44 38"/>
      </svg>`;
    case "pixel-cursor": // cursor estilo Windows-95
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="${sz}" height="${sz}">
        <polygon fill="#FFFFFF" stroke="#${C.NAVY}" stroke-width="2" stroke-linejoin="round" points="4,3 4,24 10,18 13,28 17,26 14,16 22,16"/>
      </svg>`;
    case "letter-a-frame": // composición letra A con marco rectangular y squiggle (brandbook §04)
      return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="${sz}" height="${sz}">
        <rect fill="none" stroke="#FFFFFF" stroke-width="2" x="14" y="22" width="36" height="32"/>
        <path fill="none" stroke="#${C.YELLOW}" stroke-width="4" stroke-linecap="round" d="M 18 14 q 6 -6 12 0 q 6 6 12 0 q 6 -6 12 0"/>
        <text x="32" y="58" font-family="Space Grotesk, Arial, sans-serif" font-size="48" font-weight="bold" fill="#${C.BLUE}" text-anchor="middle">A</text>
      </svg>`;
    default:
      return null;
  }
}

const BRAND_ICONS = ["lightbulb", "calculator", "flask", "globe", "star", "check", "exclamation", "quote", "target", "tools", "pixel-smiley", "pixel-cursor", "letter-a-frame"];

// Legacy Font Awesome map — DEPRECATED. Kept for backwards compatibility with
// existing class slides. New code should use brand icons (BRAND_ICONS above).
const faIconMap = {
  lightbulb: FaLightbulb, tools: FaTools, exclamation: FaExclamationTriangle,
  quote: FaQuoteLeft, check: FaCheckCircle, target: FaTarget,
  code: FaCode, sitemap: FaSitemap, arrow: FaArrowRight,
  puzzle: FaPuzzlePiece, search: FaSearch, file: FaFileAlt,
  diagram: FaProjectDiagram, star: FaStar,
};

function colorVal(name) {
  if (!name) return null;
  const upper = name.toUpperCase();
  return C[upper] || name;
}

async function svgToDataUri(svg) {
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "data:image/png;base64," + buf.toString("base64");
}

async function iconToDataUri(iconName, colorOverride, size) {
  // Brand icon first
  if (BRAND_ICONS.includes(iconName)) {
    return svgToDataUri(brandIconSvg(iconName, size || 256));
  }
  // Fallback to legacy Font Awesome
  const Icon = faIconMap[iconName];
  if (!Icon) return null;
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Icon, { color: colorOverride || "#000000", size: String(size || 256) })
  );
  return svgToDataUri(svg);
}

// ─── Visible-content limits (brandbook: max 5 lines, min 18pt body) ───
function countVisibleLines(body) {
  if (!body) return 0;
  if (typeof body === "string") return body.split("\n").length;
  if (!Array.isArray(body)) return 1;
  return body.reduce((n, item) => {
    const text = (item && item.text) ? item.text : (typeof item === "string" ? item : "");
    return n + Math.max(1, text.split("\n").length);
  }, 0);
}

function warnIfTooDense(slideType, body, max) {
  const limit = max || 5;
  const n = countVisibleLines(body);
  if (n > limit) {
    console.warn(
      `[brandbook] Slide '${slideType}' tiene ${n} líneas visibles (máx ${limit}). ` +
      `Mueve detalle a las speaker notes.`
    );
  }
}

// ─── 5E sequence validator ───
// Validates standard class structure. Special sessions (welcome / closing-ceremony)
// are exempt — they have their own conventions per facilitation-guide.md.
function validateSlideSequence(slides) {
  if (!Array.isArray(slides) || slides.length === 0) return;
  const types = slides.map(s => s && s.type);
  const warn = (msg) => console.warn(`[pedagogía] ${msg}`);

  // Detect session kind from slide types
  const isWelcome = types.includes("welcome") && !types.includes("objectives");
  const isClosingCeremony = types.includes("closing-ceremony") && !types.includes("objectives");
  if (isWelcome || isClosingCeremony) {
    // Special sessions: only verify there's a title and a logo-bearing first slide
    if (types[0] !== "title" && types[0] !== "welcome" && types[0] !== "closing-ceremony") {
      warn(`Sesión especial: slide 0 debería ser 'title', 'welcome' o 'closing-ceremony'. Es '${types[0]}'.`);
    }
    return;
  }

  // Standard class validation
  if (types[0] !== "title") {
    warn(`Slide 0 debería ser 'title', es '${types[0]}'.`);
  }
  // Title tagline check (per tone-and-narrative-guide.md)
  if (slides[0] && slides[0].type === "title" && !slides[0].tagline) {
    warn(`Slide 0 (title): falta 'tagline' poético — el guide de tono lo marca como no-negociable.`);
  }
  // Skip bridge / block-transition / section slides when looking for first content
  const structural = new Set(["title", "section", "block-transition", "bridge"]);
  const firstContentIdx = types.findIndex((t, i) => i > 0 && !structural.has(t));
  if (firstContentIdx >= 0 && types[firstContentIdx] !== "engage") {
    warn(`Primera slide de contenido debería ser 'engage' (regla "no cold lecture"). Es '${types[firstContentIdx]}'.`);
  }
  // Bridge backward (slide 02) — non-negotiable except for class 01
  const classNumber = slides[0] && slides[0].classNumber;
  const isClassOne = String(classNumber) === "1" || String(classNumber) === "01";
  if (!isClassOne && !types.slice(0, 4).includes("bridge")) {
    warn(`Falta 'bridge' en las primeras 4 slides — toda clase (excepto C01) debe abrir con un conector "LO QUE YA SABES / LO QUE VIENE HOY". Ver tone-and-narrative-guide.md.`);
  }
  // At least one analogy per class — warm tone requirement
  if (!types.includes("analogy")) {
    warn(`Falta slide 'analogy' — toda clase debe tener al menos una analogía cálida de la vida real (VIDA REAL ↔ TÉRMINO TÉCNICO). Ver tone-and-narrative-guide.md.`);
  }
  if (!types.includes("objectives")) {
    warn(`Falta slide 'objectives' (objetivos de aprendizaje con verbos Bloom observables).`);
  }
  if (!types.includes("steam")) {
    warn(`Falta slide 'steam' (mujer STEAM por clase es no-negociable).`);
  }
  // STEAM goes at the BEGINNING of the class (after objectives, before lecture content).
  // No warning if it appears before closing — that's the new structure.
  // `reflection` slide is opcional.
  // (legacy check removed — steam at the beginning is now the recommended pattern)
  // Bridge forward — recommended near the end (before farewell)
  const farewellIdx = types.lastIndexOf("farewell");
  const forwardBridgeIdx = slides.findIndex(s => s && s.type === "bridge" && s.direction === "forward");
  if (farewellIdx >= 0 && forwardBridgeIdx < 0) {
    warn(`Falta 'bridge' con direction: "forward" antes del farewell — toda clase debe cerrar con preview de la siguiente.`);
  }
  // STEAM woman: warn if pedagogically-critical fields are missing on any steam slide
  slides.forEach((s, i) => {
    if (s && s.type === "steam" && !s.connectionToClass) {
      warn(`Slide ${i} (steam): falta 'connectionToClass' — el guide pide conectar explícitamente con el contenido de la clase.`);
    }
  });
  // Warm closing on objectives
  slides.forEach((s, i) => {
    if (s && s.type === "objectives" && !s.warmClosing && !s.subtitle) {
      // Soft hint — objectives builder has a default, but a custom one is better
    }
  });
}

// ─── Pedagogical skeleton builder ───
// Returns a valid 5E-compliant + narrative-connected slide sequence per
// tone-and-narrative-guide.md. Insert content slides between objectives and kahoot.
function buildClassSkeleton(opts) {
  const {
    classNumber, title, tagline, subtitle, courseName,
    previousKnowledge, todayContent, // for the backward bridge
    nextClassTitle, nextClassTagline, nextClassWhy, // for the forward bridge
    objectives, warmClosing,
    analogy, // optional inline analogy slide content
    steamWoman, glossary, hoyAprendimos, proximaClase, tarea,
  } = opts;
  const isClassOne = String(classNumber) === "1" || String(classNumber) === "01";
  const slides = [
    {
      type: "title",
      classNumber: String(classNumber),
      title,
      tagline: tagline || "<Sustituye con una frase poética que comprima la idea central de la clase>",
      subtitle: subtitle || courseName || "",
      notes: "",
    },
  ];
  // Bridge backward — non-negotiable except for class 01
  if (!isClassOne) {
    slides.push({
      type: "bridge",
      sectionLabel: "00 · APERTURA",
      direction: "backward",
      title: title,
      subtitle: "Lo que ya sabes y lo que viene hoy.",
      leftLabel: "LO QUE YA SABES",
      leftSubtitle: `De la clase ${parseInt(classNumber, 10) - 1} y anteriores:`,
      left: previousKnowledge || ["<concepto 1>", "<concepto 2>", "<concepto 3>"],
      rightLabel: "LO QUE VIENE HOY",
      right: todayContent || ["<contenido 1>", "<contenido 2>", "<contenido 3>"],
      notes: "BRIDGE backward. Conecta explícitamente con la clase anterior. Ver tone-and-narrative-guide.md.",
    });
  }
  slides.push(
    {
      type: "engage",
      title: "¿…?",
      icon: "lightbulb",
      body: ["Sustituye este body con una pregunta gancho conectada a la experiencia de las alumnas.", { text: "Nunca abrir con cátedra fría.", bold: true }],
      notes: "ENGAGE. Pregunta gancho. Levantar manos / pedir un ejemplo.",
    },
    {
      type: "objectives",
      sectionLabel: "00 · APERTURA",
      title: "Objetivos de hoy",
      subtitle: "Al final de hoy podrás…",
      objectives: objectives || [
        { verb: "Identificar", content: "<contenido>", bloom: 2 },
        { verb: "Aplicar", content: "<contenido>", bloom: 3 },
        { verb: "Analizar", content: "<contenido>", bloom: 4 },
      ],
      warmClosing: warmClosing || "Si al final puedes hacer estas cosas, la clase fue exitosa.",
      notes: "EXPLAIN setup. Leer los objetivos. Recordar que son verbos observables.",
    },
  );
  // STEAM woman al INICIO de la clase (después de objetivos, antes de la cátedra).
  // Per feedback de la usuaria: la mujer STEAM da el contexto y la motivación,
  // no es un epílogo. Es la referente que justifica el contenido.
  slides.push({
    type: "steam",
    name: (steamWoman && steamWoman.name) || "<Mujer STEAM>",
    years: steamWoman && steamWoman.years,
    area: steamWoman && steamWoman.area,
    country: steamWoman && steamWoman.country,
    connectionToClass: steamWoman && steamWoman.connectionToClass,
    facts: (steamWoman && steamWoman.facts) || [],
    quote: steamWoman && steamWoman.quote,
    quoteAttribution: steamWoman && steamWoman.quoteAttribution,
    reflectionQuestion: steamWoman && steamWoman.reflectionQuestion,
    source: steamWoman && steamWoman.source,
    bgColor: (steamWoman && steamWoman.bgColor) || "PURPLE",
    notes: "STEAM al inicio. Da contexto y motivación para el contenido.",
  });
  // Optional inline analogy (recommended — at least one analogy per class)
  if (analogy) {
    slides.push({
      type: "analogy",
      sectionLabel: analogy.sectionLabel || "01 · IDEA",
      title: analogy.title || "Una analogía",
      subtitle: analogy.subtitle,
      leftLabel: analogy.leftLabel || "VIDA REAL",
      rightLabel: analogy.rightLabel || "EN PROGRAMACIÓN",
      rows: analogy.rows || [],
      notes: analogy.notes || "EXPLAIN/ENGAGE. Analogía cálida. El momento donde la clase deja de sentirse fría.",
    });
  }
  // Contenido de la cátedra va aquí — insertar slides de explain/elaborate después de este punto
  // NOTA: el Kahoot y la práctica de Dodona NO aparecen como slides separadas en el .pptx.
  // La mentora los anuncia verbalmente; el flujo de la clase está en las speaker notes.
  // Bridge forward — connects to next class
  if (nextClassTitle || nextClassWhy) {
    slides.push({
      type: "bridge",
      sectionLabel: "06 · CIERRE",
      direction: "forward",
      title: "Lo que viene…",
      subtitle: `Próxima clase: ${nextClassTagline || nextClassTitle || ""}.`,
      leftLabel: "ESTA CLASE TE DIO",
      left: (hoyAprendimos || []).slice(0, 3),
      rightLabel: `CLASE ${parseInt(classNumber, 10) + 1}`,
      rightSubtitle: nextClassTitle || "",
      right: nextClassWhy ? [nextClassWhy] : ["<por qué la próxima clase conecta>"],
      narrative: nextClassWhy ? `Trae lo de hoy bien fijado.` : null,
      notes: "BRIDGE forward. Cierra el loop narrativo con la próxima clase.",
    });
  }
  // Farewell slide — cierre rico, reemplaza al closing tradicional
  const nextClassNum = parseInt(classNumber, 10) + 1;
  slides.push({
    type: "farewell",
    title: "¡Gracias por hoy!",
    subtitle: tagline ? `Lo de hoy: ${tagline.split(".")[0].toLowerCase()}.` : null,
    achievements: (hoyAprendimos || []).slice(0, 3),
    nextClass: nextClassTitle ? `Clase ${nextClassNum} — ${nextClassTitle}` : `Clase ${nextClassNum}`,
    signature: "NIÑAS PRO  ·  PROGRAMACIÓN COMPETITIVA BÁSICO",
    notes: "Cierre cálido. Quedarse 1-2 min después de la última slide mientras se van las alumnas — algunas se acercan a preguntar cosas que no preguntaron en público.",
  });
  // Optional: insert glossary slide before closing
  if (glossary && Array.isArray(glossary) && glossary.length) {
    slides.splice(slides.length - 1, 0, {
      type: "glossary",
      title: "Vocabulario clave de hoy",
      subtitle: "Estos términos están en tu glosario completo (04-glossary.md).",
      terms: glossary,
      notes: "EVALUATE — refuerzo de vocabulario antes del cierre.",
    });
  }
  return slides;
}

async function generateSlides(config) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Niñas Pro";
  pres.title = `${config.className} — ${config.classTitle}`;

  // Run pedagogical sequence validator (warnings only)
  if (config.validate !== false) {
    validateSlideSequence(config.slides);
  }

  // ─── Logo ───
  let LOGO_DATA = null;
  const logoPath = config.logoPath || path.resolve(__dirname, "../../..", "Logo.png");
  try {
    if (fs.existsSync(logoPath)) {
      LOGO_DATA = "image/png;base64," + fs.readFileSync(logoPath).toString("base64");
    }
  } catch (_) { /* logo optional */ }

  function placeLogo(slide) {
    if (!LOGO_DATA) return;
    slide.addImage({
      data: LOGO_DATA, x: 8.3, y: 0.15, w: 0.9, h: 0.9,
      sizing: { type: "contain", w: 0.9, h: 0.9 }
    });
  }

  function placeStar(slide, x, y, sz, c) {
    slide.addText("✦", {
      x, y, w: sz || 0.35, h: sz || 0.35,
      fontSize: (sz || 0.35) * 64, fontFace: C.FONT, color: c || C.YELLOW,
      align: "center", valign: "middle", margin: 0
    });
  }

  function placeAccent(slide, y, c) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y || 0.35, w: 0.06, h: 0.55, fill: { color: colorVal(c) || C.PURPLE }
    });
  }

  // ─── Retro frame (Brandbook Windows-95 aesthetic) ───
  // Draws a pixel-art window outline. Used in title and section dividers.
  function placeRetroFrame(slide, x, y, w, h, opts) {
    const o = opts || {};
    const strokeColor = o.stroke || C.WHITE;
    // Outer frame outline
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h,
      fill: { color: o.bg || C.NAVY, transparency: o.bg ? 0 : 100 },
      line: { color: strokeColor, width: 2 }
    });
    // Title bar
    const tbH = 0.22;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: tbH,
      fill: { color: o.bg || C.NAVY, transparency: o.bg ? 0 : 100 },
      line: { color: strokeColor, width: 2 }
    });
    // Three "circle" dots in title bar (Windows-style)
    for (let i = 0; i < 3; i++) {
      slide.addShape(pres.shapes.OVAL, {
        x: x + 0.08 + i * 0.13, y: y + 0.07, w: 0.08, h: 0.08,
        fill: { color: strokeColor }, line: { color: strokeColor, width: 0 }
      });
    }
  }

  // Yellow squiggle decoration (brandbook recurring element)
  function placeSquiggle(slide, x, y, w, h, color) {
    const c = color || C.YELLOW;
    // Approximate the squiggle with three connected curves using freeform-ish shapes
    const segW = (w || 0.6) / 3;
    const segH = h || 0.3;
    for (let i = 0; i < 3; i++) {
      slide.addShape(pres.shapes.ARC, {
        x: x + i * segW, y: y + (i % 2 === 0 ? 0 : segH / 2),
        w: segW, h: segH,
        fill: { type: "solid", color: c, transparency: 100 },
        line: { color: c, width: 3 }
      });
    }
  }

  // Pink numbered square (brandbook explicit pattern for numbered lists)
  function placePinkNumber(slide, n, x, y, sz) {
    const size = sz || 0.35;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: size, h: size,
      fill: { color: C.ROSE }, line: { color: C.ROSE, width: 0 }
    });
    slide.addText(String(n), {
      x, y, w: size, h: size,
      fontSize: 14, fontFace: C.FONT, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0
    });
  }

  // Section label in brandbook style: [ 0N ]  / SECTION NAME
  function placeSectionLabel(slide, number, label, x, y) {
    slide.addText(`[ ${String(number).padStart(2, "0")} ]`, {
      x, y, w: 1.5, h: 0.35,
      fontSize: 14, fontFace: C.FONT, color: C.NAVY, bold: true, margin: 0
    });
    slide.addText(label, {
      x, y: y + 0.32, w: 4, h: 0.35,
      fontSize: 16, fontFace: C.FONT, color: C.NAVY, bold: true,
      letterSpacing: 2, margin: 0
    });
  }

  // ─── Brand furniture (applied to every content slide) ───
  // Per brandbook §00–§08: each page has section label upper-left + logo upper-right
  // + scattered decorative stars / squiggle. Without this, slides feel generic.
  function placeBrandFurniture(slide, opts) {
    const o = opts || {};
    const darkBg = o.darkBg;
    // Section label upper-left: [NN]▾ / SECTION NAME
    if (o.sectionLabel) {
      // Parse "00 · APERTURA" or "01 · IDEA" into number + name
      const parts = o.sectionLabel.split(/\s*[·\.\-]\s*/);
      const num = (parts[0] || "00").trim();
      const name = (parts[1] || "").trim().toUpperCase();
      const labelColor = darkBg ? C.WHITE : C.NAVY;
      slide.addText([
        { text: `[${num}]`, options: { color: labelColor, bold: true } },
        { text: "▾", options: { color: labelColor } },
      ], {
        x: 0.4, y: 0.12, w: 1.4, h: 0.3,
        fontSize: 13, fontFace: C.FONT, margin: 0
      });
      if (name) {
        slide.addText(name, {
          x: 0.4, y: 0.38, w: 4, h: 0.28,
          fontSize: 11, fontFace: C.FONT, color: labelColor, bold: true,
          letterSpacing: 3, margin: 0
        });
      }
    }
    // Decorative stars in safe corners (don't overlap content)
    if (!o.skipDecorations) {
      const stars = o.stars || [
        { x: 9.45, y: 5.4, sz: 0.28, c: darkBg ? C.YELLOW : C.ROSE },
        { x: 0.2,  y: 5.4, sz: 0.22, c: darkBg ? C.ROSE  : C.YELLOW },
      ];
      stars.forEach(s => placeStar(slide, s.x, s.y, s.sz, s.c));
    }
    // Optional yellow squiggle accent (replaces the plain colored accent bar)
    if (o.squiggle) {
      const sq = o.squiggle === true ? { x: 8.0, y: 0.3 } : o.squiggle;
      for (let i = 0; i < 3; i++) {
        slide.addShape(pres.shapes.ARC, {
          x: sq.x + i * 0.15, y: sq.y + (i % 2 === 0 ? 0 : 0.08),
          w: 0.18, h: 0.18,
          fill: { color: C.YELLOW, transparency: 100 },
          line: { color: C.YELLOW, width: 3 }
        });
      }
    }
  }

  // Pixel-smiley peeking from a corner (brand warmth on content slides)
  async function placeSmileyPeek(slide, position) {
    const data = await iconToDataUri("pixel-smiley", null, 256);
    if (!data) return;
    const positions = {
      "bottom-right": { x: 8.5, y: 4.4, w: 1.0, h: 1.0 },
      "bottom-left":  { x: 0.2, y: 4.4, w: 0.9, h: 0.9 },
      "top-right":    { x: 8.5, y: 0.4, w: 0.9, h: 0.9 },
    };
    const p = positions[position || "bottom-right"];
    slide.addImage({ data, ...p });
  }

  // Windows-95 style modal alert (used by plan-b)
  function placeAlertWindow(slide, x, y, w, h, opts) {
    const o = opts || {};
    const titleBarColor = o.titleBarColor || C.YELLOW;
    const bodyBg = o.bodyBg || C.WHITE;
    // Body
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: y + 0.3, w, h: h - 0.3,
      fill: { color: bodyBg }, line: { color: C.NAVY, width: 2 }
    });
    // Title bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 0.3,
      fill: { color: titleBarColor }, line: { color: C.NAVY, width: 2 }
    });
    slide.addText(o.titleText || "ALERTA", {
      x: x + 0.15, y, w: w - 0.6, h: 0.3,
      fontSize: 14, fontFace: C.FONT, color: C.NAVY, bold: true,
      valign: "middle", margin: 0
    });
    // Close button "x"
    slide.addText("— □ x", {
      x: x + w - 0.7, y, w: 0.55, h: 0.3,
      fontSize: 12, fontFace: C.FONT_CODE, color: C.NAVY,
      align: "right", valign: "middle", margin: 0
    });
  }

  // ─── Slide builders ───

  async function buildTitle(slide, d) {
    slide.background = { color: C.PURPLE };
    placeLogo(slide);
    // Yellow underline accent
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 1.8, w: 2.5, h: 0.05, fill: { color: C.YELLOW }
    });
    slide.addText("CLASE " + d.classNumber, {
      x: 0.6, y: 0.9, w: 4, h: 0.6,
      fontSize: 16, fontFace: C.FONT, color: C.YELLOW, bold: true,
      letterSpacing: 4, margin: 0
    });
    // Title — slightly smaller to leave room for tagline
    const hasTagline = !!d.tagline;
    slide.addText(d.title, {
      x: 0.6, y: 2.0, w: 5.6, h: hasTagline ? 2.2 : 2.8,
      fontSize: hasTagline ? 38 : 42, fontFace: C.FONT, color: C.WHITE, bold: true,
      valign: "middle", margin: 0, lineSpacingMultiple: 1.1
    });
    // Poetic tagline (per tone-and-narrative-guide.md — non-negotiable)
    if (d.tagline) {
      slide.addText(d.tagline, {
        x: 0.6, y: 4.3, w: 5.8, h: 0.7,
        fontSize: 16, fontFace: C.FONT, color: C.YELLOW, italic: true,
        valign: "top", margin: 0, lineSpacingMultiple: 1.3
      });
    }
    slide.addText(d.subtitle || "Programación Competitiva Básico", {
      x: 0.6, y: hasTagline ? 5.1 : 4.8, w: 5.8, h: 0.4,
      fontSize: 13, fontFace: C.FONT, color: "CCCCDD", margin: 0
    });
    // Retro pixel-art composition (right side) — brandbook signature
    // Big frame holding the pixel-smiley
    placeRetroFrame(slide, 6.3, 1.8, 2.0, 2.2, { stroke: C.WHITE });
    const smileyData = await iconToDataUri("pixel-smiley", null, 256);
    if (smileyData) slide.addImage({ data: smileyData, x: 6.55, y: 2.15, w: 1.5, h: 1.5 });
    // Pixel cursor pointing to the smiley
    const cursorData = await iconToDataUri("pixel-cursor", null, 128);
    if (cursorData) slide.addImage({ data: cursorData, x: 8.2, y: 3.3, w: 0.5, h: 0.5 });
    // Letter-A frame composition with squiggle
    placeRetroFrame(slide, 7.4, 4.1, 1.7, 1.4, { stroke: C.WHITE });
    const letterAData = await iconToDataUri("letter-a-frame", null, 256);
    if (letterAData) slide.addImage({ data: letterAData, x: 7.55, y: 4.25, w: 1.4, h: 1.2 });
    placeStar(slide, 6.0, 4.6, 0.3, C.ROSE);
    placeStar(slide, 8.8, 1.4, 0.25, C.YELLOW);
    slide.addNotes(d.notes || "");
  }

  async function buildSection(slide, d) {
    // Chapter-divider style straight from brandbook: navy bg, retro frame with smiley, label
    slide.background = { color: C.NAVY };
    placeLogo(slide);
    // Big retro frame on the left half — holds smiley + label (brandbook signature)
    placeRetroFrame(slide, 0.8, 1.4, 3.6, 3.6, { stroke: C.WHITE });
    const smileyData = await iconToDataUri("pixel-smiley", null, 256);
    if (smileyData) slide.addImage({ data: smileyData, x: 1.6, y: 1.9, w: 2.0, h: 2.0 });
    // Section label inside frame: [ 0N ] / NAME
    slide.addText(`[ ${String(d.number || "01").padStart(2, "0")} ]`, {
      x: 1.0, y: 4.0, w: 2.5, h: 0.4,
      fontSize: 22, fontFace: C.FONT, color: C.WHITE, bold: true, margin: 0
    });
    slide.addText((d.label || "SECCIÓN").toUpperCase(), {
      x: 1.0, y: 4.4, w: 3.4, h: 0.5,
      fontSize: 26, fontFace: C.FONT, color: C.WHITE, bold: true,
      letterSpacing: 2, margin: 0
    });
    // Cursor pointing to smiley
    const cursorData = await iconToDataUri("pixel-cursor", null, 128);
    if (cursorData) slide.addImage({ data: cursorData, x: 4.4, y: 3.2, w: 0.5, h: 0.5 });
    // Letter-A + squiggle composition on the right (brandbook signature)
    const letterAData = await iconToDataUri("letter-a-frame", null, 256);
    if (letterAData) slide.addImage({ data: letterAData, x: 6.0, y: 2.0, w: 2.3, h: 2.0 });
    // Description text below
    if (d.description) {
      slide.addText(d.description, {
        x: 5.6, y: 4.2, w: 3.8, h: 1.2,
        fontSize: 14, fontFace: C.FONT, color: "CCCCDD",
        valign: "top", lineSpacingMultiple: 1.4, margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildEngage(slide, d) {
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, {
      sectionLabel: d.sectionLabel || "00 · APERTURA",
      squiggle: { x: 7.5, y: 0.3 }
    });
    placeAccent(slide, 0.85, C.YELLOW);
    slide.addText(d.title, {
      x: 0.75, y: 0.85, w: 7.3, h: 0.9,
      fontSize: 30, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0, lineSpacingMultiple: 1.1
    });
    warnIfTooDense("engage", d.body, 5);
    let iconImg = null;
    if (d.icon) {
      const data = await iconToDataUri(d.icon, C.PURPLE, 256);
      if (data) iconImg = data;
    }
    if (iconImg) {
      slide.addImage({ data: iconImg, x: 0.75, y: 1.95, w: 0.6, h: 0.6 });
    }
    const bodyPara = (d.body || []).map((t) => ({
      text: t.text || t,
      options: {
        breakLine: true,
        fontSize: 20,
        color: t.bold ? C.PURPLE : C.NAVY,
        bold: !!t.bold
      }
    }));
    slide.addText(bodyPara, {
      x: iconImg ? 1.6 : 0.75, y: 1.95, w: iconImg ? 6.6 : 7.5, h: 2.5,
      fontFace: C.FONT, valign: "top", lineSpacingMultiple: 1.5
    });
    // Pixel smiley peeking from bottom-right corner (brand warmth)
    await placeSmileyPeek(slide, "bottom-right");
    slide.addNotes(d.notes || "");
  }

  async function buildObjectives(slide, d) {
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, C.PURPLE);
    slide.addText(d.title || "Objetivos de aprendizaje", {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 28, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    // Poetic subtitle (per tone-and-narrative-guide.md)
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.35,
        fontSize: 14, fontFace: C.FONT, color: C.GRAY_MED, italic: true, margin: 0
      });
    }
    // Banner
    const bannerY = d.subtitle ? 1.95 : 1.55;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: bannerY, w: 3.4, h: 0.4, fill: { color: C.PURPLE }
    });
    slide.addText("AL FINAL DE HOY PODRÁS", {
      x: 0.6, y: bannerY, w: 3.4, h: 0.4,
      fontSize: 12, fontFace: C.FONT, color: C.WHITE, bold: true,
      letterSpacing: 2, align: "center", valign: "middle", margin: 0
    });
    // Numbered list with rose squares (brandbook pattern)
    // Bloom levels are PEDAGOGICAL METADATA — kept in data for the speaker notes,
    // but NEVER rendered to students. Per tone-and-narrative-guide.md.
    const objs = d.objectives || [];
    const listStartY = bannerY + 0.7;
    objs.forEach((o, i) => {
      const y = listStartY + i * 0.6;
      placePinkNumber(slide, i + 1, 0.7, y, 0.38);
      const verb = (o.verb || "").toUpperCase();
      const content = o.content || "";
      slide.addText([
        { text: verb + " ", options: { bold: true, color: C.PURPLE } },
        { text: content, options: { color: C.NAVY } },
      ], {
        x: 1.2, y: y - 0.03, w: 7.9, h: 0.5,
        fontSize: 16, fontFace: C.FONT, valign: "middle", margin: 0,
        lineSpacingMultiple: 1.3
      });
    });
    // Warm closing message (per tone-and-narrative-guide.md — non-negotiable)
    const warm = d.warmClosing || "Si al final puedes hacer estas cosas, la clase fue exitosa.";
    slide.addText(warm, {
      x: 0.75, y: 5.1, w: 8.5, h: 0.4,
      fontSize: 13, fontFace: C.FONT, color: C.PURPLE, italic: true, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    // Auto-augment speaker notes with Bloom levels (for the mentora, not the students)
    const bloomNotes = objs
      .filter(o => o.bloom != null)
      .map((o, i) => `Obj ${i + 1}: Bloom ${o.bloom}`)
      .join(" · ");
    const allNotes = [d.notes, bloomNotes].filter(Boolean).join("\n\n[Pedagogía interna — no mostrar]: " + bloomNotes ? "" : "") || "";
    slide.addNotes((d.notes || "") + (bloomNotes ? `\n\n[Pedagogía — no mostrar a las alumnas]: ${bloomNotes}` : ""));
  }

  async function buildSteps(slide, d) {
    slide.background = { color: C.WHITE };
    placeAccent(slide, 0.35, colorVal(d.accentColor) || C.PURPLE);
    slide.addText(d.title, {
      x: 0.75, y: 0.3, w: 8.5, h: 0.7,
      fontSize: 32, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    const items = d.steps || [];
    const startY = 1.3;
    const gap = items.length > 5 ? 0.6 : 0.75;
    items.forEach((s, i) => {
      const y = startY + i * gap;
      placePinkNumber(slide, i + 1, 0.7, y, 0.4);
      const keyword = s.keyword || s.label || "";
      const desc = s.desc || s.text || (typeof s === "string" ? s : "");
      slide.addText([
        ...(keyword ? [{ text: keyword + ". ", options: { bold: true, color: C.NAVY } }] : []),
        { text: desc, options: { color: C.NAVY } },
      ], {
        x: 1.25, y: y - 0.05, w: 7.8, h: gap,
        fontSize: 18, fontFace: C.FONT, valign: "middle", margin: 0,
        lineSpacingMultiple: 1.3,
      });
    });
    slide.addNotes(d.notes || "");
  }

  async function buildActivity(slide, d) {
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.BLUE);
    slide.addText(d.title, {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 26, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    const cardY = d.cardY || 1.75;
    const cardH = d.cardH || 3.6;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: cardY, w: 8.8, h: cardH,
      fill: { color: C.ROSE },
      shadow: { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.08 }
    });
    if (d.activityLabel) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: cardY, w: 2.8, h: 0.45, fill: { color: colorVal(d.accentColor) || C.BLUE }
      });
      slide.addText(d.activityLabel, {
        x: 0.6, y: cardY, w: 2.8, h: 0.45,
        fontSize: 14, fontFace: C.FONT, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    if (d.icon) {
      const data = await iconToDataUri(d.icon, colorVal(d.accentColor) || C.BLUE, 256);
      if (data) slide.addImage({ data, x: 8.3, y: cardY + 0.3, w: 0.5, h: 0.5 });
    }
    warnIfTooDense("activity", d.body, 5);
    const bodyPara = (d.body || []).map((t) => ({
      text: t.text || t,
      options: {
        breakLine: true,
        fontSize: t.bold ? 18 : (d.bodyFontSize || 18),
        color: t.bold ? C.PURPLE : C.NAVY, bold: !!t.bold
      }
    }));
    slide.addText(bodyPara, {
      x: 1.2, y: cardY + 0.7, w: 7.5, h: cardH - 0.9,
      fontFace: C.FONT, valign: "top", lineSpacingMultiple: 1.5
    });
    slide.addNotes(d.notes || "");
  }

  async function buildContent(slide, d) {
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.PURPLE);
    slide.addText(d.title, {
      x: 0.75, y: 0.85, w: 8.0, h: 0.6,
      fontSize: 28, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.4,
        fontSize: 16, fontFace: C.FONT, color: C.NAVY, italic: true, margin: 0
      });
    }
    warnIfTooDense("content", d.body, 5);
    const bodyPara = (d.body || []).map((t) => ({
      text: t.text || t,
      options: { breakLine: true, fontSize: 18, color: C.NAVY, bold: !!t.bold }
    }));
    slide.addText(bodyPara, {
      x: 0.75, y: d.subtitle ? 2.0 : 1.65, w: 8.5, h: 3.2,
      fontFace: C.FONT, valign: "top", lineSpacingMultiple: 1.5
    });
    slide.addNotes(d.notes || "");
  }

  async function buildDefinition(slide, d) {
    slide.background = { color: C.NAVY };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel, darkBg: true, stars: [
      { x: 9.45, y: 5.4, sz: 0.28, c: C.YELLOW },
      { x: 0.2, y: 5.4, sz: 0.22, c: C.ROSE }
    ]});
    placeLogo(slide);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 0.85, w: 0.06, h: 0.55, fill: { color: colorVal(d.accentColor) || C.PURPLE }
    });
    slide.addText(d.title, {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 28, fontFace: C.FONT, color: C.WHITE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.definition) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: 1.65, w: 8.8, h: 1.0, fill: { color: colorVal(d.accentColor) || C.PURPLE }
      });
      slide.addText(d.definition, {
        x: 0.8, y: 1.65, w: 8.4, h: 1.0,
        fontSize: 20, fontFace: C.FONT, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    (d.cards || []).forEach((c, i) => {
      const gap = d.cardGap || 3.1;
      const xPos = 0.6 + i * gap;
      const cardW = d.cardW || 2.8;
      const cardY = d.cardY || 3.0;
      const cardH = d.cardH || 2.2;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: xPos, y: cardY, w: cardW, h: cardH,
        fill: { color: "1E2240" },
        shadow: { type: "outer", blur: 3, offset: 1, angle: 135, color: "000000", opacity: 0.2 }
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: xPos, y: cardY, w: cardW, h: 0.06,
        fill: { color: colorVal(c.color) || C.GREEN }
      });
      slide.addText(c.label, {
        x: xPos + 0.15, y: cardY + 0.2, w: cardW - 0.3, h: 0.6,
        fontSize: 22, fontFace: C.FONT, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0
      });
      slide.addText(c.desc, {
        x: xPos + 0.15, y: cardY + 0.9, w: cardW - 0.3, h: 1.0,
        fontSize: 16, fontFace: C.FONT, color: "BBBBCE",
        align: "center", valign: "top", margin: 0
      });
    });
    slide.addNotes(d.notes || "");
  }

  async function buildPillars(slide, d) {
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.BLUE);
    slide.addText(d.title, {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 26, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.35,
        fontSize: 14, fontFace: C.FONT, color: C.NAVY, italic: true, margin: 0
      });
    }
    const startY = d.subtitle ? 1.95 : 1.55;
    (d.pillars || []).forEach((p, i) => {
      const yPos = startY + i * 0.85;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: yPos, w: 0.08, h: 0.85,
        fill: { color: colorVal(p.color) || C.PURPLE }
      });
      slide.addText(p.label, {
        x: 0.9, y: yPos, w: 2.5, h: 0.45,
        fontSize: 20, fontFace: C.FONT, color: C.NAVY, bold: true,
        valign: "middle", margin: 0
      });
      slide.addText(p.desc, {
        x: 0.9, y: yPos + 0.4, w: 8, h: 0.45,
        fontSize: 16, fontFace: C.FONT, color: C.GRAY_MED, valign: "top", margin: 0
      });
    });
    slide.addNotes(d.notes || "");
  }

  // Token-aware code segment normalization.
  // Default: segments within a line do NOT break — only the last segment of each line breaks.
  // The line-flattener (buildCodeLines below) forces the last segment to break.
  function normalizeCodeSegment(seg) {
    if (typeof seg === "string") {
      return { text: seg, options: { color: C.CODE_DEFAULT, fontFace: C.FONT_CODE, fontSize: 18 } };
    }
    const tokenColor = seg.token && CODE_TOKENS[seg.token];
    return {
      text: seg.text || "",
      options: {
        ...(seg.options || {}),
        breakLine: seg.breakLine === true,  // only break when explicitly true
        color: seg.color || tokenColor || (seg.options && seg.options.color) || C.CODE_DEFAULT,
        bold: !!seg.bold,
        fontSize: seg.fontSize || 18,
        fontFace: seg.fontFace || C.FONT_CODE,
      }
    };
  }

  // Convert a `code: [[...], [...]]` structure into a single flat array of segments.
  // Each line ends with breakLine:true on its last segment. Empty lines become a single space.
  // If fontSize is provided, apply to all segments (used by adaptive sizing in buildCode).
  function buildCodeLines(codeLines, forceFontSize) {
    const out = [];
    (codeLines || []).forEach(line => {
      const segments = typeof line === "string" ? [line] : (Array.isArray(line) ? line : [line]);
      if (segments.length === 0) {
        out.push({ text: " ", options: { breakLine: true, fontFace: C.FONT_CODE, fontSize: forceFontSize || 16, color: C.CODE_DEFAULT } });
        return;
      }
      const normalized = segments.map(normalizeCodeSegment);
      if (forceFontSize) {
        normalized.forEach(s => { s.options.fontSize = forceFontSize; });
      }
      // Force last segment of this line to break
      normalized[normalized.length - 1].options.breakLine = true;
      out.push(...normalized);
    });
    return out;
  }

  async function buildCode(slide, d) {
    // Layout limpio: title arriba, filename discreto, code area amplia.
    // SIN mock editor title bar (causaba overlap).
    slide.background = { color: C.NAVY };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel, darkBg: true, stars: [] });
    placeLogo(slide);
    const accentColor = colorVal(d.accentColor) || C.GREEN;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: 0.85, w: 0.06, h: 0.55, fill: { color: accentColor }
    });
    slide.addText(d.title, {
      x: 0.75, y: 0.82, w: 7.0, h: 0.55,
      fontSize: 22, fontFace: C.FONT, color: C.WHITE, bold: true,
      valign: "middle", margin: 0, lineSpacingMultiple: 1.0
    });
    // Filename as discreet pill (top-right) — replaces the noisy editor bar
    if (d.filename) {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 7.8, y: 0.92, w: 1.5, h: 0.35,
        fill: { color: "2A2D4A" }, line: { color: accentColor, width: 1 },
        rectRadius: 0.05
      });
      slide.addText(d.filename, {
        x: 7.8, y: 0.92, w: 1.5, h: 0.35,
        fontSize: 11, fontFace: C.FONT_CODE, color: accentColor,
        align: "center", valign: "middle", margin: 0
      });
    }
    if (d.description) {
      slide.addText(d.description, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.4,
        fontSize: 13, fontFace: C.FONT, color: "BBBBCE", italic: true,
        valign: "top", lineSpacingMultiple: 1.3, margin: 0
      });
    }
    const codeY = d.codeY || (d.description ? 1.95 : 1.5);
    const codeH = d.codeH || (5.4 - codeY);
    const codeW = d.codeW || 8.8;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: codeY, w: codeW, h: codeH,
      fill: { color: "0D0F1A" },
      line: { color: "1E2240", width: 1 }
    });
    // Adapt font size to number of lines so code fits inside the box
    const lineCount = (d.code || []).length;
    const autoFontSize = lineCount > 12 ? 12 : (lineCount > 9 ? 14 : 16);
    const codeContent = buildCodeLines(d.code, autoFontSize);
    slide.addText(codeContent, {
      x: 0.85, y: codeY + 0.15, w: codeW - 0.5, h: codeH - 0.3,
      fontSize: autoFontSize, fontFace: C.FONT_CODE, color: C.CODE_DEFAULT,
      valign: "top", margin: 0, lineSpacingMultiple: 1.25
    });
    slide.addNotes(d.notes || "");
  }

  async function buildManualTrace(slide, d) {
    // Programming UDL adaptation: trace variable state step-by-step before coding.
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.TEAL);
    slide.addText(d.title || "Trazado manual", {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 26, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    // Code on the left
    const codeX = 0.6, codeY = 1.2, codeW = 4.2, codeH = 3.8;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: codeX, y: codeY, w: codeW, h: codeH, fill: { color: C.NAVY }
    });
    slide.addText("CÓDIGO", {
      x: codeX + 0.1, y: codeY + 0.1, w: codeW - 0.2, h: 0.3,
      fontSize: 11, fontFace: C.FONT, color: C.YELLOW, bold: true,
      letterSpacing: 2, margin: 0
    });
    const codeContent = buildCodeLines(d.code);
    slide.addText(codeContent, {
      x: codeX + 0.15, y: codeY + 0.45, w: codeW - 0.3, h: codeH - 0.6,
      fontSize: 14, fontFace: C.FONT_CODE, color: C.CODE_DEFAULT,
      valign: "top", margin: 0, lineSpacingMultiple: 1.25
    });
    // Trace table on the right
    const tblX = 5.0, tblY = 1.2, tblW = 4.4, tblH = 3.8;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: tblX, y: tblY, w: tblW, h: tblH,
      fill: { color: C.ROSE }, line: { color: C.NAVY, width: 0 }
    });
    slide.addText("ESTADO DE VARIABLES", {
      x: tblX + 0.15, y: tblY + 0.1, w: tblW - 0.3, h: 0.3,
      fontSize: 11, fontFace: C.FONT, color: C.PURPLE, bold: true,
      letterSpacing: 2, margin: 0
    });
    const headers = d.headers || ["Paso", "Variable", "Valor"];
    const colWs = d.colW || [0.7, 1.4, 2.0];
    let colX = tblX + 0.15;
    headers.forEach((h, i) => {
      slide.addText(h, {
        x: colX, y: tblY + 0.45, w: colWs[i], h: 0.3,
        fontSize: 12, fontFace: C.FONT, color: C.NAVY, bold: true,
        valign: "middle", margin: 0
      });
      colX += colWs[i];
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: tblX + 0.15, y: tblY + 0.78, w: tblW - 0.3, h: 0.02, fill: { color: C.PURPLE }
    });
    (d.rows || []).forEach((row, i) => {
      const yRow = tblY + 0.9 + i * 0.35;
      let cellX = tblX + 0.15;
      row.forEach((cell, j) => {
        slide.addText(String(cell), {
          x: cellX, y: yRow, w: colWs[j], h: 0.3,
          fontSize: 13, fontFace: j === 2 ? C.FONT_CODE : C.FONT,
          color: C.NAVY, valign: "middle", margin: 0
        });
        cellX += colWs[j];
      });
    });
    if (d.footer) {
      slide.addText(d.footer, {
        x: 0.6, y: 5.1, w: 8.8, h: 0.3,
        fontSize: 13, fontFace: C.FONT, color: C.GRAY_MED,
        align: "center", valign: "middle"
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildReflection(slide, d) {
    // Exercise 10 — written reflection, Evaluate phase of 5E. Non-negotiable.
    slide.background = { color: C.PURPLE };
    placeLogo(slide);
    placeStar(slide, 0.6, 0.4, 0.35, C.YELLOW);
    placeStar(slide, 8.6, 0.4, 0.3, C.ROSE);
    slide.addText("EJERCICIO 10 · REFLEXIÓN", {
      x: 0.6, y: 0.5, w: 6, h: 0.4,
      fontSize: 13, fontFace: C.FONT, color: C.YELLOW, bold: true,
      letterSpacing: 3, margin: 0
    });
    slide.addText(d.title || "Reflexión final", {
      x: 0.6, y: 1.0, w: 8.5, h: 0.8,
      fontSize: 36, fontFace: C.FONT, color: C.WHITE, bold: true,
      valign: "middle", margin: 0
    });
    // Quote icon
    const quoteIcon = await iconToDataUri("quote", C.YELLOW, 128);
    if (quoteIcon) {
      slide.addImage({ data: quoteIcon, x: 0.6, y: 2.0, w: 0.5, h: 0.5 });
    }
    const prompts = d.prompts || [];
    prompts.forEach((p, i) => {
      const y = 2.2 + i * 0.85;
      slide.addText(p, {
        x: 1.4, y, w: 7.8, h: 0.8,
        fontSize: 22, fontFace: C.FONT, color: C.WHITE,
        italic: true, valign: "middle", margin: 0,
        lineSpacingMultiple: 1.2
      });
    });
    // Footer instruction
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 5.0, w: 8.8, h: 0.5, fill: { color: C.YELLOW }
    });
    slide.addText(`Escribe en tu cuaderno — ${d.duration || "8 min"}`, {
      x: 0.6, y: 5.0, w: 8.8, h: 0.5,
      fontSize: 14, fontFace: C.FONT, color: C.NAVY, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    slide.addNotes(d.notes || "");
  }

  async function buildRubric(slide, d) {
    slide.background = { color: C.WHITE };
    placeAccent(slide, 0.35, colorVal(d.accentColor) || C.GREEN);
    slide.addText(d.title || "Rúbrica formativa", {
      x: 0.75, y: 0.3, w: 8.5, h: 0.65,
      fontSize: 28, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    const defaultLevels = [
      { range: "1–2", desc: "No accede a la tarea básica — vacío fundacional", color: "RED" },
      { range: "3–4", desc: "Completa calentamiento, vacíos en práctica central", color: "YELLOW" },
      { range: "5–6", desc: "Completa ejercicios centrales — nivel Aplicar de Bloom", color: "BLUE" },
      { range: "7–8", desc: "Completa la mayoría incluyendo desafíos — nivel Analizar", color: "TEAL" },
      { range: "9–10", desc: "Todos los ejercicios, casos borde, justifica decisiones — Evaluar", color: "GREEN" },
    ];
    const levels = d.levels || defaultLevels;
    levels.forEach((lv, i) => {
      const y = 1.2 + i * 0.75;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y, w: 0.08, h: 0.6, fill: { color: colorVal(lv.color) || C.PURPLE }
      });
      slide.addText(lv.range, {
        x: 0.85, y, w: 0.9, h: 0.6,
        fontSize: 22, fontFace: C.FONT, color: C.PURPLE, bold: true,
        valign: "middle", margin: 0
      });
      slide.addText(lv.desc, {
        x: 1.85, y, w: 7.4, h: 0.6,
        fontSize: 15, fontFace: C.FONT, color: C.NAVY,
        valign: "middle", margin: 0, lineSpacingMultiple: 1.2
      });
    });
    if (d.footer) {
      slide.addText(d.footer, {
        x: 0.6, y: 5.1, w: 8.8, h: 0.3,
        fontSize: 13, fontFace: C.FONT, color: C.GRAY_MED,
        italic: true, align: "center", valign: "middle"
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildPlanB(slide, d) {
    // Plan B activity for inspirational-talk classes (Classes 3, 5, 9).
    // Rendered as a Windows-95 modal alert — brandbook signature.
    slide.background = { color: C.YELLOW };
    placeStar(slide, 0.5, 0.5, 0.35, C.ROSE);
    placeStar(slide, 8.8, 0.5, 0.3, C.PURPLE);
    placeStar(slide, 0.5, 4.8, 0.3, C.NAVY);
    placeStar(slide, 8.8, 4.8, 0.35, C.NAVY);
    // Modal window
    placeAlertWindow(slide, 1.2, 1.0, 7.6, 4.0, {
      titleBarColor: C.RED, bodyBg: C.WHITE,
      titleText: d.titleBar || "ALERTA · PLAN B"
    });
    slide.addText(d.title || "Si la invitada no puede asistir", {
      x: 1.5, y: 1.6, w: 7.0, h: 0.5,
      fontSize: 24, fontFace: C.FONT, color: C.NAVY, bold: true,
      valign: "middle", margin: 0
    });
    slide.addText(d.subtitle || "Activamos la actividad de respaldo:", {
      x: 1.5, y: 2.1, w: 7.0, h: 0.4,
      fontSize: 16, fontFace: C.FONT, color: C.GRAY_MED,
      valign: "middle", margin: 0
    });
    const steps = d.steps || [];
    steps.forEach((s, i) => {
      const y = 2.7 + i * 0.5;
      placePinkNumber(slide, i + 1, 1.5, y, 0.32);
      slide.addText(typeof s === "string" ? s : (s.text || s.desc || ""), {
        x: 1.95, y, w: 6.6, h: 0.4,
        fontSize: 15, fontFace: C.FONT, color: C.NAVY,
        valign: "middle", margin: 0
      });
    });
    // OK button (Windows-95 style)
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 7.4, y: 4.5, w: 1.0, h: 0.4,
      fill: { color: C.ROSE }, line: { color: C.NAVY, width: 2 },
      rectRadius: 0.05
    });
    slide.addText("¡ OK !", {
      x: 7.4, y: 4.5, w: 1.0, h: 0.4,
      fontSize: 14, fontFace: C.FONT, color: C.NAVY, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    slide.addNotes(d.notes || "");
  }

  async function buildFlowchart(slide, d) {
    slide.background = { color: C.WHITE };
    placeAccent(slide, 0.35, colorVal(d.accentColor) || C.PURPLE);
    slide.addText(d.title, {
      x: 0.75, y: 0.3, w: 8.5, h: 0.65,
      fontSize: 32, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    (d.symbols || []).forEach((s, i) => {
      const gap = d.symbolGap || 2.4;
      const xPos = 0.5 + i * gap;
      const shapeX = xPos + 0.55;
      const shapeY = 1.2;
      const shapeW = 0.9;
      const shapeH = 0.6;
      const col = colorVal(s.color) || C.PURPLE;
      if (i === 0) slide.addShape(pres.shapes.OVAL, { x: shapeX, y: shapeY, w: shapeW, h: shapeH, fill: { color: col } });
      else if (i === 1) slide.addShape(pres.shapes.RECTANGLE, { x: shapeX, y: shapeY, w: shapeW, h: shapeH, fill: { color: col } });
      else if (i === 2) slide.addShape(pres.shapes.DIAMOND, { x: shapeX + 0.05, y: shapeY - 0.05, w: shapeW - 0.1, h: shapeH + 0.1, fill: { color: col } });
      else slide.addShape(pres.shapes.PARALLELOGRAM, { x: shapeX, y: shapeY, w: shapeW, h: shapeH, fill: { color: col } });
      slide.addText(s.shape, {
        x: xPos, y: 1.95, w: 2.1, h: 0.35,
        fontSize: 16, fontFace: C.FONT, color: C.NAVY, bold: true,
        align: "center", valign: "middle", margin: 0
      });
      slide.addText(s.meaning, {
        x: xPos, y: 2.35, w: 2.1, h: 0.5,
        fontSize: 14, fontFace: C.FONT, color: C.GRAY_MED,
        align: "center", valign: "top", margin: 0
      });
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.6, y: 3.3, w: 8.8, h: 0.5,
        fontSize: 16, fontFace: C.FONT, color: C.NAVY, align: "center", valign: "middle"
      });
    }
    if (d.example) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: 4.0, w: 8.8, h: 1.2, fill: { color: C.ROSE }
      });
      slide.addText(d.example, {
        x: 0.8, y: 4.0, w: 8.4, h: 1.2,
        fontSize: 16, fontFace: C.FONT, color: C.NAVY, align: "center", valign: "middle"
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildErrorCode(slide, d) {
    slide.background = { color: C.WHITE };
    placeAccent(slide, 0.35, colorVal(d.accentColor) || C.YELLOW);
    slide.addText(d.title, {
      x: 0.75, y: 0.3, w: 8.5, h: 0.65,
      fontSize: 32, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.icon) {
      const data = await iconToDataUri(d.icon, C.RED, 256);
      if (data) slide.addImage({ data, x: 0.75, y: 1.15, w: 0.5, h: 0.5 });
    }
    if (d.question) {
      slide.addText(d.question, {
        x: d.icon ? 1.4 : 0.75, y: 1.1, w: 7.5, h: 0.6,
        fontSize: 18, fontFace: C.FONT, color: C.NAVY, bold: true,
        valign: "middle", margin: 0
      });
    }
    const codeY = d.codeY || 1.9;
    const codeH = d.codeH || 2.3;
    const codeW = d.codeW || 8.8;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: codeY, w: codeW, h: codeH,
      fill: { color: C.NAVY },
      shadow: { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.15 }
    });
    const codeContent = buildCodeLines(d.code);
    slide.addText(codeContent, {
      x: 0.9, y: codeY + 0.1, w: codeW - 0.6, h: codeH - 0.2,
      fontSize: 16, fontFace: C.FONT_CODE, color: C.CODE_DEFAULT,
      valign: "top", margin: 0, lineSpacingMultiple: 1.2
    });
    if (d.errorMessage) {
      const errY = d.errY || codeY + codeH + 0.2;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: errY, w: 8.8, h: 0.5, fill: { color: C.RED }
      });
      slide.addText(d.errorMessage, {
        x: 0.8, y: errY, w: 8.4, h: 0.5,
        fontSize: 15, fontFace: C.FONT, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    // Mentor protocol callout (per facilitation-guide.md: "leer el error CON la alumna, no para ella")
    if (d.protocol) {
      const pY = d.errorMessage ? (d.errY || codeY + codeH + 0.2) + 0.6 : codeY + codeH + 0.2;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: pY, w: 8.8, h: 0.45, fill: { color: C.YELLOW }
      });
      slide.addText([
        { text: "Mentora: ", options: { bold: true, color: C.NAVY } },
        { text: d.protocol, options: { color: C.NAVY, italic: true } },
      ], {
        x: 0.8, y: pY, w: 8.4, h: 0.45,
        fontSize: 13, fontFace: C.FONT, valign: "middle", margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildReference(slide, d) {
    slide.background = { color: C.WHITE };
    placeAccent(slide, 0.35, colorVal(d.accentColor) || C.BLUE);
    slide.addText(d.title, {
      x: 0.75, y: 0.3, w: 8.5, h: 0.65,
      fontSize: 28, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    const cardX = 0.6, cardY = 1.2, cardW = 8.8, cardH = d.cardH || 3.8;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cardX, y: cardY, w: cardW, h: cardH,
      fill: { color: C.ROSE },
      shadow: { type: "outer", blur: 3, offset: 1, angle: 135, color: "000000", opacity: 0.08 }
    });
    if (d.headers) {
      const colX = d.colX || [1.0, 2.5, 5.0];
      const colW = d.colW || [1.5, 2.5, 4.0];
      d.headers.forEach((h, i) => {
        slide.addText(h, {
          x: colX[i], y: cardY + 0.1, w: colW[i], h: 0.4,
          fontSize: 14, fontFace: C.FONT, color: C.PURPLE, bold: true,
          valign: "middle", margin: 0
        });
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: colX[0], y: cardY + 0.55, w: colX[0] + colW[0] + colW[1] + colW[2] - colX[0], h: 0.02,
        fill: { color: C.PURPLE }
      });
    }
    (d.rows || []).forEach((row, i) => {
      const colX = d.colX || [1.0, 2.5, 5.0];
      const colW = d.colW || [1.5, 2.5, 4.0];
      const yPos = cardY + 0.8 + i * 0.55;
      if (d.headers) {
        row.forEach((cell, j) => {
          slide.addText(cell, {
            x: colX[j], y: yPos, w: colW[j], h: 0.45,
            fontSize: j === 0 ? 24 : (j === 1 ? 16 : 15),
            fontFace: C.FONT, color: j === 0 ? C.PURPLE : C.NAVY,
            bold: j <= 1, align: j === 0 ? "center" : "left",
            valign: "middle", margin: 0
          });
        });
      } else {
        slide.addText(row.sym, {
          x: colX[0], y: yPos, w: colW[0], h: 0.45,
          fontSize: 24, fontFace: C.FONT, color: C.PURPLE, bold: true,
          align: "center", valign: "middle", margin: 0
        });
        slide.addText(row.name, {
          x: colX[1], y: yPos, w: colW[1], h: 0.45,
          fontSize: 16, fontFace: C.FONT, color: C.NAVY, bold: true,
          valign: "middle", margin: 0
        });
        slide.addText(row.desc, {
          x: colX[2], y: yPos, w: colW[2], h: 0.45,
          fontSize: 15, fontFace: C.FONT, color: C.NAVY,
          valign: "middle", margin: 0
        });
      }
    });
    if (d.footer) {
      slide.addText(d.footer, {
        x: cardX, y: cardY + cardH + 0.15, w: cardW, h: 0.3,
        fontSize: 13, fontFace: C.FONT, color: C.GRAY_MED, align: "center", valign: "middle"
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildKahoot(slide, d) {
    slide.background = { color: C.YELLOW };
    placeLogo(slide);
    placeStar(slide, 0.5, 0.8, 0.4, C.ROSE);
    placeStar(slide, 9.0, 0.8, 0.4, C.PURPLE);
    placeStar(slide, 0.5, 4.2, 0.35, C.TEAL);
    placeStar(slide, 9.0, 4.2, 0.35, C.ROSE);
    slide.addText("¡Kahoot!", {
      x: 0, y: 1.2, w: 10, h: 1.2,
      fontSize: 56, fontFace: C.FONT, color: C.NAVY, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    slide.addText("Preparen sus dispositivos\nEntren a kahoot.it", {
      x: 0, y: 2.8, w: 10, h: 1.5,
      fontSize: 24, fontFace: C.FONT, color: C.NAVY,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.3
    });
    slide.addNotes(d.notes || "");
  }

  async function buildSteam(slide, d) {
    // Re-diseño: layout limpio con respiro entre elementos. Aligned with
    // references/steam-woman-guide.md. Source goes to speaker notes, not on slide.
    const bgColor = colorVal(d.bgColor) || C.PURPLE;
    slide.background = { color: bgColor };
    placeStar(slide, 0.4, 0.4, 0.35, C.YELLOW);
    // Section label upper-left (consistent with brand furniture)
    slide.addText("MUJER STEAM", {
      x: 0.85, y: 0.45, w: 4, h: 0.35,
      fontSize: 12, fontFace: C.FONT, color: C.YELLOW, bold: true,
      letterSpacing: 3, valign: "middle", margin: 0
    });
    // Name — large, single line if possible
    slide.addText(d.name, {
      x: 0.6, y: 0.95, w: 7.5, h: 0.7,
      fontSize: 36, fontFace: C.FONT, color: C.WHITE, bold: true,
      valign: "middle", margin: 0
    });
    // Meta: years · area · country (small, single line)
    const metaPieces = [d.years, d.area, d.country].filter(Boolean);
    if (metaPieces.length) {
      slide.addText(metaPieces.join("  ·  "), {
        x: 0.6, y: 1.65, w: 8.5, h: 0.3,
        fontSize: 13, fontFace: C.FONT, color: "CCCCDD", margin: 0
      });
    }
    // Photo (smaller, right side) — leaves room for content
    const photoX = 6.6, photoY = 2.1, photoW = 2.8, photoH = 2.4;
    let hasPhoto = false;
    if (d.photoPath) {
      try {
        if (fs.existsSync(d.photoPath)) {
          slide.addImage({
            path: d.photoPath, x: photoX, y: photoY, w: photoW, h: photoH,
            sizing: { type: "cover", w: photoW, h: photoH }
          });
          placeStar(slide, photoX + photoW - 0.3, photoY - 0.2, 0.4, C.YELLOW);
          hasPhoto = true;
        }
      } catch (_) { /* photo optional */ }
    }
    // Connection card — the most important pedagogical field
    if (d.connectionToClass) {
      const cardW = hasPhoto ? 5.6 : 8.8;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: 2.1, w: cardW, h: 1.4,
        fill: { color: "1E1640" },
        line: { color: C.YELLOW, width: 0 }
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: 2.1, w: 0.08, h: 1.4, fill: { color: C.YELLOW }
      });
      slide.addText("SU CONEXIÓN CON HOY", {
        x: 0.85, y: 2.2, w: cardW - 0.4, h: 0.25,
        fontSize: 10, fontFace: C.FONT, color: C.YELLOW, bold: true,
        letterSpacing: 2, margin: 0
      });
      slide.addText(d.connectionToClass, {
        x: 0.85, y: 2.5, w: cardW - 0.35, h: 0.95,
        fontSize: 12, fontFace: C.FONT, color: C.WHITE,
        valign: "top", margin: 0, lineSpacingMultiple: 1.35
      });
    }
    // Facts — max 2 lines to avoid clutter (more lines go to speaker notes)
    const factsStartY = 3.7;
    const factWidth = hasPhoto ? 5.6 : 8.8;
    (d.facts || []).slice(0, 2).forEach((line, i) => {
      const yPos = factsStartY + i * 0.35;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: yPos + 0.05, w: 0.06, h: 0.25, fill: { color: colorVal(d.factAccent) || C.TEAL }
      });
      slide.addText(line, {
        x: 0.85, y: yPos, w: factWidth - 0.3, h: 0.35,
        fontSize: 12, fontFace: C.FONT, color: "EEEEF5",
        valign: "middle", margin: 0
      });
    });
    // Quote band (bottom) — clean separator above
    if (d.quote) {
      const quoteY = 4.6;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: quoteY, w: 8.8, h: 0.03, fill: { color: C.YELLOW }
      });
      slide.addText("\"" + d.quote + "\"", {
        x: 0.6, y: quoteY + 0.12, w: 8.8, h: 0.4,
        fontSize: 13, fontFace: C.FONT, color: C.YELLOW, italic: true,
        align: "center", valign: "middle", margin: 0
      });
      if (d.quoteAttribution) {
        slide.addText("— " + d.quoteAttribution, {
          x: 0.6, y: quoteY + 0.55, w: 8.8, h: 0.22,
          fontSize: 9, fontFace: C.FONT, color: "AAAACC",
          align: "center", valign: "middle", margin: 0
        });
      }
    }
    // Reflection question at the very bottom (single line, very subtle)
    if (d.reflectionQuestion) {
      const cleanQ = d.reflectionQuestion.replace(/^¿|\?$/g, "");
      slide.addText("¿" + cleanQ + "?", {
        x: 0.6, y: 5.25, w: 8.8, h: 0.3,
        fontSize: 11, fontFace: C.FONT, color: C.ROSE, italic: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    // Source goes to speaker notes (not on slide — too noisy)
    const sourceNote = d.source ? `\n\nFuente verificable: ${d.source}` : "";
    const extraFactsNote = (d.facts || []).slice(2).length
      ? `\n\nDatos adicionales (no en slide):\n- ${(d.facts || []).slice(2).join("\n- ")}`
      : "";
    slide.addNotes((d.notes || "") + sourceNote + extraFactsNote);
  }

  async function buildClosing(slide, d) {
    slide.background = { color: C.NAVY };
    placeLogo(slide);
    placeStar(slide, 8.8, 0.4, 0.4, C.ROSE);
    if (d.label) {
      slide.addText(d.label, {
        x: 0.6, y: 0.3, w: 6, h: 0.6,
        fontSize: 14, fontFace: C.FONT, color: C.YELLOW, bold: true,
        letterSpacing: 3, margin: 0
      });
    }
    (d.sections || []).forEach((sec, i) => {
      const gap = d.sectionGap || 3.2;
      const xPos = 0.5 + i * gap;
      const secW = d.sectionW || 2.9;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: xPos, y: 1.3, w: secW, h: 0.5, fill: { color: colorVal(sec.color) || C.PURPLE }
      });
      slide.addText(sec.title, {
        x: xPos, y: 1.3, w: secW, h: 0.5,
        fontSize: 16, fontFace: C.FONT, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0
      });
      sec.items.forEach((item, j) => {
        const yPos = 2.0 + j * 0.6;
        slide.addShape(pres.shapes.OVAL, {
          x: xPos + 0.15, y: yPos + 0.13, w: 0.08, h: 0.08,
          fill: { color: colorVal(sec.color) || C.PURPLE }
        });
        slide.addText(item, {
          x: xPos + 0.35, y: yPos, w: secW - 0.35, h: 0.4,
          fontSize: 14, fontFace: C.FONT, color: "CCCCDD",
          valign: "middle", margin: 0
        });
      });
    });
    placeStar(slide, 0.6, 4.6, 0.3, C.TEAL);
    placeStar(slide, 9.0, 4.8, 0.25, C.YELLOW);
    slide.addNotes(d.notes || "");
  }

  async function buildGlossary(slide, d) {
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.TEAL);
    slide.addText(d.title || "Vocabulario clave", {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 26, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.35,
        fontSize: 13, fontFace: C.FONT, color: C.GRAY_MED, margin: 0
      });
    }
    const terms = (d.terms || []).slice(0, 4);
    const startY = d.subtitle ? 1.95 : 1.55;
    const cardH = terms.length > 2 ? 1.6 : 1.9;
    terms.forEach((t, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.6 + col * 4.5;
      const y = startY + row * (cardH + 0.2);
      // Card background
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w: 4.2, h: cardH,
        fill: { color: C.ROSE },
        line: { color: C.ROSE, width: 0 }
      });
      // Pink number badge
      placePinkNumber(slide, i + 1, x + 0.2, y + 0.2, 0.35);
      // Term in monospace if code-like (wrapped in backticks)
      const isCode = typeof t.term === "string" && /^`.+`$/.test(t.term);
      const termText = isCode ? t.term.replace(/^`|`$/g, "") : t.term;
      slide.addText(termText, {
        x: x + 0.7, y: y + 0.15, w: 3.4, h: 0.4,
        fontSize: 18, fontFace: isCode ? C.FONT_CODE : C.FONT,
        color: C.PURPLE, bold: true, valign: "middle", margin: 0
      });
      // Simple definition
      if (t.definition) {
        slide.addText(t.definition, {
          x: x + 0.2, y: y + 0.65, w: 3.9, h: 0.55,
          fontSize: 12, fontFace: C.FONT, color: C.NAVY,
          valign: "top", margin: 0, lineSpacingMultiple: 1.2
        });
      }
      // Mini example
      if (t.example) {
        const isExCode = typeof t.example === "string" && /[{};=<>()]/.test(t.example);
        slide.addText("Ej: " + t.example, {
          x: x + 0.2, y: y + cardH - 0.5, w: 3.9, h: 0.4,
          fontSize: 11, fontFace: isExCode ? C.FONT_CODE : C.FONT,
          color: C.GRAY_MED, italic: !isExCode, valign: "top", margin: 0
        });
      }
    });
    if (d.footer) {
      slide.addText(d.footer, {
        x: 0.6, y: 5.1, w: 8.8, h: 0.3,
        fontSize: 12, fontFace: C.FONT, color: C.GRAY_MED,
        italic: true, align: "center", valign: "middle"
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildWelcome(slide, d) {
    // Welcome / opening ceremony slide. Per facilitation-guide.md and time-distribution-guide.md.
    // Uses retro frame + pink-square numbered agenda — brandbook signature.
    slide.background = { color: C.PURPLE };
    placeLogo(slide);
    placeStar(slide, 0.5, 0.4, 0.4, C.YELLOW);
    placeStar(slide, 8.9, 4.8, 0.35, C.ROSE);
    // Big greeting
    slide.addText(d.greeting || "¡BIENVENIDAS!", {
      x: 0.6, y: 0.7, w: 9, h: 1.0,
      fontSize: 48, fontFace: C.FONT, color: C.WHITE, bold: true,
      letterSpacing: 2, margin: 0
    });
    // Course name
    slide.addText(d.courseName || "", {
      x: 0.6, y: 1.7, w: 9, h: 0.4,
      fontSize: 18, fontFace: C.FONT, color: C.YELLOW, bold: true,
      letterSpacing: 3, margin: 0
    });
    // Yellow squiggle accent under course name
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: 2.15, w: 2.5, h: 0.05, fill: { color: C.YELLOW }
    });
    // Pixel-smiley + cursor on the right (brandbook signature)
    const smileyData = await iconToDataUri("pixel-smiley", null, 256);
    if (smileyData) slide.addImage({ data: smileyData, x: 7.2, y: 0.6, w: 1.4, h: 1.4 });
    const cursorData = await iconToDataUri("pixel-cursor", null, 128);
    if (cursorData) slide.addImage({ data: cursorData, x: 8.5, y: 1.8, w: 0.4, h: 0.4 });
    // Agenda — numbered items inside a light card
    const agenda = d.agenda || [];
    if (agenda.length) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: 2.6, w: 8.8, h: 0.5 + agenda.length * 0.5,
        fill: { color: C.WHITE },
        line: { color: C.WHITE, width: 0 }
      });
      slide.addText("HOY VAMOS A", {
        x: 0.85, y: 2.7, w: 4, h: 0.3,
        fontSize: 11, fontFace: C.FONT, color: C.PURPLE, bold: true,
        letterSpacing: 2, margin: 0
      });
      agenda.forEach((item, i) => {
        const y = 3.1 + i * 0.5;
        placePinkNumber(slide, i + 1, 0.85, y, 0.32);
        const label = typeof item === "string" ? item : (item.text || item.label || "");
        const duration = typeof item === "object" ? item.duration : null;
        slide.addText([
          { text: label, options: { color: C.NAVY, bold: true } },
          ...(duration ? [{ text: `   (${duration})`, options: { color: C.GRAY_MED, italic: true } }] : []),
        ], {
          x: 1.3, y, w: 7.9, h: 0.4,
          fontSize: 14, fontFace: C.FONT, valign: "middle", margin: 0
        });
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildClosingCeremony(slide, d) {
    // Closing ceremony / Demo Day slide. Per time-distribution-guide.md Demo Day structure.
    slide.background = { color: C.NAVY };
    placeLogo(slide);
    placeStar(slide, 0.5, 0.5, 0.5, C.YELLOW);
    placeStar(slide, 8.9, 0.5, 0.4, C.ROSE);
    placeStar(slide, 0.5, 5.0, 0.35, C.TEAL);
    placeStar(slide, 8.9, 5.0, 0.45, C.PURPLE);
    // Big closing label
    slide.addText("CEREMONIA DE CIERRE", {
      x: 0.6, y: 0.5, w: 6, h: 0.4,
      fontSize: 13, fontFace: C.FONT, color: C.YELLOW, bold: true,
      letterSpacing: 3, margin: 0
    });
    slide.addText(d.title || "¡Lo lograron!", {
      x: 0.6, y: 1.0, w: 9, h: 1.0,
      fontSize: 44, fontFace: C.FONT, color: C.WHITE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.6, y: 2.0, w: 9, h: 0.4,
        fontSize: 18, fontFace: C.FONT, color: "CCCCDD", margin: 0
      });
    }
    // Three milestones — what they can do now they couldn't before
    const milestones = d.milestones || [];
    milestones.slice(0, 3).forEach((m, i) => {
      const gap = 3.05;
      const x = 0.55 + i * gap;
      const color = ["PURPLE", "BLUE", "GREEN"][i] || "PURPLE";
      // Card
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.8, w: 2.9, h: 1.8,
        fill: { color: "1E2240" },
        line: { color: colorVal(color), width: 2 }
      });
      // Top color bar
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 2.8, w: 2.9, h: 0.08, fill: { color: colorVal(color) }
      });
      slide.addText(m.label || m.title || "", {
        x: x + 0.15, y: 2.95, w: 2.6, h: 0.4,
        fontSize: 14, fontFace: C.FONT, color: colorVal(color), bold: true,
        letterSpacing: 2, valign: "middle", margin: 0
      });
      slide.addText(m.text || m.desc || "", {
        x: x + 0.15, y: 3.4, w: 2.6, h: 1.15,
        fontSize: 14, fontFace: C.FONT, color: C.WHITE,
        valign: "top", margin: 0, lineSpacingMultiple: 1.3
      });
    });
    // Bottom message
    if (d.message) {
      slide.addText(d.message, {
        x: 0.6, y: 4.85, w: 8.8, h: 0.4,
        fontSize: 14, fontFace: C.FONT, color: C.YELLOW, italic: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildBlockTransition(slide, d) {
    // Transition between class blocks (Apertura → Cátedra → Kahoot → Práctica → Cierre).
    // Per time-distribution-guide.md. Brandbook: navy bg with bold block label and pink-square countdown.
    slide.background = { color: C.NAVY };
    placeLogo(slide);
    // 5E phase chip top-left
    if (d.phase) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: 0.4, w: 1.6, h: 0.35,
        fill: { color: colorVal(d.phaseColor) || C.YELLOW }
      });
      slide.addText(`5E · ${d.phase.toUpperCase()}`, {
        x: 0.6, y: 0.4, w: 1.6, h: 0.35,
        fontSize: 10, fontFace: C.FONT, color: C.NAVY, bold: true,
        letterSpacing: 2, align: "center", valign: "middle", margin: 0
      });
    }
    // Big block name
    slide.addText(d.block || "BLOQUE", {
      x: 0.6, y: 1.2, w: 9, h: 1.2,
      fontSize: 56, fontFace: C.FONT, color: C.WHITE, bold: true,
      letterSpacing: 1, valign: "middle", margin: 0
    });
    // Duration pill
    if (d.duration) {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.6, y: 2.55, w: 1.8, h: 0.5,
        fill: { color: C.YELLOW }, line: { color: C.YELLOW, width: 0 },
        rectRadius: 0.25
      });
      slide.addText(d.duration, {
        x: 0.6, y: 2.55, w: 1.8, h: 0.5,
        fontSize: 18, fontFace: C.FONT, color: C.NAVY, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    // What we'll do during this block
    const items = d.items || [];
    items.slice(0, 4).forEach((it, i) => {
      const y = 3.5 + i * 0.5;
      placePinkNumber(slide, i + 1, 0.6, y, 0.35);
      slide.addText(typeof it === "string" ? it : (it.text || it.label || ""), {
        x: 1.1, y, w: 7.8, h: 0.4,
        fontSize: 16, fontFace: C.FONT, color: "EEEEF5",
        valign: "middle", margin: 0
      });
    });
    // Optional "ready check" footer
    if (d.readyCheck) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.6, y: 5.3, w: 8.8, h: 0.4, fill: { color: C.YELLOW }
      });
      slide.addText([
        { text: "✓ ", options: { bold: true, color: C.NAVY } },
        { text: d.readyCheck, options: { color: C.NAVY } },
      ], {
        x: 0.8, y: 5.3, w: 8.4, h: 0.4,
        fontSize: 12, fontFace: C.FONT, valign: "middle", margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildBridge(slide, d) {
    // Narrative connector between classes. Two columns: "LO QUE YA SABES" + "LO QUE VIENE HOY".
    // Per tone-and-narrative-guide.md — non-negotiable except for class 01.
    // Direction "backward" (slide 02) or "forward" (penultimate slide).
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, C.PURPLE);
    slide.addText(d.title || "Lo que ya sabes, lo que viene", {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 26, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.35,
        fontSize: 14, fontFace: C.FONT, color: C.GRAY_MED, italic: true, margin: 0
      });
    }
    // Left card: previous knowledge
    const leftLabel = d.leftLabel || (d.direction === "forward" ? "ESTA CLASE TE DIO" : "LO QUE YA SABES");
    const rightLabel = d.rightLabel || (d.direction === "forward" ? "LO QUE VIENE" : "LO QUE VIENE HOY");
    const leftItems = d.left || d.previousKnowledge || [];
    const rightItems = d.right || d.todayContent || [];
    const cardY = 2.0;
    const cardH = 3.05;
    // Left card (rose/purple)
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: cardY, w: 4.3, h: cardH,
      fill: { color: C.LAVENDER_LIGHT },
      line: { color: C.LAVENDER_LIGHT, width: 0 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: cardY, w: 4.3, h: 0.45, fill: { color: C.PURPLE }
    });
    slide.addText(leftLabel, {
      x: 0.6, y: cardY, w: 4.3, h: 0.45,
      fontSize: 12, fontFace: C.FONT, color: C.WHITE, bold: true,
      letterSpacing: 2, align: "center", valign: "middle", margin: 0
    });
    if (d.leftSubtitle) {
      slide.addText(d.leftSubtitle, {
        x: 0.8, y: cardY + 0.6, w: 3.9, h: 0.35,
        fontSize: 13, fontFace: C.FONT, color: C.PURPLE, italic: true, margin: 0
      });
    }
    leftItems.forEach((it, i) => {
      const text = typeof it === "string" ? it : (it.text || it.label || "");
      const y = cardY + (d.leftSubtitle ? 1.0 : 0.65) + i * 0.45;
      slide.addShape(pres.shapes.OVAL, {
        x: 0.85, y: y + 0.15, w: 0.1, h: 0.1, fill: { color: C.PURPLE }
      });
      slide.addText(text, {
        x: 1.05, y, w: 3.7, h: 0.4,
        fontSize: 14, fontFace: C.FONT, color: C.NAVY,
        valign: "middle", margin: 0
      });
    });
    // Right card (yellow accent)
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.1, y: cardY, w: 4.3, h: cardH,
      fill: { color: C.LAVENDER_LIGHT },
      line: { color: C.LAVENDER_LIGHT, width: 0 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.1, y: cardY, w: 4.3, h: 0.45, fill: { color: C.YELLOW }
    });
    slide.addText(rightLabel, {
      x: 5.1, y: cardY, w: 4.3, h: 0.45,
      fontSize: 12, fontFace: C.FONT, color: C.NAVY, bold: true,
      letterSpacing: 2, align: "center", valign: "middle", margin: 0
    });
    if (d.rightSubtitle) {
      slide.addText(d.rightSubtitle, {
        x: 5.3, y: cardY + 0.6, w: 3.9, h: 0.35,
        fontSize: 13, fontFace: C.FONT, color: C.NAVY, italic: true, margin: 0
      });
    }
    rightItems.forEach((it, i) => {
      const text = typeof it === "string" ? it : (it.text || it.label || "");
      const y = cardY + (d.rightSubtitle ? 1.0 : 0.65) + i * 0.45;
      slide.addShape(pres.shapes.OVAL, {
        x: 5.35, y: y + 0.15, w: 0.1, h: 0.1, fill: { color: C.YELLOW }
      });
      slide.addText(text, {
        x: 5.55, y, w: 3.7, h: 0.4,
        fontSize: 14, fontFace: C.FONT, color: C.NAVY,
        valign: "middle", margin: 0
      });
    });
    // Optional narrative line at the bottom (kept inside slide bounds)
    if (d.narrative) {
      slide.addText(d.narrative, {
        x: 0.75, y: 5.2, w: 8.5, h: 0.35,
        fontSize: 12, fontFace: C.FONT, color: C.PURPLE, italic: true, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildPullQuote(slide, d) {
    // A single poetic phrase on a colored background. Use to set up a concept or break rhythm.
    // Per Clase-09-Funciones.pptx slide 04.
    const bgColor = colorVal(d.bgColor) || C.LAVENDER_LIGHT;
    slide.background = { color: bgColor };
    placeStar(slide, 0.6, 0.5, 0.4, C.YELLOW);
    placeStar(slide, 8.8, 4.8, 0.35, C.PURPLE);
    placeStar(slide, 0.8, 4.7, 0.3, C.ROSE);
    placeStar(slide, 8.6, 0.6, 0.3, C.ROSE);
    // The quote itself — big, centered
    const quote = d.quote || d.text || "";
    slide.addText(quote, {
      x: 1.0, y: 1.5, w: 8.0, h: 2.5,
      fontSize: d.fontSize || 36, fontFace: C.FONT, color: C.PURPLE, bold: true,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2
    });
    // Optional subtitle (smaller line below the quote)
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 1.0, y: 4.0, w: 8.0, h: 0.6,
        fontSize: 18, fontFace: C.FONT, color: C.NAVY, italic: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    // Optional attribution (e.g. for a quote from a famous person)
    if (d.attribution) {
      slide.addText("— " + d.attribution, {
        x: 1.0, y: 4.8, w: 8.0, h: 0.4,
        fontSize: 14, fontFace: C.FONT, color: C.GRAY_MED,
        align: "center", valign: "middle", margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildAnalogy(slide, d) {
    // VIDA REAL ↔ TÉRMINO TÉCNICO. The slide where the lesson stops feeling cold.
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.PURPLE);
    slide.addText(d.title || "Analogía", {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 26, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    // Poetic subtitle — the heart of the analogy
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.35,
        fontSize: 14, fontFace: C.FONT, color: C.NAVY, italic: true, margin: 0
      });
    }
    // Two column headers
    const tableY = 2.0;
    const leftLabel = d.leftLabel || "VIDA REAL";
    const rightLabel = d.rightLabel || d.techLabel || "EN PROGRAMACIÓN";
    slide.addText(leftLabel, {
      x: 1.7, y: tableY, w: 3.0, h: 0.35,
      fontSize: 12, fontFace: C.FONT, color: C.PURPLE, bold: true,
      letterSpacing: 3, align: "center", valign: "middle", margin: 0
    });
    slide.addText(rightLabel, {
      x: 6.0, y: tableY, w: 3.0, h: 0.35,
      fontSize: 12, fontFace: C.FONT, color: C.PURPLE, bold: true,
      letterSpacing: 3, align: "center", valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.0, y: tableY + 0.4, w: 8.0, h: 0.02, fill: { color: C.PURPLE }
    });
    // Rows
    const rows = d.rows || [];
    rows.slice(0, 5).forEach((r, i) => {
      const y = tableY + 0.6 + i * 0.6;
      // Pink number
      placePinkNumber(slide, i + 1, 0.6, y, 0.32);
      // Left side (vida real)
      const isCodeLeft = typeof r.left === "string" && /[{};()=<>]/.test(r.left);
      slide.addText(r.left || "", {
        x: 1.05, y, w: 3.6, h: 0.5,
        fontSize: 14, fontFace: isCodeLeft ? C.FONT_CODE : C.FONT,
        color: C.NAVY, valign: "middle", margin: 0
      });
      // Arrow
      slide.addText("→", {
        x: 4.65, y, w: 0.6, h: 0.5,
        fontSize: 20, fontFace: C.FONT, color: C.YELLOW, bold: true,
        align: "center", valign: "middle", margin: 0
      });
      // Right side (tech)
      const isCodeRight = typeof r.right === "string" && /[{};()=<>]/.test(r.right);
      slide.addText([
        ...(r.rightLabel ? [{ text: r.rightLabel + "\n", options: { color: C.PURPLE, bold: true } }] : []),
        { text: r.right || "", options: { color: C.NAVY, fontFace: isCodeRight ? C.FONT_CODE : C.FONT } },
      ], {
        x: 5.3, y, w: 4.0, h: 0.5,
        fontSize: 13, fontFace: C.FONT, valign: "middle", margin: 0,
        lineSpacingMultiple: 1.2
      });
    });
    slide.addNotes(d.notes || "");
  }

  async function buildComparison(slide, d) {
    // Side-by-side "sin X / con X" or "antes / después".
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.PURPLE);
    slide.addText(d.title, {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 24, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.35,
        fontSize: 14, fontFace: C.FONT, color: C.NAVY, italic: true, margin: 0
      });
    }
    // Two columns. Each can render either body text or code (auto-detected).
    const sides = [
      { x: 0.6, label: d.leftLabel || "ANTES", labelColor: C.RED, body: d.left, file: d.leftFile },
      { x: 5.1, label: d.rightLabel || "DESPUÉS", labelColor: C.GREEN, body: d.right, file: d.rightFile },
    ];
    const cardY = 2.0;
    const cardH = 3.2;
    sides.forEach(s => {
      // Card background
      slide.addShape(pres.shapes.RECTANGLE, {
        x: s.x, y: cardY, w: 4.3, h: cardH,
        fill: { color: C.NAVY }, line: { color: C.NAVY, width: 0 }
      });
      // Label tag
      slide.addShape(pres.shapes.RECTANGLE, {
        x: s.x, y: cardY, w: 1.4, h: 0.4, fill: { color: s.labelColor }
      });
      slide.addText(s.label, {
        x: s.x, y: cardY, w: 1.4, h: 0.4,
        fontSize: 11, fontFace: C.FONT, color: C.WHITE, bold: true,
        letterSpacing: 2, align: "center", valign: "middle", margin: 0
      });
      // Optional filename (e.g. "sin-funciones.cpp")
      if (s.file) {
        slide.addText(s.file, {
          x: s.x + 1.5, y: cardY + 0.05, w: 2.7, h: 0.3,
          fontSize: 11, fontFace: C.FONT_CODE, color: C.YELLOW, margin: 0
        });
      }
      // Body — accept either array of code-line segments OR plain string array
      const body = s.body || [];
      const isCodeBlock = Array.isArray(body) && body.length && Array.isArray(body[0]);
      if (isCodeBlock) {
        // Auto-scale font for code so it fits inside the comparison card.
        // Card body area: 3.9" x (cardH - 0.7)". For cardH=3.2, that's 2.5" tall.
        const lineCount = body.length;
        const autoFontSize = lineCount > 11 ? 9 : (lineCount > 8 ? 10 : 12);
        const codeContent = buildCodeLines(body, autoFontSize);
        slide.addText(codeContent, {
          x: s.x + 0.2, y: cardY + 0.55, w: 3.9, h: cardH - 0.7,
          fontSize: autoFontSize, fontFace: C.FONT_CODE, color: C.CODE_DEFAULT,
          valign: "top", margin: 0, lineSpacingMultiple: 1.2
        });
      } else {
        const para = body.map(line => ({
          text: typeof line === "string" ? line : (line.text || ""),
          options: { breakLine: true, color: C.WHITE, fontSize: 14 }
        }));
        slide.addText(para, {
          x: s.x + 0.2, y: cardY + 0.55, w: 3.9, h: cardH - 0.7,
          fontFace: C.FONT, valign: "top", margin: 0, lineSpacingMultiple: 1.35
        });
      }
    });
    // Optional takeaway line (kept inside slide bounds)
    if (d.takeaway) {
      slide.addText(d.takeaway, {
        x: 0.6, y: 5.3, w: 8.8, h: 0.3,
        fontSize: 12, fontFace: C.FONT, color: C.PURPLE, italic: true, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildAnatomy(slide, d) {
    // Break a structure into 4 labeled parts.
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.PURPLE);
    slide.addText(d.title || "Anatomía", {
      x: 0.75, y: 0.85, w: 8.0, h: 0.55,
      fontSize: 26, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.45, w: 8.5, h: 0.35,
        fontSize: 14, fontFace: C.FONT, color: C.NAVY, italic: true, margin: 0
      });
    }
    // Optional centered code/structure example at top
    if (d.example) {
      const exY = 2.0;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 1.4, y: exY, w: 7.2, h: 0.8,
        fill: { color: C.NAVY }, line: { color: C.NAVY, width: 0 }
      });
      slide.addText(d.example, {
        x: 1.6, y: exY, w: 6.8, h: 0.8,
        fontSize: 18, fontFace: C.FONT_CODE, color: C.WHITE,
        align: "center", valign: "middle", margin: 0
      });
    }
    // Up to 4 parts in 2x2 grid (or row of 4 if shorter)
    const parts = (d.parts || []).slice(0, 4);
    const cardsY = d.example ? 3.05 : 2.0;
    const cols = parts.length <= 2 ? parts.length : 2;
    const rows = Math.ceil(parts.length / cols);
    const cardW = (9.4 - (cols + 1) * 0.2) / cols;
    const cardH = rows === 1 ? 2.4 : (rows === 2 ? 1.2 : 0.9);
    parts.forEach((p, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = 0.6 + col * (cardW + 0.2);
      const y = cardsY + row * (cardH + 0.2);
      const color = colorVal(p.color) || [C.PURPLE, C.BLUE, C.GREEN, C.YELLOW][i % 4];
      // Card
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w: cardW, h: cardH,
        fill: { color: C.LAVENDER_LIGHT },
        line: { color, width: 2 }
      });
      // Top color bar
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y, w: cardW, h: 0.08, fill: { color }
      });
      // Label
      slide.addText((p.label || "").toUpperCase(), {
        x: x + 0.15, y: y + 0.18, w: cardW - 0.3, h: 0.35,
        fontSize: 12, fontFace: C.FONT, color, bold: true,
        letterSpacing: 2, valign: "middle", margin: 0
      });
      // Code snippet (Space Mono) — emphasized
      if (p.code) {
        slide.addText(p.code, {
          x: x + 0.15, y: y + 0.55, w: cardW - 0.3, h: 0.4,
          fontSize: 15, fontFace: C.FONT_CODE, color: C.NAVY, bold: true,
          valign: "middle", margin: 0
        });
      }
      // Description
      if (p.desc) {
        slide.addText(p.desc, {
          x: x + 0.15, y: y + (p.code ? 0.95 : 0.55), w: cardW - 0.3, h: cardH - (p.code ? 1.0 : 0.6),
          fontSize: 12, fontFace: C.FONT, color: C.NAVY,
          valign: "top", margin: 0, lineSpacingMultiple: 1.25
        });
      }
    });
    slide.addNotes(d.notes || "");
  }

  async function buildFlowchartDiagram(slide, d) {
    // Render an ACTUAL flowchart with shapes + arrows (not just symbol legend).
    // Use after `flowchart` (symbol legend) to show students a concrete example.
    slide.background = { color: C.LAVENDER_BG };
    placeBrandFurniture(slide, { sectionLabel: d.sectionLabel });
    placeAccent(slide, 0.85, colorVal(d.accentColor) || C.PURPLE);
    slide.addText(d.title || "Diagrama de flujo", {
      x: 0.75, y: 0.85, w: 8.0, h: 0.5,
      fontSize: 24, fontFace: C.FONT, color: C.PURPLE, bold: true,
      valign: "middle", margin: 0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 0.75, y: 1.4, w: 8.5, h: 0.3,
        fontSize: 13, fontFace: C.FONT, color: C.NAVY, italic: true, margin: 0
      });
    }
    // Adaptive sizing: shrink for many nodes to fit inside the slide
    const nodes = d.nodes || [];
    const startY = d.subtitle ? 1.85 : 1.55;
    const availableH = 5.4 - startY;  // 5.4" max
    // Reserve 0.18" arrow + node height — compute node height that fits
    const arrowH = nodes.length > 6 ? 0.15 : 0.2;
    const maxNodeH = Math.min(0.42, (availableH - (nodes.length - 1) * arrowH) / nodes.length);
    const nodeH = Math.max(0.28, maxNodeH);
    const stepH = nodeH + arrowH;
    const totalH = nodes.length * nodeH + (nodes.length - 1) * arrowH;
    const verticalShift = Math.max(0, (availableH - totalH) / 2);
    const nodeW = 3.6;
    const nodeX = 3.2;  // centered (slide is 10" wide)
    nodes.forEach((n, i) => {
      const y = startY + verticalShift + i * stepH;
      const shape = (n.shape || "rectangle").toLowerCase();
      const color = colorVal(n.color) || (shape === "oval" ? C.PURPLE : shape === "rhombus" || shape === "rombo" || shape === "diamond" ? C.GREEN : C.BLUE);
      // Draw shape
      if (shape === "oval" || shape === "ovalo" || shape === "óvalo") {
        slide.addShape(pres.shapes.OVAL, {
          x: nodeX, y, w: nodeW, h: nodeH, fill: { color }, line: { color, width: 0 }
        });
      } else if (shape === "rhombus" || shape === "rombo" || shape === "diamond") {
        slide.addShape(pres.shapes.DIAMOND, {
          x: nodeX, y, w: nodeW, h: nodeH, fill: { color }, line: { color, width: 0 }
        });
      } else if (shape === "parallelogram" || shape === "paralelogramo") {
        slide.addShape(pres.shapes.PARALLELOGRAM, {
          x: nodeX, y, w: nodeW, h: nodeH, fill: { color }, line: { color, width: 0 }
        });
      } else {
        slide.addShape(pres.shapes.RECTANGLE, {
          x: nodeX, y, w: nodeW, h: nodeH, fill: { color }, line: { color, width: 0 }
        });
      }
      // Text inside the shape — font scales with node height
      const textSize = nodeH < 0.35 ? 11 : (nodeH < 0.4 ? 12 : 13);
      slide.addText(n.text || "", {
        x: nodeX, y, w: nodeW, h: nodeH,
        fontSize: textSize, fontFace: C.FONT, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0
      });
      // Arrow below (except after last node)
      if (i < nodes.length - 1) {
        slide.addText("↓", {
          x: nodeX, y: y + nodeH, w: nodeW, h: arrowH,
          fontSize: 14, fontFace: C.FONT, color: C.NAVY, bold: true,
          align: "center", valign: "middle", margin: 0
        });
      }
    });
    slide.addNotes(d.notes || "");
  }

  async function buildTryIt(slide, d) {
    // Warm invitation before the practice block. "Inténtalo tú. Antes de seguir, prueba tú."
    // Per Clase-09-Funciones.pptx slide 15.
    slide.background = { color: C.YELLOW };
    placeLogo(slide);
    placeStar(slide, 0.6, 0.5, 0.5, C.PURPLE);
    placeStar(slide, 8.8, 4.8, 0.45, C.PURPLE);
    placeStar(slide, 0.6, 4.8, 0.35, C.ROSE);
    placeStar(slide, 8.8, 0.5, 0.4, C.ROSE);
    // Big "INTÉNTALO TÚ" label
    slide.addText("INTÉNTALO TÚ", {
      x: 0.6, y: 0.7, w: 9, h: 0.5,
      fontSize: 14, fontFace: C.FONT, color: C.NAVY, bold: true,
      letterSpacing: 4, align: "center", valign: "middle", margin: 0
    });
    // The challenge title
    slide.addText(d.title || "Antes de seguir, prueba tú.", {
      x: 0.6, y: 1.3, w: 9, h: 0.8,
      fontSize: 32, fontFace: C.FONT, color: C.NAVY, bold: true,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.1
    });
    // The challenge prompt
    if (d.prompt) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 1.0, y: 2.3, w: 8.0, h: 2.0,
        fill: { color: C.WHITE }, line: { color: C.NAVY, width: 2 }
      });
      slide.addText(d.prompt, {
        x: 1.3, y: 2.4, w: 7.4, h: 1.8,
        fontSize: 16, fontFace: C.FONT, color: C.NAVY,
        valign: "middle", margin: 0, lineSpacingMultiple: 1.4
      });
    }
    // Time / mode footer
    const meta = [d.duration, d.mode].filter(Boolean).join("  ·  ");
    if (meta) {
      slide.addText(meta, {
        x: 0.6, y: 4.5, w: 9, h: 0.4,
        fontSize: 14, fontFace: C.FONT, color: C.NAVY, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    // Warm hint at the very bottom
    if (d.hint) {
      slide.addText(d.hint, {
        x: 0.6, y: 5.0, w: 9, h: 0.4,
        fontSize: 13, fontFace: C.FONT, color: C.NAVY, italic: true,
        align: "center", valign: "middle", margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  async function buildFarewell(slide, d) {
    // Cierre rico de la clase. Reemplaza al simple pull-quote.
    // Estructura: fondo púrpura + smiley + cursor pixel + mensaje grande + 3 logros + bridge a próxima clase
    slide.background = { color: C.PURPLE };
    placeStar(slide, 0.5, 0.4, 0.45, C.YELLOW);
    placeStar(slide, 8.9, 0.4, 0.4, C.ROSE);
    placeStar(slide, 0.5, 5.0, 0.35, C.TEAL);
    placeStar(slide, 8.9, 5.0, 0.5, C.YELLOW);
    // Pixel smiley + cursor (brandbook signature)
    const smileyData = await iconToDataUri("pixel-smiley", null, 256);
    if (smileyData) slide.addImage({ data: smileyData, x: 1.2, y: 0.9, w: 1.4, h: 1.4 });
    const cursorData = await iconToDataUri("pixel-cursor", null, 128);
    if (cursorData) slide.addImage({ data: cursorData, x: 2.4, y: 2.0, w: 0.45, h: 0.45 });
    // Main message
    slide.addText(d.title || "¡Gracias por hoy!", {
      x: 3.0, y: 1.0, w: 6.5, h: 0.9,
      fontSize: 40, fontFace: C.FONT, color: C.WHITE, bold: true,
      valign: "middle", margin: 0, lineSpacingMultiple: 1.0
    });
    if (d.subtitle) {
      slide.addText(d.subtitle, {
        x: 3.0, y: 1.95, w: 6.5, h: 0.4,
        fontSize: 16, fontFace: C.FONT, color: C.YELLOW, italic: true,
        valign: "middle", margin: 0
      });
    }
    // Logros — hasta 3 cosas que las alumnas pueden hacer ahora
    const achievements = (d.achievements || []).slice(0, 3);
    if (achievements.length) {
      slide.addText("HOY APRENDISTE A", {
        x: 0.6, y: 2.85, w: 9, h: 0.3,
        fontSize: 11, fontFace: C.FONT, color: C.YELLOW, bold: true,
        letterSpacing: 3, align: "center", margin: 0
      });
      const colW = 9.0 / Math.max(1, achievements.length);
      achievements.forEach((a, i) => {
        const x = 0.5 + i * colW;
        // Card on lavender bg
        slide.addShape(pres.shapes.RECTANGLE, {
          x: x + 0.15, y: 3.3, w: colW - 0.3, h: 1.3,
          fill: { color: "8550F0" },
          line: { color: "8550F0", width: 0 }
        });
        slide.addShape(pres.shapes.RECTANGLE, {
          x: x + 0.15, y: 3.3, w: colW - 0.3, h: 0.06, fill: { color: C.YELLOW }
        });
        slide.addText(typeof a === "string" ? a : (a.text || a.label || ""), {
          x: x + 0.25, y: 3.4, w: colW - 0.5, h: 1.15,
          fontSize: 13, fontFace: C.FONT, color: C.WHITE,
          align: "center", valign: "middle", margin: 0,
          lineSpacingMultiple: 1.3
        });
      });
    }
    // Bridge a próxima clase / mensaje cálido final
    const nextLine = d.nextClass
      ? `Nos vemos en la ${d.nextClass}.`
      : (d.farewell || "Nos vemos pronto.");
    slide.addText(nextLine, {
      x: 0.6, y: 4.85, w: 9, h: 0.4,
      fontSize: 18, fontFace: C.FONT, color: C.WHITE, italic: true,
      align: "center", valign: "middle", margin: 0
    });
    if (d.signature) {
      slide.addText(d.signature, {
        x: 0.6, y: 5.25, w: 9, h: 0.3,
        fontSize: 12, fontFace: C.FONT, color: C.ROSE, bold: true,
        letterSpacing: 2, align: "center", valign: "middle", margin: 0
      });
    }
    slide.addNotes(d.notes || "");
  }

  // ─── Dispatch ───
  const builders = {
    title: buildTitle,
    section: buildSection,
    welcome: buildWelcome,
    bridge: buildBridge,
    "pull-quote": buildPullQuote,
    farewell: buildFarewell,
    "block-transition": buildBlockTransition,
    engage: buildEngage,
    objectives: buildObjectives,
    analogy: buildAnalogy,
    steps: buildSteps,
    activity: buildActivity,
    content: buildContent,
    definition: buildDefinition,
    pillars: buildPillars,
    code: buildCode,
    comparison: buildComparison,
    anatomy: buildAnatomy,
    "manual-trace": buildManualTrace,
    flowchart: buildFlowchart,
    "flowchart-diagram": buildFlowchartDiagram,
    "error-code": buildErrorCode,
    reference: buildReference,
    glossary: buildGlossary,
    "try-it": buildTryIt,
    kahoot: buildKahoot,
    steam: buildSteam,
    reflection: buildReflection,
    rubric: buildRubric,
    "plan-b": buildPlanB,
    closing: buildClosing,
    "closing-ceremony": buildClosingCeremony,
  };

  for (const slideData of config.slides) {
    const builder = builders[slideData.type];
    if (!builder) {
      console.warn("Unknown slide type:", slideData.type);
      continue;
    }
    const slide = pres.addSlide();
    await builder(slide, slideData);
    // Ensure logo is on EVERY slide (brandbook requirement) — applied last to be on top.
    // Builders that already place the logo are idempotent enough; the second placement
    // overlaps the first.
    placeLogo(slide);
  }

  const outputPath = config.outputPath || "./slides.pptx";
  await pres.writeFile({ fileName: outputPath });
  console.log("Slides created:", outputPath);
}

module.exports = {
  generateSlides,
  validateSlideSequence,
  buildClassSkeleton,
  C,
  CODE_TOKENS,
  BRAND_ICONS,
};
