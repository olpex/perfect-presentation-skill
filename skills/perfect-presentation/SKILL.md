---
name: perfect-presentation
description: >
  Use this skill when creating ANY presentation, slide deck, or .pptx file that requires
  professional quality. Triggers: "зроби презентацію", "create presentation", "make slides",
  "pitch deck", "slide deck", any topic + "presentation". Requirements: minimum 15 slides,
  deep topic research, infographics, premium design, speaker notes ≥500 chars per slide.
version: 1.0.0
---

# Perfect Presentation Skill

## Overview

This skill produces **professional-grade presentations** with:
- **≥15 slides** with rich, researched content
- **Infographics** — charts, diagrams, timelines, process flows
- **Premium design** — cohesive palette, typography, layout variety
- **Speaker notes** — ≥500 characters per slide (detailed talking points)

---

## Phase 1 — Deep Research (MANDATORY, do not skip)

Before writing a single slide, conduct thorough research on the topic.

### Research Checklist

```
□ Core definition and scope of the topic
□ Historical background and evolution
□ Key statistics and data points (with years)
□ Major trends shaping the field right now
□ Top players / stakeholders / organizations
□ Challenges and pain points
□ Opportunities and future outlook
□ Case studies or real-world examples (min. 2)
□ Expert opinions or frameworks
□ Contrarian or nuanced perspectives
```

### Research Output Format

Organize findings before building slides:

```markdown
## Research Summary: [TOPIC]

### Core Facts
- [fact 1 + source/year]
- [fact 2 + source/year]

### Key Statistics
- [stat]: [number/percentage] ([year])
- [stat]: [number/percentage] ([year])

### Timeline
- [year]: [event]
- [year]: [event]

### Case Studies
1. [Company/Person]: [what they did] → [result]
2. [Company/Person]: [what they did] → [result]

### Trends
1. [trend name]: [explanation]
2. [trend name]: [explanation]

### Challenges
- [challenge]: [details]

### Future Outlook
- [prediction/opportunity]
```

---

## Phase 2 — Slide Architecture (≥15 slides)

### Required Structure

Every presentation MUST follow this structure:

| # | Slide Type | Purpose |
|---|-----------|---------|
| 1 | **Title Slide** | Topic, subtitle, presenter, date |
| 2 | **Agenda / Table of Contents** | Visual overview of sections |
| 3 | **Executive Summary** | Key takeaways (max 4 points) |
| 4 | **Context / Why This Matters** | Problem statement or relevance |
| 5 | **Historical Background** | Origin, evolution, timeline |
| 6 | **Key Statistics Infographic** | Data visualization slide |
| 7 | **Section 1 — Core Concept** | Deep dive #1 |
| 8 | **Section 1 — Details / Examples** | Supporting data or case |
| 9 | **Section 2 — Analysis** | Deep dive #2 |
| 10 | **Section 2 — Infographic** | Chart, diagram, or process flow |
| 11 | **Section 3 — Trends** | Current state + trends |
| 12 | **Section 3 — Data** | Numbers that support trends |
| 13 | **Case Study** | Real-world example with outcome |
| 14 | **Challenges & Risks** | Honest assessment |
| 15 | **Opportunities & Solutions** | Forward-looking recommendations |
| 16 | **Future Outlook** | Predictions, roadmap, or scenarios |
| 17 | **Key Takeaways** | Summary of top 5 insights |
| 18 | **Call to Action / Next Steps** | What the audience should do |
| 19 | **Q&A / Thank You** | Closing slide |

> For technical/complex topics: add slides 20–25 as deep-dives or appendix

---

## Phase 3 — Infographic Design

### Infographic Slide Types

Use at least **4 different infographic types** across the presentation:

#### 1. Stat Callout Slide
```
Layout: 3–4 large stat boxes side by side
Each box: BIG number (60pt) + label (14pt) + context (12pt)
Background: accent color or dark background
Example:
  [87%]          [$4.2T]        [3.2x]
  Adoption Rate  Market Size    ROI Multiplier
  by 2025        projected      vs. baseline
```

#### 2. Timeline Infographic
```
Layout: horizontal or vertical timeline
Elements: year markers, event labels, icons, connecting lines
Use shapes: circles for points, rectangles for labels
Color: alternate between primary and accent colors
```

#### 3. Process Flow / Steps
```
Layout: numbered steps with arrows
Elements: icon circle + step number + title + description
Direction: left-to-right or top-to-bottom
Colors: gradient progression through palette
```

#### 4. Comparison Grid
```
Layout: 2–4 column comparison
Elements: category headers, checkmarks/X marks, feature rows
Use: before/after, option A vs B, pros vs cons
```

#### 5. Chart Infographic
```
Chart types by use case:
- BAR: comparisons between categories
- LINE: trends over time
- PIE/DOUGHNUT: market share, composition
- SCATTER: correlations

Styling: match presentation palette, hide gridlines, show data labels
```

#### 6. Icon Grid
```
Layout: 2×3 or 3×3 grid
Each cell: icon (0.6") + title (bold) + 2-line description
Background: subtle card with shadow
```

---

## Phase 4 — Design System

### Step 1: Choose Color Palette

Pick ONE palette that matches the topic:

| Theme | Primary | Secondary | Accent | Text |
|-------|---------|-----------|--------|------|
| **Tech / AI** | `1E3A5F` | `0EA5E9` | `38BDF8` | `F8FAFC` |
| **Finance / Business** | `1C3144` | `D4AF37` | `F5F5F5` | `1C3144` |
| **Health / Medical** | `065F46` | `10B981` | `D1FAE5` | `064E3B` |
| **Education** | `4338CA` | `818CF8` | `EEF2FF` | `1E1B4B` |
| **Environment** | `14532D` | `4ADE80` | `DCFCE7` | `052E16` |
| **Marketing** | `9D174D` | `F472B6` | `FDF2F8` | `500724` |
| **Strategy** | `1E2761` | `7A2FBE` | `C4B5FD` | `F5F3FF` |
| **Innovation** | `C2410C` | `FB923C` | `FFF7ED` | `431407` |
| **Dark Premium** | `0F172A` | `6366F1` | `818CF8` | `F1F5F9` |
| **Minimal White** | `18181B` | `71717A` | `F4F4F5` | `09090B` |

### Step 2: Typography Rules

```
Title slides:     fontFace: "Georgia",    fontSize: 44, bold: true
Section headers:  fontFace: "Arial Black", fontSize: 28, bold: true
Body text:        fontFace: "Calibri",    fontSize: 14-16
Captions:         fontFace: "Calibri",    fontSize: 11, color: muted
Stats/Numbers:    fontFace: "Arial Black", fontSize: 52-72, bold: true
```

### Step 3: Layout Variety

Use a **different layout for every 2–3 slides**. Never repeat the same layout twice in a row.

| Layout Name | Description |
|-------------|-------------|
| `hero-text` | Full-bleed background, centered large text |
| `split-50-50` | Text left 50%, visual right 50% |
| `split-40-60` | Text 40%, large visual 60% |
| `card-grid` | 2×2 or 3×2 cards with icons |
| `stat-row` | 3–4 stat callout boxes in a row |
| `timeline` | Horizontal timeline with events |
| `full-chart` | Chart takes 80% of slide |
| `quote-callout` | Large italic quote + attribution |
| `process-steps` | Numbered steps with arrows |
| `two-column-list` | Two columns of bullet points |

### Step 4: Background Strategy

```
Slide 1 (Title):      Dark primary background
Slides 2–3:           Light background (off-white or very pale)
Content slides:       Alternate light / slightly colored
Infographic slides:   Dark or rich color for emphasis
Final slides:         Dark primary background (mirror title)
```

### Step 5: Visual Motif

Pick ONE motif and apply to every slide:
- Colored left-side accent bar (0.12" rectangle)
- Rounded card containers with shadow
- Icon circles (colored circles with white icons)
- Diagonal color split on background

---

## Phase 5 — Speaker Notes (MANDATORY: ≥500 chars per slide)

Every slide MUST have detailed speaker notes. Notes serve as the full script.

### Notes Template

```
SPEAKER NOTES — Slide [N]: [Title]

OPENING (what to say first):
[2-3 sentences to open this slide]

KEY POINTS TO EMPHASIZE:
• [Point 1]: [detailed explanation — 2 sentences]
• [Point 2]: [detailed explanation — 2 sentences]
• [Point 3]: [detailed explanation — 2 sentences]

DATA / EVIDENCE:
[Specific numbers, studies, examples to cite verbally]

TRANSITION:
[How to move to the next slide — bridge sentence]

TIMING: ~[X] minutes
```

### Notes Quality Standards

- **Minimum**: 500 characters (~75-90 words)
- **Target**: 700-900 characters (~110-140 words)
- **Tone**: Conversational, as if the speaker is talking naturally
- **Must include**: At least one specific data point or example
- **Must include**: Transition phrase to next slide

---

## Phase 6 — Implementation (PptxGenJS)

### Setup

```bash
npm install -g pptxgenjs react react-dom react-icons sharp
```

### Boilerplate Structure

```javascript
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";  // 10" × 5.625"

// ============================================================
// DESIGN TOKENS — define ONCE, use everywhere
// ============================================================
const C = {
  primary:   "1E3A5F",   // Dark blue
  secondary: "0EA5E9",   // Sky blue
  accent:    "38BDF8",   // Light blue
  textDark:  "0F172A",
  textLight: "F8FAFC",
  muted:     "64748B",
  white:     "FFFFFF",
  cardBg:    "F1F5F9",
};

const F = {
  title:  { fontFace: "Georgia",    fontSize: 40, bold: true,  color: C.textLight },
  h1:     { fontFace: "Arial Black", fontSize: 28, bold: true,  color: C.primary   },
  h2:     { fontFace: "Arial Black", fontSize: 20, bold: true,  color: C.primary   },
  body:   { fontFace: "Calibri",    fontSize: 15,              color: C.textDark  },
  small:  { fontFace: "Calibri",    fontSize: 11,              color: C.muted     },
  stat:   { fontFace: "Arial Black", fontSize: 56, bold: true,  color: C.secondary },
};

const makeShadow = () => ({
  type: "outer", blur: 8, offset: 3, angle: 135,
  color: "000000", opacity: 0.12
});

// ============================================================
// HELPER — add slide with common elements
// ============================================================
function addContentSlide(title, notes) {
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  // Left accent bar (visual motif)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.12, h: 5.625,
    fill: { color: C.primary }, line: { color: C.primary }
  });

  // Title
  slide.addText(title, {
    x: 0.3, y: 0.2, w: 9.4, h: 0.7,
    ...F.h1, margin: 0
  });

  // Bottom bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.3, w: 10, h: 0.325,
    fill: { color: C.primary }, line: { color: C.primary }
  });

  // Notes
  slide.addNotes(notes);

  return slide;
}

// ============================================================
// HELPER — stat box
// ============================================================
function addStatBox(slide, x, y, w, h, number, label, sublabel, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: color || C.cardBg },
    shadow: makeShadow(),
    line: { color: "E2E8F0", width: 0.5 }
  });
  slide.addText(number, {
    x: x + 0.1, y: y + 0.1, w: w - 0.2, h: h * 0.5,
    ...F.stat, align: "center", margin: 0
  });
  slide.addText(label, {
    x: x + 0.1, y: y + h * 0.55, w: w - 0.2, h: h * 0.22,
    fontFace: "Calibri", fontSize: 13, bold: true,
    color: C.textDark, align: "center", margin: 0
  });
  if (sublabel) {
    slide.addText(sublabel, {
      x: x + 0.1, y: y + h * 0.77, w: w - 0.2, h: h * 0.18,
      ...F.small, align: "center", margin: 0
    });
  }
}
```

### Slide Patterns

#### Pattern: Title Slide
```javascript
function buildTitleSlide(title, subtitle, presenter, date) {
  const slide = pres.addSlide();
  slide.background = { color: C.primary };

  // Large brand stripe
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 3.8, w: 10, h: 0.06,
    fill: { color: C.secondary }, line: { color: C.secondary }
  });

  // Main title
  slide.addText(title, {
    x: 0.8, y: 1.2, w: 8.4, h: 1.6,
    ...F.title, fontSize: 44, align: "left", valign: "middle",
    margin: 0
  });

  // Subtitle
  slide.addText(subtitle, {
    x: 0.8, y: 2.9, w: 8.4, h: 0.6,
    fontFace: "Calibri", fontSize: 18, color: C.accent,
    align: "left", margin: 0
  });

  // Presenter + date
  slide.addText(`${presenter}  |  ${date}`, {
    x: 0.8, y: 4.9, w: 8.4, h: 0.4,
    ...F.small, color: "94A3B8", align: "left", margin: 0
  });

  slide.addNotes("Привітайте аудиторію. Представтесь і поясніть чому ця тема важлива саме зараз. Дайте 30 секунд для того щоб аудиторія зрозуміла тему і налаштувалась. Поясніть структуру: скільки часу займе, чи будуть питання в кінці чи по ходу. Назвіть 3 головних питання на які слухачі отримають відповіді.");
  return slide;
}
```

#### Pattern: Agenda Slide
```javascript
function buildAgendaSlide(sections) {
  // sections = [{number: "01", title: "...", description: "..."}]
  const slide = addContentSlide("Agenda", "...");
  const cols = sections.length <= 4 ? 1 : 2;
  // Draw numbered cards for each section
  sections.forEach((s, i) => {
    const x = cols === 1 ? 1.0 : (i % 2 === 0 ? 0.5 : 5.2);
    const y = cols === 1 ? 1.2 + i * 0.85 : 1.2 + Math.floor(i / 2) * 0.85;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cols === 1 ? 8 : 4.5, h: 0.7,
      fill: { color: i % 2 === 0 ? C.cardBg : C.white },
      shadow: makeShadow(), line: { color: "E2E8F0", width: 0.5 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.45, h: 0.7,
      fill: { color: C.secondary }, line: { color: C.secondary }
    });
    slide.addText(s.number, {
      x: x, y: y, w: 0.45, h: 0.7,
      fontFace: "Arial Black", fontSize: 16, bold: true,
      color: C.white, align: "center", valign: "middle", margin: 0
    });
    slide.addText(s.title, {
      x: x + 0.55, y: y + 0.05, w: (cols === 1 ? 7.4 : 3.9), h: 0.35,
      fontFace: "Calibri", fontSize: 14, bold: true,
      color: C.primary, margin: 0
    });
    slide.addText(s.description, {
      x: x + 0.55, y: y + 0.38, w: (cols === 1 ? 7.4 : 3.9), h: 0.27,
      ...F.small, margin: 0
    });
  });
  return slide;
}
```

#### Pattern: Stat Infographic Slide (3 stats)
```javascript
function buildStatSlide(title, stats, notes) {
  // stats = [{number, label, sublabel}]
  const slide = addContentSlide(title, notes);
  const w = (10 - 1.2) / stats.length - 0.2;
  stats.forEach((s, i) => {
    addStatBox(slide, 0.4 + i * (w + 0.2), 1.2, w, 3.6,
      s.number, s.label, s.sublabel);
  });
  return slide;
}
```

#### Pattern: Timeline Slide
```javascript
function buildTimelineSlide(title, events, notes) {
  // events = [{year, title, description}]
  const slide = addContentSlide(title, notes);
  const y_line = 3.0;
  const startX = 0.5;
  const endX = 9.5;
  const step = (endX - startX) / (events.length - 1);

  // Main line
  slide.addShape(pres.shapes.LINE, {
    x: startX, y: y_line, w: endX - startX, h: 0,
    line: { color: C.secondary, width: 2.5 }
  });

  events.forEach((ev, i) => {
    const x = startX + i * step;
    const isTop = i % 2 === 0;

    // Circle dot
    slide.addShape(pres.shapes.OVAL, {
      x: x - 0.12, y: y_line - 0.12, w: 0.24, h: 0.24,
      fill: { color: C.secondary }, line: { color: C.white, width: 1.5 }
    });

    // Year
    slide.addText(ev.year, {
      x: x - 0.5, y: isTop ? y_line - 0.5 : y_line + 0.2,
      w: 1, h: 0.3,
      fontFace: "Arial Black", fontSize: 13, bold: true,
      color: C.primary, align: "center", margin: 0
    });

    // Event title
    slide.addText(ev.title, {
      x: x - 0.75, y: isTop ? y_line - 1.3 : y_line + 0.5,
      w: 1.5, h: 0.8,
      fontFace: "Calibri", fontSize: 11, bold: false,
      color: C.textDark, align: "center", margin: 0
    });
  });
  return slide;
}
```

#### Pattern: Card Grid (2×2 or 2×3)
```javascript
function buildCardGridSlide(title, cards, notes) {
  // cards = [{icon: FaIcon, title, description}]
  const slide = addContentSlide(title, notes);
  const cols = 2;
  const rows = Math.ceil(cards.length / cols);
  const cardW = 4.2, cardH = rows === 2 ? 1.9 : 1.3;
  const startX = 0.3, startY = 1.1;
  const gapX = 0.3, gapY = 0.25;

  cards.forEach((card, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.white }, shadow: makeShadow(),
      line: { color: "E2E8F0", width: 0.75 }
    });
    // Left accent
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: cardH,
      fill: { color: C.secondary }, line: { color: C.secondary }
    });
    slide.addText(card.title, {
      x: x + 0.15, y: y + 0.1, w: cardW - 0.25, h: 0.38,
      fontFace: "Calibri", fontSize: 14, bold: true,
      color: C.primary, margin: 0
    });
    slide.addText(card.description, {
      x: x + 0.15, y: y + 0.48, w: cardW - 0.25, h: cardH - 0.6,
      ...F.body, fontSize: 12, margin: 0
    });
  });
  return slide;
}
```

#### Pattern: Split Layout (text + chart)
```javascript
function buildSplitChartSlide(title, textPoints, chartData, chartType, notes) {
  const slide = addContentSlide(title, notes);

  // Text column
  slide.addText(textPoints.map(t => ({
    text: t, options: { bullet: true, breakLine: true,
      fontFace: "Calibri", fontSize: 13, color: C.textDark,
      paraSpaceAfter: 8 }
  })), { x: 0.3, y: 1.1, w: 3.8, h: 4.0 });

  // Chart column
  slide.addChart(pres.charts[chartType], chartData, {
    x: 4.4, y: 0.9, w: 5.2, h: 4.4,
    chartColors: [C.secondary, C.accent, C.primary],
    chartArea: { fill: { color: C.white }, roundedCorners: true },
    catAxisLabelColor: C.muted,
    valAxisLabelColor: C.muted,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelColor: C.textDark,
    showLegend: true, legendPos: "b",
    legendColor: C.muted,
  });
  return slide;
}
```

#### Pattern: Process Steps
```javascript
function buildProcessSlide(title, steps, notes) {
  // steps = [{number, title, description}]
  const slide = addContentSlide(title, notes);
  const stepW = (9.2) / steps.length - 0.15;
  const colors = [C.primary, C.secondary, "0D9488", "7C3AED", "DC2626"];

  steps.forEach((step, i) => {
    const x = 0.4 + i * (stepW + 0.15);
    const y = 1.3;
    const color = colors[i % colors.length];

    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: stepW, h: 3.8,
      fill: { color: "F8FAFC" }, shadow: makeShadow(),
      line: { color: "E2E8F0", width: 0.5 }
    });

    // Number badge
    slide.addShape(pres.shapes.OVAL, {
      x: x + stepW / 2 - 0.32, y: y + 0.2, w: 0.64, h: 0.64,
      fill: { color: color }, line: { color: color }
    });
    slide.addText(`${i + 1}`, {
      x: x + stepW / 2 - 0.32, y: y + 0.2, w: 0.64, h: 0.64,
      fontFace: "Arial Black", fontSize: 18, bold: true,
      color: C.white, align: "center", valign: "middle", margin: 0
    });

    // Arrow connector (except last)
    if (i < steps.length - 1) {
      slide.addShape(pres.shapes.LINE, {
        x: x + stepW, y: y + 1.5, w: 0.15, h: 0,
        line: { color: C.muted, width: 1.5, dashType: "dash" }
      });
    }

    // Step title
    slide.addText(step.title, {
      x: x + 0.1, y: y + 1.0, w: stepW - 0.2, h: 0.55,
      fontFace: "Calibri", fontSize: 13, bold: true,
      color: color, align: "center", margin: 0
    });

    // Description
    slide.addText(step.description, {
      x: x + 0.1, y: y + 1.55, w: stepW - 0.2, h: 2.15,
      fontFace: "Calibri", fontSize: 11, color: C.textDark,
      align: "center", valign: "top", margin: 0
    });
  });
  return slide;
}
```

#### Pattern: Quote / Highlight Callout
```javascript
function buildQuoteSlide(title, quote, attribution, context, notes) {
  const slide = pres.addSlide();
  slide.background = { color: C.primary };

  // Large quotation mark
  slide.addText("\u201C", {
    x: 0.4, y: 0.3, w: 2, h: 1.5,
    fontFace: "Georgia", fontSize: 100, bold: true,
    color: C.secondary, align: "left", margin: 0
  });

  slide.addText(quote, {
    x: 1.2, y: 1.3, w: 7.8, h: 2.5,
    fontFace: "Georgia", fontSize: 22, italic: true,
    color: C.textLight, align: "left", valign: "middle", margin: 0
  });

  slide.addText(`— ${attribution}`, {
    x: 1.2, y: 3.9, w: 7.8, h: 0.45,
    fontFace: "Calibri", fontSize: 14, bold: true,
    color: C.accent, align: "left", margin: 0
  });

  slide.addText(context, {
    x: 1.2, y: 4.4, w: 7.8, h: 0.7,
    ...F.small, color: "94A3B8", align: "left", margin: 0
  });

  slide.addNotes(notes);
  return slide;
}
```

---

## Phase 7 — Speaker Notes Template Per Slide Type

### MANDATORY: Notes must be ≥500 characters

```javascript
// Example of proper notes for each slide type:

// Title slide notes (~600 chars):
`ВСТУП: Привітайте аудиторію та представтесь. Поясніть свою кваліфікацію та чому ця тема важлива для них особисто. 
СТРУКТУРА: Ця презентація складається з [N] частин і займе приблизно [X] хвилин. 
ГОЛОВНЕ ПИТАННЯ: Поставте перед аудиторією питання, на яке вони отримають відповідь до кінця. Наприклад: "Після сьогоднішньої презентації ви зрозумієте...".
ПЕРЕХІД: Давайте почнемо з огляду нашого плану.`

// Data/stats slide notes (~700 chars):
`КОНТЕКСТ: Поясніть звідки взяті ці дані — назвіть джерело та рік. 
ПЕРША ЦИФРА: [Number] означає [explanation]. Щоб краще розуміти масштаб — це [analogy].
ДРУГА ЦИФРА: [Number] — це [explanation]. Цей показник зріс/знизився на [%] порівняно з [year].
ТРЕТЯ ЦИФРА: [Number] показує [insight]. Це важливо тому що [reason].
ВИСНОВОК: Разом ці три числа говорять нам про [overall insight].
ПЕРЕХІД: Тепер подивимось детальніше на [next topic].`
```

---

## Phase 8 — QA Checklist

Before finalizing, verify ALL of the following:

```
CONTENT
□ ≥15 slides present
□ Each slide has substantive content (not just titles)
□ Data points are specific (numbers, years, names)
□ No placeholder text remaining
□ Logical flow from slide to slide

DESIGN
□ Consistent color palette throughout
□ Every slide has at least one visual element
□ No two consecutive slides use identical layout
□ Text contrast is sufficient (dark on light OR light on dark)
□ Font sizes hierarchy respected (title > header > body)
□ Consistent margins (≥0.3" from edges)
□ Left accent bar or visual motif on every content slide

INFOGRAPHICS
□ At least 4 different infographic types used
□ At least 1 chart (bar, line, or pie)
□ At least 1 stat callout slide (large numbers)
□ At least 1 timeline or process flow
□ At least 1 card/icon grid

SPEAKER NOTES
□ Every slide has notes
□ Each notes section ≥500 characters
□ Notes include opening phrase
□ Notes include key talking points
□ Notes include transition to next slide

FILE QUALITY
□ Run markitdown QA check
□ Convert to images and inspect visually
□ No overlapping elements
□ No text cutoff
□ Consistent bottom/top bars across content slides
```

---

## Phase 9 — Final Output Commands

```bash
# 1. Generate presentation
node presentation.js

# 2. Text QA
python -m markitdown output.pptx

# 3. Check for placeholders
python -m markitdown output.pptx | grep -iE "TODO|lorem|ipsum|\[insert|\bXXX\b"

# 4. Convert to images for visual QA
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
rm -f slide-*.jpg
pdftoppm -jpeg -r 150 output.pdf slide
ls -1 "$PWD"/slide-*.jpg

# 5. Copy to outputs
cp output.pptx /mnt/user-data/outputs/
```

---

## Quick Start Template

```javascript
// presentation.js — replace all [TOPIC], [TITLE], etc.

const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "[TITLE]";

// == DESIGN TOKENS (copy from Phase 6) ==
// ... C, F, makeShadow, addContentSlide, addStatBox

// == SLIDES ==
buildTitleSlide("[TITLE]", "[SUBTITLE]", "[PRESENTER]", "[DATE]");
buildAgendaSlide([...sections]);
// ... add all 15-19 slides
// ... each slide type from the patterns above

pres.writeFile({ fileName: "output.pptx" });
```

