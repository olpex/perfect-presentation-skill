# 🎓 perfect-presentation-skill

> An AI agent skill for generating complete, production-ready Python scripts that build **professional, high-quality educational PowerPoint presentations** using the [`python-pptx`](https://python-pptx.readthedocs.io/) library.

---

## ✨ What It Does

The `perfect-presentation` skill transforms a presentation topic into a **fully runnable Python script** that produces a polished `.pptx` file. It combines the expertise of an instructional designer, UI/UX specialist, and `python-pptx` master into a guided three-step dialogue.

**Key highlights:**
- 📋 **Guided configuration** — 3-step dialogue to define topic, audience, sources, slide count, and tone
- 🎨 **Premium design system** — High-contrast palette, academic typography, and structured slide layouts
- 📊 **Real data visualizations** — Actual `pptx.chart` objects and tables (no placeholders)
- 🗒️ **Deep speaker notes** — 700–2,000 words per slide with actionable steps and navigation paths
- 📝 **APA/URL citations** — Embedded per-slide in both the notes and a slide textbox
- 🏗️ **Clean architecture** — Class-based `EducationalPresentationBuilder` with type hints and docstrings

---

## 📁 Repository Structure

```
perfect-presentation-skill/
└── skills/
    └── perfect-presentation/
        └── SKILL.md        # Skill definition & full instructions for the AI agent
```

---

## 🚀 Usage

### Prerequisites

```bash
pip install python-pptx
```

### Workflow

1. **Invoke the skill** — Point your AI agent at `skills/perfect-presentation/SKILL.md`.
2. **Answer the 3-step dialogue:**
   - *Phase 1* — Provide the presentation topic.
   - *Phase 2* — Choose target audience, sources, slide count, style, and tone.
   - *Phase 3* — The agent generates the complete Python script.
3. **Run the generated script** — Execute it with Python to produce your `.pptx` file.

```bash
python generated_presentation.py
```

---

## 🛠️ Generated Script Capabilities

The agent produces a Python script with the following slide builder methods:

| Method | Description |
|---|---|
| `add_title_slide` | Title, subtitle, and author block |
| `add_agenda_slide` | Numbered agenda overview |
| `add_concept_slide` | Concept explanation with bullet points |
| `add_chart_slide` | Bar/line/pie charts using real `pptx.chart` data |
| `add_table_slide` | Structured comparison or data tables |
| `add_infographic_slide` | Icon-based visual layout |
| `add_engagement_slide` | Discussion questions or interactive prompts |
| `add_summary_slide` | Key takeaways and call-to-action |

---

## ⚙️ Configuration Options

| Option | Details |
|---|---|
| **Language** | All content (titles, bullets, charts, notes) in **English** |
| **Slide Count** | Default ≥ 15 slides; fully configurable |
| **Styles** | Scientific · Journalistic · Business · Conversational · Artistic · Technical · Educational · and more |
| **Tones** | Official · Friendly · Inspiring · Authoritative · Serious · Passionate · and more |
| **Sources** | Web research · Uploaded documents · Transcripts · Database search · Custom sources |

Styles and tones can be **combined or customized**.

---

## 💡 Example Prompts

```
Topic: Cybersecurity for Seniors
  → Agent starts Phase 1 dialogue.

Topic chosen. Configure: Academic style, 20 slides, Source: Web research
  → Agent proceeds to Phase 3 and generates the full Python script.
```

---

## 📄 License

This project is open source. See the repository for license details.
