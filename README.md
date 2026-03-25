# AI Presentation Redesign | Python & PPTX

Transforming AI-generated PowerPoint presentations into polished, brand-consistent slides using Python and a custom corporate design system.

## What This Project Does

Takes a typical AI-generated PPTX — default 4:3 format, system fonts, arbitrary colors, no brand identity — and rebuilds it as a fully branded 16:9 presentation assembled programmatically from reusable design assets and a structured style module.

## Before & After

| Before | After |
|--------|-------|
| Default 4:3 format | 16:9 widescreen |
| System fonts, no hierarchy | Inter typeface, structured layout |
| Arbitrary default colors | Corporate palette #457B9D / #B1D2DA |
| No brand identity | Logo, header, footer on every slide |
| Chart with acid colors | Chart redesigned and integrated |

## Files

- `pptx_style.py` — style module: constants, colors, fonts, layout functions
- `pptx_after.py` — main file, generates the redesigned presentation
- `pptx_before.py` — baseline AI-generated version

## Tools

- Python 3
- python-pptx
- Matplotlib
- Adobe Illustrator (brand assets)
- Inter typeface

## Portfolio

Full project with visuals: <a href="https://www.behance.net/gallery/246230397/AI-Presentation-Redesign-Python-PPTX" target="_blank">Behance</a>
