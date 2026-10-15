# `universities/` — one JSON file per school

`purdue.json` is the default. To support another university, copy it and change the values — no Python needed:

1. Copy `purdue.json` to `<yourschool>.json` in this folder.
2. Update the `school`, `handbook`, and rule values (page size, margins, font, heading names, caption pattern, word limits, contrast threshold, …).
3. Run it: `python3 purdue_format_check.py MyThesis.docx --university yourschool`
   (or `--config path/to/yourschool.json` from anywhere).
4. Send it as a pull request so students at your school can use it too.

Rules for a new file:

- Only override keys the school actually changes; anything absent keeps the built-in default. Unknown keys are rejected loudly so typos fail instead of silently doing nothing.
- `caption_pattern` is a regular expression with the same groups as Purdue's: `(Label) (part).(number).(space)`.
- `heading_numbering` maps levels 1–5 to `[number-format, level-text, indent-left-twips, indent-hanging-twips]` (twips = twentieths of a point; 1440 = 1 inch).
- Accessibility thresholds (`contrast_min` etc.) follow WCAG 2.1 AA and rarely need changing; leave them unless your school's guide says otherwise.
- Spacing and indentation measurements:
  - `body_line_spacing`: line spacing in twips (360 = 1.5 lines, Handbook §1.9).
  - `body_indent_twips`: first-line paragraph indent in twips (720 = 0.5 in, Handbook §1.10).
  - `figure_gap_above_twips`: minimum spacing above a figure in twips (480 = 24 pt, Handbook §1.13).
  - `table_gap_below_twips`: minimum spacing after a table in twips (480 = 24 pt, Handbook §1.13).
  - `table_caption_spacing`: `[before_twips, after_twips]` for table captions (480, 120 = 24 pt before, 6 pt after, Handbook §1.13).
  - `figure_caption_spacing`: `[before_twips, after_twips]` for figure captions (120, 480 = 6 pt before, 24 pt after, Handbook §1.13).
  - `caption_line_spacing`: caption line spacing in twips (240 = single, Handbook §1.13).
  - `reference_hanging_indent_twips`: hanging indent for references in twips (720 = 0.5 in, Handbook §1.21).
  - `reference_line_spacings`: allowed reference line spacing values in twips ([240, 360, 480] = 1.0, 1.5, 2.0 lines, Handbook §1.21).
