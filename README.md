# Thesis ER — will your thesis pass review?

*Resuscitate your formatting before the deadline.* A deterministic format checker (and auto-fixer) for thesis Word files — built for Purdue's strict thesis format review, designed so any university's rules can be plugged in.

> **Unofficial community project.** This tool is not sponsored by, endorsed by, or affiliated with Purdue University. It encodes one student's reading of the thesis handbook and template; always confirm against the official handbook before you submit.

If it saves you a rejected submission, please ⭐ **star this repo** — it helps other students find it.

## Privacy: there is no database

This project cannot afford a database, and that is a feature: uploaded files are analyzed and deleted the moment the report is done. Nothing is stored, nothing is tracked, nothing trains any model. If your thesis is under confidentiality or embargo and even that is too much trust, don't upload it — clone the repo and run the checker on your own machine. It works fully offline.

## Why this exists

Purdue enforces a strict thesis format policy, and a missed deadline can mean the difference between graduating on time and not. Word's built-in checks don't know the handbook. This tool reads your `.docx` directly and tells you, rule by rule, what fails and what the handbook wants instead — including accessibility (ADA/WCAG 2.1 AA) findings like missing alt text, low-contrast chart colors, floating images that break reading order, and layout tables.

## Quick start

Requirements: Python 3.10+, and `pip install -r requirements.txt` (`lxml`, `Pillow`, `numpy`).

```bash
# Check a document — every finding carries a stable rule ID and a location
python3 purdue_format_check.py MyThesis.docx

# Emit machine-readable JSON report for CI and editor integrations
python3 purdue_format_check.py MyThesis.docx --json

# Prove every rule fires (mutates a copy of the file, one rule at a time)
python3 purdue_format_check.py MyThesis.docx --selftest

# Auto-fix what can be fixed deterministically, then re-check to convergence
python3 purdue_format_fix.py MyThesis.docx MyThesis_fixed.docx
```

The checker reads the OOXML directly — no Word install, no LLM, no network. The same input always produces the same output.

## Rule families

| Prefix | Source |
|--------|--------|
| `H1.*`  | Thesis handbook sections (e.g. `H1.8` = fonts, `H1.9`–`H1.11` = body text, `H1.13` = figures/tables) |
| `TPL`   | The official Word template's styles, numbering, and front matter |
| `A11Y`  | Accessibility: headings, alt text, 3:1 chart contrast, scanned pages, layout tables, text boxes, reading order |
| `REV`   | Submission hygiene: tracked changes and comments must be resolved |
| `SK`    | Extra strictness checks the handbook implies but never states outright (caption alignment, mention-before-appear) |

The full catalog — every check in plain English, generated from the checker source — is [RULES.md](RULES.md).

Run `--selftest` to see all 89 rules demonstrated.

## For other universities

School-specific values (page size, font, heading names, caption pattern, word limits, …) live in `universities/*.json`, with Purdue as the default. Support your school by copying `purdue.json` and sending a PR — see [universities/README.md](universities/README.md). Run any school with `--university <name>` or `--config <path>`.

## Deploy the browser API on Render

The repository includes `render.yaml`; no command or secret needs to be typed
by hand:

1. In Render, choose **New → Blueprint** and connect this GitHub repository.
2. Select `mogiiee/thesis-er` and apply the `thesis-er-api` blueprint.
3. Wait for `/api/universities` to report healthy, then copy the public
   `https://…onrender.com` service URL.
4. Put that URL in `site/config.js` as `BACKEND_URL` and redeploy GitHub Pages.

The API permits browser requests only from `https://mogiiee.github.io` by
default. Uploaded files use a temporary file and are deleted in the same
request; there is no database.

## Contributing

Contributions are welcome — Hacktoberfest-friendly. Please read [CONTRIBUTING.md](CONTRIBUTING.md) first: small, rule-scoped PRs with a self-test mutation for every new rule.

## License

MIT — see [LICENSE](LICENSE).
