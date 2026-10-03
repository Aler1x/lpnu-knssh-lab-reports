# lpnu-lab-reports

English | [Українська](README.uk.md)

An agent skill that writes and formats lab reports for the Department of Artificial Intelligence Systems, Lviv Polytechnic National University (LPNU).

- Reports are in Ukrainian.
- LaTeX is the source. The skill exports PDF and an editable Word (DOCX) file from it.
- Formatting follows DSTU 3008:2015. Department and teacher rules apply where they differ.
- Not for bachelor, master or course theses.

## Install

```sh
npx skills add Aler1x/lpnu-knssh-lab-reports
```

Then ask your agent, for example: "Use lpnu-lab-reports to make a report for lab 3 from my code and results."

## Requirements

| Output | You need |
|--------|----------|
| PDF | XeLaTeX (`latexmk -xelatex`) and the Times New Roman font |
| DOCX | [Pandoc](https://pandoc.org) and Python with `python-docx` |

## Contents

| Path | What it is |
|------|------------|
| `SKILL.md` | Instructions for the agent |
| `references/format.md` | Formatting rules (DSTU 3008:2015) |
| `references/export.md` | PDF and DOCX export and checks |
| `assets/report.tex` | Report template with the title page |
| `assets/university-logo.png` | University logo for the title page |
| `scripts/export_docx.py` | LaTeX to DOCX export with the template formatting |
