# owui-pptx-kit

Restyle a PowerPoint deck in the corporate style of another deck by **cloning the reference deck's slides and
layouts** and filling their slots with the source content. Built to be driven by an LLM agent (Open WebUI +
Open Terminal), but works as a plain CLI.

```
python3 pptx_kit.py doctor
python3 pptx_kit.py profile REF.pptx --out work/ref        # style passport, template catalogue, annotated renders
python3 pptx_kit.py outline SRC.pptx --out work/src        # content outline, images, renders
python3 pptx_kit.py build   REF.pptx work/plan.json --out OUT.pptx
python3 pptx_kit.py qa      OUT.pptx --plan work/plan.json --ref-dir work/ref --out work/qa
```

The only creative step is `plan.json` (which template slide/layout to use for each source slide and what text goes
into which slot) — written by the model. Everything mechanical is deterministic:

- slide cloning with relationship remapping (images, charts, media), original slides removed;
- text replacement that keeps run/paragraph formatting of the template (per-paragraph styles, spacers, bullets);
- auto-removal of unfilled template text, picture-bullet markers trimmed to the number of items;
- `respace` (stretch remaining cards over the row), `add` table/image into a slot area, image replacement with crop;
- layouts with placeholders (inherited font sizes/positions resolved through layout → master);
- QA: LibreOffice render, side-by-side template/result pairs, leftover-template-text check, text-overflow check
  from rendered word positions (`pdftotext -bbox`), size/aspect check, missing-font warning.

Requirements: Python 3.10+, `python-pptx`, `lxml`, `Pillow`, LibreOffice (`soffice`), poppler (`pdftoppm`, `pdftotext`).

`skills/pptx-restyle.md` — the Open WebUI skill that drives the kit.

License: MIT
