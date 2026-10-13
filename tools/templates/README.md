# Default templates

`build_templates.py` generates the default DOCX templates in [`web/src/main/resources/documents/`](../../web/src/main/resources/documents/). The design tokens, styles, page setup, cover, header and footer, and template tags all live in the script. To change the look, edit the script, not the `.docx` files, and regenerate the templates:

```bash
python tools/templates/build_templates.py
```

The script needs only the Python 3 standard library and runs from any folder. Then run `mvn test`; one test checks that the default templates pass template validation without issues. Commit the regenerated `.docx` files with the script change.

## Fonts

The templates embed subsets of Noto Sans (Regular, Bold, Italic, Bold Italic) and Noto Sans SemiBold (headings). The fonts are licensed under the [SIL Open Font License](fonts/OFL.txt), which allows embedding. Because the fonts are embedded, Word shows the document exactly as the PDF conversion does. Gotenberg also ships Noto (`fonts-noto-core`).

The subsets cover Latin (including Latin-1 and Latin Extended-A), Greek and basic Cyrillic: every language the OpenCDMP UI ships. The five files together add about 160 KB to a document. They were made from the [notofonts](https://github.com/notofonts/notofonts.github.io/tree/main/fonts/NotoSans/hinted/ttf) static TTFs with `pyftsubset` from [fontTools](https://github.com/fonttools/fonttools):

```bash
pyftsubset NotoSans-<Style>.ttf --output-file=fonts/NotoSans-<Style>.ttf \
  --unicodes="U+0020-007E,U+00A0-017F,U+0192,U+0218-021B,U+02C6-02DD,U+0300-0308,U+0370-03FF,U+0400-045F,U+0490-0491,U+2010-2027,U+2030-203A,U+2044,U+20AC,U+2116,U+2122,U+2190-2193,U+2212,U+2264,U+2265,U+25CF,U+25E6" \
  --layout-features='kern,liga,calt,ccmp,locl,mark,mkmk' --no-hinting --desubroutinize --name-IDs='*' --name-legacy --notdef-outline
```

Word and LibreOffice render characters outside the subsets, such as CJK, with a fallback font.
