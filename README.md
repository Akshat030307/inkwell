# Inkwell

A word processor modelled on Microsoft Word 2013, running in the browser. It reads and writes real `.docx` files, so
documents open in Word, LibreOffice and Google Docs.

## Run it

Open `index.html` in Chrome, Edge or Firefox. There's no build step and no server needed:

```bash
xdg-open index.html
```

Documents autosave in the browser (IndexedDB). **FILE → Save** (Ctrl+S) downloads a `.docx`.

## Features

| Tab | What it does |
| --- | --- |
| **File** | Info; New from templates (blank, resume, letter, report, meeting notes, essay); Open recent documents or a `.docx` / `.html` / `.txt` file; Save; Export to `.docx`, PDF, HTML or text; Print; Options (user name, AutoCorrect) |
| **Home** | Paste, cut, copy, Format Painter; font, size, grow/shrink, Change Case, bold, italic, underline, strikethrough, sub/superscript, font colour, highlight, clear formatting; bullet, numbering and multilevel list libraries, indent, alignment, line spacing, shading, borders, Paragraph dialog; Styles gallery with Modify Style, Create a Style and custom styles; find, replace, select |
| **Insert** | Page break, tables, pictures, hyperlinks, comments, header, footer, page numbers, date, horizontal line, symbols |
| **Page Layout** | Margins, orientation, size, columns (1–3), page and column breaks, paragraph indent and spacing fields |
| **References** | Automatic table of contents with page numbers (built from Heading 1–3), captions |
| **Review** | Spell check, word count; comments (new, delete, previous/next, resolve); Track Changes with accept/reject, previous/next and show/hide markup |
| **View** | Read Mode, Print Layout, ruler with draggable indent markers, Navigation Pane (headings outline and search results), zoom, dark mode |
| **Table** / **Picture** | Appear in context: insert/delete rows and columns; text wrapping, alignment, exact size, rotate/flip, crop, alt text, captions |

The page view is paginated like Word's Print Layout: each page is its own sheet with its header and footer, and a
paragraph that won't fit moves to the next page. AutoCorrect converts smart quotes, dashes, ellipses, arrows,
(c)/(r)/(tm), starts lists from `* ` or `1. `, and turns web addresses into links.

Most Word keyboard shortcuts work, including Ctrl+B/I/U, Ctrl+Z/Y, Ctrl+S/O/P, Ctrl+F/H, Ctrl+K, Ctrl+E/L/R/J,
Ctrl+Alt+1/2/3, Ctrl+Shift+N/L, Ctrl+Enter, Ctrl+]/[, Ctrl+= and Ctrl+Space.

## Code layout

```
index.html               ribbon, panes, menus, dialogs
css/app.css              application chrome, light and dark
js/docx.js               Office Open XML reader and writer
js/editor.js             core: selection, undo history, commands, find/replace, tables, pictures, file I/O
js/features/             one module per feature area; each registers commands and hooks on window.Ink
  pages.js               pagination, page sheets, headers/footers, print
  styles.js              style definitions → document CSS, gallery, styles.xml
  paragraph.js           Paragraph dialog, ruler, borders/shading, Format Painter, Change Case, clipboard
  lists.js  autocorrect.js  navigation.js  sections.js  pictures.js  review.js  views.js  pdf.js  library.js
js/main.js               starts the app
vendor/                  JSZip 3.10.1, html2canvas 1.4.1, jsPDF 2.5.1 (MIT)
```

## How the .docx format works

Inkwell's `.docx` support was built by reverse engineering files that Word writes (REA's archive inspector plus
reading their XML), following the published ECMA-376 standard:

- **Package:** a `.docx` is a ZIP package of XML parts listed in `[Content_Types].xml` and linked by `_rels/*.rels`. The parts are `word/document.xml` (content), `styles.xml`, `numbering.xml`, `settings.xml`, `header1.xml`, `footer1.xml`, `comments.xml` and `media/`.
- **Body:** a flat sequence of paragraphs (`<w:p>`) and tables (`<w:tbl>`). Paragraphs hold runs (`<w:r>`), stretches of text with one set of formatting (`<w:rPr>`). The writer records only formatting that differs from the paragraph's style, as Word does.
- **Lists:** paragraphs that point into `numbering.xml` (`<w:numPr>`). The reader rebuilds nested lists from their levels.
- **Sections:** a section ends at a paragraph carrying `<w:sectPr>`, which is how multi-column regions are stored.
- **Revisions and comments:** Track Changes become `<w:ins>` / `<w:del>`. Comments are a range in the text (`commentRangeStart` / `commentRangeEnd`) plus an entry in `comments.xml`.
- **Header and footer fields:** page numbers are `PAGE` and `NUMPAGES` fields.
- **Units:**

  | Measure | Unit |
  | --- | --- |
  | Page size and margins | twips (1/1440 inch) |
  | Font sizes | half-points |
  | Pictures | EMUs (914,400 per inch) |

- **Element order:** child elements must follow the schema's order, or Word reports the file as corrupt.

## Limitations

- **Pagination:** works per paragraph. A single paragraph or table taller than a page runs across the page gap instead of splitting.
- **Comments and Track Changes:** comments are single-level, with no replies. Track Changes records typing, deleting and pasting, but not formatting changes or paragraph joins.
- **PDF export:** one-click PDF renders pages as images, so the text isn't selectable. Print → Save as PDF keeps selectable text.
- **Not supported:** footnotes, text boxes, shapes, mail merge, equations, and per-section orientation.
- **Fonts:** Calibri isn't on most Linux systems. The metric-compatible Carlito font (`ttf-carlito`) gives Word-accurate line breaks.

## Deploying

Inkwell is static files. Serve the repository root with any web server. Nothing is stored server-side, because
documents live in each visitor's browser.
