# Taglift

Taglift converts Word formatting into HTML tags. Drop in a .docx and get a new .docx where bold, italic, underlined, struck-through, subscript and superscript text is wrapped in tags such as `<b>`, `<i>` and `<u>`. Tables, lists and layout stay intact, and you can review every change before downloading. Your files are converted privately in your browser.

## Example

| In Word | Converted |
| --- | --- |
| in the case of each **ABR Loan** that is an <ins>Initial</ins> <ins>Term</ins> <ins>B Loan</ins>, … 2.25% *per annum* | `in the case of each <b>ABR Loan</b> that is an <u>Initial</u> <u>Term</u> <u>B Loan</u>, … 2.25% <i>per annum</i>` |

## Using it

Open `index.html` in any modern browser (Chrome, Edge, Firefox, or Safari 16.4+). There is nothing to install and no server is needed.

1. Drag one or more `.docx` files onto the page, or click **choose files**.
2. Adjust the settings if needed. Files already added update instantly.
3. Click **Review the changes** to compare each paragraph before and after.
4. Click **Download .docx**. The result is saved as `<name> (tagged).docx`.

### Settings

| Setting | Default | What it does |
| --- | --- | --- |
| Formatting to convert | All six | Bold, italic, underline, strike-through, subscript, superscript |
| Tag names | Standard | `<b> <i> <s>`, semantic `<strong> <em> <del>`, or legacy `<b> <i> <strike>` |
| Remove the original formatting | On | Converted text becomes plain so only the tags show the formatting |
| Include formatting from styles | On | Tags text that is bold/underlined because of its style (headings, hyperlinks) |
| Headers, footers and notes | On | Also converts headers, footers, footnotes and endnotes |
| Keep edge spaces outside tags | On | Produces `<b>Loan</b> ` rather than `<b>Loan </b>` |
| Escape existing `< > &` | Off | Writes them as `&lt; &gt; &amp;` so the output is valid HTML |

### How conversion works

- Adjacent runs with the same formatting are merged into one tag, and overlapping formatting nests correctly, e.g. `<b>bold <i>both</i></b><i> italic</i>`.
- Only text changes. Tables, lists, numbering, hyperlinks, fonts, sizes, spacing and images are left as they were.
- Deleted tracked changes and hidden field codes are skipped.
- Password-protected files and old `.doc` files are rejected with a message explaining how to fix them.

## Privacy

All processing happens in the browser using built-in compression APIs and the browser's XML parser. Documents are never uploaded, so any static host can serve the app without ever seeing a document.

## Hosting

`index.html` is a single self-contained file (fonts load from Google Fonts and fall back to system fonts offline).

**GitHub Pages:** in this repository go to **Settings → Pages**, set **Source** to *Deploy from a branch*, choose `main` and `/ (root)`, and save. The app will be live at `https://reshard-porter.github.io/Taglift/` within a minute or two.

Any other static host (Netlify, Cloudflare Pages, Vercel, an intranet or SharePoint) works the same way: upload `index.html`.

## Development

| File | Purpose |
| --- | --- |
| `src/engine.js` | Conversion engine: reads and writes the .docx zip, resolves styles, inserts tags |
| `src/template.html` | Page layout, styles and interface code |
| `src/sample.docx` | Sample document behind the “Convert the sample agreement” link |
| `build.py` | Combines the files above into `index.html` |

After editing anything in `src/`, rebuild with:

```sh
python3 build.py
```

The engine also runs in Node (18+) for testing, given a DOM implementation such as `@xmldom/xmldom`:

```js
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const Taglift = require('./src/engine.js');
const { blob, stats } = await Taglift.convert(arrayBuffer, { tagStyle: 'classic' }, { DOMParser, XMLSerializer });
```
