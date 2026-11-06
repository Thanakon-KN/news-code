# News Code

A browser-based DOCX to CMS HTML formatter that follows the `update-news` skill conventions. The interface guides the editor from adding source files, through conversion settings, to reviewing and copying the output.

## Use

Open the GitHub Pages tool at:

`https://thanakon-kn.github.io/news-code/`

Add one or more `.docx` files, choose whether to skip detected headlines, and select **สร้าง HTML**. Thai and English outputs are separated automatically. Each output starts with the standard figure placeholder. Review, then copy or download the HTML for the CMS.

## Preserved content

- Paragraph boundaries and hard line breaks
- Bold, italic, and underline formatting
- A rendered preview alongside the raw HTML code
- Basic numbered and bulleted lists
- External hyperlinks, with `target="_blank"`
- News text stored in a labeled `เนื้อข่าว` or `Body` table cell

The tool does not rewrite or summarize text. It reads DOCX files locally in the browser; the documents are not uploaded to a server. Legacy `.doc` files are not supported.

## Files

- `index.html` — interface and accessible controls
- `styles.css` — responsive layout
- `app.js` — DOCX ZIP/XML reader and HTML formatter

No build step or third-party runtime dependency is required. GitHub Pages serves the static files from this folder.
