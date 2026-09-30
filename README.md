# FILO

FILO is a local-first file refinement toolkit for documents, images, and everyday files.

**Drop a file. Make it better.**

FILO brings practical file tools into one calm workspace:

- Convert documents and files into useful output formats
- Render DOC/DOCX files into multi-page PDFs with embedded images and page structure
- Convert JPEG and PNG images with quality controls
- Compress or upscale images toward a target size
- Explore 50 curated fonts and preview user copy in each typeface
- Remove simple solid backgrounds from portrait and passport-style photos
- Add a solid color background and export as PNG or JPEG
- Work locally in the browser with no default file uploads to a server

## Run locally

```bash
npm install
npm run dev
```

Then open `http://localhost:5173`.

## Build

```bash
npm run build
```

## Tech

React, TypeScript, Vite, jsPDF, docx-preview, html2canvas, and browser Canvas APIs.

## Product notes

FILO is designed as a focused workspace rather than a directory of disconnected tools. The interface keeps a persistent preview, clear file metadata, responsive controls, light/dark mode, and small feedback states close to the action.

Complex Word documents can vary across Word versions and embedded fonts. FILO renders DOC/DOCX files in-browser using `docx-preview`; exact parity still depends on the fonts and document features available in the browser.
