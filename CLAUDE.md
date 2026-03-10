# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm install       # Install dependencies
npm run dev       # Start dev server at http://localhost:3000
npm run build     # Production build
npm run lint      # Type-check with tsc --noEmit
npm run preview   # Preview production build
```

## Environment

Copy `.env.example` to `.env.local` and set `GEMINI_API_KEY` (required for Gemini API calls). The `GEMINI_API_KEY` is exposed to the frontend via Vite's `define` in `vite.config.ts`.

## Architecture

This is a **single-page React app** (no backend) built with Vite + TypeScript + Tailwind CSS v4.

**The entire application logic lives in `src/App.tsx`** — there are no other components or modules. The app:

1. Accepts an `.xlsx`/`.xls` file via drag-and-drop or file picker
2. Parses it with the `xlsx` library using raw row arrays (not keyed objects) — column mapping is index-based (A=0, B=1, …)
3. Displays a searchable list of `TrainingProduct` items on the left
4. Shows a preview of the selected product's HTML content on the right
5. Exports the selected product either as a `.doc` file (HTML-wrapped with Word namespace for compatibility) via `file-saver`, or copies it as `text/html` to the clipboard

**Column mapping logic** (in `handleFileUpload`):
- Columns A (0) and B (1) are always skipped
- Title/product name: searched by header keyword (`produkt`, `nazwa`), fallback col C (2)
- Identifier: column E (4)
- Code: column D (3)
- Description (`opis`): searched by header keyword, fallback col F (5)
- Scope (`zakres`): searched by header keyword, fallback col G (6)
- Extra data: columns X–AC (indices 23–28), included only if non-empty

**Word export** generates an HTML blob with `application/msword` MIME type and a `.doc` extension — Word opens this as a legacy HTML document, not a true `.docx`.

**Stack:** React 19, Vite 6, Tailwind CSS v4 (via `@tailwindcss/vite` plugin — no `tailwind.config.js`), `motion/react` for animations, `lucide-react` for icons.
