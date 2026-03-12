# Batch Export — Design Spec
Date: 2026-03-12

## Overview
Add batch export capability to the Konwerter Szkoleń app: users can select multiple training products from the list and download them all as separate `.docx` files packaged in a single ZIP archive.

## UI / Interaction

- Each product row in the left-side list gains a **checkbox** (always visible).
- Clicking a product name still opens the preview on the right; clicking the checkbox toggles selection without changing the preview.
- The list header shows filtered count + selected count: `Produkty (10) · Zaznaczono: 5` — where `(10)` is `filteredProducts.length` (consistent with current behaviour).
- When ≥1 product is selected, the header shows:
  - **"Eksportuj zaznaczone (N)"** button (indigo, primary style)
  - **"Odznacz wszystkie"** text link
- No "Zaznacz wszystkie" affordance — intentionally out of scope.
- Products that are selected but hidden by the current search filter are still exported (selectedIds is not filtered by filteredProducts).
- The right-side preview panel and single-product export are unchanged.

## State

```ts
const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
const [batchProgress, setBatchProgress] = useState<{ current: number; total: number } | null>(null);
```

Toggling a checkbox calls `setSelectedIds` with an updated Set copy.
`batchProgress` is `null` when idle; set during generation to drive button label.

## Refactor: `generateDocxBlob(product, templateBuffer)`

Extract `.docx` generation logic from `downloadWord()` into:

```ts
async function generateDocxBlob(product: TrainingProduct, templateBuffer: ArrayBuffer): Promise<Blob>
```

- Receives the pre-fetched template `ArrayBuffer` as a parameter — the template is **fetched once** before the batch loop, not once per product.
- Both single-file export and batch export use this function.
- `downloadWord()` becomes: fetch template → call `generateDocxBlob` → `saveAs`.

## Batch Export: `downloadBatch()`

```
templateBuffer = await fetch('/SZABLON2.docx').arrayBuffer()

zip = new PizZip()
for each selectedId (i = 0..N-1):
  setBatchProgress({ current: i + 1, total: N })
  product = products.find(p => p.id === selectedId)
  blob = await generateDocxBlob(product, templateBuffer)
  arrayBuffer = await blob.arrayBuffer()           // Blob → ArrayBuffer for PizZip
  safeName = product.name.replace(/[^a-z0-9]/gi, '_')
  // Deduplicate: if safeName already in zip, append product.id
  entryName = zip.files[`${safeName}.docx`] ? `${safeName}_${product.id}.docx` : `${safeName}.docx`
  zip.file(entryName, arrayBuffer)

zipBlob = zip.generate({ type: 'blob' })
saveAs(zipBlob, `eksport_${today}.zip`)
setBatchProgress(null)
setSelectedIds(new Set())   // clear selection after successful export
```

- Button and "Odznacz wszystkie" link are **disabled** while `batchProgress !== null`.
- Button label during generation: `Generowanie (3/5)…`
- On error: `alert()` (consistent with existing error handling), `setBatchProgress(null)` to restore button.
- After successful export: selection is cleared.

## Filename Rules

- Sanitise with `replace(/[^a-z0-9]/gi, '_')` — same as single export.
- Duplicate names: append `_${product.id}` suffix.
- ZIP filename: `eksport_YYYY-MM-DD.zip` using current date.

## Files Changed

- `src/App.tsx` only — no new files or dependencies.

## Out of Scope

- Shift+click range selection
- "Zaznacz wszystkie" button
- Export-all without selection
- Progress bar UI beyond the button label counter
