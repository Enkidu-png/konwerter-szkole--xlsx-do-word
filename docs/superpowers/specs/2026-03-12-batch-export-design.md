# Batch Export — Design Spec
Date: 2026-03-12

## Overview
Add batch export capability to the Konwerter Szkoleń app: users can select multiple training products from the list and download them all as separate `.docx` files packaged in a single ZIP archive.

## UI / Interaction

- Each product row in the left-side list gains a **checkbox** (always visible, not just on hover).
- Clicking a product name still opens the preview on the right; clicking the checkbox toggles selection without changing the preview.
- The list header updates dynamically: `Produkty (120) · Zaznaczono: 5`
- When ≥1 product is selected, the header shows:
  - **"Eksportuj zaznaczone (N)"** button (indigo, primary style)
  - **"Odznacz wszystkie"** text link
- The right-side preview panel and single-product export are unchanged.

## State

```ts
const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
```

Toggling a checkbox calls `setSelectedIds` with an updated Set copy.

## Refactor: `generateDocxBlob(product)`

Extract the `.docx` generation logic from `downloadWord()` into a standalone async function `generateDocxBlob(product: TrainingProduct): Promise<Blob>`. Both the single-file export and the batch export use this function.

## Batch Export: `downloadBatch()`

```
for each selectedId:
  product = products.find(p => p.id === selectedId)
  blob = await generateDocxBlob(product)
  add blob to PizZip as `${product.name}.docx`
zip.generate({ type: 'blob' })
saveAs(zip, `eksport_YYYY-MM-DD.zip`)
```

- Uses **PizZip** (already a project dependency) for ZIP creation — no new packages needed.
- ZIP filename: `eksport_2026-03-12.zip` (current date).
- During generation the button shows a spinner and progress: `Generowanie (3/5)…`
- On error, an `alert()` is shown (consistent with existing error handling).

## Files Changed

- `src/App.tsx` only — no new files or dependencies.

## Out of Scope

- Shift+click range selection
- Export-all button (no granular selection)
- Progress bar UI beyond the button label counter
