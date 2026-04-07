# Sizer for Photoshop and Illustrator

`Sizer` is a pair of production-focused `ExtendScript` tools for Adobe `Photoshop` and `Illustrator`.

They take a pasted order email, match each line item to a file in a chosen folder, resize/export everything at a fixed `300 DPI/PPI`, and generate review reports that make it easy to catch size issues, missing files, pricing differences, and naming/sorting mistakes before anything reaches print.

This project is built for real DTF-style batch production work, not as a demo script.

## What This Project Does

You paste the raw order email into the script UI.

The script then:

- parses each ordered item
- extracts `Width`, `Height`, `Quantity`, uploaded filename, optional customer `Message`, and price
- detects the print type from the product label
- finds the matching local artwork file using tolerant filename matching
- resizes the artwork according to the selected mode
- exports a normalized `PNG` output at `300 DPI/PPI`
- optionally sorts by print type
- generates review artifacts for production and customer communication

There are two host-specific scripts:

- `Sizer_PS_v1.5.jsx` for `Adobe Photoshop`
- `Sizer_AI_v1.5.jsx` for `Adobe Illustrator`

Both scripts follow the same workflow and produce the same family of reports.

## Why It Exists

Order emails are messy.

Downloaded assets are messy.

Filename mismatches happen.

Print types get mixed.

Artwork dimensions do not always match what the customer paid for.

This tool exists to turn that mess into a repeatable batch process with visibility.

## Core Features

- Fixed output at `300 DPI/PPI`
- Three resize modes:
  - `Respect Width`
  - `Respect Height`
  - `Stretch`
- Three naming modes:
  - `Filename___Qty`
  - `Qty___Filename`
  - `Filename`
- Three print-type output modes:
  - `Folder`
  - `Prefix`
  - `None`
- Strict item parsing from pasted order emails
- Print type detection from each item's own product label
- Tolerant file matching:
  - exact
  - normalized
  - canonical
  - ultra-loose
  - suggestion fallback
- `Customer Proof` generation
- `Pricing Audit` generation
- `_Export_LOG.txt` fallback logging for reliability
- Sortable `Export Report`
- Sticky table header inside the report
- Per-row review checkbox to mute highlight without changing the actual status

## Supported Print Type Detection

The current built-in print type rules recognize:

- `UV`
- `COOL`
- `HEAT`
- `Glitter`
- `Dyeblocker`

Detection is intentionally strict and tied to the product label for the current order item, which prevents cross-item bleed such as a `HEAT` item being misclassified as `UV` because the next line item is `UV DTF`.

## Inputs

Each run expects two things:

1. A folder containing the customer artwork files
2. The raw order email pasted into the script dialog

The parser is designed around real pasted order text, for example:

```text
DTF Heat Transfers
Width:
3

Height:
3.5

Image file upload:
Zach-Crest-1.png

1    $0.30
```

The scripts extract:

- product label
- width
- height
- uploaded filename
- quantity
- price
- currency
- customer message or note when present

They also extract order-level financial values such as:

- `Subtotal`
- `Shipping`
- `HST`
- `Total`

## File Matching Strategy

Matching is one of the most important parts of this project.

The scripts do not rely only on raw exact filenames.

They progressively try:

1. exact filename match
2. normalized match
3. canonical match
4. ultra-loose match
5. best suggestion when no safe match is found

This allows the scripts to survive common differences such as:

- case changes
- underscores vs hyphens
- multiple spaces
- Unicode dash variants
- mild filename cleanup differences

The search is limited to the selected folder.

It does not recurse through subfolders for source file discovery.

## Outputs

Each run writes into an `Export` folder and can generate the following files:

- `_Export_REPORT.html`
- `_Export_LOG.txt`
- `_Customer_Proof.html`
- `_Pricing_Audit.html`

Depending on options, it also exports the processed artwork files themselves.

## Report Types

### `_Export_REPORT.html`

This is the primary production QA view.

It includes:

- thumbnail preview
- filename
- quantity
- print type
- price
- note
- match quality
- ordered size
- output size
- delta
- status

The table supports:

- sticky header
- client-side sorting
- temporary review muting with a checkbox

Status coloring is directional, not generic:

- output larger than ordered is treated as more dangerous
- output smaller than ordered is treated as a lower-severity review case
- mixed distortions are highlighted separately

### `_Export_LOG.txt`

This is the reliability fallback.

It exists because writing rich `HTML` reports can fail on some environments, especially cloud-backed or heavily managed file locations.

The log is append-oriented and intended to preserve a usable audit trail even if the richer report write fails.

### `_Customer_Proof.html`

This is the customer-facing proof sheet.

It is intentionally cleaner than the QC report and focuses on:

- file preview
- clear dimensions
- visually readable presentation

It is suitable for quick review, browser display, and possible conversion to `PDF` outside the script.

### `_Pricing_Audit.html`

This is the internal pricing adjustment tool.

It compares ordered dimensions with measured or manually adjusted dimensions and recalculates what the item price should be.

Rows stay in the report even if the source file is missing or problematic.

That makes it possible to continue the pricing review manually.

The report supports manual entry for:

- adjusted width
- adjusted height
- adjusted quantity

Rows are sorted by the highest price difference first so the most expensive problems rise to the top.

## Pricing Logic

Pricing is recalculated by proportional area.

Formula:

```text
adjusted price = current price * (new width * new height) / (ordered width * ordered height)
```

This is applied per item.

The `Pricing Audit` then computes:

- total price adjustment
- `HST` on the adjustment only
- final amount due

It is designed to answer a very practical question:

How much more should the customer pay because the real produced size is larger than what was ordered?

## Photoshop vs Illustrator

The two scripts are intentionally aligned, but they are still host-specific.

### Photoshop

- Opens files in `Photoshop`
- Resizes the document directly
- Can optionally run the `WeMust / WeMust` action
- Saves the exported output and reports through the `Photoshop` host workflow

### Illustrator

- Opens files in `Illustrator`
- Fits and exports through the `Illustrator` document model
- Restores user interaction state safely
- Detects locked content and skips those files instead of silently processing them

For locked-content files, `Illustrator` reports the issue instead of pretending everything is fine.

## Running the Scripts

### Photoshop

1. Open `Adobe Photoshop`
2. Run `File > Scripts > Browse...`
3. Choose `Sizer_PS_v1.5.jsx`
4. Select the artwork folder
5. Paste the order email
6. Choose naming, resize, and print-type options
7. Enable optional reports if needed
8. Run

### Illustrator

1. Open `Adobe Illustrator`
2. Run `File > Scripts > Other Script...`
3. Choose `Sizer_AI_v1.5.jsx`
4. Select the artwork folder
5. Paste the order email
6. Choose naming, resize, and print-type options
7. Enable optional reports if needed
8. Run

## UI Options

Both scripts expose the same main controls:

- `Filename Format`
- `Resize Mode`
- `Sort by Print Type`
- `Run Action: WeMust / WeMust`
- `Generate Customer Proof HTML`
- `Generate Pricing Audit HTML`
- `Files Folder`
- `Paste Order Email`

Output resolution is fixed:

- `300 DPI` in `Photoshop`
- `300 PPI` in `Illustrator`

## Reliability Notes

This repo contains multiple safeguards added for real production use:

- strict print-type parsing to prevent cross-item contamination
- fallback text logging
- safer report writing with fallback behavior
- explicit missing/problem tracking
- locked-content detection in `Illustrator`

If the richer reports fail to write in a managed environment, `_Export_LOG.txt` is the first file to check.

## Current Constraints

This project is intentionally pragmatic.

It does not try to be everything.

Current limitations include:

- no direct `PDF` proof export inside the script
- no direct `Windows Explorer` file-selection integration from the report
- no clipboard automation
- no recursive source-file scanning
- no formal automated test suite

## Recommended Workflow

1. Download customer assets into a clean folder
2. Paste the raw order email into the script
3. Run export
4. Review `_Export_REPORT.html`
5. Open `_Pricing_Audit.html` if dimensions or billing need adjustment
6. Open `_Customer_Proof.html` if you need a clean customer-facing preview
7. Use `_Export_LOG.txt` if anything seems missing or if report writing partially fails

## Repo Contents

```text
Sizer_js_for_Photoshop_and_illustrator/
├─ Sizer_PS_v1.5.jsx
├─ Sizer_AI_v1.5.jsx
├─ README.md
└─ AGENTS.md
```

## Who This Is For

This repo is for print shops, production operators, and internal workflow builders who need:

- speed
- repeatability
- visibility
- less manual checking
- fewer expensive print mistakes

## Status

Current tracked script version in this repo:

- `Sizer_PS_v1.5.jsx`
- `Sizer_AI_v1.5.jsx`

The project is actively shaped around real shop-floor feedback.

If a report feels noisy, a matching rule feels risky, or a pricing edge case appears in production, the scripts are designed to evolve with those realities.
