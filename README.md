# CRM XML Processor

Tool to downgrade CRM Dynamics solutions exported from **v9** so they import into **v8** — by stripping incompatible elements/attributes and injecting `ObjectTypeCode` values that v8 requires.

Accepts either the full **solution `.zip`** (recommended) or a raw `customizations.xml`.

## What's new in v2.0

- 📦 **ZIP-native** — point it at the exported `.zip`; it edits only the embedded `customizations.xml` + `solution.xml` in place, leaves every other entry byte-identical, and creates a timestamped backup of the ZIP.
- 🧭 **Two-tier ObjectTypeCode resolution** — target (v8) CSV wins, source (v9) CSV is the seed for new entities, `##` if neither knows the entity.
- 🔧 **`solution.xml` version rewrite** — `version="9.*"` and `SolutionPackageVersion="9.*"` on `<ImportExportXml>` are rewritten to the configured v8 target (defaults `8.2.0009.0019` / `8.2`). The solution's own `<Version>` element is left alone.
- 📊 **Rich CLI summary** — per-tag removal counts, target-vs-source-vs-`##` breakdown, version diff, size delta, duration.
- 🔁 **Backward-compatible** — raw-XML mode still works, and the legacy `entityTypeCodesFile` config key is still accepted (treated as the target CSV).

## Quick Start

```bash
# Install
npm install

# Run on the exported solution zip (in-place update + auto backup)
node crm-xml-processor.js config.json data/MySolution_1_0_0_0.zip

# Or run on a raw customizations.xml (legacy mode)
node crm-xml-processor.js config.json data/customizations.xml
```

The input type is auto-detected from the file extension (`.zip` vs `.xml`).

## What it does

- ✓ Auto-detects `.zip` vs `.xml` input
- ✓ Auto backup with timestamp before any change (`<file>.backup.<timestamp>`)
- ✓ For ZIPs: opens the archive, edits **only** the embedded `customizations.xml`, then writes the same zip back — other entries are left byte-identical (most efficient path, no re-extraction)
- ✓ Removes incompatible XML elements (configurable list)
- ✓ Removes specified attributes from given tags
- ✓ Injects `<ObjectTypeCode>` for every `<Entity>` — uses the CSV mapping when available, falls back to `##` placeholder otherwise
- ✓ Rewrites `solution.xml` root-element version attributes (`version=` and `SolutionPackageVersion=`) from any `9.x` value to the configured v8 target
- ✓ Prints a clean CLI summary at the end (counts per element/attribute, size delta, version changes, missing entities, duration)

> **Note:** Any entity not found in `data/entity-types.csv` gets a `##` placeholder. Search for `##` in the resulting XML and replace with the real ObjectTypeCode for your target v8 environment before importing.

## Files

| File | Purpose |
|---|---|
| `crm-xml-processor.js` | Main tool |
| `config.json` | Modification rules (elements/attrs to strip, etc.) |
| `data/entity-types-target.csv` | `EntityName,TypeCode` mapping from the **target v8 org** (authoritative) |
| `data/entity-types-source.csv` | `EntityName,TypeCode` mapping from the **source v9 org** (fallback for new entities) |

## ObjectTypeCode resolution — two-tier lookup

For every `<Entity>` in the XML the tool decides which `<ObjectTypeCode>` to inject:

1. **Target CSV first.** If the entity already exists in the v8 org, that code is authoritative.
2. **Source CSV as fallback.** For a *brand-new* custom entity that doesn't exist in v8 yet, seed the value with the source-org code so the XML stays valid. v8 will reassign its own code when the solution imports; the seed is only a hint.
3. **`##` placeholder** if the entity is in neither CSV — you must hand-edit the XML before importing.

Generate each CSV by running this SQL against the corresponding organization DB (e.g. `YourOrg_MSCRM`):

```sql
SELECT Name, ObjectTypeCode
FROM MetadataSchema.Entity
WHERE ObjectTypeCode > 0
ORDER BY Name;
```

Or pipe straight to file with `sqlcmd`:

```bash
sqlcmd -S <SqlServer> -d <Org>_MSCRM -E -h -1 -s "," -W ^
  -Q "SET NOCOUNT ON; SELECT Name, ObjectTypeCode FROM MetadataSchema.Entity WHERE ObjectTypeCode > 0 ORDER BY Name" ^
  -o data\entity-types-target.csv
```

Run it once against the v8 target DB → `entity-types-target.csv`, once against the v9 source DB → `entity-types-source.csv`.

## Config Example

```json
{
  "removeElements": ["IsDataSourceSecret", "IsBPFEntity", "ExternalTypeName"],
  "removeAttributes": [
    { "tagName": "option", "attributes": ["ExternalValue", "Color"] }
  ],
  "addObjectTypeCode": true,
  "entityTypeCodes": {
    "target": "data/entity-types-target.csv",
    "source": "data/entity-types-source.csv"
  },
  "solutionVersion": {
    "targetVersion": "8.2.0009.0019",
    "targetPackageVersion": "8.2"
  }
}
```

### `solution.xml` version rewrite

Dynamics export writes the source version into the root element:

```xml
<ImportExportXml version="9.1.7.5" SolutionPackageVersion="9.1" ...>
```

v8 rejects that. The tool replaces **any** `9.x[.y.z…]` value on those two attributes with the configured v8 target:

| Attribute | Before | After |
|---|---|---|
| `version` | `9.1.7.5` (or any `9.*`) | `8.2.0009.0019` |
| `SolutionPackageVersion` | `9.1` (or any `9.*`) | `8.2` |

The solution's own `<Version>1.0.0.0</Version>` element is **not** touched. If the `solutionVersion` config block is omitted, `solution.xml` is left alone.

> Legacy single-CSV config (`"entityTypeCodesFile": "data/entity-types.csv"`) is still accepted and treated as the **target** CSV.

## Sample Output

```
═══════════════════════════════════════════════════════════
  CRM Solution Downgrade Summary  (v9 → v8)
═══════════════════════════════════════════════════════════
  Mode          : ZIP
  Input         : MySolution_1_0_0_0.zip
  Backup        : MySolution_1_0_0_0.zip.backup.2026-04-15T14-25-30-000Z
  XML entry     : customizations.xml
  Zip size      : 412.3 KB → 398.1 KB (-3.4%)
  XML size      : 1.43 MB → 1.31 MB (-8.4%)
  Duration      : 1.12s

  Elements Removed
  ────────────────────────────────────────
    IsDataSourceSecret                    42
    IsBPFEntity                           17
    ...
    TOTAL                                156

  ObjectTypeCode Elements Added
  ────────────────────────────────────────
    Mapped from target CSV                8
    Seeded from source CSV (new)          2
    Used '##' placeholder                 1
    TOTAL                                11

  ℹ  New entities not yet in target v8 (seeded with source code):
     - my_new_entity_a                        10047
     - my_new_entity_b                        10052
  → v8 will assign a fresh ObjectTypeCode on import; the seed is only a hint.

  ⚠  Entities missing from BOTH CSVs (using '##'):
     - legacy_orphan_entity
  → Replace '##' in the XML with real type codes before importing.

  solution.xml Version Rewrite
  ────────────────────────────────────────
    version                  9.1.7.5      →  8.2.0009.0019
    SolutionPackageVersion   9.1          →  8.2
═══════════════════════════════════════════════════════════
  ✓ Done
═══════════════════════════════════════════════════════════
```
