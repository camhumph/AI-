# CMS Geometry AI Training Workflow

This is the workflow for making Qwen classify DME/PCS mold-base parts the way CMS wants.

## 1. Install Ollama

Open PowerShell:

```powershell
irm https://ollama.com/install.ps1 | iex
```

Then open a new Command Prompt:

```cmd
ollama --version
```

## 2. Download Qwen

For the older ThinkPad, start with:

```cmd
ollama pull qwen3.5:4b
```

If it is too slow:

```cmd
ollama pull qwen3:1.7b
```

## 3. Collect Existing CAD Geometry Examples

```cmd
python C:\CMS_AI\geometry_classifier\collect_geometry_dataset.py
```

This creates:

```text
C:\CMS_AI\geometry_classifier\data\all_xt_export_rows.csv
C:\CMS_AI\geometry_classifier\data\job_geometry_summary.csv
C:\CMS_AI\geometry_classifier\data\geometry_ai_examples.jsonl
```

## 4. Ask Qwen To Analyze One XT Export

Use the path to any `XT_Export_CAD_Dimensions.csv`:

```cmd
python C:\CMS_AI\geometry_classifier\qwen_classify_xt_csv.py "C:\CMS_Local_Workspace\CMS_ACTIVE_QUOTE_8418 2 cavity funnel with water and baffle change 7-3-13_20260707_145953\XT_Export_CAD_Dimensions.csv"
```

Outputs go here:

```text
C:\CMS_AI\geometry_classifier\outputs
```

It writes:

- `<name>_qwen_classification.json`
- `<name>_qwen_classification.csv`
- `<name>_CORRECT_ME.csv`

## 5. Correct The AI

Open the `*_CORRECT_ME.csv`.

Fill in `CorrectRole` only where Qwen is wrong.

Allowed role names:

```text
top_clamp_plate
a_plate
b_plate
sc_retainer_plate
sc_backup_plate
stripper_plate
support_plate
bottom_clamp_plate
rail
ejector_plate
bottom_ejector_plate
ejector_retainer_plate (deprecated, legacy corrections only)
ejector_backup_plate
latch_lock
leader_pin
leader_pin_bushing
guided_ejector_bushing
return_pin
ejector_pin
support_pillar
pullcore
insert_or_core_detail
hardware_other
ignore
```

## 6. What "Training" Means At First

At first, training means:

1. Give Qwen CMS mold-base rules.
2. Give Qwen examples.
3. Correct Qwen's output.
4. Reuse those corrected examples in future prompts.

This is safer than fine-tuning right away.

## 6A. Important CMS Rules Learned From Real Jobs

Do not put these only in chat. They belong in the classifier rules, knowledge file, and corrected examples.

- Use `A Plate` / `B Plate`, not cavity/core, for CMS outputs.
- Exact shop tokens are strong anchor evidence, not weak notes -- they override generic geometry when a component name carries one: `A-PLATE`, `B-PLATE`, `SC-RETAINER`, `SC-BACKUP`, `EJ-RET`, `EJ-BACKUP`, `RAIL`, `LDR-PIN`, `LBB`. Only fall back to pure geometry when names are generic, stale, or missing.
- In the ejector assembly, the thinner plate is always `Ejector Plate` and the thicker/lower plate is `Bottom Ejector Plate`. Never call the thinner plate `Ejector Retainer Plate`.
- Ejector-stack and pin-plate rows must never be mapped or merged into the `A Plate` quote row.
- Decide the bottom of the stack from rails/ejector-stack plates first. Leader pins/bushings only decide orientation when rails/ejector plates are missing. Do not let leader-pin direction (including reversed/seated pins on SC bases) flip a clear rail/ejector stack or a strong `A-PLATE`/`B-PLATE` token.
- Latch lock parts (`PLC75`, `LATCH-LOCK`, `SAFETY-STRAP`, also seen misspelled `SAFTEY-STRAP`) mean the base may be a sequenced/latch-lock standard base, not a plain A/B/support stack.
- Latches help identify secondary opening/parting lines and which plates move together. Leader pins set guide direction; latches do not.
- Leader pins and bushings help identify guide direction, but multiple guide sets can exist. Do not flip the B Plate away from a strong `B-PLATE` token just because SC plates are present.
- For T001015-style bases, the correct pattern is:

```text
A Plate
B Plate
SC Retainer Plate
SC Backup Plate
Bottom Clamp Plate
Rails
Ejector Retainer Plate
Ejector Plate
```

## 7. Real Fine-Tuning Later

After 30-50 corrected jobs, make a fine-tuning dataset.

Then train a small model or table classifier.

Best final setup:

```text
VBA exports XT_Export_CAD_Dimensions.csv
Python geometry rules classify high-confidence parts
Qwen analyzes the whole mold and fixes uncertain rows
Corrected examples improve future runs
VBA fills quote from final classification CSV
```
