# Mixer Compound Planner – automation for `Mixing_Production_Plan.xlsm`

This automates the daily mixer planning for MX01 / MX02 / MX03.
**You type the same inputs as today. The workbook decides priority, chooses the
mixer, and builds the 1440-minute plan with line clearance.**

## What gets automated

| Step | Today (manual) | After automation |
|---|---|---|
| Stock | Paste LX02 → STOCK updates | Same |
| Weekly plan (Thursday) | Type image into STOCK H/K, then reduce K by hand every day | Type the image once (H = extrusion batches/day, K = `CONT` or the **full** weekly number). Released POs are deducted automatically. |
| Priority | Read STOCK, decide by eye | **Priority** sheet: days of cover, status (CRITICAL / WATCH / OK / LIMIT DONE / CHECK DATA), plan quantity, rank |
| Mixer choice | Decide by experience | Macro picks the mixer that has the recipe (Recipe Details) and the most free time, and you can override it |
| MX01/02/03 plan | Type every row, 3 + LC + 1 clean-out blocks, shift split | Macro writes the full sequence with dates, shifts, Sr. no, 22 codes, clean-out rows, line clearance and iterations |
| PO numbers | Copy from PO sheets into the plan | **Fill PO Numbers** reads MX01 PO / MX02 PO / MX03 PO and the clean-out sheets |
| Iterations from/to | Remember where each PO stopped | **Iteration Log** remembers the last iteration of every PO, and the next day continues from it |
| Tote requirement | Type codes and batches | **Build Tote Requirement** fills it from the plan |

## Planning rules used (from your process)

1. **Days of cover** = STOCK `Todays Stock` ÷ `Extrusion batches required/day`.
2. Cover **< 1.5 days → 1-CRITICAL** (press will stop). This must be planned today.
3. Cover **< 2.5 days → 2-WATCH**. These go in only if the mixer still has time after the critical ones.
4. **Plan quantity**
   * `CONT` → produce **3.5 days** of requirement (your 3–4 day rule; you can change it in Settings).
   * A number in STOCK K → plan **only the remaining weekly quantity**.
5. **Mixer batches** = ROUNDUP(extrusion batches × pallet kg ÷ mixer batch weight).
   I checked this against your own MX01/MX02 history and it matches
   (for example, FM708/2: 12 EX batches → 36 mixer batches; H5-28-10/0 625: 15 → 46).
6. **Mixer allocation**: a mixer is used only if `Recipe Details` has a batch weight for that 21 code on
   that mixer. Among the possible mixers, the one with the least planned time gets the job.
7. **22 codes with the same 21 code run back-to-back with no line clearance.** For example, FM457/0 289,
   475 and 750 run as one campaign with a single LC. This alone saves 80 minutes per avoided change-over.
8. **Line clearance block** between two different 21 codes: 3 clean-out (29 code) batches, then physical LC,
   then 1 clean-out batch, 80 minutes in total. If today's first compound is the same as the last one
   yesterday, no LC is added at the start.
9. Most urgent first. If two items have equal cover, the one consumed faster goes first.
   Batches are split across shifts A / B / C.
10. If the day is full, the item is planned **partially** and marked in the Priority remark column.

## One-time installation (about 5 minutes)

1. Take a backup copy of `Mixing_Production_Plan.xlsm`.
2. Open it, press **Alt + F11** (VBA editor), then **File → Import File…** and choose `MixerPlanner.bas`.
3. Close the editor, press **Alt + F8**, and run **`Setup_Planner`**. It creates these sheets:
   * **Control Panel**: buttons for the daily steps.
   * **Planner Settings**: edit the yellow cells. **Set the real mixing cycle time (minutes per batch) for each mixer.**
     6.5 min is only a placeholder.
   * **Priority**: formula sheet. It mirrors STOCK row by row, so **do not sort it**. Filter by Rank or Status instead.
   * **CO Map**: which clean-out 29 code to use before each compound (default 29000024, BB2030).
   * **Iteration Log**: last iteration used per PO. You can type the current values once for POs that are already running.
4. Save as `.xlsm`.

Your existing sheets, formulas and SAP layouts are not changed. Exceptions: the macro writes only the input
columns of MX01/02/03 (A–D, F–H, J–L), uses hidden helper columns BA:BE, and rewrites the
Tote sheet formulas. The Tote formulas previously pointed to an external file (`[8]Raw Data`); they now use
the workbook's own `Raw Data`.

## Daily routine

1. Paste SAP **LX02** (as today). STOCK and Priority update.
2. Check the **Priority** sheet. Type `MX01/MX02/MX03` in column S only if you want to force a mixer.
3. **Control Panel → 1. Build Mixer Plan**. MX01, MX02 and MX03 are filled.
4. Release the 21/22/29 POs in SAP and enter them in MX01 PO / MX02 PO / MX03 PO and the clean-out sheets
   (as today).
5. **2. Fill PO Numbers**. PO numbers and iteration from/to appear. Pink cells mean the PO was not found.
6. **3. Build Tote Requirement**.
7. At end of day: **4. Close Day**. It stores the iterations so tomorrow continues correctly.
   (Building tomorrow's plan also does this automatically for the old plan.)

**Thursday:** update STOCK D/H/K from the planning email image and set *Week start date* in Planner Settings.

## Files

* `MixerPlanner.bas`: the VBA module to import.
* `dry_run_planner.py`: the same rules in Python. It reads the xlsm without changing it and writes a preview:
  `python dry_run_planner.py Mixing_Production_Plan.xlsm preview.xlsx`

## Next improvements (phase 2)

* **RM Request**: add a BOM sheet (21 code → raw material kg per batch). The requirement column can then be
  calculated from the plan instead of typed.
* **Weekly plan image**: ask the planning team for the Excel behind the image. A paste-in sheet would then
  fill STOCK H/K with no typing. (Reading the emailed image with OCR is possible but error-prone for
  production numbers.)
* A per-compound cycle time if some compounds mix much longer than others.
