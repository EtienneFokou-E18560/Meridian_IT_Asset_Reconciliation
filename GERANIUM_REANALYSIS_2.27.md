# Geranium re-analysis — failed 2.26 submit → rev 2.27

## Verdict

The failed submit was not primarily a golden-workbook emptiness problem. Local golden still carries Dashboard/LR formula caches and the verified totals (40/36/28/28, $332,115, $61,526.60, MD-00034 $479.88). The rubric paste itself still failed Geranium-style atomicity / clarity checks that earlier feedback already named.

## Earlier Geranium findings vs what 2.26 actually shipped

| Earlier finding | 2.26 response | Still failing? | 2.27 action |
|-----------------|---------------|----------------|-------------|
| C5 non-atomic (Evidence Source: populated + false-absence + threshold assets) | Softened / reordered wording; still one criterion | Yes | Split → #5, #6, #7 |
| C7 non-atomic (per-row Acquisition Cost + inventory total) | Still one criterion | Yes | Split → #8, #9 |
| C8 non-atomic (per-row RBV + MD-00034 + aggregate) | Still one criterion | Yes | Split → #10, #11, #12 |
| C12 mixed accept / escalate / block | Partially clarified | Partially | Split → #51, #52, #53 |
| C11 subjective sign-off | Fixed: worksheet + Name/Signature/Date | No (keep) | Kept as #49 |
| Answer keys required | Present | Keep | Kept (#9, #11, #12, #41) |
| Stale platform criteria (C24–27, 32/33/34/33) | Documented in PLATFORM_REPASTE | Operational risk | #41 is sole category answer key; status/type dashboard criteria removed |

## Why “wording soften” failed

Geranium grades **one observable check per criterion**. A single Pass/Fail that requires three independent facts (e.g. every Evidence Source filled **and** no false FA absence **and** MD-00130/131/132 under-threshold language) is still non-atomic even if the prose is clearer. The fix is structural split, not synonym polish.

## Tradeoffs under the 60-cap

To add +5 atoms from C5/C7/C8 splits without exceeding 60:

- Dropped Last Verified Date
- Dropped Dashboard Status Counts and Dashboard Exception Type Counts (Category answer key retained)
- Merged Exception Action + Owner → #44
- Dropped negative “Missing as In Stock” (Missing atoms #16–#27 already cover)

## What the user must do on the platform

1. Delete / replace the entire prior criteria list (do not edit 2.26 in place).
2. Paste all 60 blocks from `platform_criterion_revisions.txt`.
3. Re-upload golden `Yanou_IT_Asset_Reconciliation.xlsx` from this branch.
4. Confirm no leftover criteria still mention 32/33/34/33 or old C24–27 location/category wording.
