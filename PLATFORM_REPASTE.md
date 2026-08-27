# Platform re-paste checklist — Yanou rev 2.27

Replace the **entire** platform criteria list with `platform_criterion_revisions.txt` (IDs 1–60). Do not merge with the failed 2.26 paste.

## Why 2.26 failed (Geranium re-analysis)

Softening wording on non-atomic criteria was not enough. Geranium still saw bundled checks:

| Failed 2.26 item | Problem | 2.27 fix |
|------------------|---------|----------|
| Evidence Source (old #5) | Populated + no false FA absence + under-threshold assets in one criterion | **#5 / #6 / #7** (three atoms) |
| Acquisition Cost (old #7) | Per-row match + inventory total in one criterion | **#8 / #9** |
| Remaining Book Value (old #8) | Per-row + MD-00034 + aggregate in one criterion | **#10 / #11 / #12** |
| Disposition / C12-style | Accept vs escalate vs block mixed | **#51 / #52 / #53** |
| Stale platform IDs | Judges citing old C24–27 and counts 32/33/34/33 | Category answer key is **#41** only: **40 / 36 / 28 / 28** |

## Current answer keys (must match golden)

- Dashboard categories: Laptop **40**, Mobile **36**, Monitor **28**, Network Asset **28**
- Acquisition Cost inventory total: **$332,115**
- Remaining Book Value aggregate: **$61,526.60**
- MD-00034 Remaining Book Value: **$479.88**

## Slots freed to stay at 60

Dropped: Last Verified Date; Dashboard Status Counts; Dashboard Exception Type Counts; separate Exception Action vs Owner (merged into **#44**); “Missing as In Stock” negative (covered by Missing atoms **#16–#27**).

## Golden note

Re-upload `Yanou_IT_Asset_Reconciliation.xlsx` from this branch after pasting all 60 criteria.
