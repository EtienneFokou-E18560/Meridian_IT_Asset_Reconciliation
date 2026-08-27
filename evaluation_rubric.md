# Evaluation Rubric — Meridian IT Asset Reconciliation

**Revision:** 2.27 (post–Geranium re-analysis of failed 2.26 submit)  
**Weight model:** All criteria are **equally weighted** (weight = 1).  
**Scoring:** Each criterion is independently Pass or Fail.  
**Atomicity:** One observable check per criterion. Non-atomic bundles that Geranium flagged (C5, C7, C8, C12) are **split** into separate criteria — wording soften was not enough.  
**Platform limit:** Exactly **60** criteria (56 scoring + 4 anti-cheating negatives).

**What changed vs the failed 2.26 paste (Geranium follow-up):**
1. **C5 → three criteria:** Evidence Source populated; no false Fixed Assets absence; MD-00130/131/132 threshold wording.
2. **C7 → two criteria:** Per-row Acquisition Cost match; Inventory Acquisition Cost total **$332,115**.
3. **C8 → three criteria:** Per-row Remaining Book Value match/blank; MD-00034 = **$479.88**; Corrected Register Remaining Book Value aggregate **$61,526.60**.
4. **C12 clarified:** Accept under-threshold cost gaps; escalate/block only when blockers require.
5. **Slots freed (still 60):** Dropped Dashboard Status Counts and Dashboard Exception Type Counts (keep Category answer-key); merged Exception Action+Owner; dropped Last Verified Date; dropped “Missing as In Stock” negative (covered by Missing atoms).
6. **C11 kept:** Certification Sign-Off **worksheet** with Name, Signature, Date.
7. **Answer keys kept:** Category 40/36/28/28; RBV $61,526.60; MD-00034 $479.88; Acquisition total $332,115.

---

## Criteria (paste each block as Weight 1)

1. Inventory Number Fidelity  
The Corrected Register Inventory Number column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Inventory Number string on both sides for that asset).

2. Serial Number Fidelity  
The Corrected Register Serial Number column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Serial Number string on both sides for that asset).

3. Vendor Fidelity  
The Corrected Register Vendor column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Vendor string on both sides for that asset).

4. PO Number Fidelity  
The Corrected Register PO Number column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same PO Number string on both sides for that asset).

5. Evidence Source Populated  
Every Corrected Register row that represents a Fixed Assets asset has a non-blank Evidence Source cell.

6. Evidence Source — No False Fixed Assets Absence  
Evidence Source does not claim Fixed Assets has no row for an Inventory Number that actually exists in Fixed Assets (for example, it does not say Fixed Assets has no matching Inventory Number for MD-00001 when MD-00001 is present in Fixed Assets).

7. Evidence Source — Under-Threshold Cost Assets  
For MD-00130, MD-00131, and MD-00132, Evidence Source states that the cost difference is under the $100 threshold (or equivalent under-threshold wording) and does not treat those three assets as Cost Mismatch exceptions requiring escalation.

8. Acquisition Cost — Per-Row Match  
For every Corrected Register row that represents a Fixed Assets asset, Acquisition Cost equals that asset’s Acquisition Cost in Fixed Assets.

9. Acquisition Cost — Inventory Total  
The sum of Acquisition Cost across all Corrected Register Inventory rows equals $332,115.

10. Remaining Book Value — Per-Row Match  
For every Corrected Register row that represents a Fixed Assets asset: if Fixed Assets lists Remaining Book Value for that asset, Corrected Register Remaining Book Value matches it; if Fixed Assets has no Remaining Book Value for that asset, Corrected Register Remaining Book Value is blank.

11. Remaining Book Value — MD-00034  
MD-00034 Remaining Book Value equals $479.88.

12. Remaining Book Value — Corrected Register Aggregate  
The sum of Remaining Book Value across all Corrected Register rows that have a Remaining Book Value equals $61,526.60.

13. Status Fidelity  
The Corrected Register Status column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Status string on both sides for that asset).

14. Location Fidelity  
The Corrected Register Location column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Location string on both sides for that asset).

15. Department Fidelity  
The Corrected Register Department column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Department string on both sides for that asset).

16. Missing Asset — MD-00017  
MD-00017 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

17. Missing Asset — MD-00029  
MD-00029 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

18. Missing Asset — MD-00041  
MD-00041 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

19. Missing Asset — MD-00053  
MD-00053 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

20. Missing Asset — MD-00068  
MD-00068 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

21. Missing Asset — MD-00079  
MD-00079 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

22. Missing Asset — MD-00088  
MD-00088 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

23. Missing Asset — MD-00097  
MD-00097 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

24. Missing Asset — MD-00108  
MD-00108 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

25. Missing Asset — MD-00119  
MD-00119 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

26. Missing Asset — MD-00125  
MD-00125 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

27. Missing Asset — MD-00128  
MD-00128 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

28. Inter-Department Transfer — MD-00003  
MD-00003 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

29. Inter-Department Transfer — MD-00015  
MD-00015 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

30. Inter-Department Transfer — MD-00027  
MD-00027 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

31. Inter-Department Transfer — MD-00039  
MD-00039 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

32. Inter-Department Transfer — MD-00052  
MD-00052 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

33. Custody Gap — MD-00004  
MD-00004 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

34. Custody Gap — MD-00016  
MD-00016 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

35. Custody Gap — MD-00028  
MD-00028 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

36. Custody Gap — MD-00040  
MD-00040 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

37. Custody Gap — MD-00054  
MD-00054 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

38. Custody Gap — MD-00066  
MD-00066 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

39. Custody Gap — MD-00078  
MD-00078 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

40. Custody Gap — MD-00089  
MD-00089 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

41. Dashboard Category Counts  
Dashboard category counts equal Laptop 40, Mobile 36, Monitor 28, and Network Asset 28.

42. Critical Certification Blockers Present  
The workbook identifies the critical certification blockers required before sign-off (cost mismatches that exceed threshold, missing assets, and other blocker-class findings called out in Certification / Exception Register).

43. Exception Identity  
Each Exception Register finding row has a unique Exception ID and an Exception Type that correctly classifies the finding (for example Missing Asset, Cost Mismatch, Inter-Department Transfer, Custody Gap).

44. Exception Disposition  
Each Exception Register finding row states a clear recommended Action and an Owner (role or named party) responsible for follow-up.

45. Exception Severity  
Each Exception Register finding row has a Severity appropriate to the finding class (for example Critical for certification blockers such as over-threshold cost mismatches and missing assets; lower severity for non-blocking items).

46. Exception Source References  
Each Exception Register finding row cites the source systems or evidence used (for example Fixed Assets, ITAM, Purchase Orders, HR) so a reviewer can trace the finding.

47. Exception Register Financial Fields  
Exception Register financial fields for cost-related findings are populated from Fixed Assets / Purchase Orders amounts (not invented figures), including Acquisition Cost and variance where applicable.

48. Location Reconciliation Cross-Check  
Location Reconciliation compares Corrected Register locations to floor-plan / location evidence and flags material location disagreements that require follow-up.

49. Certification Sign-Off Structure  
A Certification Sign-Off worksheet exists and includes Name, Signature, and Date fields for the certifying party.

50. Certification Completeness Gate  
Certification content states that sign-off is blocked (or not complete) while critical blockers remain open — it does not present the register as fully certified while those blockers are unresolved.

51. Under-Threshold Cost Gaps — Accept  
MD-00130, MD-00131, and MD-00132 cost differences are treated as accept / no-escalation (under the $100 threshold), not as Cost Mismatch exceptions that block certification.

52. Over-Threshold Cost Mismatches — Escalate  
Over-threshold cost mismatches (the Cost Mismatch exception set) are escalated / listed as exceptions requiring remediation before clean certification.

53. Missing Assets — Block Certification  
Missing assets are treated as certification blockers (listed such that certification remains incomplete until addressed).

54. No Fabricated Inventory Numbers  
Corrected Register does not invent Inventory Numbers that do not appear in Fixed Assets or ITAM source extracts.

55. No Silent Drop of Fixed Assets Rows  
Every Inventory Number present in Fixed Assets appears in Corrected Register (no Fixed Assets asset silently omitted from the corrected inventory).

56. Negative — Blank Evidence Source Allowed  
FAIL the submission if Corrected Register Evidence Source is left blank for assets that exist in Fixed Assets.

57. Negative — Wrong Golden Category Counts  
FAIL the submission if Dashboard category counts are presented as Laptop 32, Mobile 33, Monitor 34, Network Asset 33 (or any set other than Laptop 40, Mobile 36, Monitor 28, Network Asset 28).

58. Negative — Under-Threshold Treated as Cost Mismatch Exceptions  
FAIL the submission if MD-00130, MD-00131, or MD-00132 are listed as Cost Mismatch exceptions requiring escalation despite under-threshold differences.

59. Negative — Certification Signed Off With Open Blockers  
FAIL the submission if Certification Sign-Off is presented as complete / signed off while critical blockers (over-threshold cost mismatches or missing assets) remain open.

60. Negative — Fabricated PO or Cost Figures  
FAIL the submission if Exception Register or Corrected Register uses Purchase Order numbers or dollar amounts that do not appear in the Fixed Assets / Purchase Orders source extracts for that asset.

---

## Paste checklist

- Paste **all 60** criteria above into the platform (replace any prior list).
- Confirm no stale criteria remain that cite location 24–27 as category counts or counts 32/33/34/33.
- Re-upload golden `Yanou_IT_Asset_Reconciliation.xlsx` from this branch after pasting.
- Do not leave `#61` or placeholder rows on the platform.

## Slot math (2.26 → 2.27)

| Change | Delta |
|--------|------:|
| Split Evidence Source (1→3) | +2 |
| Split Acquisition Cost (1→2) | +1 |
| Split Remaining Book Value (1→3) | +2 |
| Drop Last Verified Date | −1 |
| Drop Dashboard Status Counts | −1 |
| Drop Dashboard Exception Type Counts | −1 |
| Merge Exception Action + Owner | −1 |
| Drop Negative “Missing as In Stock” | −1 |
| **Net** | **0 (still 60)** |
