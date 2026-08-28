# Evaluation Rubric — Meridian IT Asset Reconciliation

**Revision:** 2.31 (no cross-criterion references in criteria text)  
**Weight model:** Positive criteria sum to **+100**. Negative (anti-cheating) criteria are **penalties** (−5 each).  
**Scoring:** Each criterion is independently Pass or Fail.  
- Pass a **positive** criterion → add its points.  
- Trigger a **negative** criterion (prohibited pattern present) → subtract its points.  
- Perfect clean submission = **100%** (all positives pass, no negatives triggered).  
**Atomicity:** One observable check per criterion.  
**Platform limit:** Exactly **60** criteria (55 positive + 5 negative).

**Weight bands:**
| Band | Score | Criteria |
|------|------:|----------|
| Deliverable filename | +2 | 1 |
| Answer keys / high-stakes gates | +3 | 6, 9–13, 42–43, 49–55 |
| Core fidelity / exception structure | +2 | 2–5, 7–8, 14–16, 44–48, 56 |
| Atomic missing / transfer / custody | +1 | 17–41 |
| Anti-cheating penalties | −5 | 56–60 |

---

## Criteria (paste with the stated Weight)


1. Deliverable Filename [+2]
Weight: +2
The submission is delivered as a workbook named exactly Yanou_IT_Asset_Reconciliation.xlsx (case-sensitive, no _Final/_v2/(1)/spacing variants).

2. Inventory Number Fidelity [+2]
Weight: +2
The Corrected Register Inventory Number column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Inventory Number string on both sides for that asset).

3. Serial Number Fidelity [+2]
Weight: +2
The Corrected Register Serial Number column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Serial Number string on both sides for that asset).

4. Vendor Fidelity [+2]
Weight: +2
The Corrected Register Vendor column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Vendor string on both sides for that asset).

5. PO Number Fidelity [+2]
Weight: +2
The Corrected Register PO Number column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same PO Number string on both sides for that asset).

6. Evidence Source Populated [+3]
Weight: +3
Every Corrected Register row that represents a Fixed Assets asset has a non-blank Evidence Source cell.

7. Evidence Source — No False Fixed Assets Absence [+2]
Weight: +2
Evidence Source does not claim Fixed Assets has no row for an Inventory Number that actually exists in Fixed Assets (for example, it does not say Fixed Assets has no matching Inventory Number for MD-00001 when MD-00001 is present in Fixed Assets).

8. Evidence Source — Under-Threshold Cost Assets [+2]
Weight: +2
For MD-00130, MD-00131, and MD-00132, Evidence Source states that the cost difference is under the $100 threshold (or equivalent under-threshold wording) and does not treat those three assets as Cost Mismatch exceptions requiring escalation.

9. Acquisition Cost — Per-Row Match [+3]
Weight: +3
For every Corrected Register row that represents a Fixed Assets asset, Acquisition Cost equals that asset’s Acquisition Cost in Fixed Assets.

10. Acquisition Cost — Inventory Total [+3]
Weight: +3
The sum of Acquisition Cost across all Corrected Register Inventory rows equals $332,115.

11. Remaining Book Value — Per-Row Match [+3]
Weight: +3
For every Corrected Register row that represents a Fixed Assets asset: if Fixed Assets lists Remaining Book Value for that asset, Corrected Register Remaining Book Value matches it; if Fixed Assets has no Remaining Book Value for that asset, Corrected Register Remaining Book Value is blank.

12. Remaining Book Value — MD-00034 [+3]
Weight: +3
MD-00034 Remaining Book Value equals $479.88.

13. Remaining Book Value — Corrected Register Aggregate [+3]
Weight: +3
The sum of Remaining Book Value across all Corrected Register rows that have a Remaining Book Value equals $61,526.60.

14. Status Fidelity [+2]
Weight: +2
The Corrected Register Status column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Status string on both sides for that asset).

15. Location Fidelity [+2]
Weight: +2
The Corrected Register Location column matches the Fixed Assets extract for every asset that appears in Fixed Assets (same Location string on both sides for that asset).

16. Missing Asset — MD-00017 [+1]
Weight: +1
MD-00017 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

17. Missing Asset — MD-00029 [+1]
Weight: +1
MD-00029 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

18. Missing Asset — MD-00041 [+1]
Weight: +1
MD-00041 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

19. Missing Asset — MD-00053 [+1]
Weight: +1
MD-00053 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

20. Missing Asset — MD-00068 [+1]
Weight: +1
MD-00068 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

21. Missing Asset — MD-00079 [+1]
Weight: +1
MD-00079 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

22. Missing Asset — MD-00088 [+1]
Weight: +1
MD-00088 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

23. Missing Asset — MD-00097 [+1]
Weight: +1
MD-00097 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

24. Missing Asset — MD-00108 [+1]
Weight: +1
MD-00108 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

25. Missing Asset — MD-00119 [+1]
Weight: +1
MD-00119 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

26. Missing Asset — MD-00125 [+1]
Weight: +1
MD-00125 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

27. Missing Asset — MD-00128 [+1]
Weight: +1
MD-00128 appears in Fixed Assets and does not appear in ITAM; it is listed as Missing (or equivalent missing disposition) in Exception Register and/or Corrected Register.

28. Inter-Department Transfer — MD-00003 [+1]
Weight: +1
MD-00003 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

29. Inter-Department Transfer — MD-00015 [+1]
Weight: +1
MD-00015 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

30. Inter-Department Transfer — MD-00027 [+1]
Weight: +1
MD-00027 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

31. Inter-Department Transfer — MD-00039 [+1]
Weight: +1
MD-00039 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

32. Inter-Department Transfer — MD-00052 [+1]
Weight: +1
MD-00052 shows an ITAM vs Fixed Assets department disagreement and is identified as an inter-department transfer (or equivalent transfer disposition) in Exception Register and/or Corrected Register.

33. Custody Gap — MD-00004 [+1]
Weight: +1
MD-00004 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

34. Custody Gap — MD-00016 [+1]
Weight: +1
MD-00016 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

35. Custody Gap — MD-00028 [+1]
Weight: +1
MD-00028 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

36. Custody Gap — MD-00040 [+1]
Weight: +1
MD-00040 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

37. Custody Gap — MD-00054 [+1]
Weight: +1
MD-00054 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

38. Custody Gap — MD-00066 [+1]
Weight: +1
MD-00066 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

39. Custody Gap — MD-00078 [+1]
Weight: +1
MD-00078 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

40. Custody Gap — MD-00089 [+1]
Weight: +1
MD-00089 has a custody / assigned-user gap relative to policy expectations and is identified with a custody-related disposition in Exception Register and/or Corrected Register.

41. Dashboard Category Counts [+3]
Weight: +3
Dashboard category counts equal Laptop 40, Mobile 36, Monitor 28, and Network Asset 28.

42. Critical Certification Blockers Present [+3]
Weight: +3
The workbook identifies the critical certification blockers required before sign-off (cost mismatches that exceed threshold, missing assets, and other blocker-class findings called out in Certification / Exception Register).

43. Exception Identity [+2]
Weight: +2
Each Exception Register finding row has a unique Exception ID and an Exception Type that correctly classifies the finding (for example Missing Asset, Cost Mismatch, Inter-Department Transfer, Custody Gap).

44. Exception Disposition [+2]
Weight: +2
Each Exception Register finding row states a clear recommended Action and an Owner (role or named party) responsible for follow-up.

45. Exception Severity [+2]
Weight: +2
Each Exception Register finding row has a Severity appropriate to the finding class (for example Critical for certification blockers such as over-threshold cost mismatches and missing assets; lower severity for non-blocking items).

46. Exception Source References [+2]
Weight: +2
Each Exception Register finding row cites the source systems or evidence used (for example Fixed Assets, ITAM, Purchase Orders, HR) so a reviewer can trace the finding.

47. Exception Register Financial Fields [+2]
Weight: +2
Exception Register financial fields for cost-related findings are populated from Fixed Assets / Purchase Orders amounts (not invented figures), including Acquisition Cost and variance where applicable.

48. Location Reconciliation Cross-Check [+3]
Weight: +3
Location Reconciliation compares Corrected Register locations to floor-plan / location evidence and flags material location disagreements that require follow-up.

49. Certification Sign-Off Structure [+3]
Weight: +3
A Certification Sign-Off worksheet exists and includes Name, Signature, and Date fields for the certifying party.

50. Certification Completeness Gate [+3]
Weight: +3
Certification content states that sign-off is blocked (or not complete) while critical blockers remain open — it does not present the register as fully certified while those blockers are unresolved.

51. Exception Register — MD-00130/131 Non-Blockers [+3]
Weight: +3
MD-00130 and MD-00131 each appear in Exception Register with an exception Type that classifies the missing ledger row as below capitalization / under-threshold (for example Below Capitalization Threshold), Severity is non-Critical (for example Low or Informational), and Required Action is informational or no-escalation (not presented as a certification blocker).

52. Over-Threshold Cost Mismatches — Escalate [+3]
Weight: +3
Over-threshold cost mismatches (the Cost Mismatch exception set) are escalated / listed as exceptions requiring remediation before clean certification.

53. Missing Assets — Block Certification [+3]
Weight: +3
Missing assets are treated as certification blockers (listed such that certification remains incomplete until addressed).

54. No Fabricated Inventory Numbers [+3]
Weight: +3
Corrected Register does not invent Inventory Numbers that do not appear in Fixed Assets or ITAM source extracts.

55. No Silent Drop of Fixed Assets Rows [+2]
Weight: +2
Every Inventory Number present in Fixed Assets appears in Corrected Register (no Fixed Assets asset silently omitted from the corrected inventory).

56. Negative — Blank Evidence Source Allowed [-5]
Weight: -5
FAIL the submission if Corrected Register Evidence Source is left blank for assets that exist in Fixed Assets.

57. Negative — Wrong Golden Category Counts [-5]
Weight: -5
FAIL the submission if Dashboard category counts are presented as Laptop 32, Mobile 33, Monitor 34, Network Asset 33 (or any set other than Laptop 40, Mobile 36, Monitor 28, Network Asset 28).

58. Negative — Under-Threshold Treated as Cost Mismatch Exceptions [-5]
Weight: -5
FAIL the submission if MD-00130, MD-00131, or MD-00132 are listed as Cost Mismatch exceptions requiring escalation despite under-threshold differences.

59. Negative — Certification Signed Off With Open Blockers [-5]
Weight: -5
FAIL the submission if Certification Sign-Off is presented as complete / signed off while critical blockers (over-threshold cost mismatches or missing assets) remain open.

60. Negative — Fabricated PO or Cost Figures [-5]
Weight: -5
FAIL the submission if Exception Register or Corrected Register uses Purchase Order numbers or dollar amounts that do not appear in the Fixed Assets / Purchase Orders source extracts for that asset. ---

---

## Paste checklist

- Paste **all 60** criteria with their **Weight** values (not flat 1).
- Criterion **#1** must quote the exact basename `Yanou_IT_Asset_Reconciliation.xlsx`.
- Confirm positive weights sum to **+100** and negatives are **−5** each.
- Re-upload golden `Yanou_IT_Asset_Reconciliation.xlsx` from this branch after pasting.

## Score math

| Bucket | Count | Points |
|--------|------:|-------:|
| Positive (+3) | 15 | +45 |
| Positive (+2) | 15 | +30 |
| Positive (+1) | 25 | +25 |
| **Positive total (max)** | **55** | **+100** |
| Negative (−5) | 5 | −25 if all triggered |
| **Perfect clean score** | | **100%** |
