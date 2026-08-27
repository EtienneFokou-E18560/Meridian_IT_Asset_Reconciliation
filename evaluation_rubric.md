# Yanou IT Asset Reconciliation Rubric (rev 2.22 — PLATFORM PASTE)

Rev 2.22: Remove golden-only EX-0001/$85 from financial-exposure criterion; split Missing assets, unapproved transfers, and Critical-blocker custody chains into atomic criteria; raise custody-chain weight (≥12 pts total); drop serial-mismatch positive/negative double-count (keep positive only). Rev 2.21: Rewrite criteria 7/8/9 as pure declarative facts (remove grader-facing “fails” / “is not sufficient” scaffolding). Rev 2.20: Golden fidelity — LR PO citations aligned to hardware_purchase_orders.csv; LR control totals and Dashboard breakouts carry OOXML cached values; Certification blocker table; Critical certification blocker language on disposal/shipment exceptions; MD-00034 evidence NBV resolved (no raw XLOOKUP text). Rev 2.19: Lead criteria 7/8/9 with named expected values and correctness checks (not presence-only); update criterion 27 category counts to Laptop 40 / Mobile Device 36 / Monitor 28 / Network Asset 28.

Score `Yanou_IT_Asset_Reconciliation.xlsx` from the deliverable alone. Weights in weight field only.

**Counts:** 56 positive + 8 negative = **64** (platform IDs 1–64).

**Platform ID map:** 1–56 positives (incl. Critical certification blocker definition as 56) · 57–64 = N1–N8.

---

## Criteria (platform order)

### Positive criteria (1–56)

1. **Deliverable filename** (weight +5) — Produces Yanou_IT_Asset_Reconciliation.xlsx.  
   *Objective · Style/Formatting (identity; exclude from style share)*

2. **Corrected register: verified custodian** (weight +5) — For reviewed assets, the corrected asset register populates a non-blank verified custodian value (not merely a column header).  
   *Objective · Content*

3. **Corrected register: verified location** (weight +3) — For reviewed assets, the corrected asset register populates a non-blank verified location value.  
   *Objective · Content*

4. **Corrected register: verified lifecycle status** (weight +3) — For reviewed assets, the corrected asset register populates a non-blank verified lifecycle status value.  
   *Objective · Content*

5. **Corrected register: evidence source** (weight +3) — For reviewed assets, the corrected asset register populates a non-blank evidence source, and that text is consistent with the fixed-asset ledger extract: it must not claim there is no FA/ledger row for an asset tag that appears on the ledger; for MD-00130, MD-00131, and MD-00132 it may note ledger absence (with the under-threshold vs capital-qualifying distinction addressed elsewhere).  
   *Objective · Content*

6. **Corrected register: last verified date** (weight +3) — For reviewed assets, the corrected asset register populates a non-blank last verified date.  
   *Objective · Content*

7. **Corrected register: acquisition cost** (weight +3) — Each reviewed asset row has a non-blank numeric/currency acquisition cost matching source evidence, and the sum of corrected-register acquisition costs for reviewed assets equals $332,115 (±$1).  
   *Objective · Content*

8. **Corrected register: remaining book value** (weight +3) — Where a ledger NBV exists, remaining book value is non-blank and matches the ledger figure (blank only when no ledger row exists); MD-00034/FA-000034 remaining book value equals $479.88 (acquisition $930 − accumulated depreciation $450.12), and aggregate remaining book value across reviewed assets equals $61,526.60 (±$0.05).  
   *Objective · Content*

9. **Corrected register: confidence** (weight +3) — Every reviewed asset has a non-blank confidence value drawn only from the closed set {High, Medium, Low}.  
   *Objective · Content*

10. **Certification not approved while Critical blockers remain** (weight +3) — States the quarterly inventory is NOT APPROVED (or equivalent hold) while Critical certification blockers remain; does not give an unconditional approval.  
    *Subjective · Content*

11. **Certification page sign-off lines** (weight +4) — Includes a certification/sign-off page with sign-off lines for the IT Operations Manager, Finance Controller, HR Operations Lead, and Internal Auditor.  
    *Objective · Content*

12. **Certification page accept / escalate / block structure** (weight +4) — The certification page distinguishes what can be accepted (or accepted with conditions), what must escalate, and what blocks signing (Critical certification blockers).  
    *Objective · Content*

13. **Evidence precedence documented** (weight +4) — The deliverable documents a precedence order consistent with policy (ledger / verified disposal certificates / carrier acceptance+delivery+receiving-scan over approved transfers over HR over ticket status over technician notes and dock/exception images) and shows at least one rejected lower-precedence claim with the higher-precedence record that overruled it.  
    *Objective · Content*

14. **MD-00082 shipment mismatch held** (weight +4) — Marks MD-00082 as In Transit - Exception / shipment mismatch / equivalent unresolved custody, citing tracking 1ZMD00000082 and serial mismatch (MISMATCH-0082 vs MD-MO-050082), and does not clear it to Available.  
    *Objective · Content*

15. **Missing asset MD-00089** (weight +1) — Classifies MD-00089 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

16. **Missing asset MD-00090** (weight +1) — Classifies MD-00090 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

17. **Missing asset MD-00091** (weight +1) — Classifies MD-00091 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

18. **Missing asset MD-00092** (weight +1) — Classifies MD-00092 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

19. **Missing asset MD-00093** (weight +1) — Classifies MD-00093 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

20. **Missing asset MD-00094** (weight +1) — Classifies MD-00094 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

21. **Missing asset MD-00095** (weight +1) — Classifies MD-00095 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

22. **Missing asset MD-00096** (weight +1) — Classifies MD-00096 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

23. **Missing asset MD-00097** (weight +1) — Classifies MD-00097 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

24. **Missing asset MD-00098** (weight +1) — Classifies MD-00098 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

25. **Missing asset MD-00099** (weight +1) — Classifies MD-00099 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

26. **Missing asset MD-00100** (weight +1) — Classifies MD-00100 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
    *Objective · Content*

27. **Duplicate serials flagged** (weight +3) — Flags duplicate serial MD-LA-050021 across MD-00021 and MD-00025, and duplicate serial MD-NE-050088 across MD-00088 and MD-00092.  
    *Objective · Content*

28. **Unapproved transfer TR-00058** (weight +1) — Flags missing approval on transfer TR-00058.  
    *Objective · Content*

29. **Unapproved transfer TR-00059** (weight +1) — Flags missing approval on transfer TR-00059.  
    *Objective · Content*

30. **Unapproved transfer TR-00064** (weight +1) — Flags missing approval on transfer TR-00064.  
    *Objective · Content*

31. **Unapproved transfer TR-00065** (weight +1) — Flags missing approval on transfer TR-00065.  
    *Objective · Content*

32. **Unapproved transfer TR-00127** (weight +1) — Flags missing approval on transfer TR-00127.  
    *Objective · Content*

33. **RN-0132 under-threshold claim rejected** (weight +4) — Rejects regional note RN-0132 (or equivalent note claiming MD-00132 was expensed under threshold) as non-authoritative; does not clear MD-00132’s missing FA row on that basis.  
    *Objective · Content*

34. **Exception rows include related-record citation(s)** (weight +3) — Each exception row populates at least one related-record citation drawn from: employee ID, ticket ID, transfer ID, PO number, tracking number, certificate ID, or ledger asset ID (non-blank field value, not header-only).  
    *Objective · Content*

35. **Exception rows include financial exposure** (weight +3) — Each exception row populates a financial exposure value calculated from acquisition costs and book values.  
    *Objective · Content*

36. **Exception rows include required action** (weight +3) — Each exception row populates a required action.  
    *Objective · Content*

37. **Exception rows include owner by role** (weight +3) — Each exception row populates an owner by role.  
    *Objective · Content*

38. **Exception rows include action deadline** (weight +3) — Each unresolved exception row for a Critical certification blocker includes a populated resolution action deadline (specific date by which the escalation owner must complete the required action).  
    *Objective · Content*

39. **Custody chain: MD-00068** (weight +2) — Provides custody-chain detail for MD-00068 covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

40. **Custody chain: MD-00074** (weight +2) — Provides custody-chain detail for MD-00074 covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

41. **Custody chain: MD-00076** (weight +2) — Provides custody-chain detail for MD-00076 covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

42. **Custody chain: MD-00082** (weight +2) — Provides custody-chain detail for MD-00082 covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

43. **Custody chain: MD-00084** (weight +2) — Provides custody-chain detail for MD-00084 covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

44. **Custody chain: MD-00114** (weight +2) — Provides custody-chain detail for MD-00114 covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

45. **Custody chain: MD-00118** (weight +2) — Provides custody-chain detail for MD-00118 covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

46. **Custody chain: MD-00132** (weight +2) — Provides custody-chain detail for MD-00132 covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

47. **Dashboard breakout by verified status** (weight +4) — Summary dashboard includes asset counts and book value broken out by verified status.  
    *Objective · Content*

48. **Dashboard breakout by verified location** (weight +3) — Summary dashboard includes asset counts and book value broken out by verified location.  
    *Objective · Content*

49. **Dashboard breakout by device category** (weight +3) — Summary dashboard includes asset counts and book value broken out by device category with correct counts from the corrected register: Laptop 40, Mobile Device 36, Monitor 28, Network Asset 28 (counts ±0).  
    *Objective · Content*

50. **Dashboard breakout by exception type** (weight +3) — Summary dashboard includes a breakout by exception type (count and financial exposure or book value).  
    *Objective · Content*

51. **Ledger reconciliation explains expected vs Critical gaps** (weight +5) — Reconciles corrected inventory to the ledger and explains remaining differences with an explicit distinction between expected under-threshold FA absence (MD-00130/MD-00131) and Critical capital-qualifying absence (MD-00132), plus cost-basis mismatches and disposal-status disagreements.  
    *Objective · Content*

52. **Overdue returns without receiving proof block certification** (weight +4) — Treats overdue / label-only returns that lack carrier acceptance + delivery + receiving-scan proof as Critical certification blockers (for example MD-00068).  
    *Objective · Content*

53. **Shipment serial mismatches treated as Critical certification blockers** (weight +4) — Treats unresolved shipment serial mismatches as Critical certification blockers for at least 3 of these 4 assets: MD-00074, MD-00076, MD-00082, and MD-00084.  
    *Objective · Content*

54. **Missing disposal certificates block certification** (weight +4) — Treats retired/pending-disposal assets lacking a verified disposal certificate as Critical certification blockers (MD-00114 and MD-00118).  
    *Objective · Content*

55. **Capital-asset ledger gaps block certification** (weight +4) — Treats capital-qualifying assets (acquisition cost ≥ $2,500) that are missing from the fixed-asset ledger as Critical certification blockers.  
    *Objective · Content*

56. **Critical certification blocker definition** (weight +2) — Defines a Critical certification blocker as an unresolved asset that meets any of: (a) shipment serial mismatch; (b) Missing / Cannot Locate; (c) retired/pending-disposal without a verified disposal certificate; (d) overdue / label-only return lacking carrier acceptance + delivery + receiving-scan proof; or (e) capital-qualifying asset (acquisition cost ≥ $2,500) missing from the fixed-asset ledger.  
    *Objective · Content*


### Negative criteria N1–N8 (57–64)

57. **N1. Unconditional certification approval** (weight -5) — Approves/certifies the quarterly inventory as complete despite Critical certification blockers.  
    *Objective · Content*

58. **N2. Label-created treated as returned** (weight -5) — Marks Label Created shipments as Available/returned/completed without carrier acceptance + delivery + receiving scan evidence.  
    *Objective · Content*

59. **N3. Dock image used as receiving clearance** (weight -4) — Uses receiving_exception_scan_1ZMD00000082.png as proof MD-00082 was successfully received/cleared.  
    *Objective · Content*

60. **N4. Untraceable transfer/return/approval/disposal IDs** (weight -5) — Cites or uses transfer IDs, return receipts, approval IDs, or disposal certificates that cannot be traced to a source record in the provided files.  
    *Objective · Content*

61. **N5. Missing assets marked verified in use/stock** (weight -4) — Treats unresolved MD-00089–MD-00100 assets as verified In Use or Available stock.  
    *Objective · Content*

62. **N6. Under-threshold absence treated as Critical capital gap** (weight -5) — Treats MD-00130 and/or MD-00131 ledger absence as a Critical certification blocker, or uses a capitalization_approved=Yes purchase-order flag to force FA expectations for those under-threshold lines.  
    *Objective · Content*

63. **N7. RN-0132 accepted to clear MD-00132** (weight -5) — Accepts the regional note that MD-00132 was expensed under threshold (or equivalent) and clears/omits the Critical missing FA-row finding for MD-00132.  
    *Objective · Content*

64. **N8. False ledger-absence claim** (weight -5) — Claims that an asset has no fixed-asset ledger / FAR row when that asset tag appears on the provided fixed-asset ledger extract.  
    *Objective · Content*
