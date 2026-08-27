# Yanou IT Asset Reconciliation Rubric (rev 2.19 — PLATFORM PASTE)

Rev 2.21: Rewrite criteria 7/8/9 as pure declarative facts (remove grader-facing “fails” / “is not sufficient” scaffolding). Rev 2.20: Golden fidelity — LR PO citations aligned to hardware_purchase_orders.csv; LR control totals and Dashboard breakouts carry OOXML cached values; Certification blocker table; Critical certification blocker language on disposal/shipment exceptions; MD-00034 evidence NBV resolved (no raw XLOOKUP text). Rev 2.19: Lead criteria 7/8/9 with named expected values and correctness checks (not presence-only); update criterion 27 category counts to Laptop 40 / Mobile Device 36 / Monitor 28 / Network Asset 28. Rev 2.16: Align flat platform criterion IDs — Critical certification blocker definition is criterion 39 (after N1–N5); former N6–N9 are criteria 40–43. Positives without the definition are 1–33; N1–N5 are 34–38. Rev 2.15: Remove cross-criterion references from criterion bodies. Rev 2.14: Split former P10 into NOT APPROVED hold + Critical blocker definition for platform paste length. Rev 2.13: P7/P8/P9/P20/P27 name expected computed values. Rev 2.12: Critical certification blocker definition introduced; named tags MD-00068/74/76/82/84/114/118/132 on custody criterion. Rev 2.11: populated-value requirements; N9 false FA-absence claim. Rev 2.10: deadline sentence. Rev 2.9: custody instruction language.

Score `Yanou_IT_Asset_Reconciliation.xlsx` from the deliverable alone. Weights in weight field only.

**Counts:** 34 positive + 9 negative = **43** (platform IDs 1–43).

**Platform ID map:** 1–33 positives (excl. blocker definition) · 34–38 = N1–N5 · **39 = Critical certification blocker definition** · **40–43 = N6–N9**.

---

## Criteria (platform order)

### Positive criteria (1–33)

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

15. **Missing assets MD-00089–MD-00100 (threshold)** (weight +4) — Classifies at least 10 of the 12 assets MD-00089 through MD-00100 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock.  
   *Objective · Content*

16. **Duplicate serials flagged** (weight +3) — Flags duplicate serial MD-LA-050021 across MD-00021 and MD-00025, and duplicate serial MD-NE-050088 across MD-00088 and MD-00092.  
   *Objective · Content*

17. **Unapproved transfers flagged** (weight +3) — Flags missing approval on transfers TR-00058, TR-00059, TR-00064, TR-00065, and TR-00127.  
   *Objective · Content*

18. **RN-0132 under-threshold claim rejected** (weight +4) — Rejects regional note RN-0132 (or equivalent note claiming MD-00132 was expensed under threshold) as non-authoritative; does not clear MD-00132’s missing FA row on that basis.  
   *Objective · Content*

19. **Exception rows include related-record citation(s)** (weight +3) — Each exception row populates at least one related-record citation drawn from: employee ID, ticket ID, transfer ID, PO number, tracking number, certificate ID, or ledger asset ID (non-blank field value, not header-only).  
   *Objective · Content*

20. **Exception rows include financial exposure** (weight +3) — Each exception row populates a financial exposure value that matches the finding math, including EX-0001 financial exposure of $85 (non-blank, not header-only).  
   *Objective · Content*

21. **Exception rows include required action** (weight +3) — Each exception row populates a required action.  
   *Objective · Content*

22. **Exception rows include owner by role** (weight +3) — Each exception row populates an owner by role.  
   *Objective · Content*

23. **Exception rows include action deadline** (weight +3) — Each unresolved exception row for a Critical certification blocker includes a populated resolution action deadline (specific date by which the escalation owner must complete the required action).  
   *Objective · Content*

24. **Custody chains for Critical certification blockers** (weight +3) — Provides custody-chain detail for each unresolved Critical certification blocker, including at least MD-00068, MD-00074, MD-00076, MD-00082, MD-00084, MD-00114, MD-00118, and MD-00132, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
   *Objective · Content*

25. **Dashboard breakout by verified status** (weight +4) — Summary dashboard includes asset counts and book value broken out by verified status.  
   *Objective · Content*

26. **Dashboard breakout by verified location** (weight +3) — Summary dashboard includes asset counts and book value broken out by verified location.  
   *Objective · Content*

27. **Dashboard breakout by device category** (weight +3) — Summary dashboard includes asset counts and book value broken out by device category with correct counts from the corrected register: Laptop 40, Mobile Device 36, Monitor 28, Network Asset 28 (counts ±0).  
   *Objective · Content*

28. **Dashboard breakout by exception type** (weight +3) — Summary dashboard includes a breakout by exception type (count and financial exposure or book value).  
   *Objective · Content*

29. **Ledger reconciliation explains expected vs Critical gaps** (weight +5) — Reconciles corrected inventory to the ledger and explains remaining differences with an explicit distinction between expected under-threshold FA absence (MD-00130/MD-00131) and Critical capital-qualifying absence (MD-00132), plus cost-basis mismatches and disposal-status disagreements.  
   *Objective · Content*

30. **Overdue returns without receiving proof block certification** (weight +4) — Treats overdue / label-only returns that lack carrier acceptance + delivery + receiving-scan proof as Critical certification blockers (for example MD-00068).  
   *Objective · Content*

31. **Shipment serial mismatches treated as Critical certification blockers** (weight +4) — Treats unresolved shipment serial mismatches as Critical certification blockers for at least 3 of these 4 assets: MD-00074, MD-00076, MD-00082, and MD-00084.  
   *Objective · Content*

32. **Missing disposal certificates block certification** (weight +4) — Treats retired/pending-disposal assets lacking a verified disposal certificate as Critical certification blockers (MD-00114 and MD-00118).  
   *Objective · Content*

33. **Capital-asset ledger gaps block certification** (weight +4) — Treats capital-qualifying assets (acquisition cost ≥ $2,500) that are missing from the fixed-asset ledger as Critical certification blockers.  
   *Objective · Content*


### Negative criteria N1–N5 (34–38)

34. **N1. Unconditional certification approval** (weight -5) — Approves/certifies the quarterly inventory as complete despite Critical certification blockers.  
   *Objective · Content*

35. **N2. Label-created treated as returned** (weight -5) — Marks Label Created shipments as Available/returned/completed without carrier acceptance + delivery + receiving scan evidence.  
   *Objective · Content*

36. **N3. Clears serial mismatches / fails to treat as blockers** (weight -5) — Fails to treat unresolved shipment serial mismatches as Critical certification blockers for at least 3 of these 4 assets — MD-00074, MD-00076, MD-00082, and MD-00084 — or clears those mismatches to Available / completed return.  
   *Objective · Content*

37. **N4. Dock image used as receiving clearance** (weight -4) — Uses receiving_exception_scan_1ZMD00000082.png as proof MD-00082 was successfully received/cleared.  
   *Objective · Content*

38. **N5. Untraceable transfer/return/approval/disposal IDs** (weight -5) — Cites or uses transfer IDs, return receipts, approval IDs, or disposal certificates that cannot be traced to a source record in the provided files.  
   *Objective · Content*


### Critical certification blocker definition (39)

39. **Critical certification blocker definition** (weight +2) — Defines a Critical certification blocker as an unresolved asset that meets any of: (a) shipment serial mismatch; (b) Missing / Cannot Locate; (c) retired/pending-disposal without a verified disposal certificate; (d) overdue / label-only return lacking carrier acceptance + delivery + receiving-scan proof; or (e) capital-qualifying asset (acquisition cost ≥ $2,500) missing from the fixed-asset ledger.  
   *Objective · Content*


### Negative criteria N6–N9 (40–43)

40. **N6. Missing assets marked verified in use/stock** (weight -4) — Treats unresolved MD-00089–MD-00100 assets as verified In Use or Available stock.  
   *Objective · Content*

41. **N7. Under-threshold absence treated as Critical capital gap** (weight -5) — Treats MD-00130 and/or MD-00131 ledger absence as a Critical certification blocker, or uses a capitalization_approved=Yes purchase-order flag to force FA expectations for those under-threshold lines.  
   *Objective · Content*

42. **N8. RN-0132 accepted to clear MD-00132** (weight -5) — Accepts the regional note that MD-00132 was expensed under threshold (or equivalent) and clears/omits the Critical missing FA-row finding for MD-00132.  
   *Objective · Content*

43. **N9. False ledger-absence claim** (weight -5) — Claims that an asset has no fixed-asset ledger / FAR row when that asset tag appears on the provided fixed-asset ledger extract.  
   *Objective · Content*

