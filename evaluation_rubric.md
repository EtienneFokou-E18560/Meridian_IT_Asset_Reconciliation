# Yanou IT Asset Reconciliation Rubric (rev 2.12 — PLATFORM PASTE)

Rev 2.12: dropped redundant negative mirrors of P10/P15/P18/P31. Rev 2.11: P2–P9 and P19–P23 require populated values (not header-only); P5 Evidence Source must be consistent with ledger presence; remaining N5 penalizes false FAR-absence claims. Rev 2.10: P23 declarative deadline sentence. Rev 2.9: P24 instruction language.

Score `Yanou_IT_Asset_Reconciliation.xlsx` from the deliverable alone. Weights in weight field only.

**Counts:** 33 positive + 5 negative = **38**.

---

## Positive criteria

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

7. **Corrected register: acquisition cost** (weight +3) — For reviewed assets, the corrected asset register populates a numeric acquisition cost.  
   *Objective · Content*

8. **Corrected register: remaining book value** (weight +3) — For reviewed assets that have a net book value on the fixed-asset ledger extract, the corrected asset register populates remaining book value from that ledger figure (blank is acceptable only when the asset has no ledger row).  
   *Objective · Content*

9. **Corrected register: confidence** (weight +3) — For reviewed assets, the corrected asset register populates a non-blank confidence value.  
   *Objective · Content*

10. **Certification not approved / Critical certification blockers retained** (weight +5) — States the quarterly inventory is NOT APPROVED (or equivalent hold) while Critical certification blockers remain; does not give an unconditional approval.  
    *Subjective · Content*

11. **Certification page sign-off lines** (weight +4) — Includes a certification/sign-off page with sign-off lines for the IT Operations Manager, Finance Controller, HR Operations Lead, and Internal Auditor.  
    *Objective · Content*

12. **Certification page accept / escalate / block structure** (weight +4) — The certification page distinguishes what can be accepted (or accepted with conditions), what must escalate, and what blocks signing.  
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

20. **Exception rows include financial exposure** (weight +3) — Each exception row populates a financial exposure value.  
    *Objective · Content*

21. **Exception rows include required action** (weight +3) — Each exception row populates a required action.  
    *Objective · Content*

22. **Exception rows include owner by role** (weight +3) — Each exception row populates an owner by role.  
    *Objective · Content*

23. **Exception rows include action deadline** (weight +3) — Each certification-blocking unresolved exception row includes a **populated** resolution action deadline (specific date by which the escalation owner must complete the required action), not merely a deadline column header.  
    *Objective · Content*

24. **Custody chains for certification-blocking assets** (weight +3) — Provides custody-chain detail for each high-risk / certification-blocking unresolved asset, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events.  
    *Objective · Content*

25. **Dashboard breakout by verified status** (weight +4) — Summary dashboard includes asset counts and book value broken out by verified status.  
    *Objective · Content*

26. **Dashboard breakout by verified location** (weight +3) — Summary dashboard includes asset counts and book value broken out by verified location.  
    *Objective · Content*

27. **Dashboard breakout by device category** (weight +3) — Summary dashboard includes asset counts and book value broken out by device category.  
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

## Negative criteria

N1. **Label-created treated as returned** (weight -5) — Marks Label Created shipments as Available/returned/completed without carrier acceptance + delivery + receiving scan evidence.  
*Objective · Content*

N2. **Dock image used as receiving clearance** (weight -4) — Uses receiving_exception_scan_1ZMD00000082.png as proof MD-00082 was successfully received/cleared.  
*Objective · Content*

N3. **Untraceable transfer/return/approval/disposal IDs** (weight -5) — Cites or uses transfer IDs, return receipts, approval IDs, or disposal certificates that cannot be traced to a source record in the provided files.  
*Objective · Content*

N4. **Under-threshold absence treated as Critical capital gap** (weight -5) — Treats MD-00130 and/or MD-00131 ledger absence as a Critical certification blocker, or uses a capitalization_approved=Yes purchase order to force FA expectations for those under-threshold lines.  
*Objective · Content*

N5. **False fixed-asset ledger absence** (weight -5) — Claims fixed-asset ledger absence (no FA / FAR silent / no matching FAR row) for an asset tag that appears on the provided fixed-asset ledger extract.  
*Objective · Content*
