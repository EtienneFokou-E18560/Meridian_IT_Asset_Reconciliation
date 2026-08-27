# Yanou IT Asset Reconciliation Rubric (rev 2.23 — PLATFORM PASTE, ≤60)

Rev 2.23: Fit platform 60-criterion cap while keeping the quality-review “excellent” atomic splits; restore C7–C9 expected values; restore capital-gap Critical blocker; drop Critical-blocker definition, dashboard-by-location, RN-0132-accepted negative, and false-ledger-absence negative to fit. Rev 2.22: Remove golden-only EX-0001/$85; atomic Missing/transfer/custody splits; raise custody weight; drop serial-mismatch double-count.

Score `Yanou_IT_Asset_Reconciliation.xlsx` from the deliverable alone. Weights in weight field only.

**Counts:** 54 positive + 6 negative = **60** (platform max).

**Edits vs suggested “excellent” draft:** Restored C7–C9 named expected values (draft was presence-only); restored capital-qualifying FAR-gap Critical blocker (draft omitted it); dropped Critical-blocker definition, dashboard-by-location, RN-0132-accepted negative, and false-ledger-absence negative to fit the 60 cap (RN-0132 still covered by positive #33; false FA-absence still constrained by positive #5).

---

## Criteria (platform order 1–60)

1. Produces Yanou_IT_Asset_Reconciliation.xlsx. | weight: 5
2. For reviewed assets, the corrected asset register populates a non-blank verified custodian value (not merely a column header). | weight: 5
3. For reviewed assets, the corrected asset register populates a non-blank verified location value. | weight: 3
4. For reviewed assets, the corrected asset register populates a non-blank verified lifecycle status value. | weight: 3
5. For reviewed assets, the corrected asset register populates a non-blank evidence source, and that text is consistent with the fixed-asset ledger extract: it must not claim there is no FA/ledger row for an asset tag that appears on the ledger; for MD-00130, MD-00131, and MD-00132 it may note ledger absence (with the under-threshold vs capital-qualifying distinction addressed elsewhere). | weight: 3
6. For reviewed assets, the corrected asset register populates a non-blank last verified date. | weight: 3
7. Each reviewed asset row has a non-blank numeric/currency acquisition cost matching source evidence, and the sum of corrected-register acquisition costs for reviewed assets equals $332,115 (±$1). | weight: 3
8. Where a ledger NBV exists, remaining book value is non-blank and matches the ledger figure (blank only when no ledger row exists); MD-00034/FA-000034 remaining book value equals $479.88 (acquisition $930 − accumulated depreciation $450.12), and aggregate remaining book value across reviewed assets equals $61,526.60 (±$0.05). | weight: 3
9. Every reviewed asset has a non-blank confidence value drawn only from the closed set {High, Medium, Low}. | weight: 3
10. States the quarterly inventory is NOT APPROVED (or equivalent hold) while Critical certification blockers remain; does not give an unconditional approval. | weight: 3
11. Includes a certification/sign-off page with sign-off lines for the IT Operations Manager, Finance Controller, HR Operations Lead, and Internal Auditor. | weight: 4
12. The certification page distinguishes what can be accepted (or accepted with conditions), what must escalate, and what blocks signing (Critical certification blockers). | weight: 4
13. The deliverable documents a precedence order consistent with policy (ledger / verified disposal certificates / carrier acceptance+delivery+receiving-scan over approved transfers over HR over ticket status over technician notes and dock/exception images) and shows at least one rejected lower-precedence claim with the higher-precedence record that overruled it. | weight: 4
14. Marks MD-00082 as In Transit - Exception / shipment mismatch / equivalent unresolved custody, citing tracking 1ZMD00000082 and serial mismatch (MISMATCH-0082 vs MD-MO-050082), and does not clear it to Available. | weight: 4
15. Classifies MD-00089 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
16. Classifies MD-00090 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
17. Classifies MD-00091 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
18. Classifies MD-00092 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
19. Classifies MD-00093 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
20. Classifies MD-00094 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
21. Classifies MD-00095 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
22. Classifies MD-00096 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
23. Classifies MD-00097 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
24. Classifies MD-00098 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
25. Classifies MD-00099 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
26. Classifies MD-00100 as Missing / Cannot Locate (or equivalent), not as verified Available/In Use stock. | weight: 1
27. Flags duplicate serial MD-LA-050021 across MD-00021 and MD-00025, and duplicate serial MD-NE-050088 across MD-00088 and MD-00092. | weight: 4
28. Flags missing approval on transfer TR-00058. | weight: 1
29. Flags missing approval on transfer TR-00059. | weight: 1
30. Flags missing approval on transfer TR-00064. | weight: 1
31. Flags missing approval on transfer TR-00065. | weight: 1
32. Flags missing approval on transfer TR-00127. | weight: 1
33. Rejects regional note RN-0132 (or equivalent note claiming MD-00132 was expensed under threshold) as non-authoritative; does not clear MD-00132's missing FA row on that basis. | weight: 4
34. Each exception row populates at least one related-record citation drawn from: employee ID, ticket ID, transfer ID, PO number, tracking number, certificate ID, or ledger asset ID (non-blank field value, not header-only). | weight: 3
35. Each exception row populates a financial exposure value calculated from acquisition costs and book values (non-blank, not header-only). | weight: 3
36. Each exception row populates a required action. | weight: 3
37. Each exception row populates an owner by role. | weight: 3
38. Each unresolved exception row for a Critical certification blocker includes a populated resolution action deadline (specific date by which the escalation owner must complete the required action). | weight: 3
39. Provides custody-chain detail for MD-00068, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events. | weight: 2
40. Provides custody-chain detail for MD-00074, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events. | weight: 2
41. Provides custody-chain detail for MD-00076, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events. | weight: 2
42. Provides custody-chain detail for MD-00082, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events. | weight: 2
43. Provides custody-chain detail for MD-00084, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events. | weight: 2
44. Provides custody-chain detail for MD-00114, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events. | weight: 2
45. Provides custody-chain detail for MD-00118, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events. | weight: 2
46. Provides custody-chain detail for MD-00132, covering purchase/assignment/transfer and later offboarding, shipment, receipt, loss, or disposal events. | weight: 2
47. Summary dashboard includes asset counts and book value broken out by verified status. | weight: 3
48. Summary dashboard includes asset counts and book value broken out by device category with correct counts from the corrected register: Laptop 40, Mobile Device 36, Monitor 28, Network Asset 28 (counts ±0). | weight: 3
49. Summary dashboard includes a breakout by exception type (count and financial exposure or book value). | weight: 3
50. Reconciles corrected inventory to the ledger and explains remaining differences with an explicit distinction between expected under-threshold FA absence (MD-00130/MD-00131) and Critical capital-qualifying absence (MD-00132), plus cost-basis mismatches and disposal-status disagreements. | weight: 5
51. Treats overdue / label-only returns that lack carrier acceptance + delivery + receiving-scan proof as Critical certification blockers (for example MD-00068). | weight: 4
52. Treats unresolved shipment serial mismatches as Critical certification blockers for at least 3 of these 4 assets: MD-00074, MD-00076, MD-00082, and MD-00084. | weight: 4
53. Treats retired/pending-disposal assets lacking a verified disposal certificate as Critical certification blockers (MD-00114 and MD-00118). | weight: 4
54. Treats capital-qualifying assets (acquisition cost ≥ $2,500) that are missing from the fixed-asset ledger as Critical certification blockers. | weight: 4
55. Approves/certifies the quarterly inventory as complete despite Critical certification blockers. | weight: -5
56. Marks Label Created shipments as Available/returned/completed without carrier acceptance + delivery + receiving scan evidence. | weight: -5
57. Uses receiving_exception_scan_1ZMD00000082.png as proof MD-00082 was successfully received/cleared. | weight: -5
58. Cites or uses transfer IDs, return receipts, approval IDs, or disposal certificates that cannot be traced to a source record in the provided files. | weight: -5
59. Treats unresolved MD-00089–MD-00100 assets as verified In Use or Available stock. | weight: -4
60. Treats MD-00130 and/or MD-00131 ledger absence as a Critical certification blocker, or uses a capitalization_approved=Yes purchase-order flag to force FA expectations for those under-threshold lines. | weight: -5
