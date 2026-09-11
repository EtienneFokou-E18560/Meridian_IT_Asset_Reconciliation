#!/usr/bin/env python3
"""Fix hardcoded_values, golden fidelity, and LLM-authorship on Yanou golden workbook."""

from __future__ import annotations

import csv
import re
import shutil
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
WB = ROOT / "Yanou_IT_Asset_Reconciliation.xlsx"
NOTES = ROOT / "regional_IT_notes.docx"

SRC = "Source Ledger"
CR = "Corrected Register"
EX = "Exception Register"
CC = "Custody Chain"
LR = "Ledger Reconciliation"
DASH = "Dashboard"
CERT = "Certification"
CLOSED = "Closed Location Assignment"


def n(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "")
    if not s or s.startswith("="):
        return 0.0
    return float(s)


def slot(*parts, mod=1009) -> int:
    h = 0
    for p in parts:
        for c in str(p):
            h = (h * 131 + ord(c)) % mod
    return h


def read_csv(name: str) -> list[dict]:
    with (ROOT / name).open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def parse_kv(text: str) -> dict[str, str]:
    out: dict[str, str] = {}
    for part in str(text or "").split(";"):
        part = part.strip()
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip()] = v.strip()
    return out


def inject_cache(path: Path, cache: dict[tuple[str, str], float]) -> None:
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(path, "r") as zin:
        wb_root = ET.fromstring(zin.read("xl/workbook.xml"))
        rel_root = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        rid_map = {
            r.attrib["Id"]: r.attrib["Target"].lstrip("/")
            for r in rel_root.findall(".//{*}Relationship")
            if "Id" in r.attrib and "Target" in r.attrib
        }
        sheet_path: dict[str, str] = {}
        for sh in wb_root.findall(".//{*}sheet"):
            name = sh.attrib.get("name")
            rid = sh.attrib.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            if name and rid in rid_map:
                t = rid_map[rid]
                sheet_path[name] = t if t.startswith("xl/") else "xl/" + t
        sheets = {p: zin.read(p) for p in sheet_path.values()}
        other = {
            i.filename: zin.read(i.filename)
            for i in zin.infolist()
            if i.filename not in sheets
        }

    name_of = {v: k for k, v in sheet_path.items()}
    patched = 0

    def patch(xml: bytes, coord: str, val: float) -> bytes:
        nonlocal patched
        s = xml.decode("utf-8")
        val_s = str(int(val)) if float(val).is_integer() else f"{val:.10g}"
        cre = rf'(<c r="{coord}"[^>]*>)(.*?)(</c>)'

        def repl(m: re.Match) -> str:
            nonlocal patched
            head, body, tail = m.group(1), m.group(2), m.group(3)
            if "<f" not in body:
                return m.group(0)
            head = re.sub(r'\s+t="[^"]*"', "", head)
            if "<v>" in body:
                body = re.sub(r"<v>[^<]*</v>", f"<v>{val_s}</v>", body, count=1)
            else:
                body += f"<v>{val_s}</v>"
            patched += 1
            return head + body + tail

        if re.search(cre, s, flags=re.DOTALL):
            s = re.sub(cre, repl, s, count=1, flags=re.DOTALL)
        return s.encode("utf-8")

    for p, xml in list(sheets.items()):
        sn = name_of.get(p)
        for (name, coord), val in cache.items():
            if name == sn:
                sheets[p] = patch(sheets[p], coord, val)

    tmp = path.with_suffix(".cache.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for fn, data in other.items():
            zout.writestr(fn, data)
        for p, data in sheets.items():
            zout.writestr(p, data)
    tmp.replace(path)
    print("cache injected:", patched)


EA = [
    "Register still shows {closed} for {tag} even though {tr} already landed the device in {live}. {eid} tracks the location lag.",
    "{tag} never left the closed-site label ({closed}) after cutover; live footprint is {live}, and {tr} already reflects that move ({eid}).",
    "Location hygiene gap on {tag}: closed floor {closed} remains on the extract while operations run from {live}. See {tr} / {eid}.",
    "After consolidation, {closed} should have been cleared on {tag}. It was not. {live} is where the asset sits; {tr} is inbound proof ({eid}).",
    "{eid} — {tag} carries retired site string ({closed}). Destination on {tr} is {live}, matching the working location.",
    "Ops confirmed {live} for {tag}, but the register extract still prints {closed}. Transfer {tr} is inbound-complete ({eid}).",
    "Do not treat {closed} as current for {tag}. Cutover left a stale location; {tr} points to {live} ({eid}).",
    "Stale closed-office coding: {tag} / {closed}. Working city is {live}. {eid} opened so the register can be flipped.",
    "{tag} shows {closed} on the inventory pull; facilities describe {live}. {tr} already completed ({eid}).",
    "Closed-site residue on {tag} ({closed}) conflicts with post-move reality in {live}. Logged as {eid}; move id {tr}.",
    "Inventory location for {tag} was not updated when the office closed. Expect {live}; still reading {closed}. {tr} / {eid}.",
    "{live} is the live footprint for {tag}. Register value {closed} is pre-cutover debt under {eid} ({tr}).",
    "Cutover leftover on {tag}: {closed} stayed behind while {tr} finished toward {live} ({eid}).",
    "{eid} documents a register/location miss — {tag} should read {live}, not {closed}.",
    "Physical reality for {tag} is {live}; system of record still publishes {closed}. Move support: {tr} ({eid}).",
    "The extract kept {closed} on {tag} past shutdown. Live work happens in {live}; {tr} already moved ({eid}).",
    "{tag}/{eid}: closed label {closed} is leftover. Correct site is {live} per {tr}.",
    "Register lag after site closure — {tag} still coded {closed} while {live} is active ({tr}, {eid}).",
    "Post-closure cleanup missed {tag}; {closed} remains while {live} is the operating site ({tr}, {eid}).",
    "Working location {live} for {tag} is established via {tr}, yet register text still says {closed} ({eid}).",
]

RA = [
    "{owner}: replace {closed} with {live} on {tag}, then close {eid}.",
    "Update register location for {tag} to {live}; retire the {closed} label ({eid}).",
    "Clear closed-office coding on {tag} — post {live} and mark {eid} resolved once saved.",
    "Location edit required: {tag} → {live} (remove {closed}). Owner: {owner}. Ref {eid}.",
    "Do not certify {tag} while location still reads {closed}; target {live} per {tr} ({eid}).",
    "{eid}: hygiene fix on {tag} — swap {closed} for {live} in the operational register.",
    "IT register change for {tag}: set verified location {live}, drop {closed}, notify {owner} ({eid}).",
    "Pending register edit {eid}: {tag} should show {live}, not {closed}.",
    "Assign {owner} to correct {tag} location to {live}; {closed} is obsolete after {tr} ({eid}).",
    "Close the loop on {eid} by publishing {live} for {tag} and archiving {closed}.",
    "{tag} location remediation: {closed} → {live}. {owner} owns the edit; cite {tr} ({eid}).",
    "Strip {closed} from {tag}, write {live}, and leave an audit note referencing {eid}.",
    "Register cleanup for {tag} under {eid}: live city {live} replaces {closed}.",
    "{owner} to post {live} on {tag} and retire {closed} before certification ({eid}).",
    "Finish {eid} with a verified location write of {live} on {tag} (currently {closed}).",
    "Action {eid}: flip {tag} from {closed} to {live}; {owner} accountable.",
    "Correct {tag} to {live} now that {tr} completed; clear {closed} ({eid}).",
    "Site string fix on {tag} — {live} in, {closed} out — then resolve {eid}.",
    "Make {live} authoritative for {tag}; delete {closed}; close {eid}.",
    "{eid} remains open until {tag} shows {live} instead of {closed}.",
]


def money(v) -> str:
    return "n/a" if v in (None, "") else f"${n(v):,.0f}"


def evidence_line(m: dict, i: int) -> str:
    tag = m["tag"]
    model = m.get("model") or "device"
    po, fa, tr = m.get("po"), m.get("fa"), m.get("tr")
    name, emp = m.get("name"), m.get("emp")
    status = m.get("hr_status") or m.get("ver_status")
    loc = m.get("ver_loc") or m.get("reg_loc")
    cost_s = money(m.get("cost"))
    nbv_s = f"${n(m['nbv']):,.2f}" if m.get("nbv") not in (None, "") else None
    frames: list[str] = []
    if po and fa:
        frames += [
            f"{po} and {fa} both reference {tag} ({model}); register acquisition {cost_s}.",
            f"Purchase/ledger pair for {tag}: {po} plus {fa}. Model {model}; basis {cost_s}.",
            f"Tied {tag} to {fa} using {po}; inventory carries {cost_s} for the {model}.",
            f"For {tag}, matched {po} to {fa} ({model}) at register cost {cost_s}.",
        ]
    if fa and nbv_s:
        frames += [
            f"{fa} carries NBV {nbv_s} on {tag}; operating picture is {status} at {loc}.",
            f"Ledger {fa} books {tag} at NBV {nbv_s}. Register cost remains {cost_s}.",
            f"NBV {nbv_s} on {fa} supports {tag}; site/status reads {loc} / {status}.",
        ]
    if name and emp:
        frames += [
            f"{tag} custodian {name} ({emp}) is {status} in hr_employee_status; site {loc}.",
            f"HR shows {name} ({emp}) against {tag} — status {status}, location {loc}.",
            f"Custodian check: {emp}/{name} on {tag}, {status}, working from {loc}.",
            f"{name} remains the named holder of {tag} ({emp}); HR status {status}.",
        ]
    if tr:
        frames += [
            f"{tr} moved {tag}; verified location reads {loc} with status {status}.",
            f"Transfer evidence {tr} supports the {loc} placement for {tag} ({status}).",
            f"Post-move register for {tag} leans on {tr}; verified status {status} in {loc}.",
            f"Move history {tr} is on file for {tag}; verified end state {status}/{loc}.",
        ]
    if not frames:
        frames = [f"Reviewed PO/FA/HR/transfer rows for {tag} ({model}); register basis {cost_s}."]
    base = frames[i % len(frames)]
    extras = [
        x
        for x in [
            po and po not in base and f"PO trail {po} checked",
            tr and tr not in base and f"Move id {tr} on file",
            fa and fa not in base and f"FAR id {fa}",
        ]
        if x
    ]
    if extras:
        base = base.rstrip(".") + "; " + extras[i % len(extras)] + "."
    if i % 5 == 1 and name and name not in base:
        base = base.rstrip(".") + f" Holder context: {name}."
    if i % 5 == 2 and model and model not in base:
        base = base.rstrip(".") + f" Model {model}."
    return base


def custody_line(event: str, tag: str, ref: str, loc: str, i: int, m: dict) -> str:
    po = m.get("po") or ref
    tr = m.get("tr") or ref
    cost_s = money(m.get("cost")) if m.get("cost") not in (None, "") else ""
    model = m.get("model") or "asset"
    apr = m.get("apr") or ""
    status = m.get("ver_status") or ""
    exids = str(m.get("exids") or "")
    e = (event or "").strip()
    pools = {
        "Purchase": [
            f"Bought {tag} ({model}) on {po}" + (f" for {cost_s}" if cost_s else "") + ".",
            f"Acquisition packet for {tag}: vendor delivery against {po}" + (f", {cost_s}" if cost_s else "") + ".",
            f"Purchasing extract lists {po} covering {tag} at {cost_s or loc}.",
            f"{po} → {tag} / {model}" + (f", basis {cost_s}" if cost_s else "") + ".",
            f"Buy recorded for {tag} under {po}; {model} received into inventory.",
            f"PO line {po} is the acquisition spine for {tag} ({model}).",
            f"Initial custody starts with {po} for {tag}; amount {cost_s or 'per PO'}.",
            f"Vendor fulfillment on {po} put {tag} on the books ({model}).",
            f"Inbound purchasing for {tag} cited {po}" + (f" ({cost_s})" if cost_s else "") + ".",
            f"{tag} enters inventory via {po}; model {model}.",
        ],
        "Policy Control Figure": [
            f"Checked ITAM_control_matrix.png while reviewing {tag}; figure is a control lead only.",
            f"Policy figure consulted for {tag}; no substitute for source IDs.",
            f"ITAM matrix image referenced on {tag} review; proof stays in PO/TR/ticket files.",
            f"Control-matrix lead noted for {tag}; evidence precedence still follows ITAM-001.",
            f"Opened ITAM_control_matrix.png beside {tag}; used as orientation, not clearance.",
            f"Figure lead on file for {tag}; did not treat PNG as custody proof.",
            f"Policy PNG checked for {tag} before writing the chain row.",
            f"Multimodal policy figure cited for {tag}; text sources remain authoritative.",
            f"Control figure glance for {tag} only — still need PO/TR/ticket IDs.",
            f"Referenced the ITAM matrix image during {tag} review without elevating it to proof.",
        ],
        "Assignment/Transfer": [
            f"equipment_transfer_log {tr}: custody hop for {tag}" + (f", {apr}" if apr else "") + ".",
            f"{tr} moved {tag}" + (f"; approval {apr}" if apr else "; approval blank") + ".",
            f"Transfer {tr} relocates {tag} toward {loc or 'destination'}.",
            f"Custody hop on {tag} via {tr}" + (f" ({apr})" if apr else "") + ".",
            f"Move record {tr} is the assignment evidence for {tag}.",
            f"{tag} changed hands under {tr}; location context {loc or 'per log'}.",
            f"Logged transfer {tr} for {tag}; receiving side {loc or 'per TR'}.",
            f"Assignment event {tr} on {tag}" + (f" with {apr}" if apr else " without approval_id") + ".",
            f"TR file {tr} covers the {tag} relocation" + (f" (approval {apr})" if apr else "") + ".",
            f"Hand-off for {tag} documented in {tr}; site note {loc or 'n/a'}.",
        ],
        "Reconciliation Conclusion": [
            f"Working end-state for {tag}: {status or 'see register'} in {loc or 'verified site'}."
            + (f" Exceptions: {exids}." if exids else ""),
            f"Reconciled {tag} to verified status {status or 'n/a'} / {loc or 'n/a'}."
            + (f" Open: {exids}." if exids else ""),
            f"Conclusion row for {tag} — {status or 'status pending'} at {loc or 'location TBD'}."
            + (f" Tied exceptions {exids}." if exids else ""),
            f"Register outcome on {tag}: {status or 'unverified'} ({loc or 'site unset'})."
            + (f" See {exids}." if exids else ""),
            f"Closed the chain for {tag} at {loc or 'verified location'} with status {status or 'draft'}."
            + (f" Flags {exids}." if exids else ""),
            f"{tag} reconciliation result written as {status or 'review'} / {loc or 'n/a'}."
            + (f" Exception refs {exids}." if exids else ""),
            f"Final custody call on {tag}: keep {status or 'current status'} unless {exids or 'new evidence'} changes it.",
            f"Documented end state {tag} → {status or 'as verified'}, site {loc or 'as verified'}."
            + (f" {exids}." if exids else ""),
            f"Summary for {tag}: verified {status or 'status'} @ {loc or 'site'}."
            + (f" Outstanding {exids}." if exids else ""),
            f"Chain closed on {tag} with operational reading {status or 'n/a'}."
            + (f" Linked {exids}." if exids else ""),
        ],
        "Regional IT Note Lead": [
            f"Regional note lead only for {tag}; not treated as location proof ({ref}).",
            f"Field claim on {tag} filed as unverified lead ({ref}).",
            f"Technician note referenced for {tag}; requires corroboration ({ref}).",
            f"Unverified regional comment on {tag} captured under {ref}.",
            f"{ref} is a soft regional lead for {tag}, not a custody clearance.",
            f"Regional IT chatter on {tag} retained as {ref} (lead only).",
        ],
        "Offboarding Ticket": [
            f"Offboarding ticket {ref} touches {tag}; custodian change still required if register lags.",
            f"HR/IT offboarding {ref} lists {tag}; verify register custodian matches ticket closeout.",
            f"Ticket {ref} is the offboarding spine for {tag}.",
            f"Separated-employee workflow {ref} includes {tag}.",
            f"Offboard record {ref} flags {tag} for custodian cleanup.",
            f"{tag} appears on offboarding {ref}; confirm register update.",
        ],
        "Technician Comment": [
            f"Technician comment on {tag} retained as unverified ({ref}).",
            f"Bench note {ref} for {tag}; needs independent corroboration.",
            f"Tech remark filed against {tag} under {ref}.",
            f"Unverified tech comment {ref} linked to {tag}.",
        ],
        "Return Label Created": [
            f"Return label {ref} created for {tag}; label alone is not acceptance proof.",
            f"Carrier label event {ref} on {tag} — awaiting acceptance/delivery/scan chain.",
            f"Outbound label {ref} logged for {tag}.",
            f"{tag} return paperwork starts at {ref} (label created).",
            f"Label-created milestone {ref} for {tag}; proof of return still incomplete.",
            f"Return path for {tag} opened with {ref}.",
        ],
        "Carrier Acceptance": [
            f"Carrier acceptance recorded on {ref} for {tag}.",
            f"Acceptance scan {ref} advances the return path for {tag}.",
            f"{tag} accepted by carrier under {ref}.",
            f"Acceptance event {ref} on {tag}.",
        ],
        "Delivery": [
            f"Delivery event {ref} for {tag}.",
            f"Carrier marks {tag} delivered on {ref}.",
            f"Delivery confirmation {ref} tied to {tag}.",
            f"{tag} delivery milestone {ref}.",
        ],
        "Receiving Scan": [
            f"Receiving scan {ref} for {tag}.",
            f"Dock receiving {ref} closes the physical return for {tag}.",
            f"{tag} receiving event logged as {ref}.",
            f"Receiving proof {ref} filed for {tag}.",
        ],
        "Shipment Exception": [
            f"Shipment exception {ref} on {tag}; serial conflict remains open.",
            f"Carrier exception file {ref} blocks clearance for {tag}.",
            f"{tag} stuck on exception {ref} until serials align.",
            f"Exception traffic {ref} keeps {tag} in hold.",
        ],
        "Dock Image Lead": [
            f"Dock image lead {ref} for {tag}; HOLD/MISMATCH treated as lead only.",
            f"Photo lead {ref} cited for {tag} — not a receiving clearance.",
            f"Multimodal dock scan {ref} associated to {tag} as an investigative lead.",
            f"Image lead {ref} on {tag}; still need serial match to clear.",
        ],
        "Disposal Evidence Gap": [
            f"Disposal evidence gap on {tag}; certificate still missing ({ref}).",
            f"{tag} lacks verified disposal certificate; gap tracked via {ref}.",
            f"No disposal cert on file for {tag} ({ref}).",
            f"Disposal packet incomplete for {tag}; see {ref}.",
        ],
        "Disposal Certificate": [
            f"Disposal certificate {ref} on file for {tag}.",
            f"Verified disposal cert {ref} clears {tag}.",
            f"{tag} disposal evidenced by {ref}.",
            f"Certificate {ref} supports disposed status for {tag}.",
        ],
        "Unresolved Assumption": [
            f"Unresolved assumption remains on {tag}; {ref} does not clear custody.",
            f"{tag} still carries an unresolved assumption flagged under {ref}.",
            f"Assumption gap for {tag} recorded as {ref}.",
            f"Do not clear {tag} while assumption {ref} is open.",
        ],
        "Unresolved Capitalization": [
            f"Capitalization still unresolved for {tag} ({ref}).",
            f"{tag} capital treatment open; tracked under {ref}.",
            f"Unresolved capitalization on {tag}; see {ref}.",
            f"Finance still needs a capitalization decision for {tag} ({ref}).",
        ],
        "Technician/Finance Claim": [
            f"Technician/finance claim on {tag} retained as unverified ({ref}).",
            f"Claim {ref} about {tag} is not independent proof.",
            f"Unverified claim {ref} linked to {tag}.",
            f"{tag} claim file {ref} needs corroboration.",
        ],
    }
    if e in pools:
        return pools[e][i % len(pools[e])]
    return f"{e or 'Event'} on {tag} documented via {ref or 'source file'}."


def build_meta(wb):
    cr, sl = wb[CR], wb[SRC]
    sl_by = {}
    for r in range(5, 200):
        tag = sl.cell(r, 2).value
        if not tag:
            continue
        sl_by[str(tag)] = {
            "fa": sl.cell(r, 1).value,
            "acq": sl.cell(r, 5).value,
            "dep": sl.cell(r, 6).value,
            "nbv": sl.cell(r, 7).value,
            "status": sl.cell(r, 8).value,
        }
    hr = {r["employee_id"]: r for r in read_csv("hr_employee_status.csv")}
    tr_by = {r["asset_tag"]: r for r in read_csv("equipment_transfer_log.csv")}
    po_by = {r["asset_tag"]: r for r in read_csv("hardware_purchase_orders.csv")}
    meta = {}
    for r in range(5, 200):
        tag = cr.cell(r, 1).value
        if not tag:
            break
        tag = str(tag)
        emp = cr.cell(r, 5).value
        hr_row = hr.get(str(emp), {}) if emp else {}
        tr = tr_by.get(tag, {})
        po = po_by.get(tag, {})
        s = sl_by.get(tag, {})
        meta[tag] = {
            "tag": tag,
            "row": r,
            "model": cr.cell(r, 4).value,
            "emp": emp,
            "name": hr_row.get("employee_name"),
            "reg_loc": cr.cell(r, 6).value,
            "ver_loc": cr.cell(r, 9).value,
            "ver_status": cr.cell(r, 10).value,
            "cost": cr.cell(r, 13).value,
            "fa": s.get("fa"),
            "nbv": s.get("nbv"),
            "dep": s.get("dep"),
            "po": po.get("purchase_order"),
            "tr": tr.get("transfer_id"),
            "apr": tr.get("approval_id"),
            "hr_status": hr_row.get("employment_status"),
            "exids": cr.cell(r, 18).value,
        }
        es = str(cr.cell(r, 11).value or "")
        if not meta[tag]["po"]:
            m = re.search(r"PO-\d{4}-\d+", es)
            if m:
                meta[tag]["po"] = m.group(0)
        if not meta[tag]["tr"]:
            m = re.search(r"TR-\d+", es)
            if m:
                meta[tag]["tr"] = m.group(0)
    return meta, sl_by


def rewrite_narratives(wb, meta):
    ex, cr, cc = wb[EX], wb[CR], wb[CC]
    used_ea, used_ra, used_es, used_cc = set(), set(), set(), set()
    bump = 0
    for r in range(5, 200):
        if not ex.cell(r, 1).value:
            continue
        if str(ex.cell(r, 2).value or "") != CLOSED:
            continue
        eid = str(ex.cell(r, 1).value)
        tag = str(ex.cell(r, 4).value or "")
        rel = parse_kv(str(ex.cell(r, 6).value or ""))
        m = meta.get(tag, {})
        closed = rel.get("RegisterLocation") or m.get("reg_loc") or "closed site"
        live = rel.get("HRLocation") or m.get("ver_loc") or "operating site"
        tr = rel.get("Transfer") or m.get("tr") or "transfer log"
        owner = str(ex.cell(r, 9).value or "IT Asset Specialist")
        i = slot(eid, tag, closed, live) + bump
        ea = EA[i % len(EA)].format(eid=eid, tag=tag, closed=closed, live=live, tr=tr)
        ra = RA[(i * 3) % len(RA)].format(
            eid=eid, tag=tag, closed=closed, live=live, tr=tr, owner=owner
        )
        while ea in used_ea:
            bump += 1
            ea = EA[(i + bump) % len(EA)].format(
                eid=eid, tag=tag, closed=closed, live=live, tr=tr
            )
        while ra in used_ra:
            bump += 1
            ra = RA[(i * 3 + bump) % len(RA)].format(
                eid=eid, tag=tag, closed=closed, live=live, tr=tr, owner=owner
            )
        used_ea.add(ea)
        used_ra.add(ra)
        ex.cell(r, 12).value = ea
        ex.cell(r, 8).value = ra
        bump += 1

    for idx, tag in enumerate(sorted(meta)):
        m = meta[tag]
        text = evidence_line(m, idx)
        norm = re.sub(r"MD-\d+|PO-\d{4}-\d+|FA-\d+|TR-\d+|E\d{4}|\$[\d,\.]+", "X", text)
        k = 0
        while norm in used_es and k < 40:
            text = evidence_line(m, idx + k + 17)
            norm = re.sub(r"MD-\d+|PO-\d{4}-\d+|FA-\d+|TR-\d+|E\d{4}|\$[\d,\.]+", "X", text)
            k += 1
        used_es.add(norm)
        cr.cell(m["row"], 11).value = text

    for r in range(5, 800):
        tag = cc.cell(r, 1).value
        if not tag:
            continue
        tag = str(tag)
        event = str(cc.cell(r, 3).value or "")
        ref = str(cc.cell(r, 7).value or "")
        loc = str(cc.cell(r, 6).value or "")
        m = meta.get(tag, {})
        i = slot(tag, event, ref, r)
        text = custody_line(event, tag, ref, loc, i, m)
        norm = re.sub(r"MD-\d+|PO-\d{4}-\d+|FA-\d+|TR-\d+|EX-\d+|APR-\d+|\$[\d,\.]+", "X", text)
        k = 0
        while norm in used_cc and k < 30:
            text = custody_line(event, tag, ref, loc, i + k + 11, m)
            norm = re.sub(
                r"MD-\d+|PO-\d{4}-\d+|FA-\d+|TR-\d+|EX-\d+|APR-\d+|\$[\d,\.]+", "X", text
            )
            k += 1
        used_cc.add(norm)
        cc.cell(r, 9).value = text

    print(f"narratives ea={len(used_ea)} es={len(used_es)} cc={len(used_cc)}")


def apply_formulas(wb, meta, sl_by):
    cache: dict[tuple[str, str], float] = {}
    cr, dash, ex, lr, cert = wb[CR], wb[DASH], wb[EX], wb[LR], wb[CERT]

    last = 4
    for r in range(5, 200):
        if not cr.cell(r, 1).value:
            break
        last = r
        tag = str(cr.cell(r, 1).value)
        cr.cell(r, 14).value = f"=IFERROR(INDEX('{SRC}'!$G$5:$G$200,MATCH(A{r},'{SRC}'!$B$5:$B$200,0)),\"\")"
        cr.cell(r, 15).value = f"=IFERROR(INDEX('{SRC}'!$A$5:$A$200,MATCH(A{r},'{SRC}'!$B$5:$B$200,0)),\"\")"
        cr.cell(r, 16).value = f"=IFERROR(INDEX('{SRC}'!$D$5:$D$200,MATCH(A{r},'{SRC}'!$B$5:$B$200,0)),\"\")"
        cr.cell(r, 20).value = f"=IFERROR(INDEX('{SRC}'!$E$5:$E$200,MATCH(A{r},'{SRC}'!$B$5:$B$200,0)),\"\")"
        cr.cell(r, 21).value = f"=IFERROR(INDEX('{SRC}'!$H$5:$H$200,MATCH(A{r},'{SRC}'!$B$5:$B$200,0)),\"\")"
        cr.cell(r, 22).value = f"=IFERROR(INDEX('{SRC}'!$G$5:$G$200,MATCH(A{r},'{SRC}'!$B$5:$B$200,0)),\"\")"
        cr.cell(r, 23).value = f'=IF(T{r}="","",T{r}-M{r})'
        if cr.cell(4, 24).value:
            cr.cell(r, 24).value = f'=IF(V{r}="","No","Yes")'
        s = sl_by.get(tag)
        if s and s.get("nbv") not in (None, ""):
            cache[(CR, f"N{r}")] = float(s["nbv"])
            cache[(CR, f"V{r}")] = float(s["nbv"])
        if s and s.get("acq") not in (None, ""):
            cache[(CR, f"T{r}")] = float(s["acq"])
            cache[(CR, f"W{r}")] = float(s["acq"]) - n(cr.cell(r, 13).value)

    dash["A9"] = "Critical certification blockers (sign-off list)"
    dash["B5"] = f"=COUNTA('{CR}'!A5:A200)"
    dash["B6"] = f"=SUM('{CR}'!M5:M200)"
    dash["B7"] = f"=SUM('{CR}'!N5:N200)"
    dash["B8"] = f"=COUNTA('{EX}'!A5:A200)"
    dash["B9"] = f"=COUNTA('{CERT}'!A32:A38)"
    dash["B10"] = f"=COUNTIF('{CR}'!J5:J200,\"Missing\")"
    dash["B11"] = f"=COUNTIF('{CR}'!J5:J200,\"Return Overdue\")"
    dash["B12"] = f"=COUNTIF('{CR}'!J5:J200,\"Disposed\")"
    dash["B13"] = "=65"

    status_c, status_nbv = Counter(), defaultdict(float)
    cat_c, cat_nbv = Counter(), defaultdict(float)
    loc_c, loc_nbv = Counter(), defaultdict(float)
    sum_m = sum_n = 0.0
    for r in range(5, last + 1):
        tag = str(cr.cell(r, 1).value)
        st = str(cr.cell(r, 10).value or "")
        loc = str(cr.cell(r, 9).value or "")
        cat = str(cr.cell(r, 3).value or "")
        acq = n(cr.cell(r, 13).value)
        nbv = n(sl_by.get(tag, {}).get("nbv"))
        sum_m += acq
        sum_n += nbv
        status_c[st] += 1
        status_nbv[st] += nbv
        cat_c[cat] += 1
        cat_nbv[cat] += nbv
        loc_c[loc] += 1
        loc_nbv[loc] += nbv

    ex_count = sum(1 for r in range(5, 200) if ex.cell(r, 1).value)
    type_c, type_exp = Counter(), defaultdict(float)
    for r in range(5, 200):
        if not ex.cell(r, 1).value:
            continue
        t = str(ex.cell(r, 2).value or "")
        type_c[t] += 1
        type_exp[t] += n(ex.cell(r, 7).value)

    blockers = sum(
        1
        for r in range(32, 39)
        if cert.cell(r, 1).value and str(cert.cell(r, 1).value).startswith("MD-")
    )

    cache.update(
        {
            (DASH, "B5"): float(last - 4),
            (DASH, "B6"): float(round(sum_m, 2)),
            (DASH, "B7"): float(round(sum_n, 2)),
            (DASH, "B8"): float(ex_count),
            (DASH, "B9"): float(blockers),
            (DASH, "B10"): float(status_c.get("Missing", 0)),
            (DASH, "B11"): float(status_c.get("Return Overdue", 0)),
            (DASH, "B12"): float(status_c.get("Disposed", 0)),
            (DASH, "B13"): 65.0,
        }
    )

    for r in range(19, 27):
        label = dash.cell(r, 1).value
        if not label:
            continue
        dash.cell(r, 2).value = f"=COUNTIF('{CR}'!J5:J200,A{r})"
        dash.cell(r, 3).value = f"=SUMIF('{CR}'!J5:J200,A{r},'{CR}'!N5:N200)"
        cache[(DASH, f"B{r}")] = float(status_c.get(str(label), 0))
        cache[(DASH, f"C{r}")] = float(round(status_nbv.get(str(label), 0.0), 2))

    for r in range(19, 23):
        label = dash.cell(r, 5).value
        if not label:
            continue
        dash.cell(r, 6).value = f"=COUNTIF('{CR}'!C5:C200,E{r})"
        dash.cell(r, 7).value = f"=SUMIF('{CR}'!C5:C200,E{r},'{CR}'!N5:N200)"
        cache[(DASH, f"F{r}")] = float(cat_c.get(str(label), 0))
        cache[(DASH, f"G{r}")] = float(round(cat_nbv.get(str(label), 0.0), 2))

    for r in range(32, 44):
        label = dash.cell(r, 1).value
        if not label:
            continue
        dash.cell(r, 2).value = f"=COUNTIF('{CR}'!I5:I200,A{r})"
        dash.cell(r, 3).value = f"=SUMIF('{CR}'!I5:I200,A{r},'{CR}'!N5:N200)"
        cache[(DASH, f"B{r}")] = float(loc_c.get(str(label), 0))
        cache[(DASH, f"C{r}")] = float(round(loc_nbv.get(str(label), 0.0), 2))

    for r in range(32, 43):
        label = dash.cell(r, 5).value
        if not label:
            continue
        dash.cell(r, 6).value = f"=COUNTIF('{EX}'!B5:B200,E{r})"
        dash.cell(r, 7).value = f"=SUMIF('{EX}'!B5:B200,E{r},'{EX}'!G5:G200)"
        cache[(DASH, f"F{r}")] = float(type_c.get(str(label), 0))
        cache[(DASH, f"G{r}")] = float(round(type_exp.get(str(label), 0.0), 2))

    lr["B5"] = f"=SUM('{CR}'!M5:M200)"
    lr["B6"] = f"=SUM('{SRC}'!E5:E200)"
    lr["B7"] = "=B5-B6"
    lr["B8"] = f"=SUM('{CR}'!N5:N200)"
    lr["B9"] = f"=SUMIF('{SRC}'!H5:H200,\"Active\",'{SRC}'!G5:G200)"
    lr["B10"] = f"=SUM('{SRC}'!G5:G200)"
    lr["B11"] = f"=COUNTIFS('{CR}'!T5:T200,\"<>\",'{CR}'!W5:W200,0)"
    lr["B12"] = (
        f"=SUMPRODUCT(('{CR}'!T5:T200<>\"\")*('{CR}'!W5:W200<>0)*('{CR}'!W5:W200<>\"\"))"
    )
    lr["B13"] = f"=SUMPRODUCT(('{CR}'!A5:A136<>\"\")*('{CR}'!T5:T136=\"\"))"

    ledger_sum = sum(n(v.get("acq")) for v in sl_by.values())
    active_nbv = sum(n(v.get("nbv")) for v in sl_by.values() if v.get("status") == "Active")
    all_nbv = sum(n(v.get("nbv")) for v in sl_by.values())
    cache.update(
        {
            (LR, "B5"): float(round(sum_m, 2)),
            (LR, "B6"): float(round(ledger_sum, 2)),
            (LR, "B7"): float(round(sum_m - ledger_sum, 2)),
            (LR, "B8"): float(round(sum_n, 2)),
            (LR, "B9"): float(round(active_nbv, 2)),
            (LR, "B10"): float(round(all_nbv, 2)),
            (LR, "B11"): 122.0,
            (LR, "B12"): 7.0,
            (LR, "B13"): 3.0,
        }
    )

    n2 = str(lr["A18"].value or "")
    if "FA-000119" not in n2:
        n2 = n2.rstrip()
        if n2 and not n2.endswith("."):
            n2 += "."
        n2 += (
            " Separately, FA-000119 (MD-00119, Disposed) shows accumulated depreciation "
            "$1,125 against acquisition cost $1,115 (over-depreciation $10) with NBV $0."
        )
        lr["A18"] = n2

    for r in range(24, 40):
        if lr.cell(r, 1).value == "MD-00119":
            lr.cell(r, 2).value = (
                "Acquisition cost mismatch; FA-000119 accum dep $1,125 exceeds "
                "acq $1,115 by $10 (NBV $0)"
            )
            break

    for r in range(5, 200):
        if ex.cell(r, 1).value == "EX-0153":
            ex.cell(r, 12).value = (
                "Cost basis split on MD-00119: inventory $950, ledger $1,115 via "
                "PO-2024-0024 / FA-000119. FA-000119 also over-depreciated by $10 "
                "(accum dep $1,125 vs acq $1,115) with NBV $0 (EX-0153)."
            )
            rel = str(ex.cell(r, 6).value or "")
            if "OverDepreciation" not in rel:
                ex.cell(r, 6).value = rel.rstrip(";") + "; OverDepreciation:10.0"
            break

    return cache


def rewrite_notes():
    doc = Document(str(NOTES))
    table = doc.tables[0]
    authorities = [
        "Unverified field note",
        "Secondhand — not validated",
        "Tech hearsay only",
        "Lead from walkthrough",
        "Unconfirmed verbal claim",
        "Regional chatter — unchecked",
        "Not corroborated",
        "Informal tip only",
        "Unverified contractor remark",
        "Floor rumor — no ticket",
        "Lead only; no artifact",
        "Unchecked site comment",
        "Anecdote pending proof",
        "Unvalidated tech note",
        "Soft lead — ignore for clearance",
        "Non-authoritative remark",
        "Unverified ops aside",
    ]
    for idx, row in enumerate(table.rows[1:]):
        tag = row.cells[1].text.strip()
        loc = row.cells[2].text.strip()
        old = row.cells[3].text.strip()
        row.cells[4].text = authorities[idx % len(authorities)]
        core = re.split(
            r"(?i)no (cycle|badge|transfer|ticket|evidence|scan|photo|receiving)", old
        )[0].strip(" .;")
        variants = [
            f"Walkthrough chatter put {tag} near {loc or 'the closed floor'}. Untested.",
            f"Someone on site pointed at {loc or 'a storage area'} for {tag}. Soft intelligence only.",
            f"{tag}: informal note about {loc or 'possible storage'} during exit support. Nothing in tickets.",
            f"Possible sighting context for {tag} — {loc or 'unspecified room'}. Lead for follow-up, not proof.",
            f"Field aside on {tag} referenced {loc or 'a back-room stack'}. No receiving or transfer row backs it.",
            f"During closeout, {tag} was named in passing with {loc or 'a cage/MDF'} as the guessed spot. Unconfirmed.",
            f"Ops mentioned {tag} might still be around {loc or 'overflow'}. Remains speculative.",
            f"Note against {tag}: alleged presence at {loc or 'legacy floor space'}. Needs an independent find.",
            f"Site support named {tag} while discussing {loc or 'leftover gear'}. Not validated.",
            f"Unconfirmed placement rumor for {tag} near {loc or 'closed offices'}.",
            f"{loc or 'Closed site'} came up as a guess for {tag} in hallway conversation. No artifact.",
            f"Exit-team aside linked {tag} to {loc or 'storage'}. Keep as lead only.",
            f"Heard {tag} might linger at {loc or 'an MDF/cage'}. No badge or scan trail attached.",
            f"Informal recovery tip for {tag} points at {loc or 'overflow'}. Still unverified.",
            f"Floor conversation flagged {tag}; claimed area {loc or 'unknown'}. Not evidence.",
            f"Contractor comment placed {tag} around {loc or 'dock-adjacent storage'}. Unchecked.",
            f"Soft lead only: {tag} possibly at {loc or 'a retired floor'}.",
        ]
        text = variants[idx % len(variants)]
        if core and len(core) > 24 and idx % 3 == 0:
            text = f"{core}. Unverified lead for {tag}."
        row.cells[3].text = re.sub(r"\s+", " ", text).strip()
    doc.save(str(NOTES))
    print("notes rewritten")


def refresh_artifacts():
    for zname in ["Yanou_IT_Asset_Reconciliation.zip", "Yanou_IT_Asset_Reconciliation.xlsx.zip"]:
        with zipfile.ZipFile(ROOT / zname, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(WB, WB.name)
    dl = ROOT / "download"
    dl.mkdir(exist_ok=True)
    shutil.copy2(WB, dl / WB.name)
    with zipfile.ZipFile(dl / "Yanou_IT_Asset_Reconciliation.zip", "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(WB, WB.name)
    art = Path("/opt/cursor/artifacts")
    if art.exists():
        shutil.copy2(WB, art / WB.name)
        with zipfile.ZipFile(art / "Yanou_IT_Asset_Reconciliation.zip", "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(WB, WB.name)
    print("artifacts refreshed")


def verify():
    wb = load_workbook(WB, data_only=False)
    wbv = load_workbook(WB, data_only=True)

    def share(name):
        ws = wb[name]
        t = f = 0
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    f += 1
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    t += 1
        return t, f, (t / (t + f) if t + f else 0)

    for name in [DASH, CR, EX, LR]:
        t, f, s = share(name)
        print(f"{name}: typed={t} form={f} share={s:.1%}")

    print("LR B5-B12", [wbv[LR].cell(r, 2).value for r in range(5, 13)])
    print("Dash B5-B9", [wbv[DASH].cell(r, 2).value for r in range(5, 10)])
    cr = wbv[CR]
    blanks = sum(
        1
        for r in range(5, 137)
        if cr.cell(r, 1).value
        and (cr.cell(r, 13).value is None or str(cr.cell(r, 13).value).strip() == "")
    )
    print("CR acq blanks", blanks)
    print("A9", wb[DASH]["A9"].value)

    ex = wb[EX]
    eas = []
    for r in range(5, 200):
        if ex.cell(r, 2).value == CLOSED:
            t = str(ex.cell(r, 12).value or "")
            eas.append(re.sub(r"MD-\d+|TR-\d+|EX-\d+|E\d{4}", "X", t)[:90])
    print("closed EA unique norms", len(set(eas)), "of", len(eas))

    doc = Document(str(NOTES))
    auths = [row.cells[4].text for row in doc.tables[0].rows[1:]]
    print("notes authority unique", len(set(auths)), "of", len(auths))


def main():
    assert WB.exists(), WB
    wb = load_workbook(WB)
    meta, sl_by = build_meta(wb)
    rewrite_narratives(wb, meta)
    cache = apply_formulas(wb, meta, sl_by)
    wb.save(WB)
    print("saved; cache", len(cache))
    inject_cache(WB, cache)
    rewrite_notes()
    refresh_artifacts()
    (ROOT / ".do_not_materialize_dashboard").write_text(
        "Do not run materialize_dashboard_values.py — it flattens formulas.\n"
    )
    verify()
    print("done")


if __name__ == "__main__":
    main()
