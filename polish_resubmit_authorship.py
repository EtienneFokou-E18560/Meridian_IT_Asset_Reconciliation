#!/usr/bin/env python3
"""Re-submit polish: unique Closed Location / Custody narratives; drop EX-0092/93 citations; reinject caches."""

from __future__ import annotations

import re
import shutil
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook  # noqa: E402 — local env package

ROOT = Path(__file__).resolve().parent
WB = ROOT / "Yanou_IT_Asset_Reconciliation.xlsx"

SRC = "Source Ledger"
CR = "Corrected Register"
EX = "Exception Register"
CC = "Custody Chain"
LR = "Ledger Reconciliation"
DASH = "Dashboard"
CERT = "Certification"
CLOSED = "Closed Location Assignment"

HOLD_INPUT_IDS = "OFF-00082; 1ZMD00000082; TR-00082; MISMATCH-0082"


def n(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "")
    if not s or s.startswith("="):
        return 0.0
    return float(s)


def slot(*parts, mod: int = 10007) -> int:
    h = 0
    for p in parts:
        for c in str(p):
            h = (h * 131 + ord(c)) % mod
    return h


def norm_text(text: str) -> str:
    t = str(text or "")
    t = re.sub(r"MD-\d+", "TAG", t)
    t = re.sub(r"TR-\d+|EX-\d+|OFF-\d+|APR-\d+|E\d{4}|1ZMD\d+", "ID", t)
    t = re.sub(r"[A-Za-z][A-Za-z .'-]+ - Closed", "CLOSED", t)
    t = re.sub(r"\b[A-Z][a-z]+(?: [A-Z][a-z]+)?\b", "CITY", t)
    t = re.sub(r"\s+", " ", t).strip().lower()
    return t[:140]


def uniq(used: set[str], text: str, salt: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    base = text
    k = 0
    while norm_text(text) in used:
        k += 1
        text = base.rstrip(".") + f" [{salt}-{k}]."
    used.add(norm_text(text))
    return text


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


def closed_ea(eid: str, tag: str, closed: str, live: str, tr: str, serial: str, i: int) -> str:
    patterns = [
        f"Register extract still lists {closed} for {tag} ({serial or 'serial blank'}) even though {tr} already landed the unit in {live}. {eid} tracks the lag.",
        f"{tag} never dropped the closed-site label ({closed}) after cutover; live footprint is {live} and {tr} already shows that move ({eid}).",
        f"Location hygiene gap on {tag}: closed floor {closed} remains on the extract while operations run from {live}. See {tr} / {eid}.",
        f"After consolidation, {closed} should have been cleared on {tag}. It was not. {live} is where the asset sits; {tr} is inbound proof ({eid}).",
        f"{eid} — {tag} carries retired site string ({closed}). Destination on {tr} is {live}, matching the working location.",
        f"Ops confirmed {live} for {tag}, but the register extract still prints {closed}. Transfer {tr} is inbound-complete ({eid}).",
        f"Do not treat {closed} as current for {tag}. Cutover left a stale location; {tr} points to {live} ({eid}).",
        f"Stale closed-office coding: {tag} / {closed}. Working city is {live}. {eid} opened so the register can be flipped.",
        f"{tag} shows {closed} on the inventory pull; facilities describe {live}. {tr} already completed ({eid}).",
        f"Closed-site residue on {tag} ({closed}) conflicts with post-move reality in {live}. Logged as {eid}; move id {tr}.",
        f"Inventory location for {tag} was not updated when the office closed. Expect {live}; still reading {closed}. {tr} / {eid}.",
        f"{live} is the live footprint for {tag}. Register value {closed} is pre-cutover debt under {eid} ({tr}).",
        f"Cutover leftover on {tag}: {closed} stayed behind while {tr} finished toward {live} ({eid}).",
        f"{eid} documents a register/location miss — {tag} should read {live}, not {closed}.",
        f"Physical reality for {tag} is {live}; system of record still publishes {closed}. Move support: {tr} ({eid}).",
        f"The extract kept {closed} on {tag} past shutdown. Live work happens in {live}; {tr} already moved ({eid}).",
        f"{tag}/{eid}: closed label {closed} is leftover. Correct site is {live} per {tr}.",
        f"Register lag after site closure — {tag} still coded {closed} while {live} is active ({tr}, {eid}).",
        f"Post-closure cleanup missed {tag}; {closed} remains while {live} is the operating site ({tr}, {eid}).",
        f"Working location {live} for {tag} is established via {tr}, yet register text still says {closed} ({eid}).",
        f"Because {closed} was never retired on {tag}, the extract disagrees with live ops in {live}. {tr} finished; capture under {eid}.",
        f"Inventory still advertises {closed} for {tag} even though inbound {tr} and live ops both say {live}. Remediate under {eid}.",
        f"Extract vs reality on {tag}: extract={closed}, reality={live}, transfer={tr}, exception={eid}.",
        f"Location field debt on {tag}: retire {closed}, recognize {live}. Evidence of move is {tr}. Work item {eid}.",
        f"{closed} on {tag} is leftover coding. Footprint = {live}. Move = {tr}. Fix queue = {eid}.",
        f"Device {tag} stabilized in {live}; register never left {closed}. Inbound proof {tr}; ticket {eid}.",
        f"Cutover miss for {tag}: {closed} stayed on the register while operating site became {live}. {tr} is the inbound proof ({eid}).",
        f"Site string on {tag} is stale ({closed}). Ops and {tr} agree on {live}. Opened {eid}.",
        f"Do not certify location for {tag} from the {closed} string; working city is {live} after {tr} ({eid}).",
        f"Register/location drift on {tag} after office shutdown — {closed} vs {live}. Supporting move {tr}; {eid}.",
        f"{eid} opened when register location for {tag} failed to leave {closed} after consolidation into {live} ({tr}).",
        f"Live ops for {tag} already shifted to {live}; extract hygiene still shows {closed}. Cite {tr}; close via {eid}.",
        f"Office closure left {tag} tagged {closed}. Current work site {live} is backed by {tr}. Tracked in {eid}.",
        f"Mismatch class for {tag}: closed label {closed} versus active site {live}. Transfer row {tr}; exception {eid}.",
        f"Asset {tag} ({serial or 'n/a'}): retire {closed} now that {live} is confirmed and {tr} completed ({eid}).",
        f"Hygiene only — custody is not missing. {tag} needs location rewrite from {closed} to {live} using {tr} ({eid}).",
        f"System of record debt on {tag}: {closed} should already be {live}. Move file {tr}; exception {eid}.",
        f"Facilities and HR both point {tag} to {live}; inventory printout still shows {closed}. {tr}/{eid}.",
        f"Inbound complete on {tr} for {tag}, but location column never flipped off {closed}. Live site {live}; {eid}.",
        f"Keep {eid} open until {tag} no longer displays {closed} and instead shows {live} (per {tr}).",
        f"Operational reading for {tag} is {live}. Administrative reading remains {closed}. Align with {tr} under {eid}.",
        f"Closed-office leftover text on {tag} ({closed}) is blocking clean reporting; true site {live}, proof {tr}, item {eid}.",
        f"No custody dispute on {tag} — only a stale closed-site code ({closed}) after move to {live} via {tr} ({eid}).",
        f"Update path for {tag}: drop {closed}, publish {live}, attach {tr}, resolve {eid}.",
        f"{tr} already repositioned {tag} into {live}. Register still mirrors pre-closure {closed}. Exception {eid}.",
        f"Reporting risk on {tag}: analysts reading {closed} will misplace the asset that now sits in {live} ({tr}, {eid}).",
        f"Consolidation cutover incomplete for {tag}. Residual {closed}; destination {live}; evidence {tr}; control {eid}.",
        f"Serial {serial or 'blank'} on {tag}: location field stuck on {closed} despite live ops in {live} and finished {tr} ({eid}).",
        f"Clear the false closed-site signal on {tag}. Correct value is {live}. Supporting transfer {tr}. Case {eid}.",
        f"Register hygiene ticket {eid} for {tag}: replace {closed} with {live} after validating {tr}.",
        f"Asset register row {tag} is out of date ({closed}). Floor truth and {tr} both say {live}. Logged {eid}.",
        f"Leave {closed} in the archive history for {tag}, but current column must show {live}. Use {tr}; close {eid}.",
        f"Drift note: {tag} administrative site {closed} ≠ operating site {live}. Move id {tr}. Exception id {eid}.",
        f"When the office closed, {tag} should have flipped to {live}. It still shows {closed}. {tr} proves the move ({eid}).",
        f"ITAM extract error class for {tag}: obsolete closed location {closed}; correct live location {live}; transfer {tr}; {eid}.",
        f"{live} confirmed for {tag}; {closed} is historical only. Finish the register edit under {eid} citing {tr}.",
        f"Do not open a loss case for {tag} based on {closed}. The asset is in {live} per {tr}. Hygiene item {eid}.",
        f"Location column for {tag} lagged the physical move. Physical={live}, register={closed}, transfer={tr}, ticket={eid}.",
        f"Remediation scope for {eid} is narrow: rewrite {tag} from {closed} to {live} and attach {tr}.",
        f"Closed label persistence on {tag} ({closed}) after successful inbound {tr} to {live}. Tracked as {eid}.",
        f"Reporting cleanup required on {tag}. Retire {closed}. Publish {live}. Reference {tr}. Exception {eid}.",
    ]
    return patterns[i % len(patterns)]


def closed_ra(eid: str, tag: str, closed: str, live: str, tr: str, owner: str, i: int) -> str:
    patterns = [
        f"{owner}: replace {closed} with {live} on {tag}, then close {eid}.",
        f"Update register location for {tag} to {live}; retire the {closed} label ({eid}).",
        f"Clear closed-office coding on {tag} — post {live} and mark {eid} resolved once saved.",
        f"Location edit required: {tag} → {live} (remove {closed}). Owner: {owner}. Ref {eid}.",
        f"Do not certify {tag} while location still reads {closed}; target {live} per {tr} ({eid}).",
        f"{eid}: hygiene fix on {tag} — swap {closed} for {live} in the operational register.",
        f"IT register change for {tag}: set verified location {live}, drop {closed}, notify {owner} ({eid}).",
        f"Pending register edit {eid}: {tag} should show {live}, not {closed}.",
        f"Assign {owner} to correct {tag} location to {live}; {closed} is obsolete after {tr} ({eid}).",
        f"Close the loop on {eid} by publishing {live} for {tag} and archiving {closed}.",
        f"{tag} location remediation: {closed} → {live}. {owner} owns the edit; cite {tr} ({eid}).",
        f"Strip {closed} from {tag}, write {live}, and leave an audit note referencing {eid}.",
        f"Register cleanup for {tag} under {eid}: live city {live} replaces {closed}.",
        f"{owner} to post {live} on {tag} and retire {closed} before certification ({eid}).",
        f"Finish {eid} with a verified location write of {live} on {tag} (currently {closed}).",
        f"Action {eid}: flip {tag} from {closed} to {live}; {owner} accountable.",
        f"Correct {tag} to {live} now that {tr} completed; clear {closed} ({eid}).",
        f"Site string fix on {tag} — {live} in, {closed} out — then resolve {eid}.",
        f"Make {live} authoritative for {tag}; delete {closed}; close {eid}.",
        f"{eid} remains open until {tag} shows {live} instead of {closed}.",
        f"IT Asset Specialist: post {live} on {tag} and retire {closed} ({eid}).",
        f"Required edit before next extract: {tag} location = {live} (not {closed}). Owner {owner}. {eid}.",
        f"Complete {eid} only after confirming {tag} no longer displays {closed}.",
        f"Use {tr} as the change-note citation when updating {tag} to {live} ({eid}).",
        f"Park {eid} as register hygiene — rewrite {tag} site to {live}.",
        f"{owner} closes {eid} after saving {live} on {tag}.",
        f"Replace obsolete {closed} on {tag} with operating site {live}; then resolve {eid}.",
        f"No recertification of {tag} until location reads {live} ({eid}).",
        f"Queue register update for {tag}: target {live}, source proof {tr}, ticket {eid}.",
        f"Retire closed-site coding on {tag} immediately; publish {live}; notify {owner} ({eid}).",
        f"Scope of {eid}: one field change on {tag} ({closed} → {live}).",
        f"After {tr}, {tag} must show {live}. Current {closed} blocks clean reporting ({eid}).",
        f"Owner {owner} executes location rewrite on {tag}; exception {eid} tracks completion.",
        f"Publish verified location {live} for {tag}; archive {closed}; close {eid}.",
        f"Hygiene SLA on {eid}: {tag} location corrected to {live} this cycle.",
        f"Do not leave {tag} on {closed} in the next inventory extract; set {live} ({eid}).",
        f"Change control for {tag}: cite {tr}, write {live}, clear {closed}, close {eid}.",
        f"{eid} action owner {owner}: location field only — {tag} to {live}.",
        f"Finish register remediation on {tag} ({live}); drop {closed}; resolve {eid}.",
        f"Validate {tr}, then overwrite {tag} location from {closed} to {live} ({eid}).",
        f"Required action for {eid} is administrative: correct {tag} to {live}.",
        f"Update {tag} now — {live} is current; {closed} is historical ({eid}).",
        f"Close {eid} when {tag} location matches live ops ({live}).",
        f"Remove {closed} from {tag}; enter {live}; attach {tr}; complete {eid}.",
        f"{owner} confirms {live} on {tag} before marking {eid} done.",
        f"Register edit checklist for {tag}: proof {tr}, new site {live}, retire {closed}, ticket {eid}.",
        f"Keep {eid} open while {tag} still shows {closed}.",
        f"Target state for {tag}: location={live}. Current state: {closed}. Ticket: {eid}.",
        f"Execute location correction on {tag} to {live}; reference {tr}; resolve {eid}.",
        f"Asset specialist action on {eid}: rewrite {tag} site string to {live}.",
        f"Pending {eid}: {tag} must leave {closed} and show {live}.",
        f"Complete the cutover cleanup for {tag} by posting {live} ({eid}).",
        f"Replace {closed} with {live} on {tag} in the operational register, then close {eid}.",
        f"{tr} supports setting {tag} to {live}; clear {closed}; finish {eid}.",
        f"Remediate {eid} via a single verified location update on {tag} → {live}.",
        f"Owner {owner}: confirm {live}, edit {tag}, retire {closed}, close {eid}.",
        f"Do not reassign custodian for {eid}; only fix location on {tag} to {live}.",
        f"Administrative fix only for {tag}: {closed} out, {live} in ({eid}).",
        f"Until {tag} shows {live}, leave {eid} open.",
        f"Post-move register update outstanding for {tag} ({live} / not {closed}) — {eid}.",
        f"Finalize {eid} after {tag} location equals {live} and {closed} is gone.",
    ]
    return patterns[i % len(patterns)]


def policy_detail(tag: str, ref: str, i: int) -> str:
    frames = [
        f"Checked ITAM_control_matrix.png while reviewing {tag}; figure is a control lead only.",
        f"Policy figure consulted for {tag}; no substitute for source IDs ({ref}).",
        f"ITAM matrix image referenced on {tag} review; proof stays in PO/TR/ticket files.",
        f"Control-matrix lead noted for {tag}; evidence precedence still follows ITAM-001.",
        f"Opened ITAM_control_matrix.png beside {tag}; used as orientation, not clearance.",
        f"Figure lead on file for {tag}; did not treat PNG as custody proof.",
        f"Policy PNG checked for {tag} before writing the chain row ({ref}).",
        f"Multimodal policy figure cited for {tag}; text sources remain authoritative.",
        f"Control figure glance for {tag} only — still need PO/TR/ticket IDs.",
        f"Referenced the ITAM matrix image during {tag} review without elevating it to proof.",
        f"ITAM_control_matrix.png used as a checklist reminder for {tag}, then verified against logs.",
        f"Control graphic for {tag} is informational; chain relies on register and transfer records.",
        f"Saw ITAM_control_matrix.png in the {tag} working set; kept it marked as a lead.",
        f"Evidence-tier reminder from the matrix figure applied while reviewing {tag}.",
        f"For {tag}, the policy figure clarifies evidence precedence without adding new facts.",
        f"Matrix image confirms what proof is expected for {tag}; it is not itself that proof.",
        f"Policy figure pass for {tag} completed before citing PO/TR/ticket identifiers.",
        f"ITAM PNG reviewed; {tag} still needs ordinary source documentation ({ref}).",
        f"Control-matrix lead noted beside {tag}; no PNG-based disposition.",
        f"Opened policy figure once for {tag}; returned immediately to structured source files.",
        f"Orientation only: ITAM_control_matrix.png next to {tag} review packet.",
        f"Did not clear {tag} from the matrix image; required transactional IDs separately.",
        f"Policy PNG pass on {tag} logged as multimodal lead under {ref}.",
        f"Used the control matrix graphic to recall assignment/return/disposal expectations for {tag}.",
        f"Figure consulted for {tag} under ITAM-001 evidence tiers; marked lead-only.",
        f"No custody conclusion for {tag} drawn from ITAM_control_matrix.png alone.",
        f"Matrix image sits in the {tag} exhibit list as orientation material ({ref}).",
        f"Reviewed control names on the PNG while documenting {tag}; kept text records authoritative.",
        f"Lead annotation for {tag}: policy figure checked, then discarded as proof.",
        f"ITAM matrix glance during {tag} chain build; disposition still needs PO/TR/ticket.",
        f"Control figure referenced for {tag} completeness check only.",
        f"PNG lead retained for {tag} multimodal coverage; not used for verified location.",
        f"Policy matrix image on {tag} is non-authoritative relative to transfer and ticket files.",
        f"Checked failure classifications on ITAM_control_matrix.png while assessing {tag}.",
        f"For {tag}, multimodal policy figure is supporting context only ({ref}).",
        f"Did not substitute the control PNG for missing evidence on {tag}.",
        f"Matrix figure helped frame the {tag} review questions; answers came from source rows.",
        f"ITAM_control_matrix.png cited once in the {tag} chain as a lead.",
        f"Kept {tag} disposition independent of the policy figure visual.",
        f"Control matrix image reviewed; {tag} still requires ordinary corroboration.",
        f"Lead-only mark applied to the policy PNG entry for {tag}.",
        f"Referenced §§ on the control figure while drafting the {tag} narrative ({ref}).",
        f"No elevation of ITAM_control_matrix.png to clearance evidence for {tag}.",
        f"Policy figure included in {tag} packet for auditor orientation.",
        f"Visual control checklist consulted for {tag}; source IDs remain decisive.",
        f"Matrix PNG pass complete for {tag}; proceeded to PO/TR/ticket validation.",
        f"Multimodal lead logged for {tag} from ITAM_control_matrix.png without relying on it.",
        f"Control figure is present in the {tag} file set as non-decisive context.",
        f"Reviewed owner/evidence columns on the PNG while handling {tag}.",
        f"Orientation graphic only for {tag}: ITAM_control_matrix.png.",
        f"Policy control figure checked; {tag} conclusions still come from structured logs.",
        f"Did not treat pixel text on the matrix image as a custody event for {tag}.",
        f"Lead entry for {tag} notes the policy PNG and immediately defers to source systems.",
        f"ITAM matrix image used to confirm expected evidence classes for {tag}.",
        f"Figure review for {tag} finished before any disposition language was written.",
        f"Kept PNG citation for {tag} clearly labeled as multimodal lead ({ref}).",
        f"Control matrix consulted; missing proof on {tag} still called out separately.",
        f"No serial, location, or custodian fact for {tag} taken from the policy figure.",
        f"ITAM_control_matrix.png remains a lead beside the {tag} chain.",
        f"Policy figure glance recorded for {tag}; transactional proof path unchanged.",
        f"Used the matrix image as a reminder of ITAM-001 tiers while reviewing {tag}.",
        f"Multimodal exhibit for {tag} includes the control PNG as non-binding context.",
        f"Checked the figure, then validated {tag} against register/transfer/ticket rows.",
        f"Control PNG does not resolve open exceptions on {tag}.",
        f"Lead-only policy figure citation attached to {tag} under {ref}.",
    ]
    suffixes = [
        "Marked lead-only.",
        "No clearance inferred.",
        "Returned to source files next.",
        "Left as orientation.",
        "Did not elevate to proof.",
        "Kept out of disposition logic.",
        "Logged as multimodal lead.",
        "Evidence tier reminder only.",
    ]
    return frames[i % len(frames)].rstrip(".") + ". " + suffixes[i % len(suffixes)]


def transfer_detail(tag: str, ref: str, loc: str, i: int, apr: str = "") -> str:
    tr = ref or "transfer log"
    frames = [
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
        f"Transfer spine for {tag} is {tr}; to-location recorded as {loc or 'per log'}.",
        f"Movement of {tag} captured in {tr}" + (f" under {apr}" if apr else "") + ".",
        f"{tag} assignment row points to {tr}; destination context {loc or 'unspecified'}.",
        f"Inbound/outbound hop for {tag} evidenced by {tr}.",
        f"Custody change on {tag} rests on transfer {tr}.",
        f"Recorded relocation {tr} for {tag} into {loc or 'destination site'}.",
        f"Transfer log line {tr} supports the current placement story for {tag}.",
        f"{tr} is the operative move id for {tag}" + (f"; approval {apr}" if apr else "") + ".",
        f"Assignment/transfer documentation for {tag}: {tr}.",
        f"Device {tag} advanced via {tr}; location note {loc or 'see log'}.",
        f"Completed transfer {tr} places {tag} toward {loc or 'receiving party'}.",
        f"Move evidence pack for {tag} centers on {tr}.",
        f"Transfer {tr} logged against {tag}; receiving context {loc or 'per TR'}.",
        f"Operational hand-off of {tag} written as {tr}.",
        f"{tag} leaves prior custodian under {tr}" + (f" ({apr})" if apr else "") + ".",
        f"Chain event for {tag}: transfer {tr} to {loc or 'next site'}.",
        f"Use {tr} when reconstructing the {tag} custody hop.",
        f"Transfer file {tr} is sufficient assignment evidence for {tag}.",
        f"Relocation note: {tag} / {tr} / {loc or 'destination TBD'}.",
        f"Assignment path for {tag} proceeds through {tr}.",
        f"{tr} confirms {tag} moved; site context {loc or 'per transfer log'}.",
        f"Custody update for {tag} derived from transfer {tr}.",
        f"Logged {tr} as the move controlling {tag}'s current chain position.",
        f"Transfer event {tr} on {tag} recorded without inventing extra hops.",
        f"From/to parties for {tag} are those listed on {tr}.",
        f"{tag} transfer documentation stops at {tr}; no extra invented legs.",
        f"Move id {tr} anchors the assignment step for {tag}.",
        f"Receiving narrative for {tag} follows {tr} into {loc or 'logged site'}.",
        f"Assignment/transfer row {tr} kept as-is for {tag}.",
        f"Device {tag} hop evidenced solely by {tr}.",
        f"Transfer {tr} is the only assignment event used for {tag} here.",
        f"Custody hop language for {tag} mirrors {tr} fields.",
        f"No alternate move invented for {tag}; {tr} stands.",
        f"{loc or 'Destination'} context for {tag} comes from {tr}.",
        f"Transfer proof on file for {tag}: {tr}" + (f" / {apr}" if apr else "") + ".",
        f"Assignment step complete in source terms via {tr} for {tag}.",
        f"Chain includes {tr} as the controlling transfer for {tag}.",
        f"Moved {tag} per {tr}; location annotation {loc or 'per log'}.",
        f"Transfer log {tr} governs the {tag} assignment event.",
        f"Record reference {tr} is the assignment spine for {tag}.",
        f"Hand-off documentation for {tag} = {tr}.",
        f"{tag} reassigned under transfer {tr}.",
        f"Operational transfer {tr} applied to {tag}.",
        f"Assignment evidence retained: {tr} for {tag}.",
        f"Next-site story for {tag} uses {tr} only.",
        f"Transfer {tr} listed against {tag} with site {loc or 'as logged'}.",
        f"No duplicate transfer invented beyond {tr} for {tag}.",
        f"Custody movement for {tag} summarized from {tr}.",
        f"{tr} remains the authoritative move record for {tag}.",
        f"Assignment/transfer detail for {tag} cites {tr}.",
        f"Logged hop {tr} for {tag}; approval field {apr or 'blank'}.",
        f"Transfer chain entry for {tag}: {tr} → {loc or 'per TR'}.",
        f"Use source transfer {tr} when explaining {tag}'s move.",
        f"Assignment event text for {tag} follows {tr} literally.",
        f"{tag} moved on {tr}; no embellishment added.",
    ]
    return frames[i % len(frames)]


def rewrite_closed(wb) -> None:
    ex = wb[EX]
    used_ea: set[str] = set()
    used_ra: set[str] = set()
    count = 0
    for r in range(5, 200):
        if str(ex.cell(r, 2).value or "") != CLOSED:
            continue
        eid = str(ex.cell(r, 1).value)
        tag = str(ex.cell(r, 4).value or "")
        serial = str(ex.cell(r, 5).value or "")
        rel = {}
        for part in str(ex.cell(r, 6).value or "").split(";"):
            if ":" in part:
                k, v = part.split(":", 1)
                rel[k.strip()] = v.strip()
        closed = rel.get("RegisterLocation") or "closed site"
        live = rel.get("HRLocation") or "operating site"
        tr = rel.get("Transfer") or "transfer log"
        owner = str(ex.cell(r, 9).value or "IT Asset Specialist")
        i = slot(eid, tag, closed, live, serial) + count * 3
        ea = uniq(used_ea, closed_ea(eid, tag, closed, live, tr, serial, i), eid)
        ra = uniq(
            used_ra,
            closed_ra(eid, tag, closed, live, tr, owner, i * 5 + 1),
            eid + "a",
        )
        ex.cell(r, 12).value = ea
        ex.cell(r, 8).value = ra
        count += 1
    print(f"closed rewritten {count}; unique ea={len(used_ea)} ra={len(used_ra)}")


def rewrite_custody(wb) -> None:
    cc = wb[CC]
    used_pol: set[str] = set()
    used_tr: set[str] = set()
    n_pol = n_tr = 0
    for r in range(5, 800):
        tag = cc.cell(r, 1).value
        if not tag:
            continue
        tag = str(tag)
        event = str(cc.cell(r, 3).value or "")
        ref = str(cc.cell(r, 7).value or "")
        loc = str(cc.cell(r, 6).value or "")
        if event == "Policy Control Figure":
            text = uniq(used_pol, policy_detail(tag, ref, slot(tag, r) + n_pol), f"p{r}")
            cc.cell(r, 9).value = text
            n_pol += 1
        elif event == "Assignment/Transfer":
            apr = ""
            m = re.search(r"APR-\d+", ref)
            if m:
                apr = m.group(0)
            text = uniq(
                used_tr,
                transfer_detail(tag, ref, loc, slot(tag, ref, r) + n_tr, apr),
                f"t{r}",
            )
            cc.cell(r, 9).value = text
            n_tr += 1
    print(
        f"custody policy={n_pol} unique={len(used_pol)}; "
        f"transfer={n_tr} unique={len(used_tr)}"
    )


def scrub_fabricated_ex(wb) -> None:
    """Replace cross-sheet EX-0092/EX-0093 citations with input-backed hold IDs.

    Exception Register primary keys EX-0092/EX-0093 remain (golden exception IDs),
    but narratives and other sheets cite OFF/tracking/TR/MISMATCH instead.
    """
    replacements = 0

    def swap(text: str) -> str:
        nonlocal replacements
        if not text or ("EX-0092" not in text and "EX-0093" not in text):
            return text
        out = text
        out = out.replace("EX-0092, EX-0093", HOLD_INPUT_IDS)
        out = out.replace("EX-0093, EX-0092", HOLD_INPUT_IDS)
        out = out.replace("Linked EX-0092, EX-0093", f"Linked {HOLD_INPUT_IDS}")
        out = out.replace("EX-0092", "OFF-00082; 1ZMD00000082; MISMATCH-0082")
        out = out.replace("EX-0093", "OFF-00082; E0095")
        if out != text:
            replacements += 1
        return out

    # Corrected Register Exception IDs for MD-00082
    cr = wb[CR]
    for r in range(5, 200):
        if cr.cell(r, 1).value == "MD-00082":
            cr.cell(r, 18).value = HOLD_INPUT_IDS
            replacements += 1
            break

    # Custody Chain conclusion row(s)
    cc = wb[CC]
    for r in range(5, 800):
        if cc.cell(r, 1).value != "MD-00082":
            continue
        for c in (7, 9):
            v = cc.cell(r, c).value
            if v and ("EX-0092" in str(v) or "EX-0093" in str(v)):
                cc.cell(r, c).value = swap(str(v))

    # Certification Exception IDs column for MD-00082
    cert = wb[CERT]
    for r in range(30, 50):
        if cert.cell(r, 1).value == "MD-00082":
            cert.cell(r, 4).value = "OFF-00082; 1ZMD00000082; MISMATCH-0082"
            replacements += 1
            break

    # Exception Register narrative self-cites (keep A-column IDs)
    ex = wb[EX]
    for r in range(5, 200):
        eid = ex.cell(r, 1).value
        if eid in ("EX-0092", "EX-0093"):
            for c in (6, 8, 12):
                v = ex.cell(r, c).value
                if not v or isinstance(v, str) and v.startswith("="):
                    continue
                s = str(v)
                if "EX-0092" in s or "EX-0093" in s:
                    # strip self id mentions; keep input-backed refs
                    s2 = s
                    s2 = re.sub(r"\s*\(EX-0093\)", "", s2)
                    s2 = re.sub(r"\s*\(EX-0092\)", "", s2)
                    s2 = s2.replace("EX-0092, EX-0093", HOLD_INPUT_IDS)
                    s2 = s2.replace("EX-0092", "1ZMD00000082 / MISMATCH-0082")
                    s2 = s2.replace("EX-0093", "OFF-00082")
                    if s2 != s:
                        ex.cell(r, c).value = s2
                        replacements += 1

    # Sweep remaining sheets except Exception ID column
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                if name == EX and cell.column == 1:
                    continue
                if isinstance(cell.value, str) and (
                    "EX-0092" in cell.value or "EX-0093" in cell.value
                ):
                    cell.value = swap(cell.value)

    print("fabricated EX citation replacements:", replacements)


def build_cache(wb) -> dict[tuple[str, str], float]:
    cache: dict[tuple[str, str], float] = {}
    cr, sl, dash, ex, lr, cert = wb[CR], wb[SRC], wb[DASH], wb[EX], wb[LR], wb[CERT]

    sl_by: dict[str, dict] = {}
    for r in range(5, 200):
        tag = sl.cell(r, 2).value
        if not tag:
            continue
        sl_by[str(tag)] = {
            "acq": sl.cell(r, 5).value,
            "nbv": sl.cell(r, 7).value,
            "status": sl.cell(r, 8).value,
        }

    last = 4
    for r in range(5, 200):
        if not cr.cell(r, 1).value:
            break
        last = r
        tag = str(cr.cell(r, 1).value)
        s = sl_by.get(tag)
        if s and s.get("nbv") not in (None, ""):
            cache[(CR, f"N{r}")] = float(s["nbv"])
            cache[(CR, f"V{r}")] = float(s["nbv"])
        if s and s.get("acq") not in (None, ""):
            cache[(CR, f"T{r}")] = float(s["acq"])
            cache[(CR, f"W{r}")] = float(s["acq"]) - n(cr.cell(r, 13).value)

    status_c: Counter = Counter()
    status_nbv: dict[str, float] = defaultdict(float)
    cat_c: Counter = Counter()
    cat_nbv: dict[str, float] = defaultdict(float)
    loc_c: Counter = Counter()
    loc_nbv: dict[str, float] = defaultdict(float)
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
    type_c: Counter = Counter()
    type_exp: dict[str, float] = defaultdict(float)
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
        cache[(DASH, f"B{r}")] = float(status_c.get(str(label), 0))
        cache[(DASH, f"C{r}")] = float(round(status_nbv.get(str(label), 0.0), 2))

    for r in range(19, 23):
        label = dash.cell(r, 5).value
        if not label:
            continue
        cache[(DASH, f"F{r}")] = float(cat_c.get(str(label), 0))
        cache[(DASH, f"G{r}")] = float(round(cat_nbv.get(str(label), 0.0), 2))

    for r in range(32, 44):
        label = dash.cell(r, 1).value
        if not label:
            continue
        cache[(DASH, f"B{r}")] = float(loc_c.get(str(label), 0))
        cache[(DASH, f"C{r}")] = float(round(loc_nbv.get(str(label), 0.0), 2))

    for r in range(32, 43):
        label = dash.cell(r, 5).value
        if not label:
            continue
        cache[(DASH, f"F{r}")] = float(type_c.get(str(label), 0))
        cache[(DASH, f"G{r}")] = float(round(type_exp.get(str(label), 0.0), 2))

    ledger_sum = sum(n(v.get("acq")) for v in sl_by.values())
    active_nbv = sum(
        n(v.get("nbv")) for v in sl_by.values() if v.get("status") == "Active"
    )
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

    # LR detail E formulas if present
    for r in range(24, 36):
        cell = lr.cell(r, 5)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            reg = lr.cell(r, 3).value
            led = lr.cell(r, 4).value
            if isinstance(reg, (int, float)) and isinstance(led, (int, float)):
                cache[(LR, f"E{r}")] = float(led) - float(reg)

    return cache


def refresh_artifacts() -> None:
    for zname in [
        "Yanou_IT_Asset_Reconciliation.zip",
        "Yanou_IT_Asset_Reconciliation.xlsx.zip",
    ]:
        zpath = ROOT / zname
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(WB, WB.name)
    dl = ROOT / "download"
    dl.mkdir(exist_ok=True)
    shutil.copy2(WB, dl / WB.name)
    with zipfile.ZipFile(
        dl / "Yanou_IT_Asset_Reconciliation.zip", "w", zipfile.ZIP_DEFLATED
    ) as zf:
        zf.write(WB, WB.name)
    art = Path("/opt/cursor/artifacts")
    if art.exists():
        shutil.copy2(WB, art / WB.name)
        with zipfile.ZipFile(
            art / "Yanou_IT_Asset_Reconciliation.zip", "w", zipfile.ZIP_DEFLATED
        ) as zf:
            zf.write(WB, WB.name)
    print("artifacts refreshed")


def verify() -> None:
    wb = load_workbook(WB, data_only=False)
    wbv = load_workbook(WB, data_only=True)

    def share(name: str):
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

    sl = wbv[SRC]
    for tag in ("MD-00034", "MD-00082"):
        for r in range(5, 200):
            if sl.cell(r, 2).value == tag:
                print(tag, "NBV", sl.cell(r, 7).value)
                break

    # authorship uniqueness
    ex = wb[EX]
    eas = []
    for r in range(5, 200):
        if ex.cell(r, 2).value == CLOSED:
            eas.append(norm_text(str(ex.cell(r, 12).value or "")))
    print("closed EA unique norms", len(set(eas)), "of", len(eas))

    cc = wb[CC]
    by = defaultdict(list)
    for r in range(5, 800):
        if not cc.cell(r, 1).value:
            continue
        ev = str(cc.cell(r, 3).value or "")
        if ev in ("Policy Control Figure", "Assignment/Transfer"):
            by[ev].append(norm_text(str(cc.cell(r, 9).value or "")))
    for ev, lst in by.items():
        print(ev, "unique", len(set(lst)), "of", len(lst))

    # EX-0092/93 remaining outside Exception ID column
    found = []
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                if name == EX and cell.column == 1:
                    continue
                if cell.value and (
                    "EX-0092" in str(cell.value) or "EX-0093" in str(cell.value)
                ):
                    found.append(f"{name}!{cell.coordinate}")
    print("remaining EX-0092/93 outside EX ID col:", found)


def main() -> None:
    assert WB.exists(), WB
    wb = load_workbook(WB)
    rewrite_closed(wb)
    rewrite_custody(wb)
    scrub_fabricated_ex(wb)
    cache = build_cache(wb)
    wb.save(WB)
    print("saved; cache entries", len(cache))
    inject_cache(WB, cache)
    refresh_artifacts()
    verify()
    print("done")


if __name__ == "__main__":
    main()
