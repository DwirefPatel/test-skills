"""
Turns a tied-out brokerage 1099 (brokerage_1099 RunResult) into input-workbook rows:

  8949  Schedule D: one summary line per basis-reported box (A/D/G/J) with the wash-sale
        W adjustment, plus one line per noncovered / unknown-term lot (B/E/H/K/unknown)
  DIV / INT / OID / 99M  Schedule B and Schedule 1: one row per payer with the box totals

Items Drake needs that these screens don't carry (margin interest, Section 1256, state
subtraction, the digital-asset question...) come back as "manual items" for the tracker.
"""
from datetime import date

FORM_SCREENS = {"DIV": "DIV", "INT": "INT", "OID": "OID", "MISC": "99M"}
STATE_BOX_FIELDS = {
    "DIV": {"14": "box14_state", "15": "box15_state_id", "16": "box16_state_wh"},
    "INT": {"15": "box15_state", "16": "box16_state_id", "17": "box17_state_wh"},
    "OID": {"12": "box12_state", "14": "box14_state_wh"},
    "MISC": {"15": "box15_state_wh", "16": "box16_state"},
}
MANUAL = {
    "margin_interest": ("Margin interest paid", "Form 4952 investment interest (Drake screen 4952)"),
    "section_1256": ("Section 1256 aggregate profit/(loss)", "Form 6781 (Drake screen 6781)"),
    "accrued_interest_paid": ("Accrued interest paid on bond purchases",
                              "Schedule B reduction (adjustment on the INT screen)"),
    "partnership_distribution": ("Publicly traded partnership distributions",
                                 "Request the K-1(s); not entered from the 1099"),
    "state_withholding": ("State tax withheld (supplemental)", "State withholding screen"),
    "other": ("Other supplemental item", "Review"),
}
NOT_DEDUCTIBLE = {"adr_fee", "investment_fee"}


def money_str(d):
    return "" if d is None else f"{d:.2f}"


def date_str(d):
    if isinstance(d, date):
        return d.strftime("%m/%d/%Y")
    if d in (None, "", "N/A"):
        return ""
    return str(d)


def payer_for(pkg, form):
    for p in pkg.payers:
        if any(f.upper().replace("FORM ", "") == form for f in p.forms):
            return p
    return pkg.payers[0] if pkg.payers else None


def build_rows(result, client, source_tag, forms):
    """-> (rows_by_screen_code, manual_items). client: dict(ssn, name, client_id, ts)."""
    from brokerage_1099.analysis import state_treasury_amount

    pkg, an = result.package, result.analysis
    common = {"client_ssn": client["ssn"], "client_id": client.get("client_id", ""),
              "client_name": client.get("name", ""), "ts": client.get("ts", "T"),
              "source_doc": source_tag}
    rows = {code: [] for code in ("8949", "DIV", "INT", "OID", "99M")}
    manual = []

    # ---- Schedule D: 8949 summary lines (basis reported)
    for line in an.entry_lines:
        if not line.lot_count:
            continue
        digital = line.box in "GJ"
        payer = payer_for(pkg, "1099-DA" if digital else "1099-B")
        bt = pkg.box_total(line.box)
        codes = ("W" if line.wash else "") + ("D" if line.market_discount else "")
        adj = line.wash - line.market_discount
        desc = f"{pkg.broker[:34]} - Box {line.box} summary" + (" ordinary" if line.ordinary else "")
        rows["8949"].append({
            **common, "box_8949": line.box, "description": desc[:60],
            "date_acquired": "VARIOUS", "date_sold": "VARIOUS",
            "proceeds": money_str(line.proceeds), "cost": money_str(line.cost),
            "adj_code": codes, "adj_amount": money_str(adj) if adj else "",
            "entry_type": "Summary (basis reported to IRS)"
                          + (" - ordinary portion" if line.ordinary else ""),
            "lot_ref": f"{line.box}-SUMMARY" + ("-ORD" if line.ordinary else ""),
            "broker_gain": money_str(line.gain),
            "payer_name": payer.name if payer else "", "payer_ein": payer.ein if payer else "",
            "source_page": str(bt.source_page) if bt and bt.source_page else "",
        })

    # ---- Schedule D: noncovered / unknown-term lots, one line each
    n = 0
    for box, lots in an.noncovered.items():
        for lot in lots:
            n += 1
            payer = payer_for(pkg, "1099-DA" if lot.is_digital else "1099-B")
            qty = f"{lot.quantity.normalize():f} " if lot.quantity else ""
            unknown = box == "X"
            rows["8949"].append({
                **common, "box_8949": "" if unknown else box,
                "description": f"{qty}{lot.description}"[:60],
                "date_acquired": date_str(lot.date_acquired), "date_sold": date_str(lot.date_sold),
                "proceeds": money_str(lot.proceeds), "cost": money_str(lot.cost),
                "adj_code": "W" if lot.wash_sale else "",
                "adj_amount": money_str(lot.wash_sale) if lot.wash_sale else "",
                "entry_type": "Individual lot (UNKNOWN TERM - client must supply term and basis)"
                              if unknown else "Individual lot (basis NOT reported to IRS)",
                "lot_ref": f"{box}-{(lot.cusip or lot.description)[:14]}-{date_str(lot.date_sold)}-{n}",
                "broker_gain": money_str(lot.broker_gain),
                "payer_name": payer.name if payer else "", "payer_ein": payer.ein if payer else "",
                "source_page": str(lot.source_page or ""),
            })

    # ---- Schedule B / Schedule 1: payer box totals
    for form, screen in FORM_SCREENS.items():
        totals, pages = {}, set()
        for l in pkg.income_lines:
            if l.form == form:
                totals[l.box] = totals.get(l.box, 0) + l.amount
                if l.source_page:
                    pages.add(l.source_page)
        if not any(totals.values()):
            continue
        payer = payer_for(pkg, f"1099-{form}")
        row = {**common, "payer_name": payer.name if payer else "",
               "payer_ein": payer.ein if payer else "",
               "source_page": ",".join(str(p) for p in sorted(pages))}
        names = {f.name for f in forms[screen].fields}
        for box, amt in totals.items():
            if not amt:
                continue
            field = STATE_BOX_FIELDS[form].get(box, f"box{box}")
            if field in names:
                row[field] = money_str(amt)
            else:
                manual.append((f"1099-{form} box {box}", amt,
                               f"Not on the Drake {screen} screen definition - enter manually"))
        rows[screen].append(row)

    # ---- items Drake needs that the screens above don't carry
    for item in pkg.other_items:
        if item.category in NOT_DEDUCTIBLE:
            continue
        label, dest = MANUAL.get(item.category, MANUAL["other"])
        manual.append((item.description or label, item.amount, dest))
    st = state_treasury_amount(pkg)
    if st:
        manual.append(("U.S. Treasury / federal obligation interest", st,
                       "State return subtraction (e.g. Colorado)"))
    if an.digital_asset_question_yes:
        manual.append(("Digital asset transactions", None, "Answer YES to the 1040 digital asset question"))
    if pkg.corrected:
        manual.append(("CORRECTED 1099", None, "If the return was already filed, decide on an amendment"))
    open_flags = sum(1 for f in an.flags if f.status == "Open")
    if open_flags:
        manual.append((f"{open_flags} open flag(s) on the workpaper", None,
                       "Review the 'Flags & Questions' tab before filing"))
    return {k: v for k, v in rows.items() if v}, manual
