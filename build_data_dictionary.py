#!/usr/bin/env python3
"""
build_data_dictionary.py  --  data dictionary for AFRO_admin_data.xlsx
=====================================================================

Builds AFRO_admin_data_dictionary.xlsx (to upload to SharePoint next to the admin data) from the admin data
file itself, so the completeness, ranges and examples always match the data.

    python build_data_dictionary.py AFRO_admin_data.xlsx AFRO_admin_data_dictionary.xlsx

The descriptions are written in this file (COLUMNS below). If a column is added to the admin data, the script
stops and asks for its description, so the dictionary never falls behind the data.
"""
import re
import sys

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

SHEET = "Sheet1"

# --------------------------------------------------------------------------------------------- descriptions
VACC = [  # (suffix, long name)
    ("Polio", "polio"), ("Measles", "measles"), ("YF", "yellow fever"), ("Tetanus", "tetanus"),
    ("Mening", "meningitis"), ("VitA", "vitamin A"), ("Deparasit", "deworming"),
    ("Albendazole", "albendazole"), ("MII", "insecticide-treated nets (MII)"),
    ("AntMal", "antimalarial drugs"), ("Cholera", "cholera"), ("LLIN", "long-lasting insecticidal nets"),
    ("IPV", "IPV"), ("fIPV", "fractional IPV"), ("Other_A", "other intervention A"), ("Other_B", "other intervention B"),
]
GROUPS = {
    "id": "Identification", "integ": "Integration with other interventions", "pop": "Target population",
    "age": "Target age group", "polio": "Polio vaccination", "other": "Other interventions",
    "aefi": "Safety and evaluation", "campaign": "Campaign / response", "coadmin": "Co-administration (nOPV2 + bOPV)",
    "tech": "Technical / helper",
}
COLUMNS = {}


def add(name, group, desc, notes=""):
    COLUMNS[name] = (GROUPS[group], desc, notes)


add("Ctry_Code", "id", "Country code (3 letters, e.g. GUI, NIE, DRC).", "43 codes for 38 country names: 10 countries have two codes (e.g. DRC, Kenya, Tanzania).")
add("Country", "id", "Country name in capitals.", "Some countries appear under two names (CAMEROON / CAMEROUN, CENTRAL AFRICAN REPUBLIC / RÉPUBLIQUE CENTRAFRIQUE): group them before counting countries.")
add("Province", "id", "First administrative level below the country (province, region, state).", "Names are cleaned against the geographic reference; keep spelling as is when joining to the shapefiles.")
add("District", "id", "Health district / second administrative level in which the SIA took place. One row = one district in one campaign round and vaccine.", "Empty for a few records.")
add("Admin_3", "id", "Third administrative level (sub-district), when the data is reported at that level.", "Almost always empty (21 records).")
for suf, nm in VACC:
    add("Integ" + suf, "integ", "Whether %s was given together with the polio SIA (Yes / No)." % nm,
        "Values written in different ways (Yes, yes, YES, No, NO): compare case-insensitively." if suf in ("Polio", "Measles") else "")
add("Main_Act", "integ", "Main activity of the campaign: Polio, or Polio_MR_Integrated (polio with measles-rubella).", "'polio' also appears in lower case.")
add("TotPop", "pop", "Total population of the district used for planning the campaign.", "Can be smaller than the vaccinated total: see the coverage rules.")
for suf, nm in VACC:
    add("Pop" + suf, "pop", "Target population for %s in the district." % nm,
        "PopPolio is the denominator of CV_nOPV2 and CV_bOPV. Corrected where clearly wrong, see sheet Coverage rules." if suf == "Polio" else "")
for suf, nm in VACC:
    add("AgeG" + suf, "age", "Age group targeted for %s." % nm,
        "Polio: mostly 0-59m; also 0-10y, 0-14y, 0-119m. Written in several ways (0-119m / 0-119mth)." if suf == "Polio" else
        ("Contains 0 when measles was not part of the campaign." if suf == "Measles" else ""))
add("Nbr0dosesVaccPolio_0_11M", "polio", "Children aged 0-11 months vaccinated who had never received a polio dose (zero-dose children).")
add("Nbr_1doses_Plus_VacPolio_0_11M", "polio", "Children aged 0-11 months vaccinated who had already received 1 or more polio doses.")
add("Nbr0dosesVacPolio_12_59M", "polio", "Children aged 12-59 months vaccinated who had never received a polio dose (zero-dose children).")
add("Nbr_1doses_Plus_VacPolio_12_59M", "polio", "Children aged 12-59 months vaccinated who had already received 1 or more polio doses.")
add("Total_Nbr_0doseVaccPolio_0_59M", "polio", "Zero-dose children vaccinated, 0-59 months (sum of the two age groups).")
add("Prop_0dosesPolio", "polio", "Proportion of the vaccinated children who were zero-dose.", "Ratio, not %. A few values above 1 are data-entry errors.")
add("Total_Nbr_1dose_Plus_vaccPolio_0_59M", "polio", "Children with 1 or more previous doses vaccinated, 0-59 months (sum of the two age groups).")
add("TotalNbrVaccPolio", "polio", "Total number of children vaccinated with polio vaccine in the district and round, as reported. For a co-administered record (Vaccine_type 'nOPV2 & bOPV') it is the larger of Tot_Vacc_nOPV2 and Total_Vacc_bOPV, so a child vaccinated with both vaccines is counted once.",
    "Reported value; about 60% of records equal zero-dose + one-or-more-dose children. Not the sum of the two vaccines. Placed just before Response.")
add("CVPolio", "polio", "Polio vaccination coverage as stored in the source (ratio: 1.00 = 100%).", "Not always equal to TotalNbrVaccPolio / PopPolio. Used to correct PopPolio where it was clearly wrong (coverage above 1.5 or below 0.5). Some values are above 1.5.")
add("Doses.UsedPolio", "polio", "Polio vaccine doses used.", "Almost always empty.")
add("WastRPolio", "polio", "Polio vaccine wastage rate.", "Mostly empty.")
add("NbrVacMeasles", "other", "Children vaccinated against measles during the campaign.")
add("CVMeasles", "other", "Measles vaccination coverage (ratio).")
add("ZeroDoseMeasles", "other", "Zero-dose children among those vaccinated against measles.")
add("PropZeroDoseMeasles", "other", "Proportion of zero-dose children among those vaccinated against measles.")
add("UsedMeasles", "other", "Measles vaccine doses used.")
add("WastRMeasles", "other", "Measles vaccine wastage rate.")
add("NbrVacYF", "other", "Children vaccinated against yellow fever.")
add("CVYF", "other", "Yellow fever coverage (ratio).")
add("ZeroDoseYF", "other", "Zero-dose children among those vaccinated against yellow fever.")
add("PropZeroDoseYF", "other", "Proportion of zero-dose children, yellow fever.")
add("UsedYF", "other", "Yellow fever doses used.")
add("WastRYF", "other", "Yellow fever wastage rate.")
add("NbrVacTetanus_1st_Dose", "other", "People vaccinated with the 1st tetanus dose.")
add("NbrVacTetanus_2nd_Dose", "other", "People vaccinated with the 2nd tetanus dose.")
add("NbrVacTetanus_3rd_Dose", "other", "People vaccinated with the 3rd tetanus dose.")
add("TotalNbrVacTetanus", "other", "Total vaccinated against tetanus.")
add("CVTetanus", "other", "Tetanus coverage (ratio).")
add("UsedTetanus", "other", "Tetanus doses used.")
add("WastRTetanus", "other", "Tetanus wastage rate.")
for suf, nm, nbr in [("Mening", "meningitis", "NbrVacMening"), ("VitA", "vitamin A", "NbrVacVitA"), ("Deparasit", "deworming", "NbrVacDeparasit"),
                     ("Albendazole", "albendazole", "NbrVacAlbendazole"), ("MII", "insecticide-treated nets", "NbrVacMII"),
                     ("AntMal", "antimalarial drugs", "NbrVacAntMal")]:
    add(nbr, "other", "People reached with %s." % nm)
    add("CV" + suf, "other", "Coverage for %s (ratio)." % nm)
    add("Used" + suf, "other", "Quantity of %s used." % nm)
    add("WastR" + suf, "other", "Wastage rate, %s." % nm)
add("ZeroDoseCholera", "other", "Zero-dose children among those vaccinated against cholera.")
add("NbrVacCholera", "other", "People vaccinated against cholera.")
add("PropZeroDoseCholera", "other", "Proportion of zero-dose children, cholera.")
add("CVCholera", "other", "Cholera vaccination coverage (ratio).")
add("UsedCholera", "other", "Cholera doses used.")
add("WastRCholera", "other", "Cholera wastage rate.")
add("ZeroLLIN", "other", "Households without any net reached by the net distribution.")
add("NbrLLINDistributed", "other", "Long-lasting insecticidal nets distributed.")
add("PropZeroLLIN", "other", "Proportion of households without a net.")
add("CVLLIN", "other", "Net distribution coverage (ratio).")
add("UsedLLIN", "other", "Nets used.")
add("WastRLLIN", "other", "Net wastage rate.")
add("Total_IPV1_0_6wk", "other", "Children given the 1st IPV dose, 0-6 weeks.")
add("Total_IPV2_6_14wk", "other", "Children given the 2nd IPV dose, 6-14 weeks.")
add("TotalNbrVaccIPV", "other", "Total children vaccinated with IPV.")
add("CVIPV", "other", "IPV coverage (ratio).")
add("UsedIPV", "other", "IPV doses used.")
add("WastRIPV", "other", "IPV wastage rate.")
add("Total_fIPV", "other", "Children vaccinated with fractional IPV.")
add("CVfIPV", "other", "fIPV coverage (ratio).")
add("UsedfIPV", "other", "fIPV doses used.")
add("WastRfIPV", "other", "fIPV wastage rate.")
for x in ("A", "B"):
    add("NbrVacOther_" + x, "other", "People reached by other intervention %s." % x)
    add("CVOther_" + x, "other", "Coverage of other intervention %s (ratio)." % x)
    add("UsedOther_" + x, "other", "Quantity used, other intervention %s." % x)
    add("WastROther_" + x, "other", "Wastage rate, other intervention %s." % x)
add("AEFIMin", "aefi", "Minor adverse events following immunisation (AEFI) reported.", "Mostly empty.")
add("AEFISerious", "aefi", "Serious adverse events following immunisation reported.", "Mostly empty (58 records).")
add("Death", "aefi", "Deaths reported after vaccination.", "Empty.")
add("NumberMonitored", "aefi", "Children checked in the post-campaign evaluation (monitoring).", "Empty. Independent monitoring data is in the IM repository.")
add("NumberVaccinated_Evaluation", "aefi", "Children found vaccinated in the post-campaign evaluation.", "Empty.")
add("Prop_Missed_Eval", "aefi", "Proportion of children found missed in the post-campaign evaluation.", "Empty.")
add("SIA_date", "campaign", "Start date of the SIA round in the district.", "Date.")
add("Date_Report", "campaign", "Date the report was received.", "Date. Filled for about a quarter of the records.")
add("SIA_Integrated", "campaign", "Whether the SIA was integrated with another intervention (Yes / No).", "Written as Yes / No / NO.")
add("TypeofActivity", "campaign", "Type of activity: Reactive, Preventive, SIA, NIDs, Outbreak Response, Polio_MR_Integrated, ...", "Filled for about a quarter of the records. 'Poilo_MR_Integrated' is a spelling error for Polio_MR_Integrated.")
add("EpidNumber", "campaign", "EPID number of the outbreak case or environmental sample that triggered an outbreak response.", "Filled only for response campaigns (about 470 records).")
add("DateOfOnset", "campaign", "Date of onset of the case that triggered the response.", "Stored in mixed formats (dates and Excel serial numbers saved as text, for example 45828).")
add("Year", "campaign", "Year of the campaign (year of SIA_date).")
add("Round_Add", "campaign", "Campaign round: Rnd0 to Rnd6.", "Round of the response, not of the calendar year.")
add("Tot_Vacc_nOPV2", "coadmin", "Children vaccinated with nOPV2 in the district and round.", "0 when no nOPV2 was given. Combined records: same value as TotalNbrVaccPolio unless the country gave a separate figure.")
add("CV_nOPV2", "coadmin", "nOPV2 coverage = Tot_Vacc_nOPV2 / PopPolio (ratio: 1.00 = 100%).", "Above 1.5 means the reported population is probably too low. Such records are listed for the countries.")
add("Total_Vacc_bOPV", "coadmin", "Children vaccinated with bOPV in the district and round.", "0 when no bOPV was given. Combined records: same value as TotalNbrVaccPolio unless the country gave a separate figure.")
add("CV_bOPV", "coadmin", "bOPV coverage = Total_Vacc_bOPV / PopPolio (ratio).", "Same caution as CV_nOPV2.")
add("Vaccine_type", "coadmin", "Vaccine used: nOPV2, bOPV, mOPV2, MR, or 'nOPV2 & bOPV' when both were given in the same campaign (co-administration).", "See sheet Vaccine_type. Placed just before Response.")
add("Response", "campaign", "Name of the response / campaign, e.g. GUI-2025-01-NID_bOPV-nOPV (country code, year, number, type, vaccine).", "Empty for about 1,900 older records. A campaign is co-administered when the name contains both 'nOPV' and 'bOPV'.")
add("Index", "tech", "Running record number inside the data-entry batch of a campaign (1 = first row).", "Internal; no analytical meaning.")
add("SIAMTH", "tech", "Month of the SIA (1-12).", "Same as the month of SIA_date for 98% of records.")
add("check_dp", "tech", "Helper key used to find duplicated records: formula joining Country, Province, District, SIA_date, TotalNbrVaccPolio and Response.", "Internal. A live formula in the admin data; it follows the columns, so keep it when columns are added.")
add("AFRO_block", "tech", "Programme block of the country: WA (West Africa), LCB (Lake Chad Basin), ESA (East and Southern Africa), DRC, CA (Central Africa).", "'LBC' appears for 25 records of the Central African Republic: it should be LCB.")
add("prov_candidates", "tech", "Helper column used when the province names were repaired.", "Empty.")

VACCINE_TYPES = [
    ("nOPV2", "Novel oral polio vaccine type 2 given alone. One record per district and round."),
    ("bOPV", "Bivalent oral polio vaccine (types 1 and 3) given alone. One record per district and round."),
    ("nOPV2 & bOPV", "Co-administration: nOPV2 and bOPV in the same campaign. One record per district and round, with both vaccine totals."),
    ("nOPV2_bOPV_IPV", "Co-administration of nOPV2, bOPV and IPV (Nigeria, Kebbi, 2026). Treated as a combined record."),
    ("mOPV2", "Monovalent oral polio vaccine type 2."),
    ("MR", "Measles-rubella vaccine given in an integrated polio / MR campaign. TotalNbrVaccPolio is 0 on these records; use NbrVacMeasles."),
]

COADMIN = [
    ("How to recognise a co-administered campaign", "The Response name contains both 'nOPV' and 'bOPV', or Vaccine_type already holds both vaccines."),
    ("Reported as separate records", "An nOPV2 record and a bOPV record of the same district, round and Response are merged into one record: Tot_Vacc_nOPV2 and Total_Vacc_bOPV take the two totals, Vaccine_type = 'nOPV2 & bOPV'."),
    ("Reported as one combined record", "Tot_Vacc_nOPV2 = Total_Vacc_bOPV = TotalNbrVaccPolio by default (the countries are asked to confirm the real split)."),
    ("TotalNbrVaccPolio of a co-administered record", "The larger of Tot_Vacc_nOPV2 and Total_Vacc_bOPV (the same children are vaccinated with both, so the totals are not added)."),
    ("Not co-administered", "Vaccine_type unchanged. TotalNbrVaccPolio goes in the column of the vaccine used (nOPV2 or bOPV); the other column stays empty."),
    ("Single record in a co-administered Response", "Where other districts report separate nOPV2 and bOPV records but a district has only nOPV2, it is kept as nOPV2 only (no bOPV was reported there)."),
    ("Coverage", "CV_nOPV2 = Tot_Vacc_nOPV2 / PopPolio and CV_bOPV = Total_Vacc_bOPV / PopPolio."),
    ("Vaccine_type clean-up", "'Nopv2' and 'nOPV3' to 'nOPV14' (drag-fill errors) were corrected to 'nOPV2'; 'bopv' to 'bOPV'."),
]

QUALITY = [
    ("Coverage above 1.5", "More children vaccinated than the target population: the population is probably too low. Records are listed for the countries to correct."),
    ("PopPolio from CVPolio", "Where vaccinated / PopPolio is above 1.5 and the stored CVPolio is plausible, PopPolio = vaccinated / CVPolio (whole number). Also when the coverage is above 5 and CVPolio is less than half of it."),
    ("PopPolio from CVPolio (coverage below 0.5)", "Where vaccinated / PopPolio is below 0.5 and the stored CVPolio is plausible (0.5 to 1.5), PopPolio = vaccinated / CVPolio (whole number). In these records PopPolio was usually the total population or a province figure repeated in every district."),
    ("PopPolio from the district's other rounds", "Where coverage is still above 1.5, PopPolio = median of the same district's other polio records in the same year, when this gives a coverage between 0.5 and 1.5."),
    ("PopPolio from a model", "Where it is still above 1.5, PopPolio is estimated by a model trained on the district's own history (typical error about 10%), when this gives a coverage between 0.5 and 1.5. These are estimates, not reported data; countries are asked to confirm them."),
    ("Whole numbers", "PopPolio is always a whole number of children."),
    ("Records still above 1.5", "Left as reported (no reliable history). Listed for the countries."),
    ("Audit", "Every changed record is listed in the audit workbook produced with the data (sheet PopPolio_corrected, with the old value and the reason)."),
]

CODES = [
    ("Round_Add", "Rnd0 to Rnd6", "Campaign round."),
    ("AFRO_block", "WA", "West Africa"), ("AFRO_block", "LCB", "Lake Chad Basin"), ("AFRO_block", "ESA", "East and Southern Africa"),
    ("AFRO_block", "DRC", "Democratic Republic of the Congo"), ("AFRO_block", "CA", "Central Africa"),
    ("Integ* flags", "Yes / No", "Written in several ways (Yes, yes, YES, No, NO): compare case-insensitively."),
    ("Coverage columns", "ratio", "1.00 = 100%. Values are not capped."),
    ("Dates", "date", "SIA_date, Date_Report, DateOfOnset."),
]


def kind(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return "Date"
    if pd.api.types.is_numeric_dtype(s):
        return "Number"
    return "Text"


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    src, dst = sys.argv[1], sys.argv[2]
    print("reading", src, flush=True)
    df = pd.read_excel(src, sheet_name=SHEET)
    missing = [c for c in df.columns if c not in COLUMNS]
    if missing:
        sys.exit("No description for these columns, add them to COLUMNS: " + ", ".join(missing))
    extra = [c for c in COLUMNS if c not in df.columns]
    if extra:
        print("note: described but not in the data:", ", ".join(extra))

    rows = []
    for j, c in enumerate(df.columns, start=1):
        s = df[c]
        grp, desc, notes = COLUMNS[c]
        k = kind(s)
        nn = s.notna().sum()
        ex = s.dropna().iloc[0] if nn else ""
        if isinstance(ex, pd.Timestamp):
            ex = ex.strftime("%Y-%m-%d")
        rng = ""
        if k == "Number" and nn:
            rng = "%s to %s" % (format(s.min(), ",.4g"), format(s.max(), ",.10g"))
        elif k == "Text" and nn and s.nunique() <= 12:
            rng = " | ".join(map(str, s.value_counts().index[:12]))
        elif k == "Date" and nn:
            rng = "%s to %s" % (s.min().strftime("%Y-%m-%d"), s.max().strftime("%Y-%m-%d"))
        rows.append({"No": j, "Column": c, "Excel column": L(j), "Group": grp, "Description": desc, "Type": k,
                     "Filled (%)": round(100 * nn / len(df), 1), "Distinct values": int(s.nunique()),
                     "Range / values": rng, "Example": ex, "Notes / known issues": notes})
    dd = pd.DataFrame(rows)

    vt = df["Vaccine_type"].value_counts().rename_axis("Vaccine_type").reset_index(name="Records")
    meaning = dict(VACCINE_TYPES)
    vt["Meaning"] = vt["Vaccine_type"].map(meaning).fillna("")
    vt = pd.concat([vt, pd.DataFrame([{"Vaccine_type": k, "Records": 0, "Meaning": v} for k, v in VACCINE_TYPES
                                      if k not in set(vt["Vaccine_type"])])], ignore_index=True)

    readme = pd.DataFrame({"": [
        "Data dictionary of AFRO_admin_data.xlsx (WHO AFRO administrative vaccination data, polio SIAs).",
        "One row of the data = one district, one campaign round and one vaccine (or one combined nOPV2 & bOPV record).",
        "Records: {:,}   |   Columns: {}   |   Countries: {}   |   Years: {} to {}".format(
            len(df), df.shape[1], df["Country"].nunique(), int(df["Year"].min()), int(df["Year"].max())),
        "Sheets: Data dictionary (every column), Vaccine_type (values), Co-administration (how the new columns are filled), "
        "Coverage rules (data-quality corrections), Codes.",
        "Co-administration columns Tot_Vacc_nOPV2, CV_nOPV2, Total_Vacc_bOPV, CV_bOPV, Vaccine_type and TotalNbrVaccPolio sit just before Response.",
        "Coverage is a ratio (1.00 = 100%). Coverage above 1.5 is not credible and is under review with the countries.",
        "Built with build_data_dictionary.py from the data file; completeness and ranges are computed from the data.",
    ]})

    with pd.ExcelWriter(dst, engine="openpyxl") as xw:
        readme.to_excel(xw, sheet_name="Read me", index=False)
        dd.to_excel(xw, sheet_name="Data dictionary", index=False)
        vt.to_excel(xw, sheet_name="Vaccine_type", index=False)
        pd.DataFrame(COADMIN, columns=["Topic", "Rule"]).to_excel(xw, sheet_name="Co-administration", index=False)
        pd.DataFrame(QUALITY, columns=["Topic", "Rule"]).to_excel(xw, sheet_name="Coverage rules", index=False)
        pd.DataFrame(CODES, columns=["Column", "Value", "Meaning"]).to_excel(xw, sheet_name="Codes", index=False)
        widths = {"Read me": [150], "Data dictionary": [5, 34, 9, 24, 70, 8, 9, 10, 34, 24, 60],
                  "Vaccine_type": [20, 10, 90], "Co-administration": [42, 120], "Coverage rules": [38, 120], "Codes": [20, 20, 70]}
        for name, ws in xw.sheets.items():
            ws.sheet_view.showGridLines = False
            for j, w in enumerate(widths[name], start=1):
                ws.column_dimensions[L(j)].width = w
            if name != "Read me":
                for c in ws[1]:
                    c.font = Font(bold=True, color="FFFFFF")
                    c.fill = PatternFill("solid", fgColor="005C97")
                    c.alignment = Alignment(vertical="center", wrap_text=True)
                ws.freeze_panes = "C2" if name == "Data dictionary" else "A2"
                ws.auto_filter.ref = ws.dimensions
            for row in ws.iter_rows(min_row=1 if name == "Read me" else 2):
                for c in row:
                    c.alignment = Alignment(vertical="top", wrap_text=True)
        # emphasise the co-administration group
        ws = xw.sheets["Data dictionary"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 4).value == GROUPS["coadmin"]:
                for j in range(1, 12):
                    ws.cell(r, j).fill = PatternFill("solid", fgColor="E8F3FA")
    print("written:", dst, "|", len(dd), "columns described")


if __name__ == "__main__":
    main()
