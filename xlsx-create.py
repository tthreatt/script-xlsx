import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation
import os

input_dir = '/Users/tthreatt/Desktop/BCBS-SC'
output_file = 'BCBS-251110.xlsx'

csv_files = [f for f in os.listdir(input_dir) if f.endswith('.csv')]

search_layout = [
    {
        "sheet_prefix": "dnpi_license_bcbs_sc_",
        "row": 3,
        "headers": [
            "first_name", "middle_name", "last_name", "organization_name", "monitored_product", "issuer", "license_type", "license_source", "multi_stata",
            "license_category", "verified_first_name", "verified_middle_name", "verified_last_name", "verified_org_name", "verified_license_issued", "verified_license_number",
            "verified_license_status", "verified_license_details", "verified_license_expiration", "calculated_license_status", "abms_moc_status", "abms_renewal_date",
            "abms_duration_type", "abms_reverification_date", "dea_schedules", "dea_license_state"
        ]
    },
    {
        "sheet_prefix": "dnpi_preclusion_BCBS_SC_",
        "row": 101,
        "headers": [
            "monitored_product",
            "first_name",
            "middle_name",
            "last_name",
            "organization_name",
            "dob",
            "ein",
            "address_lines",
            "city",
            "state",
            "postal",
            "general",
            "speciality",
            "business_name",
            "preclusion_npi",
            "preclusion_id",
            "excluded_date",
            "reinstated_date",
            "claim_rejected_date",
            "preclusion_first_name",
            "preclusion_last_name",
            "preclusion_middle_name",
            "preclusion_former_last_name"
        ]
    },
    {
        "sheet_prefix": "dnpi_exclusion_BCBS_SC_",
        "row": 110,
        "headers": [
            "first_name", "middle_name", "last_name", "organization_name", "monitored_product", "akas", "cage", "dbas", "npis", "type", "upin", "source",
            "address_of_residence", "address_line_2_of_residence", "fax_of_residence", "zip_of_residence", "city_of_residence", "state_of_residence", "county_of_residence",
            "country_of_residence", "telephone_of_residence", "comments", "source_id", "speciality", "dob", "duns_numbers", "exclusion_code", "start_date", "exclusion_date",
            "exclusion_term", "reinstate_date", "delisted_date", "classification", "exclusion_notes", "prefix", "suffix", "exclusion_last", "exclusion_first", "exclusion_middle",
            "exclusion_former_last", "exclusion_license_number", "excluding_agency", "provider_number", "exclusion_organization_name"
        ]
    },
    {
        "sheet_prefix": "dnpi_opt_out_BCBS_SC_",
        "row": 120,
        "headers": [
            "monitored_product", "first_name", "middle_name", "last_name", "organization_name", "address_lines", "city", "state", "postal", "speciality",
            "opt_out_id", "optout_npi", "effective_date", "end_date", "optout_first_name", "optout_last_name", "optout_former_last_name", "eligible_to_order_and_refer"
        ]
    },
    {
        "sheet_prefix": "dnpi_npi_BCBS_SC_",
        "row": 130,
        "headers": [
            "first_name", "middle_name", "last_name", "organization_name", "monitored_product", "ein", "nppes_npi", "entity_type", "last_update_date",
            "replacement_npi", "nppes_last_name", "nppes_first_name", "nppes_credentials", "nppes_middle_name", "nppes_name_prefix", "nppes_name_suffix",
            "nppes_organization_name", "is_sole_proprietor", "provider_gender_code", "npi_deactivation_date", "npi_reactivation_date", "npi_deactivation_reason",
            "provider_enumeration_date", "mailing_address_of_residence", "mailing_address_line_2_of_residence", "mailing_fax_of_residence",
            "mailing_zip_of_residence", "mailing_city_of_residence", "mailing_state_of_residence", "mailing_county_of_residence", "mailing_country_of_residence",
            "mailing_telephone_of_residence", "practice_address_of_residence", "practice_address_line_2_of_residence", "practice_fax_of_residence",
            "practice_zip_of_residence", "practice_city_of_residence", "practice_state_of_residence", "practice_county_of_residence", "practice_country_of_residence",
            "practice_telephone_of_residence", "nppes_licenses_taxonomies"
        ]
    }
]

# --- Merge sheets by prefix ---
sheet_prefixes = [block["sheet_prefix"] for block in search_layout]
sheet_data = {prefix: pd.DataFrame() for prefix in sheet_prefixes}
other_sheets = []

for csv_file in csv_files:
    path = os.path.join(input_dir, csv_file)
    df = pd.read_csv(path, dtype=str)
    for prefix in sheet_prefixes:
        if csv_file.startswith(prefix):
            sheet_data[prefix] = pd.concat([sheet_data[prefix], df], ignore_index=True)
            break
    else:
        sheet_name = os.path.splitext(csv_file[:30])[0]
        other_sheets.append((df, sheet_name))

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for prefix, df in sheet_data.items():
        if not df.empty:
            # Write using original filename for uniqueness if >1 file per prefix
            # Or use prefix as default
            sheet_name = None
            prefixed_files = [f for f in csv_files if f.startswith(prefix)]
            if prefixed_files:
                sheet_name = os.path.splitext(prefixed_files[0])[0]  # Take first by default
            else:
                sheet_name = prefix.rstrip("_")
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    # Write other sheets
    for df, sheet_name in other_sheets:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    pd.DataFrame().to_excel(writer, index=False, sheet_name='Search_Tab')

# --- Setup Excel workbook for headers, formulas, and validation drop-down ---
wb = openpyxl.load_workbook(output_file)
ws = wb['Search_Tab']

def find_sheet_by_prefix(wb, prefix):
    for name in wb.sheetnames:
        if name.startswith(prefix):
            return name
    raise Exception(f"No sheet found with prefix '{prefix}'")

# Move Search_Tab to be the first sheet
if "Search_Tab" in wb.sheetnames and wb.sheetnames[-1] == "Search_Tab":
    wb.move_sheet("Search_Tab", offset=-len(wb.sheetnames)+1)

license_sheetname = None

for block in search_layout:
    prefix = block["sheet_prefix"]
    try:
        sheetname = find_sheet_by_prefix(wb, prefix)
    except Exception as e:
        print(f"Warning: {e}")
        continue
    row = block["row"]
    headers = block["headers"]
    # Write headers
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=row, column=col_idx, value=header).font = openpyxl.styles.Font(bold=True)
    # Determine which formula to use
    if prefix == "dnpi_license_bcbs_sc_":
        formula = f"FILTER('{sheetname}'!B2:AA70000, ('{sheetname}'!A2:A70000 = $A$1) * (($B$1 = \"All\") + ('{sheetname}'!G2:G70000 = $B$1)), \"\")"
        license_sheetname = sheetname
    else:
        formula = f"FILTER('{sheetname}'!B2:AS70000, '{sheetname}'!A2:A70000 = $A$1, \"\")"
    ws.cell(row=row + 1, column=1, value=formula)

# Before using license_sheetname, ADD:
if license_sheetname is None:
    raise Exception("No license sheet name found after writing formulas. Check your file prefixes and written worksheet names.")
license_ws = wb[license_sheetname]

# --- Issuer drop-down validation for B1, robust handling for big lists ---
search_ws = ws
license_ws = wb[license_sheetname]
search_ws['B1'] = "All"  # Default filter set

issuers = set()
for row in license_ws.iter_rows(min_row=2, min_col=7, max_col=7, values_only=True):  # col G
    val = row[0]
    if val and val.strip():
        issuers.add(val.strip())
issuer_list = sorted(issuers)
issuer_list.insert(0, "All")  # Add 'All' at the top

if sum(len(i) for i in issuer_list) + len(issuer_list) - 1 <= 255:
    dv = DataValidation(type="list", formula1='"{}"'.format(",".join(issuer_list)), allow_blank=True)
    search_ws.add_data_validation(dv)
    dv.add(search_ws["B1"])
else:
    if 'IssuerList' not in wb.sheetnames:
        wb.create_sheet('IssuerList')
    issuer_ws = wb['IssuerList']
    for idx, issuer in enumerate(issuer_list, start=1):
        issuer_ws.cell(row=idx, column=1, value=issuer)
    dv = DataValidation(
        type="list",
        formula1='=IssuerList!$A$1:$A${}'.format(len(issuer_list)),
        allow_blank=True
    )
    search_ws.add_data_validation(dv)
    dv.add(search_ws["B1"])
    issuer_ws.sheet_state = 'hidden'

wb.save(output_file)