import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime
import os

input_dir = '/Users/tthreatt/Desktop/BCBS-SC'
output_file = 'BCBS-251110-new.xlsx'
today_str = datetime.now().strftime('%y-%m-%d')  # e.g., '25-11-03'

csv_files = [f for f in os.listdir(input_dir) if f.endswith('.csv')]
print("*** All files being processed: ***")
for fname in csv_files:
    print("-", fname)

search_layout = [
    {
        "sheet_prefix": "license_",
        "row": 3,
        "headers": [
            "first_name", "middle_name", "last_name", "organization_name", "monitored_product", "issuer", "license_type", "license_source", "multi_stata",
            "license_category", "verified_first_name", "verified_middle_name", "verified_last_name", "verified_org_name", "verified_license_action", "verified_license_issued", "verified_license_number",
            "verified_license_status", "verified_license_details", "verified_license_expiration", "calculated_license_status", "board_action_text", "abms_moc_status", "abms_renewal_date",
            "abms_duration_type", "abms_reverification_date", "dea_schedules", "dea_license_state"
        ]
    },
    {
        "sheet_prefix": "npi_",
        "row": 101,
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
    },
    {
        "sheet_prefix": "exclusion_",
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
        "sheet_prefix": "preclusion_",
        "row": 120,
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
        "sheet_prefix": "opt_out_",
        "row": 130,
        "headers": [
            "monitored_product", "first_name", "middle_name", "last_name", "organization_name", "address_lines", "city", "state", "postal", "speciality",
            "opt_out_id", "optout_npi", "effective_date", "end_date", "optout_first_name", "optout_last_name", "optout_former_last_name", "eligible_to_order_and_refer"
        ]
    },
    {
        "sheet_prefix": "ofac_",
        "row": 140,
        "headers": [
            "npi", "first_name", "middle_name", "last_name", "organization_name", "monitored_product", "monitored_start_date", "monitored_end_date",
            "ofac_organization_name", "ofac_first_name", "ofac_middle_name", "ofac_last_name", "ofac_suffix", "ofac_date_of_birth", "ofac_npi",
            "ofac_tin", "ofac_specialty", "ofac_country", "ofac_date", "ofac_reinstate_date", "ofac_terms", "ofac_comments"
        ]
    },
    {
        "sheet_prefix": "ssdmf_",
        "row": 150,
        "headers": [
            "npi", "first_name", "middle_name", "last_name", "ssn", "monitored_product", "monitored_start_date", "monitored_end_date", "status",
            "last_verify_time", "verification_result", "ssdmf_first", "ssdmf_last", "ssdmf_date_of_birth", "ssdmf_date_of_death"
        ]
    },
    {
        "sheet_prefix": "missing_",
        "special_match": "npi_missing_license_creds",
        "row": 160,
        "headers": ["npi"]
    }  
]

SHORT_PREFIX_MAP = {
    'dnpi_preclusion_bcbs_sc_': 'preclusion_',
    'dnpi_exclusion_bcbs_sc_': 'exclusion_',
    'dnpi_license_bcbs_sc_': 'license_',
    'dnpi_opt_out_bcbs_sc_': 'opt_out_',
    'dnpi_ssdmf_bcbs_sc_': 'ssdmf_',
    'dnpi_ofac_bcbs_sc_': 'ofac_',
    'dnpi_npi_bcbs_sc_': 'npi_'
}

# --- INVERT the map to get logical -> file prefix mapping ---
LOGICAL_TO_FILE_PREFIX = {v: k for k, v in SHORT_PREFIX_MAP.items()}
# LOGICAL_TO_FILE_PREFIX['license_'] => 'dnpi_license_bcbs_sc_'

sheet_data = {}
sheetname_lookup = {}
sheet_rowcount = {}
sheet_colcount = {}

for block in search_layout:
    prefix = block["sheet_prefix"]
    special = block.get("special_match", "")
    row = block["row"]
    file_prefix = LOGICAL_TO_FILE_PREFIX.get(prefix, prefix)
    for csv_file in csv_files:
        match = False
        if special:
            match = special in csv_file
            if not match:
                continue
        else:
            match = csv_file.startswith(file_prefix) and ('npi_missing_license_creds' not in csv_file)
            if not match:
                continue

        path = os.path.join(input_dir, csv_file)
        base_sheet_name = f"{prefix}{today_str}"
        sheet_name = base_sheet_name
        c = 1
        while sheet_name in sheet_data:
            sheet_name = f"{base_sheet_name}_{c}"
            c += 1

        df = pd.read_csv(path, dtype=str)
        sheet_data[sheet_name] = df
        # Save real Excel end row and col for each sheet
        sheet_rowcount[sheet_name] = len(df) + 1   # data rows + header row = Excel last row #
        sheet_colcount[sheet_name] = get_column_letter(df.shape[1])  # Excel letter for last col
        if prefix not in sheetname_lookup:
            sheetname_lookup[prefix] = sheet_name

print("\nSheet mapping:")
for k, v in sheetname_lookup.items():
    print(f"  {k} --> {v}")

print("\nSheet names that will be created:")
for sheet_name in sheet_data.keys():
    print("-", sheet_name)

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, df in sheet_data.items():
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    pd.DataFrame().to_excel(writer, index=False, sheet_name='Search_Tab')

wb = openpyxl.load_workbook(output_file)
ws = wb['Search_Tab']

if "Search_Tab" in wb.sheetnames and wb.sheetnames[-1] == "Search_Tab":
    wb.move_sheet("Search_Tab", offset=-len(wb.sheetnames)+1)

for block in search_layout:
    prefix = block["sheet_prefix"]
    headers = block["headers"]
    row = block["row"]
    sheetname = sheetname_lookup.get(prefix)
    if not sheetname or sheetname not in wb.sheetnames:
        print(f"Warning: Sheet for prefix '{prefix}' not found in workbook. Skipping.")
        continue

    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=row, column=col_idx, value=header).font = Font(bold=True)
    
    filter_end_col = sheet_colcount.get(sheetname, 'A')       # true Excel end col (e.g., 'Z')
    filter_end_row = sheet_rowcount.get(sheetname, 2)         # true Excel end row (e.g., 105)

    # Example Excel range you can use in your FILTER formulas:
    excel_range = f'B2:{filter_end_col}{filter_end_row}'

    if prefix == "license_":
        formula = (
            f'FILTER(\'{sheetname}\'!{excel_range}, '
            f'(\'{sheetname}\'!A2:A{filter_end_row}=$A$1)*(($B$1="All")+'
            f'(\'{sheetname}\'!G2:G{filter_end_row}=$B$1)), "")'
        )
    else:
        formula = (
            f'FILTER(\'{sheetname}\'!{excel_range}, \'{sheetname}\'!A2:A{filter_end_row}=$A$1, "")'
        )
    ws.cell(row=row + 1, column=1, value=formula)

license_sheetname = sheetname_lookup.get("license_")
if license_sheetname and license_sheetname in wb.sheetnames:
    search_ws = ws
    license_ws = wb[license_sheetname]
    search_ws['B1'] = "All"
    issuers = set()
    for row in license_ws.iter_rows(min_row=2, min_col=7, max_col=7, values_only=True):
        val = row[0]
        if val and val.strip():
            issuers.add(val.strip())
    issuer_list = sorted(issuers)
    issuer_list.insert(0, "All")
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
            formula1=f'=IssuerList!$A$1:$A${len(issuer_list)}',
            allow_blank=True
        )
        search_ws.add_data_validation(dv)
        dv.add(search_ws["B1"])
        issuer_ws.sheet_state = 'hidden'
else:
    print("Warning: License sheet not found. Issuer drop-down will be skipped.")

wb.save(output_file)
print("Go and Filter!")