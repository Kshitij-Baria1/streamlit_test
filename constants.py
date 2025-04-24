import configparser
from pathlib import Path

config = configparser.ConfigParser()
properties_file_path = Path(__file__).resolve().parent / "config.properties"
config.read(properties_file_path)

ORIGINAL_EXCEL_ROOT_FOLDER = config.get("FILE_PATH", "ORIGINAL_EXCEL_ROOT_FOLDER")

EXPECTED_HEADERS = [
    "License Name",
    "Entity Name",
    "Entity Jurisdiction",
    "License Type Code",
    "License jurisdiction type",
    "License jurisdiction name",
    "Department name",
    "License Number",
    "MFR Template",
    "Does license expire?",
    "expiration date",
    "renewal frequency (years)",
    "does license number change when renews?",
    "renewal fee",
    "submit by date",
    "manager reminder",
    "client reminder",
    "status",
    "filed date",
    "approved date",
    "Account Number",
    "Renewal Manager",
    "License Manager",
    "Partner Order Number",
    "Payment Method",
    "Filing Fee",
    "Research Address",
    "Research City",
    "Research County",
    "Research State",
    "Research Postal Code",
    "Research Country",
    "Expedite?"
]

USA_STATE_ABBREVIATIONS = [
    'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA',
    'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
    'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
    'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
    'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY'
]

