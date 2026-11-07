"""Settings for the MIS pack. Edit this file, not the pipeline, when the business changes.

Most routine changes land here:
  - a branch renames a column  -> add the new heading to COLUMN_ALIASES
  - a threshold changes        -> update THRESHOLDS
  - a new service line starts  -> add it to SERVICES and SERVICE_ALIASES
"""

from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_DIR = PROJECT_ROOT / "data" / "input"
BRANCH_MASTER = PROJECT_ROOT / "data" / "reference" / "branch_master.csv"
HISTORY_FILE = PROJECT_ROOT / "data" / "history" / "clean_history.csv"
OUTPUT_DIR = PROJECT_ROOT / "output"

SERVICES = ["Express", "Standard"]

# One row is expected per branch, per day, per service.
DAYS_IN_WEEK = 7

ID_COLUMNS = ["date", "branch", "service"]
COUNT_COLUMNS = ["booked", "delivered", "on_time", "returned"]
MONEY_COLUMNS = ["revenue", "fuel_cost"]
NUMERIC_COLUMNS = COUNT_COLUMNS + MONEY_COLUMNS
REQUIRED_COLUMNS = ID_COLUMNS + NUMERIC_COLUMNS

# Headings seen in branch files, mapped to the standard column name. Matching ignores
# case, spaces and punctuation, so "Parcels Booked" and "parcels_booked" are the same.
COLUMN_ALIASES = {
    "date": ["date", "report date", "dt", "day", "activity date"],
    "branch": [
        "branch", "branch code", "branch id", "branch name", "stn", "station", "location",
    ],
    "service": ["service", "service type", "svc", "product", "shipment class"],
    "booked": [
        "booked", "parcels booked", "booked qty", "bkd", "consignments booked",
    ],
    "delivered": [
        "delivered", "parcels delivered", "delivered qty", "dlvd", "consignments delivered",
    ],
    "on_time": [
        "on time", "on-time", "delivered on time", "ontime qty", "ot", "on-time deliveries",
    ],
    "returned": ["returned", "parcels returned", "returned qty", "rtn", "returns"],
    "revenue": ["revenue", "revenue (rs)", "sales", "rev", "gross revenue"],
    "fuel_cost": ["fuel cost", "fuel", "fuel (rs)", "fuel spend", "fuel expense"],
}

# A heading such as "Exp Booked" means one file holds both services side by side.
WIDE_SERVICE_PREFIXES = {
    "exp": "Express",
    "express": "Express",
    "std": "Standard",
    "standard": "Standard",
}

SERVICE_ALIASES = {
    "express": "Express",
    "exp": "Express",
    "ex": "Express",
    "standard": "Standard",
    "std": "Standard",
    "stnd": "Standard",
}

# A header row must match at least this many standard columns to count as the header.
MIN_HEADER_MATCHES = 5

# Fuzzy matching only corrects near-certain branch-name typos.
BRANCH_FUZZY_CUTOFF = 0.85

# Anything outside this range is treated as a likely keying error and is excluded.
REVENUE_PER_PARCEL_RANGE = (20.0, 400.0)

# Repeated auto-fixes of the same kind in one file collapse into a single exceptions line.
FIX_COLLAPSE_LIMIT = 5

# These start the Thresholds block in the report, where they can be changed by hand.
THRESHOLDS = {
    "on_time_target": 0.92,
    "on_time_tolerance": 0.03,
    "return_rate_limit": 0.05,
    "min_coverage": 0.85,
}
