"""Diagnostic: download the whole SharePoint list (all columns), independent of the app's local copy."""
import argparse
import sys

import pandas as pd

from sharepoint_sync import COLUMNS, LIST_URL, LOCAL_TZ, open_session, sign_out, sp_get

# Columns Oświadczenia.py relies on - the test checks they are present.
REQUIRED_COLUMNS = list(COLUMNS)
# Built-in columns worth keeping; other built-in ones (Author, Attachments, ...) are skipped.
BASE_FIELDS_TO_KEEP = {"ID", "Title", "Created", "Modified"}
LOOKUP_TYPES = ("User", "UserMulti", "Lookup", "LookupMulti")


# ------------------------- SharePoint REST -------------------------
def lookup_target(field):
    """Property of an expanded person/lookup value that holds the display text."""
    if field["TypeAsString"].startswith("User"):
        return "Title"
    return field.get("LookupField") or "Title"


def get_fields(session):
    data = sp_get(session, f"{LIST_URL}/fields", {"$filter": "Hidden eq false"})
    return [f for f in data["value"]
            if f["TypeAsString"] != "Computed"
            and (not f["FromBaseType"] or f["InternalName"] in BASE_FIELDS_TO_KEEP)]


def get_items(session, fields, row_limit):
    # Person and lookup columns only return an ID unless they are expanded.
    lookups = [f for f in fields if f["TypeAsString"] in LOOKUP_TYPES]
    params = {
        "$select": ",".join(["*"] + [f"{f['InternalName']}/{lookup_target(f)}" for f in lookups]),
        "$top": min(row_limit, 5000) if row_limit else 5000,
    }
    if lookups:
        params["$expand"] = ",".join(f["InternalName"] for f in lookups)

    url = f"{LIST_URL}/items"
    items = []
    while url:
        data = sp_get(session, url, params)
        items.extend(data["value"])
        url = data.get("odata.nextLink")
        params = None  # the next link already carries the query
        if row_limit and len(items) >= row_limit:
            break
    return items[:row_limit] if row_limit else items


def simplify(value, field):
    """Flatten expanded person/lookup and multi-choice values to plain text."""
    if field["TypeAsString"] in LOOKUP_TYPES:
        key = lookup_target(field)
        if isinstance(value, list):
            return "; ".join(str(v[key]) for v in value if v.get(key)) or None
        if isinstance(value, dict):
            return value.get(key)
        return None
    if isinstance(value, list):
        return "; ".join(map(str, value)) or None
    return value


def to_dataframe(items, fields):
    df = pd.DataFrame({
        # Internal names starting with "_" come back prefixed with "OData_".
        f["Title"]: [simplify(item.get(f["InternalName"], item.get("OData_" + f["InternalName"])), f)
                     for item in items]
        for f in fields
    })
    for f in fields:
        if f["TypeAsString"] == "DateTime":
            df[f["Title"]] = (pd.to_datetime(df[f["Title"]], utc=True)
                              .dt.tz_convert(LOCAL_TZ).dt.tz_localize(None))
    return df


# ------------------------- Main -------------------------
def main():
    parser = argparse.ArgumentParser(description="Read the SharePoint list via REST API.")
    parser.add_argument("--rows", type=int, default=0, help="row limit, 0 = all rows")
    parser.add_argument("--csv", default=None, help="optional path to save the rows as CSV")
    parser.add_argument("--logout", action="store_true", help="forget the cached sign-in")
    args = parser.parse_args()

    if args.logout:
        sign_out()
        print("Cached sign-in removed.")
        return 0

    try:
        session = open_session()
    except Exception as e:
        print(f"FAILED to sign in: {e}")
        return 1
    print("Signed in.")

    try:
        fields = get_fields(session)
        items = get_items(session, fields, args.rows)
    except Exception as e:
        print(f"FAILED to read the list: {e}")
        return 1

    df = to_dataframe(items, fields)
    print(f"\nRetrieved {len(df)} rows, {len(df.columns)} columns.")
    print("Columns:", list(df.columns))

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        print("WARNING - columns required by Oświadczenia.py are missing:", missing)
    else:
        print("OK - all columns required by Oświadczenia.py are present.")

    if not df.empty:
        with pd.option_context("display.max_columns", None, "display.width", 200):
            print("\nFirst rows:")
            print(df[[c for c in REQUIRED_COLUMNS if c in df.columns]].head(10))

    if args.csv:
        df.to_csv(args.csv, index=False, encoding="utf-8-sig")
        print(f"\nSaved to {args.csv}")

    return 0 if not missing else 2


if __name__ == "__main__":
    sys.exit(main())
