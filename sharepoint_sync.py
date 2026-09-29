"""Local SQLite copy of the SharePoint task list used by Oświadczenia.py.

The first sync downloads the whole list. Later syncs ask SharePoint when the list was
last changed and download only the items modified since then (and, when something was
deleted, the list of IDs). The whole list is downloaded again every FULL_SYNC_EVERY,
or whenever the local copy does not add up.
"""
import json
import os
import sqlite3
from collections import namedtuple
from datetime import datetime, timedelta

import msal
import pandas as pd
import requests
from msal_extensions import PersistedTokenCache, build_encrypted_persistence

# ------------------------- Configuration -------------------------
# Values taken from the data connection stored in "Dane sharepoint.xlsx".
SERVER_URL = "https://uhypl.sharepoint.com"
SITE_URL = SERVER_URL + "/sites/DAA"
LIST_ID = "5b1a818d-9b0b-4a9e-8cbe-396f485f0e66"
LIST_URL = f"{SITE_URL}/_api/web/lists(guid'{LIST_ID}')"

TENANT = "uhy-pl.com"
# Application (client) ID of the app registration - can be overridden with SP_CLIENT_ID.
CLIENT_ID = os.environ.get("SP_CLIENT_ID") or "d73af729-3bb4-445e-b763-e6a49097dd0c"
SCOPES = [SERVER_URL + "/.default"]

# Per-user files, kept outside OneDrive so the database is never synced between users.
DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "Oswiadczenia")
TOKEN_CACHE_PATH = os.path.join(DATA_DIR, "sharepoint_token_cache.bin")
DB_PATH = os.path.join(DATA_DIR, "sharepoint_cache.db")
LOCAL_TZ = "Europe/Warsaw"  # SharePoint returns dates in UTC

# Columns Oświadczenia.py uses: column title -> SharePoint internal name.
COLUMNS = {
    "Nazwa firmy": "Nazwa_x0020_firmy",
    "Data rozpoczęcia": "Data_x0020_wp_x0142_yni_x0119_ci",
    "Osoba odpowiedzialna": "Osoba_x0020_odpowiedzialna",
    "Typ zadania": "Typanalizy",
    "Rodzaj sprawozdania": "Rodzaj_x0020_sprawozdania",
}
PERSON_COLUMNS = {"Osoba odpowiedzialna"}  # return only an ID unless expanded
DATE_COLUMNS = {"Data rozpoczęcia"}
# Edits that keep an item's old "Modified" date are only picked up by a full download.
FULL_SYNC_EVERY = timedelta(days=7)

SyncResult = namedtuple("SyncResult", "full changed deleted")


# ------------------------- Sign-in -------------------------
def _msal_app():
    return msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT}",
        token_cache=PersistedTokenCache(build_encrypted_persistence(TOKEN_CACHE_PATH)),
        # Sign in through the Windows account broker (WAM) instead of a browser tab.
        enable_broker_on_windows=True,
    )


def get_token(parent_window_handle=msal.PublicClientApplication.CONSOLE_WINDOW_HANDLE):
    """Access token for SharePoint; asks the user to sign in only if the cached sign-in can't be used.

    parent_window_handle is the window the sign-in dialog opens over."""
    app = _msal_app()
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
    if not result:
        result = app.acquire_token_interactive(SCOPES, parent_window_handle=parent_window_handle)
    if "access_token" not in result:
        raise RuntimeError(f"{result.get('error')}: {result.get('error_description')}")
    return result["access_token"]


def sign_out():
    app = _msal_app()
    for account in app.get_accounts():
        app.remove_account(account)


# ------------------------- SharePoint REST -------------------------
class ThrottledError(RuntimeError):
    """SharePoint refused a query that matches more than 5000 items (list view threshold)."""


def open_session(parent_window_handle=msal.PublicClientApplication.CONSOLE_WINDOW_HANDLE):
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {get_token(parent_window_handle)}",
                            "Accept": "application/json;odata=nometadata"})
    return session


def sp_get(session, url, params=None):
    response = session.get(url, params=params, timeout=60)
    if not response.ok:
        error = ThrottledError if "SPQueryThrottledException" in response.text else RuntimeError
        raise error(f"HTTP {response.status_code} for {response.url}\n{response.text[:500]}")
    return response.json()


def sp_get_all(session, url, params):
    """All rows of a list query, following SharePoint's paging."""
    rows = []
    while url:
        data = sp_get(session, url, params)
        rows.extend(data["value"])
        url = data.get("odata.nextLink")
        params = None  # the next link already carries the query
    return rows


def _fetch_rows(session, filter_=None):
    """Items as rows of ID followed by the COLUMNS values (dates stay UTC ISO strings)."""
    select = ["ID"] + [f"{name}/Title" if title in PERSON_COLUMNS else name
                       for title, name in COLUMNS.items()]
    params = {"$select": ",".join(select), "$top": 5000,
              "$expand": ",".join(COLUMNS[title] for title in PERSON_COLUMNS)}
    if filter_:
        params["$filter"] = filter_

    rows = []
    for item in sp_get_all(session, f"{LIST_URL}/items", params):
        row = [item["ID"]]
        for title, name in COLUMNS.items():
            value = item.get(name)
            if title in PERSON_COLUMNS:
                value = value.get("Title") if isinstance(value, dict) else None
            row.append(value)
        rows.append(row)
    return rows


# ------------------------- Local database -------------------------
_ITEM_COLUMNS = ", ".join(f'"{title}"' for title in COLUMNS)
_INSERT_ITEM = (f"INSERT OR REPLACE INTO items (ID, {_ITEM_COLUMNS}) "
                f"VALUES ({', '.join('?' * (len(COLUMNS) + 1))})")
_SCHEMA = json.dumps(COLUMNS, ensure_ascii=False)  # a change forces a full download


def _connect():
    os.makedirs(DATA_DIR, exist_ok=True)
    # Transactions are started explicitly, so that a full reload (drop, create and
    # insert) is saved all at once or not at all.
    con = sqlite3.connect(DB_PATH, isolation_level=None)
    con.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
    return con


def _read_meta(con):
    return dict(con.execute("SELECT key, value FROM meta"))


def load():
    """The local copy as a DataFrame (dates in local time) and the time of the last sync.

    Before the first successful sync the DataFrame is empty and the time is None."""
    con = _connect()
    try:
        meta = _read_meta(con)
        if meta.get("schema") == _SCHEMA:
            df = pd.read_sql_query(f"SELECT {_ITEM_COLUMNS} FROM items", con)
        else:
            df = pd.DataFrame(columns=list(COLUMNS))
    finally:
        con.close()
    for title in DATE_COLUMNS:
        df[title] = pd.to_datetime(df[title], utc=True).dt.tz_convert(LOCAL_TZ).dt.tz_localize(None)
    return df, meta.get("last_sync")


def _apply_changes(session, con, meta, info):
    """Save items edited or deleted since the last sync; returns (changed, deleted).

    Returns None when a full download is needed instead."""
    changed = deleted = 0
    if info["LastItemModifiedDate"] != meta["last_modified"]:
        try:
            # "ge" because the dates are whole seconds; saving an item again is harmless.
            rows = _fetch_rows(session, f"Modified ge datetime'{meta['last_modified']}'")
        except ThrottledError:
            return None  # over 5000 changed items
        con.executemany(_INSERT_ITEM, rows)
        changed = len(rows)

    if info["LastItemDeletedDate"] != meta["last_deleted"]:
        remote_ids = {item["ID"] for item in
                      sp_get_all(session, f"{LIST_URL}/items", {"$select": "ID", "$top": 5000})}
        gone = [(item_id,) for (item_id,) in con.execute("SELECT ID FROM items")
                if item_id not in remote_ids]
        con.executemany("DELETE FROM items WHERE ID = ?", gone)
        deleted = len(gone)

    # The counts differ when a change was missed, e.g. an item restored from the recycle bin.
    if con.execute("SELECT COUNT(*) FROM items").fetchone()[0] != info["ItemCount"]:
        return None
    return changed, deleted


def _reload_all(session, con):
    rows = _fetch_rows(session)
    con.execute("DROP TABLE IF EXISTS items")
    con.execute(f"CREATE TABLE items (ID INTEGER PRIMARY KEY, {_ITEM_COLUMNS})")
    con.executemany(_INSERT_ITEM, rows)
    return len(rows)


def sync(full=False, parent_window_handle=msal.PublicClientApplication.CONSOLE_WINDOW_HANDLE):
    """Bring the local copy up to date with SharePoint; full=True downloads the whole list."""
    session = open_session(parent_window_handle)
    # Read the list's dates before its items: anything changed during the sync is
    # newer than the dates saved below, so the next sync picks it up.
    info = sp_get(session, LIST_URL,
                  {"$select": "ItemCount,LastItemModifiedDate,LastItemDeletedDate"})
    now = datetime.now().isoformat(timespec="seconds")

    con = _connect()
    try:
        meta = _read_meta(con)
        last_full = meta.get("last_full_sync")
        full = (full or meta.get("schema") != _SCHEMA or not last_full
                or datetime.now() - datetime.fromisoformat(last_full) > FULL_SYNC_EVERY)

        con.execute("BEGIN")
        changes = None if full else _apply_changes(session, con, meta, info)
        full = changes is None
        if full:
            changes = (_reload_all(session, con), 0)
            meta["last_full_sync"] = now
        meta.update(schema=_SCHEMA, last_sync=now,
                    last_modified=info["LastItemModifiedDate"],
                    last_deleted=info["LastItemDeletedDate"])
        con.executemany("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)", meta.items())
        con.execute("COMMIT")
    finally:
        con.close()  # a transaction that was not committed is rolled back
    return SyncResult(full, *changes)
