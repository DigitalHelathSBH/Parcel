import codecs
import os

import pandas as pd
import pyodbc
from dotenv import load_dotenv

load_dotenv()


def _escape(value: str) -> str:
    return "{" + value.replace("}", "}}") + "}"


_REPLACE_SUFFIX = "_replace_errors"


def _replace_errors_lookup(name):
    if not name.endswith(_REPLACE_SUFFIX):
        return None
    base_name = name[: -len(_REPLACE_SUFFIX)]
    try:
        base = codecs.lookup(base_name)
    except LookupError:
        return None

    def decode(data, errors="replace"):
        return codecs.decode(bytes(data), base_name, "replace"), len(data)

    return codecs.CodecInfo(encode=base.encode, decode=decode, name=name)


# some legacy rows contain bytes that aren't valid in the configured encoding
# (data-entry glitches); decode those with U+FFFD replacement instead of
# raising and taking down the request
codecs.register(_replace_errors_lookup)


def get_connection() -> pyodbc.Connection:
    conn_str = (
        f"DRIVER={{{os.environ['DB_DRIVER']}}};"
        f"SERVER={os.environ['DB_SERVER']},{os.environ.get('DB_PORT', '1433')};"
        f"DATABASE={os.environ['DB_NAME']};"
        f"UID={_escape(os.environ['DB_USER'])};"
        f"PWD={_escape(os.environ['DB_PASSWORD'])};"
    )
    conn = pyodbc.connect(conn_str)
    # Legacy Thai columns are stored server-side as Windows-874 (cp874). Whether the
    # driver hands that to us raw or already transcodes it depends on platform/driver
    # (e.g. msodbcsql on Linux transcodes SQL_CHAR to UTF-8 before it reaches pyodbc,
    # while on Windows it typically doesn't) - override via DB_CHAR_ENCODING if needed.
    char_encoding = os.environ.get("DB_CHAR_ENCODING", "cp874")
    conn.setdecoding(pyodbc.SQL_CHAR, encoding=f"{char_encoding}{_REPLACE_SUFFIX}")
    conn.setdecoding(pyodbc.SQL_WCHAR, encoding="utf-16le")
    conn.setencoding(encoding="utf-16le")
    return conn


def run_query(sql: str, params=None) -> pd.DataFrame:
    with get_connection() as conn:
        return pd.read_sql(sql, conn, params=params)
