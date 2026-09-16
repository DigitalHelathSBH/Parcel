import codecs
import os

import pandas as pd
import pyodbc
from dotenv import load_dotenv

load_dotenv()


def _escape(value: str) -> str:
    return "{" + value.replace("}", "}}") + "}"


def _cp874_replace_decode(data, errors="replace"):
    return codecs.decode(bytes(data), "cp874", "replace"), len(data)


def _cp874_replace_lookup(name):
    if name != "cp874_replace":
        return None
    base = codecs.lookup("cp874")
    return codecs.CodecInfo(
        encode=base.encode,
        decode=_cp874_replace_decode,
        name="cp874_replace",
    )


# some legacy rows contain bytes that aren't valid cp874 (data-entry glitches);
# decode those with U+FFFD replacement instead of raising and taking down the request
codecs.register(_cp874_replace_lookup)


def get_connection() -> pyodbc.Connection:
    conn_str = (
        f"DRIVER={{{os.environ['DB_DRIVER']}}};"
        f"SERVER={os.environ['DB_SERVER']},{os.environ.get('DB_PORT', '1433')};"
        f"DATABASE={os.environ['DB_NAME']};"
        f"UID={_escape(os.environ['DB_USER'])};"
        f"PWD={_escape(os.environ['DB_PASSWORD'])};"
    )
    conn = pyodbc.connect(conn_str)
    # legacy Thai columns are stored as Windows-874 (cp874), not the ODBC default
    conn.setdecoding(pyodbc.SQL_CHAR, encoding="cp874_replace")
    conn.setdecoding(pyodbc.SQL_WCHAR, encoding="utf-16le")
    conn.setencoding(encoding="utf-16le")
    return conn


def run_query(sql: str, params=None) -> pd.DataFrame:
    with get_connection() as conn:
        return pd.read_sql(sql, conn, params=params)
