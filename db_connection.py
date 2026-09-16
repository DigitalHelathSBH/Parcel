import os

import pandas as pd
import pyodbc
from dotenv import load_dotenv

load_dotenv()


def _escape(value: str) -> str:
    return "{" + value.replace("}", "}}") + "}"


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
    conn.setdecoding(pyodbc.SQL_CHAR, encoding="cp874")
    conn.setdecoding(pyodbc.SQL_WCHAR, encoding="utf-16le")
    conn.setencoding(encoding="utf-16le")
    return conn


def run_query(sql: str, params=None) -> pd.DataFrame:
    with get_connection() as conn:
        return pd.read_sql(sql, conn, params=params)
