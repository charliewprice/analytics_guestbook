import psycopg2 as pg
import pandas.io.sql as psql
import pandas as pd

conn_str = "host={0} port={1} dbname={2} user={3} password={4}".format("localhost", 5432, "guestbookdb", "gbuser", "w0lfpack")

def guestbookDbConnect():
  try:
    conn = pg.connect(conn_str)
    print("Welcome to Jupyter Notebook.  You are connected to the Opportunity House guestbook database!")
    return conn
  except pg.OperationalError:
    print("You are not connected to the database.")
    return None
