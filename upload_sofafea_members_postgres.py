"""upload_sofafea_members_postgres.py

Uploads Ancestors_Database_v10_copy.xlsx into the PostgreSQL SOFAFEA database.

Tables populated:
  - locations          (city, county, state, coordinates, country)
  - family_members     (person data with family tree links)
  - sofafea_members    (extra Excel columns not in family_members)

Requirements:  pip install openpyxl pandas psycopg2-binary python-dotenv
Env:           DB_CONNECTION_STRING in .env or backend/.env
"""

from __future__ import annotations

import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import psycopg2
from dotenv import load_dotenv

# Resolve .env
SCRIPT_DIR = Path(__file__).resolve().parent
ENV_PATH = SCRIPT_DIR / "backend" / ".env"
if not ENV_PATH.exists():
    ENV_PATH = SCRIPT_DIR / ".env"
load_dotenv(ENV_PATH)

DB_DSN = os.getenv("DB_CONNECTION_STRING", "")
if not DB_DSN:
    print("ERROR: DB_CONNECTION_STRING not set in .env")
    sys.exit(1)

INPUT_FILE = SCRIPT_DIR / "camp_william_penn_civil_war_troops_v4.xlsx"
SHEET_NAME = "Soldier_records"

# Logging
def log(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"    [{ts}] {msg}", flush=True)


# Value helpers
def is_empty(val: Any) -> bool:
    if val is None: return True
    if isinstance(val, float) and pd.isna(val): return True
    s = str(val).strip()
    return s.lower() in ("", "nan", "nat", "none", "null", "-", "--", "----", "n/a", "na")

def cell(row: pd.Series, *keys: str) -> str:
    for k in keys:
        v = row.get(k)
        if v is not None and not is_empty(v):
            return str(v).strip()
    return ""

def date_cell(row: pd.Series, *keys: str) -> Optional[str]:
    s = cell(row, *keys)
    if not s: return None
    m = re.match(r"^(\d{4}-\d{2}-\d{2})", s)
    if m: return m.group(1)
    m = re.match(r"^(\d{4})$", s)
    if m: return f"{m.group(1)}-01-01"
    return None


# Location
def insert_location(cur, city: str = "", county: str = "", state: str = "", coords: str = "") -> Optional[int]:
    if not any([city, county, state, coords]): return None
    cur.execute(
        "SELECT locations_id FROM public.locations WHERE city=%s AND county=%s AND state=%s ORDER BY locations_id LIMIT 1",
        (city, county, state))
    row = cur.fetchone()
    if row:
        if coords:
            cur.execute("UPDATE public.locations SET coordinates=%s WHERE locations_id=%s AND (coordinates IS NULL OR coordinates='')", (coords, row[0]))
        return row[0]
    cur.execute("INSERT INTO public.locations (city,county,state,coordinates) VALUES (%s,%s,%s,%s) RETURNING locations_id", (city, county, state, coords))
    return cur.fetchone()[0]


# Member
def insert_person(cur, row: pd.Series, first: str, last: str, alias: str = "",
                  race: str = "", ethnicity: str = "",
                  birth_date = None, birth_loc_id = None,
                  death_date = None, death_loc_id = None,
                  marriage_date = None, marriage_loc_id = None,
                  mil_svc: str = "", branch: str = "", war: str = "",
                  gen_number: int = 1, father_id = None, mother_id = None, spouse_id = None) -> Optional[int]:
    if is_empty(first) and is_empty(last): return None
    now = datetime.now()
    cur.execute(
        "INSERT INTO public.family_members (first_name,last_name,alias,race,ethnicity, birth_date,birth_location_id, death_date,death_location_id, marriage_date,marriage_location_id, military_service,branch,war, generation_number,father_id,mother_id,spouse_id,created_at) VALUES (%s,%s,%s,%s,%s, %s,%s, %s,%s, %s,%s, %s,%s,%s, %s,%s,%s,%s,%s) RETURNING member_id",
        (first, last, alias, race, ethnicity, birth_date, birth_loc_id, death_date, death_loc_id, marriage_date, marriage_loc_id, mil_svc, branch, war, gen_number, father_id, mother_id, spouse_id, now))
    return cur.fetchone()[0]


# sofafea_members
def insert_sofafea(cur, member_id: int, row: pd.Series) -> None:
    cur.execute(
        "INSERT INTO public.sofafea_members (member_id,contact_member_number,email,contact_membership_type) VALUES (%s,%s,%s,%s) ON CONFLICT (member_id) DO UPDATE SET contact_member_number=EXCLUDED.contact_member_number, email=EXCLUDED.email, contact_membership_type=EXCLUDED.contact_membership_type",
        (member_id, cell(row,"contact_member_#"), cell(row,"email"), cell(row,"contact_membership_type")))


# Main
def main() -> None:
    overall_start = time.time()
    log(f"Reading {INPUT_FILE} ...")
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    log(f"Loaded {len(df)} rows x {len(df.columns)} columns.")

    # Normalize column names
    new_cols = []
    for c in df.columns:
        new_cols.append(re.sub(r"\s+", "_", str(c).strip()).lower())
    df.columns = new_cols
    log(f"Column names normalized.")

    conn = psycopg2.connect(DB_DSN)
    conn.autocommit = False

    try:
        with conn.cursor() as cur:
            cur.execute("""CREATE TABLE IF NOT EXISTS public.sofafea_members (
                sofafea_member_id SERIAL PRIMARY KEY,
                member_id INTEGER UNIQUE REFERENCES public.family_members(member_id) ON DELETE CASCADE,
                contact_member_number TEXT, email TEXT, contact_membership_type TEXT,
                documentation TEXT, generation_number INTEGER,
                created_at TIMESTAMPTZ DEFAULT now()
            )""")
            conn.commit()
            log("sofafea_members table ready.")

            total_rows = len(df)
            total_inserted = 0
            skipped = 0

            for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
                fname = cell(row, "first_name")
                lname = cell(row, "last_name")
                if is_empty(fname) and is_empty(lname):
                    skipped += 1
                    continue

                log(f"Row {row_idx}/{total_rows}: {fname} {lname}")
                gen_ids = {}

                # Gen 1
                m_id = insert_person(cur, row,
                    first=fname, last=lname,
                    race=cell(row,"race"), ethnicity=cell(row,"ethnicity"),
                    birth_date=date_cell(row,"gen_1:_birth_date"),
                    birth_loc_id=insert_location(cur, cell(row,"city,_county,_state"), cell(row,"county"), cell(row,"state.1"), cell(row,"coordinates")),
                    death_date=date_cell(row,"gen_1:_death_date"),
                    death_loc_id=insert_location(cur, cell(row,"city,_county,_state.1"), cell(row,"county.1"), cell(row,"state.2"), cell(row,"coordinates.1")),
                    marriage_date=date_cell(row,"gen_1:_marriage_date"),
                    marriage_loc_id=insert_location(cur, cell(row,"city,_county,_state.2"), cell(row,"county.2"), cell(row,"state.3"), cell(row,"coordinates.2")),
                    mil_svc=cell(row,"military_service"), branch=cell(row,"branch"), war=cell(row,"war"),
                    gen_number=1)
                if m_id:
                    gen_ids[1] = m_id
                    insert_sofafea(cur, m_id, row)
                    total_inserted += 1

                    # Spouse
                    sp_id = insert_person(cur, row,
                        first=cell(row,"gen_1:_spouse_first_name"),
                        last=cell(row,"gen_1:_spouse_surname/maiden_name"),
                        race=cell(row,"race.1"), ethnicity=cell(row,"ethnicity.1"),
                        birth_date=date_cell(row,"spouse_gen_1:_birth_date"),
                        birth_loc_id=insert_location(cur, cell(row,"city,_county,_state.3"), cell(row,"county.3"), cell(row,"state.4"), cell(row,"coordinates.3")),
                        death_date=date_cell(row,"spouse_gen_1:_death_date"),
                        death_loc_id=insert_location(cur, cell(row,"city,_county,_state.4"), cell(row,"county.4"), cell(row,"state.5"), cell(row,"coordinates.4")),
                        gen_number=1)
                    if sp_id:
                        cur.execute("UPDATE public.family_members SET spouse_id=%s WHERE member_id=%s", (sp_id, m_id))
                        cur.execute("UPDATE public.family_members SET spouse_id=%s WHERE member_id=%s", (m_id, sp_id))
                        total_inserted += 1

                # Gen 2-10
                loc_off = 5
                for gen in range(2, 11):
                    gs = f"gen_{gen}"
                    if gen == 2:
                        f_first = cell(row, "gen_2:_father's:_first_name", "gen_2:_father's_first_name")
                        f_last = cell(row, "gen_2:_father's_surname", "gen_2:_father's:_surname")
                        f_bd = date_cell(row, "gen_2:_father's_birth_date", "gen_2:_father's:_birth_date")
                        f_dd = date_cell(row, "gen_2:_father's_death_date", "gen_2:_father's:_death_date")
                        f_md = date_cell(row, "gen_2:_father,_mother's_marriage_date")
                    else:
                        f_first = cell(row, f"{gs}_father:_first_name")
                        f_last = cell(row, f"{gs}_father_surname", f"{gs}:_father_surname")
                        f_bd = date_cell(row, f"{gs}:_father's_birth_date", f"{gs}_father's_birth_date")
                        f_dd = date_cell(row, f"{gs}:_father's_death_date", f"{gs}_father's_death_date")
                        f_md = date_cell(row, f"{gs}:_father,_mother's_marriage_date")

                    f_race = cell(row, f"race.{gen-1}" if gen>2 else "race.3")
                    f_eth = cell(row, f"ethnicity.{gen-2}" if gen>2 else "ethnicity.2")
                    f_alias = cell(row, f"alias_name.{gen-2}" if gen>2 else "alias_name")
                    f_mil = cell(row, f"military_service.{gen-1}" if gen>2 else "military_service.1")
                    f_br = cell(row, f"branch.{gen-1}" if gen>2 else "branch.1")
                    f_war = cell(row, f"war.{gen-2}" if gen>3 else ("war.1" if gen==2 else "war.2"))

                    f_bloc = insert_location(cur, cell(row,f"city,_county,_state.{loc_off}"), cell(row,f"county.{loc_off-4}"), cell(row,f"state.{loc_off+1}"), cell(row,f"coordinates.{loc_off}"))
                    f_dloc = insert_location(cur, cell(row,f"city,_county,_state.{loc_off+1}"), cell(row,f"county.{loc_off-3}"), cell(row,f"state.{loc_off+2}"), cell(row,f"coordinates.{loc_off+1}"))
                    f_mloc = insert_location(cur, cell(row,f"city,_county,_state.{loc_off+2}"), cell(row,f"county.{loc_off-2}"), cell(row,f"state.{loc_off+3}"), cell(row,f"coordinates.{loc_off+2}"))

                    father_id = None
                    if not is_empty(f_first) or not is_empty(f_last):
                        father_id = insert_person(cur, row,
                            first=f_first, last=f_last, alias=f_alias,
                            race=f_race, ethnicity=f_eth,
                            birth_date=f_bd, birth_loc_id=f_bloc,
                            death_date=f_dd, death_loc_id=f_dloc,
                            marriage_date=f_md, marriage_loc_id=f_mloc,
                            mil_svc=f_mil, branch=f_br, war=f_war,
                            gen_number=gen)
                        if father_id:
                            total_inserted += 1
                            if gen-1 in gen_ids:
                                cur.execute("UPDATE public.family_members SET father_id=%s WHERE member_id=%s", (father_id, gen_ids[gen-1]))

                    # Mother
                    m_first = cell(row, f"{gs}:_mother's_first_name", f"{gs}_mother's_first_name")
                    m_last = cell(row, f"{gs}:_mother's_surname/maiden_name", f"{gs}_mother's_surname/maiden_name", f"{gs}:_mother's_surname")
                    m_bd = date_cell(row, f"_{gs}_mother's_birth_date", f"{gs}_mother's_birth_date")
                    m_dd = date_cell(row, f"{gs}_mother's_death_date")
                    m_race = cell(row, f"race.{gen}" if gen>2 else "race.4")
                    m_bloc = insert_location(cur, cell(row,f"city,_county,_state.{loc_off+3}"), cell(row,f"county.{loc_off-1}"), cell(row,f"state.{loc_off+4}"), cell(row,f"coordinates.{loc_off+3}"))
                    m_dloc = insert_location(cur, cell(row,f"city,_county,_state.{loc_off+4}"), cell(row,f"county.{loc_off}"), cell(row,f"state.{loc_off+5}"), cell(row,f"coordinates.{loc_off+4}"))

                    mother_id = None
                    if not is_empty(m_first) or not is_empty(m_last):
                        mother_id = insert_person(cur, row,
                            first=m_first, last=m_last, race=m_race,
                            birth_date=m_bd, birth_loc_id=m_bloc,
                            death_date=m_dd, death_loc_id=m_dloc,
                            gen_number=gen)
                        if mother_id:
                            total_inserted += 1
                            if gen-1 in gen_ids:
                                cur.execute("UPDATE public.family_members SET mother_id=%s WHERE member_id=%s", (mother_id, gen_ids[gen-1]))

                    if father_id: gen_ids[gen] = father_id
                    elif mother_id: gen_ids[gen] = mother_id
                    loc_off += 5

                conn.commit()

    except Exception as exc:
        conn.rollback()
        log(f"ERROR: {exc}")
        raise
    finally:
        conn.close()

    elapsed = time.time() - overall_start
    log(f"Done. {total_inserted} members inserted, {skipped} rows skipped.")
    log(f"Completed in {elapsed:.1f}s ({elapsed/60:.1f} min).")

if __name__ == "__main__":
    main()
