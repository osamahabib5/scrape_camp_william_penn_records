DROP PUBLICATION IF EXISTS archive_pub;

SELECT pubname FROM pg_publication;




CREATE SUBSCRIPTION archive_sub
CONNECTION 'host=sofafea-postgres.postgres.database.azure.com port=5432 dbname=postgres user=replicator_user password=P@ssword_B@ckup
 sslmode=require'
PUBLICATION archive_pub;


SELECT COUNT(*) FROM locations;

SELECT subname, subenabled, subslotname, subpublications, subconninfo 
FROM pg_subscription;

SELECT srrelid::regclass AS table_name, srsubstate 
FROM pg_subscription_rel;

SELECT 
    schemaname,
    relname AS tablename,
    n_live_tup AS estimated_rows
FROM pg_stat_user_tables
WHERE relname IN ('locations', 'usct_connecticut', 'family_members', 'book_of_negroes')
ORDER BY relname;

SELECT 'locations' AS table_name, COUNT(*) AS row_count FROM locations
UNION ALL
SELECT 'usct_connecticut', COUNT(*) FROM usct_connecticut
UNION ALL
SELECT 'family_members', COUNT(*) FROM family_members
UNION ALL
SELECT 'book_of_negroes', COUNT(*) FROM book_of_negroes;

SELECT 
    slot_name,
    active,
    pg_size_pretty(pg_wal_lsn_diff(pg_current_wal_lsn(), confirmed_flush_lsn)) AS lag_size
FROM pg_replication_slots
WHERE slot_name = 'archive_sub';

SELECT subname, subenabled, substate, subslotname 
FROM pg_subscription;

SELECT * FROM pg_stat_subscription;


ALTER SUBSCRIPTION archive_sub REFRESH PUBLICATION;

SELECT srrelid::regclass, srsubstate FROM pg_subscription_rel;

SELECT * FROM pg_stat_subscription WHERE worker_type = 'apply' AND last_error IS NOT NULL;

SELECT srrelid::regclass AS table_name, srsubstate FROM pg_subscription_rel;

SELECT subname, subenabled, subslotname, subconninfo 
FROM pg_subscription;

SELECT pid, application_name, state, backend_start 
FROM pg_stat_activity 
WHERE application_name LIKE '%archive_sub%' OR application_name LIKE 'logical replication worker%';

SELECT subname, subenabled, subslotname, subconninfo 
FROM pg_subscription;

SELECT subname, subenabled, subslotname, subconninfo 
FROM pg_subscription;

ALTER SUBSCRIPTION archive_sub DISABLE;
ALTER SUBSCRIPTION archive_sub ENABLE;

SELECT pid, application_name, state 
FROM pg_stat_activity 
-- WHERE application_name LIKE '%archive_sub%' OR application_name LIKE 'logical replication worker%';


SELECT subname, last_error, last_error_time, pid 
FROM pg_stat_subscription;

SELECT column_name, data_type 
FROM information_schema.columns 
WHERE table_name = 'pg_stat_subscription'
ORDER BY ordinal_position;

SELECT subname, pid, received_lsn, last_msg_receipt_time 
FROM pg_stat_subscription;

SELECT pid, worker_type, leader_pid 
FROM pg_stat_subscription;

DROP SUBSCRIPTION archive_sub;


SELECT tablename FROM pg_tables WHERE schemaname = 'public' AND tablename IN ('locations', 'usct_connecticut', 'family_members', 'book_of_negroes');

CREATE TABLE test_repl (
    id INTEGER PRIMARY KEY,
    name TEXT
);

DROP TABLE IF EXISTS test_repl;

CREATE SUBSCRIPTION archive_sub
CONNECTION 'host=sofafea-postgres.postgres.database.azure.com port=5432 dbname=postgres user=replicator_user password=SofAFeA35.1 sslmode=require'
PUBLICATION archive_pub;

SELECT pid, worker_type FROM pg_stat_subscription;

SELECT srrelid::regclass AS table_name, srsubstate FROM pg_subscription_rel;

SELECT subname, subenabled, subslotname, subpublications 
FROM pg_subscription;

SELECT srrelid::regclass AS table_name, srsubstate FROM pg_subscription_rel;

SELECT pid, worker_type, last_msg_receipt_time 
FROM pg_stat_subscription;

DROP SUBSCRIPTION archive_sub;

SELECT tablename FROM pg_tables WHERE schemaname = 'public' AND tablename IN ('locations', 'usct_connecticut', 'family_members', 'book_of_negroes');

SELECT srrelid::regclass, srsubstate FROM pg_subscription_rel;

SELECT received_lsn, latest_end_lsn FROM pg_stat_subscription;

SELECT pid, locktype, mode, granted 
FROM pg_locks 
WHERE NOT granted;


SELECT conname
FROM pg_constraint
WHERE contype = 'f'
  AND conrelid::regclass IN ('locations'::regclass, 'usct_connecticut'::regclass, 'family_members'::regclass, 'book_of_negroes'::regclass);


ALTER TABLE book_of_negroes DROP CONSTRAINT book_of_negroes_arrival_location_id_fkey;
ALTER TABLE book_of_negroes DROP CONSTRAINT book_of_negroes_departure_location_id_fkey;
ALTER TABLE book_of_negroes DROP CONSTRAINT book_of_negroes_member_id_fkey;
ALTER TABLE family_members DROP CONSTRAINT family_members_birth_location_id_fkey;
ALTER TABLE family_members DROP CONSTRAINT family_members_father_id_fkey;
ALTER TABLE family_members DROP CONSTRAINT family_members_mother_id_fkey;
ALTER TABLE family_members DROP CONSTRAINT family_members_spouse_id_fkey;
ALTER TABLE usct_connecticut DROP CONSTRAINT usct_connecticut_enlistment_location_id_fkey;
ALTER TABLE usct_connecticut DROP CONSTRAINT usct_connecticut_member_id_fkey;
ALTER TABLE usct_connecticut DROP CONSTRAINT usct_connecticut_pob_location_id_fkey;
ALTER TABLE usct_connecticut DROP CONSTRAINT usct_connecticut_residence_location_id_fkey;

SELECT COUNT(*) FROM locations;

SELECT COUNT(*) FROM locations;
SELECT COUNT(*) FROM usct_connecticut;
SELECT COUNT(*) FROM family_members;
SELECT COUNT(*) FROM book_of_negroes;


ALTER TABLE book_of_negroes ADD CONSTRAINT book_of_negroes_arrival_location_id_fkey FOREIGN KEY (arrival_location_id) REFERENCES locations(locations_id) ON UPDATE CASCADE ON DELETE SET NULL;
ALTER TABLE book_of_negroes ADD CONSTRAINT book_of_negroes_departure_location_id_fkey FOREIGN KEY (departure_location_id) REFERENCES locations(locations_id) ON UPDATE CASCADE ON DELETE SET NULL;
ALTER TABLE book_of_negroes ADD CONSTRAINT book_of_negroes_member_id_fkey FOREIGN KEY (member_id) REFERENCES family_members(member_id);
ALTER TABLE family_members ADD CONSTRAINT family_members_birth_location_id_fkey FOREIGN KEY (birth_location_id) REFERENCES locations(locations_id) ON UPDATE CASCADE ON DELETE SET NULL;
ALTER TABLE family_members ADD CONSTRAINT family_members_father_id_fkey FOREIGN KEY (father_id) REFERENCES family_members(member_id);
ALTER TABLE family_members ADD CONSTRAINT family_members_mother_id_fkey FOREIGN KEY (mother_id) REFERENCES family_members(member_id);
ALTER TABLE family_members ADD CONSTRAINT family_members_spouse_id_fkey FOREIGN KEY (spouse_id) REFERENCES family_members(member_id);
ALTER TABLE usct_connecticut ADD CONSTRAINT usct_connecticut_enlistment_location_id_fkey FOREIGN KEY (enlistment_location_id) REFERENCES locations(locations_id) ON UPDATE CASCADE ON DELETE SET NULL;
ALTER TABLE usct_connecticut ADD CONSTRAINT usct_connecticut_member_id_fkey FOREIGN KEY (member_id) REFERENCES family_members(member_id);
ALTER TABLE usct_connecticut ADD CONSTRAINT usct_connecticut_pob_location_id_fkey FOREIGN KEY (pob_location_id) REFERENCES locations(locations_id) ON UPDATE CASCADE ON DELETE SET NULL;
ALTER TABLE usct_connecticut ADD CONSTRAINT usct_connecticut_residence_location_id_fkey FOREIGN KEY (residence_location_id) REFERENCES locations(locations_id) ON UPDATE CASCADE ON DELETE SET NULL;

CREATE SUBSCRIPTION archive_sub
CONNECTION 'host=sofafea-postgres.postgres.database.azure.com port=5432 dbname=postgres user=replicator_user password=SofAFeA35.1 sslmode=require'
PUBLICATION archive_pub
WITH (copy_data = false);

SELECT subname, subenabled FROM pg_subscription;

SELECT * FROM locations WHERE locations_id = 9999;


-- ============================================================================
--  REPLICATE AUDIT TRAIL & USERS TABLES TO BACKUP
-- ============================================================================
-- The backup database must have identical table structures before
-- replication can populate them.  Run this section on the backup.

-- Step 1: Create the users_information table
CREATE TABLE IF NOT EXISTS public.users_information (
    id              BIGSERIAL PRIMARY KEY,
    username        TEXT        NOT NULL UNIQUE,
    password        TEXT        NOT NULL,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    role            VARCHAR(255) NOT NULL DEFAULT 'viewer'
);

-- Step 2: Create the audit schema and logged_actions table
CREATE SCHEMA IF NOT EXISTS audit;

CREATE TABLE IF NOT EXISTS audit.logged_actions (
    id              BIGSERIAL PRIMARY KEY,
    schema_name     TEXT        NOT NULL DEFAULT 'public',
    table_name      TEXT        NOT NULL,
    action          TEXT        NOT NULL CHECK (action IN ('INSERT', 'UPDATE', 'DELETE')),
    row_id          TEXT        NOT NULL,
    changed_by      TEXT,
    changed_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    old_values      JSONB,
    new_values      JSONB,
    changed_columns TEXT[]
);

CREATE INDEX IF NOT EXISTS idx_audit_table_time
    ON audit.logged_actions (table_name, changed_at DESC);

CREATE INDEX IF NOT EXISTS idx_audit_changed_by
    ON audit.logged_actions (changed_by);


-- ============================================================================
--  REPLICATE sofafea_members TABLE TO BACKUP
-- ============================================================================
-- The backup database must have the sofafea_members table before replication
-- can populate it.  Run this section on the backup database.

-- Step 1: Create the sofafea_members table (must match source exactly)
CREATE TABLE IF NOT EXISTS public.sofafea_members (
    sofafea_member_id       SERIAL PRIMARY KEY,
    member_id               INTEGER UNIQUE REFERENCES public.family_members(member_id)
                                ON DELETE CASCADE,
    contact_member_number   TEXT,
    email                   TEXT,
    contact_membership_type TEXT,
    documentation           TEXT,
    generation_number       INTEGER,
    created_at              TIMESTAMPTZ DEFAULT now()
);

-- Step 2: Create indexes for sofafea_members
CREATE INDEX IF NOT EXISTS idx_sofafea_member_id
    ON public.sofafea_members (member_id);

CREATE INDEX IF NOT EXISTS idx_sofafea_gen
    ON public.sofafea_members (generation_number);

-- Step 3: Generate any FK constraints on the source that should be recreated here
-- (Run on source database to get the DDL)
-- SELECT 'ALTER TABLE ' || conrelid::regclass ||
--   ' ADD CONSTRAINT ' || conname || ' ' ||
--   pg_get_constraintdef(oid) || ';' AS add_constraint_stmt
-- FROM pg_constraint
-- WHERE contype = 'f'
--   AND (conrelid::regclass = 'sofafea_members'::regclass
--        OR confrelid::regclass = 'sofafea_members'::regclass)
-- ORDER BY conname;

-- Step 3: Verify the tables now exist on the backup
SELECT tablename, schemaname
FROM pg_tables
WHERE (schemaname = 'public' AND tablename = 'users_information')
   OR (schemaname = 'audit' AND tablename = 'logged_actions');

-- Step 2: Check current row counts (should be zero before first refresh)
SELECT 'users_information' AS tbl, COUNT(*) AS cnt FROM public.users_information
UNION ALL
SELECT 'audit.logged_actions', COUNT(*) FROM audit.logged_actions;

-- Step 3: Drop any FK constraints on users_information that may
-- interfere with replication (none expected, but defensive)
DO $$
DECLARE
    r RECORD;
BEGIN
    FOR r IN
        SELECT conname, conrelid::regclass AS tbl
        FROM pg_constraint
        WHERE contype = 'f'
          AND (conrelid::regclass = 'users_information'::regclass
               OR confrelid::regclass = 'users_information'::regclass)
    LOOP
        EXECUTE 'ALTER TABLE ' || r.tbl || ' DROP CONSTRAINT ' || r.conname;
    END LOOP;
END;
$$;

-- Step 4: Drop and recreate the subscription with all tables included.
-- copy_data = false avoids re-syncing existing data.
-- This means only NEW changes after this point will be replicated.

DROP SUBSCRIPTION IF EXISTS archive_sub;

CREATE SUBSCRIPTION archive_sub
CONNECTION 'host=sofafea-postgres.postgres.database.azure.com port=5432 dbname=postgres user=replicator_user password=SofAFeA35.1 sslmode=require'
PUBLICATION archive_pub
WITH (copy_data = false);

-- Step 5: Check the subscription state for the new tables
SELECT srrelid::regclass AS table_name, srsubstate
FROM pg_subscription_rel
WHERE srrelid::regclass::text IN ('users_information', 'audit.logged_actions');

-- Step 5b: Copy existing data from source for the new tables.
-- Since copy_data = false, existing rows on the source won't be
-- replicated automatically.  Use dblink or pg_dump to backfill them.
-- Option A — using dblink (run on backup):
--   CREATE EXTENSION IF NOT EXISTS dblink;
--   INSERT INTO public.users_information
--   SELECT * FROM dblink('host=sofafea-postgres.postgres.database.azure.com
--                         port=5432 dbname=postgres user=replicator_user
--                         password=SofAFeA35.1 sslmode=require',
--                        'SELECT id, username, password, created_at, updated_at, role
--                         FROM public.users_information')
--   AS t (id bigint, username text, password text,
--         created_at timestamptz, updated_at timestamptz, role varchar(255));
--   INSERT INTO audit.logged_actions
--   SELECT * FROM dblink('...same connection...',
--                        'SELECT id, schema_name, table_name, action, row_id,
--                                changed_by, changed_at, old_values, new_values,
--                                changed_columns
--                         FROM audit.logged_actions')
--   AS t (id bigint, schema_name text, table_name text, action text,
--         row_id text, changed_by text, changed_at timestamptz,
--         old_values jsonb, new_values jsonb, changed_columns text[]);
--
-- Option B — using pg_dump (run locally on your machine):
--   pg_dump --host=sofafea-postgres.postgres.database.azure.com
--           --port=5432 --dbname=postgres --username=replicator_user
--           --table=public.users_information --table=audit.logged_actions
--           --data-only --column-inserts > /tmp/backfill.sql
--   psql --host=backup-host --dbname=postgres --username=admin -f /tmp/backfill.sql

-- Step 5: Check the subscription state for the new tables
SELECT srrelid::regclass AS table_name, srsubstate
FROM pg_subscription_rel
WHERE srrelid::regclass::text IN ('users_information', 'audit.logged_actions');

-- Step 6: Monitor replication lag after adding new tables
SELECT
    slot_name,
    active,
    pg_size_pretty(pg_wal_lsn_diff(pg_current_wal_lsn(), confirmed_flush_lsn)) AS lag_size
FROM pg_replication_slots
WHERE slot_name = 'archive_sub';

-- Step 7: Verify row counts match the source (run after replication catches up)
SELECT 'users_information' AS tbl, COUNT(*) AS cnt FROM public.users_information
UNION ALL
SELECT 'audit.logged_actions', COUNT(*) FROM audit.logged_actions;

-- Step 8: Check for any replication errors
SELECT 
*
FROM pg_stat_subscription
WHERE subname = 'archive_sub';