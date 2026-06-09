SELECT pubname FROM pg_publication;

ALTER USER replicator_user WITH PASSWORD 'SofAFeA35.1';

SELECT pubname, puballtables FROM pg_publication;

SELECT schemaname, tablename FROM pg_publication_tables;

SELECT 'locations' AS table_name, COUNT(*) AS row_count FROM locations
UNION ALL
SELECT 'usct_connecticut', COUNT(*) FROM usct_connecticut
UNION ALL
SELECT 'family_members', COUNT(*) FROM family_members
UNION ALL
SELECT 'book_of_negroes', COUNT(*) FROM book_of_negroes;


-- Remove test_repl from publication and drop the table
ALTER PUBLICATION archive_pub DROP TABLE test_repl;
DROP TABLE IF EXISTS test_repl;


SELECT slot_name, active, 
       pg_wal_lsn_diff(pg_current_wal_lsn(), confirmed_flush_lsn) AS pending_bytes
FROM pg_replication_slots 
WHERE slot_name = 'archive_sub';

SELECT pg_drop_replication_slot('archive_sub');

SELECT slot_name, active FROM pg_replication_slots;

-- On source (as admin user or user with replication role)
SELECT slot_name, active FROM pg_replication_slots WHERE slot_name = 'archive_sub';

SELECT pubname, puballtables FROM pg_publication WHERE pubname = 'archive_pub';

DROP PUBLICATION archive_pub;
CREATE PUBLICATION archive_pub FOR TABLE 
    locations, 
    usct_connecticut, 
    family_members, 
    book_of_negroes;

SELECT relname, pg_size_pretty(pg_total_relation_size(relid)) AS total_size
FROM pg_catalog.pg_statio_user_tables
WHERE relname IN ('locations', 'usct_connecticut', 'family_members', 'book_of_negroes')
ORDER BY pg_total_relation_size(relid) DESC;

SELECT table_name, constraint_name 
FROM information_schema.table_constraints 
WHERE constraint_type = 'PRIMARY KEY' 
  AND table_name IN ('locations', 'usct_connecticut', 'family_members', 'book_of_negroes');


  SELECT COUNT(*) FROM locations;
SELECT COUNT(*) FROM usct_connecticut;
SELECT COUNT(*) FROM family_members;
SELECT COUNT(*) FROM book_of_negroes;


SELECT
  'ALTER TABLE ' || conrelid::regclass || 
  ' ADD CONSTRAINT ' || conname || ' ' || 
  pg_get_constraintdef(oid) || ';' AS add_constraint_stmt
FROM pg_constraint
WHERE contype = 'f'
  AND conrelid::regclass IN (
    'locations'::regclass,
    'usct_connecticut'::regclass,
    'family_members'::regclass,
    'book_of_negroes'::regclass
  )
ORDER BY conname;

INSERT INTO locations (locations_id) VALUES (9999);


-- ============================================================================
--  ADD AUDIT TRAIL & USERS TABLES TO PUBLICATION
-- ============================================================================
-- The audit.logged_actions and public.users_information tables need to be
-- replicated to the backup database so the audit trail is preserved.

-- Step 1: Verify the current publication members
SELECT schemaname, tablename
FROM pg_publication_tables
WHERE pubname = 'archive_pub';

-- Step 2: Add users_information and audit.logged_actions to the publication
ALTER PUBLICATION archive_pub ADD TABLE public.users_information;
ALTER PUBLICATION archive_pub ADD TABLE audit.logged_actions;

-- Step 3: Verify the tables are now in the publication
SELECT schemaname, tablename
FROM pg_publication_tables
WHERE pubname = 'archive_pub'
ORDER BY schemaname, tablename;

-- Step 4: Check row counts for the new tables
SELECT 'users_information' AS table_name, COUNT(*) FROM public.users_information
UNION ALL
SELECT 'audit.logged_actions', COUNT(*) FROM audit.logged_actions;

-- Step 5: Generate FK constraints for users_information (if any exist)
SELECT
  'ALTER TABLE ' || conrelid::regclass ||
  ' ADD CONSTRAINT ' || conname || ' ' ||
  pg_get_constraintdef(oid) || ';' AS add_constraint_stmt
FROM pg_constraint
WHERE contype = 'f'
  AND (conrelid::regclass = 'users_information'::regclass
       OR confrelid::regclass = 'users_information'::regclass)
ORDER BY conname;


-- ============================================================================
--  ADD sofafea_members TO PUBLICATION (new table for SOFAFEA member data)
-- ============================================================================

-- Step 1: Add sofafea_members to the publication
ALTER PUBLICATION archive_pub ADD TABLE public.sofafea_members;

-- Step 2: Verify the table is in the publication
SELECT schemaname, tablename
FROM pg_publication_tables
WHERE pubname = 'archive_pub'
ORDER BY schemaname, tablename;

-- Step 3: Check row counts for all replicated tables
SELECT 'locations' AS table_name, COUNT(*) FROM public.locations
UNION ALL
SELECT 'usct_connecticut', COUNT(*) FROM public.usct_connecticut
UNION ALL
SELECT 'family_members', COUNT(*) FROM public.family_members
UNION ALL
SELECT 'book_of_negroes', COUNT(*) FROM public.book_of_negroes
UNION ALL
SELECT 'users_information', COUNT(*) FROM public.users_information
UNION ALL
SELECT 'audit.logged_actions', COUNT(*) FROM audit.logged_actions
UNION ALL
SELECT 'sofafea_members', COUNT(*) FROM public.sofafea_members;

-- Step 4: Generate FK constraints for sofafea_members
SELECT
  'ALTER TABLE ' || conrelid::regclass ||
  ' ADD CONSTRAINT ' || conname || ' ' ||
  pg_get_constraintdef(oid) || ';' AS add_constraint_stmt
FROM pg_constraint
WHERE contype = 'f'
  AND (conrelid::regclass = 'sofafea_members'::regclass
       OR confrelid::regclass = 'sofafea_members'::regclass)
ORDER BY conname;
