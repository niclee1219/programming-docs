-- Enable RLS on all tables (implicit deny for non-service-role connections).
-- The Flask server uses the service role key which bypasses RLS — app behaviour unchanged.
-- Direct access via psql, Supabase dashboard anon/authenticated roles, or leaked DB URL is blocked.

ALTER TABLE user_settings       ENABLE ROW LEVEL SECURITY;
ALTER TABLE field_mappings      ENABLE ROW LEVEL SECURITY;
ALTER TABLE director_history    ENABLE ROW LEVEL SECURITY;
ALTER TABLE company_snapshots   ENABLE ROW LEVEL SECURITY;
