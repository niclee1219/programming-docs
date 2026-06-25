-- Add updated_at audit timestamp to per-user tables.
-- DEFAULT NOW() populates existing rows on migration; the server sets it explicitly on every upsert.

ALTER TABLE user_settings    ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT NOW();
ALTER TABLE field_mappings   ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT NOW();
ALTER TABLE director_history ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT NOW();
