SUPABASE POSTGRES MIGRATION

Files:
- db.py: replaces SQLite persistence with Supabase PostgreSQL.
- requirements.txt: adds psycopg2-binary.
- schema.sql: optional SQL Editor setup (the app also runs CREATE TABLE IF NOT EXISTS).
- .env.example: connection-string template.
- app.py: same app with its backup download correctly labelled as JSON.

Setup:
1. In Supabase Dashboard, open Project Settings / Connect and copy a PostgreSQL connection string.
2. Prefer the Session pooler for app deployments where direct IPv6 connectivity is unavailable.
3. Set SUPABASE_DB_URL to that complete connection string. Do not commit it to Git.
4. Install dependencies: pip install -r requirements.txt
5. Run the SQL in schema.sql in Supabase SQL Editor, or let db.init_db() create the tables.
6. Run: streamlit run app.py

IMPORTANT:
- Existing SQLite data is NOT automatically copied to Supabase. Export your current app data to Excel first, then import that Excel in the migrated app.
- Supabase credentials must remain server-side. Never put the database password or service-role key in frontend/browser code.
- The JSON backup is an application-level data export, not a replacement for Supabase's managed backups.
- If you enable Supabase RLS on these tables, configure appropriate policies for the database role used by the app. A server-side Streamlit app using a protected database connection should not expose that connection to users.
