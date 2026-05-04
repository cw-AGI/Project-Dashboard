# Supabase (optional cloud data)

1. Create a project at [supabase.com](https://supabase.com).
2. In the SQL Editor, run `migrations/0001_dashboard_projects.sql`.
3. In **Authentication → Users**, add at least one user (email + password) for admins who will save data online.
4. In `WN_Dashboard_v4_6-dark.html`, set `SUPABASE_REMOTE.url` and `SUPABASE_REMOTE.anonKey` (Project Settings → API).
5. Deploy the HTML to GitHub Pages (or any static host). Open the site, click the **cloud** icon, sign in with the Supabase user, then use **Admin** and **Save** as usual.

If the `dashboard_projects` table has **no rows**, the app keeps using `projects/index.json` + JSON files from the repo. After the first successful cloud save (or manual SQL insert), the app loads the project list from Supabase.
