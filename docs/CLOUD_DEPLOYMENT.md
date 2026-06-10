# OrderFlow Cloud Deployment

Recommended stack:

- Render for the Python web app.
- Supabase for shared application state.

## 1. Create Supabase State Table

In Supabase, open SQL Editor and run:

```sql
-- docs/SUPABASE_SETUP.sql
create table if not exists public.orderflow_app_state (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

create or replace function public.set_orderflow_app_state_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists orderflow_app_state_updated_at on public.orderflow_app_state;

create trigger orderflow_app_state_updated_at
before update on public.orderflow_app_state
for each row
execute function public.set_orderflow_app_state_updated_at();

alter table public.orderflow_app_state enable row level security;

drop policy if exists "OrderFlow service role can manage app state" on public.orderflow_app_state;

create policy "OrderFlow service role can manage app state"
on public.orderflow_app_state
for all
using (auth.role() = 'service_role')
with check (auth.role() = 'service_role');
```

## 2. Render Environment Variables

Set these in Render:

```text
HOST=0.0.0.0
GREENOPS_AUTH_SECRET=<long-random-secret>
SUPABASE_URL=https://<project-ref>.supabase.co
SUPABASE_SERVICE_ROLE_KEY=<service-role-key>
ORDERFLOW_USERS_JSON=[{"email":"planner@company.com","name":"Planner","role":"planner","password":"change-me"}]
```

`SUPABASE_URL` can be the base project URL or the REST URL from Supabase. Both work:

```text
https://<project-ref>.supabase.co
https://<project-ref>.supabase.co/rest/v1/
```

Use the Supabase service role key only on the Render server. Do not expose it in browser code.

## 3. Data Stored in Supabase

OrderFlow stores JSON state under these keys:

- `team_state:v1`: work sessions, client workspaces, leverschema and laadschema data.
- `shortage_sessions:v1`: Manco order sessions.
- `shortage_day_sessions:v1`: Manco day/session list.

The local JSON files remain as fallback for development when Supabase env vars are not configured.

## 4. Render Deploy

Render can use `render.yaml` from this repository.

Build command:

```text
pip install -r requirements.txt
```

Start command:

```text
python app.py
```

## 5. First Migration From Local Data

After Supabase env vars are configured, migrate the current local JSON files with:

```powershell
$env:SUPABASE_URL="https://<project-ref>.supabase.co"
$env:SUPABASE_SERVICE_ROLE_KEY="<service-role-key>"
py tools\migrate_local_state_to_supabase.py
```

This copies:

- `data/team_state.json` -> key `team_state:v1`
- `data/shortage_sessions.json` -> key `shortage_sessions:v1`
- `data/shortage_day_sessions.json` -> key `shortage_day_sessions:v1`
