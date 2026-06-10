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
