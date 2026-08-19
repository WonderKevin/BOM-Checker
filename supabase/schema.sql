-- Run in Supabase SQL Editor for project: WM BOM Checker

create extension if not exists pgcrypto;

create table if not exists public.items (
  id uuid primary key default gen_random_uuid(),
  description text not null,
  item_code text not null unique,
  row_order integer,
  created_at timestamptz not null default now()
);

alter table public.items drop constraint if exists items_description_key;

create table if not exists public.bom_uploads (
  id uuid primary key default gen_random_uuid(),
  bom_type text not null check (bom_type in ('WM','CBF')),
  month text not null,
  batch text not null,
  file_name text not null,
  storage_path text,
  created_at timestamptz not null default now()
);

create table if not exists public.bom_usage_rows (
  id uuid primary key default gen_random_uuid(),
  upload_id uuid references public.bom_uploads(id) on delete set null,
  bom_type text not null check (bom_type in ('WM','CBF')),
  month text not null,
  batch text not null,
  item_code text not null,
  description text not null,
  production_code text not null,
  usage_lbs numeric not null default 0,
  wm_usage numeric,
  cbf_usage numeric,
  row_order integer,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists bom_usage_rows_lookup_idx
  on public.bom_usage_rows (month, batch, item_code, production_code, bom_type);

insert into storage.buckets (id, name, public)
values ('bom-files', 'bom-files', false)
on conflict (id) do nothing;

alter table public.items enable row level security;
alter table public.bom_uploads enable row level security;
alter table public.bom_usage_rows enable row level security;

-- Local/dev open policies. Tighten these before production auth rollout.
drop policy if exists "dev read items" on public.items;
drop policy if exists "dev write items" on public.items;
drop policy if exists "dev read uploads" on public.bom_uploads;
drop policy if exists "dev write uploads" on public.bom_uploads;
drop policy if exists "dev read usage rows" on public.bom_usage_rows;
drop policy if exists "dev write usage rows" on public.bom_usage_rows;
drop policy if exists "dev storage read" on storage.objects;
drop policy if exists "dev storage write" on storage.objects;
drop policy if exists "dev storage update" on storage.objects;

create policy "dev read items" on public.items for select using (true);
create policy "dev write items" on public.items for all using (true) with check (true);
create policy "dev read uploads" on public.bom_uploads for select using (true);
create policy "dev write uploads" on public.bom_uploads for all using (true) with check (true);
create policy "dev read usage rows" on public.bom_usage_rows for select using (true);
create policy "dev write usage rows" on public.bom_usage_rows for all using (true) with check (true);
create policy "dev storage read" on storage.objects for select using (bucket_id = 'bom-files');
create policy "dev storage write" on storage.objects for insert with check (bucket_id = 'bom-files');
create policy "dev storage update" on storage.objects for update using (bucket_id = 'bom-files');

insert into public.items (description, item_code, row_order) values
('Crystalline Allulose','IW741',0),
('Monk Fruit Concentrate','IW796',1),
('Erythritol FN Powder','IN954',2),
('Erythritol STD Granular','IW810',3),
('Organic Coconut Sugar','IW816',4),
('Wonder Monday Crumbs - Erythritol','IW815',5),
('Wonder Monday Crumbs - Coconut Sugar','IW797',6),
('Wonder Monday Crumbs - Cane Sugar','IW793',7),
('Natural Graham Flavor','IN838',8),
('White Chocolate Raspberry Flavor','IN843',9),
('Maple Pecan Flavor','IN840',10),
('Salted Caramel Type Extract','IN864',11),
('Vanilla Bean','IN926',12),
('Lemon Meringue Nat Flavor','IN928',13),
('Enrobe','IW777',14),
('SB 2-Pack Bites Labels','PFW1020-12721',15),
('PB 2-Pack Bites Labels','PFW1020-12730',16),
('Wonder Monday 3" Lids','PP5002',17),
('Wonder Monday Blank 3" Bases','PP5003',18),
('Wonder Monday 3" Clear Base','PP5004',19),
('Wonder Monday 3" Lids','PP2723',20),
('Black 3" Bases','PP2721',21),
('Wonder Monday 3" Clear Base','PP2722',22),
('Target 6 pack carton','PC2894',23),
('SRP Mastercase 12pk','PC2891',24)
on conflict (item_code) do update set
  description = excluded.description,
  row_order = excluded.row_order;
