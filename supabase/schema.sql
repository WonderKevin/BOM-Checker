-- Run in Supabase SQL Editor for project: WM BOM Checker

create extension if not exists pgcrypto;

create table if not exists public.items (
  id uuid primary key default gen_random_uuid(),
  description text not null unique,
  item_code text not null unique,
  created_at timestamptz not null default now()
);

create table if not exists public.uploads (
  id uuid primary key default gen_random_uuid(),
  month_label text not null,
  batch_name text not null,
  upload_type text not null check (upload_type in ('WM','CBF')),
  file_name text not null,
  storage_path text not null,
  created_at timestamptz not null default now()
);

create table if not exists public.bom_values (
  id uuid primary key default gen_random_uuid(),
  upload_id uuid references public.uploads(id) on delete cascade,
  month_label text not null,
  batch_name text not null,
  upload_type text not null check (upload_type in ('WM','CBF')),
  material_name text not null,
  item_code text,
  production_code text,
  usage numeric not null default 0,
  created_at timestamptz not null default now()
);

insert into storage.buckets (id, name, public)
values ('bom-files', 'bom-files', false)
on conflict (id) do nothing;

alter table public.items enable row level security;
alter table public.uploads enable row level security;
alter table public.bom_values enable row level security;

-- Local/dev open policies. Tighten these before production auth rollout.
drop policy if exists "dev read items" on public.items;
drop policy if exists "dev write items" on public.items;
drop policy if exists "dev read uploads" on public.uploads;
drop policy if exists "dev write uploads" on public.uploads;
drop policy if exists "dev read bom values" on public.bom_values;
drop policy if exists "dev write bom values" on public.bom_values;
create policy "dev read items" on public.items for select using (true);
create policy "dev write items" on public.items for all using (true) with check (true);
create policy "dev read uploads" on public.uploads for select using (true);
create policy "dev write uploads" on public.uploads for all using (true) with check (true);
create policy "dev read bom values" on public.bom_values for select using (true);
create policy "dev write bom values" on public.bom_values for all using (true) with check (true);

create policy "dev storage read" on storage.objects for select using (bucket_id = 'bom-files');
create policy "dev storage write" on storage.objects for insert with check (bucket_id = 'bom-files');
create policy "dev storage update" on storage.objects for update using (bucket_id = 'bom-files');

insert into public.items (description, item_code) values
('Crystalline Allulose','IW741'),
('Monk Fruit Concentrate','IW796'),
('Erythritol FN Powder','IN954'),
('Erythritol STD Granular','IW810'),
('Organic Coconut Sugar','IW816'),
('Wonder Monday Crumbs - Erythritol','IW815'),
('Wonder Monday Crumbs - Coconut Sugar','IW797'),
('Wonder Monday Crumbs - Cane Sugar','IW793'),
('Natural Graham Flavor','IN838'),
('White Chocolate Raspberry Flavor','IN843'),
('Maple Pecan Flavor','IN840'),
('Salted Caramel Type Extract','IN864'),
('Vanilla Bean','IN926'),
('Lemon Meringue Nat Flavor','IN928'),
('Enrobe','IW777'),
('SB 2-Pack Bites Labels','PFW1020-12721'),
('PB 2-Pack Bites Labels','PFW1020-12730'),
('Wonder Monday 3" Lids','PP5002'),
('Wonder Monday Blank 3" Bases','PP5003'),
('Wonder Monday 3" Clear Base','PP5004'),
('M167 - Emerson Temp Monitor','9991-1'),
('M166 - tempTRIP Temp Monitor','9993-1'),
('M169 - Sensitech Temp Monitor','9999-1'),
('Factor US Carton - Lemon Blueberry','PB2604-14459 LBB'),
('Factor US Carton - Vanilla Raspberry','PB2604-14403 VR'),
('Factor US Carton - Double Chocolate','PB2604-14417'),
('Factor US Carton - Creme Brulee','PB2604-14461'),
('HelloFresh Carton - Classic Plain','PB2607-14301'),
('HelloFresh Carton - Strawberry Bliss','PB2607-14321'),
('Factor Canada Carton - Double chocolate','PB2604-14617'),
('Factor Canada Carton - Pumpkin spice','PB2604-14608'),
('Factor Canada Carton - Vanilla Caramel','PB2604-14689'),
('Factor Black Cartons','PB2604-Black'),
('Wildgrain Blank Cartons','PB2607'),
('HelloFresh / Factor Master Case','PC2838'),
('Target 6 pack carton','PC2894'),
('SRP Mastercase 12pk','PC2891'),
('Red Velvet Flavor','IN841'),
('SupraRed','IN947'),
('BetaCarotene/Altratene','IN948')
on conflict do nothing;
