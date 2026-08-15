-- Allowlist eksplisit paket yang boleh dipantau scheduler GCal.
-- Jalankan manual di Supabase SQL Editor; DDL tidak dijalankan oleh aplikasi.
create table if not exists public.calendar_sync_targets (
    scope        text not null default 'POKJA2026',
    jenis_paket  text not null check (jenis_paket in ('tender', 'pl')),
    kode_paket   text not null,
    nama_paket   text,
    folder_name  text,
    enabled      boolean not null default true,
    source       text not null default 'manual',
    note         text,
    created_at   timestamptz not null default now(),
    updated_at   timestamptz not null default now(),
    primary key (scope, jenis_paket, kode_paket)
);

alter table public.calendar_sync_targets
    add column if not exists folder_name text;

create index if not exists idx_calendar_sync_targets_active
    on public.calendar_sync_targets (scope, jenis_paket, enabled);

comment on table public.calendar_sync_targets is
    'Explicit allowlist paket milik scope yang boleh discrape ke Google Calendar.';
