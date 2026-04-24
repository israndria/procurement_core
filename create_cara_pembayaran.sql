-- Jalankan SQL ini di Supabase SQL Editor:
-- https://supabase.com/dashboard/project/iubvqphzalodqqhpatcy/sql/new

CREATE TABLE IF NOT EXISTS public.cara_pembayaran (
  id TEXT PRIMARY KEY,
  label TEXT NOT NULL,
  keyword TEXT[] NOT NULL DEFAULT '{}',
  teks TEXT NOT NULL
);

INSERT INTO public.cara_pembayaran (id, label, keyword, teks) VALUES
(
  'monthly_certificate',
  'Monthly Certificate',
  ARRAY['monthly', 'monthly certificate', 'bulanan'],
  '(Monthly Certificate) Pembayaran dilakukan dengan cara didasarkan pada hasil pengukuran bersama atas pekerjaan yang benar-benar telah dilaksanakan secara bulanan'
),
(
  'termin',
  'Termin',
  ARRAY['termin', 'pembayaran termin', 'angsuran'],
  '(Termin) Pembayaran dilakukan dengan cara pembayarannya didasarkan pada hasil pengukuran bersama atas pekerjaan yang benar-benar telah dilaksanakan secara cara angsuran'
)
ON CONFLICT (id) DO UPDATE SET
  label = EXCLUDED.label,
  keyword = EXCLUDED.keyword,
  teks = EXCLUDED.teks;
