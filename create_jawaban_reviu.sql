-- Tabel jawaban reviu DPP — 1 row per sesi rapat
-- Jalankan di: https://supabase.com/dashboard/project/<YOUR_SUPABASE_PROJECT_ID>/sql/new

CREATE TABLE IF NOT EXISTS public.jawaban_reviu (
  id_sesi         TEXT PRIMARY KEY DEFAULT 'default',
  updated_at      TIMESTAMPTZ DEFAULT now(),

  -- Rekomendasi hasil reviu (8 seksi A-H)
  rekomen_1       TEXT DEFAULT '',
  rekomen_2       TEXT DEFAULT '',
  rekomen_3       TEXT DEFAULT '',
  rekomen_4       TEXT DEFAULT '',
  rekomen_5       TEXT DEFAULT '',
  rekomen_6       TEXT DEFAULT '',
  rekomen_7       TEXT DEFAULT '',
  rekomen_8       TEXT DEFAULT '',

  -- Tanggapan PPK (8 seksi A-H)
  tanggapan_1     TEXT DEFAULT '',
  tanggapan_2     TEXT DEFAULT '',
  tanggapan_3     TEXT DEFAULT '',
  tanggapan_4     TEXT DEFAULT '',
  tanggapan_5     TEXT DEFAULT '',
  tanggapan_6     TEXT DEFAULT '',
  tanggapan_7     TEXT DEFAULT '',
  tanggapan_8     TEXT DEFAULT '',

  -- Catatan kolom per tabel A-H (JSON array string per tabel)
  catatan_tabel_1 TEXT DEFAULT '[]',
  catatan_tabel_2 TEXT DEFAULT '[]',
  catatan_tabel_3 TEXT DEFAULT '[]',
  catatan_tabel_4 TEXT DEFAULT '[]',
  catatan_tabel_5 TEXT DEFAULT '[]',
  catatan_tabel_6 TEXT DEFAULT '[]',
  catatan_tabel_7 TEXT DEFAULT '[]',
  catatan_tabel_8 TEXT DEFAULT '[]'
);

-- Drop dulu jika sudah ada versi lama (tanpa catatan_tabel)
-- Uncomment baris berikut jika perlu reset:
-- DROP TABLE IF EXISTS public.jawaban_reviu;

-- Insert row default kosong
INSERT INTO public.jawaban_reviu (id_sesi)
VALUES ('default')
ON CONFLICT (id_sesi) DO NOTHING;

-- Tambah kolom catatan_tabel jika tabel sudah ada tapi belum punya kolom ini
DO $$
BEGIN
  FOR i IN 1..8 LOOP
    IF NOT EXISTS (
      SELECT 1 FROM information_schema.columns
      WHERE table_name='jawaban_reviu'
      AND column_name='catatan_tabel_' || i
    ) THEN
      EXECUTE 'ALTER TABLE public.jawaban_reviu ADD COLUMN catatan_tabel_' || i || ' TEXT DEFAULT ''[]''';
    END IF;
  END LOOP;
END $$;
