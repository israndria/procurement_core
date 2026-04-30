create table if not exists draft_paket (
    kode_tender       text primary key,
    nama_tender       text,
    mak               text,
    kode_rup          text,
    nilai_pagu        text,
    nilai_hps         text,
    link_pdf          text,
    kode_pokja        text,
    nomor_pp          text,
    nomor_surat_dinas text,
    nama_dinas        text,
    nama_ppk          text,
    jangka_waktu      text,
    sumber_anggaran   text,
    diambil_pada      timestamptz default now()
);
