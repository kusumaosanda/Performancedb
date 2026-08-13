#!/usr/bin/env python3
"""
build_techprod.py — bangun data/technician_productivity.json dari CSV mentah.

Sumber : ~/Library/CloudStorage/OneDrive-PT.TelekomunikasiIndonesia/
         Performancedb/technician_productivity/produktivitas_teknisi.csv
Keluaran: data/technician_productivity.json (dibaca halaman Technician Productivity)

    python3 build_techprod.py

Metrik (mengikuti pivot pemilik data):
    produktivitas harian = jumlah order ÷ CHIEF_CODE unik pada tanggal itu
    angka periode        = rata-rata dari nilai harian
Baris tanpa CHIEF_CODE dibuang, sama seperti perilaku pivot Excel.

Wilayah diambil dari sto_hierarchy_map.json, BUKAN kolom REGIONAL/WITEL di CSV.

Jumlah teknisi dihitung ulang di tiap level, tidak dijumlahkan dari level bawah:
~31% teknisi bekerja di lebih dari satu STO, sehingga menjumlahkan STO
menghasilkan kelebihan hitung.

CSV mentah memuat SERVICENUM (nomor pelanggan). Yang ditulis ke JSON hanya
angka agregat — tidak ada identitas pelanggan maupun nama teknisi.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-PT.TelekomunikasiIndonesia/"
    "Performancedb/technician_productivity/produktivitas_teknisi.csv"
)
OUT = os.path.join(HERE, "data", "technician_productivity.json")
MAP = os.path.join(HERE, "sto_hierarchy_map.json")

# Salah ketik yang ada di data sumber; tanpa ini hitungannya terpecah dua.
NORMALISASI = {"KENDALA SYSTEM": "KENDALA SISTEM"}


def main():
    try:
        import pandas as pd
    except ImportError:
        print("❌  pandas belum terpasang. Jalankan:")
        print("    pip3 install pandas openpyxl --break-system-packages")
        sys.exit(1)

    if not os.path.isfile(CSV):
        print(f"❌  File sumber tidak ditemukan:\n    {CSV}")
        print("\n    Taruh produktivitas_teknisi.csv di folder itu, lalu ulangi.")
        sys.exit(1)
    if not os.path.isfile(MAP):
        print(f"❌  sto_hierarchy_map.json tidak ditemukan di {HERE}")
        sys.exit(1)

    print(f"📖  Membaca {os.path.basename(CSV)} "
          f"({os.path.getsize(CSV)/1048576:.1f} MB) …", flush=True)
    df = pd.read_csv(CSV, sep="|", dtype=str, keep_default_na=False,
                     na_values=[""], low_memory=False)
    df.columns = [c.strip().strip('"') for c in df.columns]

    wajib = {"CHIEF_CODE", "STO", "C_STATUSDATE", "FLAG_HOMEPASSID",
             "ERRORCODE_AKHIR", "SUBERRORCODE_AKHIR"}
    kurang = wajib - set(df.columns)
    if kurang:
        print(f"❌  Kolom yang dibutuhkan tidak ada: {', '.join(sorted(kurang))}")
        print(f"    Kolom yang terbaca: {', '.join(df.columns[:12])} …")
        sys.exit(1)

    total_awal = len(df)
    df = df[df["CHIEF_CODE"].notna()].copy()
    print(f"    {total_awal:,} baris → {len(df):,} dipakai "
          f"({total_awal-len(df):,} tanpa CHIEF_CODE dibuang)")

    peta = json.load(open(MAP, encoding="utf-8"))
    df["STO"] = df["STO"].str.strip().str.upper()

    tak_terpeta = sorted(set(df["STO"]) - set(peta))
    if tak_terpeta:
        n = df["STO"].isin(tak_terpeta).sum()
        print(f"⚠   {len(tak_terpeta)} STO tidak ada di sto_hierarchy_map.json "
              f"({n:,} order) — dikeluarkan dari breakdown wilayah:")
        print("    " + ", ".join(tak_terpeta[:15]) + (" …" if len(tak_terpeta) > 15 else ""))
        df = df[~df["STO"].isin(tak_terpeta)]

    for lvl, kol in [("area", "tsel_area"), ("regional", "tsel_regional"),
                     ("branch", "tsel_branch")]:
        df[lvl] = df["STO"].map(lambda s: peta[s][kol])

    dt = pd.to_datetime(df["C_STATUSDATE"], errors="coerce")
    if dt.isna().any():
        print(f"⚠   {dt.isna().sum():,} baris dengan C_STATUSDATE tidak terbaca — dibuang")
        df = df[dt.notna()]
        dt = dt[dt.notna()]
    df["tgl"] = dt.dt.strftime("%Y-%m-%d")
    df["bulan"] = dt.dt.strftime("%Y-%m")
    df["err"] = df["ERRORCODE_AKHIR"].str.upper().str.strip().replace(NORMALISASI)
    df["sub"] = df["SUBERRORCODE_AKHIR"].str.upper().str.strip().replace(NORMALISASI)

    TGL = sorted(df["tgl"].unique())
    IDX = {t: i for i, t in enumerate(TGL)}
    print(f"    periode: {TGL[0]} → {TGL[-1]}  ({len(TGL)} hari)")

    def seri(d):
        """[[indeks_hari, jumlah_order, teknisi_unik], …]"""
        return sorted([IDX[t], int(len(g)), int(g["CHIEF_CODE"].nunique())]
                      for t, g in d.groupby("tgl"))

    VARIAN = {
        "ALL":  df,
        "HPID": df[df["FLAG_HOMEPASSID"].notna()],
        "REG":  df[df["FLAG_HOMEPASSID"].isna()],
    }

    hasil = {"tgl": TGL, "bulan": sorted(df["bulan"].unique()),
             "hier": {}, "lvl": {}, "err": {}}

    for a, ga in df.groupby("area"):
        hasil["hier"][a] = {
            r: {b: sorted(gb["STO"].unique()) for b, gb in gr.groupby("branch")}
            for r, gr in ga.groupby("regional")
        }

    print("🧮  Menghitung produktivitas …", flush=True)
    for vk, vd in VARIAN.items():
        L = {"NASIONAL|": seri(vd)}
        for lvl in ("area", "regional", "branch", "STO"):
            for nama, g in vd.groupby(lvl):
                L[f"{lvl}|{nama}"] = seri(g)
        hasil["lvl"][vk] = L

    def ringkas_kendala(d):
        ber = d[d["err"].notna()]
        return {
            "n": int(len(d)),
            "ok": int(d["err"].isna().sum()),
            "kat": {c: {"n": int(len(g)),
                        "top": [[k, int(v)] for k, v in g["sub"].value_counts().items()]}
                    for c, g in ber.groupby("err")},
        }

    print("🧮  Meringkas kendala …", flush=True)
    hasil["err"]["NASIONAL|"] = ringkas_kendala(df)
    for lvl in ("area", "regional", "branch", "STO"):
        for nama, g in df.groupby(lvl):
            hasil["err"][f"{lvl}|{nama}"] = ringkas_kendala(g)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(hasil, f, separators=(",", ":"))

    nas = hasil["lvl"]["ALL"]["NASIONAL|"]
    prod = sum(o / c for _, o, c in nas if c) / len([x for x in nas if x[2]])
    e = hasil["err"]["NASIONAL|"]

    print(f"\n✅  {OUT}  ({os.path.getsize(OUT)/1024:.0f} KB)")
    print(f"    order              : {len(df):,}")
    print(f"    produktivitas ALL  : {prod:.4f}")
    print(f"    tanpa kendala      : {e['ok']:,} ({e['ok']/e['n']*100:.1f}%)")
    print(f"    entitas wilayah    : {len(hasil['lvl']['ALL']):,}")
    print("\n    Selanjutnya:  ./deploy.sh --source")
    print("    (atau lain kali cukup satu perintah: ./deploy.sh --techprod)\n")


if __name__ == "__main__":
    main()
