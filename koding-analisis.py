import sqlite3
import pandas as pd
import matplotlib.pyplot as plt

# =====================================================
# 1. KONEKSI DATABASE
# =====================================================

db_file = "dummy_sampah.db"
conn = sqlite3.connect(db_file)

# Mengecek nama tabel dalam database
query = """
SELECT name
FROM sqlite_master
WHERE type = 'table'
"""

tabel = pd.read_sql(query, conn)

print("DAFTAR TABEL DALAM DATABASE:")
print(tabel)

print("=" * 60)
print("ANALISIS DATA SAMPAH KOTA SURABAYA")
print("=" * 60)

# =====================================================
# 2. INFORMASI UMUM DATASET
# =====================================================

query = """
SELECT
    COUNT(*) AS jumlah_data,
    SUM(timbulan_kg) AS total_timbulan,
    SUM(sampah_terkelola_kg) AS total_terkelola,
    SUM(sampah_tidak_terkelola_kg) AS total_tidak_terkelola
FROM data_sampah
"""

informasi = pd.read_sql(query, conn)

print("\nINFORMASI UMUM DATASET")
print("-" * 60)
print(informasi)

total_timbulan = informasi.loc[0, "total_timbulan"]
total_terkelola = informasi.loc[0, "total_terkelola"]
total_tidak = informasi.loc[0, "total_tidak_terkelola"]

persen_terkelola = total_terkelola / total_timbulan * 100
persen_tidak = total_tidak / total_timbulan * 100

print(f"\nPersentase sampah terkelola: {persen_terkelola:.2f}%")
print(f"Persentase sampah tidak terkelola: {persen_tidak:.2f}%")

# =====================================================
# 3. ANALISIS JENIS SAMPAH
# =====================================================

query = """
SELECT
    jenis_sampah,
    SUM(timbulan_kg) AS total_timbulan
FROM data_sampah
GROUP BY jenis_sampah
ORDER BY total_timbulan DESC
"""

jenis_sampah = pd.read_sql(query, conn)

print("\nTIMBULAN BERDASARKAN JENIS SAMPAH")
print("-" * 60)
print(jenis_sampah)

print("\nJenis sampah terbanyak:",
      jenis_sampah.iloc[0]["jenis_sampah"])

# =====================================================
# 4. ANALISIS SUMBER SAMPAH
# =====================================================

query = """
SELECT
    sumber_sampah,
    SUM(timbulan_kg) AS total_timbulan
FROM data_sampah
GROUP BY sumber_sampah
ORDER BY total_timbulan DESC
"""

sumber_sampah = pd.read_sql(query, conn)

print("\nTIMBULAN BERDASARKAN SUMBER SAMPAH")
print("-" * 60)
print(sumber_sampah)

print("\nSumber sampah terbesar:",
      sumber_sampah.iloc[0]["sumber_sampah"])

# =====================================================
# 5. ANALISIS PENGELOLAAN SAMPAH
# =====================================================

query = """
SELECT
    SUM(sampah_terkelola_kg) AS total_terkelola,
    SUM(sampah_tidak_terkelola_kg) AS total_tidak_terkelola
FROM data_sampah
"""

pengelolaan = pd.read_sql(query, conn)

print("\nANALISIS PENGELOLAAN SAMPAH")
print("-" * 60)
print(pengelolaan)

print(f"\nSampah terkelola: {persen_terkelola:.2f}%")
print(f"Sampah tidak terkelola: {persen_tidak:.2f}%")

# =====================================================
# 6. ANALISIS METODE PENGELOLAAN
# =====================================================

query = """
SELECT
    metode_pengelolaan,
    SUM(sampah_terkelola_kg) AS total_terkelola
FROM data_sampah
GROUP BY metode_pengelolaan
ORDER BY total_terkelola DESC
"""

metode = pd.read_sql(query, conn)

print("\nSAMPAH TERKELOLA BERDASARKAN METODE")
print("-" * 60)
print(metode)

# =====================================================
# 7. ANALISIS 10 KECAMATAN TERTINGGI
# =====================================================

query = """
SELECT
    kecamatan,
    SUM(timbulan_kg) AS total_timbulan
FROM data_sampah
GROUP BY kecamatan
ORDER BY total_timbulan DESC
LIMIT 10
"""

kecamatan = pd.read_sql(query, conn)

print("\nTOP 10 KECAMATAN BERDASARKAN TIMBULAN SAMPAH")
print("-" * 60)
print(kecamatan)

# =====================================================
# 8. ANALISIS TREN TIMBULAN BULANAN
# =====================================================

query = """
SELECT
    tahun,
    bulan,
    SUM(timbulan_kg) AS total_timbulan
FROM data_sampah
GROUP BY tahun, bulan
ORDER BY tahun, bulan
"""

bulanan = pd.read_sql(query, conn)

print("\nTREN TIMBULAN SAMPAH BULANAN")
print("-" * 60)
print(bulanan)

# =====================================================
# 9. MENYIMPAN HASIL ANALISIS KE EXCEL
# =====================================================

with pd.ExcelWriter("hasil_analisis_sampah.xlsx") as writer:

    informasi.to_excel(
        writer,
        sheet_name="Informasi Umum",
        index=False
    )

    jenis_sampah.to_excel(
        writer,
        sheet_name="Jenis Sampah",
        index=False
    )

    sumber_sampah.to_excel(
        writer,
        sheet_name="Sumber Sampah",
        index=False
    )

    pengelolaan.to_excel(
        writer,
        sheet_name="Pengelolaan",
        index=False
    )

    metode.to_excel(
        writer,
        sheet_name="Metode Pengelolaan",
        index=False
    )

    kecamatan.to_excel(
        writer,
        sheet_name="Top 10 Kecamatan",
        index=False
    )

    bulanan.to_excel(
        writer,
        sheet_name="Tren Bulanan",
        index=False
    )

print("\nHASIL ANALISIS BERHASIL DISIMPAN KE EXCEL!")
print("Nama file: hasil_analisis_sampah.xlsx")

# =====================================================
# 10. VISUALISASI MATPLOTLIB
# =====================================================

# GRAFIK 1: JENIS SAMPAH

plt.figure(figsize=(10, 6))

plt.bar(
    jenis_sampah["jenis_sampah"],
    jenis_sampah["total_timbulan"],
    color="seagreen"
)

plt.title("Timbulan Sampah Berdasarkan Jenis")
plt.xlabel("Jenis Sampah")
plt.ylabel("Total Timbulan (kg)")
plt.xticks(rotation=30)
plt.grid(axis="y", linestyle="--", alpha=0.5)

plt.tight_layout()
plt.show()

# GRAFIK 2: SUMBER SAMPAH

plt.figure(figsize=(10, 6))

plt.bar(
    sumber_sampah["sumber_sampah"],
    sumber_sampah["total_timbulan"],
    color="skyblue"
)

plt.title("Timbulan Sampah Berdasarkan Sumber")
plt.xlabel("Sumber Sampah")
plt.ylabel("Total Timbulan (kg)")
plt.xticks(rotation=30)
plt.grid(axis="y", linestyle="--", alpha=0.5)

plt.tight_layout()
plt.show()

# GRAFIK 3: PENGELOLAAN SAMPAH

plt.figure(figsize=(7, 7))

plt.pie(
    [total_terkelola, total_tidak],
    labels=["Terkelola", "Tidak Terkelola"],
    autopct="%1.1f%%",
    startangle=90,
    colors=["seagreen", "salmon"]
)

plt.title("Persentase Pengelolaan Sampah")
plt.tight_layout()
plt.show()

# GRAFIK 4: 10 KECAMATAN TERTINGGI

plt.figure(figsize=(10, 6))

plt.barh(
    kecamatan["kecamatan"],
    kecamatan["total_timbulan"],
    color="cornflowerblue"
)

plt.title("Top 10 Kecamatan dengan Timbulan Sampah Tertinggi")
plt.xlabel("Total Timbulan (kg)")
plt.ylabel("Kecamatan")
plt.gca().invert_yaxis()
plt.grid(axis="x", linestyle="--", alpha=0.5)

plt.tight_layout()
plt.show()

# GRAFIK 5: TREN TIMBULAN BULANAN

plt.figure(figsize=(10, 6))

plt.plot(
    bulanan["bulan"].astype(str) + "-" + bulanan["tahun"].astype(str),
    bulanan["total_timbulan"],
    marker="o",
    color="darkorange"
)

plt.title("Tren Timbulan Sampah Bulanan")
plt.xlabel("Bulan")
plt.ylabel("Total Timbulan (kg)")
plt.xticks(rotation=45)
plt.grid(True, linestyle="--", alpha=0.5)

plt.tight_layout()
plt.show()

# GRAFIK 6: METODE PENGELOLAAN

plt.figure(figsize=(10, 6))

plt.bar(
    metode["metode_pengelolaan"],
    metode["total_terkelola"],
    color="mediumpurple"
)

plt.title("Sampah Terkelola Berdasarkan Metode Pengelolaan")
plt.xlabel("Metode Pengelolaan")
plt.ylabel("Total Sampah Terkelola (kg)")
plt.xticks(rotation=30)
plt.grid(axis="y", linestyle="--", alpha=0.5)

plt.tight_layout()
plt.show()

# =====================================================
# 11. MENUTUP DATABASE
# =====================================================

conn.close()

print("\n" + "=" * 60)
print("ANALISIS DATA SELESAI")
print("=" * 60)