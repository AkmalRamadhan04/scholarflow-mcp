# ScholarFlow MCP (Bahasa Indonesia)

**MCP Server & Agent Skill Terpadu untuk Riset Akademik dan Systematic Literature Review (SLR) yang 100% Reprodusibel**

> **English documentation:** Lihat [README.md](README.md).

**ScholarFlow MCP** menggabungkan orkestrasi metodologi riset akademik (`skills/scholarflow/SKILL.md`) dengan **14 tool MCP deterministik** berbasis Python (`FastMCP`) agar seluruh komponen riset Anda—berkas PDF fisik, log pencarian 5 jalur, matriks ekstraksi kode tertutup (ADO–TCM), lembar uji reliabilitas pengode manusia buta (*Cohen's Kappa* & *PABAK*), berkas referensi `.bib`/`.ris`, gambar 300 DPI, hingga naskah `.docx`/`.md`—selalu **sinkron 100% dan teraudit otomatis**.

---

## 7 Keunggulan Utama Dibanding Alur Riset AI Biasa

1. **Rekonsiliasi 5 Jalur SPAR-4-SLR & PRISMA 2020** — memverifikasi secara matematis seluruh persamaan *Assembling → Arranging → Assessing* (Scopus, Non-Scopus, Web Pelengkap, Korpus Awal Tim, Snowballing Maju/Mundur).
2. **Manajemen Fisik Berkas Literatur & Routing Non-Destruktif** — menata otomatis folder `1_Inti`, `2_Belum_Diklasifikasi`, `3_Dikecualikan`, `6_Referensi_Latar`, dan `arsip_salah_saring_kueri` (+ manifest CSV), sekaligus memverifikasi keterbacaan setiap PDF menggunakan `PyMuPDF`.
3. **Akuisisi Teks Lengkap Hibrida (Legal OA + Antrean Manual)** — mengunduh PDF Open Access melalui OpenAlex & metadata `citation_pdf_url`/OJS (secara tegas menolak/melewati CAPTCHA), lalu menyusun sheet `Unduh_Teks_Lengkap` berisi tautan klik aktif untuk kandidat yang memerlukan akses institusi.
4. **Matriks Kode Tertutup + Catatan Bacaan 1-ke-1** — mendukung dimensi **ADO–TCM** (*Antecedents, Decisions, Outcomes, Theories, Contexts, Methods*) dan menerapkan *dropdown* Excel aman berbasis rentang sel (menghindari galat korupsi Excel `>255` karakter).
5. **Uji Reliabilitas Pengode Manusia 100% Buta (*Blinded*)** — mengekspor lembar kerja Excel terpisah tanpa membocorkan hasil Pengode 1, lengkap dengan pembagian beban kerja berstrata + irisan bersama, lalu menghitung otomatis `% Agreement`, **Cohen's Kappa ($\kappa$)**, dan **PABAK**.
6. **Sinkronisasi Naskah *Single Source of Truth*** — membangun ulang naskah `.docx` & `.md`, tabel lampiran korpus, dan 3 gambar 300 DPI langsung dari `Matriks_Induk.xlsx` & `Log_Pencarian.xlsx`.
7. **Audit Integritas 4-Arah Deterministik (`20+` Pengujian)** — memeriksa serentak kecocokan lintas berkas fisik PDF ↔ Log Pencarian ↔ Matriks & Catatan Bacaan ↔ Naskah, Sitasi 100% di batang tubuh, `.bib`/`.ris` (0 galat penulis `&`), dan Protokol.

---

## Dukungan Preset Bahasa Indonesia (`language="id"`)

Saat memanggil `init_research_project` dengan parameter `language="id"`, ScholarFlow secara otomatis membuat struktur direktori berbahasa Indonesia:

```text
00_Admin/Protokol_Tinjauan.docx
01_Naskah/Naskah_Tersinkronisasi.docx  (+ gambar/ 300 DPI)
02_Literatur/{1_Inti, 2_Belum_Diklasifikasi, 3_Dikecualikan, 4_Referensi_Metodologi, 5_Regulasi, 6_Referensi_Latar, arsip_salah_saring_kueri}
03_Ekstraksi_Koding/{Log_Pencarian.xlsx, Matriks_Induk.xlsx, Lembar_Kerja_Buta_Pengode2.xlsx, Catatan_Bacaan/}
04_Tools/{Referensi.bib, Referensi.ris, integrity_audit_report.json}
```

---

## Instalasi Cepat

```bash
# Jalankan langsung via uvx dari GitHub
uvx --from git+https://github.com/AkmalRamadhan04/scholarflow-mcp scholarflow-mcp
```

Konfigurasi di **Google Antigravity** (`mcp_config.json`), **Claude Desktop**, atau **Cursor**:

```json
{
  "mcpServers": {
    "scholarflow": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/AkmalRamadhan04/scholarflow-mcp",
        "scholarflow-mcp"
      ]
    }
  }
}
```

---

## Lisensi & Penulis

Dikembangkan oleh **Akmal Ramadhan** ([@AkmalRamadhan04](https://github.com/AkmalRamadhan04)) di bawah [Lisensi MIT](LICENSE).
