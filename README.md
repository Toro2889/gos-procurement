# GOS Group - Procurement Portal

Aplikasi web internal untuk pengajuan dan approval pengadaan barang/jasa di GOS Group. Single-page app statis (HTML/CSS/JS murni, tanpa build step) yang login via Microsoft 365 SSO dan menyimpan seluruh data di **SharePoint Lists**, diakses lewat Microsoft Graph API.

## Fitur

- **Pengajuan Baru** — form pengajuan pengadaan (Barang/Jasa) dengan:
  - Pemilihan Company → Client → Project → Cabang (bertingkat, sesuai Master Data).
  - Preview otomatis siapa approver L1/L2 begitu Entitas & Proyek terisi lengkap.
  - Upload dokumen pendukung, nomor pengajuan auto-generate.
- **Dashboard** — daftar seluruh pengajuan dengan:
  - Search, sort per kolom (klik header), dan filter tick per kolom ala Excel.
  - Filter status, company, rentang tanggal, serta toggle "Perlu Saya Approve".
  - Detail per pengajuan: approve/reject L1 & L2, update status lanjutan oleh GA (Approved → Submitted to Finance → Delivered), input harga real, export PDF.
  - Pengaju bisa membatalkan (**Cancel**) pengajuannya sendiri selama masih Pending L1/L2.
  - Export ke Excel.
- **Master Data** (khusus role GA/Approver):
  - Entitas (Company/Client/Project), Barang, Vendor/Jasa — CRUD + aktif/nonaktifkan, search, sort, filter tick per kolom.
  - Setting Approval — konfigurasi threshold L2 dan rule approval L1/L2 per Project + Cabang.
  - Daftar Pengaju — kontrol siapa saja yang boleh membuat pengajuan baru.
- **Dark/Light mode** — toggle di header, preferensi tersimpan di browser (localStorage).

## Tech Stack

- Vanilla HTML/CSS/JavaScript (ES5-style, tanpa framework/bundler).
- **MSAL Browser** (`@azure/msal-browser`) — autentikasi Microsoft 365 (Azure AD / Entra ID).
- **Microsoft Graph API** — seluruh baca/tulis data (SharePoint Lists) dan pengiriman notifikasi email.
- **SheetJS (xlsx)** — export data Dashboard ke Excel.

## Struktur File

```
index.html   Markup halaman (login, form pengajuan, dashboard, master data, modal-modal)
app.js       Seluruh logika: auth, Graph API calls, rendering, validasi, approval workflow
style.css    Styling + tema terang/gelap (CSS custom properties)
```

## Konfigurasi (Azure AD)

Diatur di awal `app.js` (`CONFIG`):

| Key | Keterangan |
|---|---|
| `clientId` | App registration (client) ID di Azure AD |
| `tenantId` | Directory (tenant) ID |
| `siteUrl` | URL SharePoint site tempat semua list disimpan |
| `masterList` / `resultList` | Nama list utama (lihat tabel list di bawah) |
| `approverEmail` | Email fallback jika rule approval belum ketemu |
| `redirectUri` | Otomatis mengikuti origin halaman — **harus didaftarkan sebagai Redirect URI (SPA)** di App Registration Azure AD agar login berhasil di domain/host tersebut |

Daftar email dengan akses penuh (GA/Master Data) diatur di `APPROVERS`.

## SharePoint Lists yang Dibutuhkan

Semua list berada di site yang sama (`CONFIG.siteUrl`). Kolom di bawah **harus ada** di masing-masing list (tipe Choice untuk `Status` harus memuat semua opsi yang tercantum):

### `MasterPengadaan` (Entitas)
`Company`, `CompanyCode`, `Client`, `Project`, `Aktif` (Yes/No)

### `MasterBarang`
`NamaBarang`, `KategoriBarang`, `Subkategori`, `Satuan`, `HargaEstimasi`, `Aktif`

### `MasterJasa`
`NamaVendor`, `Kategori`, `DomisiliVendor`, `PICVendor`, `NoTelpon`, `Aktif`

### `KonfigurasiApproval`
Satu list, dibedakan lewat kolom `TipeKonfigurasi` (`L2Setting` / `ApprovalRule` / `SubmitterAccess` / `Deleted`):
`TipeKonfigurasi`, `Project`, `Cabang`, `L1Email`, `L1Name`, `L2Email`, `L2Name`, `L2Threshold`, `SubmitterEmail`

### `HasilPengajuan` (data pengajuan)
`NomorPengajuan`, `NomorUrut`, `TanggalPengajuan`, `JenisPengadaan`, `Company`, `Client`, `Project`, `Cabang`, `TujuanPermintaan`, `JenisProduk`, `EstimasiHarga`, `DokumenPendukung`, `DetailItem`, `DetailItemReal`, `HargaReal`, `SubmittedBy`, `SubmittedByEmail`, `L1ApproverEmail`, `L1ApproverName`, `L1Notes`, `L1ApprovalDate`, `L2ApproverEmail`, `L2ApproverName`, `L2Notes`, `L2ApprovalDate`, `ApproverNotes`, `TanggalSubmittedToFinance`, `NamaPenerima`, `TanggalTerima`, `CancelNotes`, `CancelDate`,
`Status` (Choice): `Pending L1`, `Pending L2`, `Approved`, `Submitted to Finance`, `Delivered`, `Rejected`, `Cancelled`

## Menjalankan Lokal

```bash
python3 -m http.server 5500
# buka http://127.0.0.1:5500/index.html
```

> Login SSO hanya akan berhasil jika host yang diakses (misal `http://127.0.0.1:5500/index.html`) sudah didaftarkan sebagai Redirect URI di App Registration Azure AD. Tanpa itu, halaman bisa dibuka tapi proses login akan gagal — ini bukan bug, melainkan pembatasan keamanan OAuth yang memang harus dikonfigurasi lebih dulu oleh admin Azure AD.

## Keamanan

- Semua teks bebas dari input pengguna (nama cabang, catatan, alasan pembatalan, dll.) di-escape lewat helper `eA()` sebelum dirender sebagai HTML — jangan pernah menyisipkan input pengguna ke `innerHTML` tanpa lewat `eA()`.
- Akses ke Master Data dan approval dibatasi lewat daftar `APPROVERS` dan rule di `KonfigurasiApproval` — validasi akses tetap harus dianggap sebagai lapisan UI saja; kontrol akses sesungguhnya ada di permission SharePoint/Graph API di sisi Microsoft 365.
