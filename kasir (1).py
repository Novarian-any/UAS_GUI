import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os

# ===== CLASS: DataManager =====
# Mengelola data produk dan rekap penjualan menggunakan pandas
class DataManager:
    def __init__(self):
        # Path file excel produk dan rekap penjualan
        self.produk_file = os.path.join(os.path.dirname(__file__), "list_produk.xlsx")
        self.recap_file = os.path.join(os.path.dirname(__file__), "recap_penjualan.xlsx")
        # Load data produk dari excel ke dataframe pandas
        self.df_produk = pd.read_excel(self.produk_file, dtype={"Kode": str})

    def cari_produk(self, query):
        # Cari produk berdasarkan kode atau nama (case-insensitive)
        query = query.strip().lower()
        mask = (self.df_produk["Kode"].str.lower() == query) | \
               (self.df_produk["Nama Produk"].str.lower().str.contains(query, na=False))
        hasil = self.df_produk[mask]
        # Kembalikan None jika tidak ditemukan, list dict jika ada
        return None if hasil.empty else hasil.to_dict("records")

    def simpan_rekap(self, tanggal, keranjang):
        # Buat dataframe dari item yang dibeli
        rows = [{"Tanggal": tanggal, "Kode": i["Kode"], "Nama": i["Nama Produk"],
                 "Qty": i["qty"], "Harga Satuan": i["Harga"],
                 "Subtotal": i["Harga"] * i["qty"]} for i in keranjang]
        df_baru = pd.DataFrame(rows)
        # Jika file rekap sudah ada, tambahkan; jika belum, buat baru
        if os.path.exists(self.recap_file):
            df_lama = pd.read_excel(self.recap_file)
            df_gabung = pd.concat([df_lama, df_baru], ignore_index=True)
        else:
            df_gabung = df_baru
        df_gabung.to_excel(self.recap_file, index=False)


# ===== CLASS: KasirApp =====
# Mengelola seluruh tampilan GUI aplikasi kasir
class KasirApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Jack Owi Mart - Kasir")
        self.root.state("zoomed")
        self.root.configure(bg="#F0F4F8")

        self.dm = DataManager()         # Instance DataManager
        self.keranjang = []             # List item belanja
        self.tanggal_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

        self._build_header()            # Bangun header toko
        self._build_main()              # Bangun area utama

    def _build_header(self):
        # Frame header dengan background biru toko
        header = tk.Frame(self.root, bg="#0056B3", pady=10)
        header.pack(fill=tk.X)
        # Label nama toko dengan warna berbeda per kata
        nama_frame = tk.Frame(header, bg="#0056B3")
        nama_frame.pack()
        tk.Label(nama_frame, text="Jack Owi", font=("Arial", 22, "bold"),
                 fg="#FFD700", bg="#0056B3").pack(side=tk.LEFT)
        tk.Label(nama_frame, text=" Mart", font=("Arial", 22, "bold"),
                 fg="#FFFFFF", bg="#0056B3").pack(side=tk.LEFT)
        # Tanggal transaksi di bawah nama toko
        tgl_f = tk.Frame(header, bg="#0056B3")
        tgl_f.pack()
        tk.Label(tgl_f, text="Tanggal: ", font=("Arial", 9),
                 fg="#FFFFFF", bg="#0056B3").pack(side=tk.LEFT)
        tk.Entry(tgl_f, textvariable=self.tanggal_var, font=("Arial", 9),
                 width=12, justify="center").pack(side=tk.LEFT)

    def _build_main(self):
        # Frame utama untuk dua kolom: pencarian + keranjang
        main = tk.Frame(self.root, bg="#F0F4F8")
        main.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        # ==== KOLOM KIRI: Pencarian produk ====
        kiri = tk.Frame(main, bg="#F0F4F8", width=450)
        kiri.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        kiri.pack_propagate(False)

        tk.Label(kiri, text="Cari Produk", font=("Arial", 11, "bold"),
                 bg="#F0F4F8").pack(anchor="w")
        # Input pencarian dan tombol cari
        cari_f = tk.Frame(kiri, bg="#F0F4F8")
        cari_f.pack(fill=tk.X, pady=4)
        self.entry_cari = tk.Entry(cari_f, font=("Arial", 10), width=22)
        self.entry_cari.pack(side=tk.LEFT, ipady=4, padx=(0, 4))
        self.entry_cari.bind("<Return>", lambda e: self._cari())  # Enter untuk cari
        tk.Button(cari_f, text="Cari", font=("Arial", 9, "bold"), bg="#0056B3",
                  fg="white", relief=tk.FLAT, padx=8, command=self._cari).pack(side=tk.LEFT)

        # Label status pencarian (produk habis, dll)
        self.lbl_status = tk.Label(kiri, text="", font=("Arial", 9),
                                   bg="#F0F4F8", fg="red")
        self.lbl_status.pack(anchor="w")

        # Frame hasil pencarian dengan scrollbar
        tk.Label(kiri, text="Hasil Pencarian:", font=("Arial", 9),
                 bg="#F0F4F8", fg="#555").pack(anchor="w", pady=(4, 0))
        hasil_frame = tk.Frame(kiri, bg="#F0F4F8")
        hasil_frame.pack(fill=tk.BOTH, expand=True)
        self.listbox_hasil = tk.Listbox(
            hasil_frame,
            font=("Arial", 10), height=25,
            selectbackground="#0056B3",
            activestyle="none"
        )
        sb = tk.Scrollbar(hasil_frame, command=self.listbox_hasil.yview)
        self.listbox_hasil.config(yscrollcommand=sb.set)
        self.listbox_hasil.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        # Double click untuk tambah ke keranjang
        self.listbox_hasil.bind("<Double-Button-1>", lambda e: self._tambah_dari_hasil())

        tk.Label(kiri, text="Double-click untuk menambah ke keranjang",
                 font=("Arial", 8), bg="#F0F4F8", fg="#888").pack(anchor="w", pady=3)

        # ==== KOLOM KANAN: Keranjang belanja ====
        kanan = tk.Frame(main, bg="#F0F4F8")
        kanan.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tk.Label(kanan, text="Keranjang Belanja", font=("Arial", 11, "bold"),
                 bg="#F0F4F8").pack(anchor="w")

        # Tabel keranjang: No, Nama, Harga, Qty, Subtotal, Hapus
        cols = ("No", "Nama Produk", "Harga", "Qty", "Subtotal", "")
        self.tree = ttk.Treeview(kanan, columns=cols, show="headings", height=12)
        widths = [50, 450, 150, 80, 180, 70]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")
        self.tree.column("Nama Produk", anchor="w")
        # Scrollbar untuk tabel
        vsb = ttk.Scrollbar(kanan, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.LEFT, fill=tk.Y)

        # Klik pada kolom hapus (kolom ke-6) untuk kurangi qty
        self.tree.bind("<ButtonRelease-1>", self._klik_tree)

        # Area total dan tombol bayar
        bawah = tk.Frame(self.root, bg="#F0F4F8")
        bawah.pack(fill=tk.X, padx=15, pady=6)
        self.lbl_total = tk.Label(bawah, text="Total: Rp 0",
                                  font=("Arial", 13, "bold"), bg="#F0F4F8", fg="#0056B3")
        self.lbl_total.pack(side=tk.LEFT)
        tk.Button(bawah, text="  BAYAR  ", font=("Arial", 11, "bold"),
                  bg="#FFD700", fg="#0056B3", relief=tk.FLAT, padx=10,
                  command=self._buka_bayar).pack(side=tk.RIGHT)
        self._tampilkan_semua_produk()
        self._simpan_hasil = []         # Menyimpan list hasil pencarian sementara
    
    def _tampilkan_semua_produk(self):
        self.listbox_hasil.delete(0, tk.END)
        self._simpan_hasil = self.dm.df_produk.to_dict("records")
        for p in self._simpan_hasil:
            self.listbox_hasil.insert(
                tk.END,
                f"[{p['Kode']}] {p['Nama Produk']} - Rp {p['Harga']:,}"
            )
            
    # --- Metode pencarian produk ---
    def _cari(self):
        query = self.entry_cari.get()
        if not query.strip():
            return
        hasil = self.dm.cari_produk(query)
        self.listbox_hasil.delete(0, tk.END)
        # Jika produk tidak ditemukan, tampilkan pesan "Produk habis"
        if not hasil:
            self.lbl_status.config(text="Produk habis / tidak tersedia")
            self._simpan_hasil = []
        else:
            self.lbl_status.config(text="")
            self._simpan_hasil = hasil
            # Tampilkan setiap hasil di listbox
            for p in hasil:
                self.listbox_hasil.insert(tk.END,
                    f"[{p['Kode']}] {p['Nama Produk']} - Rp {p['Harga']:,}")

    # --- Tambah produk dari hasil pencarian ke keranjang ---
    def _tambah_dari_hasil(self):
        sel = self.listbox_hasil.curselection()
        if not sel or not self._simpan_hasil:
            return
        idx = sel[0]
        if idx >= len(self._simpan_hasil):
            return
        produk = self._simpan_hasil[idx]
        # Cek apakah produk sudah ada di keranjang
        for item in self.keranjang:
            if item["Kode"] == produk["Kode"]:
                item["qty"] += 1           # Tambah qty jika sudah ada
                self._render_keranjang()
                return
        # Jika belum ada, tambahkan sebagai item baru
        item_baru = dict(produk)
        item_baru["qty"] = 1
        self.keranjang.append(item_baru)
        self._render_keranjang()

    # --- Render ulang isi tabel keranjang ---
    def _render_keranjang(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        total = 0
        for i, item in enumerate(self.keranjang, 1):
            sub = item["Harga"] * item["qty"]
            total += sub
            # Kolom terakhir adalah tombol kurangi (scroll "-")
            self.tree.insert("", tk.END, iid=str(i-1), values=(
                i, item["Nama Produk"], f"Rp {item['Harga']:,}",
                item["qty"], f"Rp {sub:,}", "[-]"))
        self.lbl_total.config(text=f"Total: Rp {total:,}")

    # --- Handle klik di tabel: tombol [-] untuk kurangi qty ---
    def _klik_tree(self, event):
        region = self.tree.identify_region(event.x, event.y)
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        # Hanya proses klik di kolom ke-6 (kolom hapus)
        if region == "cell" and col == "#6" and row:
            idx = int(row)
            if idx < len(self.keranjang):
                if self.keranjang[idx]["qty"] > 1:
                    self.keranjang[idx]["qty"] -= 1   # Kurangi 1 qty
                else:
                    self.keranjang.pop(idx)            # Hapus jika qty = 0
                self._render_keranjang()

    # --- Buka window pembayaran ---
    def _buka_bayar(self):
        if not self.keranjang:
            messagebox.showwarning("Keranjang Kosong", "Tambahkan produk terlebih dahulu.")
            return
        # Buat window baru untuk konfirmasi pembayaran
        win = tk.Toplevel(self.root)
        win.title("Pembayaran")
        win.geometry("480x460")
        win.resizable(False, False)
        win.configure(bg="#F0F4F8")
        win.grab_set()

        # Header window pembayaran
        tk.Label(win, text="Konfirmasi Pembayaran", font=("Arial", 13, "bold"),
                 bg="#0056B3", fg="white", pady=8).pack(fill=tk.X)

        # Tabel ringkasan: No, Kode, Nama Produk, Harga
        cols = ("No", "Kode", "Nama Produk", "Harga")
        tree2 = ttk.Treeview(win, columns=cols, show="headings", height=10)
        for col, w in zip(cols, [35, 75, 200, 100]):
            tree2.heading(col, text=col)
            tree2.column(col, width=w, anchor="center")
        tree2.column("Nama Produk", anchor="w")

        total = 0
        for i, item in enumerate(self.keranjang, 1):
            sub = item["Harga"] * item["qty"]
            total += sub
            # Tampilkan setiap baris item
            tree2.insert("", tk.END, values=(
                i, item["Kode"], f"{item['Nama Produk']} (x{item['qty']})",
                f"Rp {sub:,}"))
        tree2.pack(padx=15, pady=10, fill=tk.X)

        tk.Label(win, text=f"Total: Rp {total:,}", font=("Arial", 12, "bold"),
                 bg="#F0F4F8", fg="#0056B3").pack(pady=4)

        # Tombol bayar sekarang
        tk.Button(win, text="Bayar Sekarang", font=("Arial", 11, "bold"),
                  bg="#28A745", fg="white", relief=tk.FLAT, padx=15, pady=6,
                  command=lambda: self._proses_bayar(win, total)).pack(pady=8)

    # --- Proses pembayaran dan tampilkan konfirmasi ---
    def _proses_bayar(self, win, total):
        win.destroy()
        # Simpan rekap ke excel menggunakan DataManager
        self.dm.simpan_rekap(self.tanggal_var.get(), self.keranjang)

        # Window sukses pembayaran
        swin = tk.Toplevel(self.root)
        swin.title("Pembayaran Berhasil")
        swin.geometry("350x200")
        swin.resizable(False, False)
        swin.configure(bg="#F0F4F8")
        swin.grab_set()

        # Ikon centang dan pesan berhasil
        tk.Label(swin, text="✅", font=("Arial", 38), bg="#F0F4F8").pack(pady=10)
        tk.Label(swin, text="Pembayaran Berhasil!", font=("Arial", 13, "bold"),
                 bg="#F0F4F8", fg="#28A745").pack()
        tk.Label(swin, text=f"Total: Rp {total:,}", font=("Arial", 10),
                 bg="#F0F4F8").pack()
        # Tombol cetak struk
        keranjang_copy = list(self.keranjang)
        tk.Button(swin, text="Cetak Struk", font=("Arial", 10, "bold"),
                  bg="#0056B3", fg="white", relief=tk.FLAT, padx=10,
                  command=lambda: [swin.destroy(),
                                   self._cetak_struk(keranjang_copy, total)]).pack(pady=10)
        # Reset keranjang setelah bayar
        self.keranjang = []
        self._render_keranjang()

    # --- Tampilkan struk belanja seperti template receipt ---
    def _cetak_struk(self, keranjang, total):
        swin = tk.Toplevel(self.root)
        swin.title("Struk Belanja")
        swin.geometry("360x520")
        swin.resizable(False, False)
        swin.configure(bg="white")
        swin.grab_set()

        # Frame struk dengan padding
        f = tk.Frame(swin, bg="white", padx=24, pady=16)
        f.pack(fill=tk.BOTH, expand=True)

        # Header struk - nama toko
        tk.Label(f, text="Jack Owi Mart", font=("Arial", 14, "bold"),
                 bg="white", anchor="center").pack(fill=tk.X)
        tk.Label(f, text="Jl. Sudirman No. 23-10\nTelp: 021-1122334",
                 font=("Arial", 8), bg="white", anchor="center").pack(fill=tk.X)
        tk.Label(f, text="CASH RECEIPT", font=("Arial", 10, "bold"),
                 bg="white").pack(pady=(6, 2))
        tk.Label(f, text=f"Tanggal: {self.tanggal_var.get()}",
                 font=("Arial", 8), bg="white").pack(anchor="w")

        # Tabel item struk
        tbl = tk.Frame(f, bg="white")
        tbl.pack(fill=tk.X, pady=6)
        # Header kolom tabel struk
        tk.Label(tbl, text="Nama Produk", font=("Arial", 8, "bold"),
                 bg="white", width=22, anchor="w").grid(row=0, column=0, sticky="w")
        tk.Label(tbl, text="Harga", font=("Arial", 8, "bold"),
                 bg="white", width=12, anchor="e").grid(row=0, column=1, sticky="e")

        # Isi baris item per produk
        for i, item in enumerate(keranjang, 1):
            sub = item["Harga"] * item["qty"]
            label_nama = f"{item['Nama Produk']} x{item['qty']}"
            tk.Label(tbl, text=label_nama, font=("Arial", 8),
                     bg="white", anchor="w").grid(row=i, column=0, sticky="w")
            tk.Label(tbl, text=f"Rp {sub:,}", font=("Arial", 8),
                     bg="white", anchor="e").grid(row=i, column=1, sticky="e")

        # Garis pemisah total
        tk.Frame(f, bg="#CCCCCC", height=1).pack(fill=tk.X, pady=4)

        # Tampilkan total, kas, dan kembalian (simulasi kas = total dibulatkan ke 1000)
        kas = ((total // 1000) + 1) * 1000
        kembalian = kas - total
        total_f = tk.Frame(f, bg="white")
        total_f.pack(fill=tk.X)
        for label, nilai in [("Total", total), ("Cash", kas), ("Kembalian", kembalian)]:
            row = tk.Frame(total_f, bg="white")
            row.pack(fill=tk.X)
            bold = "bold" if label == "Total" else "normal"
            tk.Label(row, text=label, font=("Arial", 9, bold),
                     bg="white", anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text=f"Rp {nilai:,}", font=("Arial", 9, bold),
                     bg="white", anchor="e").pack(side=tk.RIGHT)

        # Footer struk
        tk.Frame(f, bg="#CCCCCC", height=1).pack(fill=tk.X, pady=6)
        tk.Label(f, text="TERIMA KASIH!", font=("Arial", 10, "bold"),
                 bg="white").pack()
        tk.Label(f, text="Selamat berbelanja kembali",
                 font=("Arial", 8), bg="white", fg="#888").pack()

# ===== ENTRY POINT =====
if __name__ == "__main__":
    root = tk.Tk()
    app = KasirApp(root)
    root.mainloop()