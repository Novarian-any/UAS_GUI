# Import library GUI tkinter
import tkinter as tk

# ttk digunakan untuk tabel dan combobox
# messagebox digunakan untuk popup pesan
from tkinter import ttk, messagebox

# Dictionary nama produk berdasarkan kode
produk = {
    "001":"Indomie Goreng",
    "002":"Indomie Soto Lamongan",
    "003":"Indomie Kari Ayam",
    "004":"Indomie Soto",
    "005":"Indomie Ayam Bawang",

    "021":"Lifebuoy Body Wash",
    "022":"Lifebuoy Lemon Fresh",

    "041":"Rinso Anti Noda",
    "042":"Rinso Molto Strawberry",

    "061":"SilverQueen Almond",
    "062":"SilverQueen Cashew",

    "081":"Aqua 600ml",
    "082":"Le Minerale 600ml"
}

# Dictionary harga produk
harga = {
    "001":3500,
    "002":3500,
    "003":3500,
    "004":3500,
    "005":3500,

    "021":12000,
    "022":12000,

    "041":23000,
    "042":24000,

    "061":15000,
    "062":15000,

    "081":3500,
    "082":3300
}

# Menyimpan data barang yang dipilih user
keranjang = {}

# Menampilkan nama dan harga barang otomatis
def tampilkan_info(event=None):

    # Mengambil kode dari input
    kode = entry_kode.get()

    # Jika kode ditemukan
    if kode in produk:

        # Tampilkan nama barang
        label_barang.config(
            text=f"Nama Barang : {produk[kode]}"
        )

        # Tampilkan harga barang
        label_harga.config(
            text=f"Harga : Rp{harga[kode]}"
        )

    # Jika kode tidak ditemukan
    else:

        # Reset label
        label_barang.config(
            text="Nama Barang : -"
        )

        label_harga.config(
            text="Harga : -"
        )

# Menambahkan barang ke keranjang
def tambah_barang():

    # Mengambil kode barang
    kode = entry_kode.get()

    # Validasi kode barang
    if kode not in produk:
        messagebox.showerror(
            "Error",
            "Kode barang tidak ditemukan"
        )
        return

    # Validasi jumlah barang
    try:
        jumlah = int(spin_jumlah.get())

    except:
        messagebox.showerror(
            "Error",
            "Jumlah tidak valid"
        )
        return

    # Menambahkan barang ke keranjang
    keranjang[kode] = keranjang.get(kode, 0) + jumlah

    # Refresh tabel keranjang
    tampilkan_keranjang()

    # Menghapus input kode
    entry_kode.delete(0, tk.END)

    # Reset jumlah barang
    spin_jumlah.delete(0, tk.END)
    spin_jumlah.insert(0, "1")

    # Reset informasi barang
    label_barang.config(text="Nama Barang : -")
    label_harga.config(text="Harga : -")

# Menampilkan isi keranjang ke tabel
def tampilkan_keranjang():

    # Menghapus isi tabel sebelumnya
    tree_keranjang.delete(*tree_keranjang.get_children())

    total = 0

    # Loop isi keranjang
    for kode, jumlah in keranjang.items():

        # Hitung subtotal
        subtotal = harga[kode] * jumlah

        # Tambah ke total
        total += subtotal

        # Tampilkan ke tabel
        tree_keranjang.insert(
            "",
            tk.END,
            values=(
                produk[kode],
                jumlah,
                f"Rp{subtotal}"
            )
        )

    # Tampilkan total belanja
    label_total.config(
        text=f"TOTAL : Rp{total}"
    )

# Mengunci input sebelum pembayaran
def selesai_input():

    # Validasi keranjang kosong
    if len(keranjang) == 0:
        messagebox.showerror(
            "Error",
            "Keranjang masih kosong"
        )
        return

    total = 0

    # Menghitung total harga
    for kode, jumlah in keranjang.items():
        total += harga[kode] * jumlah

    diskon = 0

    # Menghitung diskon
    if total > 100000:
        diskon = total * 0.10

    elif total > 50000:
        diskon = total * 0.05

    elif total > 25000:
        diskon = total * 0.02

    # Total setelah diskon
    total_bayar = int(total - diskon)

    # Menampilkan informasi pembayaran
    messagebox.showinfo(
        "Input Selesai",
        f"Total Belanja : Rp{total}\n"
        f"Diskon : Rp{int(diskon)}\n"
        f"Total Bayar : Rp{total_bayar}\n\n"
        f"Silahkan lanjut ke pembayaran"
    )

    # Menonaktifkan input
    entry_kode.config(state="disabled")
    spin_jumlah.config(state="disabled")
    btn_tambah.config(state="disabled")

# Mengembalikan program ke kondisi awal
def transaksi_baru():

    # Menghapus isi keranjang
    keranjang.clear()

    # Refresh tabel
    tampilkan_keranjang()

    # Mengaktifkan input kembali
    entry_kode.config(state="normal")
    spin_jumlah.config(state="normal")
    btn_tambah.config(state="normal")

    # Menghapus isi input
    entry_kode.delete(0, tk.END)

    # Reset jumlah barang
    spin_jumlah.delete(0, tk.END)
    spin_jumlah.insert(0, "1")

    # Reset metode pembayaran
    metode_pembayaran.set("")

    # Menghapus input uang
    entry_uang.delete(0, tk.END)

    # Reset label barang
    label_barang.config(text="Nama Barang : -")
    label_harga.config(text="Harga : -")

# Memproses pembayaran
def bayar():

    # Validasi keranjang kosong
    if len(keranjang) == 0:
        messagebox.showerror(
            "Error",
            "Keranjang kosong"
        )
        return

    total = 0

    # Menghitung total harga
    for kode, jumlah in keranjang.items():
        total += harga[kode] * jumlah

    diskon = 0

    # Menghitung diskon
    if total > 100000:
        diskon = total * 0.10

    elif total > 50000:
        diskon = total * 0.05

    elif total > 25000:
        diskon = total * 0.02

    # Total akhir pembayaran
    total_bayar = int(total - diskon)

    # Mengambil metode pembayaran
    metode = metode_pembayaran.get()

    # Validasi metode pembayaran
    if metode == "":
        messagebox.showerror(
            "Error",
            "Pilih metode pembayaran"
        )
        return

    # Pembayaran cash
    if metode == "Cash":

        # Mengambil input uang
        uang = entry_uang.get()

        # Validasi input kosong
        if uang == "":
            messagebox.showerror(
                "Error",
                "Masukkan uang pembayaran"
            )
            return

        # Konversi ke integer
        try:
            uang = int(uang)

        except:
            messagebox.showerror(
                "Error",
                "Input uang harus angka"
            )
            return

        # Validasi uang kurang
        if uang < total_bayar:
            messagebox.showerror(
                "Error",
                "Uang tidak cukup"
            )
            return

        # Hitung kembalian
        kembalian = uang - total_bayar

        # Popup pembayaran berhasil
        messagebox.showinfo(
            "Pembayaran Berhasil",
            f"Metode : Cash\n"
            f"Total : Rp{total_bayar}\n"
            f"Uang Bayar : Rp{uang}\n"
            f"Kembalian : Rp{kembalian}"
        )

        # Reset transaksi
        transaksi_baru()

    # Pembayaran QRIS
    elif metode == "QRIS":

        messagebox.showinfo(
            "Pembayaran QRIS",
            f"Silahkan scan QR\n"
            f"Total Bayar : Rp{total_bayar}"
        )

        transaksi_baru()

    # Pembayaran Debit
    elif metode == "Debit":

        messagebox.showinfo(
            "Pembayaran Debit",
            f"Silahkan tempel kartu\n"
            f"Total Bayar : Rp{total_bayar}"
        )

        transaksi_baru()

# Membuat window utama
root = tk.Tk()

# Judul aplikasi
root.title("APLIKASI KASIR")

# Membuat fullscreen
root.state("zoomed")

# Background putih
root.config(bg="white")

# Label judul aplikasi
judul = tk.Label(
    root,
    text="APLIKASI KASIR",
    font=("Arial", 28, "bold"),
    bg="white",
    fg="black"
)

judul.pack(pady=20)

# Frame bagian atas
frame_atas = tk.Frame(
    root,
    bg="white"
)

frame_atas.pack(fill="x", pady=10)

# Frame input barang
frame_input = tk.LabelFrame(
    frame_atas,
    text="INPUT PRODUK",
    font=("Arial", 14, "bold"),
    padx=20,
    pady=20,
    bg="white"
)

frame_input.pack(
    side="left",
    padx=20
)

# Label kode barang
tk.Label(
    frame_input,
    text="Kode Barang",
    font=("Arial", 12),
    bg="white"
).grid(row=0, column=0, pady=10, sticky="w")

# Input kode barang
entry_kode = tk.Entry(
    frame_input,
    font=("Arial", 12),
    width=25
)

entry_kode.grid(row=0, column=1, padx=10)

# Event keyboard
entry_kode.bind("<KeyRelease>", tampilkan_info)
