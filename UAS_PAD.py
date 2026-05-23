import tkinter as tk
from tkinter import ttk, messagebox

# ================= DATA PRODUK =================

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

# ================= DATA =================

keranjang = {}

# ================= FUNCTION =================

def tampilkan_info(event=None):

    kode = entry_kode.get()

    if kode in produk:

        label_barang.config(
            text=f"Nama Barang : {produk[kode]}"
        )

        label_harga.config(
            text=f"Harga : Rp{harga[kode]}"
        )

    else:

        label_barang.config(
            text="Nama Barang : -"
        )

        label_harga.config(
            text="Harga : -"
        )


def tambah_barang():

    kode = entry_kode.get()

    if kode not in produk:
        messagebox.showerror(
            "Error",
            "Kode barang tidak ditemukan"
        )
        return

    try:
        jumlah = int(spin_jumlah.get())

    except:
        messagebox.showerror(
            "Error",
            "Jumlah tidak valid"
        )
        return

    keranjang[kode] = keranjang.get(kode, 0) + jumlah

    tampilkan_keranjang()

    entry_kode.delete(0, tk.END)

    spin_jumlah.delete(0, tk.END)
    spin_jumlah.insert(0, "1")

    label_barang.config(text="Nama Barang : -")
    label_harga.config(text="Harga : -")


def tampilkan_keranjang():

    tree_keranjang.delete(*tree_keranjang.get_children())

    total = 0

    for kode, jumlah in keranjang.items():

        subtotal = harga[kode] * jumlah

        total += subtotal

        tree_keranjang.insert(
            "",
            tk.END,
            values=(
                produk[kode],
                jumlah,
                f"Rp{subtotal}"
            )
        )

    label_total.config(
        text=f"TOTAL : Rp{total}"
    )


def selesai_input():

    if len(keranjang) == 0:
        messagebox.showerror(
            "Error",
            "Keranjang masih kosong"
        )
        return

    total = 0

    for kode, jumlah in keranjang.items():
        total += harga[kode] * jumlah

    diskon = 0

    if total > 100000:
        diskon = total * 0.10

    elif total > 50000:
        diskon = total * 0.05

    elif total > 25000:
        diskon = total * 0.02

    total_bayar = int(total - diskon)

    messagebox.showinfo(
        "Input Selesai",
        f"Total Belanja : Rp{total}\n"
        f"Diskon : Rp{int(diskon)}\n"
        f"Total Bayar : Rp{total_bayar}\n\n"
        f"Silahkan lanjut ke pembayaran"
    )

    entry_kode.config(state="disabled")
    spin_jumlah.config(state="disabled")
    btn_tambah.config(state="disabled")


def transaksi_baru():

    keranjang.clear()

    tampilkan_keranjang()

    entry_kode.config(state="normal")
    spin_jumlah.config(state="normal")
    btn_tambah.config(state="normal")

    entry_kode.delete(0, tk.END)

    spin_jumlah.delete(0, tk.END)
    spin_jumlah.insert(0, "1")

    metode_pembayaran.set("")

    entry_uang.delete(0, tk.END)

    label_barang.config(text="Nama Barang : -")
    label_harga.config(text="Harga : -")


def bayar():

    if len(keranjang) == 0:
        messagebox.showerror(
            "Error",
            "Keranjang kosong"
        )
        return

    total = 0

    for kode, jumlah in keranjang.items():

        total += harga[kode] * jumlah

    diskon = 0

    if total > 100000:
        diskon = total * 0.10

    elif total > 50000:
        diskon = total * 0.05

    elif total > 25000:
        diskon = total * 0.02

    total_bayar = int(total - diskon)

    metode = metode_pembayaran.get()

    if metode == "":
        messagebox.showerror(
            "Error",
            "Pilih metode pembayaran"
        )
        return

    if metode == "Cash":

        uang = entry_uang.get()

        if uang == "":
            messagebox.showerror(
                "Error",
                "Masukkan uang pembayaran"
            )
            return

        try:
            uang = int(uang)

        except:
            messagebox.showerror(
                "Error",
                "Input uang harus angka"
            )
            return

        if uang < total_bayar:
            messagebox.showerror(
                "Error",
                "Uang tidak cukup"
            )
            return

        kembalian = uang - total_bayar

        messagebox.showinfo(
            "Pembayaran Berhasil",
            f"Metode : Cash\n"
            f"Total : Rp{total_bayar}\n"
            f"Uang Bayar : Rp{uang}\n"
            f"Kembalian : Rp{kembalian}"
        )

        transaksi_baru()

    elif metode == "QRIS":

        messagebox.showinfo(
            "Pembayaran QRIS",
            f"Silahkan scan QR\n"
            f"Total Bayar : Rp{total_bayar}"
        )

        transaksi_baru()

    elif metode == "Debit":

        messagebox.showinfo(
            "Pembayaran Debit",
            f"Silahkan tempel kartu\n"
            f"Total Bayar : Rp{total_bayar}"
        )

        transaksi_baru()

# ================= WINDOW =================

root = tk.Tk()

root.title("APLIKASI KASIR")

root.state("zoomed")

root.config(bg="white")

# ================= TITLE =================

judul = tk.Label(
    root,
    text="APLIKASI KASIR",
    font=("Arial", 28, "bold"),
    bg="white",
    fg="black"
)

judul.pack(pady=20)

# ================= FRAME ATAS =================

frame_atas = tk.Frame(
    root,
    bg="white"
)

frame_atas.pack(fill="x", pady=10)

# ================= INPUT PRODUK =================

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

# Kode Barang
tk.Label(
    frame_input,
    text="Kode Barang",
    font=("Arial", 12),
    bg="white"
).grid(row=0, column=0, pady=10, sticky="w")

entry_kode = tk.Entry(
    frame_input,
    font=("Arial", 12),
    width=25
)

entry_kode.grid(row=0, column=1, padx=10)

entry_kode.bind("<KeyRelease>", tampilkan_info)

# Jumlah Barang
tk.Label(
    frame_input,
    text="Jumlah",
    font=("Arial", 12),
    bg="white"
).grid(row=1, column=0, pady=10, sticky="w")

spin_jumlah = tk.Spinbox(
    frame_input,
    from_=1,
    to=100,
    width=10,
    font=("Arial", 12)
)

spin_jumlah.grid(row=1, column=1, sticky="w")

# Nama Barang
label_barang = tk.Label(
    frame_input,
    text="Nama Barang : -",
    font=("Arial", 12, "bold"),
    bg="white",
    fg="green"
)

label_barang.grid(
    row=2,
    column=0,
    columnspan=2,
    pady=10
)

# Harga Barang
label_harga = tk.Label(
    frame_input,
    text="Harga : -",
    font=("Arial", 12),
    bg="white"
)

label_harga.grid(
    row=3,
    column=0,
    columnspan=2
)

# Tombol Tambah
btn_tambah = tk.Button(
    frame_input,
    text="Tambah Produk",
    font=("Arial", 12, "bold"),
    bg="green",
    fg="white",
    width=20,
    command=tambah_barang
)

btn_tambah.grid(
    row=4,
    column=0,
    columnspan=2,
    pady=10
)

# Tombol Selesai Input
btn_selesai = tk.Button(
    frame_input,
    text="Selesai Input",
    font=("Arial", 12, "bold"),
    bg="orange",
    fg="white",
    width=20,
    command=selesai_input
)

btn_selesai.grid(
    row=5,
    column=0,
    columnspan=2,
    pady=5
)

# ================= TABEL PRODUK =================

frame_produk = tk.LabelFrame(
    frame_atas,
    text="DAFTAR PRODUK",
    font=("Arial", 14, "bold"),
    padx=10,
    pady=10,
    bg="white"
)

frame_produk.pack(
    side="left",
    padx=20
)

kolom_produk = ("Kode", "Nama Produk", "Harga")

tree_produk = ttk.Treeview(
    frame_produk,
    columns=kolom_produk,
    show="headings",
    height=12
)

for col in kolom_produk:
    tree_produk.heading(col, text=col)

tree_produk.column("Kode", width=100)
tree_produk.column("Nama Produk", width=350)
tree_produk.column("Harga", width=150)

for kode in produk:

    tree_produk.insert(
        "",
        tk.END,
        values=(
            kode,
            produk[kode],
            f"Rp{harga[kode]}"
        )
    )

tree_produk.pack()

# ================= KERANJANG =================

frame_bawah = tk.LabelFrame(
    root,
    text="KERANJANG BELANJA",
    font=("Arial", 14, "bold"),
    padx=20,
    pady=20,
    bg="white"
)

frame_bawah.pack(
    fill="both",
    expand=True,
    padx=20,
    pady=20
)

kolom_keranjang = (
    "Nama Barang",
    "Jumlah",
    "Subtotal"
)

tree_keranjang = ttk.Treeview(
    frame_bawah,
    columns=kolom_keranjang,
    show="headings",
    height=12
)

for col in kolom_keranjang:
    tree_keranjang.heading(col, text=col)

tree_keranjang.column("Nama Barang", width=500)
tree_keranjang.column("Jumlah", width=150)
tree_keranjang.column("Subtotal", width=250)

tree_keranjang.pack(
    fill="both",
    expand=True
)

# ================= TOTAL =================

label_total = tk.Label(
    root,
    text="TOTAL : Rp0",
    font=("Arial", 20, "bold"),
    bg="white",
    fg="blue"
)

label_total.pack(pady=10)

# ================= METODE PEMBAYARAN =================

frame_pembayaran = tk.LabelFrame(
    root,
    text="METODE PEMBAYARAN",
    font=("Arial", 14, "bold"),
    padx=20,
    pady=20,
    bg="white"
)

frame_pembayaran.pack(pady=10)

# Combobox metode
tk.Label(
    frame_pembayaran,
    text="Metode",
    font=("Arial", 12),
    bg="white"
).grid(row=0, column=0, padx=10)

metode_pembayaran = ttk.Combobox(
    frame_pembayaran,
    values=["Cash", "QRIS", "Debit"],
    state="readonly",
    width=20,
    font=("Arial", 12)
)

metode_pembayaran.grid(row=0, column=1, padx=10)

# Input uang cash
tk.Label(
    frame_pembayaran,
    text="Uang Cash",
    font=("Arial", 12),
    bg="white"
).grid(row=1, column=0, padx=10, pady=10)

entry_uang = tk.Entry(
    frame_pembayaran,
    font=("Arial", 12),
    width=23
)

entry_uang.grid(row=1, column=1)

# ================= BUTTON BAYAR =================

btn_bayar = tk.Button(
    root,
    text="BAYAR",
    font=("Arial", 14, "bold"),
    bg="blue",
    fg="white",
    width=25,
    height=2,
    command=bayar
)

btn_bayar.pack(pady=20)

root.mainloop()