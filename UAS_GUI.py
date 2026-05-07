import tkinter as tk
from tkinter import ttk, messagebox

# ================= DATA PRODUK =================

produk = {
    "001":"Indomie Goreng","002":"Indomie Soto Lamongan","003":"Indomie Kari Ayam","004":"Indomie Soto",
    "005":"Indomie Ayam Bawang","006":"Indomie Rendang","007":"Indomie Iga Penyet","008":"Indomie Aceh",
    "009":"Indomie Sambal Matah","010":"Indomie Cabe Ijo",

    "021":"Lifebuoy Body Wash Total Protect 100ml",
    "022":"Lifebuoy Lemon Fresh 100ml",

    "041":"Rinso Anti Noda Classic 800g",
    "042":"Rinso Molto Korean Strawberry 800g",

    "061":"SilverQueen Almond 58g",
    "062":"SilverQueen Cashew 58g",

    "081":"Aqua Air Mineral 600ml",
    "082":"Le Minerale 600ml"
}

harga = {
    "001":3500,
    "002":3500,
    "003":3500,
    "004":3500,
    "005":3500,
    "006":3500,
    "007":3500,
    "008":3500,
    "009":3500,
    "010":3500,

    "021":12000,
    "022":12000,

    "041":23000,
    "042":24000,

    "061":15000,
    "062":15000,

    "081":3500,
    "082":3300
}

keranjang = {}

# ================= FUNCTION =================

def tambah_ke_keranjang():
    kode = entry_kode.get()
    jumlah = entry_jumlah.get()

    if kode not in produk:
        messagebox.showerror("Error", "Kode tidak ditemukan")
        return

    if not jumlah.isdigit():
        messagebox.showerror("Error", "Jumlah harus angka")
        return

    jumlah = int(jumlah)

    keranjang[kode] = keranjang.get(kode, 0) + jumlah

    tampilkan_keranjang()

    entry_kode.delete(0, tk.END)
    entry_jumlah.delete(0, tk.END)


def tampilkan_keranjang():
    listbox_keranjang.delete(0, tk.END)

    total = 0

    for kode, jumlah in keranjang.items():
        subtotal = harga[kode] * jumlah
        total += subtotal

        teks = f"{produk[kode]} x{jumlah} = Rp{subtotal}"
        listbox_keranjang.insert(tk.END, teks)

    label_total.config(text=f"Total: Rp{total}")


def bayar():
    total = 0

    for kode, jumlah in keranjang.items():
        total += harga[kode] * jumlah

    if total > 100000:
        total *= 0.9
    elif total > 50000:
        total *= 0.95
    elif total > 25000:
        total *= 0.98

    messagebox.showinfo("Pembayaran", f"Total Bayar = Rp{int(total)}")


# ================= GUI =================

root = tk.Tk()
root.title("Aplikasi Kasir")
root.geometry("700x500")
root.config(bg="white")

judul = tk.Label(
    root,
    text="APLIKASI KASIR",
    font=("Arial", 20, "bold"),
    bg="white"
)
judul.pack(pady=10)

# ================= FRAME INPUT =================

frame_input = tk.Frame(root, bg="white")
frame_input.pack(pady=10)

label_kode = tk.Label(frame_input, text="Kode Barang", bg="white")
label_kode.grid(row=0, column=0, padx=5, pady=5)

entry_kode = tk.Entry(frame_input)
entry_kode.grid(row=0, column=1, padx=5)

label_jumlah = tk.Label(frame_input, text="Jumlah", bg="white")
label_jumlah.grid(row=1, column=0, padx=5, pady=5)

entry_jumlah = tk.Entry(frame_input)
entry_jumlah.grid(row=1, column=1, padx=5)

btn_tambah = tk.Button(
    frame_input,
    text="Tambah",
    command=tambah_ke_keranjang,
    bg="green",
    fg="white"
)
btn_tambah.grid(row=2, column=0, columnspan=2, pady=10)

# ================= TABEL PRODUK =================

frame_produk = tk.Frame(root)
frame_produk.pack(pady=10)

kolom = ("Kode", "Produk", "Harga")

tree = ttk.Treeview(frame_produk, columns=kolom, show="headings", height=10)

for col in kolom:
    tree.heading(col, text=col)


tree.column("Kode", width=80)
tree.column("Produk", width=350)
tree.column("Harga", width=120)

for kode in produk:
    tree.insert("", tk.END, values=(kode, produk[kode], harga[kode]))


tree.pack()

# ================= KERANJANG =================

label_keranjang = tk.Label(root, text="Keranjang", font=("Arial", 14, "bold"))
label_keranjang.pack()

listbox_keranjang = tk.Listbox(root, width=70, height=8)
listbox_keranjang.pack(pady=5)

label_total = tk.Label(root, text="Total: Rp0", font=("Arial", 14, "bold"))
label_total.pack(pady=10)

btn_bayar = tk.Button(
    root,
    text="Bayar",
    command=bayar,
    bg="blue",
    fg="white",
    width=20
)
btn_bayar.pack(pady=10)

root.mainloop()

