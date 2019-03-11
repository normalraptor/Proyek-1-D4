# Fungsi Hlookup

Fungsi atau Rumus Excel HLOOKUP adalah salah satu fungsi ms. Excel yang digunakan untuk mencari data pada kolom pertama sebuah tabel data, kemudian mengambil nilai dari sel mana pun di baris yang sama pada tabel data tersebut.

> =HLOOKUP\(lookup\_value, table\_array, row\_index\_num, \[range\_lookup\]\)

**Deskripsi**

* _lookup\_value_,  Nilai yang dicari pada baris pertama tabel. lookup\_value bisa berupa nilai, referensi, atau string teks.
* _table\_array_,  Tabel data dimana lookup\_value akan di temukan pada baris pertamanya. Gunakan referensi ke sebuah range, tabel, nama range atau nama tabel.
*  _row\_index\_num_, Nomor baris dalam table\_array yang akan menghasilkan nilai yang cocok. 
* _range\_lookup_, berisi TRUE atau FALSE. Nilai logika ini akan menentukan apakah kita ingin Hlookup mencari kecocokan yang sama persis atau kecocokan yang mendekati. Jika berisi TRUE atau dikosongkan \(tidak diisi\) maka Hlookup menghasilkan kecocokan yang mendekati. Jika FALSE, HLOOKUP akan menemukan kecocokan persis. Jika tidak ditemukan, akan menghasilkan nilai kesalahan \#N/A.

lookup\_value, table\_array, row\_index\_num merupakan komponen yang wajib ada saat menggunakan rumus HLOOKUP, sedangkan range\_lookup bersifat opsional.

**Contoh Penggunaan**

![](../.gitbook/assets/contohexcel3.png)

Pada gambar tersebut dicari sebuah nilai yang berada pada kolom C1 dan terletak pada baris ketiga

