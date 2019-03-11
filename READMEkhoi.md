---
description: >-
  Fungsi-fungsi berikut merupakan fungsi yang paling sering digunakan dalam
  Microsoft Excel
---

# Fungsi dalam Microsoft Excel

## Fungsi ABS

Fungsi ABS adalah fungsi yang akan memberi nilai absolut pada suatu angka

> =ABS\(number\)

**Deskripsi**  
_Number_, berisi sel angka yang akan dimutlak

**Contoh Penggunaan**

![](.gitbook/assets/image%20%281%29.png)

Dapat kita amati bahwa nilai pada kolom A setelah diberi fungsi abs akan menghasilkan nilai mutlak

## Fungsi Average

Fungsi average akan menghitung rata-rata dari sejumlah sel

> =AVERAGE\(number1;number2; ...\)

**Deskripsi**  
_Number1-...-Number\(n\)_, berisi sel yang akan dihitung rata-ratanya.

**Contoh Penggunaan**

![](.gitbook/assets/image%20%287%29.png)

Dapat kita amati fungsi average pada kolom b5, fungsi ini akan menghitung nilai rata-rata dari sel B2, B3, B4 yaitu rata-rata nya adalah 3,33..

## Fungsi Lower

Mengonversi teks menjadi huruf kecil

> =LOWER\(text\)

**Deskripsi**  
text, berisi alamat sel yang akan dikenakan fungsi lower.

**Contoh Penggunaan**

![](.gitbook/assets/image%20%286%29.png)

Dapat kita amati bahwa fungsi lower akan mengubah text pada kolom A menjadi huruf kecil.

## Fungsi Upper

Mengonversi teks menjadi huruf besar

> =UPPER\(text\)

**Deskripsi**  
text, berisi alamat sel yang akan dikenakan fungsi upper.

**Contoh Penggunaan**

![](.gitbook/assets/image.png)

Dapat kita amati bahwa fungsi upper akan mengubah text pada kolom A menjadi huruf kapital.

## Fungsi Proper

Fungsi Proper adalah fungsi untuk mengubah setiap huruf pertama menggunakan huruf besar dan sisanya menggunakan huruf kecil atau dengan istilah profesional text.

> =PROPER\(text\)

**Deskripsi**  
text, berisi alamat sel yang akan dikenakan fungsi clean.

**Contoh Penggunaan**

![](.gitbook/assets/image%20%283%29.png)

Dapat kita amati pada tiap baris di kolom nama. Fungsi proper pada kolom B2 akan merubah huruf awal dari setiap kata menjadi huruf besar.

## Fungsi Clean

Fungsi dari rumus CLEAN pada Microsoft Excel adalah untuk menghapus semua karakter yang tak dapat dicetak dari teks. Jika biasanya anda mengimport teks dari aplikasi lain ke dalam lembar kerja Microsoft Excel dan tidak dapat tercetak pada Sistem Operasi anda, anda dapat menggunakan fungsi rumus CLEAN ini pada lembar kerja Microsoft Excel anda.

> =CLEAN\(text\)

**Deskripsi**  
text, berisi alamat sel yang akan dikenakan fungsi clean.

**Contoh Penggunaan**

![](.gitbook/assets/image%20%285%29.png)

Dapat kita amati pada kolom A2 terdapat suatu karakter yang tidak sempurna. Maka hasil yang akan muncul pada kolom A2 adalah Rumus CLEAN. Penggunaan rumus CLEAN ini akan bekerja untuk menghilangkan karakter yang tidak sempurna pada komputer anda, anda tinggal menuliskan =CLEAN\(A2\) misalnya pada kolom B2. Maka hasil yang akan keluar pada kolom B2 adalah Rumus CLEAN.

## Fungsi Trim

Fungsi trim merupakan sebuah fungsi yang digunakan untuk menghapus spasi pada teks kecuali spasi ruang antar kata-kata yang mungkin memiliki jarak yang tidak teratur pada ms.excel.

Dengan menggunakan fungsi trim ini, akan dihapus semua kelebihan spasi yang terdapat pada suatu strings teks tanpa harus memperbaiki satu-persatu.

> =TRIM\(text\)

**Deskripsi**  
text, berisi alamat sel yang akan dikenakan fungsi trim.

**Contoh Penggunaan**

![](.gitbook/assets/image%20%282%29.png)

Pada contoh penggunaan fungsi trim tersebut dapat kita amati pada tiap-tiap baris di kolom nama. Text yang terdapat di kolom nama tersebut memiliki _double space_ dan saat setelah menggunakan fungsi trim,  spasi lebih pada suatu text akan terhapus. 

## Fungsi IF

Fungsi ini digunakan untuk mengembalikan suatu nilai jika kondisi bernilai true dan nilai lain jika kondisi bernilai false. Berikut adalah syntax untuk fungsi if:

> =IF\(logical\_test, value\_if\_true, \[value\_if\_false\]\)

**Deskripsi** 

* _logical\_test_, adalah kondisi yang ingin diuji
* _value\_if\_true,_ adalah nilai yang diberikan jika logical\_test bernilai benar
* _value\_if\_false_, adalah nilai yang diberikan jika logical\_test bernilai salah

logical\_test dan value\_if\_true wajib ada ketika menggunakan fungsi if, sedangkan value\_if\_false bersifat opsional.

{% hint style="info" %}
Jika Anda akan menggunakan teks dalam rumus, Anda perlu membungkus teks dalam tanda petik \(misalnya "teks"\).  Jika bukan teks, Anda tidak perlu menggunakan tanda petik.
{% endhint %}

## Fungsi MATCH

 **Fungsi MATCH** adalah fungsi Microsoft Excel yang digunakan untuk mencari _posisi_ suatu nilai pada suatu range yang terdapat pada sebuah kolom atau baris, tapi tidak kedua-duanya. Dengan kata lain Fungsi MATCH mencari item tertentu di rentang sel, kemudian menghasilkan posisi relatif item tersebut dalam rentang \(range\).

> =MATCH\(lookup\_value, lookup\_array, \[match\_type\]\)

**Deskripsi**

* _lookup\_value_, berisi objek yang ingin dicari dalam lookup\_array. lookup\_value dapat berisi angka, huruf, atau referensi sel.
* _lookup\_array_, berisi rentang sel untuk mencari keberadaan lookup\_value.
* _match\_type_,  berisi bilangan -1, 0, atau 1. Argumen match\_type menentukan bagaimana Excel mencocokkan lookup\_value dengan nilai dalam lookup\_array. **-1** : MATCH menemukan nilai yang lebih besar mendekati lookup\_value atau sama dengan lookup\_value **0** : MATCH menemukan nilai yang sama persis dengan lookup\_value **1** **atau dikosongkan** : MATCH menemukan nilai yang lebih kecil mendekati lookup\_value atau sama dengan lookup\_value.   lookup\_value dan lookup\_array wajib ada dalam penulisan rumus tersebut, sedangkan match\_type bersifat opsional.

**Contoh Penggunaan**

![](.gitbook/assets/contohexcel2.png)

Pada gambar tersebut terlihat pencarian nilai 82 dalam kolom C2 sampai C6, pada saat match\_type bernilai 1 maka hasilnya adalah 3 yang berarti nilai yang mendekati 82 dari bawah atau sama dengan 82 ada di baris ke 3 dalam kolom C, atau bisa disebut berada dalam sel C3.

{% hint style="info" %}
MATCH tidak membedakan antara huruf besar dan huruf kecil ketika mencocokkan nilai teks.

Jika MATCH tidak berhasil menemukan kecocokan, maka mengembalikan nilai kesalahan \#N/A.

Jika match\_type sama dengan 0 dan lookup\_value berupa string teks, Anda dapat menggunakan karakter wildcard berupa tanda tanya \(?\) dan tanda bintang \(\*\) dalam argumen lookup\_value. Tanda tanya \(?\) mencocokkan dengan karakter tunggal apa pun sedangkan tanda bintang \(\*\) mencocokan dengan urutan karakter apa pun.
{% endhint %}

## Fungsi HLOOKUP

Fungsi atau Rumus Excel HLOOKUP adalah salah satu fungsi ms. Excel yang digunakan untuk mencari data pada kolom pertama sebuah tabel data, kemudian mengambil nilai dari sel mana pun di baris yang sama pada tabel data tersebut.

> =HLOOKUP\(lookup\_value, table\_array, row\_index\_num, \[range\_lookup\]\)

**Deskripsi**

* _lookup\_value_,  Nilai yang dicari pada baris pertama tabel. lookup\_value bisa berupa nilai, referensi, atau string teks.
* _table\_array_,  Tabel data dimana lookup\_value akan di temukan pada baris pertamanya. Gunakan referensi ke sebuah range, tabel, nama range atau nama tabel.
*  _row\_index\_num_, Nomor baris dalam table\_array yang akan menghasilkan nilai yang cocok. 
* _range\_lookup_, berisi TRUE atau FALSE. Nilai logika ini akan menentukan apakah kita ingin Hlookup mencari kecocokan yang sama persis atau kecocokan yang mendekati. Jika berisi TRUE atau dikosongkan \(tidak diisi\) maka Hlookup menghasilkan kecocokan yang mendekati. Jika FALSE, HLOOKUP akan menemukan kecocokan persis. Jika tidak ditemukan, akan menghasilkan nilai kesalahan \#N/A.

lookup\_value, table\_array, row\_index\_num merupakan komponen yang wajib ada saat menggunakan rumus HLOOKUP, sedangkan range\_lookup bersifat opsional.

**Contoh Penggunaan**

![](.gitbook/assets/contohexcel3.png)

Pada gambar tersebut dicari sebuah nilai yang berada pada kolom C1 dan terletak pada baris ketiga

## Fungsi VLOOKUP

 Fungsi atau Rumus Excel VLOOKUP adalah salah satu fungsi ms. Excel yang digunakan untuk mencari data pada kolom pertama sebuah tabel data, kemudian mengambil nilai dari sel mana pun di baris yang sama pada tabel data tersebut.

> =HLOOKUP\(lookup\_value, table\_array, col\_index\_num, \[range\_lookup\]\)

**Deskripsi**

* _lookup\_value_,  Nilai yang dicari pada baris pertama tabel. lookup\_value bisa berupa nilai, referensi, atau string teks.
* _table\_array_,  Tabel data dimana lookup\_value akan di temukan pada baris pertamanya. Gunakan referensi ke sebuah range, tabel, nama range atau nama tabel.
*  col_\_index\_num_, Nomor kolom dalam table\_array yang akan menghasilkan nilai yang cocok. 
* _range\_lookup_, berisi TRUE atau FALSE. Nilai logika ini akan menentukan apakah kita ingin VLOOKUP mencari kecocokan yang sama persis atau kecocokan yang mendekati. Jika berisi TRUE atau dikosongkan \(tidak diisi\) maka VLOOKUP menghasilkan kecocokan yang mendekati. Jika FALSE, VLOOKUP akan menemukan kecocokan persis. Jika tidak ditemukan, akan menghasilkan nilai kesalahan \#N/A.

lookup\_value, table\_array, row\_index\_num merupakan komponen yang wajib ada saat menggunakan rumus HLOOKUP, sedangkan range\_lookup bersifat opsional.

**Contoh Penggunaan**

![](.gitbook/assets/contohexcel4.png)

Pada gambar tersebut dicari sebuah nilai yang terletak pada nomor yang terdapat pada sel H2, dalam rentang tabel dari A1 sampai E6 dan berada pada kolom ke 3.

## Fungsi ROUND

Fungsi ROUND berguna untuk membulatkan angka. Pembulatan yang dilakukan mengikuti kaidah matematika, yaitu membulatkan ke bawah jika angka terakhir suatu bilangan adalah 0, 1, 2, 3, 4 dan membulatkan ke atas jika angka terakhir suatu bilangan adalah 5, 6, 7, 8, 9. Jika bilangan tersebut adalah desimal, pembulatan kebawah dilakukan jika angka desimal kurang dari 0,5 dan pembulatan keatas jika lebih dari atau sama dengan 0,5.

> =ROUND\(number, num\_digit\)

**Deskripsi**

*  _number_, angka desimal atau pecahan yang aka dibulatkan ke beberapa tempat tertentu. Dapat berupa referensi suatu sel.
* _num\_digit_, banyak digit pembulatan yang ingin ditetapkan pada angka tersebut.

**Contoh Penggunaan**  
  

![](.gitbook/assets/contohexcel5.png)

 Saat menggunakan rumus Round excel, apabila **JumlahDigit lebih besar dari 0 \(nol\)**, maka angka desimal dibulatkan ke sejumlah tempat desimal yang ditentukan.  Apabila **JumlahDigit adalah 0**, angka desimal akan dibulatkan ke bilangan bulat \(tanpa nilai desimal/pecahan\) terdekat.  Apabila **JumlahDigit lebih kecil dari 0 \(negatif\)**, maka angka desimal akan dibulatkan ke sebelah kiri koma desimal.  


