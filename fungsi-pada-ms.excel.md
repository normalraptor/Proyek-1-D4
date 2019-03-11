# Fungsi Pada MS.EXCEL

## 1. Fungsi REPT

Penerapan dari fungsi REPT ini adalah akan menampilkan teks sebanyak jumlah yang diinginkan.

### Syntax                               

###                          $$REPT(teks, numbertimes)$$ 

### Contoh 

|  | A | B |
| :--- | :--- | :--- |
| 1 | Saya |  |
| 2 | Saya Adalah |  |
| 3 | Saya Adalah Manusia |  |

### Output

=REPT\(A1,3\)

output= Saya Saya Saya 

=REPT\(A2,3\)

output= Saya Adalah Saya Adalah Saya Adalah 

=REPT\(A3,3\)

output= Saya Adalah Manusia Saya Adalah Manusia Saya Adalah Manusia 

{% hint style="info" %}
 **Keterangan** 

Jika number\_times adalah 0 \(nol\), REPT Menampilkan cell kosong. 

Jika number\_times bukan bilangan bulat Contoh 2,5 2 Untuk Angka Dibelakang koma akan dihilangkan.

Hasil dari fungsi REPT tidak bisa lebih panjang dari 32.767 karakter, Apabila Karakter Melebihi angka tersebut Maka Hasilnya adalah  \#VALUE.
{% endhint %}

## Fungsi Date, Month, Year 

| A | B | C |
| :--- | :--- | :--- |
| Fungsi  | Hasil |  |
| Date | 10/10/2010 |  |
| Month | 10 |  |
| Year | 2010 |  |

{% tabs %}
{% tab title="Fungsi Date" %}
fungsi date, digunakan untuk menuliskan tanggal dengan format tahun. bulan dan hari.

### Syntax

=DATE\(year;month;day\)

output 

=DATE\(2010;10;10\)

output 10/10/2010

{% hint style="info" %}
Year, diisi tahun yang diinginkan 

Month, diisi bulan yang diinginkan 

Day, diisi hari yang diinginkan 
{% endhint %}
{% endtab %}

{% tab title="Fungsi Month" %}
Fungsi month, digunakan untuk mengembalikan dari sebuah tanggal yang dinyatakan oleh nomor seri, Bulan ditentukan sebagai bilangan bulatnya, rentannya 1-12

### Syntax

### MONTH\(serial\_number\)

### Contoh

=MONTH\(DATE\(2010;10;10\)\)

output 10.

{% hint style="info" %}
Serial number, Diperlukan. Tanggal yang ingin Anda cari bulannya. Tanggal harus dimasukkan menggunakan fungsi DATE, atau sebagai hasil dari rumus atau fungsi lain. Misalnya, gunakan DATE\(2010;10;10\) untuk tanggal 10 november 2010
{% endhint %}
{% endtab %}

{% tab title="Fungsi Year" %}
Fungsi year, mengembalikan tahun yang terkait dalam suatu tanggal, Tahun dikembalikan dalam bilangan bulat 

### Syntax

### YEAR\(serial\_number\)

### Contoh 

### YEAR\(DATE\(2010;10;10\)\)

output 2010.
{% endtab %}
{% endtabs %}

