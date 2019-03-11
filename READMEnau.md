# 5 Fungsi Penting EXCEL

## **1.FUNGSI MAX**

Fungsi ini berfungsi untuk mengembalikan nilai terbesar pada sekumpulan nilai yang di input user.Berikut Sintak dalam penggunaan fungsi MAX.

### Syntax

$$
MAX(number1, [number2], ...)
$$

### Contoh

|  | A | B | C |
| :--- | :--- | :--- | :--- |
| 1 |  | 12 |  |
| 2 |  | 10 |  |
| 3 |  | 3 |  |
| 4 |  | 2 |  |
| 5 |  | 15 |  |
| 6 |  | 1 |  |
| 7 |  | 3 |  |

{% tabs %}
{% tab title="Menggunakan Rentan Pada Tabel" %}
#### MAX\(B1:B7\) 

Output  :  15
{% endtab %}

{% tab title="Menggunakan Angka" %}
#### **MAX**\(12,10,3,2,15,1,3\)

Output : 15
{% endtab %}

{% tab title="Mix" %}
#### MAX\(B1:B7,16,20\)

Output : 20  
{% endtab %}

{% tab title="Menggunakan Operasi Logika" %}
#### MAX\(if\(1&gt;2,2,100\),3,4,5,B1:B7\)

Output : 100
{% endtab %}
{% endtabs %}

{% hint style="info" %}
#### Keterangan

* Argumen bisa berupa angka atau nama, array, atau referensi yang berisi angka-angka.
* Nilai logika dan teks representasi angka yang Anda ketikkan secara langsung ke dalam daftar argumen akan dihitung.
* Jika argumen tidak berisi angka, maka MAX mengembalikan 0 \(nol\).

  Argumen yang merupakan nilai kesalahan atau teks yang tidak dapat diterjemahkan menjadi angka menyebabkan kesalahan.
{% endhint %}

## 2.Fungsi

Fungsi ini berfungsi untuk mengembalikan nilai terkecil pada sekumpulan nilai yang di input user.Berikut Sintak dalam penggunaan fungsi MIN.

### Syntax

$$
MIN(number1, [number2], ...)
$$

### Contoh

|  | A | B | C | D | E |
| :--- | :--- | :--- | :--- | :--- | :--- |
| 1 |  |  |  |  |  |
| 2 | 3 | 12 | 10 | 51 | 1 |
| 3 |  |  |  |  |  |

{% tabs %}
{% tab title="Menggunakan Rentan Pada Tabel" %}
#### MIN\(A2:E1\) 

Output  :  1
{% endtab %}

{% tab title="Menggunakan Angka" %}
#### MIN\(1,3,7,6,103,123,5\) 

Output  :  1
{% endtab %}

{% tab title="Menggunakan Operasi Logika " %}
#### MAX\(if\(1&gt;2,2,-5\),3,4,5,A2:E2\)

Output : -5
{% endtab %}
{% endtabs %}

{% hint style="info" %}
#### Keterangan

* Argumen bisa berupa angka atau nama, array, atau referensi yang berisi angka-angka.
* Nilai logika dan teks representasi angka yang Anda ketikkan secara langsung ke dalam daftar argumen akan dihitung.
* Jika argumen tidak berisi angka, maka MIN mengembalikan 0 \(nol\).
* Argumen yang merupakan nilai kesalahan atau teks yang tidak dapat diterjemahkan menjadi angka menyebabkan kesalahan.
{% endhint %}

## 3.FUNGSI CONCENATE

 Satu keahlian utama untuk berkerja dengan menggabungkan cell, juga dikenal sebagai **concatenation / penggabungan**. Mari kita lihat pada bagaimana menggabungkan data dalam Microsoft Excel.

### Syntax

$$
CONCATENATE([String 1] , [String 2] ...... , )
$$

### Contoh

|  | A | B | C |
| :--- | :--- | :--- | :--- |
| 1 | Dwinanda | Ganteng | =CONCATENATE\(A1:B2\) |
| 2 | Haqi | Tampan | =CONCATENATE\(A2,B2\) |
| 3 | Suplana | Cerdas | =CONCATENATE\(A3," ",B3\) |
| 4 | Riris | Pintar | =CONCATENATE\(A4," ","Sangat"," ",B4\) |

{% tabs %}
{% tab title="C1" %}
#### Output : \#VALUE!

Terjadi error karena penggabungan tidak dapat dilakukan dengan melakukan drag terhadap cell
{% endtab %}

{% tab title="C2" %}
#### Output : HaqiTampan

 Fungsi ini tidak memisahkan antar kalimat dengan spas i
{% endtab %}

{% tab title="C3" %}
#### Output : Suplana Cerdas

Output terdapat spasi karena ditambahkan string \( "\[spasi\] "\) 
{% endtab %}

{% tab title="C4" %}
#### Output : Riris Sangat Pintar

Seperti halnya penambahan spasi , kita juga dapat menambahkan kata-kata yang tidak ada pada cell ,setiap kata-kata yang ingin ditambahkan harus diawali dan diakhiri tanda kutip 2 \(" \[Kata\] "\)
{% endtab %}
{% endtabs %}

## 4.FUNGSI  LEFT , RIGHT , DAN MID

Ketiga Fungsi diatas memiliki inti yang sama ,yaitu mengambil beberapa huruf dari suatu cell .Perbedaan lebih detailnya akan dijelaskan pada contoh.

|  | A | B | C |
| :--- | :--- | :--- | :--- |
| 1 | DWINANDA | =LEFT\(A1,3\)   | =LEFT\("DWINANDA",3\) |
| 2 |  | =MID\(A1,3,4\) | =MID\("DWINANDA",3,4\) |
| 3 |  | =RIGHT\(A1,4\) | =RIGHT\("DWINANDA",4\) |

{% tabs %}
{% tab title="" %}

{% endtab %}

{% tab title="Fungsi Left \(B1&C1\)" %}
### SYNTAX : 

#### LEFT\(\[CELL\],\[Banyak Huruf\]\)

Output B1 & C1 : DWI

Penggunaan Fungsi Left adalah untuk mengambil sebanyak **\[Banyak Huruf\]** dari huruf paling kiri pada **\[CELL\]**.

{% hint style="info" %}
**\[CELL\]** dapat berisi alamat cell \(contoh B1\) atau teks yang diinput langsung dengan diapit dua tanda kutip \(contoh C1\)
{% endhint %}
{% endtab %}

{% tab title="Fungsi MID\(B2&C2\)" %}
### SYNTAX : 

#### MID\(\[CELL\],\[HURUF PERTAMA\],\[Banyak Huruf\]\)

Output B2 & C2 : INAN

Penggunaan Fungsi Left adalah untuk mengambil sebanyak **\[Banyak Huruf\]** dari huruf ke **\[HURUF PERTAMA\]** pada **\[CELL\]**.

{% hint style="info" %}
**\[CELL\]** dapat berisi alamat cell \(contoh B2\) atau teks yang diinput langsung dengan diapit dua tanda kutip \(contoh C2\)

**\[HURUF PERTAMA\]**  adalah huruf yang akan menjadi patokan awal misal pada contoh 

\[HURUF PERTAMA\] adalah huruf ke 3 yang berati dimulai dari huruf I
{% endhint %}
{% endtab %}

{% tab title="Fungsi Left\(B3&C3\)" %}
### SYNTAX : 

#### RIGHT\(\[CELL\],\[Banyak Huruf\]\)

Output B3 & C3 : ANDA

Penggunaan Fungsi Left adalah untuk mengambil sebanyak **\[Banyak Huruf\]** dari huruf paling kana pada **\[CELL\]**.

{% hint style="info" %}
**\[CELL\]** dapat berisi alamat cell \(contoh B3\) atau teks yang diinput langsung dengan diapit dua tanda kutip \(contoh C3\)
{% endhint %}
{% endtab %}
{% endtabs %}

