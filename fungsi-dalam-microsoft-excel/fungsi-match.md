# Fungsi Match

**Fungsi MATCH** adalah fungsi Microsoft Excel yang digunakan untuk mencari _posisi_ suatu nilai pada suatu range yang terdapat pada sebuah kolom atau baris, tapi tidak kedua-duanya. Dengan kata lain Fungsi MATCH mencari item tertentu di rentang sel, kemudian menghasilkan posisi relatif item tersebut dalam rentang \(range\).

> =MATCH\(lookup\_value, lookup\_array, \[match\_type\]\)

**Deskripsi**

* _lookup\_value_, berisi objek yang ingin dicari dalam lookup\_array. lookup\_value dapat berisi angka, huruf, atau referensi sel.
* _lookup\_array_, berisi rentang sel untuk mencari keberadaan lookup\_value.
* _match\_type_,  berisi bilangan -1, 0, atau 1. Argumen match\_type menentukan bagaimana Excel mencocokkan lookup\_value dengan nilai dalam lookup\_array. **-1** : MATCH menemukan nilai yang lebih besar mendekati lookup\_value atau sama dengan lookup\_value **0** : MATCH menemukan nilai yang sama persis dengan lookup\_value **1** **atau dikosongkan** : MATCH menemukan nilai yang lebih kecil mendekati lookup\_value atau sama dengan lookup\_value.   lookup\_value dan lookup\_array wajib ada dalam penulisan rumus tersebut, sedangkan match\_type bersifat opsional.

**Contoh Penggunaan**

![](../.gitbook/assets/contohexcel2.png)

Pada gambar tersebut terlihat pencarian nilai 82 dalam kolom C2 sampai C6, pada saat match\_type bernilai 1 maka hasilnya adalah 3 yang berarti nilai yang mendekati 82 dari bawah atau sama dengan 82 ada di baris ke 3 dalam kolom C, atau bisa disebut berada dalam sel C3.

{% hint style="info" %}
MATCH tidak membedakan antara huruf besar dan huruf kecil ketika mencocokkan nilai teks.

Jika MATCH tidak berhasil menemukan kecocokan, maka mengembalikan nilai kesalahan \#N/A.

Jika match\_type sama dengan 0 dan lookup\_value berupa string teks, Anda dapat menggunakan karakter wildcard berupa tanda tanya \(?\) dan tanda bintang \(\*\) dalam argumen lookup\_value. Tanda tanya \(?\) mencocokkan dengan karakter tunggal apa pun sedangkan tanda bintang \(\*\) mencocokan dengan urutan karakter apa pun.
{% endhint %}

