---
description: >-
  Berikut adalah beberapa catatan dari saya tentang penggunaan beberapa fungsi
  yang cukup sering digunakan dalam pengolahan data menggunakan Excel.
---

# Fungsi pada Excel yang sering digunakan

## Fungsi Count

Fungsi count adalah fungsi untuk menghitung berapa banyak angka dari suatu rentang sel.  Misalnya, dengan menggunakan fungsi ini seseorang bisa memasukkan rumus berikut untuk menghitung angka dalam rentang A1:A20: =COUNT\(A1:A20\). Dalam contoh ini, jika ada lima sel dalam rentang berisikan angka, hasilnya adalah 5.

## Fungsi Counta

Fungsi counta adalah fungsi untuk menghitung isi dari suatu rentang sel apapun isi dari sel tersebut.  Fungsi counta menghitung sel yang berisi setiap tipe informasi, termasuk nilai kesalahan dan teks kosong \(**""**\). Misalnya, jika rentang tersebut berisi rumus yang mengembalikan string kosong, fungsi counta menghitung nilai itu. Fungsi counta tidak menghitung sel kosong.

## Fungsi Countif

Fungsi countif adalah fungsi untuk menghitung value spesifik dari suatu rentang sel. 

Secara sederhana, COUNTIF berbentuk sebagai berikut:

* =COUNTIF\(Rentang sel yang ingin dicari?, Apa yang ingin dicari?\)

## Fungsi Sum

Fungsi Sum adalah fungsi untuk menjumlah angka pada suatu rentang sel. Seseorang dapat menambahkan nilai individual, referensi sel atau rentang, atau campuran ketiganya.

## Fungsi Sumif

Fungsi Sumif adalah fungsi untuk menjumlah angka yang memenuhi suatu kondisi pada suatu rentang sel. Sebagai contoh, di dalam kolom yang berisi angka, Anda hanya ingin menjumlahkan nilai-nilai yang lebih besar dari 5. Anda bisa menggunakan rumus berikut:=SUMIF\(A1:A5,"&gt;5"\)

## Fungsi Sumifs

Fungsi Sumifs adalah fungsi untuk menjumlahkan angka yang memnuhi beberapa kondisi dalam suatu rentang sel. Sebagai contoh, gunakan SUMIFS untuk menjumlahkan jumlah pengecer di negara yang \(1\) berada dalam satu kode pos dan \(2\) yang labanya melebihi nilai dolar tertentu.

