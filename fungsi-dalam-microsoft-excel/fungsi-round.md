# Fungsi Round

Fungsi ROUND berguna untuk membulatkan angka. Pembulatan yang dilakukan mengikuti kaidah matematika, yaitu membulatkan ke bawah jika angka terakhir suatu bilangan adalah 0, 1, 2, 3, 4 dan membulatkan ke atas jika angka terakhir suatu bilangan adalah 5, 6, 7, 8, 9. Jika bilangan tersebut adalah desimal, pembulatan kebawah dilakukan jika angka desimal kurang dari 0,5 dan pembulatan keatas jika lebih dari atau sama dengan 0,5.

> =ROUND\(number, num\_digit\)

**Deskripsi**

*  _number_, angka desimal atau pecahan yang aka dibulatkan ke beberapa tempat tertentu. Dapat berupa referensi suatu sel.
* _num\_digit_, banyak digit pembulatan yang ingin ditetapkan pada angka tersebut.

**Contoh Penggunaan**  
  

![](../.gitbook/assets/contohexcel5.png)

 Saat menggunakan rumus Round excel, apabila **JumlahDigit lebih besar dari 0 \(nol\)**, maka angka desimal dibulatkan ke sejumlah tempat desimal yang ditentukan.  Apabila **JumlahDigit adalah 0**, angka desimal akan dibulatkan ke bilangan bulat \(tanpa nilai desimal/pecahan\) terdekat.  Apabila **JumlahDigit lebih kecil dari 0 \(negatif\)**, maka angka desimal akan dibulatkan ke sebelah kiri koma desimal.  


