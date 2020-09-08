Tools untuk auto collect cpu usage, memory usage, uptime, version, Serial Number pada switch Huawei yang akan digenerate jadi excel.

Yang perlu disiapkan :

1. File Log Switch Huawei yang berisi 'display diagnostic-information'
2. filenamenya harus sudah di rename sesuai hostnamenya karena mempengaruhi output hostname di excel.
3. Install package 'xlwt'


Cara Pakai :

1. pip install -r requirements.txt
2. Siapkan Folder 'Source' yang berisi kumpulan file/log switch huawei
3. Lalu jalankan script nya dan hasilnya akan berada di file 'result.xls'

Jika masih bingung bisa bertanya langsung di IG @am.cello

Thanks
