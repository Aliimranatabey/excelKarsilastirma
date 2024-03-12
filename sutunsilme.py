import pandas as pd

# Excel dosyalarını oku
dosya1 = pd.read_excel('5GhzKapali10K.xlsx')
dosya2 = pd.read_excel('cikti.xlsx')

# '5Ghz_Kapali_Cihaz_Listesi_070324.xlsx' dosyasındaki MAC sütunundaki verileri içeren satırları seç
eslesen_satirlar = dosya1['MAC'].isin(dosya2['EŞLEŞME'])

# Eşleşmeyen verileri 'soncikti' isimli yeni bir Excel dosyasına yaz
soncikti_dosyasi = dosya1[~eslesen_satirlar]
soncikti_dosyasi.to_excel('Book1.xlsx', index=False)

