import pandas as pd
import numpy as np
# Excel dosyasının yolu
excel_dosya_yolu = 'tryExcel.xlsx'  

# Excel dosyasını oku
veri_cercevesi = pd.read_excel(excel_dosya_yolu)                               

# Rastgele sayılar ürettik
rastgele_sayilar = np.random.randint(low=1, high=30, size=20)

# Veri çerçevesinde belirli bir sütunu seçtik
sutun_adi1 = veri_cercevesi.columns[0]  # Sütun adını değiştirdik
veri_cercevesi[sutun_adi1]=rastgele_sayilar
veri_cercevesi.to_excel(excel_dosya_yolu, index=False)
# Rastgele sayılar üret
rastgele_sayilar = np.random.randint(low=1, high=30, size=20)
sutun_adi2 = veri_cercevesi.columns[1]  
# Veri çerçevesinde belirli bir sütunu seçtik
veri_cercevesi[sutun_adi2]=rastgele_sayilar
# Veriyi güncellendikten sonra Excel dosyasına yazdık
veri_cercevesi.to_excel(excel_dosya_yolu, index=False)

# Veri çerçevesini görüntüledik
print("Güncellenmiş Veri Çerçevesi:")
print(veri_cercevesi)

bireBirAyniDeger = sum(veri_cercevesi[sutun_adi1] == veri_cercevesi[sutun_adi2])

i=1
aynıDegerDizesi=[]
while i<len(veri_cercevesi[sutun_adi1]):
    j=1
    while j<len(veri_cercevesi[sutun_adi1]):
        if veri_cercevesi[sutun_adi1][i]==veri_cercevesi[sutun_adi2][j]:  
            l=1
            while l<len(veri_cercevesi[sutun_adi1]):
                if veri_cercevesi[sutun_adi1][i]==veri_cercevesi[sutun_adi1][l]:
                    aynıDegerDizesi.append(veri_cercevesi[sutun_adi1][i])
                l=l+1
            print("Aynı Degerler : ",veri_cercevesi[sutun_adi1][i])
        j=j+1
    i=i+1

benzersizElemanlar=set(aynıDegerDizesi)
print("Aynı Deger Dizesi",benzersizElemanlar," Adedi : ",len(benzersizElemanlar))
print("İki sütunun birebir aynı değere sahip satırlarının adedi : ",bireBirAyniDeger) 