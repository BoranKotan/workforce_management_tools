import pandas as pd
import random
# Excel dosyasını okuma
excel_dosya_adı = 'ornek.xlsx'
veri = pd.read_excel(excel_dosya_adı)

# Anahtar kelime kontrolü ve haftanın tamamını değiştirme
anahtar_kelime = '13:00-23:00'
hedef_vardiya_haftaici = '17:30-02:00'
hedef_vardiya_haftasonu = '13:00-23:00'
kisi_sayisi = 0

degistir = False  # Haftanın tamamını değiştirmek için kontrol
gun_sutunlari = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar']
haftaici_sutunları = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma']
haftasonu_sutunları = ['Cumartesi', 'Pazar']


for gun in haftaici_sutunları:
    if anahtar_kelime in veri[gun].values:
        degistir = True
    if degistir:
        veri[gun] = hedef_vardiya_haftaici


if degistir:
    for gun in haftasonu_sutunları:
        veri[gun] = hedef_vardiya_haftasonu

vardiyadaki_kisi_sayisi = len(veri[gun])
kalan = vardiyadaki_kisi_sayisi % 5



katsayi1 = vardiyadaki_kisi_sayisi / 5
katsayi2 = vardiyadaki_kisi_sayisi % 5
    
def off_verme(veri, haftaici_sutunları, katsayi1, katsayi2):
    off_verilen = 0
    for index, row in veri.iterrows():
        for gun in haftaici_sutunları:
            if off_verilen < katsayi1:
                # Rastgele bir gün seçip OFF ver
                off_gun = random.choice(haftaici_sutunları)
                veri.at[index, off_gun] = 'OFF'
                off_verilen += 1
            elif off_verilen < katsayi1 + katsayi2:
                # Kalan OFF'ları rastgele günlerde dağıt
                off_gun = random.choice(haftaici_sutunları)
                if veri.at[index, off_gun] != 'OFF':
                    veri.at[index, off_gun] = 'OFF'
                    off_verilen += 1

off_verme(veri, haftaici_sutunları, katsayi1, katsayi2)

# Sonuçları yeni bir Excel dosyasına yazma
yeni_excel_dosya_adı = 'ornek.xlsx'
veri.to_excel(yeni_excel_dosya_adı, index=False)

print(f'Kural uygulama işlemi tamamlandı. Sonuçlar "{yeni_excel_dosya_adı}" dosyasına kaydedildi.')