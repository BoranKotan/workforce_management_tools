import pandas as pd
import random
# Excel dosyasını oku
excel_file_path = 'ornek.xlsx'  # Excel dosyanızın yolunu belirtin
df = pd.read_excel(excel_file_path)

# Gün sütunları
gun_sutunlari = ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma']

# Her bir satır için gez
for index, row in df.iterrows():
    # Her bir gün için kontrol et
    secilen_gun = random.choice(gun_sutunlari)
    df.at[index, secilen_gun] = "OFF"


# Sonuçları göster
#print(df)

# Sonuçları Excel dosyasına yaz
df.to_excel('ornek2.xlsx', index=False)
