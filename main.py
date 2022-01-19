import pandas as pd

dosya_yolu = input("Muavina ait tam dosya yolunu giriniz: ")
dosya_yolu.replace("/","\\")

df = pd.read_excel(dosya_yolu)

basliklar = df.columns
print("""Dosyanızdaki başlıklar aşağıdaki gibidir:""")
print("")

for i in basliklar:
    print(i)


print("")
print("""Başlıklarınızı verilen başlıklarla eşleştiriniz.""")


alt_hesap = input("Alt Hesap: ")
hesap_adi = input("Hesap Adı: ")
fis_tarihi = input("Fiş Tarihi: ")
fis_no = input("Fiş Numarası: ")
aciklama = input("Açıklama: ")
tutar= input("Tutar(TL): ")

eslestirilen_basliklar = [alt_hesap,hesap_adi,fis_tarihi,fis_no,aciklama,tutar]

for i in basliklar:
    if i not in eslestirilen_basliklar:
        df.drop(columns=[i], inplace=True)

df.rename(columns = {alt_hesap:'Alt Hesap', hesap_adi:'Hesap Adı',fis_no:'Fiş Numarası', aciklama:'Açıklama', tutar:'Tutar(TL)', fis_tarihi:'Fiş Tarihi'}, inplace = True)

secilen_hesaplar_dosya_yolu = input("İncelenecek hesap kodlarının olduğu dosyaya ait tam dosya yolunu giriniz: ")
secilen_hesaplar_dosya_yolu.replace("/","\\")

df_secilen_hesaplar = pd.read_excel(secilen_hesaplar_dosya_yolu, header=None) # kolon ismi yok varsaydım.

df_secilen_hesaplar.columns = ["Hesap_kodu"]

islenecek_liste = df_secilen_hesaplar["Hesap_kodu"].to_list()

appended_data = []
for i in islenecek_liste:
    exam_df = df.loc[df["Alt Hesap"] == i]
    amounts = exam_df["Tutar(TL)"]
    Q1 = amounts.quantile(0.25)
    Q3 = amounts.quantile(0.75)
    IQR = Q3-Q1
    alt_sinir = Q1- 1.5*IQR
    ust_sinir = Q3 + 1.5*IQR
    aykiri_tf = (amounts > ust_sinir) | (exam_df < alt_sinir)
    incele = exam_df[aykiri_tf]
    appended_data.append(incele)

appended_data = pd.concat(appended_data)

appended_data.to_excel('appended.xlsx')
