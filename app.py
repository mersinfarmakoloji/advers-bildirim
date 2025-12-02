import streamlit as st
from docx import Document
from datetime import date
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import re

st.set_page_config(page_title="Advers Bildirim v14", page_icon="🇹🇷", layout="centered")

# --- AYARLAR (BURAYI KENDİNE GÖRE DOLDUR) ---
GONDEREN_EMAIL = "mersinfarmakoloji@gmail.com"  # O yeni açtığın bot maili
ALICI_EMAIL = "mersinfarmakoloji@gmail.com"           # Rapor kime gidecek? (Senin asıl mailin)
# ŞİFREYİ KODUN İÇİNE YAZMIYORUZ! (Güvenlik için aşağıda anlatacağım yere yazacağız)

st.title("🇹🇷 T.C. Sağlık Bakanlığı - TÜFAM Bildirimi")
st.info("Formu doldurup gönderdiğinizde, rapor otomatik olarak yetkiliye e-posta ile iletilecektir.")

# --- YARDIMCI FONKSİYONLAR ---
def kutu_yap(secim, hedef):
    return "[X]" if secim == hedef else "[ ]"

def soru_cevapla(cevap):
    if cevap == "Evet": return "[X] Evet  [ ] Hayır  [ ] Bilinmiyor"
    if cevap == "Hayır": return "[ ] Evet  [X] Hayır  [ ] Bilinmiyor"
    return "[ ] Evet  [ ] Hayır  [X] Bilinmiyor"

def TR_upper(text):
    if text: return text.replace("i", "İ").upper()
    return ""

def TR_lower(text):
    if text: return text.replace("I", "ı").replace("İ", "i").lower()
    return ""

# --- A. HASTA ve CİDDİYET ---
st.header("A. HASTA & CİDDİYET")
c1, c2 = st.columns(2)
with c1:
    ad_soyad = st.text_input("1. Hasta Ad Soyad (Baş Harfler)", placeholder="Örn: A.Y.")
    dogum_tarihi = st.date_input("2. Doğum Tarihi", min_value=date(1900, 1, 1), max_value=date.today())
    
    bugun = date.today()
    yas_hesap = bugun.year - dogum_tarihi.year - ((bugun.month, bugun.day) < (dogum_tarihi.month, dogum_tarihi.day))
    st.caption(f"Hesaplanan Yaş: {yas_hesap}")

with c2:
    cinsiyet = st.radio("3. Cinsiyet", ["Kadın", "Erkek"], horizontal=True)
    boy = st.text_input("4. Boy (cm)", placeholder="170")
    kilo = st.text_input("5. Ağırlık (kg)", placeholder="70")

st.markdown("---")
st.subheader("⚠️ Ciddiyet Durumu")

ciddiyet_durumu = st.radio("Vaka Ciddi mi?", ["Ciddi Değil", "Ciddi"], horizontal=True)

k_olum_val, k_hayat_val, k_hastane_val, k_sakatlik_val, k_anomali_val, k_tibbi_val = False, False, False, False, False, False
olum_tarihi_str, olum_nedeni, otopsi = "", "", "[ ] Evet  [ ] Hayır"

if ciddiyet_durumu == "Ciddi":
    st.info("👇 Kriterleri işaretleyiniz:")
    with st.container():
        cols_cid = st.columns(2)
        with cols_cid[0]:
            k_olum_val = st.checkbox("💀 Ölüm")
            k_hayat_val = st.checkbox("❤️ Hayatı Tehdit Edici")
            k_hastane_val = st.checkbox("🏥 Hastaneye Yatış/Uzama")
        with cols_cid[1]:
            k_sakatlik_val = st.checkbox("♿ Kalıcı Sakatlık")
            k_anomali_val = st.checkbox("👶 Konjenital Anomali")
            k_tibbi_val = st.checkbox("⚕️ Tıbbi Olarak Önemli")

    if k_olum_val:
        st.error("Ölüm Detayları:")
        col_o1, col_o2 = st.columns(2)
        with col_o1:
            ot = st.date_input("Ölüm Tarihi", max_value=date.today())
            olum_tarihi_str = ot.strftime("%d.%m.%Y")
            oto = st.radio("Otopsi Yapıldı mı?", ["Evet", "Hayır"], horizontal=True)
            otopsi = "[X] Evet  [ ] Hayır" if oto == "Evet" else "[ ] Evet  [X] Hayır"
        with col_o2:
            olum_nedeni = st.text_input("Ölüm Nedeni")

# --- B. REAKSİYONLAR ---
st.header("B. ADVERS REAKSİYONLAR")
reaksiyonlar = []
for i in range(1, 6):
    with st.expander(f"Reaksiyon {i}", expanded=(i==1)):
        col_r1, col_r2, col_r3 = st.columns([3, 1, 1])
        with col_r1: r_tanim = st.text_input(f"Tanım", key=f"rt{i}")
        with col_r2: r_bas = st.date_input(f"Başlangıç", key=f"rb{i}", max_value=date.today())
        with col_r3: 
            r_devam = st.checkbox("Devam Ediyor", key=f"rd{i}")
            if r_devam:
                r_bit = "DEVAM EDİYOR"
            else:
                r_bit_date = st.date_input(f"Bitiş", key=f"rbit{i}", value=None, max_value=date.today())
                r_bit = r_bit_date

        if not r_devam and r_bit and r_bas and r_bit < r_bas:
             st.error("⚠️ HATA: Bitiş tarihi başlangıçtan önce olamaz!")

        if r_tanim: 
            reaksiyonlar.append({"tanim": r_tanim, "bas": r_bas, "bit": r_bit, "devam": r_devam})

st.subheader("Sonuç Durumu")
sonuc_secim = st.radio("Sonuç", ["İyileşti/Düzeldi", "İyileşiyor", "Sekel Bıraktı", "Devam Ediyor", "Ölümle Sonuçlandı", "Bilinmiyor"], horizontal=True)
lab_bulgu = st.text_area("3. Laboratuvar Bulguları (Tarihleriyle birlikte)", height=68)
st.info("ℹ️ **Tıbbi Öykü:** Allerji, gebelik, sigara/alkol, kronik hastalıklar vb.")
tibbi_oyku = st.text_area("4. Tıbbi Öykü / Eş Zamanlı Hastalıklar", height=68)

# --- C. İLAÇLAR ---
st.header("C. ŞÜPHELENİLEN İLAÇLAR")
ilaclar = []

for i in range(1, 6):
    with st.expander(f"💊 İlaç {i}", expanded=(i==1)):
        c_i1, c_i2, c_i3 = st.columns([2, 1, 1])
        with c_i1: 
            i_adi = st.text_input(f"İlaç Adı", key=f"ia{i}", help="Biliniyorsa TİCARİ ismini yazınız.")
        with c_i2: 
            i_yol_secim = st.selectbox(f"Veriliş Yolu", ["Oral", "IV", "IM", "SC", "Topikal", "Diğer"], key=f"iy{i}")
            if i_yol_secim == "Diğer":
                i_yol = st.text_input(f"👉 Yolu Yazınız ({i})", key=f"iy_txt{i}")
            else:
                i_yol = i_yol_secim
        with c_i3: 
            i_doz = st.text_input(f"Günlük Doz", placeholder="Örn: 500 mg", key=f"id{i}")
        
        c_i4, c_i5, c_i6 = st.columns([2, 1, 1])
        with c_i4: i_end = st.text_input(f"Endikasyon", key=f"ie{i}")
        with c_i5: i_bas = st.date_input(f"Başlama", key=f"ib{i}", max_value=date.today())
        with c_i6: 
            i_devam = st.checkbox("Kullanım Devam Ediyor", key=f"idvm{i}")
            if i_devam:
                i_bit = "DEVAM EDİYOR"
            else:
                i_bit_date = st.date_input(f"Kesilme", value=None, key=f"ibit{i}", max_value=date.today())
                i_bit = i_bit_date

        if not i_devam and i_bit and i_bas and i_bit < i_bas:
            st.error("⚠️ Kesilme tarihi başlama tarihinden önce olamaz!")

        st.markdown(f":blue[**⬇️ {i}. İlaç Değerlendirme Soruları:**]")
        q_col1, q_col2, q_col3, q_col4 = st.columns(4)
        with q_col1: q7 = st.selectbox("7. İlaç Kesildi mi?", ["Bilinmiyor", "Evet", "Hayır"], key=f"q7_{i}")
        with q_col2: q8 = st.selectbox("8. Reaksiyon azaldı mı?", ["Bilinmiyor", "Evet", "Hayır"], key=f"q8_{i}")
        with q_col3: q9 = st.selectbox("9. Yeniden verildi mi?", ["Bilinmiyor", "Evet", "Hayır"], key=f"q9_{i}")
        with q_col4: q10 = st.selectbox("10. Tekrarladı mı?", ["Bilinmiyor", "Evet", "Hayır"], key=f"q10_{i}")

        if i_adi: 
            ilaclar.append({
                "ad": i_adi, "yol": i_yol, "doz": i_doz, "bas": i_bas, "bit": i_bit, "end": i_end,
                "s7": soru_cevapla(q7), "s8": soru_cevapla(q8), "s9": soru_cevapla(q9), "s10": soru_cevapla(q10),
                "devam": i_devam
            })

st.info("ℹ️ Eş Zamanlı ilaçları virgül ile ayırarak yazınız.")
es_zamanli = st.text_area("11. Eş Zamanlı İlaçlar", height=68)
diger_gozlem = st.text_area("12. Diğer Gözlemler (Kalite sorunu vb.)", height=68)
tedavi = st.text_area("13. Advers Reaksiyonun Tedavisi", height=68)

# --- D. BİLDİREN ---
st.header("D. BİLDİRİM YAPAN KİŞİ")
c_d1, c_d2 = st.columns(2)
with c_d1:
    b_ad = st.text_input("1. Adı Soyadı")
    b_tel = st.text_input("3. Tel No")
    b_faks = st.text_input("5. Faks")
with c_d2:
    b_meslek = st.selectbox("2. Meslek", ["Doktor", "Eczacı", "Hemşire", "Diğer"])
    b_adres = st.text_area("4. Adresi", value="Mersin Üniversitesi Tıp Fakültesi", height=100)
    b_email = st.text_input("6. E-posta")

st.markdown("---")
col_r1, col_r2 = st.columns(2)
with col_r1:
    rapor_firma = st.radio("8. Rapor firmaya bildirildi mi?", ["Bilinmiyor", "Evet", "Hayır"], horizontal=True, index=None)
with col_r2:
    rapor_tipi = st.radio("10. Rapor Tipi", ["İlk", "Takip"], horizontal=True, index=None)

rapor_tarihi = st.date_input("9. Rapor Tarihi", value=date.today(), max_value=date.today())

st.markdown("---")
submitted = st.button("📤 BİLDİRİMİ GÖNDER", type="primary", use_container_width=True)

# --- KAYIT VE MAİL GÖNDERME ---
if submitted:
    if not ad_soyad or not ilaclar or not reaksiyonlar:
        st.error("⚠️ Lütfen en az Hasta Adı, Bir Reaksiyon ve Bir İlaç giriniz.")
    else:
        try:
            with st.spinner("Rapor oluşturuluyor ve mail gönderiliyor..."):
                doc = Document("Advers reaksiyon bildirim formu.docx")
                
                # --- VERİ HAZIRLIĞI (AYNI) ---
                r_list = [{"tanim":"", "bas":"", "bit":""} for _ in range(5)]
                for idx, r in enumerate(reaksiyonlar):
                    bitis_str = "DEVAM EDİYOR" if r["devam"] else (r["bit"].strftime("%d.%m.%Y") if r["bit"] else "")
                    if idx < 5:
                        r_list[idx] = {"tanim": TR_upper(r["tanim"]), "bas": r["bas"].strftime("%d.%m.%Y") if r["bas"] else "", "bit": bitis_str}

                i_list = [{"ad":"", "yol":"", "doz":"", "bas":"", "bit":"", "end":"", "s7":"", "s8":"", "s9":"", "s10":""} for _ in range(5)]
                for idx, ilac in enumerate(ilaclar):
                    bitis_str = "DEVAM EDİYOR" if ilac["devam"] else (ilac["bit"].strftime("%d.%m.%Y") if ilac["bit"] else "")
                    if idx < 5:
                        i_list[idx] = {
                            "ad": TR_upper(ilac["ad"]), "yol": TR_upper(ilac["yol"]), "doz": TR_lower(ilac["doz"]), 
                            "end": TR_upper(ilac["end"]), "bas": ilac["bas"].strftime("%d.%m.%Y") if ilac["bas"] else "", 
                            "bit": bitis_str,
                            "s7": ilac["s7"], "s8": ilac["s8"], "s9": ilac["s9"], "s10": ilac["s10"]
                        }

                def radio_kutu(secim, hedef): return "[X]" if secim == hedef else "[ ]"
                rf_str = "[ ] Evet [ ] Hayır [ ] Bilinmiyor" if rapor_firma is None else f"{radio_kutu(rapor_firma, 'Evet')} Evet  {radio_kutu(rapor_firma, 'Hayır')} Hayır  {radio_kutu(rapor_firma, 'Bilinmiyor')} Bilinmiyor"
                rt_str = "[ ] İlk [ ] Takip" if rapor_tipi is None else f"{radio_kutu(rapor_tipi, 'İlk')} İlk  {radio_kutu(rapor_tipi, 'Takip')} Takip"

                veriler = {
                    "{{hasta_adi_soyadi_basharfleri}}": TR_upper(ad_soyad), 
                    "{{dogum_tarihi}}": dogum_tarihi.strftime("%d.%m.%Y"), "{{yas}}": str(yas_hesap), "{{cinsiyet}}": cinsiyet, "{{boy}}": boy, "{{kilo}}": kilo,
                    "{{cid_yok}}": "[X]" if ciddiyet_durumu == "Ciddi Değil" else "[ ]", "{{cid_var}}": "[X]" if ciddiyet_durumu == "Ciddi" else "[ ]",
                    "{{k_olum}}": "[X]" if k_olum_val else "[ ]", "{{k_hayat}}": "[X]" if k_hayat_val else "[ ]",
                    "{{k_hastane}}": "[X]" if k_hastane_val else "[ ]", "{{k_sakatlik}}": "[X]" if k_sakatlik_val else "[ ]",
                    "{{k_anomali}}": "[X]" if k_anomali_val else "[ ]", "{{k_tibbi}}": "[X]" if k_tibbi_val else "[ ]",
                    "{{olum_tarih}}": olum_tarihi_str, "{{olum_neden}}": TR_upper(olum_nedeni), "{{otopsi}}": otopsi,
                    "{{reaksiyon_1}}": r_list[0]["tanim"], "{{bas_1}}": r_list[0]["bas"], "{{bit_1}}": r_list[0]["bit"],
                    "{{reaksiyon_2}}": r_list[1]["tanim"], "{{bas_2}}": r_list[1]["bas"], "{{bit_2}}": r_list[1]["bit"],
                    "{{reaksiyon_3}}": r_list[2]["tanim"], "{{bas_3}}": r_list[2]["bas"], "{{bit_3}}": r_list[2]["bit"],
                    "{{reaksiyon_4}}": r_list[3]["tanim"], "{{bas_4}}": r_list[3]["bas"], "{{bit_4}}": r_list[3]["bit"],
                    "{{reaksiyon_5}}": r_list[4]["tanim"], "{{bas_5}}": r_list[4]["bas"], "{{bit_5}}": r_list[4]["bit"],
                    "{{s_iyilesti}}": kutu_yap(sonuc_secim, "İyileşti/Düzeldi"), "{{s_iyilesiyor}}": kutu_yap(sonuc_secim, "İyileşiyor"), "{{s_sekel}}": kutu_yap(sonuc_secim, "Sekel Bıraktı"),
                    "{{s_devam}}": kutu_yap(sonuc_secim, "Devam Ediyor"), "{{s_olum}}": kutu_yap(sonuc_secim, "Ölümle Sonuçlandı"), "{{s_bilinmiyor}}": kutu_yap(sonuc_secim, "Bilinmiyor"),
                    "{{lab}}": TR_upper(lab_bulgu), "{{oyku}}": TR_upper(tibbi_oyku), "{{tedavi}}": TR_upper(tedavi), "{{diger_gozlem}}": TR_upper(diger_gozlem),
                    "{{ilac_1}}": i_list[0]["ad"], "{{yol_1}}": i_list[0]["yol"], "{{doz_1}}": i_list[0]["doz"], "{{ilac_bas_1}}": i_list[0]["bas"], "{{ilac_bit_1}}": i_list[0]["bit"], "{{end_1}}": i_list[0]["end"], "{{s7_1}}": i_list[0]["s7"], "{{s8_1}}": i_list[0]["s8"], "{{s9_1}}": i_list[0]["s9"], "{{s10_1}}": i_list[0]["s10"],
                    "{{ilac_2}}": i_list[1]["ad"], "{{yol_2}}": i_list[1]["yol"], "{{doz_2}}": i_list[1]["doz"], "{{ilac_bas_2}}": i_list[1]["bas"], "{{ilac_bit_2}}": i_list[1]["bit"], "{{end_2}}": i_list[1]["end"], "{{s7_2}}": i_list[1]["s7"], "{{s8_2}}": i_list[1]["s8"], "{{s9_2}}": i_list[1]["s9"], "{{s10_2}}": i_list[1]["s10"],
                    "{{ilac_3}}": i_list[2]["ad"], "{{yol_3}}": i_list[2]["yol"], "{{doz_3}}": i_list[2]["doz"], "{{ilac_bas_3}}": i_list[2]["bas"], "{{ilac_bit_3}}": i_list[2]["bit"], "{{end_3}}": i_list[2]["end"], "{{s7_3}}": i_list[2]["s7"], "{{s8_3}}": i_list[2]["s8"], "{{s9_3}}": i_list[2]["s9"], "{{s10_3}}": i_list[2]["s10"],
                    "{{ilac_4}}": i_list[3]["ad"], "{{yol_4}}": i_list[3]["yol"], "{{doz_4}}": i_list[3]["doz"], "{{ilac_bas_4}}": i_list[3]["bas"], "{{ilac_bit_4}}": i_list[3]["bit"], "{{end_4}}": i_list[3]["end"], "{{s7_4}}": i_list[3]["s7"], "{{s8_4}}": i_list[3]["s8"], "{{s9_4}}": i_list[3]["s9"], "{{s10_4}}": i_list[3]["s10"],
                    "{{ilac_5}}": i_list[4]["ad"], "{{yol_5}}": i_list[4]["yol"], "{{doz_5}}": i_list[4]["doz"], "{{ilac_bas_5}}": i_list[4]["bas"], "{{ilac_bit_5}}": i_list[4]["bit"], "{{end_5}}": i_list[4]["end"], "{{s7_5}}": i_list[4]["s7"], "{{s8_5}}": i_list[4]["s8"], "{{s9_5}}": i_list[4]["s9"], "{{s10_5}}": i_list[4]["s10"],
                    "{{bildiren_ad}}": TR_upper(b_ad), "{{bildiren_meslek}}": b_meslek, "{{bildiren_tel}}": b_tel, 
                    "{{bildiren_adres}}": TR_upper(b_adres), "{{bildiren_faks}}": b_faks, "{{bildiren_email}}": b_email,
                    "{{rapor_tarihi}}": rapor_tarihi.strftime("%d.%m.%Y"),
                    "{{rapor_firma}}": rf_str, "{{rapor_tipi}}": rt_str,
                    "{{es_zamanli}}": TR_upper(es_zamanli)
                }

                def replace_fast(doc, data):
                    for p in doc.paragraphs:
                        if "{{" in p.text: 
                            for key, value in data.items():
                                if key in p.text: p.text = p.text.replace(key, str(value))
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    if "{{" in p.text:
                                        for key, value in data.items():
                                            if key in p.text: p.text = p.text.replace(key, str(value))
                    regex = re.compile(r"\{\{.*?\}\}") 
                    for p in doc.paragraphs:
                        if "{{" in p.text: p.text = regex.sub("", p.text)
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    if "{{" in p.text: p.text = regex.sub("", p.text)

                replace_fast(doc, veriler)
                bio = BytesIO()
                doc.save(bio)
                
                # --- MAİL GÖNDERME KISMI ---
                try:
                    # GMAIL AYARLARI (Şifreyi secrets'dan alacağız)
                    # Localde çalışırken secrets yoksa hata verebilir, o yüzden try-except
                    GMAIL_SIFRE = st.secrets["GMAIL_PASS"] 
                    
                    msg = MIMEMultipart()
                    msg['From'] = GONDEREN_EMAIL
                    msg['To'] = ALICI_EMAIL
                    msg['Subject'] = f"Advers Bildirim Raporu - {TR_upper(ad_soyad)}"
                    
                    body = f"Sayın Yetkili,\n\n{TR_upper(ad_soyad)} hastasına ait Advers Reaksiyon Bildirim Formu ektedir.\n\nBildiren: {TR_upper(b_ad)}\nTarih: {date.today().strftime('%d.%m.%Y')}"
                    msg.attach(MIMEText(body, 'plain'))
                    
                    # Dosyayı ekle
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(bio.getvalue())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename="Advers_{ad_soyad}.docx"')
                    msg.attach(part)
                    
                    # Sunucuya bağlan ve gönder
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(GONDEREN_EMAIL, GMAIL_SIFRE)
                    server.sendmail(GONDEREN_EMAIL, ALICI_EMAIL, msg.as_string())
                    server.quit()
                    
                    st.success(f"✅ Rapor başarıyla oluşturuldu ve {ALICI_EMAIL} adresine gönderildi!")
                    
                except Exception as mail_err:
                    st.warning(f"⚠️ Rapor oluşturuldu ancak mail gönderilemedi. (Sebep: {mail_err})")
                    st.info("💡 Not: Kendi bilgisayarınızda (Local) çalışırken mail atması için 'secrets.toml' ayarı gerekir. Buluta yükleyince çalışacaktır.")
                
                # Her durumda indirme butonu da olsun
                st.download_button(label="📥 RAPORU İNDİR", data=bio.getvalue(), file_name=f"Advers_{ad_soyad}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
        except Exception as e:
            st.error(f"Hata: {e}")