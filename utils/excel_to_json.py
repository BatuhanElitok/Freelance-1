import pandas as pd
import json
import os
from collections import defaultdict

def get_university_details(unis_start_path):
    """
    Üniversite detaylarını (şehir, logo, banner, üniversite türü) içeren bir sözlük oluşturur.
    """
    try:
        # Unis_start Excel dosyasını oku
        df_unis = pd.read_excel(unis_start_path)
        
        # Üniversite detaylarını içeren sözlük oluştur
        university_details = {}
        
        for _, row in df_unis.iterrows():
            uni_name = row.get('Üniversite İsmi', None)
            if uni_name and isinstance(uni_name, str):
                university_details[uni_name] = {
                    "sehir": row.get('sehir', ''),
                    "logo": row.get('logo', ''),
                    "banner": row.get('banner', ''),
                    "universite_Turu": row.get('üniversite Türü', '')
                }
        
        print(f"Toplam {len(university_details)} üniversite detayı yüklendi.")
        return university_details
    except Exception as e:
        print(f"Üniversite detayları yüklenirken hata oluştu: {e}")
        return {}

def excel_to_hierarchical_json(excel_file, puan_turu, university_details):
    """
    Excel dosyasını hiyerarşik JSON'a dönüştürür ve üniversite detaylarını ekler.
    """
    try:
        # Excel dosyasını oku
        df = pd.read_excel(excel_file)
        
        # Sütun isimlerini ve ilk birkaç satırı göster - debug için
        print(f"Sütun isimleri: {df.columns.tolist()}")
        if not df.empty:
            sample_row = df.iloc[0]
            print(f"İlk satır örneği (bazı alanlar):")
            print(f"  Yerleşen son kişinin Diploma notu: {sample_row.get('Yerleşen son kişinin Diploma notu', 'YOK')}")
            print(f"  Ortalama OBP: {sample_row.get('Ortalama OBP', 'YOK')}")
            print(f"  TYT Temel Matematik: {sample_row.get('TYT Temel Matematik', 'YOK')}")
            print(f"  TYT Türkçe: {sample_row.get('TYT Türkçe', 'YOK')}")
        
        # Üniversite, fakülte ve bölüm bazında sıralama için DataFrame'i düzenle
        df = df.sort_values(by=['Üniversite İsmi', 'fakülte', 'bolum'])
        
        # Ana veri yapısı
        universities = {}
        
        # Veri çerçevesini satır satır işleyelim
        for index, row in df.iterrows():
            try:
                uni_name = row['Üniversite İsmi']
                faculty_name = row['fakülte']
                department_name = row['bolum']
                
                # Excel dosyasında sütun adları farklı olabilir - TYT için özel kontroller
                diploma_notu = row.get('Yerleşen son kişinin Diploma notu', '')
                ortalama_obp = row.get('Ortalama OBP', '')
                
                # TYT verileri için sütun isimleri kontrol et - farklı yazımlar için de dene
                tyt_mat = row.get('TYT Temel Matematik', None)
                if tyt_mat is None or pd.isna(tyt_mat):
                    tyt_mat = row.get('TYT Matematik', '')
                
                tyt_fen = row.get('TYT Fen Bilimleri', None)
                if tyt_fen is None or pd.isna(tyt_fen):
                    tyt_fen = row.get('TYT Fen', '')
                
                tyt_turkce = row.get('TYT Türkçe', '')
                tyt_sosyal = row.get('TYT Sosyal Bilimler', None)
                if tyt_sosyal is None or pd.isna(tyt_sosyal):
                    tyt_sosyal = row.get('TYT Sosyal', '')
                
                # İlk 5 satır için debug bilgisi yazdır
                if index < 5 and puan_turu.upper() == 'TYT':
                    print(f"\nTYT Satır {index} - Alanlar:")
                    print(f"  Diploma notu: {diploma_notu}")
                    print(f"  OBP: {ortalama_obp}")
                    print(f"  TYT Mat: {tyt_mat}")
                    print(f"  TYT Fen: {tyt_fen}")
                    print(f"  TYT Türkçe: {tyt_turkce}")
                    print(f"  TYT Sosyal: {tyt_sosyal}")
                
                # Dizileri doğru şekilde ayrıştır
                try:
                    son4_yerlesen = eval(row['son4_yerleşen']) if isinstance(row['son4_yerleşen'], str) else row['son4_yerleşen']
                except:
                    son4_yerlesen = row['son4_yerleşen']
                    
                try:
                    son4_siralama = eval(row['son4_sıralama']) if isinstance(row['son4_sıralama'], str) else row['son4_sıralama']
                except:
                    son4_siralama = row['son4_sıralama']
                    
                try:
                    son4_puan = eval(row['son4_puan']) if isinstance(row['son4_puan'], str) else row['son4_puan']
                except:
                    son4_puan = row['son4_puan']
                
                # Öğretim üyesi sayılarını doğru şekilde ele al
                try:
                    profesor_val = row.get('Profesör', None)
                    if pd.isna(profesor_val) or profesor_val == '---' or profesor_val is None:
                        profesor = 0
                    elif isinstance(profesor_val, (int, float)) and not pd.isna(profesor_val):
                        profesor = int(profesor_val)
                    elif isinstance(profesor_val, str):
                        # String içindeki sayıyı çıkar
                        if profesor_val.isdigit():
                            profesor = int(profesor_val)
                        else:
                            # Sayı olmayan karakterleri temizle
                            cleaned = ''.join(c for c in profesor_val if c.isdigit())
                            profesor = int(cleaned) if cleaned else 0
                    else:
                        profesor = 0
                except Exception:
                    profesor = 0
                
                # Doçent değerini işle
                try:
                    docent_val = row.get('Doçent', None)
                    if pd.isna(docent_val) or docent_val == '---' or docent_val is None:
                        docent = 0
                    elif isinstance(docent_val, (int, float)) and not pd.isna(docent_val):
                        docent = int(docent_val)
                    elif isinstance(docent_val, str):
                        # String içindeki sayıyı çıkar
                        if docent_val.isdigit():
                            docent = int(docent_val)
                        else:
                            # Sayı olmayan karakterleri temizle
                            cleaned = ''.join(c for c in docent_val if c.isdigit())
                            docent = int(cleaned) if cleaned else 0
                    else:
                        docent = 0
                except Exception:
                    docent = 0
                
                # Doktora değerini işle
                try:
                    doktora_val = row.get('Doktora', None)
                    if pd.isna(doktora_val) or doktora_val == '---' or doktora_val is None:
                        doktora = 0
                    elif isinstance(doktora_val, (int, float)) and not pd.isna(doktora_val):
                        doktora = int(doktora_val)
                    elif isinstance(doktora_val, str):
                        # String içindeki sayıyı çıkar
                        if doktora_val.isdigit():
                            doktora = int(doktora_val)
                        else:
                            # Sayı olmayan karakterleri temizle
                            cleaned = ''.join(c for c in doktora_val if c.isdigit())
                            doktora = int(cleaned) if cleaned else 0
                    else:
                        doktora = 0
                except Exception:
                    doktora = 0
                
                # Toplam değerini işle
                try:
                    toplam_val = row.get('Toplam Öğretim Görvelisi', None)
                    if pd.isna(toplam_val) or toplam_val == '---' or toplam_val is None:
                        toplam = profesor + docent + doktora  # Diğerlerinin toplamını kullan
                    elif isinstance(toplam_val, (int, float)) and not pd.isna(toplam_val):
                        toplam = int(toplam_val)
                    elif isinstance(toplam_val, str):
                        # String içindeki sayıyı çıkar
                        if toplam_val.isdigit():
                            toplam = int(toplam_val)
                        else:
                            # Sayı olmayan karakterleri temizle
                            cleaned = ''.join(c for c in toplam_val if c.isdigit())
                            if cleaned:
                                toplam = int(cleaned)
                            else:
                                toplam = profesor + docent + doktora  # Temiz bir sayı yoksa diğerlerinin toplamını kullan
                    else:
                        toplam = profesor + docent + doktora  # Veri tipi bilinmiyorsa diğerlerinin toplamını kullan
                except Exception:
                    toplam = profesor + docent + doktora
                
                # Eğer toplam için geçerli bir değer yoksa veya diğerlerinin toplamından küçükse, toplam değerini hesapla
                if toplam < profesor + docent + doktora:
                    toplam = profesor + docent + doktora
                
                # Puan türüne göre TYT ve AYT alanlarını ekle
                yerlesenlerin_yks_net_ortalamalari = {}
                
                # TYT alanları ekle - yukarıda kontrol edilen değerleri kullan
                tyt_fields = {
                    "TYT_Temel_Matematik": str(tyt_mat) if not pd.isna(tyt_mat) else '',
                    "TYT_Fen_Bilimleri": str(tyt_fen) if not pd.isna(tyt_fen) else '',
                    "TYT_Turkce": str(tyt_turkce) if not pd.isna(tyt_turkce) else '',
                    "TYT_Sosyal_Bilimler": str(tyt_sosyal) if not pd.isna(tyt_sosyal) else ''
                }
                yerlesenlerin_yks_net_ortalamalari.update(tyt_fields)
                
                # TYT haricindeki puan türleri için AYT alanlarını ekle
                puan_turu_value = row.get('puan', '').upper() if row.get('puan') else puan_turu.upper()
                
                if puan_turu_value == 'SAY':
                    yerlesenlerin_yks_net_ortalamalari.update({
                        "AYT_Sayisal": {
                            "Matematik": row.get('AYT Matematik', ''),
                            "Fizik": row.get('AYT Fizik', ''),
                            "Kimya": row.get('AYT Kimya', ''),
                            "Biyoloji": row.get('AYT Biyoloji', '')
                        }
                    })
                elif puan_turu_value == 'EA':
                    yerlesenlerin_yks_net_ortalamalari.update({
                        "AYT_Esit_Agirlik": {
                            "Turk_Dili_ve_Edebiyati": row.get('AYT Türk Dili ve Edebiyatı', ''),
                            "Matematik": row.get('AYT Matematik', ''),
                            "Cografya_1": row.get('AYT Coğrafya-1', ''),
                            "Tarih_1": row.get('AYT Tarih-1', '')
                        }
                    })
                elif puan_turu_value == 'SÖZ':
                    yerlesenlerin_yks_net_ortalamalari.update({
                        "AYT_Sozel": {
                            "Turk_Dili_ve_Edebiyati": row.get('AYT Türk Dili ve Edebiyatı', ''),
                            "Tarih_1": row.get('AYT Tarih-1', ''),
                            "Tarih_2": row.get('AYT Tarih-2', ''),
                            "Cografya_1": row.get('AYT Coğrafya-1', ''),
                            "Cografya_2": row.get('AYT Coğrafya-2', ''),
                            "Felsefe_Grubu": row.get('AYT Felsefe Grubu', ''),
                            "Din_Kulturu_ve_Ahlak_Bilgisi": row.get('AYT Din Kültürü ve Ahlak Bilgisi', '')
                        }
                    })
                elif puan_turu_value == 'DİL':
                    yerlesenlerin_yks_net_ortalamalari.update({
                        "AYT_Yabanci_Dil": {
                            "dil1": row.get('AYT Yabancı Dil 1', ''),
                            "dil2": row.get('AYT Yabancı Dil 2', ''),
                            "dil3": row.get('AYT Yabancı Dil 3', ''),
                            "dil4": row.get('AYT Yabancı Dil 4', ''),
                            "dil5": row.get('AYT Yabancı Dil 5', '')
                        }
                    })
                
                # Bölüm detayları oluştur
                bolum_detaylari = {
                    "son4_kontenjan": row.get('son4_kont', ''),
                    "son4_yerlesen": son4_yerlesen,
                    "son4_siralama": son4_siralama,
                    "son4_puan": son4_puan,
                    "YOP_Kodu": row.get('yop', ''),
                    "Yerlesen_son_kisinin_Diploma_notu": str(diploma_notu) if not pd.isna(diploma_notu) else '',
                    "Ortalama_OBP": str(ortalama_obp) if not pd.isna(ortalama_obp) else '',
                    "Yerlesenlerin_YKS_Net_Ortalamalari": yerlesenlerin_yks_net_ortalamalari,
                    "Ogretim_Uyesi_Sayisi_ve_Unvan_Dagilimi": {
                        "Profesor": profesor,
                        "Docent": docent,
                        "Doktora": doktora,
                        "Toplam_Ogretim_Gorevlisi": toplam
                    }
                }
                
                # Sıralama verilerini güvenli bir şekilde kullan
                siralama_2024 = son4_siralama[0] if isinstance(son4_siralama, list) and len(son4_siralama) > 0 else None
                siralama_2023 = son4_siralama[1] if isinstance(son4_siralama, list) and len(son4_siralama) > 1 else None
                siralama_2022 = son4_siralama[2] if isinstance(son4_siralama, list) and len(son4_siralama) > 2 else None
                
                # Bölüm bilgileri
                bolum_bilgileri = {
                    "bolum_ismi": department_name,
                    "siralama_2024": siralama_2024,
                    "siralama_2023": siralama_2023,
                    "siralama_2022": siralama_2022,
                    "bursluluk_durumu": row.get('burs', ''),
                    "ogrenim_turu": row.get('type', ''),
                    "doluluk": row.get('doluluk', ''),
                    "puan_turu": row.get('puan', ''),
                    "bolum_detaylari": bolum_detaylari
                }
                
                # Üniversite detaylarını al (varsa)
                uni_details = university_details.get(uni_name, {
                    "sehir": "",
                    "logo": "",
                    "banner": "",
                    "universite_Turu": ""
                })
                
                # Üniversite bilgileri
                if uni_name not in universities:
                    universities[uni_name] = {
                        "Universite_Ismi": uni_name,
                        "sehir": uni_details["sehir"],
                        "logo": uni_details["logo"],
                        "banner": uni_details["banner"],
                        "universite_Turu": uni_details["universite_Turu"],
                        "fakulteler": {}
                    }
                
                # Fakülte bilgileri
                if faculty_name not in universities[uni_name]["fakulteler"]:
                    universities[uni_name]["fakulteler"][faculty_name] = {
                        "fakulte_ismi": faculty_name,
                        "bolumler": {}
                    }
                
                # Bölümü fakülteye ekle
                universities[uni_name]["fakulteler"][faculty_name]["bolumler"][department_name] = bolum_bilgileri
                
            except Exception as e:
                print(f"Satır işlenirken hata oluştu (satır {index}): {e}")
                continue
        
        # Sözlük yapısını sıralı listeye dönüştür
        universities_list = []
        # Üniversite adına göre sırala
        for uni_name in sorted(universities.keys()):
            uni_data = universities[uni_name]
            fakulteler_list = []
            
            # Fakülte adına göre sırala
            for faculty_name in sorted(uni_data["fakulteler"].keys()):
                faculty_data = uni_data["fakulteler"][faculty_name]
                bolumler_list = []
                
                # Bölüm adına göre sırala
                for department_name in sorted(faculty_data["bolumler"].keys()):
                    bolumler_list.append(faculty_data["bolumler"][department_name])
                
                fakulteler_list.append({
                    "fakulte_ismi": faculty_name,
                    "bolumler": bolumler_list
                })
            
            uni_data["fakulteler"] = fakulteler_list
            universities_list.append(uni_data)
        
        # Son JSON yapısı
        final_json = {
            "Universiteler": universities_list
        }
        
        return final_json
    except Exception as e:
        print(f"Excel dosyası işlenirken hata oluştu: {e}")
        # Boş bir JSON döndür
        return {"Universiteler": []}

def save_json(data, output_file):
    """
    JSON verilerini dosyaya kaydeder.
    """
    try:
        # Dizin yoksa oluştur
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        
        return True
    except Exception as e:
        print(f"JSON dosyası kaydedilirken hata oluştu: {e}")
        return False

def create_json_files():
    """
    Tüm Excel dosyalarını işleyerek JSON dosyalarını sıfırdan oluşturur.
    """
    print("JSON dosyaları sıfırdan oluşturuluyor...")
    
    # Üniversite detaylarını yükle
    university_details = get_university_details("data/unis_details_with_banners.xlsx")
    
    # Excel dosya yolları ve puan türleri
    excel_files = {
        "say": "unis_details_say_deneme.xlsx",
        "ea": "unis_details_ea_deneme.xlsx",
        "söz": "unis_details_söz_deneme.xlsx",
        "dil": "unis_details_dil_deneme.xlsx",
        "tyt": "unis_details_tyt_deneme.xlsx"
    }
    
    # Eğer dosyalar data/finished klasöründe ise alternatif yolları kontrol et
    for puan_turu, file_path in list(excel_files.items()):
        alt_path = f"data/finished/unis_details_{puan_turu}_rj.xlsx"
        if not os.path.exists(file_path) and os.path.exists(alt_path):
            excel_files[puan_turu] = alt_path
            print(f"Alternatif yol kullanılıyor: {alt_path}")
    
    # final klasörünü temizle veya oluştur
    final_dir = "json"
    if os.path.exists(final_dir):
        # İsteğe bağlı: Mevcut dosyaları sil
        for file in os.listdir(final_dir):
            if file.endswith(".json"):
                try:
                    os.remove(os.path.join(final_dir, file))
                    print(f"{file} dosyası silindi.")
                except:
                    pass
    else:
        # Klasör yoksa oluştur
        os.makedirs(final_dir)
        print(f"'{final_dir}' klasörü oluşturuldu.")
    
    # Her dosya için dönüşümü gerçekleştir
    for puan_turu, excel_file in excel_files.items():
        try:
            print(f"\n{puan_turu.upper()} puan türü işleniyor... ({excel_file})")
            
            # Dosyanın var olup olmadığını kontrol et
            if not os.path.exists(excel_file):
                print(f"UYARI: {excel_file} dosyası bulunamadı. Bu puan türü atlanıyor.")
                continue
                
            # Dönüşümü gerçekleştir
            json_data = excel_to_hierarchical_json(excel_file, puan_turu, university_details)
            
            # JSON dosyasının çıktı yolu
            output_file = f"{final_dir}/{puan_turu}.json"
            
            # JSON'ı kaydet
            if save_json(json_data, output_file):
                print(f"Dönüşüm tamamlandı. Hiyerarşik JSON dosyası: {output_file}")
            else:
                print(f"HATA: {output_file} dosyası kaydedilemedi.")
                
        except Exception as e:
            print(f"HATA: {excel_file} dosyası işlenirken bir sorun oluştu: {e}")

if __name__ == "__main__":
    # Sıfırdan JSON dosyaları oluştur
    create_json_files()
    print("\nTüm işlemler tamamlandı.")