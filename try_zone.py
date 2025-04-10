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

def normalize_yop_code(yop_code):
    """
    YÖP kodlarını standart formata getirir.
    Bazı Excel dosyalarında int olarak, bazılarında string olarak olabilir.
    """
    if pd.isna(yop_code) or yop_code is None:
        return ""
    
    # Int veya float ise stringe çevir
    if isinstance(yop_code, (int, float)):
        return str(int(yop_code))
    
    # String ise boşlukları temizle
    if isinstance(yop_code, str):
        return yop_code.strip()
    
    return ""

def get_net_details():
    """
    Yerleşen son kişilerin net bilgilerini YÖP kodlarına göre içeren bir sözlük oluşturur.
    """
    try:
        # Tüm DataFrame'leri yükle
        net_details = {}
        
        # Dosya isimleri ve puan türleri
        df_files = {
            'say_df.xlsx': 'SAY',
            'soz_df.xlsx': 'SÖZ',
            'ea_df.xlsx': 'EA',
            'dil_df.xlsx': 'DİL',
            'tyt_df.xlsx': 'TYT'
        }
        
        # YÖP kodu bulunan toplam satır sayısı
        total_yop_rows = 0
        
        # Her dosyayı işle
        for file_name, puan_turu in df_files.items():
            try:
                # Olası dosya yollarını dene
                df = None
                loaded_path = None
                
                for path in [file_name, f"data/{file_name}"]:
                    try:
                        if os.path.exists(path):
                            df = pd.read_excel(path)
                            loaded_path = path
                            break
                    except:
                        continue
                
                if df is None:
                    print(f"{file_name} dosyası bulunamadı, atlanıyor.")
                    continue
                
                print(f"{file_name} dosyası '{loaded_path}' yolundan yüklendi. Satır sayısı: {len(df)}")
                
                # Sütun isimlerini kontrol et
                print(f"{file_name} sütunları: {df.columns.tolist()}")
                
                # Her satır için YÖP kodu ve netleri eşleştir
                yop_count = 0
                
                for _, row in df.iterrows():
                    # YÖP kodunu standartlaştır
                    yop_code = normalize_yop_code(row.get('yop', ''))
                    
                    # Geçerli bir YÖP kodu varsa
                    if yop_code:
                        yop_count += 1
                        
                        # Puan türüne göre net bilgilerini al
                        net_dict = {
                            'type': row.get('type', ''),
                            'bolum': row.get('bolum', '')
                        }
                        
                        # TYT netleri - tüm puan türlerinde ortak
                        for field in ['TYT Temel Matematik', 'TYT Fen Bilimleri', 'TYT Türkçe', 'TYT Sosyal Bilimler']:
                            if field in df.columns:
                                value = row.get(field, '')
                                if not pd.isna(value) and value != '':
                                    # Virgüllü sayıları noktalı formata çevir
                                    if isinstance(value, str) and ',' in value:
                                        value = value.replace(',', '.')
                                    net_dict[field] = value
                        
                        # Puan türüne göre AYT netleri
                        if puan_turu == 'SAY':
                            for field in ['AYT Matematik', 'AYT Fizik', 'AYT Kimya', 'AYT Biyoloji']:
                                if field in df.columns:
                                    value = row.get(field, '')
                                    if not pd.isna(value) and value != '':
                                        if isinstance(value, str) and ',' in value:
                                            value = value.replace(',', '.')
                                        net_dict[field] = value
                        elif puan_turu == 'SÖZ':
                            for field in ['AYT Türk Dili ve Edebiyatı', 'AYT Coğrafya-1', 'AYT Coğrafya-2', 
                                         'AYT Tarih-1', 'AYT Tarih-2', 'AYT Felsefe Grubu', 
                                         'AYT Din Kültürü ve Ahlak Bilgisi']:
                                if field in df.columns:
                                    value = row.get(field, '')
                                    if not pd.isna(value) and value != '':
                                        if isinstance(value, str) and ',' in value:
                                            value = value.replace(',', '.')
                                        net_dict[field] = value
                        elif puan_turu == 'EA':
                            for field in ['AYT Türk Dili ve Edebiyatı', 'AYT Matematik', 
                                         'AYT Coğrafya-1', 'AYT Tarih-1']:
                                if field in df.columns:
                                    value = row.get(field, '')
                                    if not pd.isna(value) and value != '':
                                        if isinstance(value, str) and ',' in value:
                                            value = value.replace(',', '.')
                                        net_dict[field] = value
                        elif puan_turu == 'DİL':
                            for field in ['AYT Yabancı Dil 1']:
                                if field in df.columns:
                                    value = row.get(field, '')
                                    if not pd.isna(value) and value != '':
                                        if isinstance(value, str) and ',' in value:
                                            value = value.replace(',', '.')
                                        net_dict[field] = value
                        
                        # Eğer aynı YÖP kodu için önceki kayıtlar varsa, atla
                        # İlk bulunan kaydı kullan
                        if yop_code not in net_details:
                            net_details[yop_code] = net_dict
                
                print(f"{file_name} - YÖP kodu bulunan satır sayısı: {yop_count}")
                total_yop_rows += yop_count
            
            except Exception as e:
                print(f"{file_name} dosyası işlenirken hata oluştu: {e}")
        
        print(f"Toplam {len(net_details)} YÖP kodu için net bilgisi yüklendi (toplam {total_yop_rows} satırdan).")
        
        # İlk 5 YÖP kodu ve net bilgilerini göster - debug amaçlı
        if net_details:
            print("\nÖRNEK YÖP KODLARI VE NET BİLGİLERİ (ilk 5):")
            count = 0
            for yop, details in net_details.items():
                if count < 5:
                    print(f"YÖP: {yop}")
                    net_values = {k: v for k, v in details.items() if k not in ['type', 'bolum']}
                    print(f"  Bölüm: {details.get('bolum', 'Belirtilmemiş')}")
                    print(f"  Net bilgileri: {net_values}")
                    count += 1
                else:
                    break
        
        return net_details
    except Exception as e:
        print(f"Net bilgileri yüklenirken hata oluştu: {e}")
        return {}

def excel_to_hierarchical_json(excel_file, puan_turu, university_details, net_details):
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
            print(f"  YÖP Kodu: {sample_row.get('yop', 'YOK')}")
            print(f"  Yerleşen son kişinin Diploma notu: {sample_row.get('Yerleşen son kişinin Diploma notu', 'YOK')}")
            print(f"  Ortalama OBP: {sample_row.get('Ortalama OBP', 'YOK')}")
        
        # Üniversite, fakülte ve bölüm bazında sıralama için DataFrame'i düzenle
        df = df.sort_values(by=['Üniversite İsmi', 'fakülte', 'bolum'])
        
        # Ana veri yapısı
        universiteler = []
        
        # Şu anki işlenen üniversite ve fakülte
        current_uni = None
        current_fakulte = None
        current_uni_dict = None
        
        # YÖP eşleşme sayacı
        yop_match_count = 0
        net_info_count = 0
        
        # Veri çerçevesini satır satır işleyelim
        for index, row in df.iterrows():
            try:
                uni_name = row['Üniversite İsmi']
                faculty_name = row['fakülte']
                department_name = row['bolum']
                
                # YÖP kodunu al ve normalize et
                yop_code = normalize_yop_code(row.get('yop', ''))
                
                # Excel dosyasında sütun adları farklı olabilir
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
                
                # Yerleşenlerin YKS net ortalamaları
                yerlesenlerin_yks_net_ortalamalari = {
                    "TYT_Temel_Matematik": str(tyt_mat) if not pd.isna(tyt_mat) else '',
                    "TYT_Fen_Bilimleri": str(tyt_fen) if not pd.isna(tyt_fen) else '',
                    "TYT_Turkce": str(tyt_turkce) if not pd.isna(tyt_turkce) else '',
                    "TYT_Sosyal_Bilimler": str(tyt_sosyal) if not pd.isna(tyt_sosyal) else ''
                }
                
                # Puan türüne göre AYT alanlarını ekle
                puan_turu_value = row.get('puan', '').upper() if row.get('puan') else puan_turu.upper()
                
                if puan_turu_value == 'SAY':
                    yerlesenlerin_yks_net_ortalamalari.update({
                        "AYT_Matematik": row.get('AYT Matematik', ''),
                        "AYT_Fizik": row.get('AYT Fizik', ''),
                        "AYT_Kimya": row.get('AYT Kimya', ''),
                        "AYT_Biyoloji": row.get('AYT Biyoloji', '')
                    })
                elif puan_turu_value == 'EA':
                    yerlesenlerin_yks_net_ortalamalari.update({
                        "AYT_Turk_Dili_ve_Edebiyati": row.get('AYT Türk Dili ve Edebiyatı', ''),
                        "AYT_Matematik": row.get('AYT Matematik', ''),
                        "AYT_Cografya_1": row.get('AYT Coğrafya-1', ''),
                        "AYT_Tarih_1": row.get('AYT Tarih-1', '')
                    })
                elif puan_turu_value == 'SÖZ':
                    yerlesenlerin_yks_net_ortalamalari.update({
                        "AYT_Turk_Dili_ve_Edebiyati": row.get('AYT Türk Dili ve Edebiyatı', ''),
                        "AYT_Tarih_1": row.get('AYT Tarih-1', ''),
                        "AYT_Tarih_2": row.get('AYT Tarih-2', ''),
                        "AYT_Cografya_1": row.get('AYT Coğrafya-1', ''),
                        "AYT_Cografya_2": row.get('AYT Coğrafya-2', ''),
                        "AYT_Felsefe_Grubu": row.get('AYT Felsefe Grubu', ''),
                        "AYT_Din_Kulturu_ve_Ahlak_Bilgisi": row.get('AYT Din Kültürü ve Ahlak Bilgisi', '')
                    })
                elif puan_turu_value == 'DİL':
                    yerlesenlerin_yks_net_ortalamalari.update({
                        "AYT_Yabanci_Dil_1": row.get('AYT Yabancı Dil 1', ''),
                        "AYT_Yabanci_Dil_2": row.get('AYT Yabancı Dil 2', ''),
                        "AYT_Yabanci_Dil_3": row.get('AYT Yabancı Dil 3', ''),
                        "AYT_Yabanci_Dil_4": row.get('AYT Yabancı Dil 4', ''),
                        "AYT_Yabanci_Dil_5": row.get('AYT Yabancı Dil 5', '')
                    })
                
                # Yerleşen son kişinin netleri - YÖP koduna göre alıyoruz
                yerlesen_son_kisinin_netleri = {}
                
                # YÖP kodu eşleşmesi kontrolü
                if yop_code and yop_code in net_details:
                    yop_match_count += 1
                    
                    # Type ve bolum dışındaki tüm alanları al, '-' değerleri filtrele
                    net_bilgileri = {}
                    for k, v in net_details[yop_code].items():
                        if k not in ['type', 'bolum'] and not pd.isna(v) and v != '-' and v != '':
                            net_bilgileri[k] = v
                    
                    if net_bilgileri:
                        yerlesen_son_kisinin_netleri = net_bilgileri
                        net_info_count += 1
                
                # Debug için - ilk 3 eşleşen YÖP kodu için detayları göster
                if index < 3 and yop_code in net_details:
                    print(f"\nDEBUG - Satır {index}, YÖP: {yop_code}")
                    print(f"  net_details içindeki değer: {net_details[yop_code]}")
                    print(f"  Filtrelenmiş net bilgileri: {yerlesen_son_kisinin_netleri}")
                
                # Bölüm detayları
                bolum_detaylari = {
                    "son_4_yil_toplam_kontenjan": row.get('son4_kont', ''),
                    "son_4_yil_yerlesen": son4_yerlesen,
                    "son_4_yil_siralama": son4_siralama,
                    "son_4_yil_puan": son4_puan,
                    "YOP_Kodu": yop_code,
                    "Yerlesen_son_kisinin_Diploma_notu": str(diploma_notu) if not pd.isna(diploma_notu) else '',
                    "Ortalama_OBP": str(ortalama_obp) if not pd.isna(ortalama_obp) else '',
                    "Yerlesen_son_kisinin_netleri": yerlesen_son_kisinin_netleri,
                    "Yerlesenlerin_YKS_Net_Ortalamalari": yerlesenlerin_yks_net_ortalamalari,
                    "Ogretim_Uyesi_Sayisi_ve_Unvan_Dagilimi": {
                        "Profesor": profesor,
                        "Docent": docent,
                        "Doktora": doktora,
                        "Toplam_Ogretim_Gorevlisi": toplam
                    }
                }
                
                # Bölüm bilgileri
                siralama_2024 = son4_siralama[0] if isinstance(son4_siralama, list) and len(son4_siralama) > 0 else None
                siralama_2023 = son4_siralama[1] if isinstance(son4_siralama, list) and len(son4_siralama) > 1 else None
                siralama_2022 = son4_siralama[2] if isinstance(son4_siralama, list) and len(son4_siralama) > 2 else None
                
                bolum = {
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
                
                # Üniversite detaylarını al
                uni_details = university_details.get(uni_name, {
                    "sehir": "",
                    "logo": "",
                    "banner": "",
                    "universite_Turu": ""
                })
                
                # Yeni üniversite ise yeni ekle
                if current_uni != uni_name:
                    if current_uni_dict:
                        universiteler.append(current_uni_dict)
                    
                    current_uni = uni_name
                    current_uni_dict = {
                        "Universite_Ismi": uni_name,
                        "sehir": uni_details["sehir"],
                        "logo": uni_details["logo"],
                        "banner": uni_details["banner"],
                        "universite_Turu": uni_details["universite_Turu"],
                        "fakulteler": []
                    }
                    current_fakulte = None
                
                # Yeni fakülte ise yeni ekle
                if current_fakulte != faculty_name:
                    current_fakulte = faculty_name
                    current_fakulte_dict = {
                        "fakulte_ismi": faculty_name,
                        "bolumler": []
                    }
                    current_uni_dict["fakulteler"].append(current_fakulte_dict)
                
                # Bölümü fakülteye ekle
                current_uni_dict["fakulteler"][-1]["bolumler"].append(bolum)
                
            except Exception as e:
                print(f"Satır işlenirken hata oluştu (satır {index}): {e}")
                continue
        
        # Son üniversiteyi de ekle
        if current_uni_dict:
            universiteler.append(current_uni_dict)
        
        print(f"\nİşlenen veri istatistikleri:")
        print(f"  Toplam satır sayısı: {len(df)}")
        print(f"  YÖP kodu eşleşen bölüm sayısı: {yop_match_count}")
        print(f"  Net bilgisi eklenen bölüm sayısı: {net_info_count}")
        
        # Son JSON yapısı
        final_json = {
            "Universiteler": universiteler
        }
        
        return final_json, yop_match_count, net_info_count
    except Exception as e:
        print(f"Excel dosyası işlenirken hata oluştu: {e}")
        # Boş bir JSON döndür
        return {"Universiteler": []}, 0, 0

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
    university_details = {}
    unis_detail_paths = [
        "data/unis_details_with_banners.xlsx",
        "unis_details_with_banners.xlsx"
    ]
    
    for path in unis_detail_paths:
        try:
            university_details = get_university_details(path)
            if university_details:
                print(f"Üniversite detayları '{path}' dosyasından yüklendi.")
                break
        except Exception as e:
            print(f"{path} dosyasından üniversite detayları yüklenemedi: {e}")
    
    if not university_details:
        print("Uyarı: Üniversite detayları yüklenemedi. Boş detaylar ile devam edilecek.")
    
    # Yerleşen son kişilerin net bilgilerini YÖP koduna göre yükle
    net_details = get_net_details()
    
    # Excel dosya yolları ve puan türleri
    excel_files = {
        "say": "unis_details_say_deneme.xlsx",
        "ea": "unis_details_ea_deneme.xlsx",
        "söz": "unis_details_söz_deneme.xlsx",
        "dil": "unis_details_dil_deneme.xlsx",
        "tyt": "unis_details_tyt_deneme.xlsx"
    }
    
    # Her puan türü için olası dosya yollarını kontrol et
    for puan_turu, file_path in list(excel_files.items()):
        alt_paths = [
            file_path,
            f"data/{file_path}",
            f"data/finished/unis_details_{puan_turu}_rj.xlsx",
            f"unis_details_{puan_turu}.xlsx"
        ]
        
        file_found = False
        for alt_path in alt_paths:
            if os.path.exists(alt_path):
                excel_files[puan_turu] = alt_path
                print(f"{puan_turu.upper()} için '{alt_path}' kullanılacak.")
                file_found = True
                break
        
        if not file_found:
            print(f"Uyarı: {puan_turu.upper()} için hiçbir Excel dosyası bulunamadı.")
    
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
        
    # Toplam istatistikler
    total_matches = 0
    total_net_info = 0
    total_processed = 0
    
    # Her dosya için dönüşümü gerçekleştir
    for puan_turu, excel_file in excel_files.items():
        try:
            print(f"\n{puan_turu.upper()} puan türü işleniyor... ({excel_file})")
            
            # Dosyanın var olup olmadığını kontrol et
            if not os.path.exists(excel_file):
                print(f"UYARI: {excel_file} dosyası bulunamadı. Bu puan türü atlanıyor.")
                continue
                
            # Dönüşümü gerçekleştir
            json_data, yop_match_count, net_info_count = excel_to_hierarchical_json(excel_file, puan_turu, university_details, net_details)
            
            # İstatistikleri güncelle
            total_matches += yop_match_count
            total_net_info += net_info_count
            
            if isinstance(json_data, dict) and "Universiteler" in json_data:
                total_processed += sum(
                    sum(len(fakulte["bolumler"]) for fakulte in uni["fakulteler"])
                    for uni in json_data["Universiteler"]
                )
            
            # JSON dosyasının çıktı yolu
            output_file = f"{final_dir}/{puan_turu}.json"
            
            # JSON'ı kaydet
            if save_json(json_data, output_file):
                print(f"Dönüşüm tamamlandı. Hiyerarşik JSON dosyası: {output_file}")
                print(f"  Bu puan türünde {net_info_count} bölüme yerleşen son kişinin net bilgisi eklendi.")
                print(f"  Toplam {yop_match_count} YÖP kodu eşleşmesi bulundu.")
            else:
                print(f"HATA: {output_file} dosyası kaydedilemedi.")
                
        except Exception as e:
            print(f"HATA: {excel_file} dosyası işlenirken bir sorun oluştu: {e}")
    
    print(f"\nGenel İstatistikler:")
    print(f"  Toplam işlenen bölüm sayısı: {total_processed}")
    print(f"  Toplam YÖP kodu eşleşmesi: {total_matches}")
    print(f"  Toplam net bilgisi eklenen bölüm sayısı: {total_net_info}")

if __name__ == "__main__":
    # Sıfırdan JSON dosyaları oluştur
    create_json_files()
    print("\nTüm işlemler tamamlandı.")