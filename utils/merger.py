import json
import os
from collections import defaultdict

def load_json_file(file_path):
    """
    JSON dosyasını yükler ve içeriğini döndürür.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"HATA: {file_path} dosyası yüklenirken bir sorun oluştu: {e}")
        return None

def merge_universities(json_files):
    """
    Farklı puan türlerindeki üniversite verilerini birleştirir.
    """
    print("Üniversite verileri birleştiriliyor...")
    
    # Tüm üniversiteleri ve fakültelerini saklayacak ana sözlük
    merged_universities = {}
    
    # Her JSON dosyasını işle
    for json_file, data in json_files.items():
        puan_turu = os.path.basename(json_file).split('.')[0].upper()
        print(f"\n{puan_turu} verileri işleniyor...")
        
        if not data or 'Universiteler' not in data:
            print(f"UYARI: {json_file} dosyasında 'Universiteler' alanı bulunamadı. Bu dosya atlanıyor.")
            continue
            
        universities = data['Universiteler']
        
        # Her üniversiteyi işle
        for uni in universities:
            uni_name = uni.get('Universite_Ismi', '')
            if not uni_name:
                continue
                
            # Üniversite ana sözlükte yoksa ekle
            if uni_name not in merged_universities:
                merged_universities[uni_name] = {
                    "Universite_Ismi": uni_name,
                    "sehir": uni.get('sehir', ''),
                    "logo": uni.get('logo', ''),
                    "banner": uni.get('banner', ''),
                    "universite_Turu": uni.get('universite_Turu', ''),
                    "fakulteler": {}
                }
            
            # Fakulteleri işle
            for fakulte in uni.get('fakulteler', []):
                fakulte_name = fakulte.get('fakulte_ismi', '')
                if not fakulte_name:
                    continue
                    
                # Fakülte ana sözlükte yoksa ekle
                if fakulte_name not in merged_universities[uni_name]["fakulteler"]:
                    merged_universities[uni_name]["fakulteler"][fakulte_name] = {
                        "fakulte_ismi": fakulte_name,
                        "bolumler": {}
                    }
                
                # Bölümleri işle
                for bolum in fakulte.get('bolumler', []):
                    bolum_name = bolum.get('bolum_ismi', '')
                    
                    # Bölüm ana sözlükte yoksa veya mevcut bölüm farklı bir puan türüne sahipse ekle
                    # Aynı isimli bölüm ve farklı bursluluk durumları için benzersiz anahtar oluştur
                    puan_turu_bolum = bolum.get('puan_turu', '')
                    bursluluk = bolum.get('bursluluk_durumu', '')
                    ogrenim_turu = bolum.get('ogrenim_turu', '')
                    
                    # Benzersiz anahtar oluştur
                    unique_key = f"{bolum_name}_{puan_turu_bolum}_{bursluluk}_{ogrenim_turu}"
                    
                    # Bölümü ekle
                    merged_universities[uni_name]["fakulteler"][fakulte_name]["bolumler"][unique_key] = bolum
    
    # Sözlük yapısını sıralı listeye dönüştür
    merged_universities_list = []
    
    # Üniversiteleri alfabetik olarak sırala
    for uni_name in sorted(merged_universities.keys()):
        uni_data = merged_universities[uni_name]
        fakulteler_list = []
        
        # Fakülteleri alfabetik olarak sırala
        for fakulte_name in sorted(uni_data["fakulteler"].keys()):
            fakulte_data = uni_data["fakulteler"][fakulte_name]
            bolumler_list = []
            
            # Bölümleri alfabetik olarak sırala
            # Burada unique_key'den gerçek bölüm adına geri dönüyoruz
            for unique_key in sorted(fakulte_data["bolumler"].keys()):
                bolumler_list.append(fakulte_data["bolumler"][unique_key])
            
            fakulteler_list.append({
                "fakulte_ismi": fakulte_name,
                "bolumler": bolumler_list
            })
        
        uni_data["fakulteler"] = fakulteler_list
        merged_universities_list.append(uni_data)
    
    # Son JSON yapısı
    merged_json = {
        "Universiteler": merged_universities_list
    }
    
    return merged_json

def save_merged_json(data, output_file):
    """
    Birleştirilmiş JSON verilerini dosyaya kaydeder.
    """
    try:
        # Dizin yoksa oluştur
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            
        print(f"\nBirleştirilmiş veriler başarıyla kaydedildi: {output_file}")
        return True
    except Exception as e:
        print(f"\nHATA: JSON dosyası kaydedilirken bir sorun oluştu: {e}")
        return False

def merge_json_files():
    """
    final klasöründeki JSON dosyalarını birleştirir ve tek bir dosya oluşturur.
    """
    # JSON dosyalarının bulunduğu klasör
    json_dir = "json"
    
    # Birleştirilmiş JSON dosyasının adı
    output_file = "tum_universiteler.json"
    
    # Klasörlerin varlığını kontrol et
    if not os.path.exists(json_dir):
        print(f"HATA: '{json_dir}' klasörü bulunamadı.")
        return False
    
    # JSON dosyalarını bul
    json_files = {}
    expected_files = ['say.json', 'ea.json', 'söz.json', 'dil.json', 'tyt.json']
    
    for file_name in expected_files:
        file_path = os.path.join(json_dir, file_name)
        if os.path.exists(file_path):
            data = load_json_file(file_path)
            if data:
                json_files[file_path] = data
                print(f"'{file_path}' dosyası yüklendi.")
            else:
                print(f"UYARI: '{file_path}' dosyası yüklenemedi.")
        else:
            print(f"UYARI: '{file_path}' dosyası bulunamadı.")
    
    if not json_files:
        print("HATA: Birleştirilecek JSON dosyası bulunamadı.")
        return False
    
    # Dosyaları birleştir
    merged_data = merge_universities(json_files)
    
    # Birleştirilmiş dosyayı kaydet
    return save_merged_json(merged_data, output_file)

if __name__ == "__main__":
    # Birleştirme işlemini başlat
    print("JSON dosyaları birleştiriliyor...")
    
    if merge_json_files():
        print("\nBirleştirme işlemi başarıyla tamamlandı.")
        
        # İstatistikler
        try:
            with open("tum_universiteler.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                uni_count = len(data.get("Universiteler", []))
                
                # Fakülte ve bölüm sayısını hesapla
                fakulte_count = 0
                bolum_count = 0
                
                for uni in data.get("Universiteler", []):
                    fakulteler = uni.get("fakulteler", [])
                    fakulte_count += len(fakulteler)
                    
                    for fakulte in fakulteler:
                        bolum_count += len(fakulte.get("bolumler", []))
                
                print(f"\nToplam {uni_count} üniversite, {fakulte_count} fakülte ve {bolum_count} bölüm birleştirildi.")
        except Exception as e:
            print(f"İstatistikler hesaplanırken hata oluştu: {e}")
    else:
        print("\nBirleştirme işlemi başarısız oldu.")