import re
import pandas as pd
from collections import Counter

def parse_lotr_script(script_file):
    """
    Lord of the Rings senaryosunu ayrıştıran fonksiyon.
    
    Args:
        script_file: Senaryo dosyasının yolu
    
    Returns:
        dialog_data: Karakter konuşmalarını içeren bir liste
    """
    # Dosyayı oku
    with open(script_file, 'r', encoding='utf-8', errors='ignore') as file:
        script_text = file.read()
    
    # Konuşmaları ayıklamak için regex kalıbı
    # Karakter adı büyük harfle başlar ve diyalog takip eder
    dialog_pattern = r'([A-Z][A-Z\s]+)(?:\s+\(.*?\))?\s*\n((?:.+\n)+)'
    
    dialog_data = []
    
    # Sahne yönergeleri - bunları filtreleyin
    scene_directions = ["DAY", "NIGHT", "MORNING", "EVENING", "DUSK", "DAWN", "AFTERNOON", "LATER", "CONTINUOUS",'BLACK SCREEN']
    
    # Normal konuşmaları işle
    for match in re.finditer(dialog_pattern, script_text):
        character = match.group(1).strip()
        dialog = match.group(2).strip()
        
        # Parantez içindeki sahne yönergelerini kaldır
        dialog = re.sub(r'\(.*?\)', '', dialog)
        
        # Sahne yönergelerini dışarıda bırak
        is_scene_direction = False
        for direction in scene_directions:
            if direction in character:
                is_scene_direction = True
                break
        
        if character and dialog and not is_scene_direction:
            # Karakter adında sadece voice over (V.O.) varsa, onu kaydet
            is_voiceover = False
            if "(V.O.)" in character:
                character = character.replace("(V.O.)", "").strip()
                is_voiceover = True
            
            dialog_data.append({
                'character': character,
                'dialog': dialog,
                'type': 'voiceover' if is_voiceover else 'dialog'
            })
    
    return dialog_data

def analyze_characters(dialog_data):
    """
    Karakterlerin konuşma istatistiklerini analiz eder.
    
    Args:
        dialog_data: parse_lotr_script tarafından üretilen veri
    
    Returns:
        character_stats: Karakter istatistiklerini içeren bir sözlük
    """
    character_stats = {}
    
    for entry in dialog_data:
        character = entry['character']
        
        if character not in character_stats:
            character_stats[character] = {
                'line_count': 0,
                'dialog_count': 0,
                'voiceover_count': 0,
                'total_words': 0
            }
        
        character_stats[character]['line_count'] += 1
        
        if entry['type'] == 'dialog':
            character_stats[character]['dialog_count'] += 1
        elif entry['type'] == 'voiceover':
            character_stats[character]['voiceover_count'] += 1
        
        # Kelime sayısını hesapla
        words = entry['dialog'].split()
        character_stats[character]['total_words'] += len(words)
    
    return character_stats

def analyze_themes(dialog_data, character):
    """
    Belirli bir karakterin konuşmalarındaki temaları analiz eder.
    
    Args:
        dialog_data: parse_lotr_script tarafından üretilen veri
        character: Analiz edilecek karakter adı
    
    Returns:
        theme_counts: Tema sayımlarını içeren bir sözlük
    """
    # Bu karakterin tüm konuşmalarını al
    character_dialogs = [entry['dialog'] for entry in dialog_data 
                        if entry['character'] == character]
    
    # Tüm metni birleştir ve küçük harfe çevir
    all_text = ' '.join(character_dialogs).lower()
    
    # Temalar ve ilgili anahtar kelimeler
    themes = {
        'Güç': ['power', 'ring', 'rule', 'control', 'strength', 'master'],
        'Yolculuk': ['journey', 'road', 'path', 'way', 'travel', 'walk'],
        'Dostluk': ['friend', 'fellowship', 'together', 'companion', 'trust'],
        'Kötülük': ['evil', 'dark', 'shadow', 'enemy', 'mordor', 'sauron'],
        'Umut': ['hope', 'light', 'courage', 'brave', 'faith']
    }
    
    # Tema sayımları
    theme_counts = {}
    
    for theme, keywords in themes.items():
        theme_counts[theme] = 0
        for keyword in keywords:
            # Tam kelime eşleşmesi için düzenli ifade kullan
            pattern = r'\b' + keyword + r'\b'
            matches = re.findall(pattern, all_text)
            theme_counts[theme] += len(matches)
    
    return theme_counts

def find_interactions(dialog_data):
    """
    Karakterler arasındaki etkileşimleri analiz eder.
    
    Args:
        dialog_data: parse_lotr_script tarafından üretilen veri
    
    Returns:
        interactions: Karakter etkileşimlerini içeren bir sözlük
    """
    interactions = {}
    
    for i in range(1, len(dialog_data)):
        prev_character = dialog_data[i-1]['character']
        curr_character = dialog_data[i]['character']
        
        if prev_character != curr_character:
            if prev_character not in interactions:
                interactions[prev_character] = {}
            
            if curr_character not in interactions[prev_character]:
                interactions[prev_character][curr_character] = 0
            
            interactions[prev_character][curr_character] += 1
    
    return interactions

def count_ring_mentions(dialog_data):
    """
    Her karakterin 'Ring' kelimesini kaç kez kullandığını sayar.
    
    Args:
        dialog_data: parse_lotr_script tarafından üretilen veri
    
    Returns:
        ring_mentions: Ring kelimesi sayımlarını içeren bir sözlük
    """
    ring_mentions = {}
    
    for entry in dialog_data:
        character = entry['character']
        dialog = entry['dialog']
        
        # 'ring' ve 'Ring' kelimelerini ara
        matches = re.findall(r'\bring\b|\bRing\b', dialog)
        
        if matches:
            if character not in ring_mentions:
                ring_mentions[character] = 0
            
            ring_mentions[character] += len(matches)
    
    return ring_mentions

def save_to_csv(character_stats, theme_data, interactions, ring_mentions, output_file):
    """
    Analiz sonuçlarını CSV dosyasına kaydeder.
    
    Args:
        character_stats: Karakter istatistikleri
        theme_data: Tema analizi sonuçları
        interactions: Karakter etkileşimleri
        ring_mentions: Ring kelimesi sayımları
        output_file: Çıktı CSV dosyasının adı
    """
    # Karakter istatistiklerini DataFrame'e dönüştür
    character_df = pd.DataFrame.from_dict(character_stats, orient='index')
    character_df.reset_index(inplace=True)
    character_df.rename(columns={'index': 'character'}, inplace=True)
    
    # Theme verilerini düzenle
    theme_rows = []
    for character, themes in theme_data.items():
        row = {'character': character}
        row.update(themes)
        theme_rows.append(row)
    
    theme_df = pd.DataFrame(theme_rows)
    
    # Ring sayımlarını düzenle
    ring_df = pd.DataFrame(list(ring_mentions.items()), columns=['character', 'ring_mentions'])
    
    # Tüm dataframe'leri birleştir
    result_df = character_df.merge(theme_df, on='character', how='left')
    result_df = result_df.merge(ring_df, on='character', how='left')
    
    # Etkileşimleri düzenle
    interaction_rows = []
    for char1 in interactions:
        for char2, count in interactions[char1].items():
            interaction_rows.append({
                'character1': char1,
                'character2': char2,
                'count': count
            })
    
    interaction_df = pd.DataFrame(interaction_rows)
    
    # CSV olarak kaydet
    result_df.to_csv(f"{output_file}_characters.csv", index=False)
    interaction_df.to_csv(f"{output_file}_interactions.csv", index=False)
    
    print(f"Veriler {output_file}_characters.csv ve {output_file}_interactions.csv dosyalarına kaydedildi.")

def main(script_file, output_file):
    """
    Ana fonksiyon - tüm analizleri çalıştırır ve sonuçları görüntüler.
    
    Args:
        script_file: Senaryo dosyasının yolu
        output_file: Çıktı CSV dosyasının adı
    """
    # Senaryoyu ayrıştır
    dialog_data = parse_lotr_script(script_file)
    
    # Karakter istatistiklerini analiz et
    character_stats = analyze_characters(dialog_data)
    
    # Konuşma sayılarına göre sırala
    sorted_characters = sorted(character_stats.items(), 
                              key=lambda x: x[1]['line_count'], 
                              reverse=True)
    
    # İlk 10 karakteri göster
    print("En çok konuşan 10 karakter:")
    for i, (character, stats) in enumerate(sorted_characters[:10]):
        print(f"{i+1}. {character}: {stats['line_count']} satır, {stats['total_words']} kelime")
    
    # Tüm karakterlerin temalarını analiz et
    theme_data = {}
    for character, _ in sorted_characters:
        themes = analyze_themes(dialog_data, character)
        theme_data[character] = themes
    
    # Ana karakterlerin temalarını göster
    main_characters = ['FRODO', 'GANDALF', 'ARAGORN', 'SAM', 'BOROMIR', 'GALADRIEL']
    
    print("\nTema analizi:")
    for character in main_characters:
        if character in theme_data:
            print(f"\n{character} temaları:")
            for theme, count in sorted(theme_data[character].items(), key=lambda x: x[1], reverse=True):
                print(f"  {theme}: {count}")
    
    # Karakter etkileşimlerini bul
    interactions = find_interactions(dialog_data)
    
    # En sık etkileşimleri bul
    top_interactions = []
    for char1 in interactions:
        for char2 in interactions[char1]:
            top_interactions.append((char1, char2, interactions[char1][char2]))
    
    # Etkileşim sayısına göre sırala
    top_interactions.sort(key=lambda x: x[2], reverse=True)
    
    print("\nEn sık karakter etkileşimleri:")
    for char1, char2, count in top_interactions[:10]:
        print(f"{char1} → {char2}: {count}")
    
    # Ring kelimesi kullanımını analiz et
    ring_mentions = count_ring_mentions(dialog_data)
    sorted_mentions = sorted(ring_mentions.items(), key=lambda x: x[1], reverse=True)
    
    print("\n'Ring' kelimesini en çok kullanan karakterler:")
    for character, count in sorted_mentions[:10]:
        print(f"{character}: {count}")
    
    # Verileri CSV'ye kaydet
    save_to_csv(character_stats, theme_data, interactions, ring_mentions, output_file)

if __name__ == "__main__":
    # Kullanım örneği:
    script_file = "data/LordoftheRings1-FOTR.txt"  # Senaryo dosyasının yolu
    output_file = "lotr_analysis"  # Çıktı dosyası adı (uzantısız)
    main(script_file, output_file)