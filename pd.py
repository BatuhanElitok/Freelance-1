import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap

# Tema verileri
theme_data = {
    'Tema': ['Güç', 'Kötülük', 'Dostluk', 'Yolculuk', 'Umut'],
    'Dakika': [81, 43, 24, 15, 15],
    'Yüzde': [45.6, 24.1, 13.7, 8.2, 8.4]
}

theme_df = pd.DataFrame(theme_data)

# Ana karakterlerin tema dağılımları
character_themes = [
    ["GANDALF", 56, 8, 22, 30, 5],
    ["FRODO", 20, 4, 1, 6, 2],
    ["ARAGORN", 8, 4, 13, 5, 12],
    ["SAM", 13, 4, 1, 2, 3],
    ["BOROMIR", 17, 2, 5, 2, 2],
    ["GALADRIEL", 38, 2, 3, 23, 7],
    ["ELROND", 17, 3, 3, 15, 3],
    ["BILBO", 10, 3, 3, 4, 2],
    ["LEGOLAS", 11, 1, 1, 5, 1],
    ["GIMLI", 6, 3, 5, 7, 0],
    ["SARUMAN", 4, 2, 3, 7, 0]
]

char_df = pd.DataFrame(character_themes, 
                       columns=['Karakter', 'Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut'])

# Her karakterin toplam tema referanslarını hesapla
char_df['Toplam'] = char_df.iloc[:, 1:].sum(axis=1)

# Her tema için yüzde hesapla
for theme in ['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']:
    char_df[f'{theme}_Yüzde'] = (char_df[theme] / char_df['Toplam'] * 100).round(1)

# Alt tema dağılımı
subtheme_data = [
    # Güç alt temaları
    {'Ana Tema': 'Güç', 'Alt Tema': "Yüzük'ün Gücü ve Etkisi", 'Dakika': 36},
    {'Ana Tema': 'Güç', 'Alt Tema': "İktidar ve Otorite Mücadelesi", 'Dakika': 24},
    {'Ana Tema': 'Güç', 'Alt Tema': "Güç Karşısında Direnç", 'Dakika': 21},
    
    # Kötülük alt temaları
    {'Ana Tema': 'Kötülük', 'Alt Tema': "Sauron'un Tehdidi", 'Dakika': 17},
    {'Ana Tema': 'Kötülük', 'Alt Tema': "İhanet (Saruman)", 'Dakika': 13},
    {'Ana Tema': 'Kötülük', 'Alt Tema': "İçsel Karanlık ve Ayartılma", 'Dakika': 13},
    
    # Dostluk alt temaları
    {'Ana Tema': 'Dostluk', 'Alt Tema': "Kardeşlik Bağları", 'Dakika': 10},
    {'Ana Tema': 'Dostluk', 'Alt Tema': "Fedakarlık", 'Dakika': 8},
    {'Ana Tema': 'Dostluk', 'Alt Tema': "Irk/Tür Ayrımlarını Aşma", 'Dakika': 6},
    
    # Yolculuk alt temaları
    {'Ana Tema': 'Yolculuk', 'Alt Tema': "Bilinmeyene Yolculuk", 'Dakika': 6},
    {'Ana Tema': 'Yolculuk', 'Alt Tema': "Değişim ve Büyüme", 'Dakika': 5},
    {'Ana Tema': 'Yolculuk', 'Alt Tema': "Evi Terk Etme", 'Dakika': 4},
    
    # Umut alt temaları
    {'Ana Tema': 'Umut', 'Alt Tema': "Karanlığa Karşı Dayanma", 'Dakika': 8},
    {'Ana Tema': 'Umut', 'Alt Tema': "Küçük Şeylerin Gücü", 'Dakika': 4},
    {'Ana Tema': 'Umut', 'Alt Tema': "İyiliğin Nihai Zaferi", 'Dakika': 3}
]

subtheme_df = pd.DataFrame(subtheme_data)

# Film bölümlerine göre temalar
section_data = [
    {'Bölüm': 'Giriş ve Shire', 'Dakika': 35, 'Baskın Temalar': 'Güç (Yüzük tarihi), Yolculuk (başlangıcı)'},
    {'Bölüm': "Yüzük'ün Keşfi", 'Dakika': 15, 'Baskın Temalar': "Güç (Yüzük'ün etkileri), Kötülük (Sauron'un tehdidi)"},
    {'Bölüm': "Shire'dan Kaçış", 'Dakika': 20, 'Baskın Temalar': "Yolculuk, Dostluk, Kötülük (Nazgûl takibi)"},
    {'Bölüm': 'Bree ve Aragorn', 'Dakika': 15, 'Baskın Temalar': "Dostluk, Güç (Yüzük'ün etkisi)"},
    {'Bölüm': 'Weathertop', 'Dakika': 12, 'Baskın Temalar': "Kötülük, Yolculuk"},
    {'Bölüm': 'Rivendell', 'Dakika': 25, 'Baskın Temalar': "Dostluk (Kardeşliğin kurulması), Güç (Yüzük hakkında kararlar)"},
    {'Bölüm': 'Moria', 'Dakika': 30, 'Baskın Temalar': "Dostluk, Kötülük, Yolculuk"},
    {'Bölüm': 'Lothlórien', 'Dakika': 16, 'Baskın Temalar': "Güç, Umut"},
    {'Bölüm': 'Nehir Yolculuğu', 'Dakika': 10, 'Baskın Temalar': "Yolculuk, Dostluk"}
]

section_df = pd.DataFrame(section_data)

# Temaların film boyunca gelişimi
progression_data = [
    {'Filmin Bölümü': 'Filmin Başı (0-35 dk)', 'Güç': 'Yüksek', 'Yolculuk': 'Orta', 'Dostluk': 'Düşük', 'Kötülük': 'Orta', 'Umut': 'Düşük'},
    {'Filmin Bölümü': 'Filmin İlk Çeyreği (35-70 dk)', 'Güç': 'Orta', 'Yolculuk': 'Yüksek', 'Dostluk': 'Orta', 'Kötülük': 'Orta', 'Umut': 'Düşük'},
    {'Filmin Bölümü': 'Filmin Ortası (70-105 dk)', 'Güç': 'Orta', 'Yolculuk': 'Orta', 'Dostluk': 'Yüksek', 'Kötülük': 'Orta', 'Umut': 'Orta'},
    {'Filmin Bölümü': 'Filmin Üçüncü Çeyreği (105-140 dk)', 'Güç': 'Orta', 'Yolculuk': 'Düşük', 'Dostluk': 'Orta', 'Kötülük': 'Yüksek', 'Umut': 'Düşük'},
    {'Filmin Bölümü': 'Filmin Sonu (140-178 dk)', 'Güç': 'Yüksek', 'Yolculuk': 'Orta', 'Dostluk': 'Yüksek', 'Kötülük': 'Düşük', 'Umut': 'Yüksek'}
]

progression_df = pd.DataFrame(progression_data)

# Tema yoğunluğunu sayısal değerlere dönüştür
intensity_map = {'Yüksek': 3, 'Orta': 2, 'Düşük': 1}

for col in ['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']:
    progression_df[f'{col}_Değer'] = progression_df[col].map(intensity_map)

# Veri analizini bir Excel dosyasına kaydedelim
with pd.ExcelWriter('LOTR_Tema_Analizi.xlsx') as writer:
    theme_df.to_excel(writer, sheet_name='Ana Temalar', index=False)
    char_df.to_excel(writer, sheet_name='Karakter Temaları', index=False)
    subtheme_df.to_excel(writer, sheet_name='Alt Temalar', index=False)
    section_df.to_excel(writer, sheet_name='Film Bölümleri', index=False)
    progression_df.to_excel(writer, sheet_name='Temaların Gelişimi', index=False)

print("Tema analizi Excel dosyası oluşturuldu: LOTR_Tema_Analizi.xlsx")

# Görselleştirmeler için bazı örnekler

# 1. Ana Temaların Pasta Grafiği
plt.figure(figsize=(10, 7))
plt.pie(theme_df['Dakika'], labels=theme_df['Tema'], autopct='%1.1f%%', startangle=140, 
        colors=['#4C72B0', '#DD8452', '#55A868', '#C44E52', '#8172B3'])
plt.title('Yüzüklerin Efendisi: Yüzük Kardeşliği\nTema Dağılımı (Yüzdelik)')
plt.savefig('tema_dagilimi_pasta.png', dpi=300, bbox_inches='tight')

# 2. Karakterlerin Tema Odakları (İlk 7 karakter)
plt.figure(figsize=(14, 8))
top_chars = char_df.iloc[:7].copy()
themes = ['Güç_Yüzde', 'Kötülük_Yüzde', 'Dostluk_Yüzde', 'Yolculuk_Yüzde', 'Umut_Yüzde']
theme_labels = ['Güç', 'Kötülük', 'Dostluk', 'Yolculuk', 'Umut']

# Verileri hazırla
data = []
for theme, label in zip(themes, theme_labels):
    data.append(top_chars[theme].values)

# Çubuk grafiği oluştur
x = np.arange(len(top_chars))
width = 0.15
multiplier = 0

fig, ax = plt.subplots(figsize=(14, 8))

for attribute, measurement in zip(theme_labels, data):
    offset = width * multiplier
    rects = ax.bar(x + offset, measurement, width, label=attribute)
    multiplier += 1

ax.set_xticks(x + width * 2)
ax.set_xticklabels(top_chars['Karakter'])
ax.set_ylabel('Tema Yüzdesi (%)')
ax.set_title('Ana Karakterlerin Tema Odakları')
ax.legend(loc='upper right')
plt.savefig('karakter_tema_odaklari.png', dpi=300, bbox_inches='tight')

# 3. Alt Temaların Ağaç Haritası
plt.figure(figsize=(16, 18))  # Even more increased height for better readability
colors = {'Güç': '#4C72B0', 'Kötülük': '#DD8452', 'Dostluk': '#55A868', 
          'Yolculuk': '#C44E52', 'Umut': '#8172B3'}

# Assign color for each Ana Tema
cmap = {}
for theme in subtheme_df['Ana Tema'].unique():
    cmap[theme] = colors.get(theme, '#333333')

# Assign Ana Tema's color to each subtheme
subtheme_df['Color'] = subtheme_df['Ana Tema'].map(cmap)

# Create tree map
from matplotlib.patches import Rectangle
import matplotlib.patheffects as path_effects

# Calculate y positions
y_pos = 0
rects = []
y_positions = []
heights = []
labels = []
colors = []

# Add larger spacing between themes and taller rectangles for text
theme_spacing = 15  # Increased spacing between themes
min_rect_height = 15  # Minimum rectangle height for better text display

for theme in subtheme_df['Ana Tema'].unique():
    subset = subtheme_df[subtheme_df['Ana Tema'] == theme]
    
    theme_start_y = y_pos  # Remember where this theme starts
    
    # Rectangles for subthemes with minimum height
    for _, row in subset.iterrows():
        # Ensure minimum rectangle height
        rect_height = max(row['Dakika'] * 1.5, min_rect_height)  # Scale Dakika by 1.5 and ensure minimum height
        rect = Rectangle((0, y_pos), 1, rect_height, facecolor=cmap[theme], alpha=0.8, edgecolor='white')
        rects.append((rect, row['Alt Tema'], row['Dakika']))
        y_pos += rect_height + 5  # Add 5 units spacing between subthemes
    
    # Calculate theme height and y position for label
    theme_height = y_pos - theme_start_y
    theme_y_pos = theme_start_y + theme_height / 2
    
    y_positions.append(theme_y_pos)
    heights.append(theme_height)
    labels.append(theme)
    colors.append(cmap[theme])
    
    # Add spacing between themes
    y_pos += theme_spacing

# Create figure with increased width-to-height ratio
fig, ax = plt.subplots(figsize=(16, 20))

# Add rectangles
for rect, label, value in rects:
    ax.add_patch(rect)
    rx, ry = rect.get_xy()
    cx = rx + rect.get_width() / 2.0
    cy = ry + rect.get_height() / 2.0
    
    # Break long labels into multiple lines if needed
    if len(label) > 20:
        words = label.split()
        half = len(words) // 2
        label = ' '.join(words[:half]) + '\n' + ' '.join(words[half:])
    
    # Add black outline around white text for better visibility
    ax.annotate(f"{label}\n({value} min)", (cx, cy), color='white', fontweight='bold', 
                ha='center', va='center', fontsize=14, path_effects=[
                    path_effects.withStroke(linewidth=3, foreground='black')])

# Add Ana Temas with increased font size and better positioning
for y, h, label, color in zip(y_positions, heights, labels, colors):
    ax.annotate(f"{label}", (1.2, y), color=color, fontweight='bold', fontsize=18,
                ha='left', va='center')

# Axis settings with more space for labels
ax.set_xlim(0, 2.2)  # More horizontal space
ax.set_ylim(0, y_pos)
ax.set_title('Subtheme Distribution (Minutes)', fontsize=20, pad=20)
ax.set_xticks([])
ax.set_yticks([])
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_visible(False)
ax.spines['left'].set_visible(False)

plt.tight_layout()
plt.savefig('subtheme_distribution.png', dpi=300, bbox_inches='tight')

# 4. Temaların Film Boyunca Gelişimi (Isı Haritası)
plt.figure(figsize=(14, 8))

# Isı haritası için veri hazırlığı
heatmap_data = progression_df[['Güç_Değer', 'Yolculuk_Değer', 'Dostluk_Değer', 'Kötülük_Değer', 'Umut_Değer']]
heatmap_data.columns = ['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']

# Özel renk haritası
colors = ["#f7fbff", "#6baed6", "#2171b5"]
cmap = LinearSegmentedColormap.from_list("custom_cmap", colors)

# Isı haritası
ax = sns.heatmap(heatmap_data, annot=False, cmap=cmap, linewidths=0.5, 
                fmt=".1f", cbar_kws={'label': 'Tema Yoğunluğu'})

# Görsel ayarları
ax.set_title('Temaların Film Boyunca Gelişimi', fontsize=16)
ax.set_yticklabels(progression_df['Filmin Bölümü'])
plt.tight_layout()
plt.savefig('tema_gelisimi_heatmap.png', dpi=300, bbox_inches='tight')

# 5. Film Bölümlerine Göre Süre Dağılımı
plt.figure(figsize=(12, 8))
bars = plt.barh(section_df['Bölüm'], section_df['Dakika'], color='#3182bd')

# Metin etiketleri ekle
for i, bar in enumerate(bars):
    plt.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2, 
             f"{section_df['Dakika'].iloc[i]} dk", 
             va='center', fontsize=10)

plt.xlabel('Süre (Dakika)')
plt.title('Film Bölümlerine Göre Süre Dağılımı')
plt.tight_layout()
plt.savefig('bolum_sure_dagilimi.png', dpi=300, bbox_inches='tight')

print("Görselleştirmeler tamamlandı ve kaydedildi.")