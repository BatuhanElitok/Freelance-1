import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap

# Theme data
theme_data = {
    'Theme': ['Power', 'Evil', 'Friendship', 'Journey', 'Hope'],
    'Minutes': [81, 43, 24, 15, 15],
    'Percentage': [45.6, 24.1, 13.7, 8.2, 8.4]
}

theme_df = pd.DataFrame(theme_data)

# Theme distributions of main characters
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
                       columns=['Character', 'Power', 'Journey', 'Friendship', 'Evil', 'Hope'])

# Calculate total theme references for each character
char_df['Total'] = char_df.iloc[:, 1:].sum(axis=1)

# Calculate percentages for each theme
for theme in ['Power', 'Journey', 'Friendship', 'Evil', 'Hope']:
    char_df[f'{theme}_Percent'] = (char_df[theme] / char_df['Total'] * 100).round(1)

# Subtheme distribution
subtheme_data = [
    # Power subthemes
    {'Main Theme': 'Power', 'Subtheme': "The Power and Influence of the Ring", 'Minutes': 36},
    {'Main Theme': 'Power', 'Subtheme': "Struggle for Power and Authority", 'Minutes': 24},
    {'Main Theme': 'Power', 'Subtheme': "Resistance Against Power", 'Minutes': 21},
    
    # Evil subthemes
    {'Main Theme': 'Evil', 'Subtheme': "Threat of Sauron", 'Minutes': 17},
    {'Main Theme': 'Evil', 'Subtheme': "Betrayal (Saruman)", 'Minutes': 13},
    {'Main Theme': 'Evil', 'Subtheme': "Inner Darkness and Temptation", 'Minutes': 13},
    
    # Friendship subthemes
    {'Main Theme': 'Friendship', 'Subtheme': "Bonds of Fellowship", 'Minutes': 10},
    {'Main Theme': 'Friendship', 'Subtheme': "Sacrifice", 'Minutes': 8},
    {'Main Theme': 'Friendship', 'Subtheme': "Overcoming Racial/Species Differences", 'Minutes': 6},
    
    # Journey subthemes
    {'Main Theme': 'Journey', 'Subtheme': "Journey into the Unknown", 'Minutes': 6},
    {'Main Theme': 'Journey', 'Subtheme': "Change and Growth", 'Minutes': 5},
    {'Main Theme': 'Journey', 'Subtheme': "Leaving Home", 'Minutes': 4},
    
    # Hope subthemes
    {'Main Theme': 'Hope', 'Subtheme': "Standing Against Darkness", 'Minutes': 8},
    {'Main Theme': 'Hope', 'Subtheme': "Power of Small Things", 'Minutes': 4},
    {'Main Theme': 'Hope', 'Subtheme': "Ultimate Victory of Good", 'Minutes': 3}
]

subtheme_df = pd.DataFrame(subtheme_data)

# Film sections by theme
section_data = [
    {'Section': 'Introduction and Shire', 'Minutes': 35, 'Dominant Themes': 'Power (Ring history), Journey (beginning)'},
    {'Section': "Discovery of the Ring", 'Minutes': 15, 'Dominant Themes': "Power (Ring's effects), Evil (Sauron's threat)"},
    {'Section': "Escape from the Shire", 'Minutes': 20, 'Dominant Themes': "Journey, Friendship, Evil (Nazgûl pursuit)"},
    {'Section': 'Bree and Aragorn', 'Minutes': 15, 'Dominant Themes': "Friendship, Power (Ring's influence)"},
    {'Section': 'Weathertop', 'Minutes': 12, 'Dominant Themes': "Evil, Journey"},
    {'Section': 'Rivendell', 'Minutes': 25, 'Dominant Themes': "Friendship (Formation of Fellowship), Power (Decisions about the Ring)"},
    {'Section': 'Moria', 'Minutes': 30, 'Dominant Themes': "Friendship, Evil, Journey"},
    {'Section': 'Lothlórien', 'Minutes': 16, 'Dominant Themes': "Power, Hope"},
    {'Section': 'River Journey', 'Minutes': 10, 'Dominant Themes': "Journey, Friendship"}
]

section_df = pd.DataFrame(section_data)

# Theme development throughout the film
progression_data = [
    {'Film Section': 'Beginning of Film (0-35 min)', 'Power': 'High', 'Journey': 'Medium', 'Friendship': 'Low', 'Evil': 'Medium', 'Hope': 'Low'},
    {'Film Section': 'First Quarter (35-70 min)', 'Power': 'Medium', 'Journey': 'High', 'Friendship': 'Medium', 'Evil': 'Medium', 'Hope': 'Low'},
    {'Film Section': 'Middle of Film (70-105 min)', 'Power': 'Medium', 'Journey': 'Medium', 'Friendship': 'High', 'Evil': 'Medium', 'Hope': 'Medium'},
    {'Film Section': 'Third Quarter (105-140 min)', 'Power': 'Medium', 'Journey': 'Low', 'Friendship': 'Medium', 'Evil': 'High', 'Hope': 'Low'},
    {'Film Section': 'End of Film (140-178 min)', 'Power': 'High', 'Journey': 'Medium', 'Friendship': 'High', 'Evil': 'Low', 'Hope': 'High'}
]

progression_df = pd.DataFrame(progression_data)

# Convert theme intensity to numerical values
intensity_map = {'High': 3, 'Medium': 2, 'Low': 1}

for col in ['Power', 'Journey', 'Friendship', 'Evil', 'Hope']:
    progression_df[f'{col}_Value'] = progression_df[col].map(intensity_map)

# Save the data analysis to an Excel file
with pd.ExcelWriter('LOTR_Theme_Analysis.xlsx') as writer:
    theme_df.to_excel(writer, sheet_name='Main Themes', index=False)
    char_df.to_excel(writer, sheet_name='Character Themes', index=False)
    subtheme_df.to_excel(writer, sheet_name='Subthemes', index=False)
    section_df.to_excel(writer, sheet_name='Film Sections', index=False)
    progression_df.to_excel(writer, sheet_name='Theme Progression', index=False)

print("Theme analysis Excel file created: LOTR_Theme_Analysis.xlsx")

# Examples for visualizations

# 1. Pie Chart of Main Themes
plt.figure(figsize=(10, 7))
plt.pie(theme_df['Minutes'], labels=theme_df['Theme'], autopct='%1.1f%%', startangle=140, 
        colors=['#4C72B0', '#DD8452', '#55A868', '#C44E52', '#8172B3'])
plt.title('The Lord of the Rings: The Fellowship of the Ring\nTheme Distribution (Minutes)')
plt.savefig('theme_distribution_pie.png', dpi=300, bbox_inches='tight')

# 2. Character Theme Focus (First 7 characters)
plt.figure(figsize=(14, 8))
top_chars = char_df.iloc[:7].copy()
themes = ['Power_Percent', 'Evil_Percent', 'Friendship_Percent', 'Journey_Percent', 'Hope_Percent']
theme_labels = ['Power', 'Evil', 'Friendship', 'Journey', 'Hope']

# Prepare data
data = []
for theme, label in zip(themes, theme_labels):
    data.append(top_chars[theme].values)

# Create bar chart
x = np.arange(len(top_chars))
width = 0.15
multiplier = 0

fig, ax = plt.subplots(figsize=(14, 8))

for attribute, measurement in zip(theme_labels, data):
    offset = width * multiplier
    rects = ax.bar(x + offset, measurement, width, label=attribute)
    multiplier += 1

ax.set_xticks(x + width * 2)
ax.set_xticklabels(top_chars['Character'])
ax.set_ylabel('Theme Percentage (%)')
ax.set_title('Theme Focus of Main Characters')
ax.legend(loc='upper right')
plt.savefig('character_theme_focus.png', dpi=300, bbox_inches='tight')

# 3. Further Improved Subtheme Tree Map with Better Spacing
plt.figure(figsize=(16, 18))  # Even more increased height for better readability
colors = {'Power': '#4C72B0', 'Evil': '#DD8452', 'Friendship': '#55A868', 
          'Journey': '#C44E52', 'Hope': '#8172B3'}

# Assign color for each main theme
cmap = {}
for theme in subtheme_df['Main Theme'].unique():
    cmap[theme] = colors.get(theme, '#333333')

# Assign main theme's color to each subtheme
subtheme_df['Color'] = subtheme_df['Main Theme'].map(cmap)

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

for theme in subtheme_df['Main Theme'].unique():
    subset = subtheme_df[subtheme_df['Main Theme'] == theme]
    
    theme_start_y = y_pos  # Remember where this theme starts
    
    # Rectangles for subthemes with minimum height
    for _, row in subset.iterrows():
        # Ensure minimum rectangle height
        rect_height = max(row['Minutes'] * 1.5, min_rect_height)  # Scale minutes by 1.5 and ensure minimum height
        rect = Rectangle((0, y_pos), 1, rect_height, facecolor=cmap[theme], alpha=0.8, edgecolor='white')
        rects.append((rect, row['Subtheme'], row['Minutes']))
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

# Add main themes with increased font size and better positioning
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

# 4. Theme Development Throughout the Film (Heat Map)
plt.figure(figsize=(14, 8))

# Prepare data for heat map
heatmap_data = progression_df[['Power_Value', 'Journey_Value', 'Friendship_Value', 'Evil_Value', 'Hope_Value']]
heatmap_data.columns = ['Power', 'Journey', 'Friendship', 'Evil', 'Hope']

# Custom color map
colors = ["#f7fbff", "#6baed6", "#2171b5"]
cmap = LinearSegmentedColormap.from_list("custom_cmap", colors)

# Heat map
ax = sns.heatmap(heatmap_data, annot=False, cmap=cmap, linewidths=0.5, 
                fmt=".1f", cbar_kws={'label': 'Theme Intensity'})

# Visual settings
ax.set_title('Theme Development Throughout the Film', fontsize=16)
ax.set_yticklabels(progression_df['Film Section'])
plt.tight_layout()
plt.savefig('theme_development_heatmap.png', dpi=300, bbox_inches='tight')

# 5. Duration Distribution by Film Section
plt.figure(figsize=(12, 8))
bars = plt.barh(section_df['Section'], section_df['Minutes'], color='#3182bd')

# Add text labels
for i, bar in enumerate(bars):
    plt.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2, 
             f"{section_df['Minutes'].iloc[i]} min", 
             va='center', fontsize=10)

plt.xlabel('Duration (Minutes)')
plt.title('Duration Distribution by Film Section')
plt.tight_layout()
plt.savefig('section_duration_distribution.png', dpi=300, bbox_inches='tight')

print("Visualizations completed and saved.")