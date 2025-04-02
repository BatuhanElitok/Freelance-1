import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
from collections import Counter
import os

# Global variables for dataframes
char_summary = None
fellowship_df = None
interaction_profile = None
pair_counts = None
network_df = None
theme_df = None
ring_df = None
location_df = None
scene_type_df = None

def create_excel_writer():
    """Create an Excel writer"""
    return pd.ExcelWriter('LOTR_Fellowship_Analysis.xlsx', engine='xlsxwriter')

def read_data_files():
    """Read all data files and return as pandas DataFrames"""
    # Read character data
    try:
        char_df = pd.read_csv('lotr_analysis_characters.csv')
    except Exception as e:
        print(f"Error reading character data: {e}")
        print("Using sample data instead")
        # Create sample data from the first few rows
        char_df = pd.DataFrame({
            'character': ['GANDALF', 'FRODO', 'SAM', 'ARAGORN', 'BILBO'],
            'line_count': [23, 13, 8, 9, 12],
            'dialog_count': [23, 13, 8, 9, 12],
            'voiceover_count': [0, 0, 0, 0, 0],
            'total_words': [3705, 2260, 1393, 1870, 2017],
            'Güç': [56, 20, 13, 8, 10],
            'Yolculuk': [8, 4, 4, 4, 3],
            'Dostluk': [22, 1, 1, 13, 3],
            'Kötülük': [30, 6, 2, 5, 4],
            'Umut': [5, 2, 3, 12, 2],
            'ring_mentions': [30, 8, 3, 1, 8]
        })
    
    # Read interaction data
    try:
        interactions_df = pd.read_csv('lotr_analysis_interactions.csv')
    except Exception as e:
        print(f"Error reading interaction data: {e}")
        print("Using sample data instead")
        # Create sample interaction data
        interactions_df = pd.DataFrame({
            'character1': ['GANDALF', 'FRODO', 'SAM', 'ARAGORN', 'BILBO'],
            'character2': ['FRODO', 'GANDALF', 'FRODO', 'FRODO', 'GANDALF'],
            'count': [2, 3, 1, 2, 2]
        })
    
    # Read script (optional if file is available)
    script_text = ""
    try:
        script_path = 'data/LordoftheRings1-FOTR.txt'
        with open(script_path, 'r', encoding='utf-8') as f:
            script_text = f.read()
    except Exception as e:
        print(f"Script file not available: {e}")
        print("Some analyses will be limited")
    
    return char_df, interactions_df, script_text

def analyze_characters(char_df):
    """Perform detailed character analysis"""
    # Create a basic summary
    char_summary = char_df.copy()
    
    # Add calculated columns
    char_summary['words_per_line'] = char_summary['total_words'] / char_summary['line_count']
    
    # Calculate theme density (theme references per line)
    themes = ['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']
    char_summary['total_theme_refs'] = char_summary[themes].sum(axis=1)
    char_summary['theme_density'] = char_summary['total_theme_refs'] / char_summary['line_count']
    
    # Calculate ring mention frequency
    char_summary['ring_mention_ratio'] = char_summary['ring_mentions'] / char_summary['line_count']
    
    return char_summary

def fellowship_analysis(char_df):
    """Analyze statistics for fellowship members"""
    # Define fellowship members (normalized names)
    fellowship_names = ['FRODO', 'SAM', 'MERRY', 'PIPPIN', 'ARAGORN', 'GANDALF', 'LEGOLAS', 'GIMLI', 'BOROMIR']
    
    # Filter for fellowship members (partial match)
    fellowship_df = char_df[char_df['character'].str.contains('|'.join(fellowship_names), case=False, regex=True)].copy()
    
    # Add calculated columns for fellowship
    themes = ['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']
    fellowship_df['total_theme_refs'] = fellowship_df[themes].sum(axis=1)
    fellowship_df['words_per_line'] = fellowship_df['total_words'] / fellowship_df['line_count']
    
    # Handle potential division by zero
    total_dialog = fellowship_df['dialog_count'].sum()
    if total_dialog > 0:
        fellowship_df['pct_of_total_dialog'] = fellowship_df['dialog_count'] / total_dialog * 100
    else:
        fellowship_df['pct_of_total_dialog'] = 0
    
    return fellowship_df

def analyze_interactions(interactions_df):
    """Analyze character interactions"""
    # Count interactions per character
    char_interactions = interactions_df.groupby('character1').size().reset_index(name='initiated_interactions')
    
    # Count interactions received per character
    received_interactions = interactions_df.groupby('character2').size().reset_index(name='received_interactions')
    
    # Merge to get a complete interaction profile
    interaction_profile = pd.merge(char_interactions, received_interactions, 
                                  left_on='character1', right_on='character2', how='outer')
    
    # Fill NaN with 0 and clean up
    interaction_profile = interaction_profile.fillna(0)
    interaction_profile.rename(columns={'character1': 'character'}, inplace=True)
    
    # Drop character2 column if it exists
    if 'character2' in interaction_profile.columns:
        interaction_profile.drop('character2', axis=1, inplace=True)
    
    # Calculate total interactions and interaction ratio
    interaction_profile['total_interactions'] = interaction_profile['initiated_interactions'] + interaction_profile['received_interactions']
    
    # Handle potential division by zero
    interaction_profile['interaction_ratio'] = 0
    mask = interaction_profile['total_interactions'] > 0
    interaction_profile.loc[mask, 'interaction_ratio'] = interaction_profile.loc[mask, 'initiated_interactions'] / interaction_profile.loc[mask, 'total_interactions']
    
    # Sort by total interactions
    interaction_profile = interaction_profile.sort_values('total_interactions', ascending=False)
    
    return interaction_profile

def analyze_pairs(interactions_df):
    """Analyze character pair interactions"""
    # Create character pairs
    pair_counts = interactions_df.groupby(['character1', 'character2']).size().reset_index(name='interaction_count')
    
    # Sort by count
    pair_counts = pair_counts.sort_values('interaction_count', ascending=False)
    
    return pair_counts

def character_network_analysis(interactions_df):
    """Analyze character connections and network properties"""
    # Number of unique connections per character
    char1_connections = interactions_df.groupby('character1')['character2'].nunique().reset_index(name='outgoing_connections')
    char2_connections = interactions_df.groupby('character2')['character1'].nunique().reset_index(name='incoming_connections')
    
    # Merge
    network_df = pd.merge(char1_connections, char2_connections, 
                         left_on='character1', right_on='character2', how='outer')
    
    # Clean up
    network_df = network_df.fillna(0)
    network_df.rename(columns={'character1': 'character'}, inplace=True)
    
    # Drop character2 column if it exists
    if 'character2' in network_df.columns:
        network_df.drop('character2', axis=1, inplace=True)
    
    # Calculate total unique connections
    network_df['total_connections'] = network_df['outgoing_connections'] + network_df['incoming_connections']
    
    # Sort
    network_df = network_df.sort_values('total_connections', ascending=False)
    
    return network_df

def theme_analysis(char_df):
    """Analyze theme distribution across characters"""
    # Define themes
    themes = ['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']
    
    # Calculate total theme references
    theme_df = char_df.copy()
    theme_df['total_theme_refs'] = theme_df[themes].sum(axis=1)
    
    # Filter to characters with at least one theme reference
    theme_df = theme_df[theme_df['total_theme_refs'] > 0].copy()
    
    # Calculate percentage of each theme for the character
    for theme in themes:
        theme_df[f'{theme}_pct'] = 0  # Initialize with zero
        mask = theme_df['total_theme_refs'] > 0
        theme_df.loc[mask, f'{theme}_pct'] = theme_df.loc[mask, theme] / theme_df.loc[mask, 'total_theme_refs'] * 100
    
    # Sort by total theme references
    theme_df = theme_df.sort_values('total_theme_refs', ascending=False)
    
    return theme_df

def ring_reference_analysis(char_df):
    """Analyze ring references in dialog"""
    # Filter to characters who mention the ring
    ring_df = char_df[pd.notnull(char_df['ring_mentions']) & (char_df['ring_mentions'] > 0)].copy()
    
    # Calculate ring mentions relative to dialog
    ring_df['ring_per_dialog'] = 0
    mask = ring_df['dialog_count'] > 0
    ring_df.loc[mask, 'ring_per_dialog'] = ring_df.loc[mask, 'ring_mentions'] / ring_df.loc[mask, 'dialog_count']
    
    ring_df['ring_per_word'] = 0
    mask = ring_df['total_words'] > 0
    ring_df.loc[mask, 'ring_per_word'] = ring_df.loc[mask, 'ring_mentions'] / ring_df.loc[mask, 'total_words'] * 1000  # Per 1000 words
    
    # Sort by total ring mentions
    ring_df = ring_df.sort_values('ring_mentions', ascending=False)
    
    return ring_df

def analyze_script(script_text):
    """Analyze script for locations, scene distribution, etc."""
    if not script_text:
        return None, None
    
    # Extract scenes
    scenes = re.findall(r'(INT\.|EXT\.)[^\n]+', script_text)
    
    # Extract locations
    locations = []
    for scene in scenes:
        loc_match = re.search(r'(?:INT\.|EXT\.)\s+([^-\n]+)', scene)
        if loc_match:
            locations.append(loc_match.group(1).strip())
    
    # Count locations
    location_counts = Counter(locations)
    location_df = pd.DataFrame({
        'location': list(location_counts.keys()),
        'scene_count': list(location_counts.values())
    })
    location_df['percentage'] = location_df['scene_count'] / len(scenes) * 100
    location_df = location_df.sort_values('scene_count', ascending=False)
    
    # Scene type analysis
    scene_types = {'INT': 0, 'EXT': 0}
    for scene in scenes:
        if scene.startswith('INT.'):
            scene_types['INT'] += 1
        elif scene.startswith('EXT.'):
            scene_types['EXT'] += 1
    
    scene_type_df = pd.DataFrame({
        'scene_type': list(scene_types.keys()),
        'count': list(scene_types.values())
    })
    scene_type_df['percentage'] = scene_type_df['count'] / len(scenes) * 100
    
    return location_df, scene_type_df

def format_workbook(writer):
    """Apply formatting to the Excel workbook"""
    workbook = writer.book
    
    # Create formats
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'bg_color': '#D9E1F2',
        'border': 1
    })
    
    # Apply the header format to all sheets
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        
        # Format column widths
        worksheet.set_column(0, 0, 20)  # First column (usually character name)
        worksheet.set_column(1, 20, 15)  # Data columns
        
        # The first row already contains headers, so format them
        if sheet_name == 'Character Stats':
            for col_num, value in enumerate(char_summary.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Fellowship Analysis':
            for col_num, value in enumerate(fellowship_df.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Interaction Analysis':
            for col_num, value in enumerate(interaction_profile.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Character Pairs':
            for col_num, value in enumerate(pair_counts.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Character Network':
            for col_num, value in enumerate(network_df.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Theme Analysis':
            for col_num, value in enumerate(theme_df.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Ring References':
            for col_num, value in enumerate(ring_df.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Locations' and location_df is not None:
            for col_num, value in enumerate(location_df.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Scene Types' and scene_type_df is not None:
            for col_num, value in enumerate(scene_type_df.columns):
                worksheet.write(0, col_num, value, header_format)
                
        elif sheet_name == 'Summary':
            for col_num, value in enumerate(['Metric', 'Value']):
                worksheet.write(0, col_num, value, header_format)
    
    # Add conditional formatting for certain sheets
    if 'Ring References' in writer.sheets:
        worksheet = writer.sheets['Ring References']
        # Column indices for ring_per_dialog and ring_per_word may vary
        ring_per_dialog_col = ring_df.columns.get_loc('ring_per_dialog') if 'ring_per_dialog' in ring_df.columns else -1
        if ring_per_dialog_col >= 0:
            col_letter = chr(65 + ring_per_dialog_col)  # Convert to Excel column letter (A, B, C...)
            worksheet.conditional_format(f'{col_letter}2:{col_letter}100', {
                'type': '3_color_scale',
                'min_color': '#FFFFFF',
                'mid_color': '#FFEB9C',
                'max_color': '#FFC7CE'
            })
    
    if 'Theme Analysis' in writer.sheets and any(col.endswith('_pct') for col in theme_df.columns):
        worksheet = writer.sheets['Theme Analysis']
        # Find columns that end with _pct
        pct_cols = [i for i, col in enumerate(theme_df.columns) if col.endswith('_pct')]
        for col_idx in pct_cols:
            col_letter = chr(65 + col_idx)  # Convert to Excel column letter
            worksheet.conditional_format(f'{col_letter}2:{col_letter}100', {
                'type': '3_color_scale',
                'min_color': '#FFFFFF',
                'mid_color': '#9BC2E6',
                'max_color': '#4472C4'
            })

def create_excel_report():
    """Create a comprehensive Excel report with all analyses"""
    global char_summary, fellowship_df, interaction_profile, pair_counts, network_df, theme_df, ring_df, location_df, scene_type_df
    
    # Create the Excel writer
    writer = create_excel_writer()
    
    # Write each dataframe to a different sheet
    # 1. Overall Character Statistics
    char_summary.to_excel(writer, sheet_name='Character Stats', index=False)
    
    # 2. Fellowship Analysis
    fellowship_df.to_excel(writer, sheet_name='Fellowship Analysis', index=False)
    
    # 3. Interaction Analysis
    interaction_profile.to_excel(writer, sheet_name='Interaction Analysis', index=False)
    
    # 4. Character Pairs
    pair_counts.to_excel(writer, sheet_name='Character Pairs', index=False)
    
    # 5. Character Network
    network_df.to_excel(writer, sheet_name='Character Network', index=False)
    
    # 6. Theme Analysis
    theme_df.to_excel(writer, sheet_name='Theme Analysis', index=False)
    
    # 7. Ring References
    ring_df.to_excel(writer, sheet_name='Ring References', index=False)
    
    # 8. Location Analysis (if available)
    if location_df is not None:
        location_df.to_excel(writer, sheet_name='Locations', index=False)
        
    # 9. Scene Type Analysis (if available)
    if scene_type_df is not None:
        scene_type_df.to_excel(writer, sheet_name='Scene Types', index=False)
    
    # 10. Dashboard/Summary
    # Create a summary sheet with key metrics
    summary_data = {
        'Metric': [
            'Total Characters', 
            'Total Dialog Count',
            'Total Words',
            'Most Talkative Character',
            'Most Words Spoken By',
            'Most Ring Mentions By',
            'Character with Most Interactions',
            'Most Common Character Pair',
            'Most Referenced Theme',
            'Most Common Location (if available)'
        ],
        'Value': [
            len(char_summary),
            int(char_summary['dialog_count'].sum()),
            int(char_summary['total_words'].sum()),
            char_summary.loc[char_summary['dialog_count'].idxmax(), 'character'],
            char_summary.loc[char_summary['total_words'].idxmax(), 'character'],
            ring_df.iloc[0]['character'] if not ring_df.empty else 'N/A',
            interaction_profile.iloc[0]['character'] if not interaction_profile.empty else 'N/A',
            f"{pair_counts.iloc[0]['character1']} → {pair_counts.iloc[0]['character2']}" if not pair_counts.empty else 'N/A',
            char_summary[['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']].sum().idxmax(),
            location_df.iloc[0]['location'] if location_df is not None and not location_df.empty else 'N/A'
        ]
    }
    
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    # Format the workbook
    format_workbook(writer)
    
    # Save the Excel file
    writer.save()
    
    print(f"Excel report created: LOTR_Fellowship_Analysis.xlsx")

def main():
    """Main function to run the analysis"""
    # Make variables global
    global char_summary, fellowship_df, interaction_profile, pair_counts, network_df, theme_df, ring_df, location_df, scene_type_df
    
    # Read data
    char_df, interactions_df, script_text = read_data_files()
    
    # Perform analyses
    char_summary = analyze_characters(char_df)
    fellowship_df = fellowship_analysis(char_df)
    interaction_profile = analyze_interactions(interactions_df)
    pair_counts = analyze_pairs(interactions_df)
    network_df = character_network_analysis(interactions_df)
    theme_df = theme_analysis(char_df)
    ring_df = ring_reference_analysis(char_df)
    
    # Script analysis (if available)
    location_df, scene_type_df = analyze_script(script_text)
    
    # Create Excel report
    create_excel_report()
    
    print("Analysis complete! LOTR_Fellowship_Analysis.xlsx has been created.")