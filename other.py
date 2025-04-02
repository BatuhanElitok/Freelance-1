import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
from collections import Counter
import os
import sys
import traceback

print("Script started")
print(f"Python version: {sys.version}")
print(f"Working directory: {os.getcwd()}")

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
    print("Creating Excel writer...")
    return pd.ExcelWriter('LOTR_Fellowship_Analysis.xlsx', engine='xlsxwriter')

def read_data_files():
    """Read all data files and return as pandas DataFrames"""
    print("Reading data files...")
    
    # Read character data
    try:
        print("Loading character data...")
        char_df = pd.read_csv('lotr_analysis_characters.csv')
        print(f"Successfully loaded character data with {len(char_df)} rows")
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
        print("Loading interaction data...")
        interactions_df = pd.read_csv('lotr_analysis_interactions.csv')
        print(f"Successfully loaded interaction data with {len(interactions_df)} rows")
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
        print("Attempting to load script text...")
        script_path = 'data/LordoftheRings1-FOTR.txt'
        with open(script_path, 'r', encoding='utf-8') as f:
            script_text = f.read()
        print(f"Successfully loaded script with {len(script_text)} characters")
    except Exception as e:
        print(f"Script file not available: {e}")
        print("Some analyses will be limited")
    
    return char_df, interactions_df, script_text

def analyze_characters(char_df):
    """Perform detailed character analysis"""
    print("Analyzing character data...")
    
    # Create a basic summary
    char_summary = char_df.copy()
    
    # Add calculated columns
    char_summary['words_per_line'] = char_summary['total_words'] / char_summary['line_count']
    
    # Calculate theme density (theme references per line)
    themes = ['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']
    
    # Handle missing theme columns
    for theme in themes:
        if theme not in char_summary.columns:
            print(f"Warning: Theme column '{theme}' not found, adding with zeros")
            char_summary[theme] = 0
    
    char_summary['total_theme_refs'] = char_summary[themes].sum(axis=1)
    char_summary['theme_density'] = char_summary['total_theme_refs'] / char_summary['line_count']
    
    # Calculate ring mention frequency
    if 'ring_mentions' not in char_summary.columns:
        print("Warning: 'ring_mentions' column not found, adding with zeros")
        char_summary['ring_mentions'] = 0
        
    char_summary['ring_mention_ratio'] = char_summary['ring_mentions'] / char_summary['line_count']
    
    print(f"Character analysis complete with {len(char_summary)} characters")
    return char_summary

def fellowship_analysis(char_df):
    """Analyze statistics for fellowship members"""
    print("Analyzing fellowship members...")
    
    # Define fellowship members (normalized names)
    fellowship_names = ['FRODO', 'SAM', 'MERRY', 'PIPPIN', 'ARAGORN', 'GANDALF', 'LEGOLAS', 'GIMLI', 'BOROMIR']
    
    # Filter for fellowship members (partial match)
    fellowship_df = char_df[char_df['character'].str.contains('|'.join(fellowship_names), case=False, regex=True)].copy()
    
    # Add calculated columns for fellowship
    themes = ['Güç', 'Yolculuk', 'Dostluk', 'Kötülük', 'Umut']
    
    # Handle missing theme columns
    for theme in themes:
        if theme not in fellowship_df.columns:
            print(f"Warning: Theme column '{theme}' not found in fellowship data, adding with zeros")
            fellowship_df[theme] = 0
    
    fellowship_df['total_theme_refs'] = fellowship_df[themes].sum(axis=1)
    fellowship_df['words_per_line'] = fellowship_df['total_words'] / fellowship_df['line_count']
    
    # Handle potential division by zero
    total_dialog = fellowship_df['dialog_count'].sum()
    if total_dialog > 0:
        fellowship_df['pct_of_total_dialog'] = fellowship_df['dialog_count'] / total_dialog * 100
    else:
        fellowship_df['pct_of_total_dialog'] = 0
    
    print(f"Fellowship analysis complete with {len(fellowship_df)} members")
    return fellowship_df

def analyze_interactions(interactions_df):
    """Analyze character interactions"""
    print("Analyzing character interactions...")
    
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
    
    print(f"Interaction analysis complete with {len(interaction_profile)} characters")
    return interaction_profile

def analyze_pairs(interactions_df):
    """Analyze character pair interactions"""
    print("Analyzing character pairs...")
    
    # Create character pairs
    if 'count' in interactions_df.columns:
        pair_counts = interactions_df.groupby(['character1', 'character2'])['count'].sum().reset_index(name='interaction_count')
    else:
        # If no count column, assume each row is one interaction
        pair_counts = interactions_df.groupby(['character1', 'character2']).size().reset_index(name='interaction_count')
    
    # Sort by count
    pair_counts = pair_counts.sort_values('interaction_count', ascending=False)
    
    print(f"Pair analysis complete with {len(pair_counts)} character pairs")
    return pair_counts

def ring_reference_analysis(char_df):
    """Analyze ring references in dialog"""
    print("Analyzing ring references...")
    
    # Check if ring_mentions column exists
    if 'ring_mentions' not in char_df.columns:
        print("Warning: 'ring_mentions' column not found, adding with zeros")
        char_df['ring_mentions'] = 0
    
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
    
    if ring_df.empty:
        print("No characters with ring mentions found")
    else:
        print(f"Ring reference analysis complete with {len(ring_df)} characters")
    
    return ring_df

def create_excel_report():
    """Create a comprehensive Excel report with all analyses"""
    print("Creating Excel report...")
    global char_summary, fellowship_df, interaction_profile, pair_counts, network_df, theme_df, ring_df, location_df, scene_type_df
    
    try:
        # Create the Excel writer
        writer = create_excel_writer()
        
        # Write each dataframe to a different sheet
        # 1. Overall Character Statistics
        print("Writing character stats to Excel...")
        char_summary.to_excel(writer, sheet_name='Character Stats', index=False)
        
        # 2. Fellowship Analysis
        print("Writing fellowship analysis to Excel...")
        fellowship_df.to_excel(writer, sheet_name='Fellowship Analysis', index=False)
        
        # 3. Interaction Analysis
        print("Writing interaction analysis to Excel...")
        interaction_profile.to_excel(writer, sheet_name='Interaction Analysis', index=False)
        
        # 4. Character Pairs
        print("Writing character pairs to Excel...")
        pair_counts.to_excel(writer, sheet_name='Character Pairs', index=False)
        
        # 5. Ring References
        print("Writing ring references to Excel...")
        ring_df.to_excel(writer, sheet_name='Ring References', index=False)
        
        # 6. Summary
        print("Creating summary sheet...")
        summary_data = {
            'Metric': [
                'Total Characters', 
                'Total Dialog Count',
                'Total Words',
                'Most Talkative Character',
                'Most Words Spoken By',
                'Most Ring Mentions By',
                'Character with Most Interactions'
            ],
            'Value': [
                len(char_summary),
                int(char_summary['dialog_count'].sum()),
                int(char_summary['total_words'].sum()),
                char_summary.loc[char_summary['dialog_count'].idxmax(), 'character'],
                char_summary.loc[char_summary['total_words'].idxmax(), 'character'],
                ring_df.iloc[0]['character'] if not ring_df.empty else 'N/A',
                interaction_profile.iloc[0]['character'] if not interaction_profile.empty else 'N/A'
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Save the Excel file
        print("Saving Excel file...")
        writer.save()
        
        # Verify the file was created
        if os.path.exists('LOTR_Fellowship_Analysis.xlsx'):
            file_size = os.path.getsize('LOTR_Fellowship_Analysis.xlsx')
            print(f"Excel report created: LOTR_Fellowship_Analysis.xlsx ({file_size} bytes)")
        else:
            print("WARNING: Excel file was not found after creation attempt")
    
    except Exception as e:
        print(f"Error creating Excel report: {e}")
        traceback.print_exc()

def main():
    """Main function to run the analysis"""
    print("Starting main analysis...")
    
    # Make variables global
    global char_summary, fellowship_df, interaction_profile, pair_counts, network_df, theme_df, ring_df, location_df, scene_type_df
    
    try:
        # Read data
        print("Reading data files...")
        char_df, interactions_df, script_text = read_data_files()
        
        # Perform analyses
        print("Performing analyses...")
        char_summary = analyze_characters(char_df)
        fellowship_df = fellowship_analysis(char_df)
        interaction_profile = analyze_interactions(interactions_df)
        pair_counts = analyze_pairs(interactions_df)
        ring_df = ring_reference_analysis(char_df)
        
        # Create Excel report
        print("Creating Excel report...")
        create_excel_report()
        
        print("Analysis complete! LOTR_Fellowship_Analysis.xlsx has been created.")
    
    except Exception as e:
        print(f"Error in main function: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    try:
        main()
        print("Script completed successfully")
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        traceback.print_exc()
    
    # Keep console open
    input("Press Enter to exit...")