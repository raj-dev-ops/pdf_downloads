import os
import re
import pandas as pd
from pathlib import Path

def extract_volume_issue_from_path(file_path):
    """Extract volume and issue numbers from file path"""
    match_vol = re.search(r'vol(\d+)', file_path)
    match_iss = re.search(r'iss(\d+)', file_path)
    
    volume = int(match_vol.group(1)) if match_vol else None
    issue = int(match_iss.group(1)) if match_iss else None
    
    return volume, issue

def clean_title(filename):
    """Clean the title by removing .pdf extension"""
    return filename.replace('.pdf', '')

def collect_all_titles():
    """Collect all PDF titles organized by volume and issue"""
    base_paths = [
        Path('/Users/raj/Desktop/pdf_downlaods/downloads'),
        Path('/Users/raj/Desktop/pdf_downlaods/src/downloads')
    ]
    
    all_data = []
    
    for base_path in base_paths:
        if not base_path.exists():
            continue
            
        for pdf_file in base_path.rglob('*.pdf'):
            volume, issue = extract_volume_issue_from_path(str(pdf_file))
            
            if volume and issue:
                title = clean_title(pdf_file.name)
                all_data.append({
                    'Volume': f'Volume {volume}',
                    'Issue': f'Issue {issue}',
                    'Title': title,
                    'Volume_Num': volume,
                    'Issue_Num': issue
                })
    
    return all_data

def create_excel_report(data):
    """Create Excel report with volumes ordered from 18 to 1"""
    if not data:
        print("No data found to export")
        return
    
    df = pd.DataFrame(data)
    
    df = df.sort_values(
        by=['Volume_Num', 'Issue_Num', 'Title'],
        ascending=[False, True, True]
    )
    
    df_display = df[['Volume', 'Issue', 'Title']].copy()
    
    output_file = '/Users/raj/Desktop/pdf_downlaods/journal_titles_collection.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_display.to_excel(writer, sheet_name='All Titles', index=False)
        
        worksheet = writer.sheets['All Titles']
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 15
        worksheet.column_dimensions['C'].width = 100
        
        for volume in sorted(df['Volume_Num'].unique(), reverse=True):
            volume_df = df[df['Volume_Num'] == volume][['Volume', 'Issue', 'Title']]
            volume_df = volume_df.sort_values('Title')
            sheet_name = f'Volume {volume}'
            volume_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            worksheet.column_dimensions['A'].width = 15
            worksheet.column_dimensions['B'].width = 15
            worksheet.column_dimensions['C'].width = 100
    
    print(f"Excel file created: {output_file}")
    
    print(f"\nSummary:")
    print(f"Total titles: {len(df)}")
    print(f"Volumes found: {sorted(df['Volume_Num'].unique(), reverse=True)}")
    
    for volume in sorted(df['Volume_Num'].unique(), reverse=True):
        volume_data = df[df['Volume_Num'] == volume]
        issues = sorted(volume_data['Issue_Num'].unique())
        print(f"  Volume {volume}: {len(volume_data)} titles across issues {issues}")

def main():
    print("Collecting all journal titles...")
    data = collect_all_titles()
    
    if data:
        print(f"Found {len(data)} titles")
        create_excel_report(data)
    else:
        print("No PDF files found in the expected directories")

if __name__ == "__main__":
    main()