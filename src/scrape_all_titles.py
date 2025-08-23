#!/usr/bin/env python3
"""
IJES Title Collector - Scrapes all article titles from volumes 1-18 and exports to Excel
"""

import re
import time
import logging
from typing import List, Dict, Optional
import requests
from bs4 import BeautifulSoup
import pandas as pd
from tqdm import tqdm

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

BASE_URL = "https://intjexersci.com"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
}

class IJESTitleCollector:
    """Collects all article titles from IJES website"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self.all_articles = []
    
    def _make_request(self, url: str, max_retries: int = 3) -> Optional[requests.Response]:
        """Make HTTP request with retry logic"""
        for attempt in range(max_retries):
            try:
                response = self.session.get(url, timeout=30)
                response.raise_for_status()
                return response
            except requests.exceptions.RequestException as e:
                logger.warning(f"Request failed (attempt {attempt + 1}/{max_retries}): {url}")
                if attempt < max_retries - 1:
                    time.sleep(2 * (attempt + 1))
                else:
                    logger.error(f"Failed to fetch {url} after {max_retries} attempts")
                    return None
    
    def get_max_issue_for_volume(self, volume: int) -> int:
        """Find the maximum issue number for a given volume"""
        max_issue = 0
        for issue in range(1, 20):
            url = f"{BASE_URL}/ijes/vol{volume}/iss{issue}/"
            response = self._make_request(url)
            if response and response.status_code == 200:
                max_issue = issue
            else:
                break
        return max_issue
    
    def get_article_titles(self, volume: int, issue: int) -> List[Dict]:
        """Get all article titles from a volume/issue page"""
        url = f"{BASE_URL}/ijes/vol{volume}/iss{issue}/"
        logger.info(f"Fetching titles from Volume {volume}, Issue {issue}")
        
        response = self._make_request(url)
        if not response:
            return []
        
        soup = BeautifulSoup(response.content, 'html.parser')
        articles = []
        
        all_links = soup.find_all('a', href=True)
        
        for link in all_links:
            href = link.get('href', '')
            title = link.get_text(strip=True)
            
            if not title or len(title) < 10:
                continue
            
            if re.search(rf'/ijes/vol{volume}/iss{issue}/\d+/?', href):
                cleaned_title = self.clean_title(title)
                if cleaned_title and len(cleaned_title) > 10:
                    articles.append({
                        'Title': cleaned_title,
                        'Volume': f'Volume {volume}',
                        'Issue': f'Issue {issue}',
                        'Volume_Num': volume,
                        'Issue_Num': issue
                    })
                    logger.debug(f"Found: {cleaned_title}")
        
        if not articles:
            content_divs = soup.find_all(['div', 'p', 'h3', 'h4'])
            for div in content_divs:
                text = div.get_text(strip=True)
                if self.is_likely_article_title(text):
                    cleaned_title = self.clean_title(text)
                    if cleaned_title and len(cleaned_title) > 10:
                        articles.append({
                            'Title': cleaned_title,
                            'Volume': f'Volume {volume}',
                            'Issue': f'Issue {issue}',
                            'Volume_Num': volume,
                            'Issue_Num': issue
                        })
        
        unique_articles = []
        seen_titles = set()
        for article in articles:
            if article['Title'] not in seen_titles:
                seen_titles.add(article['Title'])
                unique_articles.append(article)
        
        logger.info(f"Found {len(unique_articles)} unique articles in Volume {volume}, Issue {issue}")
        return unique_articles
    
    def clean_title(self, title: str) -> str:
        """Clean and normalize article title"""
        title = re.sub('<.*?>', '', title)
        title = re.sub(r'\s+', ' ', title)
        title = title.strip()
        
        if title.endswith('.pdf'):
            title = title[:-4]
        
        prefixes_to_remove = ['Download', 'View', 'PDF:', 'Article:', 'Full Text:']
        for prefix in prefixes_to_remove:
            if title.startswith(prefix):
                title = title[len(prefix):].strip()
        
        return title
    
    def is_likely_article_title(self, text: str) -> bool:
        """Check if text is likely an article title"""
        if len(text) < 20 or len(text) > 300:
            return False
        
        exclude_patterns = [
            'Volume', 'Issue', 'Table of Contents', 'Editorial',
            'Copyright', 'ISSN', 'Published by', 'All rights',
            'Browse', 'Search', 'Login', 'Register', 'Home'
        ]
        
        text_lower = text.lower()
        for pattern in exclude_patterns:
            if pattern.lower() in text_lower:
                return False
        
        word_count = len(text.split())
        if word_count < 3 or word_count > 50:
            return False
        
        return True
    
    def collect_all_volumes(self, start_volume: int = 1, end_volume: int = 18) -> None:
        """Collect titles from all volumes"""
        total_volumes = end_volume - start_volume + 1
        
        with tqdm(total=total_volumes, desc="Collecting volumes") as pbar:
            for volume in range(end_volume, start_volume - 1, -1):
                pbar.set_description(f"Processing Volume {volume}")
                
                max_issue = self.get_max_issue_for_volume(volume)
                if max_issue == 0:
                    logger.warning(f"No issues found for Volume {volume}")
                    pbar.update(1)
                    continue
                
                logger.info(f"Volume {volume} has {max_issue} issues")
                
                for issue in range(1, max_issue + 1):
                    articles = self.get_article_titles(volume, issue)
                    self.all_articles.extend(articles)
                    time.sleep(0.5)
                
                pbar.update(1)
    
    def export_to_excel(self, output_file: str = 'ijes_all_titles.xlsx') -> None:
        """Export collected titles to Excel with proper formatting"""
        if not self.all_articles:
            logger.error("No articles to export")
            return
        
        df = pd.DataFrame(self.all_articles)
        
        df = df.sort_values(
            by=['Volume_Num', 'Issue_Num', 'Title'],
            ascending=[False, True, True]
        )
        
        df_display = df[['Title', 'Volume', 'Issue']].copy()
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_display.to_excel(writer, sheet_name='All Titles', index=False)
            
            worksheet = writer.sheets['All Titles']
            worksheet.column_dimensions['A'].width = 100
            worksheet.column_dimensions['B'].width = 15
            worksheet.column_dimensions['C'].width = 15
            
            for volume in sorted(df['Volume_Num'].unique(), reverse=True):
                volume_df = df[df['Volume_Num'] == volume][['Title', 'Volume', 'Issue']]
                volume_df = volume_df.sort_values('Title')
                sheet_name = f'Volume {volume}'
                
                try:
                    volume_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]
                    worksheet.column_dimensions['A'].width = 100
                    worksheet.column_dimensions['B'].width = 15
                    worksheet.column_dimensions['C'].width = 15
                except Exception as e:
                    logger.warning(f"Could not create sheet for {sheet_name}: {e}")
        
        print(f"\n✅ Excel file created: {output_file}")
        print(f"Total articles collected: {len(df)}")
        
        for volume in sorted(df['Volume_Num'].unique(), reverse=True):
            volume_data = df[df['Volume_Num'] == volume]
            issues = sorted(volume_data['Issue_Num'].unique())
            print(f"  Volume {volume}: {len(volume_data)} articles across issues {list(issues)}")

def main():
    """Main function to collect all IJES titles"""
    print("=" * 60)
    print("IJES Title Collector")
    print("Collecting all article titles from Volumes 1-18")
    print("=" * 60)
    
    collector = IJESTitleCollector()
    
    try:
        print("\nStarting collection process...")
        collector.collect_all_volumes(start_volume=1, end_volume=18)
        
        print("\nExporting to Excel...")
        output_file = '/Users/raj/Desktop/pdf_downlaods/ijes_all_titles_from_web.xlsx'
        collector.export_to_excel(output_file)
        
        print("\n✅ Collection complete!")
        
    except KeyboardInterrupt:
        print("\n⏹️ Collection interrupted by user")
        if collector.all_articles:
            print("Saving partial results...")
            output_file = '/Users/raj/Desktop/pdf_downlaods/ijes_partial_titles.xlsx'
            collector.export_to_excel(output_file)
    except Exception as e:
        logger.exception("Error during collection")
        print(f"\n❌ Error: {e}")

if __name__ == '__main__':
    main()