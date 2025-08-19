#!/usr/bin/env python3
"""
IJES (International Journal of Exercise Science) Web Scraper

Scrapes articles and PDFs from the IJES website for specified volumes and issues.
"""

import os
import re
import sys
import time
import logging
from pathlib import Path
from typing import List, Tuple, Optional
from urllib.parse import urljoin, urlparse

import click
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ijes_scraper.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Constants
BASE_URL = "https://intjexersci.com"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}
RETRY_ATTEMPTS = 3
RETRY_DELAY = 2  # seconds
REQUEST_TIMEOUT = 30  # seconds


class IJESScraper:
    """Scraper for the International Journal of Exercise Science website."""
    
    def __init__(self, base_dir: str = "downloads"):
        """
        Initialize the scraper.
        
        Args:
            base_dir: Base directory for downloads
        """
        self.base_dir = Path(base_dir)
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        
    def _make_request(self, url: str, max_retries: int = RETRY_ATTEMPTS) -> Optional[requests.Response]:
        """
        Make an HTTP request with retry logic.
        
        Args:
            url: URL to request
            max_retries: Maximum number of retry attempts
            
        Returns:
            Response object or None if failed
        """
        for attempt in range(max_retries):
            try:
                response = self.session.get(url, timeout=REQUEST_TIMEOUT)
                response.raise_for_status()
                return response
            except requests.exceptions.RequestException as e:
                logger.warning(f"Request failed (attempt {attempt + 1}/{max_retries}): {url} - {e}")
                if attempt < max_retries - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                else:
                    logger.error(f"Failed to fetch {url} after {max_retries} attempts")
                    return None
    
    def _sanitize_filename(self, title: str) -> str:
        """
        Sanitize article title for use as filename.
        
        Args:
            title: Article title
            
        Returns:
            Sanitized filename
        """
        # Remove HTML tags if any
        title = re.sub('<.*?>', '', title)
        # Replace problematic characters
        title = re.sub(r'[<>:"/\\|?*]', '', title)
        # Replace multiple spaces with single space
        title = re.sub(r'\s+', ' ', title)
        # Truncate to reasonable length
        title = title[:200].strip()
        # Remove trailing dots (Windows issue)
        title = title.rstrip('.')
        return title or "untitled"
    
    def get_article_links(self, volume: int, issue: int) -> List[Tuple[str, str]]:
        """
        Get all article links from a volume/issue page.
        
        Args:
            volume: Volume number
            issue: Issue number
            
        Returns:
            List of tuples (article_url, article_title)
        """
        url = f"{BASE_URL}/ijes/vol{volume}/iss{issue}/"
        logger.info(f"Fetching article list from: {url}")
        
        response = self._make_request(url)
        if not response:
            return []
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # DEBUG: Print raw HTML structure to understand page layout
        logger.info("=== RAW HTML STRUCTURE ===")
        logger.info(f"Page title: {soup.title.get_text() if soup.title else 'No title'}")
        logger.info(f"Total links found: {len(soup.find_all('a'))}")
        logger.info("=== END RAW HTML ===")
        
        # DEBUG: Show all links found on the page
        all_links = soup.find_all('a', href=True)
        logger.info(f"=== ALL LINKS FOUND ({len(all_links)}) ===")
        for i, link in enumerate(all_links):
            href = link.get('href', '')
            text = link.get_text(strip=True)[:100]  # Truncate for readability
            logger.info(f"Link {i+1}: {href} | Text: '{text}'")
        logger.info("=== END ALL LINKS ===")
        
        articles = []
        
        # Method 1: Original regex pattern (for comparison)
        original_pattern = rf'/ijes/vol{volume}/iss{issue}/\d+/'
        article_links_original = soup.find_all('a', href=re.compile(original_pattern))
        logger.info(f"Method 1 (Original regex {original_pattern}): Found {len(article_links_original)} matches")
        
        # Method 2: More flexible patterns for article links
        flexible_patterns = [
            rf'/ijes/vol{volume}/iss{issue}/\d+/?',  # With optional trailing slash
            rf'vol{volume}/iss{issue}/\d+/?',       # Without leading /ijes/
            rf'iss{issue}/\d+/?',                   # Just issue and number
            rf'/\d+/?$',                            # Just numbers at the end
        ]
        
        for i, pattern in enumerate(flexible_patterns):
            matches = soup.find_all('a', href=re.compile(pattern))
            logger.info(f"Method {i+2} (Pattern '{pattern}'): Found {len(matches)} matches")
            for link in matches:
                href = link.get('href', '')
                text = link.get_text(strip=True)
                logger.info(f"  - {href} | '{text}'")
        
        # Method 3: Look for links containing volume and issue numbers
        vol_iss_links = []
        for link in all_links:
            href = link.get('href', '')
            text = link.get_text(strip=True)
            # Check if href contains vol and iss patterns
            if (f'vol{volume}' in href and f'iss{issue}' in href) or \
               (f'volume={volume}' in href and f'issue={issue}' in href):
                vol_iss_links.append(link)
        logger.info(f"Method 3 (Vol/Iss in href): Found {len(vol_iss_links)} matches")
        
        # Method 4: Look for PDF links directly
        pdf_links = soup.find_all('a', href=re.compile(r'\.pdf$', re.I))
        logger.info(f"Method 4 (Direct PDF links): Found {len(pdf_links)} matches")
        for link in pdf_links:
            href = link.get('href', '')
            text = link.get_text(strip=True)
            logger.info(f"  PDF: {href} | '{text}'")
        
        # Method 5: Look for common article patterns in text
        potential_article_links = []
        for link in all_links:
            href = link.get('href', '')
            text = link.get_text(strip=True).lower()
            # Look for common article indicators
            article_indicators = ['full text', 'pdf', 'article', 'download', 'view', 'read']
            if any(indicator in text for indicator in article_indicators) and href:
                potential_article_links.append(link)
        logger.info(f"Method 5 (Article indicators): Found {len(potential_article_links)} matches")
        
        # Combine all methods and deduplicate
        all_candidate_links = set()
        
        # Add original pattern matches
        for link in article_links_original:
            href = link.get('href', '')
            if href:
                all_candidate_links.add((urljoin(url, href), link.get_text(strip=True)))
        
        # Add flexible pattern matches
        for pattern in flexible_patterns:
            matches = soup.find_all('a', href=re.compile(pattern))
            for link in matches:
                href = link.get('href', '')
                if href:
                    full_url = urljoin(url, href)
                    # Only add if it looks like an article URL
                    if f'vol{volume}' in full_url and f'iss{issue}' in full_url:
                        all_candidate_links.add((full_url, link.get_text(strip=True)))
        
        # Add vol/iss matches
        for link in vol_iss_links:
            href = link.get('href', '')
            if href:
                all_candidate_links.add((urljoin(url, href), link.get_text(strip=True)))
        
        # Add potential article links if they contain volume/issue info
        for link in potential_article_links:
            href = link.get('href', '')
            full_url = urljoin(url, href)
            if href and (f'vol{volume}' in full_url or f'iss{issue}' in full_url):
                all_candidate_links.add((full_url, link.get_text(strip=True)))
        
        # Convert to list and filter
        articles = []
        for article_url, title in all_candidate_links:
            if title and title.strip():  # Ensure we have a valid title
                articles.append((article_url, title.strip()))
        
        logger.info(f"=== FINAL RESULTS ===")
        logger.info(f"Total unique articles found: {len(articles)}")
        for i, (url, title) in enumerate(articles):
            logger.info(f"Article {i+1}: {url}")
            logger.info(f"  Title: '{title}'")
        logger.info("=== END FINAL RESULTS ===")
        
        return articles
    
    def get_pdf_url(self, article_url: str) -> Optional[str]:
        """
        Extract PDF download URL from an article page.
        
        Args:
            article_url: URL of the article page
            
        Returns:
            PDF URL or None if not found
        """
        logger.info(f"Searching for PDF in: {article_url}")
        response = self._make_request(article_url)
        if not response:
            return None
        
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # DEBUG: Show all links on the article page
        all_links = soup.find_all('a', href=True)
        logger.info(f"=== ARTICLE PAGE LINKS ({len(all_links)}) ===")
        for i, link in enumerate(all_links):
            href = link.get('href', '')
            text = link.get_text(strip=True)[:100]
            logger.info(f"Link {i+1}: {href} | Text: '{text}'")
        logger.info("=== END ARTICLE PAGE LINKS ===")
        
        # Method 1: Original IJES pattern - /files/ijes/vol*/iss*/*.pdf
        original_pdf_pattern = r'/files/ijes/vol\d+/iss\d+/\d+\.pdf'
        pdf_links = soup.find_all('a', href=re.compile(original_pdf_pattern, re.I))
        logger.info(f"Method 1 (Original PDF pattern): Found {len(pdf_links)} matches")
        
        if pdf_links:
            href = pdf_links[0].get('href', '')
            if href:
                pdf_url = urljoin(article_url, href)
                logger.info(f"Found PDF via Method 1: {pdf_url}")
                return pdf_url
        
        # Method 2: More flexible PDF patterns
        pdf_patterns = [
            r'/files/.*\.pdf',           # Any file in /files/ ending with .pdf
            r'vol\d+/iss\d+/.*\.pdf',   # Volume/issue pattern anywhere
            r'/\d+\.pdf',               # Just number.pdf
            r'\.pdf$',                  # Any .pdf at end of URL
        ]
        
        for i, pattern in enumerate(pdf_patterns):
            matches = soup.find_all('a', href=re.compile(pattern, re.I))
            logger.info(f"Method {i+2} (PDF pattern '{pattern}'): Found {len(matches)} matches")
            for link in matches:
                href = link.get('href', '')
                text = link.get_text(strip=True)
                if href and 'Guide-for-Peer-Review' not in href:  # Skip peer review guide
                    pdf_url = urljoin(article_url, href)
                    logger.info(f"Found PDF via Method {i+2}: {pdf_url} | Text: '{text}'")
                    return pdf_url
        
        # Method 3: Look for common PDF link indicators in text
        pdf_text_indicators = ['pdf', 'full text', 'download', 'view pdf', 'article pdf', 'full article']
        for link in all_links:
            href = link.get('href', '')
            text = link.get_text(strip=True).lower()
            if href and any(indicator in text for indicator in pdf_text_indicators):
                # Check if the link might lead to a PDF
                if '.pdf' in href.lower() or any(term in text for term in ['pdf', 'full text']):
                    pdf_url = urljoin(article_url, href)
                    logger.info(f"Found PDF via Method 3 (text indicators): {pdf_url} | Text: '{text}'")
                    return pdf_url
        
        # Method 4: Check meta tags for PDF URLs
        meta_tags = [
            'citation_pdf_url',
            'dc.identifier.uri',
            'dc.relation.uri',
            'citation_fulltext_html_url'
        ]
        
        for meta_name in meta_tags:
            meta_pdf = soup.find('meta', {'name': meta_name})
            if meta_pdf:
                content = meta_pdf.get('content', '')
                if content and '.pdf' in content.lower():
                    logger.info(f"Found PDF via Method 4 (meta tag {meta_name}): {content}")
                    return content
        
        # Method 5: Look for iframe or embed tags that might contain PDFs
        iframes = soup.find_all(['iframe', 'embed', 'object'])
        for iframe in iframes:
            src = iframe.get('src', '') or iframe.get('data', '')
            if src and '.pdf' in src.lower():
                pdf_url = urljoin(article_url, src)
                logger.info(f"Found PDF via Method 5 (iframe/embed): {pdf_url}")
                return pdf_url
        
        # Method 6: Try to construct PDF URL based on article URL pattern
        # Extract article number from URL and try common PDF patterns
        article_match = re.search(r'/(\d+)/?$', article_url)
        if article_match:
            article_num = article_match.group(1)
            vol_match = re.search(r'/vol(\d+)/', article_url)
            iss_match = re.search(r'/iss(\d+)/', article_url)
            
            if vol_match and iss_match:
                volume = vol_match.group(1)
                issue = iss_match.group(1)
                
                # Try common PDF URL patterns
                pdf_patterns_to_try = [
                    f"{BASE_URL}/files/ijes/vol{volume}/iss{issue}/{article_num}.pdf",
                    f"{BASE_URL}/files/vol{volume}/iss{issue}/{article_num}.pdf",
                    f"{BASE_URL}/pdf/vol{volume}/iss{issue}/{article_num}.pdf",
                ]
                
                for pdf_url in pdf_patterns_to_try:
                    logger.info(f"Method 6: Trying constructed URL: {pdf_url}")
                    # Test if the URL exists with a HEAD request
                    try:
                        head_response = self.session.head(pdf_url, timeout=10)
                        if head_response.status_code == 200:
                            logger.info(f"Found PDF via Method 6 (constructed URL): {pdf_url}")
                            return pdf_url
                    except:
                        continue
        
        logger.warning(f"No PDF found for article: {article_url}")
        return None
    
    def download_pdf(self, pdf_url: str, save_path: Path) -> bool:
        """
        Download a PDF file.
        
        Args:
            pdf_url: URL of the PDF
            save_path: Path to save the file
            
        Returns:
            True if successful, False otherwise
        """
        try:
            response = self._make_request(pdf_url)
            if not response:
                return False
            
            # Ensure directory exists
            save_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Save the PDF
            with open(save_path, 'wb') as f:
                f.write(response.content)
            
            logger.info(f"Downloaded: {save_path.name}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to download {pdf_url}: {e}")
            return False
    
    def scrape_issue(self, volume: int, issue: int) -> Tuple[int, int]:
        """
        Scrape all articles from a specific volume/issue.
        
        Args:
            volume: Volume number
            issue: Issue number
            
        Returns:
            Tuple of (successful_downloads, total_articles)
        """
        logger.info(f"Starting scrape for Volume {volume}, Issue {issue}")
        
        # Create directory for this issue with new naming format
        # Volume folder: "IJES Volume 18" instead of "vol18"
        # Issue folder: "IJES 18-1" instead of "iss1"
        volume_dir = self.base_dir / f"IJES Volume {volume}"
        issue_dir = volume_dir / f"IJES {volume}-{issue}"
        issue_dir.mkdir(parents=True, exist_ok=True)
        
        # Get article links
        articles = self.get_article_links(volume, issue)
        if not articles:
            logger.warning(f"No articles found for Volume {volume}, Issue {issue}")
            return 0, 0
        
        successful = 0
        with tqdm(total=len(articles), desc=f"Vol {volume} Issue {issue}") as pbar:
            for article_url, title in articles:
                pbar.set_description(f"Processing: {title[:50]}...")
                
                # Get PDF URL
                pdf_url = self.get_pdf_url(article_url)
                if not pdf_url:
                    logger.warning(f"No PDF found for: {title}")
                    pbar.update(1)
                    continue
                
                # Prepare filename
                filename = f"{self._sanitize_filename(title)}.pdf"
                save_path = issue_dir / filename
                
                # Skip if already downloaded
                if save_path.exists():
                    logger.info(f"Already exists: {filename}")
                    successful += 1
                    pbar.update(1)
                    continue
                
                # Download PDF
                if self.download_pdf(pdf_url, save_path):
                    successful += 1
                    time.sleep(1)  # Be polite to the server
                
                pbar.update(1)
        
        logger.info(f"Downloaded {successful}/{len(articles)} articles")
        return successful, len(articles)


@click.command()
@click.option('--volume', '-v', required=True, type=int, help='Volume number')
@click.option('--issue', '-i', required=True, type=int, help='Issue number')
@click.option('--output-dir', '-o', default='downloads', help='Output directory (default: downloads)')
@click.option('--verbose', is_flag=True, help='Enable verbose logging')
def main(volume: int, issue: int, output_dir: str, verbose: bool):
    """
    IJES Web Scraper - Download articles from the International Journal of Exercise Science.
    
    Example:
        python ijes_scraper.py -v 18 -i 8
    """
    if verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    click.echo(f"🔍 Starting IJES scraper for Volume {volume}, Issue {issue}")
    
    scraper = IJESScraper(base_dir=output_dir)
    
    try:
        successful, total = scraper.scrape_issue(volume, issue)
        
        if total == 0:
            click.echo("❌ No articles found. Please check the volume and issue numbers.")
            sys.exit(1)
        elif successful == total:
            click.echo(f"✅ Successfully downloaded all {total} articles!")
        else:
            click.echo(f"⚠️  Downloaded {successful}/{total} articles. Check logs for details.")
            
    except KeyboardInterrupt:
        click.echo("\n⏹️  Scraping interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.exception("Unexpected error occurred")
        click.echo(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()