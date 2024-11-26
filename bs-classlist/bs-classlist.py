#!/usr/bin/env python3
"""
D2L Brightspace Classlist Scraper
Extracts class list information from D2L Brightspace using Selenium.
Handles Microsoft 365 authentication including MFA.
"""

import time
from pathlib import Path
from dataclasses import dataclass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, 
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
    WebDriverException
)
import pandas as pd
from typing import Optional, Dict, List, Union
import logging
import sys
import os
import argparse
from bs4 import BeautifulSoup
import getpass
import re
from urllib.parse import urlparse

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create a file handler
file_handler = logging.FileHandler('d2l_scraper.log')
file_handler.setLevel(logging.DEBUG)
file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

@dataclass
class StudentRecord:
    """Data class for storing student information"""
    last_name: str
    first_name: str
    username: str
    org_id: str
    email: str
    role: str
    last_accessed: str

class D2LScraperException(Exception):
    """Base exception class for D2L scraper errors"""
    pass

class AuthenticationError(D2LScraperException):
    """Exception raised for authentication failures"""
    pass

class NavigationError(D2LScraperException):
    """Exception raised for navigation failures"""
    pass

class D2LClasslistScraper:
    """Scrapes class list information from D2L Brightspace"""
    
    def __init__(self, course_url: str, show_browser: bool = False):
        """Initialize the D2L scraper with the course URL"""
        self._validate_url(course_url)
        self.course_url = course_url
        self.classlist_url = self._infer_classlist_url(course_url)
        self.driver = None
        self.wait = None
        self.show_browser = show_browser

    def _infer_classlist_url(self, course_url: str) -> str:
        """Infer the classlist URL from the course URL"""
        parsed_url = urlparse(course_url)
        path_parts = parsed_url.path.rstrip('/').split('/')
        if 'ou' in path_parts:
            ou_index = path_parts.index('ou')
            if ou_index + 1 < len(path_parts):
                ou_value = path_parts[ou_index + 1]
                return f"{parsed_url.scheme}://{parsed_url.netloc}/d2l/lms/classlist/classlist.d2l?ou={ou_value}"
        
        # If we can't infer the ou value, construct the URL without it
        return f"{parsed_url.scheme}://{parsed_url.netloc}/d2l/lms/classlist/classlist.d2l"

    @staticmethod
    def _validate_url(url: str) -> None:
        """Basic URL validation"""
        parsed = urlparse(url)
        if not all([parsed.scheme, parsed.netloc]):
            raise ValueError("Invalid URL. Must include scheme (http/https) and domain")
        if parsed.scheme not in ('http', 'https'):
            raise ValueError("URL must use HTTP or HTTPS protocol")

    def setup_driver(self) -> None:
        """Set up Chrome WebDriver with appropriate options"""
        options = webdriver.ChromeOptions()
        if not self.show_browser:
            options.add_argument('--headless=new')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-notifications')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')
        options.add_experimental_option('prefs', {
            'credentials_enable_service': False,
            'profile.password_manager_enabled': False
        })
        
        try:
            self.driver = webdriver.Chrome(options=options)
            self.wait = WebDriverWait(self.driver, 20)
            self.driver.implicitly_wait(10)
        except WebDriverException as e:
            raise WebDriverException(f"Failed to initialize Chrome driver: {str(e)}")

    def wait_and_find_element(self, by: By, value: str, timeout: int = 20) -> webdriver.remote.webelement.WebElement:
        """Wait for element and return it when found"""
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
        except TimeoutException:
            logger.error(f"Timeout waiting for element: {value}")
            raise

    def wait_and_click(self, by: By, value: str, timeout: int = 20, retries: int = 3) -> None:
        """Wait for element to be clickable and click it with retry logic"""
        for attempt in range(retries):
            try:
                element = WebDriverWait(self.driver, timeout).until(
                    EC.element_to_be_clickable((by, value))
                )
                self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(0.5)  # Allow time for scroll
                element.click()
                return
            except (ElementClickInterceptedException, StaleElementReferenceException) as e:
                if attempt == retries - 1:
                    logger.error(f"Failed to click element after {retries} attempts")
                    raise
                time.sleep(1)

    def login(self, username: str, password: str) -> bool:
        """Handle Microsoft 365 authentication for D2L"""
        try:
            logger.info("Starting login process...")
            self.driver.get(self.course_url)
            
            # Handle email input
            try:
                email_input = WebDriverWait(self.driver, 20).until(
                    EC.element_to_be_clickable((By.NAME, "loginfmt"))
                )
                logger.info("Email input field found")
                self.driver.execute_script("arguments[0].value = '';", email_input)
                logger.info("Email input field cleared")
                email_input.send_keys(username)
                logger.info("Username entered")
                self.wait_and_click(By.ID, "idSIButton9")
                logger.info("Next button clicked")
            except Exception as e:
                logger.error(f"Error during email input: {str(e)}")
                raise
            
            # Handle password input
            try:
                password_input = self.wait_and_find_element(By.NAME, "passwd")
                logger.info("Password input field found")
                password_input.clear()
                logger.info("Password input field cleared")
                password_input.send_keys(password)
                logger.info("Password entered")
                self.wait_and_click(By.ID, "idSIButton9")
                logger.info("Sign in button clicked")
            except Exception as e:
                logger.error(f"Error during password input: {str(e)}")
                raise
            
            # Handle MFA code input
            try:
                # Wait for MFA options to be visible
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "idDiv_SAOTCS_Proofs"))
                )
                
                # Check for different MFA options
                mfa_options = self.driver.find_elements(By.CLASS_NAME, "table")
                
                if not mfa_options:
                    raise AuthenticationError("No MFA options found")
                
                # Display available MFA options to the user
                print("Available MFA options:")
                for i, option in enumerate(mfa_options):
                    print(f"{i + 1}. {option.text}")
                
                # Let the user choose the MFA option
                choice = int(input("Enter the number of your preferred MFA option: ")) - 1
                
                if choice < 0 or choice >= len(mfa_options):
                    raise ValueError("Invalid MFA option selected")
                
                # Click the selected MFA option
                mfa_options[choice].click()
                logger.info(f"Selected MFA option: {mfa_options[choice].text}")
                
                # Wait for confirmation
                confirm = input("Press Enter after you've received the code: ")
                
                # Wait for code input field
                self.wait_and_find_element(By.NAME, "otc", timeout=30)
                mfa_code = input("Enter the verification code: ").strip()
                
                # Enter the code
                code_input = self.driver.find_element(By.NAME, "otc")
                code_input.clear()
                code_input.send_keys(mfa_code)
                
                # Click verify button
                verify_button = self.wait_and_find_element(
                    By.XPATH, 
                    "//input[@type='submit']",
                    timeout=10
                )
                verify_button.click()
                logger.info("Submitted verification code")
                
                # Wait for the code to be processed
                time.sleep(5)
                
            except TimeoutException:
                logger.error("MFA process failed or timed out")
                raise AuthenticationError("MFA process failed")
            except ValueError as e:
                logger.error(f"MFA selection error: {str(e)}")
                raise AuthenticationError(f"MFA selection failed: {str(e)}")
            
            # Handle "Stay signed in?" prompt
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "idSIButton9"))
                )
                self.wait_and_click(By.ID, "idSIButton9")
                logger.info("Handled 'Stay signed in' prompt")
            except TimeoutException:
                logger.info("No 'Stay signed in' prompt found")
                
            # Wait for D2L page to load
            success_elements = [
                'div.d2l-navigation-s-main-wrapper',
                'div.d2l-branding-navigation-dark-foreground-color',
                'nav.d2l-navigation-s',
                'd2l-navigation-main-header'
            ]
            
            for selector in success_elements:
                try:
                    WebDriverWait(self.driver, 30).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    logger.info("Successfully detected D2L page load")
                    return True
                except TimeoutException:
                    continue
            
            raise AuthenticationError("Failed to verify successful login to D2L")
            
        except Exception as e:
            logger.error(f"Login failed: {str(e)}")
            raise AuthenticationError(f"Login failed: {str(e)}")

    def navigate_to_classlist(self) -> None:
        """Navigate to the classlist page"""
        try:
            logger.info(f"Attempting direct navigation to classlist: {self.classlist_url}")
            self.driver.get(self.classlist_url)
            
            try:
                self.wait_and_find_element(By.CLASS_NAME, "d2l-table", timeout=10)
                logger.info("Direct navigation successful")
                return
            except TimeoutException:
                logger.warning("Direct navigation failed, trying menu navigation")
            
            # If direct navigation fails, try through menu
            self.driver.get(self.course_url)
            self.wait_and_click(
                By.XPATH,
                "//button[contains(@class, 'd2l-navigation-s-group')]"
            )
            
            classlist_selectors = [
                "//a[contains(text(), 'Classlist')]",
                "//d2l-menu-item-link[contains(@text, 'Classlist')]",
                "//a[contains(@href, 'classlist.d2l')]"
            ]
            
            for selector in classlist_selectors:
                try:
                    self.wait_and_click(By.XPATH, selector, timeout=5)
                    if self.wait_and_find_element(By.CLASS_NAME, "d2l-table", timeout=10):
                        logger.info("Menu navigation successful")
                        return
                except (TimeoutException, NoSuchElementException):
                    continue
                    
            raise NavigationError("Could not access classlist page")
            
        except Exception as e:
            raise NavigationError(f"Failed to navigate to classlist: {str(e)}")

    def parse_classlist(self) -> List[StudentRecord]:
        """Parse the classlist page and extract student information"""
        students = []
        retries = 3
        
        for attempt in range(retries):
            try:
                time.sleep(2)  # Wait for dynamic content
                
                logger.debug("Parsing page source")
                soup = BeautifulSoup(self.driver.page_source, 'html.parser')
                
                logger.debug("Looking for classlist table")
                table = soup.find('table', {'class': 'd2l-table d2l-grid d_gl'})
                
                if not table:
                    logger.error("Classlist table not found. Page source:")
                    logger.error(soup.prettify())
                    raise ValueError("Classlist table not found")
                
                rows = table.find_all('tr')[1:]  # Skip header
                logger.debug(f"Found {len(rows)} rows in the table")
                
                if not rows:
                    logger.error("No student rows found. Table content:")
                    logger.error(table.prettify())
                    raise ValueError("No student rows found")
                    
                for row in rows:
                    cells = row.find_all(['td', 'th'])
                    logger.debug(f"Processing row with {len(cells)} cells")
                    
                    if len(cells) >= 8:
                        name_cell = cells[2].get_text(strip=True)
                        if not name_cell or ',' not in name_cell:
                            logger.warning(f"Skipping row due to invalid name cell: {name_cell}")
                            continue
                            
                        last_name, first_name = map(str.strip, name_cell.split(',', 1))
                        
                        student = StudentRecord(
                            last_name=last_name,
                            first_name=first_name,
                            username=cells[3].get_text(strip=True),
                            org_id=cells[4].get_text(strip=True),
                            email=cells[5].get_text(strip=True),
                            role=cells[6].get_text(strip=True),
                            last_accessed=cells[7].get_text(strip=True)
                        )
                        
                        if all(vars(student).values()):
                            students.append(student)
                            logger.debug(f"Added student: {student}")
                        else:
                            logger.warning(f"Skipping student due to missing data: {vars(student)}")
                
                if students:
                    logger.info(f"Successfully parsed {len(students)} students")
                    break
                else:
                    logger.warning(f"No students parsed in attempt {attempt + 1}")
                    
            except Exception as e:
                logger.error(f"Error in parse_classlist attempt {attempt + 1}: {str(e)}")
                if attempt == retries - 1:
                    logger.error("Failed to parse classlist after all attempts")
                    raise ValueError(f"Failed to parse classlist: {str(e)}")
                time.sleep(2)
                    
        return students

    def get_classlist(self, username: Optional[str] = None, 
                     password: Optional[str] = None) -> Optional[pd.DataFrame]:
        """Main method to get the class list"""
        try:
            logger.info("Setting up WebDriver")
            self.setup_driver()
            
            if not username:
                username = input("Enter your Microsoft 365 email: ")
            if not password:
                password = getpass.getpass("Enter your password: ")
            
            logger.info("Attempting login")
            self.login(username, password)
            
            logger.info("Navigating to classlist page")
            self.navigate_to_classlist()
            
            logger.info("Parsing classlist")
            students = self.parse_classlist()
            
            if not students:
                logger.error("No student data extracted")
                raise D2LScraperException("No student data extracted")
                
            logger.info(f"Successfully extracted data for {len(students)} students")
            return pd.DataFrame([vars(s) for s in students])
            
        except Exception as e:
            logger.error(f"Error getting classlist: {str(e)}", exc_info=True)
            raise
            
        finally:
            if self.driver:
                try:
                    logger.info("Closing WebDriver")
                    self.driver.quit()
                except Exception as e:
                    logger.error(f"Error closing browser: {str(e)}", exc_info=True)


def format_table(df: pd.DataFrame, format_type: str = 'plain') -> str:
    """Format DataFrame for display"""
    if format_type == 'fancy':
        return df.to_string(index=False)
    
    headers = df.columns
    rows = [headers] + df.values.tolist()
    
    col_widths = [
        max(len(str(row[i])) for row in rows)
        for i in range(len(headers))
    ]
    
    formatted_rows = [
        " | ".join(str(cell).ljust(width) 
        for cell, width in zip(row, col_widths))
        for row in rows
    ]
    
    separator = "-" * len(formatted_rows[0])
    formatted_rows.insert(1, separator)
    
    return "\n".join(formatted_rows)


def save_file(df: pd.DataFrame, filename: str, file_format: str) -> None:
    """Save DataFrame to file"""
    path = Path(filename)
    if path.exists():
        raise IOError(f"File {filename} already exists")
        
    try:
        if file_format == 'csv':
            df.to_csv(filename, index=False)
        else:  # excel
            df.to_excel(filename, index=False)
    except Exception as e:
        raise IOError(f"Failed to save file: {str(e)}")


def main() -> None:
    """Main entry point for the scraper"""
    parser = argparse.ArgumentParser(
        description='D2L Brightspace Classlist Scraper',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('url', help='D2L course URL')
    parser.add_argument('-u', '--username', help='Microsoft 365 email')
    parser.add_argument('-p', '--password', 
                       help='Microsoft 365 password (not recommended)')
    parser.add_argument('-o', '--output', 
                       choices=['screen', 'csv', 'excel'],
                       default='screen',
                       help='Output format (default: screen)')
    parser.add_argument('-f', '--filename',
                       help='Output filename (for csv/excel)')
    parser.add_argument('--show-browser',
                       action='store_true',
                       help='Show browser window during operation')
    parser.add_argument('--fancy-output',
                       action='store_true',
                       help='Use fancy table formatting for screen output')
    parser.add_argument('--debug',
                       action='store_true',
                       help='Enable debug logging')

    args = parser.parse_args()

    # Set debug logging if requested
    if args.debug:
        logger.setLevel(logging.DEBUG)
        fh = logging.FileHandler('d2l_scraper_debug.log')
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        fh.setFormatter(formatter)
        logger.addHandler(fh)

    try:
        # Validate URL
        if not args.url.startswith(('http://', 'https://')):
            args.url = f'https://{args.url}'

        # Initialize scraper
        scraper = D2LClasslistScraper(args.url, show_browser=args.show_browser)
        
        # Get the classlist
        try:
            classlist_df = scraper.get_classlist(args.username, args.password)
        except AuthenticationError:
            logger.error("Authentication failed. Please check your credentials.")
            sys.exit(1)
        except NavigationError:
            logger.error("Failed to navigate to classlist. Please check the URL and your permissions.")
            sys.exit(1)
        except D2LScraperException as e:
            logger.error(f"Scraping failed: {str(e)}")
            sys.exit(1)

        if classlist_df is not None and not classlist_df.empty:
            if args.output == 'screen':
                print("\nClass List:")
                print(format_table(classlist_df, 
                                 'fancy' if args.fancy_output else 'plain'))
            else:
                try:
                    filename = args.filename or f"classlist.{args.output}"
                    save_file(classlist_df, filename, args.output)
                    print(f"Saved class list to {filename}")
                except IOError as e:
                    logger.error(f"Failed to save file: {str(e)}")
                    # Offer to display on screen instead
                    if input("Would you like to display the results instead? (y/n): ").lower() == 'y':
                        print("\nClass List:")
                        print(format_table(classlist_df, 
                                         'fancy' if args.fancy_output else 'plain'))
                    sys.exit(1)
        else:
            logger.error("No data was retrieved from the classlist")
            sys.exit(1)

    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
        sys.exit(130)  # Standard exit code for SIGINT
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        if args.debug:
            logger.exception("Detailed error information:")
        sys.exit(1)


if __name__ == "__main__":
    main()
    
