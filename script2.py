'''
Only sort_code field is failing
'''

import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import time
import random

def scrape_oaknorth_reviews():
    # List to store review data
    reviews_data = []
    
    # Base URL for OakNorth Bank Trustpilot reviews
    base_url = "https://www.trustpilot.com/review/www.oaknorth.co.uk?stars=1,2,3&page={}"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    # Loop through pages
    for page in range(1, 100):
        try:
            # Get the page - Added verify=False to bypass SSL verification
            response = requests.get(base_url.format(page), headers=headers, verify=False)
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Find all review containers
            reviews = soup.find_all('article', class_='review')
            
            if not reviews:
                break
                
            for review in reviews:
                try:
                    # Extract review data
                    username = review.find('span', class_='consumer-info__name').text.strip()
                    
                    # Get review date
                    date_element = review.find('time')
                    review_date = date_element['datetime'] if date_element else 'Not found'
                    
                    # Get review content
                    comment = review.find('p', class_='review-content__text').text.strip()
                    
                    # Get company reply if exists
                    reply_container = review.find('div', class_='brand-company-reply')
                    company_reply = ''
                    reply_date = ''
                    
                    if reply_container:
                        company_reply = reply_container.find('p').text.strip()
                        reply_time = reply_container.find('time')
                        reply_date = reply_time['datetime'] if reply_time else 'Not found'
                    
                    # Add to reviews list
                    reviews_data.append({
                        'Username': username,
                        'Date': review_date,
                        'Comment': comment,
                        'Company Reply': company_reply,
                        'Reply Date': reply_date
                    })
                    
                except Exception as e:
                    print(f"Error processing review: {e}")
                    continue
            
            # Add delay to be respectful to the website
            time.sleep(random.uniform(2, 4))
            
        except Exception as e:
            print(f"Error processing page {page}: {e}")
            continue
            
    # Create DataFrame and save to CSV
    df = pd.DataFrame(reviews_data)
    df.to_csv('oaknorth_reviews.csv', index=False, encoding='utf-8-sig')
    return df

# Run the scraper
if __name__ == "__main__":
    # Suppress SSL verification warnings
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    reviews_df = scrape_oaknorth_reviews()
    print(f"Scraped {len(reviews_df)} reviews successfully!")
