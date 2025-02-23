from scholarly import scholarly
import pandas as pd
import time
from datetime import datetime
import random
import re
import os

# Default URL
#DEFAULT_URL = "a google scholar profile page url" #FIXME
DEFAULT_URL = "https://scholar.google.com/citations?user=NvBZp6MAAAAJ&hl=en"

def calculate_metrics(citations_list):
    """
    Calculate h-index and i10-index from citation counts
    """
    if not citations_list:
        return 0, 0
    
    # Sort citations in descending order
    sorted_citations = sorted(citations_list, reverse=True)
    
    # Calculate h-index
    h_index = 0
    for i, citation_count in enumerate(sorted_citations, 1):
        if citation_count >= i:
            h_index = i
        else:
            break
    
    # Calculate i10-index
    i10_index = sum(1 for citation_count in sorted_citations if citation_count >= 10)
    
    return h_index, i10_index

def load_existing_data(filename):
    """
    Load existing data from file if it exists
    Returns tuple of (papers_df, metrics_df)
    """
    try:
        if os.path.exists(filename):
            # Read main papers data
            papers_df = pd.read_excel(filename, sheet_name=0)
            if papers_df.index.name != 'Title' and 'Title' in papers_df.columns:
                papers_df.set_index('Title', inplace=True)
            
            # Try to read metrics history
            try:
                metrics_df = pd.read_excel(filename, sheet_name='Metrics History')
            except:
                metrics_df = None
                
            return papers_df, metrics_df
    except Exception as e:
        print(f"Error loading existing data: {str(e)}")
    return None, None

def scrape_scholar_profile(url):
    """
    Scrape publication data from a Google Scholar profile URL
    """
    try:
        user_match = re.search(r'user=([^&]+)', url)
        if not user_match:
            raise ValueError("Could not extract user ID from URL")
        
        user_id = user_match.group(1)
        print(f"Extracted user ID: {user_id}")
        
        print("Searching for author profile...")
        search_query = scholarly.search_author_id(user_id)
        author = scholarly.fill(search_query)
        
        if not author:
            raise ValueError("Could not find author profile")
            
        author_name = author.get('name', 'unknown_author')
        print(f"Found author: {author_name}")
        
        # Dictionary to store paper data
        paper_dict = {}
        
        print("\nRetrieving publications...")
        for idx, pub in enumerate(author['publications'], 1):
            try:
                pub_filled = scholarly.fill(pub)
                citations = pub_filled.get('num_citations', 0)
                paper_name = pub_filled['bib'].get('title', '').strip()
                
                if paper_name:
                    paper_dict[paper_name] = citations
                    print(f"Processed paper {idx}: {paper_name[:50]}...")
                
                time.sleep(random.uniform(2, 5))
                
            except Exception as e:
                print(f"Error processing publication: {str(e)}")
                continue
        
        return paper_dict, author_name
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return None, None

def normalize_title(title):
    """
    Normalize paper title by removing extra whitespace and converting to lowercase
    """
    if pd.isna(title):
        return title
    return ' '.join(str(title).lower().split())

def merge_citation_data(existing_df, new_citations, current_date):
    """
    Merge existing citation data with new citations using case-insensitive title matching
    """
    citation_col = f'citations_{current_date}'
    
    if existing_df is None:
        # First run - create new DataFrame
        df = pd.DataFrame.from_dict(new_citations, orient='index', columns=[citation_col])
        df.index.name = 'Title'
        # Store original titles as a column
        df['Original_Title'] = df.index
        # Create normalized index
        df.index = df.index.map(normalize_title)
    else:
        # Add new column to existing data
        df = existing_df.copy()
        
        # If Original_Title column doesn't exist in existing data, create it
        if 'Original_Title' not in df.columns:
            df['Original_Title'] = df.index
            df.index = df.index.map(normalize_title)
        
        # Create new DataFrame with normalized titles as index
        new_df = pd.DataFrame.from_dict(new_citations, orient='index', columns=[citation_col])
        new_df.index.name = 'Title'
        # Store original titles temporarily with a different name
        new_df['New_Original_Title'] = new_df.index
        new_df.index = new_df.index.map(normalize_title)
        
        # Merge existing and new data
        df = df.join(new_df[[citation_col, 'New_Original_Title']], how='outer')
        
        # Update Original_Title for any new entries
        mask = df['Original_Title'].isna()
        df.loc[mask, 'Original_Title'] = df.loc[mask, 'New_Original_Title']
        
        # Drop the temporary column
        df = df.drop('New_Original_Title', axis=1)
    
    return df

def send_metrics_email(recipient_email, total_papers, total_citations, h_index, i10_index, author_name):
    """
    Send an email with the citation metrics summary
    
    Parameters:
    total_papers (int): Total number of papers
    total_citations (int): Total number of citations
    h_index (int): Current h-index
    i10_index (int): Current i10-index
    author_name (str): Name of the author
    """
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from datetime import datetime
    
    # Email settings - replace with your details
    sender_email = "xxx@gmail.com"  # Replace with your Gmail FIXME
    sender_password = "abcd efgh ijkl mnop"   # Replace with your app password FIXME
    
    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f'Google Scholar Metrics Update - {author_name} - {datetime.now().strftime("%Y-%m-%d")}'
    
    # Create email body
    body = f"""
    Google Scholar Metrics Update for {author_name}
    Date: {datetime.now().strftime('%Y-%m-%d')}
    
    Summary:
    - Total Papers: {total_papers}
    - Total Citations: {total_citations}
    - h-index: {h_index}
    - i10-index: {i10_index}
    
    This is an automated message from your Google Scholar Citation Scraper.
    """
    
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        # Create SMTP session
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        
        # Login
        server.login(sender_email, sender_password)
        
        # Send email
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        print(f"\nMetrics summary email sent successfully to {recipient_email}")
        
    except Exception as e:
        print(f"\nError sending email: {str(e)}")
        
    finally:
        server.quit()

def main():
    print("Google Scholar Citation Scraper")
    print("------------------------------")
    
    url = DEFAULT_URL
    
    print("\nStarting search...")
    print(f"Using URL: {url}")
    
    # Get new citation data
    paper_dict, author_name = scrape_scholar_profile(url)
    
    if paper_dict is not None and author_name is not None:
        current_date = datetime.now().strftime('%Y-%m-%d')
        citation_col = f'citations_{current_date}'
        
        # Prepare filename
        safe_author_name = re.sub(r'[<>:"/\\|?* ]', '_', author_name)
        output_filename = f'scholar_citations_{safe_author_name}.xlsx'
        
        # Load existing data
        existing_papers_df, existing_metrics_df = load_existing_data(output_filename)
        
        # Merge existing and new citation data
        papers_df = merge_citation_data(existing_papers_df, paper_dict, current_date)
        
        # Sort by latest citation count
        papers_df = papers_df.sort_values(citation_col, ascending=False)
        
        # Calculate metrics for current citations
        current_citations = papers_df[citation_col].fillna(0).astype(int)
        h_index, i10_index = calculate_metrics(current_citations.tolist())
        total_citations = current_citations.sum()
        
        # Create or update metrics history
        current_metrics = pd.DataFrame({
            'Date': [current_date],
            'Total Citations': [total_citations],
            'h-index': [h_index],
            'i10-index': [i10_index],
            'Total Papers': [len(papers_df)]
        })
        
        if existing_metrics_df is not None:
            metrics_df = pd.concat([existing_metrics_df, current_metrics], ignore_index=True)
        else:
            metrics_df = current_metrics
            
        # Sort metrics by date
        metrics_df['Date'] = pd.to_datetime(metrics_df['Date'])
        metrics_df = metrics_df.sort_values('Date', ascending=True)
        metrics_df['Date'] = metrics_df['Date'].dt.strftime('%Y-%m-%d')
        
        # Sort papers columns chronologically
        papers_df = papers_df.reindex(sorted(papers_df.columns), axis=1)

        # Reset index to Original_Title before saving
        papers_df.index = papers_df['Original_Title']
        papers_df = papers_df.drop('Original_Title', axis=1)
        papers_df.index.name = 'Title' # Ensure the index name is 'Title'
        
        # Save both sheets to Excel
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            papers_df.to_excel(writer, sheet_name='Citation Data')
            metrics_df.to_excel(writer, sheet_name='Metrics History', index=False)
        
        # Print summary
        print(f"\nTotal papers: {len(papers_df)}")
        print("\nMetrics for current crawl:")
        print(f"Total Citations: {total_citations}")
        print(f"h-index: {h_index}")
        print(f"i10-index: {i10_index}")
        
        # Send email summary
        recipient_email = "xxx@gmail.com"   #FIXME
        if recipient_email:
            send_metrics_email(
                recipient_email=recipient_email,
                total_papers=len(papers_df),
                total_citations=total_citations,
                h_index=h_index,
                i10_index=i10_index,
                author_name=author_name
            )
        else:
            print("\nNo recipient email address found in environment variables")
        print("----------------")

if __name__ == "__main__":
    main()
