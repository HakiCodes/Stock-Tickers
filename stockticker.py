import pandas as pd
from fuzzywuzzy import fuzz
import re

# Load the spreadsheet
df = pd.read_excel("C:\\Users\\nokol\\OneDrive\\Documents\\Stock ticker project.xlsx", engine="openpyxl")

# Function to preprocess names
def preprocess_name(name):
    # Convert to lowercase, remove punctuation, and strip whitespace
    return re.sub(r'\W+', ' ', name).lower().strip()

# Preprocess company names and store in a dictionary
company_dict = {row['Stock Symbol']: preprocess_name(row['Company Name']) for index, row in df.iterrows()}

# Function to extract the most likely stock ticker from a headline
def extract_ticker(headline):
    # Preprocess the headline
    processed_headline = preprocess_name(headline)
    
    # Check if any ticker symbol appears in the headline
    for ticker, name in company_dict.items():
        if ticker.lower() in processed_headline:
            return ticker
    
    # Initialize variables to store the most likely ticker and its score
    most_likely_ticker = None
    highest_score = 0
    
    # Find the most likely stock ticker based on the company name
    for ticker, name in company_dict.items():
        # Calculate similarity score between headline and company name
        similarity_score = fuzz.token_set_ratio(processed_headline, name)
        if similarity_score > highest_score:
            most_likely_ticker = ticker
            highest_score = similarity_score
    
    return most_likely_ticker

# Example usage
headline = "Innovative Launch Boosts Adobe Shares to Record Highs"
most_likely_ticker = extract_ticker(headline)
print("Company Ticker:", most_likely_ticker)
