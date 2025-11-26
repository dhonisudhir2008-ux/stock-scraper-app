import streamlit as st
import pandas as pd
import yfinance as yf
import requests
from bs4 import BeautifulSoup
import re
import time
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Stock Scraper Tool", layout="wide")

st.title("📊 Stock Financials Scraper")
st.markdown("Upload your Excel file containing a **'Stock Name'** column. This tool fetches data from Yahoo Finance & Screener.")

# --- 1. CENTRALIZED MAPPING TABLE ---
SYMBOL_MAP = {
    "INFOSYS": {"yfinance": "INFY.NS", "screener_name": "INFOSYS"}, 
    "HDFC": {"yfinance": "HDFCBANK.NS", "screener_name": "HDFCBANK"}, 
    "L&T": {"yfinance": "LT.NS", "screener_name": "LT"},
}

# --- 2. HELPER FUNCTIONS ---
def extract_table_value(soup, table_id, row_name):
    try:
        table = soup.select_one(f'section#{table_id} table')
        if table is None: return "N/A"
        
        row_td = table.find('td', string=re.compile(r'\b' + re.escape(row_name) + r'\b', re.IGNORECASE))
        
        if row_td:
            data_cols = row_td.parent.find_all('td')
            if data_cols:
                value = data_cols[-1].text.strip().replace(',', '').replace('-', '').replace(' ', '')
                if re.match(r'^-?\d+(\.\d+)?$', value):
                    return float(value)
        return "N/A"
    except Exception:
        return "N/A"

def scrape_screener_data_minimal(stock_name, screener_name):
    formatted_name = screener_name.replace(" ", "-").replace("&", "%26")
    screener_url = f"https://www.screener.in/company/{formatted_name}/consolidated/"
    
    headers = {'User-Agent': 'Mozilla/5.0'}
    data = {"P/E Ratio": "N/A", "Interest (Cr.)": "N/A"}

    try:
        response = requests.get(screener_url, headers=headers, timeout=10)
        response.raise_for_status() 
        soup = BeautifulSoup(response.content, 'html.parser')

        # Extract P/E
        ratio_list_ul = soup.select_one('div.company-ratios > ul') 
        if ratio_list_ul:
            summary_points = ratio_list_ul.find_all('li')
            for li in summary_points:
                name_tag = li.find('span', class_='name')
                value_tag = li.find('span', class_='value')
                if name_tag and value_tag:
                    name = name_tag.text.strip().replace('\n', ' ').upper()
                    value = value_tag.text.strip().replace('₹', '').replace('%', '').replace(',', '')
                    is_valid_number = re.match(r'^-?\d+(\.\d+)?$', value)
                    if "P/E" in name:
                        data["P/E Ratio"] = float(value) if is_valid_number else "N/A"
        
        # Extract Interest
        data["Interest (Cr.)"] = extract_table_value(soup, 'profit-loss', 'Interest')

    except Exception as e:
        # We handle logging in the main loop now
        pass
        
    return data

# --- 3. FRONTEND LOGIC ---

# A. File Uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read the file into Pandas
    try:
        df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully!")
        st.dataframe(df.head()) # Show preview
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    # B. The "Run" Button
    if st.button("🚀 Start Analysis"):
        if 'Stock Name' not in df.columns:
            st.error("The file must have a column named 'Stock Name'.")
        else:
            results = []
            progress_bar = st.progress(0)
            status_text = st.empty() # Placeholder for text updates
            
            # Create a scrollable container for logs so the page doesn't get too long
            with st.status("Processing Stocks...", expanded=True) as status:
                
                total_rows = len(df)
                
                for index, row in df.iterrows():
                    # Update progress
                    progress = (index + 1) / total_rows
                    progress_bar.progress(progress)
                    
                    stock_name = str(row["Stock Name"]).strip().upper()
                    if not stock_name: continue

                    # Mapping Logic
                    map_entry = SYMBOL_MAP.get(stock_name)
                    if map_entry:
                        full_symbol = map_entry["yfinance"]
                        screener_name_for_url = map_entry["screener_name"]
                    else:
                        full_symbol = stock_name.replace(" ", "") + ".NS" 
                        screener_name_for_url = stock_name
                    
                    st.write(f"🔄 Processing **{stock_name}**...")

                    # --- YFinance ---
                    current_price = "N/A"
                    market_cap_cr = "N/A" 
                    try:
                        ticker = yf.Ticker(full_symbol)
                        data_yf = ticker.info
                        current_price = data_yf.get('currentPrice', 'N/A')
                        raw_market_cap = data_yf.get('marketCap', 'N/A')
                        if isinstance(raw_market_cap, (int, float)):
                             market_cap_cr = round(raw_market_cap / 10000000, 2)
                    except Exception:
                        st.warning(f"⚠️ YFinance failed for {full_symbol}")

                    # --- Screener ---
                    screener_data = scrape_screener_data_minimal(stock_name, screener_name_for_url)
                    
                    # --- Compile ---
                    current_result = {
                        "Stock Name": row["Stock Name"], 
                        "YF Symbol Used": full_symbol,
                        "Current Price (Rs.)": current_price,
                        "Market Cap (Cr.)": market_cap_cr,
                        "P/E Ratio": screener_data["P/E Ratio"],
                        "Interest (Cr.)": screener_data["Interest (Cr.)"], 
                    }
                    
                    # Merge with original data
                    for col in df.columns:
                        if col not in current_result: 
                            current_result[col] = row[col]

                    results.append(current_result)
                    time.sleep(1) # Be polite to the server
                
                status.update(label="Analysis Complete!", state="complete", expanded=False)

            # --- 4. EXPORT RESULTS ---
            if results:
                df_results = pd.DataFrame(results)
                
                # Reorder columns
                final_new_cols = ["Stock Name", "YF Symbol Used", "Current Price (Rs.)", 
                                  "Market Cap (Cr.)", "P/E Ratio", "Interest (Cr.)"]
                original_cols = [col for col in df.columns if col not in final_new_cols]
                df_results = df_results[final_new_cols + original_cols]

                st.subheader("🎉 Results")
                st.dataframe(df_results)

                # Convert to Excel in memory (No saving to C drive needed)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_results.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Download Updated Excel",
                    data=buffer.getvalue(),
                    file_name="stocks_updated.xlsx",
                    mime="application/vnd.ms-excel"
                )
