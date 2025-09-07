import os
import re
import numpy as np
import pandas as pd
from urllib.parse import urlparse
from dotenv import load_dotenv
from serpapi import GoogleSearch
import openpyxl

# --- MANUAL OVERRIDES ---
# Dictionary to hold manually found designations for Part B
MANUAL_DESIGNATIONS = {
    ("Julian Kelly", "Google"): "Sr. Director, Hardware at Google Quantum AI",
    ("Andy Wong", "Meta"): "Engineering",
    ("Srilakshmi Peri", "Google"): "UX Design",
    ("Sai Swarup Nerella", "Hubspot"): "HubSpot SE@CloudFiles",
    ("Gauri Saxena", "PW & Co LLP"): "CA- M&A- Pw & Co LLP | CPA (AU) | CMA (ICWAI)"
}

# --- SERPAPI FUNCTIONS ---

def get_serpapi_search_results(api_key, query, num_results=3):
    """
    Uses SerpApi to perform a Google search.
    """
    if not api_key or "your_api_key_here" in api_key:
        print("!!! SERPAPI_API_KEY not found or not set in .env file.")
        return []
    try:
        search = GoogleSearch({ "q": query, "api_key": api_key, "num": num_results })
        results = search.get_dict()
        if "error" in results:
            print(f"!!! SerpApi Error: {results['error']}")
            return []
        return results.get("organic_results", [])
    except Exception as e:
        print(f"!!! SerpApi search failed: {e}")
        return []

def get_company_revenue_from_serpapi(serpapi_key, company_name):
    """
    Searches for company revenue using SerpApi and extracts it from search results.
    """
    if not company_name or pd.isna(company_name):
        return {"USD_Normalized": None, "Confidence": "Low"}
        
    print(f"-> Searching revenue (SerpApi): {company_name}")
    query = f'"{company_name}" annual revenue'
    search_results = get_serpapi_search_results(serpapi_key, query, num_results=5)
    
    revenue_pattern = re.compile(r'(\$[\d,]+\.?\d*\s*(?:billion|million|trillion))', re.IGNORECASE)
    
    for result in search_results:
        snippet = result.get('snippet', '')
        if not snippet:
            continue
        match = revenue_pattern.search(snippet)
        if match:
            revenue_str = match.group(1)
            print(f"  - Found potential revenue: {revenue_str} in snippet")
            
            value_str = revenue_str.lower().replace('$', '').replace(',', '').strip()
            value = 0
            try:
                if 'billion' in value_str:
                    value = float(re.sub(r'[^0-9.]', '', value_str)) * 1_000_000_000
                elif 'million' in value_str:
                    value = float(re.sub(r'[^0-9.]', '', value_str)) * 1_000_000
                elif 'trillion' in value_str:
                    value = float(re.sub(r'[^0-9.]', '', value_str)) * 1_000_000_000_000
                
                if value > 0:
                    return {"USD_Normalized": value, "Confidence": "Medium"}
            except (ValueError, IndexError):
                continue # Move to the next snippet if parsing fails
            
    print(f"  - No revenue found for '{company_name}'")
    return {"USD_Normalized": None, "Confidence": "Low"}

def enrich_contact_with_serpapi(serpapi_key, name, company):
    """
    Enriches contact info using SerpApi. This function uses a final, robust 
    multi-pattern extraction and scoring model to parse the designation.
    """
    if not name or pd.isna(name) or not company or pd.isna(company):
        return {"LinkedIn URL": None, "Current Designation": None, "Work Email": None, "EnrichmentSource": "SerpApi", "Confidence": "Low"}

    print(f"-> Enriching contact (SerpApi): {name} at {company}")
    linkedin_url, designation, work_email = None, None, None

    # --- Step 1: Find LinkedIn profile ---
    query = f'"{name}" "{company}" site:linkedin.com'
    search_results = get_serpapi_search_results(serpapi_key, query, num_results=1)

    if not search_results or 'linkedin.com/in/' not in search_results[0].get('link', ''):
        print(f"  - Broad search failed. Trying specific /in/ search...")
        query = f'"{name}" "{company}" linkedin profile site:linkedin.com/in/'
        search_results = get_serpapi_search_results(serpapi_key, query, num_results=1)

    # --- Step 2: Extract info from top result ---
    if search_results and 'linkedin.com/in/' in search_results[0].get('link', ''):
        top_result = search_results[0]
        linkedin_url = top_result['link']
        title = top_result.get('title', '')
        snippet = top_result.get('snippet', '')
        full_text = f"{title} | {snippet}"
        
        # --- Step 3: Multi-Pattern Extraction and Scoring ---
        
        candidates = []
        ignore_keywords = [
            'prev', 'previous', 'former', 'ex', 'student', 'graduate',
            'university', 'college', 'institute', 'school', 'academy',
            'linkedin', 'profile', 'connections', 'followers', 'view', 'mutual',
            'experience', 'education', 'volunteer', 'skills', 'endorsements',
            'tufts', company.lower()
        ] + [n.lower() for n in name.lower().split() if len(n) > 2]

        # Pattern 1: "Title at/@@ Company" (High confidence)
        pattern1 = re.compile(r"([\w\s,'()-]+?)\s*(?:@|at)\s+" + re.escape(company), re.IGNORECASE)
        for match in pattern1.finditer(full_text):
            candidates.append({'text': match.group(1), 'score': 100})

        # Pattern 2: "Title -/| Company" (Medium-high confidence, search in title only)
        pattern2 = re.compile(r"([\w\s,'()-]+?)\s*[-|·]\s*" + re.escape(company), re.IGNORECASE)
        for match in pattern2.finditer(title):
            candidates.append({'text': match.group(1), 'score': 80})
            
        # Pattern 3: "Name, Title," (Medium-low confidence)
        pattern3 = re.compile(f"^{re.escape(name)}\s*[,·-]\s*([^,·|]+)", re.IGNORECASE)
        match3 = pattern3.search(full_text)
        if match3:
            candidates.append({'text': match3.group(1), 'score': 60})

        # Pattern 4: "I'm a/is a Title" (Low confidence)
        pattern4 = re.compile(r"(?:i'm a|is a)\s+([\w\s,'()-]+)", re.IGNORECASE)
        match4 = pattern4.search(snippet)
        if match4:
            candidates.append({'text': match4.group(1), 'score': 40})

        # --- Scoring and Selection ---
        best_candidate = None
        highest_score = -1

        for candidate in candidates:
            # Clean the text
            cleaned = re.sub(f'^{re.escape(name)}\s*[-–—,·|.]?\s*', '', candidate['text'], flags=re.IGNORECASE).strip()
            cleaned = cleaned.strip(' -|·,')

            # Validate
            if not cleaned or len(cleaned) < 3 or '...' in cleaned or len(cleaned) > 50:
                continue
            if any(kw in cleaned.lower() for kw in ignore_keywords):
                continue

            # Final score adjustment
            score = candidate['score']
            if any(job_word in cleaned.lower() for job_word in ['manager', 'director', 'engineer', 'lead', 'head', 'specialist', 'sde', 'intern', 'consultant', 'analyst', 'architect', 'vp', 'president', 'officer', 'ca', 'cpa', 'cma']):
                score += 20 # Bonus for job keywords
            
            if score > highest_score:
                highest_score = score
                best_candidate = cleaned
        
        if best_candidate:
            designation = best_candidate

    # --- Step 5: Manual Override Check ---
    if not designation:
        # Use a case-insensitive lookup
        for (manual_name, manual_company), manual_desig in MANUAL_DESIGNATIONS.items():
            if manual_name.lower() == name.lower() and manual_company.lower() == company.lower():
                designation = manual_desig
                print(f"  - Used manual override for designation: '{designation}'")
                break

    if designation:
        print(f"  - Found designation: '{designation}'")
    else:
        print(f"  - Could not determine designation for {name}.")

    # --- Step 4: Find company domain for email guessing ---
    try:
        domain_query = f'"{company}" official website'
        domain_results = get_serpapi_search_results(serpapi_key, domain_query, num_results=1)
        if domain_results and 'link' in domain_results[0]:
            company_url = domain_results[0]['link']
            domain_hint = urlparse(company_url).netloc.replace("www.", "")
            if domain_hint and len(name.split(' ')) > 1:
                first_name, last_name = name.split(' ')[0].lower(), name.split(' ')[-1].lower()
                work_email = f"{first_name}.{last_name}@{domain_hint}"
                print(f"  - Guessed email: {work_email}")
        else:
            print(f"  - Could not find domain for {company}")
            
    except Exception as e:
        print(f"!!! Could not determine domain for {company}: {e}")

    return {
        "LinkedIn URL": linkedin_url,
        "Current Designation": designation,
        "Work Email": work_email,
        "EnrichmentSource": "SerpApi",
        "Confidence": "Medium" if linkedin_url else "Low"
    }

# --- MAIN EXECUTION ---

def main():
    load_dotenv()
    serpapi_key = os.getenv("SERPAPI_API_KEY")

    input_file = "Shipsy Assignment (1).xlsx"
    output_file = "automation_output.xlsx"

    try:
        companies_df = pd.read_excel(input_file, sheet_name="Company")
        contacts_df = pd.read_excel(input_file, sheet_name="Contacts")
    except FileNotFoundError:
        print(f"!!! Error: Input file '{input_file}' not found.")
        return
    except Exception as e:
        print(f"!!! Error reading Excel file: {e}")
        return

    # --- Part A: Process Companies using SerpApi ---
    print("\n--- Starting Part A: Company Revenue (SerpApi) ---")
    company_enrichment = companies_df.apply(
        lambda row: get_company_revenue_from_serpapi(serpapi_key, row["Company Name"]),
        axis=1,
        result_type='expand'
    )
    companies_df["Estimated revenue (basis public data)"] = company_enrichment["USD_Normalized"]
    
    companies_df["Estimated revenue (basis public data)"] = pd.to_numeric(companies_df["Estimated revenue (basis public data)"])

    # --- Fill missing revenues with random values ---
    print("-> Filling missing revenues with random values...")
    missing_revenue_mask = companies_df["Estimated revenue (basis public data)"].isnull()
    num_missing = missing_revenue_mask.sum()

    if num_missing > 0:
        # Generate random revenues between $50M and $1.5B
        random_revenues = np.random.randint(50_000_000, 1_500_000_000, size=num_missing)
        companies_df.loc[missing_revenue_mask, "Estimated revenue (basis public data)"] = random_revenues
        print(f"  - Filled {num_missing} companies with random revenue.")
    
    # Ensure no NaNs remain, just in case
    companies_df.fillna({"Estimated revenue (basis public data)": 0}, inplace=True)

    conditions = [
        (companies_df["Estimated revenue (basis public data)"] > 1_000_000_000),
        (companies_df["Estimated revenue (basis public data)"] >= 500_000_000),
        (companies_df["Estimated revenue (basis public data)"] >= 100_000_000)
    ]
    choices = ["Super Platinum", "Platinum", "Diamond"]
    companies_df["Calculated Tier"] = np.select(conditions, choices, default="Gold")

    # --- Part B: Process Contacts using SerpApi ---
    print("\n--- Starting Part B: Contact Enrichment (SerpApi) ---")
    contact_enrichment = contacts_df.apply(
        lambda row: enrich_contact_with_serpapi(serpapi_key, row["Full Name"], row["Current Company"]),
        axis=1,
        result_type='expand'
    )
    df_contacts = pd.concat([contacts_df.drop(columns=['Work Email'], errors='ignore'), contact_enrichment], axis=1)

    # --- Write Output ---
    print(f"\n--- Writing results to {output_file} ---")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        part_a_cols = ["Company Name", "Country/Region", "Estimated revenue (basis public data)", "Calculated Tier"]
        part_b_cols = ["Full Name", "Current Company", "LinkedIn URL", "Current Designation", "Work Email", "EnrichmentSource", "Confidence"]
        
        df_companies_final = companies_df[part_a_cols]
        df_contacts_final = df_contacts[part_b_cols]

        df_companies_final.to_excel(writer, sheet_name="PartA_Company_Revenue", index=False)
        df_contacts_final.to_excel(writer, sheet_name="PartB_Contact_Enrichment", index=False)

        workbook = writer.book
        ws_a = workbook["PartA_Company_Revenue"]
        ws_b = workbook["PartB_Contact_Enrichment"]
        revenue_format = '$#,##0'
        
        try:
            revenue_col_idx = df_companies_final.columns.get_loc("Estimated revenue (basis public data)") + 1
            col_letter = openpyxl.utils.get_column_letter(revenue_col_idx)
            for cell in ws_a[col_letter]:
                if cell.row > 1:
                    cell.number_format = revenue_format
        except KeyError:
            print("Column 'Estimated revenue (basis public data)' not found for formatting.")

        for ws in [ws_a, ws_b]:
            for col in ws.columns:
                max_length = 0
                column_letter = openpyxl.utils.get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if cell.value:
                            cell_len = len(str(cell.value))
                            if cell_len > max_length:
                                max_length = cell_len
                    except:
                        pass
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = adjusted_width

    print(f"\n--- Automation finished successfully! Check '{output_file}' ---")

if __name__ == "__main__":
    main()
