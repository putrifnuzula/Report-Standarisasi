import pandas as pd
import streamlit as st
from io import BytesIO

# Claim data functions
def filter_claim_data(df):
    return df[df['ClaimStatus'] == 'R']

def remove_duplicate_claims(df):
    dups = df[df.duplicated(subset='ClaimNo', keep=False)]
    if not dups.empty:
        st.write("Duplicated ClaimNo values:")
        st.dataframe(dups[['ClaimNo']].drop_duplicates())
    return df.drop_duplicates(subset='ClaimNo', keep='last')

def process_claim_data(df):
    df = filter_claim_data(df)
    df = remove_duplicate_claims(df)
    for col in ["TreatmentStart", "TreatmentFinish", "Date"]:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        if df[col].isnull().any():
            st.warning(f"Invalid date values detected in column '{col}'. Coerced to NaT.")
            
    # Build standardized template:
    return pd.DataFrame({
        "No": range(1, len(df) + 1),
        "Policy No": df["PolicyNo"],
        "Client Name": df["ClientName"],
        "Claim No": df["ClaimNo"],
        "Member No": df["MemberNo"],
        "Emp ID": df["EmpID"],
        "Emp Name": df["EmpName"],
        "Patient Name": df["PatientName"],
        "Membership": df["Membership"],
        "Product Type": df["ProductType"],
        "Claim Type": df["ClaimType"],
        "Room Option": df["RoomOption"].fillna('').astype(str).str.upper().str.replace(r"\s+", "", regex=True),
        "Area": df["Area"],
        "Plan": df["PPlan"],
        "Diagnosis": df["PrimaryDiagnosis"].str.upper(),
        "Treatment Place": df["TreatmentPlace"].str.upper(),
        "Treatment Start": df["TreatmentStart"],
        "Treatment Finish": df["TreatmentFinish"],
        "Settled Date": df["Date"],
        "Year": df["Date"].dt.year,
        "Month": df["Date"].dt.month,
        "Length of Stay": df["LOS"],
        "Sum of Billed": df["Billed"],
        "Sum of Accepted": df["Accepted"],
        "Sum of Excess Coy": df["ExcessCoy"],
        "Sum of Excess Emp": df["ExcessEmp"],
        "Sum of Excess Total": df["ExcessTotal"],
        "Sum of Unpaid": df["Unpaid"]
    })

# Benefit data functions
def filter_benefit_data(df):
    if 'Status_Claim' in df.columns:
        return df[df['Status_Claim'] == 'R']
    elif 'Status Claim' in df.columns:
        return df[df['Status Claim'] == 'R']
    else:
        st.warning("Column 'Status Claim' not found. Data not filtered.")
        return df

def process_benefit_data(df):
    df = filter_benefit_data(df)
    df.columns = df.columns.str.strip()
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
    rename_mapping = {
        'ClientName': 'Client Name',
        'PolicyNo': 'Policy No',
        'ClaimNo': 'Claim No',
        'MemberNo': 'Member No',
        'PatientName': 'Patient Name',
        'EmpID': 'Emp ID',
        'EmpName': 'Emp Name',
        'ClaimType': 'Claim Type',
        'TreatmentPlace': 'Treatment Place',
        'RoomOption': 'Room Option',
        'TreatmentRoomClass': 'Treatment Room Class',
        'TreatmentStart': 'Treatment Start',
        'TreatmentFinish': 'Treatment Finish',
        'ProductType': 'Product Type',
        'BenefitName': 'Benefit Name',
        'PaymentDate': 'Payment Date',
        'ExcessTotal': 'Excess Total',
        'ExcessCoy': 'Excess Coy',
        'ExcessEmp': 'Excess Emp'
    }
    df = df.rename(columns=rename_mapping)
    return df.drop(columns=["Status_Claim", "BAmount"], errors='ignore')

# Save to excel
def save_to_excel(claim_df, benefit_df, summary_top_df, claim_ratio_df, filename):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        workbook.formats[0].set_font_name('VAG Rounded Std Light')
        
        # Define excel formats:
        bold_border = workbook.add_format({'bold': True, 'border': 1, 'font_name': 'VAG Rounded Std Light'})
        plain_border = workbook.add_format({'border': 1, 'font_name': 'VAG Rounded Std Light'})
        header_border = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'font_name': 'VAG Rounded Std Light'})
        
        # Summary sheet:
        summary_sheet = workbook.add_worksheet("Summary")
        summary_sheet.hide_gridlines(2)
        row = 0
        # Write summary statistics:
        for _, data in summary_top_df.iterrows():
            summary_sheet.write(row, 0, data["Metric"], bold_border)
            summary_sheet.write(row, 1, data["Value"], plain_border)
            row += 1
        # Insert blank row without borders (seperate sum stats & CR)
        summary_sheet.write(row, 0, "")
        summary_sheet.write(row, 1, "")
        row += 1
        
        # Write header for Claim Ratio table
        cr_columns = ["Company", "Net Premi", "Billed", "Unpaid", "Excess Total", "Excess Coy", "Excess Emp", "Claim", "CR", "Est Claim"]
        for col, header in enumerate(cr_columns):
            summary_sheet.write(row, col, header, header_border)
        row += 1
        
        # Write Claim Ratio data rows
        for _, data in claim_ratio_df.iterrows():
            for col, header in enumerate(cr_columns):
                summary_sheet.write(row, col, data.get(header, ""), plain_border)
            row += 1
        
        # Claim (SC) sheet
        claim_df.to_excel(writer, index=False, sheet_name='SC')
        ws_claim = writer.sheets["SC"]
        ws_claim.hide_gridlines(2)
        rows_claim, cols_claim = claim_df.shape[0] + 1, claim_df.shape[1]
        ws_claim.conditional_format(0, 0, rows_claim - 1, cols_claim - 1,
                                     {'type': 'no_errors', 'format': plain_border})
        for col_num, value in enumerate(claim_df.columns.values):
            ws_benefit.write(0, col_num, value, header_border)
        
        # Benefit sheet
        benefit_df.to_excel(writer, index=False, sheet_name='Benefit')
        ws_benefit = writer.sheets["Benefit"]
        ws_benefit.hide_gridlines(2)
        rows_benefit, cols_benefit = benefit_df.shape[0] + 1, benefit_df.shape[1]
        ws_benefit.conditional_format(0, 0, rows_benefit - 1, cols_benefit - 1,
                                      {'type': 'no_errors', 'format': plain_border})
        for col_num, value in enumerate(benefit_df.columns.values):
            ws_benefit.write(0, col_num, value, header_border)

        
        writer.close()
    output.seek(0)
    return output, filename

# Streamlit APP UI
st.title("Template - Standardisasi Report")

uploaded_claim = st.file_uploader("Upload Claim Data", type=["csv"], key="claim")
uploaded_claim_ratio = st.file_uploader("Upload Claim Ratio Data", type=["xlsx"], key="claim_ratio")
uploaded_benefit = st.file_uploader("Upload Benefit Data", type=["csv"], key="benefit")

if uploaded_claim and uploaded_claim_ratio and uploaded_benefit:
    # Process claim data
    raw_claim = pd.read_csv(uploaded_claim)
    st.write("Processing Claim Data...")
    claim_transformed = process_claim_data(raw_claim)
    st.subheader("Claim Data Preview:")
    st.dataframe(claim_transformed.head())
    
    # Process claim ratio data
    claim_ratio_raw = pd.read_excel(uploaded_claim_ratio)
    policy_list = claim_transformed["Policy No"].unique().tolist()
    claim_ratio_filtered = claim_ratio_raw[claim_ratio_raw["Policy No"].isin(policy_list)]
    claim_ratio_unique = claim_ratio_filtered.drop_duplicates(subset="Policy No")
    desired_cols = ['Company', 'Net Premi', 'Billed', 'Unpaid', 
                    'Excess Total', 'Excess Coy', 'Excess Emp', 'Claim', 'CR', 'Est CR Total']
    missing_cols = [col for col in desired_cols if col not in claim_ratio_unique.columns]
    if missing_cols:
        st.warning(f"Missing columns in Claim Ratio Data: {missing_cols}")
    claim_ratio_unique = claim_ratio_unique[[col for col in desired_cols if col in claim_ratio_unique.columns]]
    claim_ratio_unique = claim_ratio_unique.rename(columns={'Est CR Total': 'Est Claim'})
    summary_cr_df = claim_ratio_unique[['Company', 'Net Premi', 'Billed', 'Unpaid', 
                                         'Excess Total', 'Excess Coy', 'Excess Emp', 'Claim', 'CR', 'Est Claim']]
    st.subheader("Claim Ratio Data Preview (unique by Policy No):")
    st.dataframe(summary_cr_df.head())

    # Process benefit data
    raw_benefit = pd.read_csv(uploaded_benefit)
    st.write("Processing Benefit Data...")
    benefit_transformed = process_benefit_data(raw_benefit)
    claim_no_list = claim_transformed["Claim No"].unique().tolist()
    if "ClaimNo" in benefit_transformed.columns:
        benefit_transformed = benefit_transformed[benefit_transformed["ClaimNo"].isin(claim_no_list)]
    elif "Claim No" in benefit_transformed.columns:
        benefit_transformed = benefit_transformed[benefit_transformed["Claim No"].isin(claim_no_list)]
    else:
        st.warning("Column 'ClaimNo' not found in Benefit data; skipping filtering based on ClaimNo.")
    st.subheader("Benefit Data Preview:")
    st.dataframe(benefit_transformed.head())
    
    # Prepare Summary Top Section (Claim Stats + Overall Claim Ratio)
    total_claims   = len(claim_transformed)
    total_billed   = int(claim_transformed["Sum of Billed"].sum())
    total_accepted = int(claim_transformed["Sum of Accepted"].sum())
    total_excess   = int(claim_transformed["Sum of Excess Total"].sum())
    total_unpaid   = int(claim_transformed["Sum of Unpaid"].sum())
    claim_summary_data = {
        "Metric": ["Total Claims", "Total Billed", "Total Accepted", "Total Excess", "Total Unpaid"],
        "Value": [f"{total_claims:,}", f"{total_billed:,.2f}", f"{total_accepted:,.2f}",
                  f"{total_excess:,.2f}", f"{total_unpaid:,.2f}"]
    }
    claim_summary_df = pd.DataFrame(claim_summary_data)
    
    if "Claim" in claim_ratio_unique.columns and "Net Premi" in claim_ratio_unique.columns:
        total_claim_ratio_claim = claim_ratio_unique["Claim"].sum()
        total_net_premi = claim_ratio_unique["Net Premi"].sum()
        overall_cr = (total_claim_ratio_claim / total_net_premi) * 100 if total_net_premi != 0 else 0
        claim_ratio_overall = pd.DataFrame({"Metric": ["Claim Ratio (%)"],
                                            "Value": [f"{overall_cr:.2f}%"]})
    else:
        claim_ratio_overall = pd.DataFrame({"Metric": ["Claim Ratio (%)"], "Value": ["N/A"]})
    
    summary_top_df = pd.concat([claim_summary_df, claim_ratio_overall], ignore_index=True)
    
    # Download the Excel file
    filename_input = st.text_input("Enter the Excel file name (without extension):", "SC & Benefit - - YTD")
    if filename_input:
        excel_file, final_filename = save_to_excel(claim_transformed, benefit_transformed,
                                                   summary_top_df, summary_cr_df, filename_input + ".xlsx")
        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
