import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TR Sheet Processor", layout="wide")
st.title("TR Sheet Mapper & Downloader (Fully Automated)")
st.info("💡 Yeh tool Program Structure sheet se Credit Structure aur Grand Total Structure ko accurately map karega.")

# ==========================================
# SINGLE FILE UPLOAD (TR Sheet with Structure)
# ==========================================
uploaded_file = st.file_uploader("Upload Main TR File (Excel containing Program Structure & Sem Sheet):", type=["xlsx", "xls"])

def calculate_credits_from_structure(struct_str):
    try:
        # Sum up parts of credit structure like "0-3-0" -> 3
        return sum(int(float(part)) for part in str(struct_str).strip().split('-') if part.strip().isdigit())
    except:
        try:
            return int(float(struct_str))
        except:
            return 0

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = [s.strip() for s in xls.sheet_names]
        
        # 1. Automatically find Program Structure sheet & TR Sheet
        prog_sheet_name = None
        tr_sheet_name = None
        
        for s in sheet_names:
            if 'program' in s.lower() or 'structure' in s.lower():
                prog_sheet_name = s
            elif s.lower() != 'program structure':
                tr_sheet_name = s 
                
        if not tr_sheet_name:
            tr_sheet_name = sheet_names[0]

        # Read Program Structure to build credit map automatically
        credit_map = {}
        grand_total_structure = ""
        
        if prog_sheet_name:
            # Read program structure
            prog_df = pd.read_excel(xls, sheet_name=prog_sheet_name, header=0)
            prog_df.columns = [str(c).strip().upper() for c in prog_df.columns]
            
            # Find sub code and credit structure columns dynamically
            code_col = next((c for c in prog_df.columns if 'CODE' in c or 'SUB' in c), None)
            struct_col = next((c for c in prog_df.columns if 'STRUCTURE' in c or 'CREDIT' in c), None)
            
            if code_col and struct_col:
                for _, row in prog_df.iterrows():
                    code_val = str(row[code_col]).strip()
                    struct_val = str(row[struct_col]).strip()
                    
                    if code_val.upper() == 'TOTAL' or 'TOTAL' in code_val.upper():
                        grand_total_structure = struct_val
                    elif pd.notna(row[code_col]) and code_val.lower() != 'nan':
                        credit_map[code_val.upper()] = struct_val

        # Read Main TR Sheet
        df = pd.read_excel(xls, sheet_name=tr_sheet_name, header=None)

        name_col, reg_col, sig_col = -1, -1, -1
        sgpa_col, percentage_col = -1, -1
        
        row0 = [str(x).strip().upper() for x in df.iloc[0].values]
        
        for i, val in enumerate(row0):
            if 'NAME' in val: name_col = i
            if 'REG' in val or 'ENROLL' in val: reg_col = i
            if ('ΣCIGI' in val or 'TOTAL' in val) and sig_col == -1: 
                if 'CIGI' in val or 'TOTAL' in val: sig_col = i
            if 'SGPA' in val: sgpa_col = i
            if 'PERCENTAGE' in val: percentage_col = i

        if reg_col == -1 or sig_col == -1:
            st.error("Error: TR sheet mein 'Reg' aur 'ΣCiGi' columns nahi mile.")
            st.stop()

        subjects = []
        current_col = 7 
        
        while current_col < sig_col:
            subj_name = str(df.iloc[0, current_col]).strip()
            subj_code = str(df.iloc[1, current_col]).strip()
            fallback_cred = str(df.iloc[2, current_col]).strip() 
            
            if subj_code and subj_code.lower() != 'nan':
                subjects.append({
                    'col_idx': current_col,
                    'name': subj_name,
                    'code': subj_code,
                    'fallback_cred': fallback_cred
                })
            current_col += 9 
            
        subject_count = len(subjects)
        calculated_grand_ttl_marks = subject_count * 100

        # --- CALCULATE TOTAL OBT CREDITS AUTOMATICALLY FROM MAPPED STRUCTURE ---
        total_credits_num = 0
        for subj in subjects:
            code_upper = subj['code'].upper()
            struct = credit_map.get(code_upper, subj['fallback_cred'])
            total_credits_num += calculate_credits_from_structure(struct)

        # If grand total structure wasn't explicitly found, fallback to generated sum string
        if not grand_total_structure and subject_count > 0:
            grand_total_structure = str(total_credits_num)

        # Base Columns for Truscholar Format
        base_columns = [
            'NAME', 'EMAIL', 'ENROLLMENT_NO', 'HALL_ADMIT_NO', 'EXAM_CENTER', 
            'EXAM_DATE', 'RESULT_DATE', 'DEGREE_DATE', 'GRAND_TTL_MARKS', 
            'TOTAL_OBT_MARKS', 'GRAND_TTL_CREDITS', 'TOTAL_OBT_CREDITS', 
            'CREDIT_POINT', 'GRADE_LETTER', 'GRADE_POINT', 'PERCENTAGE', 
            'PERCENTILE', 'DIVISION', 'DIVISION_IN_REGIONAL', 'SGPA', 
            'CGPA', 'RANK', 'RESULT', 'CERT_NO', 'REMARKS'
        ]

        processed_data = []

        for idx, row in df.iterrows():
            if idx < 4: continue 
            
            reg_no = str(row[reg_col]).strip()
            if not reg_no or reg_no.lower() == 'nan': continue

            tr_name = str(row[name_col]).strip() if name_col != -1 else ""
            if (not tr_name or tr_name.lower() == 'nan') and (not reg_no or reg_no.lower() == 'nan'):
                continue 

            new_row = {col: "" for col in base_columns}

            new_row["ENROLLMENT_NO"] = reg_no
            new_row["NAME"] = tr_name
            new_row["EMAIL"] = ""
            new_row["HALL_ADMIT_NO"] = ""

            # --- MAP GRAND TOTAL CREDITS FROM PROGRAM STRUCTURE SHEET ---
            new_row["GRAND_TTL_CREDITS"] = grand_total_structure 
            new_row["GRAND_TTL_MARKS"] = str(calculated_grand_ttl_marks) if calculated_grand_ttl_marks > 0 else ""
            
            new_row["TOTAL_OBT_MARKS"] = "" 
            new_row["TOTAL_OBT_CREDITS"] = str(total_credits_num) if total_credits_num > 0 else ""
            
            # --- MULTIPLY TOTAL CREDITS BY 10 FOR CREDIT_POINT ---
            new_row["CREDIT_POINT"] = str(total_credits_num * 10) if total_credits_num > 0 else ""
            
            new_row["SGPA"] = str(row[sgpa_col]) if sgpa_col != -1 and pd.notna(row[sgpa_col]) else ""
            
            if percentage_col != -1 and pd.notna(row[percentage_col]):
                try:
                    new_row["PERCENTAGE"] = f"{float(row[percentage_col]):.1f}"
                except:
                    new_row["PERCENTAGE"] = str(row[percentage_col])

            for i, subj in enumerate(subjects):
                num = str(i+1).zfill(2)
                c = subj['col_idx']
                code_upper = subj['code'].upper()
                
                new_row[f"SUBJ_NAME__{num}"] = subj['name']
                new_row[f"SUBJ_CODE__{num}"] = subj['code']
                new_row[f"MAX_MARKS_TH__{num}"] = "100" 
                new_row[f"MIN_MARKS_TH__{num}"] = "" 
                new_row[f"OBT_MARKS_TH__{num}"] = "" 
                
                # --- LOOKUP SUBJECT-WISE CREDIT STRUCTURE FROM PROGRAM STRUCTURE SHEET ---
                new_row[f"MAX_CREDS_TH__{num}"] = credit_map.get(code_upper, subj['fallback_cred'])
                
                new_row[f"OBT_CREDS_TH__{num}"] = str(row[c+3]) if pd.notna(row[c+3]) else "" 
                new_row[f"CRD_POINT_TH__{num}"] = str(row[c+8]) if pd.notna(row[c+8]) else "" # CiGi
                new_row[f"GRD_LETTR_TH__{num}"] = str(row[c+6]) if pd.notna(row[c+6]) else "" 
                new_row[f"GRD_POINT_TH__{num}"] = str(row[c+5]) if pd.notna(row[c+5]) else "" 
                new_row[f"RESULT_TH__{num}"] = "" 
                new_row[f"REMARKS_TH__{num}"] = ""

            processed_data.append(new_row)

        if not processed_data:
            st.error("No valid student data found.")
        else:
            final_df = pd.DataFrame(processed_data)
            st.success(f"Processing Complete! Successfully mapped structure for **{subject_count} Subjects** & **{len(final_df)} Students**.")
            
            st.info(f"✅ Grand Total Credit Structure: **{grand_total_structure}** | Total Obt Credits: **{total_credits_num}** | Credit Point: **{total_credits_num * 10}**")
            
            st.subheader("Data Preview")
            st.dataframe(final_df)

            st.subheader("File Rename & Download")
            output_name = st.text_input("Rename your file here:", value=f"Final_Truscholar_{subject_count}_Subjects_Auto")
            if not output_name.endswith('.xlsx'):
                output_name += '.xlsx'

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name=f"{subject_count} Subjects")
            
            st.download_button(
                label=f"⬇️ Download Final {subject_count} Subjects Excel",
                data=output.getvalue(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred: {e}")
