import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TR Sheet Processor", layout="wide")
st.title("TR Sheet Mapper & Downloader (Pro Version)")
st.info("💡 Yeh tool 3 files ka data combine karke exact Truscholar format generate karega.")

# --- Helper Function for Demo Files ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# ==========================================
# STEP 1 & 2: SUPPLEMENTARY FILES UPLOAD
# ==========================================
st.header("Step 1 & 2: Upload Mapping Files")
col1, col2 = st.columns(2)

student_map = {}
credit_map = {}
grand_total_structure = ""

with col1:
    st.subheader("1. Student Details")
    demo_student_df = pd.DataFrame({
        "Enrollment No": ["240810500024", "240810500025"], 
        "Student Name": ["Aasmant Umesh Mohadikar", "Rahul Kumar"], 
        "Email ID": ["student@example.com", "rahul@example.com"], 
        "Hall Admit No": ["HALL123", "HALL124"]
    })
    st.download_button("⬇️ Download Demo (Student Details)", data=to_excel(demo_student_df), file_name="Demo_Student_Details.xlsx")
    student_file = st.file_uploader("Upload Student Details File:", type=["xlsx", "xls"], key="stu")

    if student_file:
        try:
            sdf = pd.read_excel(student_file)
            sdf.columns = [str(c).strip().upper() for c in sdf.columns]
            
            enr_col = [c for c in sdf.columns if 'ENROLL' in c or 'ROLL' in c or 'REG' in c]
            name_col_stu = [c for c in sdf.columns if 'NAME' in c]
            email_col = [c for c in sdf.columns if 'EMAIL' in c]
            hall_col = [c for c in sdf.columns if 'HALL' in c or 'ADMIT' in c]
            
            if enr_col:
                for _, row in sdf.iterrows():
                    reg = str(row[enr_col[0]]).strip()
                    stu_name = str(row[name_col_stu[0]]).strip() if name_col_stu and pd.notna(row[name_col_stu[0]]) else ""
                    email = str(row[email_col[0]]).strip() if email_col and pd.notna(row[email_col[0]]) else ""
                    hall = str(row[hall_col[0]]).strip() if hall_col and pd.notna(row[hall_col[0]]) else ""
                    
                    student_map[reg] = {"email": email, "hall": hall, "name": stu_name}
                st.success("✅ Student Details mapped successfully!")
            else:
                st.error("Enrollment No column not found in Student Details file.")
        except Exception as e:
            st.error(f"Error reading Student file: {e}")

with col2:
    st.subheader("2. Credit Structure")
    demo_credit_df = pd.DataFrame({
        "Sub Code": ["AVP204", "AVP205", "AVP207", "Total"], 
        "Credit Structure": ["0-3-0", "0-3-0", "3-0-0", "03-10-08"]
    })
    st.download_button("⬇️ Download Demo (Credit Structure)", data=to_excel(demo_credit_df), file_name="Demo_Credit_Structure.xlsx")
    credit_file = st.file_uploader("Upload Credit Structure File:", type=["xlsx", "xls"], key="cred")

    if credit_file:
        try:
            cdf = pd.read_excel(credit_file)
            cdf.columns = [str(c).strip().upper() for c in cdf.columns]
            
            code_col = [c for c in cdf.columns if 'CODE' in c or 'SUB' in c]
            struct_col = [c for c in cdf.columns if 'CREDIT' in c or 'STRUCTURE' in c]
            
            if code_col and struct_col:
                for _, row in cdf.iterrows():
                    code = str(row[code_col[0]]).strip().upper()
                    struct = str(row[struct_col[0]]).strip()
                    if code == 'TOTAL':
                        grand_total_structure = struct
                    elif pd.notna(row[code_col[0]]):
                        credit_map[code] = struct
                st.success("✅ Credit Structure mapped successfully!")
            else:
                st.error("Sub Code or Credit Structure column not found.")
        except Exception as e:
            st.error(f"Error reading Credit file: {e}")

# ==========================================
# STEP 3: MAIN TR SHEET PROCESSING
# ==========================================
st.header("Step 3: Upload Main TR Sheet")
uploaded_file = st.file_uploader("Upload Final TR Sheet (Excel):", type=["xlsx", "xls"])

def calculate_credits_from_structure(struct_str):
    try:
        return sum(int(float(part)) for part in str(struct_str).strip().split('-') if part.strip().isdigit())
    except:
        return 0

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)

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
            st.error("Error: TR sheet mein 'Reg' aur 'ΣCiGi' columns hona zaroori hai.")
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

        total_credits_num = 0
        for subj in subjects:
            struct = credit_map.get(subj['code'].upper(), subj['fallback_cred'])
            total_credits_num += calculate_credits_from_structure(struct) if '-' in str(struct) else int(float(struct)) if str(struct).isdigit() else 0

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
            
            mapped_name = student_map.get(reg_no, {}).get("name", "")
            new_row["NAME"] = mapped_name if mapped_name else tr_name
            new_row["EMAIL"] = student_map.get(reg_no, {}).get("email", "")
            new_row["HALL_ADMIT_NO"] = student_map.get(reg_no, {}).get("hall", "")

            new_row["GRAND_TTL_CREDITS"] = grand_total_structure 
            new_row["GRAND_TTL_MARKS"] = str(calculated_grand_ttl_marks) if calculated_grand_ttl_marks > 0 else ""
            
            new_row["TOTAL_OBT_MARKS"] = "" 
            new_row["TOTAL_OBT_CREDITS"] = str(total_credits_num) if total_credits_num > 0 else ""
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
                
                new_row[f"MAX_CREDS_TH__{num}"] = credit_map.get(code_upper, subj['fallback_cred'])
                
                new_row[f"OBT_CREDS_TH__{num}"] = str(row[c+3]) if pd.notna(row[c+3]) else "" # Total (100) Marks
                
                # --- UPDATED: CRD_POINT_TH mapped to CiGi column (Offset c+8) ---
                new_row[f"CRD_POINT_TH__{num}"] = str(row[c+8]) if pd.notna(row[c+8]) else "" # CiGi Marks
                
                new_row[f"GRD_LETTR_TH__{num}"] = str(row[c+6]) if pd.notna(row[c+6]) else "" # Letter Grade
                new_row[f"GRD_POINT_TH__{num}"] = str(row[c+5]) if pd.notna(row[c+5]) else "" # Grade Point
                new_row[f"RESULT_TH__{num}"] = "" 
                new_row[f"REMARKS_TH__{num}"] = ""

            processed_data.append(new_row)

        if not processed_data:
            st.error("No valid student data found.")
        else:
            final_df = pd.DataFrame(processed_data)
            st.success(f"Processing Complete! Successfully combined data for **{subject_count} Subjects** & **{len(final_df)} Students**.")
            
            st.info("✅ **CRD_POINT_TH** column now maps directly from the **CiGi** column in the TR sheet.")
            
            st.subheader("Data Preview")
            st.dataframe(final_df)

            st.subheader("File Rename & Download")
            output_name = st.text_input("Rename your file here:", value=f"Final_Truscholar_{subject_count}_Subjects_Mapped")
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
