import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TR Sheet Processor", layout="wide")
st.title("TR Sheet Mapper & Downloader (Pro Version V2)")
st.info("💡 Yeh tool automatically credit structure detect karega aur email column se data extract karega!")

# --- Helper Function for Demo Files ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# ==========================================
# STEP 1: SUPPLEMENTARY STUDENT FILE UPLOAD
# ==========================================
st.header("Step 1: Upload Student Details (Optional)")
st.write("⚠️ Yeh file optional hai. Agar column G mein email hai to automatically use hoga.")

student_map = {}

demo_student_df = pd.DataFrame({
    "Enrollment No": ["240810500024", "240810500025"], 
    "Student Name": ["Aasmant Umesh Mohadikar", "Rahul Kumar"], 
    "Email ID": ["student@example.com", "rahul@example.com"], 
    "Hall Admit No": ["HALL123", "HALL124"]
})
st.download_button("⬇️ Download Demo (Student Details)", data=to_excel(demo_student_df), file_name="Demo_Student_Details.xlsx", key="demo_stu")

student_file = st.file_uploader("Upload Student Details File:", type=["xlsx", "xls"], key="stu_upload")

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

# ==========================================
# STEP 2: MAIN TR SHEET PROCESSING
# ==========================================
st.header("Step 2: Upload Main TR Sheet (with all sheets)")
st.write("✅ This file should have: Main TR data sheet + Program Structure sheet")

uploaded_file = st.file_uploader("Upload Final TR Sheet (Excel with multiple sheets):", type=["xlsx", "xls"], key="main_upload")

def calculate_credits_from_structure(struct_str):
    """Extract total credits from structure string like '0-3-0' or '3-0-0'"""
    try:
        parts = str(struct_str).strip().split('-')
        return sum(int(float(part)) for part in parts if part.strip())
    except:
        return 0

def extract_credit_structure_from_sheet(df):
    """
    Automatically detect and extract credit structure from Program Structure sheet
    Returns: (grand_total_structure, credit_map)
    """
    credit_map = {}
    grand_total_structure = ""
    
    try:
        # Find columns dynamically
        headers = [str(x).strip().upper() for x in df.iloc[0].values]
        code_col_idx = -1
        struct_col_idx = -1
        name_col_idx = -1
        
        for i, h in enumerate(headers):
            if 'CODE' in h or 'SUB' in h:
                code_col_idx = i
            if 'STRUCTURE' in h or 'CREDIT' in h:
                struct_col_idx = i
            if 'NAME' in h or 'SUBJECT' in h:
                name_col_idx = i
        
        if code_col_idx == -1 or struct_col_idx == -1:
            return "", {}
        
        # Extract credit structure for each subject
        for idx, row in df.iterrows():
            if idx == 0:  # Skip header
                continue
            
            code = str(row[code_col_idx]).strip().upper()
            struct = str(row[struct_col_idx]).strip()
            
            if pd.isna(code) or code == '' or code == 'NAN':
                continue
            
            if code == 'TOTAL':
                grand_total_structure = struct
            else:
                credit_map[code] = struct
        
        return grand_total_structure, credit_map
    
    except Exception as e:
        st.warning(f"Could not auto-extract credit structure: {e}")
        return "", {}

if uploaded_file is not None:
    try:
        # ===== READ ALL SHEETS =====
        xls = pd.ExcelFile(uploaded_file)
        st.info(f"📋 **Sheets found in file:** {', '.join(xls.sheet_names)}")
        
        # Detect which sheet has the TR data and which has Program Structure
        tr_sheet_name = None
        program_sheet_name = None
        
        for sheet in xls.sheet_names:
            sheet_lower = sheet.lower()
            if 'program' in sheet_lower or 'structure' in sheet_lower:
                program_sheet_name = sheet
            elif 'sem' in sheet_lower or 'sem2' in sheet_lower or 'sem4' in sheet_lower:
                tr_sheet_name = sheet
        
        # If not detected, use first two sheets
        if not tr_sheet_name and len(xls.sheet_names) > 0:
            tr_sheet_name = xls.sheet_names[0]
        if not program_sheet_name and len(xls.sheet_names) > 1:
            program_sheet_name = xls.sheet_names[1]
        
        st.write(f"✅ Using sheet **'{tr_sheet_name}'** for TR data")
        if program_sheet_name:
            st.write(f"✅ Using sheet **'{program_sheet_name}'** for Credit Structure")
        
        # ===== READ TR SHEET =====
        df = pd.read_excel(xls, sheet_name=tr_sheet_name, header=None)
        
        # ===== READ PROGRAM STRUCTURE SHEET =====
        credit_map = {}
        grand_total_structure = ""
        
        if program_sheet_name:
            try:
                program_df = pd.read_excel(xls, sheet_name=program_sheet_name, header=None)
                grand_total_structure, credit_map = extract_credit_structure_from_sheet(program_df)
                st.success(f"✅ Auto-extracted credit structure! Grand Total: {grand_total_structure}")
                st.write(f"📊 Subject-wise credits: {credit_map}")
            except Exception as e:
                st.warning(f"Could not read Program Structure sheet: {e}")
        
        # ===== PARSE TR SHEET HEADERS =====
        name_col, reg_col, sig_col = -1, -1, -1
        sgpa_col, percentage_col = -1, -1
        email_col = 6  # Column G (0-indexed) is typically column 6
        
        row0 = [str(x).strip().upper() for x in df.iloc[0].values]
        
        for i, val in enumerate(row0):
            if 'NAME' in val: 
                name_col = i
            if 'REG' in val or 'ENROLL' in val: 
                reg_col = i
            if 'ΣCIGI' in val or ('TOTAL' in val and 'CIGI' in val): 
                sig_col = i
            if 'SGPA' in val: 
                sgpa_col = i
            if 'PERCENTAGE' in val: 
                percentage_col = i
        
        if reg_col == -1 or sig_col == -1:
            st.error("❌ Error: TR sheet mein 'Reg' aur 'ΣCiGi' columns hona zaroori hai.")
            st.stop()
        
        # ===== DETECT SUBJECTS =====
        subjects = []
        current_col = 7  # Start from column H (after column G which is email)
        
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
            
            current_col += 9  # Each subject block is 9 columns
        
        subject_count = len(subjects)
        calculated_grand_ttl_marks = subject_count * 100
        
        # Calculate total credits
        total_credits_num = 0
        for subj in subjects:
            code_upper = subj['code'].upper()
            struct = credit_map.get(code_upper, subj['fallback_cred'])
            total_credits_num += calculate_credits_from_structure(struct)
        
        st.success(f"✅ Detected **{subject_count} Subjects** | Total Credits: **{total_credits_num}**")
        
        # ===== CREATE OUTPUT COLUMNS =====
        base_columns = [
            'NAME', 'EMAIL', 'ENROLLMENT_NO', 'HALL_ADMIT_NO', 'EXAM_CENTER', 
            'EXAM_DATE', 'RESULT_DATE', 'DEGREE_DATE', 'GRAND_TTL_MARKS', 
            'TOTAL_OBT_MARKS', 'GRAND_TTL_CREDITS', 'TOTAL_OBT_CREDITS', 
            'CREDIT_POINT', 'GRADE_LETTER', 'GRADE_POINT', 'PERCENTAGE', 
            'PERCENTILE', 'DIVISION', 'DIVISION_IN_REGIONAL', 'SGPA', 
            'CGPA', 'RANK', 'RESULT', 'CERT_NO', 'REMARKS'
        ]
        
        processed_data = []
        
        # ===== PROCESS EACH STUDENT ROW =====
        for idx, row in df.iterrows():
            if idx < 4: 
                continue
            
            reg_no = str(row[reg_col]).strip() if reg_col >= 0 and pd.notna(row[reg_col]) else ""
            if not reg_no or reg_no.lower() == 'nan': 
                continue
            
            tr_name = str(row[name_col]).strip() if name_col >= 0 and pd.notna(row[name_col]) else ""
            
            # ===== EMAIL PRIORITY: Column G > Student Details File =====
            email_from_tr = str(row[email_col]).strip() if email_col >= 0 and pd.notna(row[email_col]) else ""
            email_from_map = student_map.get(reg_no, {}).get("email", "")
            final_email = email_from_tr if email_from_tr and email_from_tr.lower() != 'nan' else email_from_map
            
            new_row = {col: "" for col in base_columns}
            
            new_row["ENROLLMENT_NO"] = reg_no
            
            mapped_name = student_map.get(reg_no, {}).get("name", "")
            new_row["NAME"] = mapped_name if mapped_name else tr_name
            new_row["EMAIL"] = final_email  # Priority: TR column G > Student Details
            new_row["HALL_ADMIT_NO"] = student_map.get(reg_no, {}).get("hall", "")
            
            new_row["GRAND_TTL_CREDITS"] = grand_total_structure
            new_row["GRAND_TTL_MARKS"] = str(calculated_grand_ttl_marks) if calculated_grand_ttl_marks > 0 else ""
            new_row["TOTAL_OBT_MARKS"] = ""
            new_row["TOTAL_OBT_CREDITS"] = str(total_credits_num) if total_credits_num > 0 else ""
            new_row["CREDIT_POINT"] = str(total_credits_num * 10) if total_credits_num > 0 else ""
            
            new_row["SGPA"] = str(row[sgpa_col]) if sgpa_col >= 0 and pd.notna(row[sgpa_col]) else ""
            
            if percentage_col >= 0 and pd.notna(row[percentage_col]):
                try:
                    new_row["PERCENTAGE"] = f"{float(row[percentage_col]):.2f}"
                except:
                    new_row["PERCENTAGE"] = str(row[percentage_col])
            
            # ===== PROCESS EACH SUBJECT =====
            for i, subj in enumerate(subjects):
                num = str(i + 1).zfill(2)
                c = subj['col_idx']
                code_upper = subj['code'].upper()
                
                new_row[f"SUBJ_NAME__{num}"] = subj['name']
                new_row[f"SUBJ_CODE__{num}"] = subj['code']
                new_row[f"MAX_MARKS_TH__{num}"] = "100"
                new_row[f"MIN_MARKS_TH__{num}"] = ""
                new_row[f"OBT_MARKS_TH__{num}"] = ""
                
                # ===== AUTO-POPULATE MAX_CREDS_TH from credit_map =====
                struct_from_map = credit_map.get(code_upper)
                if struct_from_map:
                    new_row[f"MAX_CREDS_TH__{num}"] = struct_from_map
                else:
                    new_row[f"MAX_CREDS_TH__{num}"] = subj['fallback_cred']
                
                # Total Marks
                new_row[f"OBT_CREDS_TH__{num}"] = str(row[c + 3]) if pd.notna(row[c + 3]) else ""
                
                # CiGi from column
                new_row[f"CRD_POINT_TH__{num}"] = str(row[c + 8]) if pd.notna(row[c + 8]) else ""
                
                # Grade Letter
                new_row[f"GRD_LETTR_TH__{num}"] = str(row[c + 6]) if pd.notna(row[c + 6]) else ""
                
                # Grade Point
                new_row[f"GRD_POINT_TH__{num}"] = str(row[c + 5]) if pd.notna(row[c + 5]) else ""
                
                new_row[f"RESULT_TH__{num}"] = ""
                new_row[f"REMARKS_TH__{num}"] = ""
            
            processed_data.append(new_row)
        
        if not processed_data:
            st.error("❌ No valid student data found.")
        else:
            final_df = pd.DataFrame(processed_data)
            
            st.success(f"✅ Processing Complete! **{subject_count} Subjects** | **{len(final_df)} Students**")
            
            st.info("""
            ✅ **Auto-Detected Features:**
            - GRAND_TTL_CREDITS: Automatically set from Program Structure sheet
            - MAX_CREDS_TH__01 to MAX_CREDS_TH__05: Subject-wise credit structure
            - EMAIL: Extracted from Column G of TR sheet (Priority over Student Details file)
            """)
            
            st.subheader("📊 Data Preview")
            col_preview = st.columns([3, 3, 1])
            with col_preview[0]:
                st.metric("Students Processed", len(final_df))
            with col_preview[1]:
                st.metric("Subjects Detected", subject_count)
            with col_preview[2]:
                st.metric("Total Credits", grand_total_structure)
            
            st.dataframe(final_df, use_container_width=True)
            
            st.subheader("📥 Download Final File")
            col1, col2 = st.columns([2, 1])
            
            with col1:
                output_name = st.text_input(
                    "Rename your file here:", 
                    value=f"Final_Truscholar_{subject_count}_Subjects_Mapped",
                    help="File will be saved as .xlsx"
                )
            
            if not output_name.endswith('.xlsx'):
                output_name += '.xlsx'
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name=f"{subject_count} Subjects")
            
            st.download_button(
                label=f"⬇️ Download Final {subject_count} Subjects Excel",
                data=output.getvalue(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_final"
            )
    
    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
        import traceback
        st.error(traceback.format_exc())
