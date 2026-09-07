import streamlit as st
import pandas as pd
import io
import re
import openpyxl
from datetime import datetime

st.set_page_config(page_title="TR Sheet Processor", page_icon="logo.png", layout="wide")

logo_col, title_col = st.columns([1, 10])
with logo_col:
    st.image("logo.png", width=80)
with title_col:
    st.title("TR Sheet Mapper & Downloader (Pro Version V2)")

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
st.write("⚠️ Yeh file optional hai.")

student_map = {}

demo_student_df = pd.DataFrame({
    "Enrollment No": ["240810500024", "240810500025"], 
    "Student Name": ["Aasmant Umesh Mohadikar", "Rahul Kumar"], 
    "Email ID": ["student@example.com", "rahul@example.com"], 
    "Hall Admit No": ["HALL123", "HALL124"]
})
st.download_button("⬇️ Download Demo (Student Details)", data=to_excel(demo_student_df), file_name="Demo_Student_Details.xlsx", key="demo_stu")

student_file = st.file_uploader("Upload Student Details File:", type=["xlsx", "xls", "csv"], key="stu_upload")

def find_exact_col(cols, *names):
    """Return the first column name from `names` that exists exactly in `cols`."""
    for n in names:
        if n in cols:
            return n
    return None

if student_file:
    try:
        # Support both Excel and CSV reference files (e.g. Truscholar student export)
        if student_file.name.lower().endswith('.csv'):
            sdf = pd.read_csv(student_file)
        else:
            sdf = pd.read_excel(student_file)
        
        sdf.columns = [str(c).strip().upper() for c in sdf.columns]
        
        # ===== PRIORITY 1: Exact known column names (Truscholar export format) =====
        reg_col_name = find_exact_col(sdf.columns, 'REGISTRATION_NO', 'ENROLLMENT_NO')
        alt_reg_col_name = find_exact_col(sdf.columns, 'ENROLLMENT_NO', 'REGISTRATION_NO')
        email_col_name = find_exact_col(sdf.columns, 'CONTACT_EMAIL', 'EMAIL_ID', 'EMAIL')
        hall_col_name = find_exact_col(sdf.columns, 'HALL_ADMIT_NO__1', 'HALL_ADMIT_NO', 'HALL ADMIT NO')
        first_name_col = find_exact_col(sdf.columns, 'FIRST_NAME')
        last_name_col = find_exact_col(sdf.columns, 'LAST_NAME')
        name_col_name = find_exact_col(sdf.columns, 'STUDENT NAME', 'NAME')
        
        # ===== PRIORITY 2 (fallback): generic keyword search — for older/simple templates =====
        if not reg_col_name:
            enr_col = [c for c in sdf.columns if 'ENROLL' in c or 'ROLL' in c or 'REG' in c]
            reg_col_name = enr_col[0] if enr_col else None
        if not email_col_name:
            email_col = [c for c in sdf.columns if 'EMAIL' in c]
            email_col_name = email_col[0] if email_col else None
        if not hall_col_name:
            hall_col = [c for c in sdf.columns if 'HALL' in c or 'ADMIT' in c]
            hall_col_name = hall_col[0] if hall_col else None
        if not name_col_name and not first_name_col and not last_name_col:
            name_col_generic = [c for c in sdf.columns if 'NAME' in c]
            name_col_name = name_col_generic[0] if name_col_generic else None
        
        if reg_col_name:
            for _, row in sdf.iterrows():
                reg = str(row[reg_col_name]).strip() if pd.notna(row[reg_col_name]) else ""
                alt_reg = str(row[alt_reg_col_name]).strip() if alt_reg_col_name and pd.notna(row[alt_reg_col_name]) else ""
                
                # Prefer FIRST_NAME + LAST_NAME combo if present, else fall back to a single NAME column
                if first_name_col or last_name_col:
                    fn = str(row[first_name_col]).strip() if first_name_col and pd.notna(row[first_name_col]) else ""
                    ln = str(row[last_name_col]).strip() if last_name_col and pd.notna(row[last_name_col]) else ""
                    stu_name = f"{fn} {ln}".strip()
                else:
                    stu_name = str(row[name_col_name]).strip() if name_col_name and pd.notna(row[name_col_name]) else ""
                
                email = str(row[email_col_name]).strip() if email_col_name and pd.notna(row[email_col_name]) else ""
                hall = str(row[hall_col_name]).strip() if hall_col_name and pd.notna(row[hall_col_name]) else ""
                
                info = {"email": email, "hall": hall, "name": stu_name}
                
                # Store under both Registration No and Enrollment No (when both exist and differ)
                # so the lookup matches whichever ID the TR sheet uses.
                if reg and reg.lower() != 'nan':
                    student_map[reg] = info
                if alt_reg and alt_reg.lower() != 'nan' and alt_reg != reg:
                    student_map[alt_reg] = info
            
            st.success(f"✅ Student Details mapped successfully! ({len(sdf)} records loaded — matched on {reg_col_name})")
        else:
            st.error("Enrollment No / Registration No column not found in Student Details file.")
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

def sum_credit_structures(struct_list):
    """
    Position-wise (L-T-P) sum of structures like '02-00-00', '00-03-00',
    '04-00-00', '00-00-08' etc.
    Returns a combined string like '06-03-08'.
    Used when the sheet has NO explicit TOTAL row and we need to
    auto-calculate the grand total credit structure from subject rows.
    """
    parsed = []
    max_len = 0
    for s in struct_list:
        if not s or str(s).strip().lower() == 'nan':
            continue
        parts = str(s).strip().split('-')
        try:
            nums = [int(float(p)) for p in parts if p.strip() != '']
        except:
            continue
        if nums:
            parsed.append(nums)
            max_len = max(max_len, len(nums))
    
    if not parsed:
        return ""
    
    totals = [0] * max_len
    for nums in parsed:
        for i, n in enumerate(nums):
            totals[i] += n
    
    # Format each part with 2-digit leading zero, like original style
    return "-".join(f"{t:02d}" for t in totals)

def extract_credit_structure_from_sheet(xls_file, sheet_name):
    """
    Automatically detect and extract credit structure from Program Structure sheet using openpyxl
    Returns: (grand_total_structure, credit_map, credit_value_map)
    """
    credit_map = {}  # Maps CODE -> Credit Structure (e.g., "0-3-0")
    credit_value_map = {}  # Maps CODE -> Credit Value (e.g., 3)
    grand_total_structure = ""
    
    try:
        wb = openpyxl.load_workbook(xls_file, data_only=True)
        if sheet_name not in wb.sheetnames:
            st.warning(f"⚠️ Sheet '{sheet_name}' not found")
            return "", {}, {}
        
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        
        if len(all_rows) < 2:
            st.warning(f"⚠️ Sheet '{sheet_name}' is empty")
            return "", {}, {}
        
        # Get headers from first row
        header_row = all_rows[0]
        headers = [str(x).strip().upper() if x else "" for x in header_row]
        
        # Find column indices
        code_col_idx = -1
        struct_col_idx = -1
        credit_col_idx = -1
        name_col_idx = -1
        
        for i, h in enumerate(headers):
            if 'SUB CODE' in h or 'CODE' in h:
                code_col_idx = i
            if 'STRUCTURE' in h:
                struct_col_idx = i
            if 'CREDIT' in h and 'STRUCTURE' not in h:
                credit_col_idx = i
            if 'NAME' in h or 'SUBJECT' in h:
                name_col_idx = i
        
        st.info(f"📍 Detected columns: Code={code_col_idx}, Structure={struct_col_idx}, Credit={credit_col_idx}, Name={name_col_idx}")
        
        if code_col_idx == -1 or struct_col_idx == -1:
            st.warning(f"❌ Could not find Code or Structure columns")
            return "", {}, {}
        
        # Extract credit structure for each subject (skip header row 0)
        for row_idx in range(1, len(all_rows)):
            row = all_rows[row_idx]
            
            code = str(row[code_col_idx]).strip().upper() if code_col_idx >= 0 and pd.notna(row[code_col_idx]) else ""
            struct = row[struct_col_idx] if struct_col_idx >= 0 else ""
            credit = row[credit_col_idx] if credit_col_idx >= 0 else ""
            name = row[name_col_idx] if name_col_idx >= 0 else ""
            
            # Handle datetime object (convert to string format)
            if isinstance(struct, datetime):
                # Convert datetime(2008, 10, 3) to "03-10-08"
                struct = f"{struct.day:02d}-{struct.month:02d}-{struct.year % 100:02d}"
            
            struct = str(struct).strip() if struct else ""
            
            # Check if this is TOTAL row (last row or if credit is present but code is not)
            is_total_row = (not code or code == 'NONE' or code == '') and (credit or struct)
            
            if is_total_row:
                # This is the TOTAL row
                grand_total_structure = struct
                st.success(f"✅ Found Grand Total Credit Structure: **{grand_total_structure}**")
            elif code and code not in ['TOTAL', 'NONE']:
                # Subject row
                credit_map[code] = struct
                if isinstance(credit, (int, float)):
                    credit_value_map[code] = int(credit)
                st.write(f"  📚 {code}: {struct} (Credit: {credit}) - {name}")
        
        # ===== AUTO-CALCULATE TOTAL if no explicit TOTAL row found =====
        if not grand_total_structure and credit_map:
            grand_total_structure = sum_credit_structures(list(credit_map.values()))
            if grand_total_structure:
                st.info(f"ℹ️ Koi explicit TOTAL row nahi mili — subjects ke credit structures ko jodkar auto-calculate kiya: **{grand_total_structure}**")
        
        return grand_total_structure, credit_map, credit_value_map
    
    except Exception as e:
        st.error(f"❌ Error extracting credit structure: {e}")
        import traceback
        st.error(traceback.format_exc())
        return "", {}, {}

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
        credit_value_map = {}
        grand_total_structure = ""
        
        if program_sheet_name:
            try:
                st.info(f"🔄 Reading credit structure from sheet: '{program_sheet_name}'...")
                grand_total_structure, credit_map, credit_value_map = extract_credit_structure_from_sheet(uploaded_file, program_sheet_name)
                if grand_total_structure:
                    st.success(f"✅ Auto-extracted Grand Total Credit Structure: **{grand_total_structure}**")
                    st.info(f"📊 Subject-wise Credit Structures:\n{chr(10).join([f'  • {k}: {v}' for k,v in credit_map.items()])}")
                else:
                    st.warning(f"⚠️ Could not extract grand total structure")
            except Exception as e:
                st.warning(f"Could not read Program Structure sheet: {e}")
                import traceback
                st.error(traceback.format_exc())
        
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
        
        # ===== FALLBACK: if Program Structure sheet gave no grand total,
        # try to build it from the per-subject structures detected on the TR sheet =====
        if not grand_total_structure:
            fallback_structs = [
                credit_map.get(subj['code'].upper(), subj['fallback_cred'])
                for subj in subjects
            ]
            grand_total_structure = sum_credit_structures(fallback_structs)
            if grand_total_structure:
                st.info(f"ℹ️ TR sheet ke subjects se GRAND_TTL_CREDITS auto-calculate kiya: **{grand_total_structure}**")
        
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
                
                # ===== AUTO-POPULATE MAX_CREDS_TH from credit_map (extracted from Program Structure) =====
                struct_from_map = credit_map.get(code_upper)
                if struct_from_map:
                    new_row[f"MAX_CREDS_TH__{num}"] = struct_from_map
                else:
                    new_row[f"MAX_CREDS_TH__{num}"] = subj['fallback_cred']
                
                # Total Marks from TR sheet
                new_row[f"OBT_CREDS_TH__{num}"] = str(row[c + 3]) if pd.notna(row[c + 3]) else ""
                
                # CiGi from TR sheet column
                new_row[f"CRD_POINT_TH__{num}"] = str(row[c + 8]) if pd.notna(row[c + 8]) else ""
                
                # Grade Letter from TR sheet
                new_row[f"GRD_LETTR_TH__{num}"] = str(row[c + 6]) if pd.notna(row[c + 6]) else ""
                
                # Grade Point from TR sheet
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
            - GRAND_TTL_CREDITS: Automatically set from Program Structure sheet (ya subject-wise structures se auto-calculate)
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
