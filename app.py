import streamlit as st
import pandas as pd
import io
import os
import zipfile
import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pypdf import PdfReader, PdfWriter, Transformation

# --- CONFIGURATION ---
LAYOUT = {
    "judge_y": 765,      # Top line
    "comp_y": 745,       # Middle line
    "contest_y": 730,    # Bottom line
    "margin_left": 50,
    "margin_right": 562, # Right align limit (612 width - 50 margin)
    "page_center": 306,  # Center of page
}

FORMAT_MAPPING = {
    "MUS": ["MUS_Long.pdf", "MUS_Short.pdf"],
    "PER": ["PER_Long.pdf", "PER_Short.pdf"],
    "SNG": ["SNG_Long.pdf", "SNG_Short.pdf"]
}

CAT_FULL_NAMES = {
    "MUS": "Musicality",
    "PER": "Performance",
    "SNG": "Singing"
}

TEMPLATE_DIR = "templates"

# --- HELPER FUNCTIONS ---

def clean_filename(text):
    """Sanitizes strings for use in filenames."""
    text = str(text).replace("/", "-").replace("\\", "-")
    return "".join(c for c in text if c.isalnum() or c in (' ', '-', '_')).strip()

def escape_rtf(text):
    """Escapes special characters for RTF output."""
    if pd.isna(text): return ""
    text = str(text)
    return text.replace('\\', '\\\\').replace('{', '\{').replace('}', '\}')

def apply_margin_to_page(page, margin_inch=0.25):
    """
    Scales the page content to ensure it fits within the specified margin.
    Used for the Overlay/Info layer only.
    """
    # Convert inches to points (1 inch = 72 pts)
    margin_pt = margin_inch * 72
    
    # Get current page dimensions
    width = float(page.mediabox.width)
    height = float(page.mediabox.height)
    
    # Calculate safe area dimensions
    safe_width = width - (2 * margin_pt)
    safe_height = height - (2 * margin_pt)
    
    # Calculate scale factors to fit safe area
    scale_x = safe_width / width
    scale_y = safe_height / height
    
    # Use the smaller scale to maintain aspect ratio
    scale = min(scale_x, scale_y)
    
    # Calculate offset to center the scaled content
    tx = (width - (width * scale)) / 2
    ty = (height - (height * scale)) / 2
    
    # Apply transformation (Scale & Translate)
    op = Transformation().scale(scale, scale).translate(tx, ty)
    page.add_transformation(op)
    
    return page

def generate_rtf_content(judges_df, competitors_df, context):
    """Generates raw RTF string content."""
    rtf = [
        r"{\rtf1\ansi\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Arial;}}",
        r"{\colortbl ;\red0\green0\blue0;}",
        r"\viewkind4\uc1\pard\sa200\sl276\slmult1\f0\fs24\lang9 "
    ]
    
    comp_list = competitors_df.to_dict('records')
    # Filter valid judges first to calculate total iterations correctly
    valid_judges = judges_df[judges_df['Number'] != 0].to_dict('records')
    
    total_items = len(comp_list) * len(valid_judges)
    current_count = 0
    
    for judge in valid_judges:
        for comp in comp_list:
            current_count += 1
            
            try: j_num = int(float(judge['Number']))
            except: j_num = judge['Number']
            
            try: c_num = int(float(comp['Number']))
            except: c_num = comp['Number']
                
            # Construct Lines
            c_name_text = f"{c_num}. {comp['Name']}"
            ctx_line = f"{context['district']} - {context['session']}, {context['date']}"
            
            # Prepare Competitor Line (Handle Director for Chorus)
            c_line_rtf = escape_rtf(c_name_text)
            
            if "Chorus" in context['session']:
                director_val = comp.get('Director', '')
                if pd.notna(director_val) and str(director_val).strip() != "":
                    # Add Director on a new line
                    c_line_rtf += r"\line " + escape_rtf(director_val)

            # 1. Judge Info: Right aligned (\qr)
            # Format: "Name - Number"
            # Name: Bold 16pt (\b\fs32)
            # Number: Bold 36pt (\fs72)
            rtf.append(r"\pard\qr\b\fs32 " + escape_rtf(judge['Name']) + r" - \fs72 " + str(j_num) + r"\b0\fs24\par")
            
            # 2. Competitor Info: Left aligned (\ql), Normal size (12pt default)
            rtf.append(r"\pard\ql " + c_line_rtf + r"\par")
            
            # 3. Contest Info: Centered (\qc), Normal size
            rtf.append(r"\pard\qc " + escape_rtf(ctx_line) + r"\par")
            
            # 4. Spacing (optional, but good for resetting)
            rtf.append(r"\pard\par") 
            
            # Page break after EVERY judge/competitor pair (except the very last one)
            if current_count < total_items:
                rtf.append(r"\page ")
            
    rtf.append("}")
    return "".join(rtf)

def generate_folder_labels_rtf(judges_df, context):
    """Generates Avery 8163 Labels (2x4 inches) in editable RTF format."""
    rtf = [
        r"{\rtf1\ansi\deff0\nouicompat\viewkind4\uc1",
        r"{\fonttbl{\f0\fnil\fcharset0 Arial;}}",
        r"{\colortbl ;\red0\green0\blue0;}",
        r"\paperw12240\paperh15840\margl225\margr225\margt720\margb720",
        r"\pard\plain\fs20 "
    ]
    
    active_judges = judges_df[judges_df['Print'] == True].to_dict('records')
    judges = [j for j in active_judges if j['Number'] != 0]
    total_judges = len(judges)
    
    for i in range(0, total_judges, 2):
        j1 = judges[i]
        j2 = judges[i+1] if (i+1) < total_judges else None
        
        rtf.append(r"\trowd\trgaph108\trleft0\trrh2880")
        rtf.append(r"\clvertalc\brdrt\brdrnil\brdrl\brdrnil\brdrb\brdrnil\brdrr\brdrnil\cellx5760")
        rtf.append(r"\clvertalc\brdrt\brdrnil\brdrl\brdrnil\brdrb\brdrnil\brdrr\brdrnil\cellx6030")
        rtf.append(r"\clvertalc\brdrt\brdrnil\brdrl\brdrnil\brdrb\brdrnil\brdrr\brdrnil\cellx11790")
        
        # Cell 1
        rtf.append(r"\pard\intbl\qc\sa0\sb0") 
        c_short = j1['Category']
        c_full = CAT_FULL_NAMES.get(c_short, c_short)
        
        # 1. Judge Name (Bold, 14pt)
        rtf.append(r"\b\f0\fs28 " + escape_rtf(j1['Name']) + r"\b0\par") 
        # 2. Category (11pt)
        rtf.append(r"\fs22 " + escape_rtf(f"{c_full} Category") + r"\par") 
        # 3. Session (10pt)
        rtf.append(r"\fs20 " + escape_rtf(context['session']) + r"\par") 
        # 4. District (10pt)
        rtf.append(escape_rtf(context['district']) + r"\par")
        # 5. Date (10pt)
        rtf.append(escape_rtf(context['date']))
        
        rtf.append(r"\cell")
        
        # Cell 2 (Gutter)
        rtf.append(r"\pard\intbl\cell")
        
        # Cell 3
        if j2:
            rtf.append(r"\pard\intbl\qc\sa0\sb0")
            c_short2 = j2['Category']
            c_full2 = CAT_FULL_NAMES.get(c_short2, c_short2)
            
            # 1. Judge Name
            rtf.append(r"\b\f0\fs28 " + escape_rtf(j2['Name']) + r"\b0\par")
            # 2. Category
            rtf.append(r"\fs22 " + escape_rtf(f"{c_full2} Category") + r"\par")
            # 3. Session
            rtf.append(r"\fs20 " + escape_rtf(context['session']) + r"\par")
            # 4. District
            rtf.append(escape_rtf(context['district']) + r"\par")
            # 5. Date
            rtf.append(escape_rtf(context['date']))
        else:
            rtf.append(r"\pard\intbl") 
            
        rtf.append(r"\cell")
        rtf.append(r"\row")
        
    rtf.append("}")
    return "".join(rtf)

def balance_and_sort_judges(df):
    """
    Ensures balanced panels (adding Absent judges if needed) but relies on 
    calculate_numbers for the final sorting and numbering.
    """
    df = df.copy()
    categories = ['MUS', 'PER', 'SNG']
    
    max_count = 0
    for cat in categories:
        count = len(df[(df['Category'] == cat) & (df['Type'] == 'Official')])
        if count > max_count:
            max_count = count
            
    if max_count > 0:
        new_rows = []
        for cat in categories:
            current_count = len(df[(df['Category'] == cat) & (df['Type'] == 'Official')])
            diff = max_count - current_count
            
            if diff > 0:
                for _ in range(diff):
                    new_rows.append({
                        "Name": f"Absent {cat} Judge",
                        "Category": cat,
                        "Type": "Official",
                        "Print": False,
                        "Number": 0
                    })
        if new_rows:
            df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
            
    return df

def calculate_numbers(df):
    """
    Sorts judges by Category -> Type -> Last Name and assigns numbers.
    """
    df = df.copy()
    
    # 1. Clean Data
    df['Category'] = df['Category'].astype(str).str.upper().str.strip()
    df['Type'] = df['Type'].astype(str).str.title().str.strip()
    
    # 2. Setup Sorting Helpers
    # Extract Last Name (using the last word in the string)
    df['Sort_Last_Name'] = df['Name'].apply(lambda x: str(x).strip().split()[-1] if len(str(x).strip()) > 0 else "")
    
    # Map Categories to ensure MUS -> PER -> SNG order
    cat_sorter = {'MUS': 0, 'PER': 1, 'SNG': 2}
    df['Sort_Cat'] = df['Category'].map(cat_sorter).fillna(99)
    
    # Map Type to ensure Official -> Practice order
    type_sorter = {'Official': 0, 'Practice': 1}
    df['Sort_Type'] = df['Type'].map(type_sorter).fillna(99)
    
    # 3. Sort the DataFrame
    df = df.sort_values(by=['Sort_Cat', 'Sort_Type', 'Sort_Last_Name'])
    
    # 4. Numbering Logic
    if 'Number' not in df.columns:
        df['Number'] = 0
    
    cat_order = ['MUS', 'PER', 'SNG']
    current_official_num = 1
    
    for cat in cat_order:
        # Number Official Judges
        mask_official = (df['Category'] == cat) & (df['Type'] == 'Official')
        count_official = mask_official.sum()
        if count_official > 0:
            df.loc[mask_official, 'Number'] = range(current_official_num, current_official_num + count_official)
            current_official_num += count_official
            
        # Number Practice Judges
        mask_practice = (df['Category'] == cat) & (df['Type'] == 'Practice')
        count_practice = mask_practice.sum()
        if count_practice > 0:
            df.loc[mask_practice, 'Number'] = range(50, 50 + count_practice)
            
    # 5. Cleanup helper columns
    df = df.drop(columns=['Sort_Last_Name', 'Sort_Cat', 'Sort_Type'])
    
    return df

def create_overlay(data, is_short=False):
    """Creates the text overlay PDF with the requested layout."""
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    
    # 1. JUDGE INFO (Right Aligned)
    try: j_num = int(float(data['judge_num']))
    except: j_num = data['judge_num']
    
    if is_short:
        # SHORT FORMAT: Big Number (36pt), Normal Name (16pt)
        can.setFont("Helvetica-Bold", 16)
        name_text = str(data['judge_name'])
        name_width = can.stringWidth(name_text, "Helvetica-Bold", 16)
        can.drawRightString(LAYOUT["margin_right"], LAYOUT["judge_y"], name_text)
        
        can.setFont("Helvetica-Bold", 36)
        num_text = str(j_num)
        can.drawRightString(LAYOUT["margin_right"] - name_width - 15, LAYOUT["judge_y"], num_text)
    else:
        # LONG FORMAT: Normal (16pt) "1. Judge Name"
        can.setFont("Helvetica-Bold", 16)
        judge_text = f"{j_num}. {data['judge_name']}"
        can.drawRightString(LAYOUT["margin_right"], LAYOUT["judge_y"], judge_text)
    
    # 2. COMPETITOR INFO (Left Aligned)
    can.setFont("Helvetica", 12)
    try: c_num = int(float(data['comp_num']))
    except: c_num = data['comp_num']
    comp_text = f"{c_num}. {data['comp_name']}"
    can.drawString(LAYOUT["margin_left"], LAYOUT["comp_y"], comp_text)
    
    # --- ADD DIRECTOR FOR CHORUS (LONG TEMPLATES) ---
    # Only if NOT is_short (i.e. Long) AND "Chorus" is in the session name
    is_chorus = "Chorus" in data.get('session', '')
    if is_chorus and not is_short:
        director = data.get('director', '')
        if director:
            can.drawString(LAYOUT["margin_left"], LAYOUT["comp_y"] - 14, director) # 14pt below competitor name

    # 3. CONTEST INFO (Center Aligned)
    can.setFont("Helvetica", 10)
    contest_text = f"{data['district']} - {data['session']}, {data['date']}"
    can.drawCentredString(LAYOUT["page_center"], LAYOUT["contest_y"], contest_text)
    
    can.save()
    packet.seek(0)
    return packet

def get_page_data(judge, comp, contest_info):
    return {
        "district": contest_info['district'],
        "session": contest_info['session'],
        "date": contest_info['date'],
        "comp_name": comp['Name'],
        "comp_num": comp['Number'],
        "director": comp.get('Director', ''),
        "judge_name": judge['Name'],
        "judge_num": judge['Number'], 
    }

def generate_blank_forms(requests):
    output_writer = PdfWriter()
    pages_added = 0
    for fmt_name, count in requests.items():
        if count > 0:
            template_path = os.path.join(TEMPLATE_DIR, f"{fmt_name}.pdf")
            if not os.path.exists(template_path): continue
            reader = PdfReader(template_path)
            for _ in range(count):
                for page in reader.pages:
                    output_writer.add_page(page)
                    pages_added += 1
    if pages_added > 0:
        pdf_bytes = io.BytesIO()
        output_writer.write(pdf_bytes)
        pdf_bytes.seek(0)
        return pdf_bytes, pages_added
    else:
        return None, 0

# --- APP INIT ---
st.set_page_config(page_title="Contest Form Generator", page_icon="📝", layout="wide")
st.title("📝 Contest Form Generator")

# Initialize Session State Tables if Empty
if 'judges_data' not in st.session_state:
    st.session_state['judges_data'] = pd.DataFrame(columns=["Number", "Name", "Category", "Type", "Print"])
if 'competitors_data' not in st.session_state:
    st.session_state['competitors_data'] = pd.DataFrame(columns=["Number", "Name", "Director", "Print"])

# 1. CONTEST INPUTS
with st.container():
    st.subheader("Contest Information")
    c1, c2, c3 = st.columns(3)
    c_district = c1.text_input("District", value="")
    c_date_obj = c2.date_input("Contest Date")
    c_session = c3.selectbox(
        "Session", 
        options=["Quartet Quarter-Finals", "Quartet Semi-Finals", "Chorus Finals", "Quartet Finals"],
        index=1 
    )

st.divider()

col_left, col_right = st.columns([1, 1])

# --- COL 1: JUDGES ---
with col_left:
    st.subheader("Judges")
    
    with st.expander("📂 Import Assignments Report", expanded=True):
        judges_file = st.file_uploader("Upload Assignment Report", type=['csv'], key="j_up", label_visibility="collapsed")
        if judges_file:
            try:
                raw_df = pd.read_csv(judges_file)
                raw_df.columns = raw_df.columns.str.strip()
                if {'Name', 'Category', 'Type'}.issubset(raw_df.columns):
                    raw_df['Category'] = raw_df['Category'].astype(str).str.upper().str.strip()
                    clean_df = raw_df[raw_df['Category'].isin(['MUS', 'PER', 'SNG'])].copy()
                    clean_df = clean_df[['Name', 'Category', 'Type']]
                    clean_df['Print'] = True
                    clean_df = balance_and_sort_judges(clean_df)
                    clean_df = calculate_numbers(clean_df)
                    clean_df = clean_df[['Number', 'Name', 'Category', 'Type', 'Print']]
                    
                    if not clean_df.equals(st.session_state['judges_data']):
                        st.session_state['judges_data'] = clean_df
                        st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

    # Manual Tools
    st.write("---")
    st.markdown("**Manage Judges List**")
    
    j_col1, j_col2 = st.columns(2)
    with j_col1:
        if st.button("Clear List", key="j_clear"):
            st.session_state['judges_data'] = pd.DataFrame(columns=["Number", "Name", "Category", "Type", "Print"])
            if "judge_editor" in st.session_state: del st.session_state["judge_editor"]
            st.rerun()
    with j_col2:
        if st.button("Auto-number Judges", help="Sorts and re-numbers the list based on Category and Name."):
            # Apply sorting and numbering logic to current data
            df = st.session_state['judges_data']
            if not df.empty:
                df = calculate_numbers(df)
                st.session_state['judges_data'] = df
                if "judge_editor" in st.session_state: del st.session_state["judge_editor"]
                st.rerun()

    edited_judges = st.data_editor(
        st.session_state['judges_data'].reset_index(drop=True),
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "Number": st.column_config.NumberColumn("Judge #", disabled=False, width="small", help="Edit manually or click Auto-number"),
            "Print": st.column_config.CheckboxColumn("Print?", default=True, width="small"),
            "Category": st.column_config.SelectboxColumn("Category", options=["MUS", "PER", "SNG"], required=True),
            "Type": st.column_config.SelectboxColumn("Type", options=["Official", "Practice"], required=True),
            "Name": st.column_config.TextColumn("Name", required=True)
        },
        column_order=["Number", "Name", "Category", "Type", "Print"],
        width='stretch',
        key="judge_editor"
    )
    
    # Save edits AND check for new rows to fill index
    if not edited_judges.equals(st.session_state['judges_data']):
         # If new rows are added, they usually have missing/NaN numbers.
         # We fill them with max_num + 1
         if edited_judges['Number'].isnull().any():
             max_num = edited_judges['Number'].max()
             if pd.isna(max_num): max_num = 0
             
             nan_rows = edited_judges[edited_judges['Number'].isnull()].index
             for i, idx in enumerate(nan_rows):
                 edited_judges.at[idx, 'Number'] = max_num + 1 + i
                 
         st.session_state['judges_data'] = edited_judges
         st.rerun()


# --- COL 2: COMPETITORS ---
with col_right:
    st.subheader("Competitors")
    
    with st.expander("📂 Import DRCJ Report", expanded=True):
        competitors_file = st.file_uploader("Upload CSV (OA, Group Name...)", type=['csv'], key="j_comp", label_visibility="collapsed")
        if competitors_file:
            try:
                raw_c = pd.read_csv(competitors_file)
                raw_c.columns = raw_c.columns.str.strip()
                if {'OA', 'Group Name'}.issubset(raw_c.columns):
                    new_comp = pd.DataFrame()
                    new_comp['Number'] = raw_c['OA']
                    new_comp['Name'] = raw_c['Group Name']
                    is_chorus_session = "Chorus" in c_session
                    if is_chorus_session and 'Director/Participant(s)' in raw_c.columns:
                         new_comp['Director'] = raw_c['Director/Participant(s)']
                    else:
                         new_comp['Director'] = ""
                    new_comp['Print'] = True
                    
                    if not new_comp.equals(st.session_state['competitors_data']):
                        st.session_state['competitors_data'] = new_comp
                        st.rerun()
            except Exception as e:
                st.error(f"Error reading CSV: {e}")

    # Manual Tools
    st.write("---")
    st.markdown("**Manage Competitors List**")
    
    if st.button("Clear List", key="c_clear"):
        st.session_state['competitors_data'] = pd.DataFrame(columns=["Number", "Name", "Director", "Print"])
        if "comp_editor" in st.session_state: del st.session_state["comp_editor"]
        st.rerun()
    
    edited_competitors = st.data_editor(
        st.session_state['competitors_data'].reset_index(drop=True),
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "Number": st.column_config.NumberColumn("OA", width="small", help="Order of Appearance"),
            "Print": st.column_config.CheckboxColumn("Print?", default=True, width="small"),
            "Name": st.column_config.TextColumn("Group Name", required=True),
            "Director": st.column_config.TextColumn("Director (Chorus Only)")
        },
        column_order=["Number", "Name", "Director", "Print"],
        width='stretch',
        key="comp_editor"
    )
    
    # Save edits AND check for new rows to fill index
    if not edited_competitors.equals(st.session_state['competitors_data']):
        if edited_competitors['Number'].isnull().any():
             max_num = edited_competitors['Number'].max()
             if pd.isna(max_num): max_num = 0
             
             nan_rows = edited_competitors[edited_competitors['Number'].isnull()].index
             for i, idx in enumerate(nan_rows):
                 edited_competitors.at[idx, 'Number'] = max_num + 1 + i
        
        st.session_state['competitors_data'] = edited_competitors
        st.rerun()


# --- GENERATE FORMS SECTION ---
st.divider()
st.subheader("Generate Forms")

has_judges = not st.session_state['judges_data'].empty
final_judges = st.session_state['judges_data']
final_competitors = st.session_state['competitors_data']

# Format Date Object to String for Forms
formatted_date = c_date_obj.strftime("%m/%d/%Y")

contest_context = {
    "district": c_district,
    "session": c_session,
    "date": formatted_date
}
safe_session = clean_filename(c_session)
safe_date = clean_filename(formatted_date)

col_gen1, col_gen2, col_gen3 = st.columns(3)

# --- OPTION 1: BY JUDGE ---
with col_gen1:
    st.markdown("### Option 1: By Judge")
    st.write("Creates a PDF packet for each judge containing forms for all competitors.")
    
    if st.button("Generate PDFs for each Judge", type="primary"):
        if not c_district:
            st.error("Please fill in District.")
        elif has_judges and not final_competitors.empty:
            with st.spinner("⏳ Generating Judge Packets... Please wait."):
                try:
                    active_judges = final_judges[final_judges['Print'] == True]
                    active_competitors = final_competitors[final_competitors['Print'] == True]
                    competitor_list = active_competitors.to_dict('records')
                    
                    if active_judges.empty or active_competitors.empty:
                        st.warning("Please select at least one Judge and one Competitor.")
                    else:
                        generated_pdfs = [] 
                        prog_bar = st.progress(0, text="Processing judges...")
                        total_j = len(active_judges)

                        for idx, (_, judge) in enumerate(active_judges.iterrows()):
                            if judge['Number'] == 0: continue
                            prog_bar.progress((idx + 1) / total_j, text=f"Processing Judge: {judge['Name']}")
                            
                            cat = judge['Category']
                            templates = FORMAT_MAPPING.get(cat, [])
                            if not templates: continue
                            
                            writer = PdfWriter()
                            pages_added = 0
                            
                            for t_name in templates:
                                if "Long" not in t_name: continue
                                t_path = os.path.join(TEMPLATE_DIR, t_name)
                                if not os.path.exists(t_path): continue
                                
                                is_short = "Short" in t_name
                                base = PdfReader(t_path)
                                
                                if is_short:
                                    base_page = base.pages[0]
                                    temp_writer = PdfWriter()
                                    temp_writer.add_page(base_page)
                                    target_page = temp_writer.pages[0]
                                    
                                    for i in range(0, len(competitor_list), 2):
                                        comp1 = competitor_list[i]
                                        comp2 = competitor_list[i+1] if (i+1) < len(competitor_list) else None
                                        
                                        base = PdfReader(t_path).pages[0]
                                        temp_writer = PdfWriter()
                                        temp_writer.add_page(base)
                                        target_page = temp_writer.pages[0]
                                        
                                        data1 = get_page_data(judge, comp1, contest_context)
                                        overlay1 = PdfReader(create_overlay(data1, is_short=True)).pages[0]
                                        apply_margin_to_page(overlay1) # MARGIN INFO ONLY
                                        target_page.merge_page(overlay1)
                                        
                                        if comp2:
                                            data2 = get_page_data(judge, comp2, contest_context)
                                            overlay2 = PdfReader(create_overlay(data2, is_short=True)).pages[0]
                                            apply_margin_to_page(overlay2) # MARGIN INFO ONLY
                                            overlay2.add_transformation(Transformation().rotate(180).translate(tx=612, ty=792))
                                            target_page.merge_page(overlay2)
                                        
                                        writer.add_page(target_page)
                                        pages_added += 1
                                else:
                                    for comp in competitor_list:
                                        page_data = get_page_data(judge, comp, contest_context)
                                        overlay = PdfReader(create_overlay(page_data, is_short=False))
                                        template_reader = PdfReader(t_path)
                                        
                                        for i_page, page in enumerate(template_reader.pages):
                                            temp_writer = PdfWriter()
                                            temp_writer.add_page(page)
                                            target_page = temp_writer.pages[0]
                                            
                                            if i_page == 0:
                                                info_page = overlay.pages[0]
                                                apply_margin_to_page(info_page) # MARGIN INFO ONLY
                                                target_page.merge_page(info_page)
                                            
                                            writer.add_page(target_page)
                                            pages_added += 1
                            
                            if pages_added > 0:
                                safe_judge = clean_filename(judge['Name'])
                                fname = f"{safe_session}_{safe_judge}_{safe_date}.pdf"
                                pdf_bytes = io.BytesIO()
                                writer.write(pdf_bytes)
                                generated_pdfs.append((fname, pdf_bytes))
                        
                        prog_bar.empty()
                        if len(generated_pdfs) == 1:
                            fname, data = generated_pdfs[0]
                            st.success(f"Created single PDF packet: {fname}")
                            st.download_button("📥 Download PDF Packet", data, fname, "application/pdf")
                        elif len(generated_pdfs) > 1:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w") as zf:
                                for fname, data in generated_pdfs:
                                    zf.writestr(fname, data.getvalue())
                            zip_buffer.seek(0)
                            st.success(f"Created {len(generated_pdfs)} Judge Packets.")
                            st.download_button("📥 Download Judge Packets", zip_buffer, f"{safe_session}_Judge_Packets.zip", "application/zip")
                        else:
                            st.warning("No files generated.")

                except Exception as e:
                    st.error(f"Error: {e}")
        else:
            st.error("Missing Data")

# --- OPTION 2: BY CATEGORY ---
with col_gen2:
    st.markdown("### Option 2: By Category")
    st.write("Creates bulk PDFs for each Category/Format.")
    
    if st.button("Generate PDFs for each Category", type="primary"):
        if not c_district:
            st.error("Please fill in District.")
        elif has_judges and not final_competitors.empty:
            with st.spinner("⏳ Generating Category Files... Please wait."):
                try:
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        active_judges = final_judges[final_judges['Print'] == True]
                        active_competitors = final_competitors[final_competitors['Print'] == True]
                        competitor_list = active_competitors.to_dict('records')
                        
                        cats = list(FORMAT_MAPPING.items())
                        prog_bar = st.progress(0, text="Processing categories...")
                        total_cats = len(cats)

                        for idx, (cat, formats) in enumerate(cats):
                            prog_bar.progress((idx + 1) / total_cats, text=f"Processing Category: {cat}")
                            cat_judges = active_judges[active_judges['Category'] == cat]
                            if cat_judges.empty: continue
                            
                            for t_name in formats:
                                t_path = os.path.join(TEMPLATE_DIR, t_name)
                                if not os.path.exists(t_path): continue
                                
                                is_short = "Short" in t_name
                                writer = PdfWriter()
                                pages_added = 0
                                
                                for _, judge in cat_judges.iterrows():
                                    if judge['Number'] == 0: continue
                                    
                                    if is_short:
                                        for i in range(0, len(competitor_list), 2):
                                            comp1 = competitor_list[i]
                                            comp2 = competitor_list[i+1] if (i+1) < len(competitor_list) else None
                                            base = PdfReader(t_path).pages[0]
                                            temp_writer = PdfWriter()
                                            temp_writer.add_page(base)
                                            target_page = temp_writer.pages[0]
                                            data1 = get_page_data(judge, comp1, contest_context)
                                            overlay1 = PdfReader(create_overlay(data1, is_short=True)).pages[0]
                                            apply_margin_to_page(overlay1) # MARGIN INFO ONLY
                                            target_page.merge_page(overlay1)
                                            if comp2:
                                                data2 = get_page_data(judge, comp2, contest_context)
                                                overlay2 = PdfReader(create_overlay(data2, is_short=True)).pages[0]
                                                apply_margin_to_page(overlay2) # MARGIN INFO ONLY
                                                overlay2.add_transformation(Transformation().rotate(180).translate(tx=612, ty=792))
                                                target_page.merge_page(overlay2)
                                            writer.add_page(target_page)
                                            pages_added += 1
                                    else:
                                        for comp in competitor_list:
                                            page_data = get_page_data(judge, comp, contest_context)
                                            overlay = PdfReader(create_overlay(page_data, is_short=False))
                                            template_reader = PdfReader(t_path)
                                            for i_page, page in enumerate(template_reader.pages):
                                                temp_writer = PdfWriter()
                                                temp_writer.add_page(page)
                                                target_page = temp_writer.pages[0]
                                                if i_page == 0:
                                                    info_page = overlay.pages[0]
                                                    apply_margin_to_page(info_page) # MARGIN INFO ONLY
                                                    target_page.merge_page(info_page)
                                                writer.add_page(target_page)
                                                pages_added += 1
                                if pages_added > 0:
                                    format_suffix = t_name.replace(".pdf", "") 
                                    fname = f"{safe_session}_{format_suffix}_{safe_date}.pdf"
                                    pdf_bytes = io.BytesIO()
                                    writer.write(pdf_bytes)
                                    zf.writestr(fname, pdf_bytes.getvalue())
                                    count += 1
                        
                        prog_bar.empty()
                    
                    zip_buffer.seek(0)
                    if count > 0:
                        st.success(f"Created {count} Category Files.")
                        st.download_button("📥 Download Category Files", zip_buffer, f"{safe_session}_Category_Files.zip", "application/zip")
                    else:
                        st.warning("No files generated.")
                except Exception as e:
                    st.error(f"Error: {e}")
        else:
            st.error("Missing Data")

# --- OPTION 3: CREATE LABELS ONLY (RTF) ---
with col_gen3:
    st.markdown("### Option 3: Create Labels Only")
    st.write("Generates an editable RTF file for each category.")
    
    if st.button("Create Labels Only (RTF)", type="primary"):
        if not c_district:
            st.error("Please fill in District.")
        elif has_judges and not final_competitors.empty:
            with st.spinner("⏳ Generating Labels..."):
                try:
                    zip_buffer = io.BytesIO()
                    count = 0
                    active_judges = final_judges[final_judges['Print'] == True]
                    active_competitors = final_competitors[final_competitors['Print'] == True]
                    
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for cat in ['MUS', 'PER', 'SNG']:
                            cat_judges = active_judges[active_judges['Category'] == cat]
                            if cat_judges.empty: continue
                            
                            rtf_content = generate_rtf_content(cat_judges, active_competitors, contest_context)
                            fname = f"{safe_session}_{cat}_Labels.rtf"
                            zf.writestr(fname, rtf_content)
                            count += 1
                    
                    zip_buffer.seek(0)
                    if count > 0:
                        st.success(f"Created {count} Label Files.")
                        st.download_button("📥 Download Labels (ZIP)", zip_buffer, f"{safe_session}_Labels.zip", "application/zip")
                    else:
                        st.warning("No labels generated.")
                except Exception as e:
                    st.error(f"Error: {e}")
        else:
            st.error("Missing Data")

# --- GENERATE FOLDER LABELS (NEW SECTION) ---
st.divider()
with st.expander("Create Folder Labels", expanded=False):
    st.write("Generates an editable RTF with Avery 8163 (2\" x 4\") labels for all selected judges.")
    if st.button("Generate Folder Labels (RTF)", type="primary"):
        if not c_district:
            st.error("Please fill in District.")
        elif has_judges:
            with st.spinner("Generating..."):
                try:
                    rtf_data = generate_folder_labels_rtf(final_judges, contest_context)
                    if rtf_data:
                        st.success("Labels generated successfully.")
                        st.download_button("📥 Download Folder Labels (RTF)", rtf_data, f"{safe_session}_Folder_Labels.rtf", "application/rtf")
                    else:
                        st.warning("No judges selected.")
                except Exception as e:
                    st.error(f"Error generating labels: {e}")
        else:
            st.error("No judges loaded.")

# --- GENERATE BLANK FORMS ---
st.divider()
with st.expander("Print Blank Forms (No Data)", expanded=False):
    st.write("Enter quantity for blank copies.")
    b_col1, b_col2, b_col3 = st.columns(3)
    blank_reqs = {}
    
    b_col1.markdown("**Music**")
    blank_reqs["MUS_Long"] = b_col1.number_input("MUS Long", min_value=0, step=1)
    blank_reqs["MUS_Short"] = b_col1.number_input("MUS Short", min_value=0, step=1)
    
    b_col2.markdown("**Performance**")
    blank_reqs["PER_Long"] = b_col2.number_input("PER Long", min_value=0, step=1)
    blank_reqs["PER_Short"] = b_col2.number_input("PER Short", min_value=0, step=1)
    
    b_col3.markdown("**Singing**")
    blank_reqs["SNG_Long"] = b_col3.number_input("SNG Long", min_value=0, step=1)
    blank_reqs["SNG_Short"] = b_col3.number_input("SNG Short", min_value=0, step=1)
    
    if st.button("Generate Blank Forms"):
        if any(v > 0 for v in blank_reqs.values()):
            with st.spinner("Generating..."):
                try:
                    b_pdf, count = generate_blank_forms(blank_reqs)
                    if count > 0:
                        st.success(f"Generated single PDF with {count} pages.")
                        st.download_button("📥 Download Blank Forms (PDF)", b_pdf, f"{safe_session}_Blank_Forms.pdf", "application/pdf")
                    else:
                        st.warning("Could not generate files.")
                except Exception as e:
                    st.error(f"Error: {e}")
        else:
            st.warning("Enter quantity > 0")