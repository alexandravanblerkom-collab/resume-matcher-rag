import streamlit as st
import chromadb
import os
import json
import io
import hashlib
import pandas as pd # NEW: For the data editor
from dotenv import load_dotenv
from google import genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT

# ---------------------------------------------------------
# 1. SETUP & AUTHENTICATION
# ---------------------------------------------------------
load_dotenv()
api_key = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=api_key)

st.set_page_config(page_title="AI Resume Matcher", page_icon="📄", layout="wide")
st.title("📄 Zero-Hallucination Resume Builder")

if "logged_in_user" not in st.session_state:
    st.session_state.logged_in_user = None

# =========================================================
# 2. REAL REGISTRATION & LOGIN SYSTEM
# =========================================================
USER_FILE = "users.json"

# Helper function to read our user database
def load_users():
    if not os.path.exists(USER_FILE):
        return {}
    with open(USER_FILE, "r") as f:
        return json.load(f)

# Helper function to save new users
def save_users(users):
    with open(USER_FILE, "w") as f:
        json.dump(users, f)

# Helper function to securely scramble passwords
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# If not logged in, show the Login/Signup screen
if st.session_state.logged_in_user is None:
    st.subheader("Welcome to the Resume Vault")
    
    # Create two tabs for the auth screen
    login_tab, signup_tab = st.tabs(["Log In", "Create Account"])


    
    # --- LOG IN TAB ---
    with login_tab:
        with st.form("login_form"):
            log_user = st.text_input("Username").lower()
            log_pass = st.text_input("Password", type="password")
            submit_login = st.form_submit_button("Log In")
            
            if submit_login:
                users = load_users()
                # Check if user exists AND the scrambled password matches
                if log_user in users and users[log_user] == hash_password(log_pass):
                    st.session_state.logged_in_user = log_user
                    st.rerun()
                else:
                    st.error("Incorrect username or password.")
                    
    # --- SIGN UP TAB ---
    with signup_tab:
        with st.form("signup_form"):
            new_user = st.text_input("Choose a Username").lower()
            new_pass = st.text_input("Choose a Password", type="password")
            submit_signup = st.form_submit_button("Sign Up")
            
            if submit_signup:
                users = load_users()
                if new_user in users:
                    st.error("Username already exists! Pick another one.")
                elif len(new_user) < 3 or len(new_pass) < 4:
                    st.warning("Username must be at least 3 characters and password at least 4 characters.")
                else:
                    # Save the new user and their scrambled password
                    users[new_user] = hash_password(new_pass)
                    save_users(users)
                    st.success("Account created successfully! You can now log in.")
                    
    # Stop running the rest of the code until someone logs in
    st.stop() 

# =========================================================
# 3. THE MULTI-TENANT DATABASE
# =========================================================
current_user = st.session_state.logged_in_user

st.sidebar.write(f"👤 Logged in as: **{current_user}**")
if st.sidebar.button("Log Out"):
    # 1. Forget the user
    st.session_state.logged_in_user = None
    
    # 2. Wipe the ingestion desk clean
    st.session_state.draft_bullets = []
    
    # 3. Shred the final resume draft
    st.session_state.final_resume = ""
    
    # 4. Refresh the page to lock it down
    st.rerun()

chroma_client = chromadb.PersistentClient(path="./resume_db")
collection_name = f"vault_{current_user}"
collection = chroma_client.get_or_create_collection(name=collection_name)

# =========================================================
# 4. THE MAIN APP (Ingestion & Matchmaker)
# =========================================================
# We keep the rest of the app exactly the same, but now 'collection'
# only points to the current user's personal vault!


if "draft_bullets" not in st.session_state:
    st.session_state.draft_bullets = []
if "final_resume" not in st.session_state:
    st.session_state.final_resume = ""

tab1, tab2, tab3 = st.tabs(["📥 1. Ingestion Desk", "🎯 2. Matchmaker & Export", "🗄️ 3. Vault Manager"])

# 2. SYNCHRONIZED SETTINGS 
density_settings = {
    1: {"margin_in": 0.5,  "font_pt": 10,   "space_after_pt": 0, "line_height": 1.0},
    2: {"margin_in": 0.6,  "font_pt": 10.5, "space_after_pt": 2, "line_height": 1.1},
    3: {"margin_in": 0.75, "font_pt": 11,   "space_after_pt": 4, "line_height": 1.15},
    4: {"margin_in": 0.85, "font_pt": 11.5, "space_after_pt": 6, "line_height": 1.25},
    5: {"margin_in": 1.0,  "font_pt": 12,   "space_after_pt": 8, "line_height": 1.5}
}

font_settings = {
    "Calibri (Modern & Clean)": {"word_name": "Calibri", "css_family": "'Calibri', sans-serif"},
    "Garamond (Classic & Academic)": {"word_name": "Garamond", "css_family": "'Garamond', serif"},
    "Arial (Standard & Safe)": {"word_name": "Arial", "css_family": "'Arial', sans-serif"},
    "Georgia (Elegant & Readable)": {"word_name": "Georgia", "css_family": "'Georgia', serif"},
    "Times New Roman (Traditional)": {"word_name": "Times New Roman", "css_family": "'Times New Roman', serif"}
}

def lock_paragraph_spacing(paragraph, space_after, space_before=0, line_ht=1.15):
    paragraph.paragraph_format.space_before = Pt(space_before)
    paragraph.paragraph_format.space_after = Pt(space_after)
    paragraph.paragraph_format.line_spacing = line_ht

# 3. EXPORT FUNCTION
def create_polished_word_doc(resume_text, level=3, font_choice="Calibri (Modern & Clean)"):
    doc = Document()
    settings = density_settings[level]
    selected_font = font_settings[font_choice]["word_name"]
    
    margin = settings["margin_in"]
    font_size = settings["font_pt"]
    spacing = settings["space_after_pt"]
    line_ht = settings["line_height"]

    # FORCE 8.5 x 11 Page Size (Overrides any international/A4 Mac defaults)
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        section.top_margin = Inches(margin)
        section.bottom_margin = Inches(margin)
        section.left_margin = Inches(margin)
        section.right_margin = Inches(margin)

    right_tab_position = 8.5 - (margin * 2) - 0.05

    style = doc.styles['Normal']
    style.font.name = selected_font
    style.font.size = Pt(font_size)

    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = header.add_run("ALEX")
    header_run.bold = True
    header_run.font.size = Pt(font_size + 6)
    lock_paragraph_spacing(header, space_after=0, line_ht=line_ht)
    
    contact = doc.add_paragraph("Boston, MA | alex@example.com | 555-0199")
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lock_paragraph_spacing(contact, space_after=12, line_ht=line_ht)

    for line in resume_text.split('\n'):
        if not line.strip(): continue 
            
        if line.strip().startswith('###'): 
            clean_header = line.replace('###', '').strip().upper()
            p = doc.add_paragraph()
            run = p.add_run(clean_header)
            run.bold = True
            run.font.size = Pt(font_size + 1)
            lock_paragraph_spacing(p, space_after=2, space_before=12, line_ht=line_ht)
            
        elif line.strip().startswith('**'):
            parts = [p.strip() for p in line.replace('**', '').split('|')]
            p = doc.add_paragraph()
            
            # THE FIX: Delete Word's default invisible tabs before adding ours
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.clear_all() 
            tab_stops.add_tab_stop(Inches(right_tab_position), WD_TAB_ALIGNMENT.RIGHT)
            
            if len(parts) >= 2:
                p.add_run(f"{parts[0]} | {parts[1]}").bold = True
            
            if len(parts) == 3:
                p.add_run(f"\t{parts[2]}").bold = False
            elif len(parts) >= 4:
                right_text = f"\t{parts[2]} | {parts[3]}".replace("N/A | ", "").replace(" | N/A", "")
                p.add_run(right_text).bold = False
                
            lock_paragraph_spacing(p, space_after=2, space_before=6, line_ht=line_ht)

        elif line.strip().startswith('-'): 
            clean_bullet = line.replace('-', '', 1).strip()
            p = doc.add_paragraph(clean_bullet, style='List Bullet')
            lock_paragraph_spacing(p, space_after=spacing, line_ht=line_ht)
            
        else:
            p = doc.add_paragraph(line.strip())
            lock_paragraph_spacing(p, space_after=spacing, line_ht=line_ht)
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 4. LIVE HTML PREVIEW
def render_live_preview(resume_text, level, font_choice):
    settings = density_settings[level]
    css_family = font_settings[font_choice]["css_family"]
    
    margin_in = settings["margin_in"]
    font_pt = settings["font_pt"]
    spacing_pt = settings["space_after_pt"]
    line_ht = settings["line_height"]
    
    html_content = f"""
    <div style="width: 100%; display: flex; justify-content: center; background: #f0f2f6; padding: 20px 0;">
        <div style="
            background-color: white; 
            color: black; 
            width: 8.5in; 
            min-height: 11in; 
            padding: {margin_in}in; 
            box-sizing: border-box;
            box-shadow: 0px 10px 20px rgba(0,0,0,0.15);
            font-family: {css_family};
            font-size: {font_pt}pt;
            line-height: {line_ht};
            zoom: 0.85;
            ">
            
            <div style="text-align: center; margin-bottom: 12pt;">
                <h1 style="margin: 0; font-size: {font_pt + 6}pt; font-weight: bold; letter-spacing: 0.5px; line-height: 1.0;">ALEX</h1>
                <p style="margin: 0; margin-top: 2pt; font-size: {font_pt}pt;">Boston, MA | alex@example.com | 555-0199</p>
            </div>
    """
    
    for line in resume_text.split('\n'):
        if not line.strip(): continue
        
        if line.strip().startswith('###'):
            clean_header = line.replace('###', '').strip().upper()
            html_content += f"<h3 style='margin: 0; margin-top: 12pt; margin-bottom: 2pt; font-size: {font_pt + 1}pt; font-weight: bold;'>{clean_header}</h3>"
            
        elif line.strip().startswith('**'):
            parts = [p.strip() for p in line.replace('**', '').split('|')]
            
            left_text = f"<b>{parts[0]}</b> | <b>{parts[1]}</b>" if len(parts) >= 2 else ""
            right_text = ""
            if len(parts) == 3:
                right_text = parts[2]
            elif len(parts) >= 4:
                right_text = f"{parts[2]} | {parts[3]}".replace("N/A | ", "").replace(" | N/A", "")
                
            html_content += f"<div style='display: flex; justify-content: space-between; margin-top: 6pt; margin-bottom: 2pt;'><span>{left_text}</span><span>{right_text}</span></div>"
            
        elif line.strip().startswith('-'):
            clean_bullet = line.replace('-', '', 1).strip()
            # Hardcoded 0.25in padding to perfectly match Word's List Bullet style
            html_content += f"<ul style='margin: 0; padding-left: 0.25in;'><li style='margin-bottom: {spacing_pt}pt; padding-left: 0.15in;'>{clean_bullet}</li></ul>"
            
        else:
            html_content += f"<p style='margin: 0; margin-bottom: {spacing_pt}pt;'>{line.strip()}</p>"
            
    html_content += """
        </div>
    </div>
    """
    return html_content

# =========================================================
# TAB 1: THE INGESTION DESK (Upgraded with Metadata)
# =========================================================
with tab1:
    st.header("Build Your Vault")
    st.write(f"**Current Vault Size:** {collection.count()} verified bullets.")
    raw_text = st.text_area("Paste raw master resume text, LinkedIn profile, or bio here:", height=150)
    
    if st.button("Shred & Extract Structured Data"):
        if raw_text:
            with st.spinner("AI is extracting bullets and metadata..."):
                # NEW PROMPT: Ask for structured JSON objects
                extraction_prompt = f"""
                You are a precise data extraction tool. Read the text below and extract every distinct professional, academic, or project achievement.
                
                CRITICAL: You must reply ONLY with a valid JSON list of objects. Do not include formatting or markdown.
                If you cannot find a specific piece of metadata (like a date or location), leave it as an empty string "".
                
                Format each object exactly like this:
                [
                    {{
                        "Approve": true,
                        "Role": "Job Title or Role",
                        "Organization": "Company or School",
                        "Location": "City, State",
                        "Date": "Timeframe",
                        "Bullet": "The verbatim achievement text"
                    }}
                ]
                
                TEXT TO PARSE:
                {raw_text}
                """
                try:
                    response = client.models.generate_content(model='gemini-2.5-flash', contents=extraction_prompt)
                    cleaned_response = response.text.replace("```json", "").replace("```", "").strip()
                    
                    # Save the structured list to session state
                    st.session_state.draft_bullets = json.loads(cleaned_response)
                    st.rerun()
                except Exception as e:
                    st.error(f"Error parsing text: {e}\n\nMake sure the AI only returns JSON.")

    # THE VERIFICATION DATA EDITOR
    if st.session_state.draft_bullets:
        st.divider()
        st.subheader("Verification Desk")
        st.write("Review, edit, and approve the extracted data. Check the boxes for the items you want to keep.")
        
        # Convert the JSON to a Pandas DataFrame for the visual editor
        df = pd.DataFrame(st.session_state.draft_bullets)
        
        # Display the interactive spreadsheet
        edited_df = st.data_editor(
            df,
            column_config={
                "Approve": st.column_config.CheckboxColumn("Keep?", default=True),
            },
            hide_index=True,
            use_container_width=True
        )
        
        if st.button("Save Approved Items to Vault", type="primary"):
            # Filter the dataframe to only include rows where 'Approve' is True
            approved_df = edited_df[edited_df["Approve"] == True]
            
            if not approved_df.empty:
                # Prepare the data for ChromaDB
                documents = []
                metadatas = []
                ids = []
                
                current_count = collection.count()
                
                for index, row in approved_df.iterrows():
                    documents.append(row["Bullet"])
                    ids.append(f"bullet_{current_count + index}")
                    
                    # Save the contextual details as hidden metadata
                    metadatas.append({
                        "Role": str(row["Role"]),
                        "Organization": str(row["Organization"]),
                        "Location": str(row["Location"]),
                        "Date": str(row["Date"])
                    })
                
                # Add everything to the Vault
                collection.add(
                    documents=documents,
                    metadatas=metadatas,
                    ids=ids
                )
                
                st.success(f"Successfully added {len(documents)} structured items to your Vault!")
                st.session_state.draft_bullets = [] # Clear the desk
            else:
                st.warning("No items were approved!")

# =========================================================
# TAB 2: THE MATCHMAKER & EXPORT (Upgraded with Metadata)
# =========================================================
with tab2:
    st.header("Draft & Download")
    job_description = st.text_area("Paste Job Description:", height=150)

    if st.button("Generate Polished Resume"):
        if not job_description or collection.count() == 0:
            st.warning("Ensure job description is pasted and Vault is not empty.")
        else:
            with st.spinner("Matching and Formatting..."):
                # 1. Search the database (Now fetching metadata too!)
                results = collection.query(
                    query_texts=[job_description], 
                    n_results=min(6, collection.count())
                )
                
                # 2. Stitch the bullets and metadata together for the AI
                retrieved_docs = results["documents"][0]
                retrieved_meta = results["metadatas"][0]
                
                enriched_bullets = []
                for doc, meta in zip(retrieved_docs, retrieved_meta):
                    # Create a rich text block for each achievement
                    block = f"""
                    Role: {meta.get('Role', 'N/A')}
                    Organization: {meta.get('Organization', 'N/A')}
                    Location: {meta.get('Location', 'N/A')}
                    Date: {meta.get('Date', 'N/A')}
                    Achievement: {doc}
                    """
                    enriched_bullets.append(block)
                
                final_context = "\n---\n".join(enriched_bullets)
                
                # 3. The Final Prompt
                match_prompt = f"""
                You are an expert resume writer. Create a structured, professional resume using ONLY the information provided in the "Enriched Database Matches" below.
                
                STRUCTURE REQUIREMENTS:
                - Add a '### PROFESSIONAL SUMMARY' (1-sentence summary based only on verified facts).
                - Group the achievements logically under '### PROFESSIONAL EXPERIENCE' or '### EDUCATION & RESEARCH'.
                - For each role, format the header exactly like this:
                  **Role** | Organization | Location | Date
                  - [Verbatim Achievement]
                
                CRITICAL RULES:
                1. Use the achievement text VERBATIM.
                2. Do not hallucinate any dates, locations, or roles. If a field says 'N/A', ignore it.
                3. Do not add skills or filler text.
                
                JOB DESCRIPTION:
                {job_description}
                
                ENRICHED DATABASE MATCHES:
                {final_context}
                """
                
                try:
                    response = client.models.generate_content(model='gemini-2.5-flash', contents=match_prompt)
                    st.session_state.final_resume = response.text
                except Exception as e:
                    st.error(f"API Error (You might still be rate limited): {e}")

    if st.session_state.final_resume:
        st.divider()
        
        # New split layout for previewing
        col_preview, col_controls = st.columns([2, 1])
        
        with col_controls:
            st.subheader("Export Settings")
            font_choice = st.selectbox("Select Resume Font:", options=list(font_settings.keys()))
            density_choice = st.slider("Spacing Density (1=Tight, 5=Loose):", 1, 5, 3, 1)
            
            docx_data = create_polished_word_doc(st.session_state.final_resume, level=density_choice, font_choice=font_choice)
            st.download_button(
                label="💾 Download as Word Document", 
                data=docx_data, 
                file_name=f"Resume_{current_user}.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary"
            )
            
        with col_preview:
            st.subheader("Live Preview")
            # This renders the print-perfect HTML preview we perfected in the lab
            preview_html = render_live_preview(st.session_state.final_resume, level=density_choice, font_choice=font_choice)
            st.components.v1.html(preview_html, height=700, scrolling=True)
# =========================================================
# TAB 3: THE VAULT MANAGER 
# =========================================================
with tab3:
    st.header("🗄️ Your Personal Vault")
    
    # 1. Fetch current contents
    vault_data = collection.get()
    
    if not vault_data['ids']:
        st.info("Your vault is currently empty.")
    else:
        # Create a dataframe for display
        df = pd.DataFrame({
            "ID": vault_data['ids'],
            "Bullet": vault_data['documents'],
            "Role": [m.get('Role', '') for m in vault_data['metadatas']],
            "Organization": [m.get('Organization', '') for m in vault_data['metadatas']],
            "Delete?": [False] * len(vault_data['ids'])
        })
        
        st.write("Review your stored achievements. Check the box to mark for deletion.")
        
        # 2. Interactive editor
        edited_df = st.data_editor(df, hide_index=True, use_container_width=True)
        
        # 3. Action button
        if st.button("Apply Changes (Delete Checked Items)"):
            to_delete = edited_df[edited_df["Delete?"] == True]["ID"].tolist()
            
            if to_delete:
                collection.delete(ids=to_delete)
                st.success(f"Deleted {len(to_delete)} items from your vault!")
                st.rerun()
            else:
                st.warning("No items selected for deletion.")