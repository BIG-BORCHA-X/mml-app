import streamlit as st
from datetime import datetime, timedelta
import openai
import os
import json
import re
from docx import Document
from generate_action_plan import write_action_plan_docx
from generate_strategy_2 import generate_strategy_docx
from dotenv import load_dotenv
import tempfile

# MODEL = "gpt-4o-mini"
MODEL = "gpt-4o"

# === Setup ===
# load_dotenv()
# openai.api_key = os.getenv("OPENAI_API_KEY")
OPENAI_API_KEY = st.secrets["openai_api_key"]
CORRECT_PASSWORD = st.secrets["app_password"]

# Set OpenAI key
openai.api_key = OPENAI_API_KEY

st.set_page_config(page_title="Action Plan Generator", layout="centered")

# === Utilities ===
def read_minutes(file_path):
    doc = Document(file_path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

def get_day_suffix(day):
    if 11 <= day <= 13:
        return "th"
    last_digit = day % 10
    return {1: "st", 2: "nd", 3: "rd"}.get(last_digit, "th")

def convert_when_to_date():
    target = datetime.today() + timedelta(weeks=4)
    day = target.day
    suffix = get_day_suffix(day)
    return f"{target.strftime('%B')} {day}{suffix}"

def extract_json_from_response(content):
    match = re.search(r"\[\s*\{.*?\}\s*\]", content, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            return []
    return []

# === Prompt Template ===
def build_prompt(minutes, company_name):
    return f"""
You are a professional business strategist who has just run a workshop for a business called "{company_name}". Below is the capture of their business planning workshop.

All writing should use British English spelling and conventions. Where appropriate, expand upon the ideas captured during the workshop to ensure clarity, completeness, and usefulness.

Your task is to create a structured Action Plan with the following columns:
- Priority
- What
- Why
- How
- When
- Success Criteria

Order the actions by priority:
- Red: High Priority
- Yellow: Medium
- Green: Low

Before generating the actions:
- Read the workshop capture below
- Extract the business's **key focus areas** (they may be labelled "Focus Areas", "Actions", or "Action Plan")
- Then generate **one action per focus area**, ordered by priority (high first, low last)
- If fewer than 6 focus areas are found, add additional actions based on any other important themes or needs identified in the workshop (to ensure at least 6 total actions are included)

Instructions:
- The “How” field should use **concise bullet points**, each a single sentence (no full paragraphs)
- The “When” field should use approximate default timeframes like “in 2 weeks” or “in 1 month” if no clear deadline is found in the minutes
- The “Success Criteria” should describe how to know the action was completed successfully

Return the result as a list of Python dictionaries, one per row, like this:
[
  {{
    "Priority: "...",
    "What": "...",
    "Why": "...",
    "How": ["...bullet point...", "...bullet point..."],
    "When": "...",
    "Success Criteria": "..."
  }},
  ...
]

Workshop Capture:
\"\"\"
{minutes}
\"\"\"
"""

# === Streamlit UI ===
# Create a password input field
password = st.text_input("🔒 Enter password to access the app:", type="password")

# Check if password is correct
if password != CORRECT_PASSWORD:
    st.warning("Access denied. Please enter the correct password to continue.")
    st.stop()

st.title("📋 Workshop Document Generator")
st.write("Upload a `.docx` minutes document and choose a document to generate.")

company_name = st.text_input("Company name", placeholder="e.g., Pal's Pickling Plant")
uploaded_file = st.file_uploader("Upload workshop minutes (.docx)", type=["docx"])

if uploaded_file and company_name:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.read())
        minutes_path = tmp.name

    minutes = read_minutes(minutes_path)

    st.header("🧩 Generate Action Plan")
    if st.button("Generate Action Plan"):
        with st.spinner("Generating Action Plan..."):
            prompt = build_prompt(minutes, company_name)
            response = openai.ChatCompletion.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=1500
            )

            content = response['choices'][0]['message']['content']
            raw_rows = extract_json_from_response(content)

            for row in raw_rows:
                row["When"] = convert_when_to_date()

            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            action_filename = f"{company_name} - Action Plan - {timestamp}.docx"
            docx_buffer = write_action_plan_docx(action_filename, raw_rows)

            st.download_button(
                label="📄 Download Action Plan",
                data=docx_buffer,
                file_name=action_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        st.success(f"✅ Action Plan Generated as: {action_filename}")

    st.header("📄 Generate Strategy Report")
    if st.button("Generate Strategy Report"):
        status_area = st.empty()
        with st.spinner("Generating Strategy Report..."):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            strategy_filename = f"{company_name} - Strategy Report - {timestamp}.docx"
            docx_buffer2 = generate_strategy_docx(minutes, strategy_filename, company_name, status_area)
            st.download_button(
                label="📄 Download Strategy Report",
                data=docx_buffer2,
                file_name=strategy_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            status_area.text("")
        st.success(f"📄 Strategy Report Generated as: {strategy_filename}")

# streamlit run app_test.py

# timestamp = datetime.now().strftime("%Y%m%d_%H%M")
# action_filename = f"Action Plan - {company_name}_{timestamp}.docx"

# timestamp = datetime.now().strftime("%Y%m%d_%H%M")
# strategy_filename = f"Strategy Report - {company_name}_{timestamp}.docx"
# status_area = st.empty()
# generate_strategy_docx(minutes, strategy_filename, status_area)

# streamlit run app.py