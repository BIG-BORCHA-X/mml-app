from datetime import datetime, timedelta
import json
import re
from docx import Document

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

Your task is to create a structured Action Plan with the following columns:
- What
- Why
- How
- When
- Success Criteria

Each row should contain a high-priority action the business should take, based on their stated needs.

Before generating the actions:
- Read the workshop capture below
- Extract the business's **key focus areas** (they may be labelled "Focus Areas" or "Actions" or "Action Plan")
- Then generate **one action per focus area**, ordered by priority

Instructions:
- The “How” field should use **concise bullet points**, each a single sentence (no full paragraphs)
- The “When” field should use approximate default timeframes like “in 2 weeks” or “in 1 month” if no clear deadline is found in the minutes
- The “Success Criteria” should describe how to know the action was completed successfully

Return the result as a list of Python dictionaries, one per row, like this:
[
  {{
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