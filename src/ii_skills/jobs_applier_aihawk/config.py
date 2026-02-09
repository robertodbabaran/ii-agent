"""
Jobs Applier AIHawk - Configuration

Directories, default settings, and resume section definitions.
"""

from pathlib import Path

# ── Directories ──────────────────────────────────────────────
SKILL_DIR = Path(__file__).parent
RESOURCES_DIR = SKILL_DIR / "resources"
OUTPUT_DIR = SKILL_DIR / "output"
TEMPLATES_DIR = SKILL_DIR / "templates"

# ── Base Documents ───────────────────────────────────────────
BASE_RESUME = RESOURCES_DIR / "User Resume (Recruiter).docx"
BASE_COVER_LETTER = RESOURCES_DIR / "Sample Cover Letter (Generalist PE).docx"

# ── Resume Sections (order matters for output) ───────────────
RESUME_SECTIONS = [
    "header",           # Name, contact info, LinkedIn
    "summary",          # Professional summary / objective
    "experience",       # Work experience
    "education",        # Education & certifications
    "skills",           # Technical & soft skills
    "achievements",     # Awards, recognition
]

# ── Tailoring Settings ───────────────────────────────────────
MAX_RESUME_PAGES = 2
MAX_COVER_LETTER_WORDS = 400
EMPHASIS_KEYWORDS_LIMIT = 15   # Max keywords to extract from JD

# ── Output Naming ────────────────────────────────────────────
OUTPUT_RESUME_TEMPLATE = "{company}_{role}_Resume_{date}.docx"
OUTPUT_COVER_LETTER_TEMPLATE = "{company}_{role}_Cover_Letter_{date}.docx"
OUTPUT_ANALYSIS_TEMPLATE = "{company}_{role}_JD_Analysis_{date}.md"
