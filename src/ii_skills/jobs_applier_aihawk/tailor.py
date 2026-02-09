"""
Jobs Applier AIHawk - Resume & Cover Letter Tailoring Engine

Reads base DOCX documents, analyzes a job description, and produces
tailored DOCX outputs with content rewritten to match the target role.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

from .config import (
    BASE_RESUME,
    BASE_COVER_LETTER,
    OUTPUT_DIR,
    RESUME_SECTIONS,
    MAX_COVER_LETTER_WORDS,
    EMPHASIS_KEYWORDS_LIMIT,
    OUTPUT_RESUME_TEMPLATE,
    OUTPUT_COVER_LETTER_TEMPLATE,
    OUTPUT_ANALYSIS_TEMPLATE,
)


# ── DOCX Helpers ─────────────────────────────────────────────

def read_docx_text(path: Path) -> str:
    """Extract all text from a DOCX file, preserving paragraph breaks."""
    if not HAS_DOCX:
        raise ImportError("python-docx is required. Run: pip install python-docx")
    if not path.exists():
        raise FileNotFoundError(f"Document not found: {path}")
    doc = Document(str(path))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def read_docx_structured(path: Path) -> List[Dict[str, str]]:
    """Extract paragraphs with their style names for structural analysis."""
    if not HAS_DOCX:
        raise ImportError("python-docx is required. Run: pip install python-docx")
    if not path.exists():
        raise FileNotFoundError(f"Document not found: {path}")
    doc = Document(str(path))
    paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip():
            paragraphs.append({
                "text": p.text.strip(),
                "style": p.style.name if p.style else "Normal",
                "bold": any(run.bold for run in p.runs if run.bold),
            })
    return paragraphs


def write_tailored_docx(sections: List[Dict], output_path: Path) -> Path:
    """Write tailored content to a new DOCX file.

    Each section dict has keys: 'heading' (str), 'content' (str or list of str).
    """
    if not HAS_DOCX:
        raise ImportError("python-docx is required. Run: pip install python-docx")

    doc = Document()

    # Set default font
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)

    for section in sections:
        heading = section.get("heading", "")
        content = section.get("content", "")

        if heading:
            doc.add_heading(heading, level=2)

        if isinstance(content, list):
            for item in content:
                doc.add_paragraph(item, style="List Bullet")
        else:
            for para_text in content.split("\n"):
                if para_text.strip():
                    doc.add_paragraph(para_text.strip())

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ── Job Description Analysis ─────────────────────────────────

def analyze_job_description(job_description: str) -> Dict:
    """Parse a job description into structured components.

    Returns dict with: role, company, requirements, skills, qualifications,
    responsibilities, keywords.
    """
    result = {
        "role": "",
        "company": "",
        "requirements": [],
        "skills": [],
        "qualifications": [],
        "responsibilities": [],
        "keywords": [],
        "raw_text": job_description.strip(),
    }

    lines = job_description.strip().split("\n")

    # Heuristic extraction
    current_section = None
    section_map = {
        "requirement": "requirements",
        "qualification": "qualifications",
        "responsibilit": "responsibilities",
        "skill": "skills",
        "what you": "requirements",
        "who you": "qualifications",
        "about the role": "responsibilities",
        "what you'll do": "responsibilities",
        "what we're looking": "requirements",
    }

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        lower = stripped.lower()

        # Detect section headers
        for key, section in section_map.items():
            if key in lower and len(stripped) < 80:
                current_section = section
                break

        # Collect bullet points and content lines
        if current_section and (stripped.startswith(("-", "*", "\u2022")) or
                                (len(stripped) > 10 and current_section)):
            clean = re.sub(r"^[-*\u2022]\s*", "", stripped)
            if clean and len(clean) > 5:
                result[current_section].append(clean)

    # Extract keywords: capitalized multi-word terms, acronyms, tech terms
    all_text = job_description.lower()
    keyword_patterns = [
        r'\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)+\b',  # Multi-word proper nouns
        r'\b[A-Z]{2,}\b',                          # Acronyms
    ]
    keywords = set()
    for pattern in keyword_patterns:
        matches = re.findall(pattern, job_description)
        keywords.update(m.strip() for m in matches if len(m) > 2)

    result["keywords"] = sorted(keywords)[:EMPHASIS_KEYWORDS_LIMIT]

    return result


# ── Tailoring Logic ──────────────────────────────────────────

def build_tailoring_context(
    base_resume_text: str,
    job_analysis: Dict,
) -> Dict:
    """Build context dict that Claude Code uses to tailor documents.

    This function prepares the structured prompt context. The actual
    LLM tailoring is done by Claude Code itself when the user invokes
    the skill — Claude reads this context and rewrites the resume/cover
    letter sections accordingly.
    """
    return {
        "base_resume": base_resume_text,
        "job_role": job_analysis.get("role", ""),
        "job_company": job_analysis.get("company", ""),
        "job_requirements": job_analysis.get("requirements", []),
        "job_skills": job_analysis.get("skills", []),
        "job_qualifications": job_analysis.get("qualifications", []),
        "job_responsibilities": job_analysis.get("responsibilities", []),
        "job_keywords": job_analysis.get("keywords", []),
        "job_raw": job_analysis.get("raw_text", ""),
        "sections_to_tailor": RESUME_SECTIONS,
        "max_cover_letter_words": MAX_COVER_LETTER_WORDS,
    }


def generate_output_paths(
    company: str,
    role: str,
) -> Dict[str, Path]:
    """Generate standardized output file paths."""
    date_str = datetime.now().strftime("%Y%m%d")
    clean = lambda s: re.sub(r"[^\w]+", "_", s.strip())[:30]

    return {
        "resume": OUTPUT_DIR / OUTPUT_RESUME_TEMPLATE.format(
            company=clean(company), role=clean(role), date=date_str
        ),
        "cover_letter": OUTPUT_DIR / OUTPUT_COVER_LETTER_TEMPLATE.format(
            company=clean(company), role=clean(role), date=date_str
        ),
        "analysis": OUTPUT_DIR / OUTPUT_ANALYSIS_TEMPLATE.format(
            company=clean(company), role=clean(role), date=date_str
        ),
    }


# ── High-Level Orchestration ─────────────────────────────────

def prepare_tailoring_session(
    job_description: str,
    resume_path: Optional[Path] = None,
    cover_letter_path: Optional[Path] = None,
) -> Dict:
    """Prepare everything needed for a tailoring session.

    Returns a dict with:
    - job_analysis: Structured JD breakdown
    - base_resume_text: Full text of base resume
    - base_cover_letter_text: Full text of base cover letter
    - tailoring_context: Combined context for LLM tailoring
    - output_paths: Where to save generated files
    - resume_structure: Paragraph-level structure of base resume
    """
    resume_path = resume_path or BASE_RESUME
    cover_letter_path = cover_letter_path or BASE_COVER_LETTER

    # Read base documents
    base_resume_text = read_docx_text(resume_path)
    base_cover_letter_text = read_docx_text(cover_letter_path)
    resume_structure = read_docx_structured(resume_path)

    # Analyze job description
    job_analysis = analyze_job_description(job_description)

    # Build tailoring context
    context = build_tailoring_context(base_resume_text, job_analysis)

    # Generate output paths
    company = job_analysis.get("company", "Company")
    role = job_analysis.get("role", "Role")
    output_paths = generate_output_paths(company, role)

    return {
        "job_analysis": job_analysis,
        "base_resume_text": base_resume_text,
        "base_cover_letter_text": base_cover_letter_text,
        "tailoring_context": context,
        "output_paths": output_paths,
        "resume_structure": resume_structure,
    }
