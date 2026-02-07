"""
Load and fill prompt templates with PrepContext placeholders.
"""

from pathlib import Path
from typing import Dict, Optional

from .config import PROMPTS_DIR, SECTIONS
from .models import PrepContext


def load_template(
    section_id: str,
    context: PrepContext,
    extra_placeholders: Optional[Dict[str, str]] = None,
) -> str:
    """Load a prompt template and fill placeholders with context data.

    Args:
        section_id: Template ID (T01-T09).
        context: PrepContext with company/role info.
        extra_placeholders: Additional key-value pairs to substitute.

    Returns:
        Filled template text.
    """
    section = SECTIONS.get(section_id)
    if not section:
        raise ValueError(f"Unknown section_id '{section_id}'. Available: {list(SECTIONS)}")

    # Find template file matching section_id prefix
    template_path = _find_template(section_id)
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    text = template_path.read_text(encoding="utf-8")

    # Build placeholder map
    placeholders = {
        "COMPANY_NAME": context.company_name,
        "ROLE_TITLE": context.role_title,
        "TARGET_FUNCTION": context.target_function,
        "JOB_DESCRIPTION": context.job_description or "(Not provided)",
        "INTERVIEW_DATE": context.interview_date or "(Not set)",
        "INTERVIEWER_NAMES": ", ".join(context.interviewer_names) if context.interviewer_names else "(Not provided)",
        "BACKGROUND_POINTS": _format_list(context.background_points, "background points"),
        "URL_LIST": _format_list(context.urls, "URLs"),
        "RESEARCH_NOTES": "(Will be populated from saved research)",
    }

    if extra_placeholders:
        placeholders.update(extra_placeholders)

    # Substitute {{PLACEHOLDER}} patterns
    for key, value in placeholders.items():
        text = text.replace(f"{{{{{key}}}}}", value)

    return text


def list_templates() -> Dict[str, Dict[str, str]]:
    """Return available templates with metadata.

    Returns:
        Dict of section_id -> {name, filename, target_words}.
    """
    result = {}
    for sid, meta in SECTIONS.items():
        template_path = _find_template(sid)
        result[sid] = {
            "name": meta["name"],
            "target_words": meta["target_words"],
            "template_exists": template_path.exists(),
        }
    return result


def _find_template(section_id: str) -> Path:
    """Find template file by section_id prefix."""
    for f in PROMPTS_DIR.iterdir():
        if f.name.startswith(section_id) and f.suffix == ".md":
            return f
    # Fallback to expected naming convention
    return PROMPTS_DIR / f"{section_id}.md"


def _format_list(items: list, label: str) -> str:
    """Format a list as bullet points or a placeholder."""
    if not items:
        return f"(No {label} provided)"
    return "\n".join(f"- {item}" for item in items)
