"""
Jobs Applier AIHawk Skill

Tailors resumes and cover letters to specific job descriptions using the
user's base documents as source material.

Actions:
- tailor_resume: Generate a job-tailored resume
- tailor_cover_letter: Generate a job-tailored cover letter
- full_application_package: Generate both in one pass
- analyze_job: Extract and summarize JD requirements
- list_resources: List available base documents
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional

from ii_skills import BaseSkill, register_skill

logger = logging.getLogger(__name__)


@register_skill
class JobsApplierSkill(BaseSkill):
    """Resume and cover letter tailoring skill."""

    name = "jobs_applier_aihawk"
    version = "1.0.0"
    description = (
        "Tailors resumes and cover letters to specific job descriptions. "
        "Reads base DOCX documents, analyzes the target JD, and produces "
        "tailored outputs emphasizing relevant experience and skills."
    )

    SKILL_DIR = Path(__file__).parent
    OUTPUT_DIR = SKILL_DIR / "output"
    RESOURCES_DIR = SKILL_DIR / "resources"

    def __init__(self, config: Optional[Dict] = None):
        super().__init__(config)
        self.OUTPUT_DIR.mkdir(exist_ok=True)

    def validate_config(self) -> List[str]:
        """Check that base documents exist."""
        from .config import BASE_RESUME, BASE_COVER_LETTER

        issues = []
        if not BASE_RESUME.exists():
            issues.append(f"Base resume not found: {BASE_RESUME}")
        if not BASE_COVER_LETTER.exists():
            issues.append(f"Base cover letter not found: {BASE_COVER_LETTER}")
        try:
            from docx import Document  # noqa: F401
        except ImportError:
            issues.append("python-docx not installed. Run: pip install python-docx")
        return issues

    def get_capabilities(self) -> List[str]:
        return [
            "tailor_resume",
            "tailor_cover_letter",
            "full_application_package",
            "analyze_job",
            "list_resources",
        ]

    def execute(self, action: str, **kwargs) -> Dict:
        """Route action to handler."""
        self._action_count += 1
        self._last_action = action

        actions = {
            "tailor_resume": self._tailor_resume,
            "tailor_cover_letter": self._tailor_cover_letter,
            "full_application_package": self._full_application_package,
            "analyze_job": self._analyze_job,
            "list_resources": self._list_resources,
        }

        handler = actions.get(action)
        if not handler:
            return {
                "success": False,
                "error": f"Unknown action: '{action}'. Available: {list(actions.keys())}",
            }

        try:
            return handler(**kwargs)
        except Exception as e:
            logger.exception(f"Action '{action}' failed")
            return {"success": False, "error": str(e)}

    # ── Action Handlers ──────────────────────────────────────

    def _tailor_resume(self, job_description: str, **kwargs) -> Dict:
        """Prepare a tailoring session for resume only."""
        from .tailor import prepare_tailoring_session

        session = prepare_tailoring_session(
            job_description=job_description,
            resume_path=kwargs.get("resume_path"),
        )

        return {
            "success": True,
            "action": "tailor_resume",
            "job_analysis": session["job_analysis"],
            "base_resume_text": session["base_resume_text"],
            "resume_structure": session["resume_structure"],
            "tailoring_context": session["tailoring_context"],
            "output_path": str(session["output_paths"]["resume"]),
            "instructions": (
                "Use the job_analysis and base_resume_text to rewrite each "
                "resume section, emphasizing skills and experience that match "
                "the job requirements. Write the result to output_path using "
                "tailor.write_tailored_docx()."
            ),
        }

    def _tailor_cover_letter(self, job_description: str, **kwargs) -> Dict:
        """Prepare a tailoring session for cover letter only."""
        from .tailor import prepare_tailoring_session

        session = prepare_tailoring_session(
            job_description=job_description,
            cover_letter_path=kwargs.get("cover_letter_path"),
        )

        return {
            "success": True,
            "action": "tailor_cover_letter",
            "job_analysis": session["job_analysis"],
            "base_cover_letter_text": session["base_cover_letter_text"],
            "tailoring_context": session["tailoring_context"],
            "output_path": str(session["output_paths"]["cover_letter"]),
            "instructions": (
                "Use the job_analysis and base_cover_letter_text to write a "
                "tailored cover letter (3 paragraphs, under 400 words). "
                "Reference the specific company and role. Write to output_path "
                "using tailor.write_tailored_docx()."
            ),
        }

    def _full_application_package(self, job_description: str, **kwargs) -> Dict:
        """Prepare a full tailoring session (resume + cover letter)."""
        from .tailor import prepare_tailoring_session

        session = prepare_tailoring_session(
            job_description=job_description,
            resume_path=kwargs.get("resume_path"),
            cover_letter_path=kwargs.get("cover_letter_path"),
        )

        return {
            "success": True,
            "action": "full_application_package",
            "job_analysis": session["job_analysis"],
            "base_resume_text": session["base_resume_text"],
            "base_cover_letter_text": session["base_cover_letter_text"],
            "resume_structure": session["resume_structure"],
            "tailoring_context": session["tailoring_context"],
            "output_paths": {
                k: str(v) for k, v in session["output_paths"].items()
            },
            "instructions": (
                "1. Rewrite each resume section to match the JD requirements. "
                "2. Write a tailored cover letter (3 paragraphs, under 400 words). "
                "3. Save both to output_paths using tailor.write_tailored_docx()."
            ),
        }

    def _analyze_job(self, job_description: str, **kwargs) -> Dict:
        """Analyze a job description without tailoring."""
        from .tailor import analyze_job_description, generate_output_paths

        analysis = analyze_job_description(job_description)
        paths = generate_output_paths(
            analysis.get("company", "Company"),
            analysis.get("role", "Role"),
        )

        return {
            "success": True,
            "action": "analyze_job",
            "job_analysis": analysis,
            "analysis_output_path": str(paths["analysis"]),
        }

    def _list_resources(self, **kwargs) -> Dict:
        """List available base documents in resources/."""
        resources = []
        if self.RESOURCES_DIR.exists():
            for f in sorted(self.RESOURCES_DIR.iterdir()):
                if f.is_file() and not f.name.startswith("."):
                    resources.append({
                        "name": f.name,
                        "path": str(f),
                        "size_kb": round(f.stat().st_size / 1024, 1),
                    })

        return {
            "success": True,
            "action": "list_resources",
            "resources": resources,
            "resources_dir": str(self.RESOURCES_DIR),
        }


# ── Convenience accessor ─────────────────────────────────────

def get_skill(config: Optional[Dict] = None) -> JobsApplierSkill:
    """Get an initialized JobsApplierSkill instance."""
    skill = JobsApplierSkill(config)
    skill.initialize()
    return skill
