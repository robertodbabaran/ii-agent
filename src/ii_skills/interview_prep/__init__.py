"""
Interview Prep Skill — Structured interview preparation materials.

Generates company research briefs, Q&A frameworks, cheat sheets,
and Anki flashcards for interview preparation. All output goes to
a Tier 3 case workspace under output/cases/.

10 Actions:
  start_prep, load_template, save_research, write_section,
  export_flashcards, compile_brief, get_prep_status, list_preps,
  get_research_context, update_context
"""

__version__ = "1.0.0"

import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

from ii_skills import BaseSkill, register_skill

from .config import (
    CASE_CONFIG_FILENAME,
    CASE_SUBFOLDERS,
    CASES_DIR,
    OUTPUT_DIR,
    RESEARCH_NOTES_FILENAME,
    SECTIONS,
    SKILL_DIR,
)
from .models import InterviewBrief, PrepContext, SectionStatus

logger = logging.getLogger(__name__)


@register_skill
class InterviewPrepSkill(BaseSkill):
    """Structured interview preparation with research briefs, Q&A, and flashcards."""

    name = "interview_prep"
    version = __version__
    description = (
        "Generate structured interview preparation materials: company research, "
        "Q&A frameworks, cheat sheets, and Anki flashcards."
    )

    SKILL_DIR = SKILL_DIR
    OUTPUT_DIR = OUTPUT_DIR

    def __init__(self, config: Optional[Dict] = None):
        super().__init__(config)
        self.OUTPUT_DIR.mkdir(exist_ok=True)

    def get_capabilities(self) -> List[str]:
        return [
            "start_prep",
            "load_template",
            "save_research",
            "write_section",
            "export_flashcards",
            "compile_brief",
            "get_prep_status",
            "list_preps",
            "get_research_context",
            "update_context",
        ]

    def execute(self, action: str, **kwargs) -> Dict:
        """Execute an interview prep action."""
        actions = {
            "start_prep": self._start_prep,
            "load_template": self._load_template,
            "save_research": self._save_research,
            "write_section": self._write_section,
            "export_flashcards": self._export_flashcards,
            "compile_brief": self._compile_brief,
            "get_prep_status": self._get_prep_status,
            "list_preps": self._list_preps,
            "get_research_context": self._get_research_context,
            "update_context": self._update_context,
        }

        handler = actions.get(action)
        if not handler:
            return {
                "success": False,
                "error": f"Unknown action '{action}'. Available: {list(actions)}",
            }

        try:
            return handler(**kwargs)
        except Exception as e:
            logger.exception(f"Action '{action}' failed")
            return {"success": False, "error": str(e)}

    # ------------------------------------------------------------------
    # Action 1: start_prep
    # ------------------------------------------------------------------
    def _start_prep(
        self,
        company_name: str,
        role_title: str,
        target_function: str = "",
        urls: Optional[List[str]] = None,
        job_description: str = "",
        interview_date: str = "",
        interviewer_names: Optional[List[str]] = None,
        interview_format: str = "",
        background_points: Optional[List[str]] = None,
        **kwargs,
    ) -> Dict:
        """Create case workspace and initialize prep context."""
        # Build safe folder name
        safe_company = company_name.replace(" ", "_").replace("/", "-").replace("\\", "-")
        date_prefix = datetime.now().strftime("%Y-%m-%d")
        folder_name = f"{date_prefix}_{safe_company}_interview_prep"

        case_folder = CASES_DIR / folder_name
        case_folder.mkdir(parents=True, exist_ok=True)

        # Create subfolders
        for sub in CASE_SUBFOLDERS:
            (case_folder / sub).mkdir(exist_ok=True)

        # Build context
        context = PrepContext(
            company_name=company_name,
            role_title=role_title,
            target_function=target_function or "",
            urls=urls or [],
            job_description=job_description,
            interview_date=interview_date,
            interviewer_names=interviewer_names or [],
            interview_format=interview_format,
            background_points=background_points or [],
            case_folder=str(case_folder),
        )

        # Initialize section statuses
        sections = []
        for sid, meta in SECTIONS.items():
            sections.append(SectionStatus(
                section_id=sid,
                section_name=meta["name"],
                status="pending",
            ))

        # Build brief state
        brief = InterviewBrief(context=context, sections=sections)

        # Write case config
        config_path = case_folder / CASE_CONFIG_FILENAME
        config_path.write_text(
            json.dumps(brief.to_dict(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

        # Initialize empty research notes
        notes_path = case_folder / RESEARCH_NOTES_FILENAME
        notes_path.write_text(
            json.dumps({}, indent=2),
            encoding="utf-8",
        )

        return {
            "success": True,
            "case_folder": str(case_folder),
            "sections_status": [s.to_dict() for s in sections],
            "message": f"Interview prep workspace created for {company_name} — {role_title}",
        }

    # ------------------------------------------------------------------
    # Action 2: load_template
    # ------------------------------------------------------------------
    def _load_template(
        self,
        section_id: str,
        prep_folder: str,
        extra_placeholders: Optional[Dict[str, str]] = None,
        **kwargs,
    ) -> Dict:
        """Load a prompt template with context placeholders filled."""
        from .template_loader import load_template

        context = self._load_context(prep_folder)
        if not context:
            return {
                "success": False,
                "error": f"No case_config.json found in {prep_folder}",
            }

        # Inject saved research notes into placeholder
        research_notes = self._load_research_notes(prep_folder)
        if research_notes:
            research_text = "\n\n".join(
                f"**Source:** {url}\n{content}"
                for url, content in research_notes.items()
            )
        else:
            research_text = "(No research saved yet)"

        extra = extra_placeholders or {}
        extra["RESEARCH_NOTES"] = research_text

        prompt_text = load_template(section_id, context, extra)

        section_meta = SECTIONS.get(section_id, {})
        return {
            "success": True,
            "prompt_text": prompt_text,
            "section_id": section_id,
            "section_name": section_meta.get("name", ""),
            "target_words": section_meta.get("target_words", 0),
        }

    # ------------------------------------------------------------------
    # Action 3: save_research
    # ------------------------------------------------------------------
    def _save_research(
        self,
        prep_folder: str,
        url: str,
        content: str,
        **kwargs,
    ) -> Dict:
        """Store raw web research for later reference."""
        notes = self._load_research_notes(prep_folder)
        notes[url] = content

        notes_path = Path(prep_folder) / RESEARCH_NOTES_FILENAME
        notes_path.write_text(
            json.dumps(notes, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

        return {
            "success": True,
            "url": url,
            "char_count": len(content),
            "total_sources": len(notes),
        }

    # ------------------------------------------------------------------
    # Action 4: write_section
    # ------------------------------------------------------------------
    def _write_section(
        self,
        prep_folder: str,
        section_id: str,
        content: str,
        **kwargs,
    ) -> Dict:
        """Write a completed section to the case workspace."""
        from .section_writer import write_section

        result = write_section(prep_folder, section_id, content)

        # Update section status in case config
        self._update_section_status(
            prep_folder,
            section_id,
            status="complete",
            file_path=result["file_path"],
            word_count=result["word_count"],
        )

        return {
            "success": True,
            "section_id": section_id,
            "file_path": result["file_path"],
            "word_count": result["word_count"],
        }

    # ------------------------------------------------------------------
    # Action 5: export_flashcards
    # ------------------------------------------------------------------
    def _export_flashcards(
        self,
        prep_folder: str,
        cards: List[Dict[str, str]],
        **kwargs,
    ) -> Dict:
        """Generate Anki TSV and readable Markdown from card data."""
        from .flashcard_exporter import export_markdown, export_tsv

        fc_dir = Path(prep_folder) / "05_flashcards"
        fc_dir.mkdir(parents=True, exist_ok=True)

        tsv_path = export_tsv(cards, str(fc_dir / "interview_prep_cards.tsv"))
        md_path = export_markdown(cards, str(fc_dir / "interview_prep_cards.md"))

        # Update T08 section status
        self._update_section_status(
            prep_folder,
            "T08",
            status="complete",
            file_path=tsv_path,
            word_count=len(cards),  # card count instead of word count
        )

        return {
            "success": True,
            "tsv_path": tsv_path,
            "md_path": md_path,
            "card_count": len(cards),
        }

    # ------------------------------------------------------------------
    # Action 6: compile_brief
    # ------------------------------------------------------------------
    def _compile_brief(self, prep_folder: str, **kwargs) -> Dict:
        """Merge all completed sections into a single brief document."""
        from .brief_compiler import compile_brief

        context = self._load_context(prep_folder)
        context_dict = context.to_dict() if context else None

        result = compile_brief(prep_folder, context_dict)

        return {
            "success": True,
            "brief_path": result["brief_path"],
            "total_words": result["total_words"],
            "sections_included": result["sections_included"],
        }

    # ------------------------------------------------------------------
    # Action 7: get_prep_status
    # ------------------------------------------------------------------
    def _get_prep_status(self, prep_folder: str, **kwargs) -> Dict:
        """Show progress on all sections."""
        brief = self._load_brief(prep_folder)
        if not brief:
            return {
                "success": False,
                "error": f"No case_config.json found in {prep_folder}",
            }

        return {
            "success": True,
            "context": brief.context.to_dict(),
            "sections": [s.to_dict() for s in brief.sections],
            "completion_pct": brief.completion_pct,
        }

    # ------------------------------------------------------------------
    # Action 8: list_preps
    # ------------------------------------------------------------------
    def _list_preps(self, **kwargs) -> Dict:
        """List all interview prep cases."""
        preps = []

        if not CASES_DIR.exists():
            return {"success": True, "preps": []}

        for folder in sorted(CASES_DIR.iterdir(), reverse=True):
            if not folder.is_dir() or not folder.name.endswith("_interview_prep"):
                continue

            config_path = folder / CASE_CONFIG_FILENAME
            if not config_path.exists():
                continue

            try:
                data = json.loads(config_path.read_text(encoding="utf-8"))
                ctx = data.get("context", {})
                sections = data.get("sections", [])
                complete = sum(1 for s in sections if s.get("status") == "complete")
                total = len(sections) if sections else len(SECTIONS)

                preps.append({
                    "company": ctx.get("company_name", "Unknown"),
                    "role": ctx.get("role_title", "Unknown"),
                    "date": ctx.get("interview_date", ""),
                    "folder": str(folder),
                    "created_at": ctx.get("created_at", ""),
                    "completion": f"{complete}/{total}",
                    "completion_pct": round(complete / total * 100, 1) if total else 0,
                })
            except (json.JSONDecodeError, KeyError):
                continue

        return {"success": True, "preps": preps}

    # ------------------------------------------------------------------
    # Action 9: get_research_context
    # ------------------------------------------------------------------
    def _get_research_context(
        self,
        prep_folder: str,
        section_id: Optional[str] = None,
        **kwargs,
    ) -> Dict:
        """Retrieve stored research and completed sections."""
        from .section_writer import get_all_sections, read_section

        research_notes = self._load_research_notes(prep_folder)

        if section_id:
            content = read_section(prep_folder, section_id)
            sections = {}
            if content:
                meta = SECTIONS.get(section_id, {})
                sections[section_id] = {
                    "name": meta.get("name", ""),
                    "content": content,
                    "word_count": len(content.split()),
                }
        else:
            sections = get_all_sections(prep_folder)

        return {
            "success": True,
            "research_notes": research_notes,
            "research_source_count": len(research_notes),
            "sections": {
                sid: {k: v for k, v in sec.items() if k != "content"}
                for sid, sec in sections.items()
            },
            "sections_with_content": sections,
        }

    # ------------------------------------------------------------------
    # Action 10: update_context
    # ------------------------------------------------------------------
    def _update_context(
        self,
        prep_folder: str,
        updates: Dict,
        **kwargs,
    ) -> Dict:
        """Update prep metadata (add interviewer names, change date, etc.)."""
        brief = self._load_brief(prep_folder)
        if not brief:
            return {
                "success": False,
                "error": f"No case_config.json found in {prep_folder}",
            }

        ctx = brief.context
        allowed_fields = {
            "company_name", "role_title", "target_function", "urls",
            "job_description", "interview_date", "interviewer_names",
            "interview_format", "background_points",
        }

        updated_fields = []
        for key, value in updates.items():
            if key in allowed_fields:
                setattr(ctx, key, value)
                updated_fields.append(key)

        # Save back
        config_path = Path(prep_folder) / CASE_CONFIG_FILENAME
        config_path.write_text(
            json.dumps(brief.to_dict(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

        return {
            "success": True,
            "updated_fields": updated_fields,
            "updated_context": ctx.to_dict(),
        }

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _load_context(self, prep_folder: str) -> Optional[PrepContext]:
        """Load PrepContext from case_config.json."""
        config_path = Path(prep_folder) / CASE_CONFIG_FILENAME
        if not config_path.exists():
            return None
        data = json.loads(config_path.read_text(encoding="utf-8"))
        return PrepContext.from_dict(data.get("context", {}))

    def _load_brief(self, prep_folder: str) -> Optional[InterviewBrief]:
        """Load full InterviewBrief from case_config.json."""
        config_path = Path(prep_folder) / CASE_CONFIG_FILENAME
        if not config_path.exists():
            return None
        data = json.loads(config_path.read_text(encoding="utf-8"))
        return InterviewBrief.from_dict(data)

    def _load_research_notes(self, prep_folder: str) -> Dict[str, str]:
        """Load research_notes.json from case folder."""
        notes_path = Path(prep_folder) / RESEARCH_NOTES_FILENAME
        if not notes_path.exists():
            return {}
        try:
            return json.loads(notes_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}

    def _update_section_status(
        self,
        prep_folder: str,
        section_id: str,
        status: str = "complete",
        file_path: str = "",
        word_count: int = 0,
    ) -> None:
        """Update a section's status in case_config.json."""
        brief = self._load_brief(prep_folder)
        if not brief:
            return

        for section in brief.sections:
            if section.section_id == section_id:
                section.status = status
                section.file_path = file_path
                section.word_count = word_count
                section.updated_at = datetime.now(timezone.utc).isoformat()
                break

        config_path = Path(prep_folder) / CASE_CONFIG_FILENAME
        config_path.write_text(
            json.dumps(brief.to_dict(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )


def get_skill(config: Optional[Dict] = None) -> InterviewPrepSkill:
    """Helper to get an initialized skill instance."""
    skill = InterviewPrepSkill(config)
    skill.initialize()
    return skill
