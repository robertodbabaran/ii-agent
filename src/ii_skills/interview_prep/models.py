"""
Data models for the Interview Prep skill.
"""

from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional


@dataclass
class PrepContext:
    """Core context for an interview preparation session."""

    company_name: str
    role_title: str
    target_function: str  # e.g., "Investor Relations", "Corporate Development"
    urls: List[str] = field(default_factory=list)
    job_description: str = ""
    interview_date: str = ""  # Optional target date
    interviewer_names: List[str] = field(default_factory=list)
    interview_format: str = ""  # Phone, video, in-person, case study
    background_points: List[str] = field(default_factory=list)  # User's key experiences
    case_folder: str = ""  # Absolute path to case workspace
    created_at: str = field(
        default_factory=lambda: datetime.now(timezone.utc).isoformat()
    )

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PrepContext":
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


@dataclass
class SectionStatus:
    """Tracks the status of a single prep section."""

    section_id: str  # T01-T09
    section_name: str
    status: str = "pending"  # pending | in_progress | complete
    file_path: str = ""
    word_count: int = 0
    updated_at: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "SectionStatus":
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


@dataclass
class InterviewBrief:
    """Full interview preparation state."""

    context: PrepContext
    sections: List[SectionStatus] = field(default_factory=list)
    research_notes: Dict[str, str] = field(default_factory=dict)  # URL -> content

    def to_dict(self) -> Dict[str, Any]:
        return {
            "context": self.context.to_dict(),
            "sections": [s.to_dict() for s in self.sections],
            "research_notes": self.research_notes,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "InterviewBrief":
        context = PrepContext.from_dict(data.get("context", {}))
        sections = [
            SectionStatus.from_dict(s) for s in data.get("sections", [])
        ]
        return cls(
            context=context,
            sections=sections,
            research_notes=data.get("research_notes", {}),
        )

    @property
    def completion_pct(self) -> float:
        if not self.sections:
            return 0.0
        complete = sum(1 for s in self.sections if s.status == "complete")
        return round(complete / len(self.sections) * 100, 1)
