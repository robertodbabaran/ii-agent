from dataclasses import dataclass
from pathlib import Path
import os


@dataclass
class DailyMacroMetalsSettings:
    """Runtime configuration for the daily macro + metals newsletter."""

    base_dir: Path = Path("Memory/daily_macro_metals_newsletter")
    gmail_address: str = os.getenv("GMAIL_ADDRESS", "")
    gmail_app_password: str = os.getenv("GMAIL_APP_PASSWORD", "")
    default_recipient: str = os.getenv("RECIPIENT_EMAIL", "babaranrob@gmail.com")

    # Data sources requested by user
    campbell_substack_feed: str = "https://campbellramble.substack.com/feed"
    x_crossbordercap: str = "crossbordercap"
    x_abcampbell: str = "abcampbell"

    nitter_instances: tuple[str, ...] = (
        "https://nitter.net",
        "https://nitter.poast.org",
        "https://nitter.privacydev.net",
    )

    def ensure_dirs(self) -> None:
        (self.base_dir / "archive").mkdir(parents=True, exist_ok=True)
