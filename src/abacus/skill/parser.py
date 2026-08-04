from dataclasses import dataclass
from typing import Any


@dataclass
class SkillStep:
    """Skill 步骤"""

    capability_name: str
    params: dict[str, Any]


@dataclass
class Skill:
    """Skill 数据类"""

    name: str
    description: str
    chapter: str
    level: str
    steps: list[SkillStep]
    content: str
