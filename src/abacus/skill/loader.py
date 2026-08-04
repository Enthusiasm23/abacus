from pathlib import Path
from typing import Any

import yaml

from .parser import Skill


class SkillLoader:
    """SKILL.md 加载器"""

    def load_from_directory(self, skills_dir: Path) -> [Skill]:
        skills = []
        for skill_file in skills_dir.rglob("SKILL.md"):
            skills.append(self.parse_skill(skill_file))
        return skills

    def parse_skill(self, path: Path) -> Skill:
        content = path.read_text(encoding="utf-8")
        frontmatter = self._parse_frontmatter(content)

        return Skill(
            name=frontmatter.get("name", ""),
            description=frontmatter.get("description", ""),
            chapter=frontmatter.get("chapter", ""),
            level=frontmatter.get("level", "rod"),
            steps=[],
            content=content,
        )

    def _parse_frontmatter(self, content: str) -> dict[str, Any]:
        if content.startswith("---"):
            end = content.index("---", 3)
            yaml_content = content[3:end]
            return yaml.safe_load(yaml_content) or {}
        return {}
