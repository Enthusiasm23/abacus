from abacus.skill.loader import SkillLoader


class TestSkillLoader:
    def test_parse_frontmatter(self):
        loader = SkillLoader()
        content = """---
name: test_skill
description: Test skill
chapter: field
level: rod
---

# Test Skill
"""
        result = loader._parse_frontmatter(content)
        assert result["name"] == "test_skill"
        assert result["chapter"] == "field"