from abacus.skill.parser import Skill, SkillStep


class TestSkill:
    def test_skill_creation(self):
        skill = Skill(
            name="test",
            description="Test skill",
            chapter="field",
            level="rod",
            steps=[],
            content="# Test"
        )
        assert skill.name == "test"
        assert skill.level == "rod"
    
    def test_skill_step_creation(self):
        step = SkillStep(
            capability_name="measure_range",
            params={"file": "test.xlsx"}
        )
        assert step.capability_name == "measure_range"