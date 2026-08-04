"""SkillIndexer 单元测试"""

import json
import pytest
from pathlib import Path

from abacus.skill.indexer import SkillIndexer


@pytest.fixture
def tmp_indexer(tmp_path):
    """创建临时 SkillIndexer 实例"""
    db_path = tmp_path / "test_index.db"
    indexer = SkillIndexer(db_path)
    yield indexer
    indexer.close()


@pytest.fixture
def sample_skill_dir(tmp_path):
    """创建包含 SKILL.md 和知识文件的临时目录"""
    skills_dir = tmp_path / "skills"

    # 创建 skill 文件
    skill_dir = skills_dir / "field"
    skill_dir.mkdir(parents=True)
    skill_file = skill_dir / "SKILL.md"
    skill_file.write_text("""---
name: abacus-field
chapter: field
description: 方田章 - 数据读取
level: rod
---

# 方田章

使用 `measure_range` 和 `measure_cells` 读取数据。
""", encoding="utf-8")

    # 创建知识文件
    shared_dir = skills_dir / "shared"
    shared_dir.mkdir(parents=True)
    kb_file = shared_dir / "formula-best-practices.md"
    kb_file.write_text("""# 公式最佳实践

## 基础规则
使用绝对引用。
""", encoding="utf-8")

    return skills_dir


class TestSkillIndexerInit:
    def test_init_creates_db(self, tmp_path):
        db_path = tmp_path / "test.db"
        indexer = SkillIndexer(db_path)
        assert db_path.exists()
        indexer.close()

    def test_init_creates_schema(self, tmp_path):
        db_path = tmp_path / "test.db"
        indexer = SkillIndexer(db_path)
        cursor = indexer.conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = {r["name"] for r in cursor.fetchall()}
        assert "skills" in tables
        assert "knowledge" in tables
        assert "relations" in tables
        indexer.close()

    def test_init_creates_fts_tables(self, tmp_path):
        db_path = tmp_path / "test.db"
        indexer = SkillIndexer(db_path)
        cursor = indexer.conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = {r["name"] for r in cursor.fetchall()}
        assert "skills_fts" in tables
        assert "knowledge_fts" in tables
        indexer.close()


class TestSkillIndexerHelpers:
    def test_hash_content(self, tmp_indexer):
        h1 = tmp_indexer._hash_content("hello")
        h2 = tmp_indexer._hash_content("hello")
        h3 = tmp_indexer._hash_content("world")
        assert h1 == h2
        assert h1 != h3
        assert len(h1) == 64  # SHA256 hex digest length

    def test_parse_skill_frontmatter_with_yaml(self, tmp_indexer):
        content = """---
name: test_skill
chapter: field
description: Test description
level: rod
---
# Content here
"""
        result = tmp_indexer._parse_skill_frontmatter(content)
        assert result["name"] == "test_skill"
        assert result["chapter"] == "field"
        assert result["level"] == "rod"

    def test_parse_skill_frontmatter_without_yaml(self, tmp_indexer):
        content = "# No frontmatter here"
        result = tmp_indexer._parse_skill_frontmatter(content)
        assert result == {}

    def test_extract_capabilities(self, tmp_indexer):
        content = "Use `measure_range` and `convert_format` for data."
        caps = tmp_indexer._extract_capabilities_from_content(content)
        assert "measure_range" in caps
        assert "convert_format" in caps

    def test_extract_capabilities_ignores_unknown(self, tmp_indexer):
        content = "Use `random_text` and `another` here."
        caps = tmp_indexer._extract_capabilities_from_content(content)
        assert "random_text" not in caps

    def test_extract_knowledge_sections(self, tmp_indexer):
        content = """# Title
## Section A
### Sub A1
## Section B
"""
        sections = tmp_indexer._extract_knowledge_sections(content)
        assert len(sections) == 4
        assert sections[0] == {"title": "Title", "level": 1}
        assert sections[1] == {"title": "Section A", "level": 2}

    def test_detect_knowledge_tags(self, tmp_indexer):
        assert "anti-pattern" in tmp_indexer._detect_knowledge_tags("anti-patterns", "content")
        assert "best-practice" in tmp_indexer._detect_knowledge_tags("best-practices", "content")
        assert "reference" in tmp_indexer._detect_knowledge_tags("reference", "content")
        assert "formula" in tmp_indexer._detect_knowledge_tags("formula", "content")
        assert "chart" in tmp_indexer._detect_knowledge_tags("chart", "content")
        assert "style" in tmp_indexer._detect_knowledge_tags("style", "content")

    def test_detect_knowledge_tags_default(self, tmp_indexer):
        tags = tmp_indexer._detect_knowledge_tags("unknown-file", "just some text")
        assert tags == ["general"]


class TestSkillIndexerScan:
    def test_scan_populates_skills(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        cursor = tmp_indexer.conn.cursor()
        cursor.execute("SELECT COUNT(*) as cnt FROM skills")
        assert cursor.fetchone()["cnt"] >= 1

    def test_scan_populates_knowledge(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        cursor = tmp_indexer.conn.cursor()
        cursor.execute("SELECT COUNT(*) as cnt FROM knowledge")
        assert cursor.fetchone()["cnt"] >= 1

    def test_scan_populates_fts(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        # FTS5 content-sync tables don't support COUNT(*) directly
        # Verify via search instead
        results = tmp_indexer.search("field")
        assert len(results) >= 1


class TestSkillIndexerSearch:
    def test_search_returns_results(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        results = tmp_indexer.search("formula")
        assert len(results) >= 1

    def test_search_skill_result(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        # FTS5 tokenizes "abacus-field" as two tokens; use single token
        results = tmp_indexer.search("field")
        assert any(r["type"] == "skill" and r["name"] == "abacus-field" for r in results)

    def test_search_limit(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        results = tmp_indexer.search("formula", limit=1)
        assert len(results) <= 1

    def test_search_no_results(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        results = tmp_indexer.search("zzzznonexistent")
        assert len(results) == 0


class TestSkillIndexerGraph:
    def test_graph_returns_skill(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        result = tmp_indexer.graph("abacus-field")
        assert result["skill"] == "abacus-field"
        assert result["chapter"] == "field"

    def test_graph_capabilities(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        result = tmp_indexer.graph("abacus-field")
        assert "measure_range" in result["capabilities"]
        assert "measure_cells" in result["capabilities"]

    def test_graph_unknown_skill(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        result = tmp_indexer.graph("nonexistent-skill")
        assert "error" in result


class TestSkillIndexerStats:
    def test_stats_after_scan(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        stats = tmp_indexer.stats()
        assert stats["skills"] >= 1
        assert stats["knowledge_files"] >= 1
        assert "chapters" in stats
        assert "sources" in stats

    def test_stats_empty_db(self, tmp_indexer):
        stats = tmp_indexer.stats()
        assert stats["skills"] == 0
        assert stats["knowledge_files"] == 0


class TestSkillIndexerExport:
    def test_export_mcp(self, tmp_indexer, sample_skill_dir):
        tmp_indexer.scan(sample_skill_dir)
        export = tmp_indexer.export_mcp()
        assert "skills" in export
        assert "knowledge" in export
        assert "relations" in export
        assert "stats" in export
        assert len(export["skills"]) >= 1
