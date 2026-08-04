"""SKILL.md 知识库索引器 - SQLite 图谱 + FTS5 全文搜索"""

import hashlib
import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml


class SkillIndexer:
    """SKILL.md 知识库索引器"""

    def __init__(self, db_path: Path | None = None):
        self.db_path = db_path or Path("skills_index.db")
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self):
        """初始化 SQLite schema"""
        cursor = self.conn.cursor()

        cursor.executescript("""
            CREATE TABLE IF NOT EXISTS skills (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                chapter TEXT NOT NULL,
                description TEXT,
                level TEXT DEFAULT 'rod',
                file_path TEXT NOT NULL,
                content_hash TEXT,
                capabilities TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS knowledge (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                source TEXT,
                file_path TEXT NOT NULL,
                content_hash TEXT,
                sections TEXT,
                tags TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS capabilities (
                name TEXT PRIMARY KEY,
                chapter TEXT NOT NULL,
                description TEXT,
                file_path TEXT
            );

            CREATE TABLE IF NOT EXISTS relations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_type TEXT NOT NULL,
                source_id TEXT NOT NULL,
                target_type TEXT NOT NULL,
                target_id TEXT NOT NULL,
                relation_type TEXT NOT NULL,
                confidence REAL DEFAULT 1.0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """)

        try:
            cursor.executescript("""
                CREATE VIRTUAL TABLE IF NOT EXISTS skills_fts USING fts5(
                    name, description, content, content=skills, content_rowid=id
                );

                CREATE VIRTUAL TABLE IF NOT EXISTS knowledge_fts USING fts5(
                    name, content, content=knowledge, content_rowid=id
                );
            """)
        except sqlite3.OperationalError:
            pass

        self.conn.commit()

    def _hash_content(self, content: str) -> str:
        """计算内容 SHA256 哈希"""
        return hashlib.sha256(content.encode("utf-8")).hexdigest()

    def _parse_skill_frontmatter(self, content: str) -> dict[str, Any]:
        """解析 SKILL.md 的 YAML frontmatter"""
        if content.startswith("---"):
            end = content.index("---", 3)
            yaml_content = content[3:end]
            return yaml.safe_load(yaml_content) or {}
        return {}

    def _extract_capabilities_from_content(self, content: str) -> [str]:
        """从 SKILL.md 内容中提取能力名称"""
        caps = []
        for match in re.finditer(r"`(\w+)`", content):
            cap = match.group(1)
            if cap.startswith(
                (
                    "abacus",
                    "measure",
                    "convert",
                    "group",
                    "distribute",
                    "summarize",
                    "find",
                    "derive",
                    "calculate",
                    "batch",
                    "create",
                    "format",
                    "update",
                    "list",
                    "delete",
                    "import",
                    "export",
                    "migrate",
                    "validate",
                    "analyze",
                    "visualize",
                    "merge",
                    "split",
                    "clean",
                    "transpose",
                )
            ):
                caps.append(cap)
        return list(set(caps))

    def _extract_knowledge_sections(self, content: str) -> list[dict[str, Any]]:
        """提取知识文件的章节结构"""
        sections = []
        for match in re.finditer(r"^(#{1,3})\s+(.+)$", content, re.MULTILINE):
            level = len(match.group(1))
            title = match.group(2).strip()
            sections.append({"title": title, "level": level})
        return sections

    def _detect_knowledge_tags(self, name: str, content: str) -> list[str]:
        """检测知识文件的标签"""
        tags = []
        name_lower = name.lower()
        content_lower = content[:500].lower()

        if "anti-pattern" in name_lower or "反模式" in content_lower:
            tags.append("anti-pattern")
        if "best-practice" in name_lower or "最佳实践" in content_lower:
            tags.append("best-practice")
        if "reference" in name_lower or "参考" in content_lower:
            tags.append("reference")
        if "pitfall" in name_lower or "陷阱" in content_lower:
            tags.append("pitfall")
        if "formula" in name_lower or "公式" in content_lower:
            tags.append("formula")
        if "chart" in name_lower or "图表" in content_lower:
            tags.append("chart")
        if "style" in name_lower or "样式" in content_lower:
            tags.append("style")

        return tags if tags else ["general"]

    def scan(self, skills_dir: Path):
        """增量扫描 skills/ 目录"""
        skills_dir = Path(skills_dir)

        for skill_file in skills_dir.rglob("SKILL.md"):
            self._index_skill(skill_file)

        knowledge_dir = skills_dir / "shared"
        if knowledge_dir.exists():
            for md_file in knowledge_dir.glob("*.md"):
                self._index_knowledge(md_file)

        self._build_relations(skills_dir)
        self.conn.commit()

    def _index_skill(self, file_path: Path):
        """索引单个 SKILL.md"""
        content = file_path.read_text(encoding="utf-8")
        content_hash = self._hash_content(content)
        frontmatter = self._parse_skill_frontmatter(content)
        capabilities = self._extract_capabilities_from_content(content)

        name = frontmatter.get("name", file_path.parent.name)
        chapter = frontmatter.get("chapter", file_path.parent.name)
        description = frontmatter.get("description", "")
        level = frontmatter.get("level", "rod")

        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT OR REPLACE INTO skills (name, chapter, description, level, file_path, content_hash, capabilities, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                name,
                chapter,
                description,
                level,
                str(file_path),
                content_hash,
                json.dumps(capabilities),
                datetime.now().isoformat(),
            ),
        )

        skill_id = cursor.lastrowid
        try:
            cursor.execute(
                """
                INSERT OR REPLACE INTO skills_fts (rowid, name, description, content)
                VALUES (?, ?, ?, ?)
            """,
                (skill_id, name, description, content),
            )
        except sqlite3.OperationalError:
            pass

    def _index_knowledge(self, file_path: Path):
        """索引单个知识文件"""
        content = file_path.read_text(encoding="utf-8")
        content_hash = self._hash_content(content)
        name = file_path.stem
        sections = self._extract_knowledge_sections(content)
        tags = self._detect_knowledge_tags(name, content)

        source = "unknown"
        if any(
            x in name
            for x in [
                "sbroenne",
                "behavioral",
                "chart",
                "pivot",
                "table",
                "range",
                "worksheet",
                "conditional",
                "dashboard",
                "data-model",
                "dmv",
                "gotchas",
                "m-code",
                "power-query",
                "slicer",
                "workflows",
            ]
        ):
            source = "sbroenne"
        elif any(x in name for x in ["formula", "excel-pitfalls", "common-pitfalls"]):
            source = "gaaiyun"
        elif any(x in name for x in ["financial", "trigger", "code-style"]):
            source = "anthropic"
        elif any(x in name for x in ["chinese", "scenario"]):
            source = "wps"
        elif name in ["best-practices", "critical-rules", "formula-diagnosis"]:
            source = "custom"

        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT OR REPLACE INTO knowledge (name, source, file_path, content_hash, sections, tags, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
            (
                name,
                source,
                str(file_path),
                content_hash,
                json.dumps(sections, ensure_ascii=False),
                json.dumps(tags),
                datetime.now().isoformat(),
            ),
        )

        knowledge_id = cursor.lastrowid
        try:
            cursor.execute(
                """
                INSERT OR REPLACE INTO knowledge_fts (rowid, name, content)
                VALUES (?, ?, ?)
            """,
                (knowledge_id, name, content),
            )
        except sqlite3.OperationalError:
            pass

    def _build_relations(self, skills_dir: Path):
        """构建关系图谱"""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM relations")

        cursor.execute("SELECT name, chapter, capabilities FROM skills")
        skills = cursor.fetchall()

        for skill in skills:
            skill_name = skill["name"]
            caps = json.loads(skill["capabilities"] or "[]")

            for cap in caps:
                cursor.execute(
                    """
                    INSERT OR IGNORE INTO relations (source_type, source_id, target_type, target_id, relation_type)
                    VALUES ('skill', ?, 'capability', ?, 'uses')
                """,
                    (skill_name, cap),
                )

        cursor.execute("SELECT name, tags FROM knowledge")
        knowledge_files = cursor.fetchall()

        for kb in knowledge_files:
            kb_name = kb["name"]
            tags = json.loads(kb["tags"] or "[]")

            for skill in skills:
                skill_name = skill["name"]
                skill_desc = skill["chapter"]

                if any(tag in skill_desc for tag in tags):
                    cursor.execute(
                        """
                        INSERT OR IGNORE INTO relations (source_type, source_id, target_type, target_id, relation_type)
                        VALUES ('knowledge', ?, 'skill', ?, 'related')
                    """,
                        (kb_name, skill_name),
                    )

    def search(self, query: str, limit: int = 10) -> list[dict[str, Any]]:
        """FTS5 全文搜索"""
        cursor = self.conn.cursor()
        results = []

        try:
            cursor.execute(
                """
                SELECT s.name, s.chapter, s.description, s.file_path,
                       rank
                FROM skills_fts
                JOIN skills s ON skills_fts.rowid = s.id
                WHERE skills_fts MATCH ?
                ORDER BY rank
                LIMIT ?
            """,
                (query, limit),
            )

            for row in cursor.fetchall():
                results.append(
                    {
                        "type": "skill",
                        "name": row["name"],
                        "chapter": row["chapter"],
                        "description": row["description"],
                        "file_path": row["file_path"],
                        "relevance": abs(row["rank"]),
                    }
                )
        except sqlite3.OperationalError:
            pass

        try:
            cursor.execute(
                """
                SELECT k.name, k.source, k.file_path, k.tags,
                       rank
                FROM knowledge_fts
                JOIN knowledge k ON knowledge_fts.rowid = k.id
                WHERE knowledge_fts MATCH ?
                ORDER BY rank
                LIMIT ?
            """,
                (query, limit),
            )

            for row in cursor.fetchall():
                results.append(
                    {
                        "type": "knowledge",
                        "name": row["name"],
                        "source": row["source"],
                        "file_path": row["file_path"],
                        "tags": json.loads(row["tags"] or "[]"),
                        "relevance": abs(row["rank"]),
                    }
                )
        except sqlite3.OperationalError:
            pass

        results.sort(key=lambda x: x.get("relevance", 0))
        return results[:limit]

    def graph(self, skill_name: str) -> dict[str, Any]:
        """获取 skill 关联图谱"""
        cursor = self.conn.cursor()

        cursor.execute("SELECT * FROM skills WHERE name = ?", (skill_name,))
        skill = cursor.fetchone()
        if not skill:
            return {"error": f"Skill '{skill_name}' not found"}

        caps = json.loads(skill["capabilities"] or "[]")

        cursor.execute(
            """
            SELECT target_id, relation_type FROM relations
            WHERE source_type = 'skill' AND source_id = ?
        """,
            (skill_name,),
        )
        uses = [
            {"capability": r["target_id"], "relation": r["relation_type"]}
            for r in cursor.fetchall()
        ]

        cursor.execute(
            """
            SELECT source_id, relation_type FROM relations
            WHERE target_type = 'skill' AND target_id = ? AND source_type = 'knowledge'
        """,
            (skill_name,),
        )
        referenced_by = [
            {"knowledge": r["source_id"], "relation": r["relation_type"]} for r in cursor.fetchall()
        ]

        cursor.execute(
            """
            SELECT DISTINCT target_id FROM relations
            WHERE source_type = 'skill' AND source_id = ?
            AND target_type = 'capability'
        """,
            (skill_name,),
        )
        cap_names = [r["target_id"] for r in cursor.fetchall()]

        related_skills = set()
        for cap in cap_names:
            cursor.execute(
                """
                SELECT source_id FROM relations
                WHERE target_type = 'capability' AND target_id = ?
                AND source_type = 'skill' AND source_id != ?
            """,
                (cap, skill_name),
            )
            for r in cursor.fetchall():
                related_skills.add(r["source_id"])

        return {
            "skill": skill_name,
            "chapter": skill["chapter"],
            "description": skill["description"],
            "capabilities": caps,
            "uses": uses,
            "referenced_by": referenced_by,
            "related_skills": list(related_skills),
        }

    def stats(self) -> dict[str, Any]:
        """获取索引统计"""
        cursor = self.conn.cursor()

        cursor.execute("SELECT COUNT(*) as cnt FROM skills")
        skill_count = cursor.fetchone()["cnt"]

        cursor.execute("SELECT COUNT(*) as cnt FROM knowledge")
        knowledge_count = cursor.fetchone()["cnt"]

        cursor.execute("SELECT COUNT(*) as cnt FROM capabilities")
        cap_count = cursor.fetchone()["cnt"]

        cursor.execute("SELECT COUNT(*) as cnt FROM relations")
        relation_count = cursor.fetchone()["cnt"]

        cursor.execute("SELECT chapter, COUNT(*) as cnt FROM skills GROUP BY chapter")
        chapters = {r["chapter"]: r["cnt"] for r in cursor.fetchall()}

        cursor.execute("SELECT source, COUNT(*) as cnt FROM knowledge GROUP BY source")
        sources = {r["source"]: r["cnt"] for r in cursor.fetchall()}

        return {
            "skills": skill_count,
            "knowledge_files": knowledge_count,
            "capabilities": cap_count,
            "relations": relation_count,
            "chapters": chapters,
            "sources": sources,
        }

    def export_mcp(self) -> dict[str, Any]:
        """导出为 MCP Server 可用格式"""
        cursor = self.conn.cursor()

        cursor.execute("SELECT name, chapter, description, capabilities FROM skills")
        skills = []
        for row in cursor.fetchall():
            skills.append(
                {
                    "name": row["name"],
                    "chapter": row["chapter"],
                    "description": row["description"],
                    "capabilities": json.loads(row["capabilities"] or "[]"),
                }
            )

        cursor.execute("SELECT name, source, tags FROM knowledge")
        knowledge = []
        for row in cursor.fetchall():
            knowledge.append(
                {
                    "name": row["name"],
                    "source": row["source"],
                    "tags": json.loads(row["tags"] or "[]"),
                }
            )

        cursor.execute(
            "SELECT source_type, source_id, target_type, target_id, relation_type FROM relations"
        )
        relations = []
        for row in cursor.fetchall():
            relations.append(
                {
                    "source": {"type": row["source_type"], "id": row["source_id"]},
                    "target": {"type": row["target_type"], "id": row["target_id"]},
                    "relation": row["relation_type"],
                }
            )

        return {
            "skills": skills,
            "knowledge": knowledge,
            "relations": relations,
            "stats": self.stats(),
        }

    def close(self):
        """关闭数据库连接"""
        self.conn.close()
