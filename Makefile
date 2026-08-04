.PHONY: install test lint format

install:
	pip install -e ".[dev]"

test:
	pytest tests/ -v --cov=abacus --cov-report=term-missing

lint:
	ruff check src/ tests/

format:
	ruff format src/ tests/
