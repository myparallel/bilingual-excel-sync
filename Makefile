.PHONY: help install test lint format clean build publish

help:
	@echo "Available commands:"
	@echo "  install     Install dependencies"
	@echo "  test        Run tests"
	@echo "  lint        Run linter"
	@echo "  format      Format code"
	@echo "  clean       Clean build artifacts"
	@echo "  build       Build package"
	@echo "  publish     Publish to PyPI"

install:
	pip install -r requirements.txt
	pip install -e .

test:
	python -m pytest tests/ -v --cov=bilingual_excel_sync --cov-report=html

lint:
	python -m flake8 bilingual_excel_sync tests
	python -m pylint bilingual_excel_sync
	python -m mypy bilingual_excel_sync

format:
	python -m black bilingual_excel_sync tests
	python -m isort bilingual_excel_sync tests

clean:
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info
	rm -rf __pycache__
	rm -rf .pytest_cache
	rm -rf .coverage
	rm -rf htmlcov/
	find . -type f -name "*.pyc" -delete
	find . -type f -name "*.pyo" -delete
	find . -type d -name "__pycache__" -delete

build: clean
	python -m build

publish: build
	python -m twine upload dist/*

demo:
	python demo.py

test-skill:
	python test_skill.py

test-cases:
	python test_cases.py

examples:
	python examples.py

analyze:
	python scripts/bilingual_excel_sync.py analyze $(input) $(output)

convert:
	python scripts/bilingual_excel_sync.py convert $(input) $(output)

verify:
	python scripts/bilingual_excel_sync.py verify $(input) $(output)