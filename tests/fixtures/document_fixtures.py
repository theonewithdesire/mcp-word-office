"""
Document fixtures for testing various document types and scenarios.

This module provides sample documents and document-related test fixtures
for comprehensive testing of the Word MCP server.
"""

import tempfile
import os
from pathlib import Path
from typing import Dict, Any, List
from dataclasses import dataclass
import json


@dataclass
class DocumentFixture:
    """Represents a test document fixture."""
    name: str
    content: str
    file_type: str
    metadata: Dict[str, Any]
    expected_structure: Dict[str, Any]


class DocumentFixtureManager:
    """Manages test document fixtures."""
    
    def __init__(self):
        """Initialize the fixture manager."""
        self.temp_dir = None
        self.created_files = []
    
    def setup_temp_directory(self) -> Path:
        """Set up a temporary directory for test files."""
        if self.temp_dir is None:
            self.temp_dir = tempfile.mkdtemp()
        return Path(self.temp_dir)
    
    def cleanup(self):
        """Clean up created test files."""
        for file_path in self.created_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception:
                pass  # Ignore cleanup errors
        
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                os.rmdir(self.temp_dir)
            except Exception:
                pass
    
    def create_simple_document(self) -> DocumentFixture:
        """Create a simple text document fixture."""
        return DocumentFixture(
            name="simple_document",
            content="This is a simple test document with basic text content.",
            file_type="docx",
            metadata={
                "title": "Simple Test Document",
                "author": "Test Author",
                "word_count": 10,
                "paragraph_count": 1
            },
            expected_structure={
                "paragraphs": 1,
                "tables": 0,
                "images": 0,
                "headers": 0,
                "footers": 0
            }
        )
    
    def create_formatted_document(self) -> DocumentFixture:
        """Create a document with various formatting."""
        content = """Title: Formatted Test Document
        
This document contains various formatting elements:

• Bold text
• Italic text  
• Underlined text
• Different font sizes

This is a paragraph with mixed formatting including bold, italic, and normal text.

Another paragraph with different styling."""
        
        return DocumentFixture(
            name="formatted_document",
            content=content,
            file_type="docx",
            metadata={
                "title": "Formatted Test Document",
                "author": "Test Author",
                "word_count": 35,
                "paragraph_count": 6
            },
            expected_structure={
                "paragraphs": 6,
                "tables": 0,
                "images": 0,
                "headers": 0,
                "footers": 0,
                "formatting": {
                    "bold_ranges": 3,
                    "italic_ranges": 2,
                    "underline_ranges": 1
                }
            }
        )
    
    def create_table_document(self) -> DocumentFixture:
        """Create a document with tables."""
        content = """Document with Tables

Table 1: Sample Data
Name    | Age | City
--------|-----|--------
John    | 25  | NYC
Jane    | 30  | LA
Bob     | 35  | Chicago

Table 2: Statistics
Metric  | Value
--------|------
Total   | 100
Average | 33.3
Max     | 35"""
        
        return DocumentFixture(
            name="table_document",
            content=content,
            file_type="docx",
            metadata={
                "title": "Document with Tables",
                "author": "Test Author",
                "word_count": 45,
                "paragraph_count": 3,
                "table_count": 2
            },
            expected_structure={
                "paragraphs": 3,
                "tables": 2,
                "images": 0,
                "headers": 0,
                "footers": 0,
                "table_details": [
                    {"rows": 4, "columns": 3},
                    {"rows": 4, "columns": 2}
                ]
            }
        )
    
    def create_large_document(self) -> DocumentFixture:
        """Create a large document for performance testing."""
        # Generate large content
        paragraphs = []
        for i in range(100):
            paragraph = f"This is paragraph {i+1} of the large test document. " * 10
            paragraphs.append(paragraph)
        
        content = "\n\n".join(paragraphs)
        
        return DocumentFixture(
            name="large_document",
            content=content,
            file_type="docx",
            metadata={
                "title": "Large Test Document",
                "author": "Test Author",
                "word_count": 10000,
                "paragraph_count": 100
            },
            expected_structure={
                "paragraphs": 100,
                "tables": 0,
                "images": 0,
                "headers": 0,
                "footers": 0
            }
        )
    
    def create_complex_document(self) -> DocumentFixture:
        """Create a complex document with multiple elements."""
        content = """Complex Test Document

Table of Contents:
1. Introduction
2. Data Analysis
3. Conclusions

Introduction
============
This document demonstrates complex formatting and structure.

Data Analysis
=============

Sample Data Table:
Name     | Score | Grade
---------|-------|-------
Alice    | 95    | A
Bob      | 87    | B
Charlie  | 92    | A

Key Statistics:
• Average Score: 91.3
• Highest Score: 95
• Lowest Score: 87

Conclusions
===========
The analysis shows positive results across all metrics.

Footer: Document created for testing purposes."""
        
        return DocumentFixture(
            name="complex_document",
            content=content,
            file_type="docx",
            metadata={
                "title": "Complex Test Document",
                "author": "Test Author",
                "word_count": 85,
                "paragraph_count": 12,
                "table_count": 1,
                "list_count": 2
            },
            expected_structure={
                "paragraphs": 12,
                "tables": 1,
                "images": 0,
                "headers": 1,
                "footers": 1,
                "lists": 2,
                "headings": 3
            }
        )
    
    def create_document_file(self, fixture: DocumentFixture) -> str:
        """Create an actual file from a document fixture."""
        temp_dir = self.setup_temp_directory()
        file_path = temp_dir / f"{fixture.name}.{fixture.file_type}"
        
        # For testing purposes, create a simple text file
        # In a real implementation, you might use python-docx to create actual Word files
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(fixture.content)
        
        self.created_files.append(str(file_path))
        return str(file_path)
    
    def get_all_fixtures(self) -> List[DocumentFixture]:
        """Get all available document fixtures."""
        return [
            self.create_simple_document(),
            self.create_formatted_document(),
            self.create_table_document(),
            self.create_large_document(),
            self.create_complex_document()
        ]
    
    def create_test_scenario_files(self) -> Dict[str, str]:
        """Create files for all test scenarios."""
        scenario_files = {}
        
        for fixture in self.get_all_fixtures():
            file_path = self.create_document_file(fixture)
            scenario_files[fixture.name] = file_path
        
        return scenario_files


class PerformanceTestData:
    """Provides data for performance testing."""
    
    @staticmethod
    def generate_large_text(size_kb: int) -> str:
        """Generate large text content for performance testing."""
        # Generate approximately size_kb kilobytes of text
        words_per_kb = 150  # Approximate words per KB
        total_words = size_kb * words_per_kb
        
        base_words = [
            "performance", "testing", "document", "content", "large", "text",
            "sample", "data", "analysis", "results", "metrics", "evaluation",
            "benchmark", "measurement", "assessment", "validation", "verification"
        ]
        
        text_parts = []
        for i in range(0, total_words, len(base_words)):
            text_parts.extend(base_words)
        
        # Truncate to exact word count and join
        return " ".join(text_parts[:total_words])
    
    @staticmethod
    def generate_table_data(rows: int, columns: int) -> List[List[str]]:
        """Generate table data for testing."""
        headers = [f"Column {i+1}" for i in range(columns)]
        data = [headers]
        
        for row in range(rows):
            row_data = [f"Row {row+1} Col {col+1}" for col in range(columns)]
            data.append(row_data)
        
        return data
    
    @staticmethod
    def create_concurrent_test_scenarios() -> List[Dict[str, Any]]:
        """Create scenarios for concurrent operation testing."""
        return [
            {
                "name": "light_load",
                "concurrent_docs": 3,
                "operations_per_doc": 5,
                "text_size_kb": 1
            },
            {
                "name": "medium_load", 
                "concurrent_docs": 5,
                "operations_per_doc": 10,
                "text_size_kb": 5
            },
            {
                "name": "heavy_load",
                "concurrent_docs": 10,
                "operations_per_doc": 20,
                "text_size_kb": 10
            }
        ]


# Global fixture manager instance
fixture_manager = DocumentFixtureManager()


def cleanup_fixtures():
    """Clean up all created fixtures."""
    fixture_manager.cleanup()


# Pytest fixtures
import pytest

@pytest.fixture(scope="session")
def document_fixtures():
    """Pytest fixture providing document fixtures."""
    return fixture_manager

@pytest.fixture(scope="session", autouse=True)
def cleanup_after_tests():
    """Automatically clean up fixtures after tests."""
    yield
    cleanup_fixtures()

@pytest.fixture
def simple_document():
    """Fixture for a simple document."""
    return fixture_manager.create_simple_document()

@pytest.fixture
def formatted_document():
    """Fixture for a formatted document."""
    return fixture_manager.create_formatted_document()

@pytest.fixture
def table_document():
    """Fixture for a document with tables."""
    return fixture_manager.create_table_document()

@pytest.fixture
def large_document():
    """Fixture for a large document."""
    return fixture_manager.create_large_document()

@pytest.fixture
def complex_document():
    """Fixture for a complex document."""
    return fixture_manager.create_complex_document()

@pytest.fixture
def test_document_files(document_fixtures):
    """Fixture providing actual test document files."""
    return document_fixtures.create_test_scenario_files()