"""
Tests for DocumentManager class.

This module tests document reading and content extraction functionality.
"""

import os
import pytest
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path

from word_mcp_server.word.document_manager import DocumentManager
from word_mcp_server.utils.errors import DocumentError


class TestDocumentManager:
    """Test cases for DocumentManager class."""
    
    @pytest.fixture
    def config(self):
        """Mock configuration for testing."""
        config = Mock()
        config.word = Mock()
        return config
    
    @pytest.fixture
    def document_manager(self, config):
        """Create DocumentManager instance for testing."""
        return DocumentManager(config)
    
    @pytest.fixture
    def mock_docx_document(self):
        """Create mock python-docx Document object."""
        doc = Mock()
        
        # Mock paragraphs
        para1 = Mock()
        para1.text = "This is the first paragraph."
        para1.style.name = "Normal"
        para1.alignment = None
        para1.runs = [Mock(text="This is the first paragraph.", bold=False, italic=False, underline=False, font=Mock(name="Arial", size=Mock(pt=12)))]
        
        para2 = Mock()
        para2.text = "Heading 1"
        para2.style.name = "Heading 1"
        para2.alignment = None
        para2.runs = [Mock(text="Heading 1", bold=True, italic=False, underline=False, font=Mock(name="Arial", size=Mock(pt=16)))]
        
        para3 = Mock()
        para3.text = "This is under heading 1."
        para3.style.name = "Normal"
        para3.alignment = None
        para3.runs = [Mock(text="This is under heading 1.", bold=False, italic=False, underline=False, font=Mock(name="Arial", size=Mock(pt=12)))]
        
        doc.paragraphs = [para1, para2, para3]
        
        # Mock tables
        table = Mock()
        row1 = Mock()
        row2 = Mock()
        
        cell1 = Mock()
        cell1.text = "Header 1"
        cell2 = Mock()
        cell2.text = "Header 2"
        cell3 = Mock()
        cell3.text = "Data 1"
        cell4 = Mock()
        cell4.text = "Data 2"
        
        row1.cells = [cell1, cell2]
        row2.cells = [cell3, cell4]
        table.rows = [row1, row2]
        table.columns = [Mock(), Mock()]
        
        doc.tables = [table]
        
        # Mock core properties
        doc.core_properties = Mock()
        doc.core_properties.title = "Test Document"
        doc.core_properties.author = "Test Author"
        doc.core_properties.subject = "Test Subject"
        doc.core_properties.created = Mock()
        doc.core_properties.created.isoformat.return_value = "2023-01-01T00:00:00"
        doc.core_properties.modified = Mock()
        doc.core_properties.modified.isoformat.return_value = "2023-01-02T00:00:00"
        
        return doc
    
    def test_init(self, config):
        """Test DocumentManager initialization."""
        manager = DocumentManager(config)
        assert manager.config == config
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    @patch('word_mcp_server.word.document_manager.Document')
    @patch('word_mcp_server.word.document_manager.Path')
    def test_read_document_success(self, mock_path, mock_document, mock_exists, mock_validate, document_manager, mock_docx_document):
        """Test successful document reading."""
        # Setup mocks
        mock_validate.return_value = True
        mock_exists.return_value = True
        mock_document.return_value = mock_docx_document
        
        # Mock Path object for file extension check
        mock_file_path = Mock()
        mock_file_path.suffix.lower.return_value = '.docx'
        mock_file_path.name = "test.docx"
        mock_file_path.stat.return_value = Mock(st_size=1024, st_ctime=1640995200, st_mtime=1641081600)
        mock_path.return_value = mock_file_path
        
        result = document_manager.read_document("test.docx")
        
        # Verify result structure
        assert result["path"] == "test.docx"
        assert "text" in result
        assert "structure" in result
        assert "tables" in result
        assert "paragraphs" in result
        assert "metadata" in result
        
        # Verify text extraction
        assert "This is the first paragraph." in result["text"]
        assert "Heading 1" in result["text"]
        assert "This is under heading 1." in result["text"]
        assert "Header 1" in result["text"]
        assert "Data 1" in result["text"]
        
        # Verify structure analysis
        structure = result["structure"]
        assert structure["paragraph_count"] == 2  # Only non-heading paragraphs
        assert structure["table_count"] == 1
        assert len(structure["headings"]) == 1
        assert structure["headings"][0]["text"] == "Heading 1"
        assert structure["headings"][0]["level"] == 1
        
        # Verify metadata
        metadata = result["metadata"]
        assert metadata["filename"] == "test.docx"
        assert metadata["title"] == "Test Document"
        assert metadata["author"] == "Test Author"
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    def test_read_document_invalid_path(self, mock_validate, document_manager):
        """Test reading document with invalid path."""
        mock_validate.return_value = False
        
        with pytest.raises(DocumentError, match="Invalid file path"):
            document_manager.read_document("invalid/path")
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    def test_read_document_file_not_found(self, mock_exists, mock_validate, document_manager):
        """Test reading non-existent document."""
        mock_validate.return_value = True
        mock_exists.return_value = False
        
        with pytest.raises(DocumentError, match="Document not found"):
            document_manager.read_document("nonexistent.docx")
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    @patch('word_mcp_server.word.document_manager.Path')
    def test_read_document_unsupported_format(self, mock_path, mock_exists, mock_validate, document_manager):
        """Test reading document with unsupported format."""
        mock_validate.return_value = True
        mock_exists.return_value = True
        
        mock_file_path = Mock()
        mock_file_path.suffix.lower.return_value = '.txt'
        mock_path.return_value = mock_file_path
        
        with pytest.raises(DocumentError, match="Unsupported file format"):
            document_manager.read_document("test.txt")
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    @patch('word_mcp_server.word.document_manager.Document')
    @patch('word_mcp_server.word.document_manager.Path')
    def test_read_document_load_failure(self, mock_path, mock_document, mock_exists, mock_validate, document_manager):
        """Test document loading failure."""
        mock_validate.return_value = True
        mock_exists.return_value = True
        mock_document.side_effect = Exception("Failed to load document")
        
        mock_file_path = Mock()
        mock_file_path.suffix.lower.return_value = '.docx'
        mock_path.return_value = mock_file_path
        
        with pytest.raises(DocumentError, match="Failed to load document"):
            document_manager.read_document("test.docx")
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    @patch('word_mcp_server.word.document_manager.Document')
    def test_extract_text_success(self, mock_document, mock_exists, mock_validate, document_manager, mock_docx_document):
        """Test successful text extraction."""
        mock_validate.return_value = True
        mock_exists.return_value = True
        mock_document.return_value = mock_docx_document
        
        result = document_manager.extract_text("test.docx")
        
        assert "This is the first paragraph." in result
        assert "Heading 1" in result
        assert "This is under heading 1." in result
        assert "Header 1" in result
        assert "Data 1" in result
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    @patch('word_mcp_server.word.document_manager.Document')
    def test_get_document_structure_success(self, mock_document, mock_exists, mock_validate, document_manager, mock_docx_document):
        """Test successful document structure analysis."""
        mock_validate.return_value = True
        mock_exists.return_value = True
        mock_document.return_value = mock_docx_document
        
        result = document_manager.get_document_structure("test.docx")
        
        assert result["paragraph_count"] == 2
        assert result["table_count"] == 1
        assert len(result["headings"]) == 1
        assert result["headings"][0]["text"] == "Heading 1"
        assert result["headings"][0]["level"] == 1
        assert len(result["sections"]) == 2  # Document start + heading section
    
    def test_extract_text_from_doc(self, document_manager, mock_docx_document):
        """Test internal text extraction method."""
        result = document_manager._extract_text_from_doc(mock_docx_document)
        
        lines = result.split('\n')
        assert "This is the first paragraph." in lines
        assert "Heading 1" in lines
        assert "This is under heading 1." in lines
        assert "Header 1" in lines
        assert "Data 1" in lines
    
    def test_analyze_document_structure(self, document_manager, mock_docx_document):
        """Test internal document structure analysis method."""
        result = document_manager._analyze_document_structure(mock_docx_document)
        
        assert result["paragraph_count"] == 2
        assert result["table_count"] == 1
        assert len(result["headings"]) == 1
        assert result["headings"][0]["text"] == "Heading 1"
        assert result["headings"][0]["level"] == 1
        
        # Check sections
        assert len(result["sections"]) == 2
        assert result["sections"][0]["title"] == "Document Start"
        assert result["sections"][1]["title"] == "Heading 1"
        assert result["sections"][1]["level"] == 1
    
    def test_extract_paragraphs(self, document_manager, mock_docx_document):
        """Test internal paragraph extraction method."""
        result = document_manager._extract_paragraphs(mock_docx_document)
        
        assert len(result) == 3
        assert result[0]["text"] == "This is the first paragraph."
        assert result[0]["style"] == "Normal"
        assert result[1]["text"] == "Heading 1"
        assert result[1]["style"] == "Heading 1"
        assert result[2]["text"] == "This is under heading 1."
        assert result[2]["style"] == "Normal"
        
        # Check run information
        assert len(result[0]["runs"]) == 1
        assert result[0]["runs"][0]["text"] == "This is the first paragraph."
        assert result[0]["runs"][0]["bold"] is False
        assert result[1]["runs"][0]["bold"] is True
    
    def test_extract_tables(self, document_manager, mock_docx_document):
        """Test internal table extraction method."""
        result = document_manager._extract_tables(mock_docx_document)
        
        assert len(result) == 1
        table = result[0]
        assert table["index"] == 0
        assert table["rows"] == 2
        assert table["columns"] == 2
        
        # Check table data
        assert len(table["data"]) == 2
        assert table["data"][0][0]["text"] == "Header 1"
        assert table["data"][0][1]["text"] == "Header 2"
        assert table["data"][1][0]["text"] == "Data 1"
        assert table["data"][1][1]["text"] == "Data 2"
    
    @patch('word_mcp_server.word.document_manager.Path')
    def test_extract_basic_metadata(self, mock_path_class, document_manager, mock_docx_document):
        """Test internal metadata extraction method."""
        # Mock file path and stats
        mock_path_instance = Mock()
        mock_path_instance.name = "test.docx"
        mock_path_instance.stat.return_value = Mock(st_size=1024, st_ctime=1640995200, st_mtime=1641081600)
        mock_path_class.return_value = mock_path_instance
        
        result = document_manager._extract_basic_metadata("test.docx", mock_docx_document)
        
        assert result["filename"] == "test.docx"
        assert result["file_size"] == 1024
        assert result["paragraph_count"] == 3
        assert result["table_count"] == 1
        assert result["title"] == "Test Document"
        assert result["author"] == "Test Author"
        assert result["subject"] == "Test Subject"
        assert result["document_created"] == "2023-01-01T00:00:00"
        assert result["document_modified"] == "2023-01-02T00:00:00"
    
    @patch('word_mcp_server.word.document_manager.Path')
    def test_extract_basic_metadata_no_properties(self, mock_path_class, document_manager):
        """Test metadata extraction when document properties are not available."""
        # Mock file path and stats
        mock_path_instance = Mock()
        mock_path_instance.name = "test.docx"
        mock_path_instance.stat.return_value = Mock(st_size=1024, st_ctime=1640995200, st_mtime=1641081600)
        mock_path_class.return_value = mock_path_instance
        
        # Mock document without properties
        doc = Mock()
        doc.paragraphs = []
        doc.tables = []
        doc.core_properties = Mock()
        doc.core_properties.title = None
        doc.core_properties.author = None
        doc.core_properties.subject = None
        doc.core_properties.created = None
        doc.core_properties.modified = None
        
        result = document_manager._extract_basic_metadata("test.docx", doc)
        
        assert result["filename"] == "test.docx"
        assert result["file_size"] == 1024
        assert result["paragraph_count"] == 0
        assert result["table_count"] == 0
        assert "title" not in result
        assert "author" not in result
        assert "subject" not in result
    
    # Tests for Task 6.2 - Metadata and statistics extraction
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    @patch('word_mcp_server.word.document_manager.Document')
    @patch('word_mcp_server.word.document_manager.Path')
    def test_get_document_info_success(self, mock_path, mock_document, mock_exists, mock_validate, document_manager, mock_docx_document):
        """Test successful document info extraction."""
        mock_validate.return_value = True
        mock_exists.return_value = True
        mock_document.return_value = mock_docx_document
        
        # Mock Path object
        mock_file_path = Mock()
        mock_file_path.name = "test.docx"
        mock_file_path.stat.return_value = Mock(st_size=1024, st_ctime=1640995200, st_mtime=1641081600)
        mock_path.return_value = mock_file_path
        
        result = document_manager.get_document_info("test.docx")
        
        # Should contain basic metadata
        assert result["filename"] == "test.docx"
        assert result["title"] == "Test Document"
        assert result["author"] == "Test Author"
        
        # Should contain statistics
        assert "word_count" in result
        assert "character_count" in result
        assert "paragraph_count" in result
        assert "reading_time_minutes" in result
        assert result["word_count"] > 0
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    @patch('word_mcp_server.word.document_manager.Document')
    def test_get_document_statistics_success(self, mock_document, mock_exists, mock_validate, document_manager, mock_docx_document):
        """Test successful document statistics calculation."""
        mock_validate.return_value = True
        mock_exists.return_value = True
        mock_document.return_value = mock_docx_document
        
        result = document_manager.get_document_statistics("test.docx")
        
        # Check required statistics
        assert "word_count" in result
        assert "character_count" in result
        assert "character_count_no_spaces" in result
        assert "paragraph_count" in result
        assert "sentence_count" in result
        assert "table_count" in result
        assert "image_count" in result
        assert "heading_counts" in result
        assert "total_headings" in result
        assert "reading_time_minutes" in result
        
        # Verify values make sense
        assert result["word_count"] > 0
        assert result["character_count"] > result["character_count_no_spaces"]
        assert result["paragraph_count"] == 3  # From mock data (all paragraphs with text)
        assert result["table_count"] == 1  # From mock data
        assert result["reading_time_minutes"] >= 1
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    @patch('word_mcp_server.word.document_manager.os.path.exists')
    @patch('word_mcp_server.word.document_manager.Document')
    def test_extract_comments_success(self, mock_document, mock_exists, mock_validate, document_manager):
        """Test successful comment extraction."""
        mock_validate.return_value = True
        mock_exists.return_value = True
        
        # Mock document with comment references
        doc = Mock()
        doc._element = Mock()
        comment_ref = Mock()
        comment_ref.get.return_value = "comment1"
        doc._element.xpath.return_value = [comment_ref]
        mock_document.return_value = doc
        
        result = document_manager.extract_comments("test.docx")
        
        assert isinstance(result, list)
        if len(result) > 0:
            comment = result[0]
            assert "id" in comment
            assert "text" in comment
            assert "author" in comment
            assert "position" in comment
    
    def test_calculate_document_statistics(self, document_manager, mock_docx_document):
        """Test internal statistics calculation method."""
        result = document_manager._calculate_document_statistics(mock_docx_document)
        
        # Check all expected statistics are present
        expected_keys = [
            "word_count", "character_count", "character_count_no_spaces",
            "paragraph_count", "sentence_count", "table_count", "image_count",
            "heading_counts", "total_headings", "reading_time_minutes"
        ]
        
        for key in expected_keys:
            assert key in result
        
        # Verify specific values from mock data
        assert result["paragraph_count"] == 3  # All paragraphs with text (including headings)
        assert result["table_count"] == 1
        assert result["word_count"] > 0
        assert result["reading_time_minutes"] >= 1
        assert "Heading 1" in result["heading_counts"]
        assert result["total_headings"] == 1
    
    def test_extract_document_properties(self, document_manager, mock_docx_document):
        """Test internal document properties extraction method."""
        result = document_manager._extract_document_properties(mock_docx_document)
        
        # Check expected properties
        assert result["title"] == "Test Document"
        assert result["author"] == "Test Author"
        assert result["subject"] == "Test Subject"
        assert result["created"] == "2023-01-01T00:00:00"
        assert result["modified"] == "2023-01-02T00:00:00"
    
    def test_extract_document_properties_no_properties(self, document_manager):
        """Test document properties extraction when properties are not available."""
        # Mock document without properties
        doc = Mock()
        doc.core_properties = Mock()
        doc.core_properties.title = None
        doc.core_properties.author = None
        doc.core_properties.subject = None
        doc.core_properties.keywords = None
        doc.core_properties.comments = None
        doc.core_properties.category = None
        doc.core_properties.language = None
        doc.core_properties.created = None
        doc.core_properties.modified = None
        doc.core_properties.last_modified_by = None
        doc.core_properties.revision = None
        doc.core_properties.version = None
        
        result = document_manager._extract_document_properties(doc)
        
        # Should return empty dict when no properties are available
        assert isinstance(result, dict)
        assert len(result) == 0
    
    def test_extract_comments_from_doc_no_comments(self, document_manager):
        """Test comment extraction when no comments are present."""
        # Mock document without comments
        doc = Mock()
        doc._element = Mock()
        doc._element.xpath.return_value = []
        
        result = document_manager._extract_comments_from_doc(doc)
        
        assert isinstance(result, list)
        assert len(result) == 0
    
    def test_extract_comments_from_doc_with_comments(self, document_manager):
        """Test comment extraction when comments are present."""
        # Mock document with comment references
        doc = Mock()
        doc._element = Mock()
        
        comment_ref1 = Mock()
        comment_ref1.get.return_value = "comment1"
        comment_ref2 = Mock()
        comment_ref2.get.return_value = "comment2"
        
        doc._element.xpath.return_value = [comment_ref1, comment_ref2]
        
        result = document_manager._extract_comments_from_doc(doc)
        
        assert isinstance(result, list)
        assert len(result) == 2
        
        # Check first comment
        comment1 = result[0]
        assert comment1["id"] == "comment1"
        assert "text" in comment1
        assert "author" in comment1
        assert comment1["position"] == 0
        
        # Check second comment
        comment2 = result[1]
        assert comment2["id"] == "comment2"
        assert comment2["position"] == 1
    
    def test_extract_comments_from_doc_exception_handling(self, document_manager):
        """Test comment extraction with exception handling."""
        # Mock document that raises exception during xpath
        doc = Mock()
        doc._element = Mock()
        doc._element.xpath.side_effect = Exception("XPath error")
        
        result = document_manager._extract_comments_from_doc(doc)
        
        # Should return empty list when exception occurs
        assert isinstance(result, list)
        assert len(result) == 0
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    def test_get_document_info_invalid_path(self, mock_validate, document_manager):
        """Test get_document_info with invalid path."""
        mock_validate.return_value = False
        
        with pytest.raises(DocumentError, match="Invalid file path"):
            document_manager.get_document_info("invalid/path")
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    def test_get_document_statistics_invalid_path(self, mock_validate, document_manager):
        """Test get_document_statistics with invalid path."""
        mock_validate.return_value = False
        
        with pytest.raises(DocumentError, match="Invalid file path"):
            document_manager.get_document_statistics("invalid/path")
    
    @patch('word_mcp_server.word.document_manager.validate_file_path')
    def test_extract_comments_invalid_path(self, mock_validate, document_manager):
        """Test extract_comments with invalid path."""
        mock_validate.return_value = False
        
        with pytest.raises(DocumentError, match="Invalid file path"):
            document_manager.extract_comments("invalid/path")