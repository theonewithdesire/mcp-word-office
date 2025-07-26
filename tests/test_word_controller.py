"""
Unit tests for WordController class with mocked COM interfaces.
"""

import pytest
import uuid
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime
from pathlib import Path

from word_mcp_server.word.controller import WordController, DocumentReference
from word_mcp_server.config.models import WordConfig
from word_mcp_server.utils.errors import ConnectionError, DocumentError, OperationError


class TestWordController:
    """Test cases for WordController class."""
    
    @pytest.fixture
    def word_config(self):
        """Create a test WordConfig instance."""
        return WordConfig(
            auto_launch=True,
            visible=False,
            save_on_exit=True,
            backup_enabled=True
        )
    
    @pytest.fixture
    def mock_word_app(self):
        """Create a mock Word application object."""
        mock_app = Mock()
        mock_app.Version = "16.0"
        mock_app.Visible = False
        mock_app.DisplayAlerts = False
        mock_app.Documents = Mock()
        return mock_app
    
    @pytest.fixture
    def mock_document(self):
        """Create a mock Word document object."""
        mock_doc = Mock()
        mock_doc.Name = "Document1"
        mock_doc.Save = Mock()
        mock_doc.SaveAs2 = Mock()
        mock_doc.Close = Mock()
        return mock_doc
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_init_with_com_available(self, word_config):
        """Test WordController initialization when COM is available."""
        controller = WordController(word_config)
        
        assert controller.config == word_config
        assert controller._word_app is None
        assert controller._documents == {}
        assert controller._max_connection_attempts == 3
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', False)
    def test_init_without_com_raises_error(self, word_config):
        """Test WordController initialization when COM is not available."""
        with pytest.raises(ConnectionError) as exc_info:
            WordController(word_config)
        
        assert "COM interface not available" in str(exc_info.value)
        assert "pip install pywin32" in exc_info.value.details
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.pythoncom')
    @patch('word_mcp_server.word.controller.win32com')
    def test_connect_to_existing_word_success(self, mock_win32com, mock_pythoncom, 
                                            word_config, mock_word_app):
        """Test successful connection to existing Word instance."""
        mock_win32com.client.GetActiveObject.return_value = mock_word_app
        
        controller = WordController(word_config)
        result = controller.connect_to_word()
        
        assert result is True
        assert controller._word_app == mock_word_app
        mock_pythoncom.CoInitialize.assert_called_once()
        mock_win32com.client.GetActiveObject.assert_called_once_with("Word.Application")
        assert mock_word_app.Visible == word_config.visible
        assert mock_word_app.DisplayAlerts is False
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.pythoncom')
    @patch('word_mcp_server.word.controller.win32com')
    def test_launch_new_word_success(self, mock_win32com, mock_pythoncom, 
                                   word_config, mock_word_app):
        """Test successful launch of new Word instance."""
        # Simulate no existing Word instance
        mock_win32com.client.GetActiveObject.side_effect = Exception("No active object")
        mock_win32com.client.Dispatch.return_value = mock_word_app
        
        controller = WordController(word_config)
        result = controller.connect_to_word()
        
        assert result is True
        assert controller._word_app == mock_word_app
        mock_win32com.client.Dispatch.assert_called_once_with("Word.Application")
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.pythoncom')
    @patch('word_mcp_server.word.controller.win32com')
    def test_connect_to_word_failure_with_retry(self, mock_win32com, mock_pythoncom, 
                                              word_config):
        """Test connection failure with retry attempts."""
        mock_win32com.client.GetActiveObject.side_effect = Exception("Connection failed")
        mock_win32com.client.Dispatch.side_effect = Exception("Launch failed")
        
        controller = WordController(word_config)
        
        with pytest.raises(ConnectionError) as exc_info:
            controller.connect_to_word()
        
        assert "Could not establish connection to Word application" in str(exc_info.value)
        # Should try GetActiveObject and Dispatch for each attempt
        assert mock_win32com.client.GetActiveObject.call_count == 3
        assert mock_win32com.client.Dispatch.call_count == 3
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_is_connected_true(self, word_config, mock_word_app):
        """Test is_connected returns True when connected."""
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        result = controller.is_connected()
        
        assert result is True
        # Should access Version property to test connection
        _ = mock_word_app.Version
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_is_connected_false_no_app(self, word_config):
        """Test is_connected returns False when no Word app."""
        controller = WordController(word_config)
        
        result = controller.is_connected()
        
        assert result is False
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_is_connected_false_app_unresponsive(self, word_config):
        """Test is_connected returns False when Word app is unresponsive."""
        from unittest.mock import PropertyMock
        
        mock_app = Mock()
        # Configure the Version property to raise an exception when accessed
        type(mock_app).Version = PropertyMock(side_effect=Exception("COM error"))
        
        controller = WordController(word_config)
        controller._word_app = mock_app
        
        result = controller.is_connected()
        
        assert result is False
        assert controller._word_app is None
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_create_document_success(self, word_config, mock_word_app, mock_document):
        """Test successful document creation."""
        mock_word_app.Documents.Add.return_value = mock_document
        
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        with patch('uuid.uuid4') as mock_uuid:
            mock_uuid.return_value = uuid.UUID('12345678-1234-5678-9012-123456789012')
            doc_id = controller.create_document()
        
        assert doc_id == '12345678-1234-5678-9012-123456789012'
        assert doc_id in controller._documents
        
        doc_ref = controller._documents[doc_id]
        assert doc_ref.doc_id == doc_id
        assert doc_ref.title == "Document1"
        assert doc_ref.is_active is True
        assert doc_ref.word_doc_ref == mock_document
        
        mock_word_app.Documents.Add.assert_called_once()
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_create_document_not_connected(self, word_config):
        """Test create_document raises error when not connected."""
        controller = WordController(word_config)
        
        with pytest.raises(ConnectionError) as exc_info:
            controller.create_document()
        
        assert "Not connected to Word application" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_create_document_operation_failure(self, word_config, mock_word_app):
        """Test create_document handles operation failures."""
        mock_word_app.Documents.Add.side_effect = Exception("Creation failed")
        
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        with pytest.raises(OperationError) as exc_info:
            controller.create_document()
        
        assert "Failed to create new document" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_open_document_success(self, word_config, mock_word_app, mock_document, tmp_path):
        """Test successful document opening."""
        # Create a temporary file
        test_file = tmp_path / "test.docx"
        test_file.write_text("test content")
        
        mock_word_app.Documents.Open.return_value = mock_document
        
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        with patch('uuid.uuid4') as mock_uuid:
            mock_uuid.return_value = uuid.UUID('12345678-1234-5678-9012-123456789012')
            doc_id = controller.open_document(str(test_file))
        
        assert doc_id == '12345678-1234-5678-9012-123456789012'
        assert doc_id in controller._documents
        
        doc_ref = controller._documents[doc_id]
        assert doc_ref.file_path == str(test_file.absolute())
        assert doc_ref.word_doc_ref == mock_document
        
        mock_word_app.Documents.Open.assert_called_once_with(str(test_file.absolute()))
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_open_document_file_not_found(self, word_config, mock_word_app):
        """Test open_document raises error for non-existent file."""
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        with pytest.raises(DocumentError) as exc_info:
            controller.open_document("/nonexistent/file.docx")
        
        assert "Document not found" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_open_document_access_denied(self, word_config, mock_word_app, tmp_path):
        """Test open_document handles access denied errors."""
        test_file = tmp_path / "test.docx"
        test_file.write_text("test content")
        
        mock_word_app.Documents.Open.side_effect = Exception("Access denied")
        
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        with pytest.raises(DocumentError) as exc_info:
            controller.open_document(str(test_file))
        
        assert "Access denied to document" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_close_document_success(self, word_config, mock_document):
        """Test successful document closing."""
        controller = WordController(word_config)
        
        # Add a document to the controller
        doc_id = "test-doc-id"
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Doc",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        controller.close_document(doc_id, save=True)
        
        assert doc_id not in controller._documents
        mock_document.Close.assert_called_once_with(SaveChanges=True)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_close_document_not_found(self, word_config):
        """Test close_document raises error for non-existent document."""
        controller = WordController(word_config)
        
        with pytest.raises(DocumentError) as exc_info:
            controller.close_document("nonexistent-doc-id")
        
        assert "Document not found" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_save_document_success(self, word_config, mock_document):
        """Test successful document saving."""
        import time
        controller = WordController(word_config)
        
        # Add a document to the controller
        doc_id = "test-doc-id"
        original_time = datetime.now()
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path="/test/path.docx",
            title="Test Doc",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=original_time,
            last_modified=original_time
        )
        controller._documents[doc_id] = doc_ref
        
        # Add a small delay to ensure time difference
        time.sleep(0.001)
        controller.save_document(doc_id)
        
        mock_document.Save.assert_called_once()
        # last_modified should be updated
        assert controller._documents[doc_id].last_modified > original_time
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_save_document_with_path(self, word_config, mock_document):
        """Test document saving with new path."""
        controller = WordController(word_config)
        
        # Add a document to the controller
        doc_id = "test-doc-id"
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Doc",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        new_path = "/new/path.docx"
        controller.save_document(doc_id, new_path)
        
        mock_document.SaveAs2.assert_called_once_with(str(Path(new_path).absolute()))
        assert controller._documents[doc_id].file_path == str(Path(new_path).absolute())
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_save_document_not_found(self, word_config):
        """Test save_document raises error for non-existent document."""
        controller = WordController(word_config)
        
        with pytest.raises(DocumentError) as exc_info:
            controller.save_document("nonexistent-doc-id")
        
        assert "Document not found" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_get_document_reference_success(self, word_config):
        """Test successful document reference retrieval."""
        controller = WordController(word_config)
        
        # Add a document to the controller
        doc_id = "test-doc-id"
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Doc",
            is_active=True,
            word_doc_ref=None,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        result = controller.get_document_reference(doc_id)
        
        assert result == doc_ref
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_get_document_reference_not_found(self, word_config):
        """Test get_document_reference raises error for non-existent document."""
        controller = WordController(word_config)
        
        with pytest.raises(DocumentError) as exc_info:
            controller.get_document_reference("nonexistent-doc-id")
        
        assert "Document not found" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_list_documents(self, word_config):
        """Test listing all documents."""
        controller = WordController(word_config)
        
        # Add some documents
        doc_ref1 = DocumentReference(
            doc_id="doc1",
            file_path=None,
            title="Doc 1",
            is_active=True,
            word_doc_ref=None,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        doc_ref2 = DocumentReference(
            doc_id="doc2",
            file_path="/path/doc2.docx",
            title="Doc 2",
            is_active=True,
            word_doc_ref=None,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        
        controller._documents["doc1"] = doc_ref1
        controller._documents["doc2"] = doc_ref2
        
        result = controller.list_documents()
        
        assert len(result) == 2
        assert result["doc1"] == doc_ref1
        assert result["doc2"] == doc_ref2
        # Should return a copy, not the original dict
        assert result is not controller._documents
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.pythoncom')
    def test_disconnect(self, mock_pythoncom, word_config, mock_word_app, mock_document):
        """Test disconnection from Word application."""
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        # Add a document
        doc_id = "test-doc-id"
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Doc",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        controller.disconnect()
        
        # Should close all documents
        mock_document.Close.assert_called_once_with(SaveChanges=True)
        assert len(controller._documents) == 0
        
        # Should quit Word application
        mock_word_app.Quit.assert_called_once_with(SaveChanges=True)
        assert controller._word_app is None
        
        # Should uninitialize COM
        mock_pythoncom.CoUninitialize.assert_called_once()
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.pythoncom')
    @patch('word_mcp_server.word.controller.win32com')
    def test_context_manager(self, mock_win32com, mock_pythoncom, word_config, mock_word_app):
        """Test WordController as context manager."""
        mock_win32com.client.GetActiveObject.return_value = mock_word_app
        
        with WordController(word_config) as controller:
            assert controller.is_connected()
            assert controller._word_app == mock_word_app
        
        # Should disconnect on exit
        mock_word_app.Quit.assert_called_once()
    
    # Tests for headers, footers, and page formatting (Task 7)
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.word_constants')
    def test_insert_header_footer_success(self, mock_constants, word_config, mock_word_app, mock_document):
        """Test successful header and footer insertion."""
        # Setup mocks
        mock_constants.wdHeaderFooterPrimary = 1
        
        mock_section = Mock()
        mock_header = Mock()
        mock_footer = Mock()
        mock_header_range = Mock()
        mock_footer_range = Mock()
        
        mock_section.Headers.return_value = mock_header
        mock_section.Footers.return_value = mock_footer
        mock_header.Range = mock_header_range
        mock_footer.Range = mock_footer_range
        
        mock_document.Sections.Count = 1
        mock_document.Sections.return_value = mock_section
        
        # Setup controller with connected Word app and document
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        doc_id = str(uuid.uuid4())
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Document",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        # Test header and footer insertion
        result = controller.insert_header_footer(doc_id, "Test Header", "Test Footer", 1)
        
        # Verify results
        assert result["header"]["success"] is True
        assert result["header"]["text"] == "Test Header"
        assert result["header"]["section"] == 1
        assert result["footer"]["success"] is True
        assert result["footer"]["text"] == "Test Footer"
        assert result["footer"]["section"] == 1
        
        # Verify mock calls
        mock_document.Sections.assert_called_with(1)
        mock_section.Headers.assert_called_with(1)
        mock_section.Footers.assert_called_with(1)
        assert mock_header_range.Text == "Test Header"
        assert mock_footer_range.Text == "Test Footer"
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_insert_header_footer_document_not_found(self, word_config):
        """Test header/footer insertion with non-existent document."""
        controller = WordController(word_config)
        
        with pytest.raises(DocumentError) as exc_info:
            controller.insert_header_footer("nonexistent", "Header", "Footer")
        
        assert "Document not found" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.word_constants')
    def test_insert_header_footer_only_header(self, mock_constants, word_config, mock_word_app, mock_document):
        """Test inserting only header text."""
        # Setup mocks
        mock_constants.wdHeaderFooterPrimary = 1
        
        mock_section = Mock()
        mock_header = Mock()
        mock_header_range = Mock()
        
        mock_section.Headers.return_value = mock_header
        mock_header.Range = mock_header_range
        
        mock_document.Sections.Count = 1
        mock_document.Sections.return_value = mock_section
        
        # Setup controller
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        doc_id = str(uuid.uuid4())
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Document",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        # Test header-only insertion
        result = controller.insert_header_footer(doc_id, header_text="Test Header")
        
        # Verify results
        assert result["header"]["success"] is True
        assert result["header"]["text"] == "Test Header"
        assert "footer" not in result
        
        # Verify mock calls
        mock_section.Headers.assert_called_with(1)
        assert mock_header_range.Text == "Test Header"
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.word_constants')
    def test_insert_page_break_success(self, mock_constants, word_config, mock_word_app, mock_document):
        """Test successful page break insertion."""
        # Setup mocks
        mock_constants.wdPageBreak = 7
        mock_constants.wdSectionBreakNextPage = 2
        
        mock_range = Mock()
        mock_range.Text = "Sample text"
        mock_document.Range.return_value = mock_range
        
        # Setup controller
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        doc_id = str(uuid.uuid4())
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Document",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        # Test page break insertion
        result = controller.insert_page_break(doc_id, position=10, break_type="page")
        
        # Verify results
        assert result["success"] is True
        assert result["break_type"] == "page"
        assert result["position"] == 10
        
        # Verify mock calls
        mock_range.SetRange.assert_called_with(10, 10)
        mock_range.InsertBreak.assert_called_with(7)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.word_constants')
    def test_insert_page_break_at_end(self, mock_constants, word_config, mock_word_app, mock_document):
        """Test page break insertion at end of document."""
        # Setup mocks
        mock_constants.wdPageBreak = 7
        
        mock_range = Mock()
        mock_range.Text = "Sample text"
        mock_document.Range.return_value = mock_range
        
        # Setup controller
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        doc_id = str(uuid.uuid4())
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Document",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        # Test page break insertion at end (position=None)
        result = controller.insert_page_break(doc_id, position=None, break_type="page")
        
        # Verify results
        assert result["success"] is True
        assert result["break_type"] == "page"
        assert result["position"] is None
        
        # Verify mock calls - should set range to end of document
        text_length = len(mock_range.Text)
        mock_range.SetRange.assert_called_with(text_length, text_length)
        mock_range.InsertBreak.assert_called_with(7)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_insert_page_break_document_not_found(self, word_config):
        """Test page break insertion with non-existent document."""
        controller = WordController(word_config)
        
        with pytest.raises(DocumentError) as exc_info:
            controller.insert_page_break("nonexistent")
        
        assert "Document not found" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.word_constants')
    def test_set_page_formatting_success(self, mock_constants, word_config, mock_word_app, mock_document):
        """Test successful page formatting."""
        # Setup mocks
        mock_constants.wdOrientPortrait = 0
        mock_constants.wdOrientLandscape = 1
        mock_constants.wdPaperLetter = 1
        
        mock_section = Mock()
        mock_page_setup = Mock()
        mock_section.PageSetup = mock_page_setup
        
        mock_document.Sections.Count = 1
        mock_document.Sections.return_value = mock_section
        
        # Setup controller
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        doc_id = str(uuid.uuid4())
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Document",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        # Test page formatting
        margins = {"top": 72, "bottom": 72, "left": 90, "right": 90}
        result = controller.set_page_formatting(
            doc_id, 
            section_index=1, 
            margins=margins, 
            orientation="landscape", 
            paper_size="letter"
        )
        
        # Verify results
        assert result["success"] is True
        assert result["section"] == 1
        assert result["top_margin"] == 72
        assert result["bottom_margin"] == 72
        assert result["left_margin"] == 90
        assert result["right_margin"] == 90
        assert result["orientation"] == "landscape"
        assert result["paper_size"] == "letter"
        
        # Verify mock calls
        mock_document.Sections.assert_called_with(1)
        assert mock_page_setup.TopMargin == 72
        assert mock_page_setup.BottomMargin == 72
        assert mock_page_setup.LeftMargin == 90
        assert mock_page_setup.RightMargin == 90
        assert mock_page_setup.Orientation == 1  # landscape
        assert mock_page_setup.PaperSize == 1  # letter
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.word_constants')
    def test_set_page_formatting_margins_only(self, mock_constants, word_config, mock_word_app, mock_document):
        """Test setting only margins."""
        mock_section = Mock()
        mock_page_setup = Mock()
        mock_section.PageSetup = mock_page_setup
        
        mock_document.Sections.Count = 1
        mock_document.Sections.return_value = mock_section
        
        # Setup controller
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        doc_id = str(uuid.uuid4())
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Document",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        # Test margins-only formatting
        margins = {"top": 36, "left": 54}
        result = controller.set_page_formatting(doc_id, margins=margins)
        
        # Verify results
        assert result["success"] is True
        assert result["top_margin"] == 36
        assert result["left_margin"] == 54
        assert "bottom_margin" not in result
        assert "right_margin" not in result
        assert "orientation" not in result
        assert "paper_size" not in result
        
        # Verify mock calls
        assert mock_page_setup.TopMargin == 36
        assert mock_page_setup.LeftMargin == 54
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    def test_set_page_formatting_document_not_found(self, word_config):
        """Test page formatting with non-existent document."""
        controller = WordController(word_config)
        
        with pytest.raises(DocumentError) as exc_info:
            controller.set_page_formatting("nonexistent", margins={"top": 72})
        
        assert "Document not found" in str(exc_info.value)
    
    @patch('word_mcp_server.word.controller.COM_AVAILABLE', True)
    @patch('word_mcp_server.word.controller.word_constants')
    def test_set_page_formatting_invalid_orientation(self, mock_constants, word_config, mock_word_app, mock_document):
        """Test page formatting with invalid orientation."""
        mock_section = Mock()
        mock_page_setup = Mock()
        mock_section.PageSetup = mock_page_setup
        
        mock_document.Sections.Count = 1
        mock_document.Sections.return_value = mock_section
        
        # Setup controller
        controller = WordController(word_config)
        controller._word_app = mock_word_app
        
        doc_id = str(uuid.uuid4())
        doc_ref = DocumentReference(
            doc_id=doc_id,
            file_path=None,
            title="Test Document",
            is_active=True,
            word_doc_ref=mock_document,
            created_at=datetime.now(),
            last_modified=datetime.now()
        )
        controller._documents[doc_id] = doc_ref
        
        # Test invalid orientation
        result = controller.set_page_formatting(doc_id, orientation="invalid")
        
        # Verify results
        assert result["success"] is True
        assert "orientation_error" in result
        assert "Invalid orientation: invalid" in result["orientation_error"]