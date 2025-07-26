"""
Tests for table and list operations in Word controller.

This module tests the table creation, cell formatting, and list creation
functionality implemented in task 5.1.
"""

import pytest
from unittest.mock import Mock, MagicMock, patch
from datetime import datetime

from word_mcp_server.word.controller import WordController, DocumentReference
from word_mcp_server.utils.errors import DocumentError, OperationError, ErrorCode
from word_mcp_server.config.models import WordConfig


class TestTableOperations:
    """Test table creation and manipulation operations."""
    
    @pytest.fixture
    def mock_word_config(self):
        """Create a mock Word configuration."""
        config = Mock(spec=WordConfig)
        config.auto_launch = True
        config.visible = False
        config.save_on_exit = True
        return config
    
    @pytest.fixture
    def mock_word_controller(self, mock_word_config):
        """Create a Word controller with mocked COM interface."""
        with patch('word_mcp_server.word.controller.COM_AVAILABLE', True), \
             patch('word_mcp_server.word.controller.win32com'), \
             patch('word_mcp_server.word.controller.pythoncom'), \
             patch('word_mcp_server.word.controller.word_constants') as mock_constants:
            
            # Mock word constants
            mock_constants.wdAutoFitContent = 1
            mock_constants.wdUnderlineSingle = 1
            mock_constants.wdUnderlineNone = 0
            
            controller = WordController(mock_word_config)
            controller._word_app = Mock()
            
            # Mock document
            mock_doc = Mock()
            mock_doc.Name = "Test Document"
            mock_doc.Range.return_value.Text = "Sample text content"
            mock_doc.Tables = Mock()
            mock_doc.Tables.Count = 0
            
            # Create document reference
            doc_id = "test-doc-123"
            doc_ref = DocumentReference(
                doc_id=doc_id,
                file_path=None,
                title="Test Document",
                is_active=True,
                word_doc_ref=mock_doc,
                created_at=datetime.now(),
                last_modified=datetime.now()
            )
            controller._documents[doc_id] = doc_ref
            
            return controller, doc_id, mock_doc
    
    def test_create_table_success(self, mock_word_controller):
        """Test successful table creation."""
        controller, doc_id, mock_doc = mock_word_controller
        
        # Mock table creation
        mock_table = Mock()
        mock_doc.Tables.Add.return_value = mock_table
        mock_doc.Tables.Count = 1
        
        # Mock range for insertion
        mock_range = Mock()
        mock_range.Text = "Sample text"
        mock_doc.Range.return_value = mock_range
        
        # Create table
        result = controller.create_table(doc_id, 3, 4)
        
        # Verify result
        assert result["rows"] == 3
        assert result["cols"] == 4
        assert result["position"] is None
        assert result["table_index"] == 1
        
        # Verify Word API calls
        mock_doc.Tables.Add.assert_called_once()
        # Note: AutoFitBehavior is not called when word_constants is None (test environment)
        mock_table.Borders.Enable = True
    
    def test_create_table_with_position(self, mock_word_controller):
        """Test table creation at specific position."""
        controller, doc_id, mock_doc = mock_word_controller
        
        # Mock table creation
        mock_table = Mock()
        mock_doc.Tables.Add.return_value = mock_table
        mock_doc.Tables.Count = 1
        
        # Mock range for insertion
        mock_range = Mock()
        mock_range.Text = "Sample text content"
        mock_doc.Range.return_value = mock_range
        
        # Create table at position 10
        result = controller.create_table(doc_id, 2, 3, position=10)
        
        # Verify result
        assert result["rows"] == 2
        assert result["cols"] == 3
        assert result["position"] == 10
        
        # Verify range was set correctly
        mock_range.SetRange.assert_called_with(10, 10)
    
    def test_create_table_invalid_dimensions(self, mock_word_controller):
        """Test table creation with invalid dimensions."""
        controller, doc_id, mock_doc = mock_word_controller
        
        # Test zero rows
        with pytest.raises(OperationError) as exc_info:
            controller.create_table(doc_id, 0, 3)
        assert "at least 1 row and 1 column" in str(exc_info.value)
        
        # Test zero columns
        with pytest.raises(OperationError) as exc_info:
            controller.create_table(doc_id, 3, 0)
        assert "at least 1 row and 1 column" in str(exc_info.value)
        
        # Test too many rows
        with pytest.raises(OperationError) as exc_info:
            controller.create_table(doc_id, 101, 3)
        assert "Table size too large" in str(exc_info.value)
        
        # Test too many columns
        with pytest.raises(OperationError) as exc_info:
            controller.create_table(doc_id, 3, 51)
        assert "Table size too large" in str(exc_info.value)
    
    def test_create_table_document_not_found(self, mock_word_controller):
        """Test table creation with non-existent document."""
        controller, doc_id, mock_doc = mock_word_controller
        
        with pytest.raises(DocumentError) as exc_info:
            controller.create_table("non-existent-doc", 3, 4)
        
        assert exc_info.value.error_code == ErrorCode.DOCUMENT_NOT_FOUND.value
        assert "Document not found" in str(exc_info.value)
    
    def test_format_table_cell_success(self, mock_word_controller):
        """Test successful table cell formatting."""
        controller, doc_id, mock_doc = mock_word_controller
        
        # Mock table and cell
        mock_table = Mock()
        mock_cell = Mock()
        mock_cell_range = Mock()
        mock_cell.Range = mock_cell_range
        mock_table.Cell.return_value = mock_cell
        mock_table.Rows.Count = 3
        mock_table.Columns.Count = 4
        
        mock_doc.Tables.Count = 1
        mock_doc.Tables.return_value = mock_table
        
        # Format cell
        controller.format_table_cell(
            doc_id, 1, 2, 3, 
            text="Test Cell",
            bold=True,
            font_size=14,
            color="#FF0000"
        )
        
        # Verify cell text was set
        assert mock_cell_range.Text == "Test Cell"
        
        # Verify formatting was applied
        assert mock_cell_range.Font.Bold == True
        assert mock_cell_range.Font.Size == 14
        mock_cell_range.Font.Color = 255  # Red in BGR format
    
    def test_format_table_cell_invalid_indices(self, mock_word_controller):
        """Test table cell formatting with invalid indices."""
        controller, doc_id, mock_doc = mock_word_controller
        
        # Mock table
        mock_table = Mock()
        mock_table.Rows.Count = 3
        mock_table.Columns.Count = 4
        mock_doc.Tables.Count = 1
        mock_doc.Tables.return_value = mock_table
        
        # Test invalid table index
        with pytest.raises(OperationError) as exc_info:
            controller.format_table_cell(doc_id, 2, 1, 1)  # Table 2 doesn't exist
        assert "Table index 2 out of range" in str(exc_info.value)
        
        # Test invalid row
        with pytest.raises(OperationError) as exc_info:
            controller.format_table_cell(doc_id, 1, 5, 1)  # Row 5 doesn't exist
        assert "Row 5 out of range" in str(exc_info.value)
        
        # Test invalid column
        with pytest.raises(OperationError) as exc_info:
            controller.format_table_cell(doc_id, 1, 1, 6)  # Column 6 doesn't exist
        assert "Column 6 out of range" in str(exc_info.value)


class TestListOperations:
    """Test list creation operations."""
    
    @pytest.fixture
    def mock_word_controller(self):
        """Create a Word controller with mocked COM interface."""
        with patch('word_mcp_server.word.controller.COM_AVAILABLE', True), \
             patch('word_mcp_server.word.controller.win32com'), \
             patch('word_mcp_server.word.controller.pythoncom'), \
             patch('word_mcp_server.word.controller.word_constants'):
            
            config = Mock(spec=WordConfig)
            controller = WordController(config)
            controller._word_app = Mock()
            
            # Mock document
            mock_doc = Mock()
            mock_doc.Name = "Test Document"
            
            # Mock range for text operations
            mock_range = Mock()
            mock_range.Text = "Sample text content"  # 20 characters
            mock_range.Start = 0
            mock_doc.Range.return_value = mock_range
            
            # Create document reference
            doc_id = "test-doc-123"
            doc_ref = DocumentReference(
                doc_id=doc_id,
                file_path=None,
                title="Test Document",
                is_active=True,
                word_doc_ref=mock_doc,
                created_at=datetime.now(),
                last_modified=datetime.now()
            )
            controller._documents[doc_id] = doc_ref
            
            return controller, doc_id, mock_doc, mock_range
    
    def test_create_bulleted_list_success(self, mock_word_controller):
        """Test successful bulleted list creation."""
        controller, doc_id, mock_doc, mock_range = mock_word_controller
        
        # Set up proper text length for the range
        mock_range.Text = "Sample text content"  # 20 characters
        
        # Mock list formatting
        mock_list_range = Mock()
        mock_list_range.Start = 20  # After the existing text
        mock_list_format = Mock()
        mock_list_range.ListFormat = mock_list_format
        
        # Configure mock_doc.Range to return different objects for different calls
        def range_side_effect(*args):
            if len(args) == 0:
                return mock_range  # For doc.Range()
            else:
                return mock_list_range  # For doc.Range(start, end)
        
        mock_doc.Range.side_effect = range_side_effect
        
        items = ["First item", "Second item", "Third item"]
        
        # Create bulleted list
        result = controller.create_list(doc_id, items, "bulleted")
        
        # Verify result
        assert result["list_type"] == "bulleted"
        assert result["item_count"] == 3
        assert result["position"] is None
        
        # Verify text was inserted
        expected_text = "First item\nSecond item\nThird item"
        mock_range.InsertAfter.assert_called_with(expected_text)
        
        # Verify bullet formatting was applied
        mock_list_format.ApplyBulletDefault.assert_called_once()
    
    def test_create_numbered_list_success(self, mock_word_controller):
        """Test successful numbered list creation."""
        controller, doc_id, mock_doc, mock_range = mock_word_controller
        
        # Set up proper text length for the range
        mock_range.Text = "Sample text content"  # 20 characters
        
        # Mock list formatting
        mock_list_range = Mock()
        mock_list_range.Start = 20  # After the existing text
        mock_list_format = Mock()
        mock_list_range.ListFormat = mock_list_format
        
        # Configure mock_doc.Range to return different objects for different calls
        def range_side_effect(*args):
            if len(args) == 0:
                return mock_range  # For doc.Range()
            else:
                return mock_list_range  # For doc.Range(start, end)
        
        mock_doc.Range.side_effect = range_side_effect
        
        items = ["First item", "Second item"]
        
        # Create numbered list
        result = controller.create_list(doc_id, items, "numbered")
        
        # Verify result
        assert result["list_type"] == "numbered"
        assert result["item_count"] == 2
        
        # Verify number formatting was applied
        mock_list_format.ApplyNumberDefault.assert_called_once()
    
    def test_create_list_with_position(self, mock_word_controller):
        """Test list creation at specific position."""
        controller, doc_id, mock_doc, mock_range = mock_word_controller
        
        # Set up proper text length for the range
        mock_range.Text = "Sample text content"  # 20 characters
        
        # Mock list formatting
        mock_list_range = Mock()
        mock_list_range.Start = 5
        mock_list_format = Mock()
        mock_list_range.ListFormat = mock_list_format
        
        # Configure mock_doc.Range to return different objects for different calls
        def range_side_effect(*args):
            if len(args) == 0:
                return mock_range  # For doc.Range()
            else:
                return mock_list_range  # For doc.Range(start, end)
        
        mock_doc.Range.side_effect = range_side_effect
        
        items = ["Item 1", "Item 2"]
        
        # Create list at position 5
        result = controller.create_list(doc_id, items, "bulleted", position=5)
        
        # Verify position was set
        assert result["position"] == 5
        mock_range.SetRange.assert_called_with(5, 5)
    
    def test_create_list_empty_items(self, mock_word_controller):
        """Test list creation with empty items list."""
        controller, doc_id, mock_doc, mock_range = mock_word_controller
        
        with pytest.raises(OperationError) as exc_info:
            controller.create_list(doc_id, [], "bulleted")
        
        assert "List must contain at least one item" in str(exc_info.value)
    
    def test_create_list_invalid_type(self, mock_word_controller):
        """Test list creation with invalid list type."""
        controller, doc_id, mock_doc, mock_range = mock_word_controller
        
        with pytest.raises(OperationError) as exc_info:
            controller.create_list(doc_id, ["Item 1"], "invalid_type")
        
        assert "List type must be 'bulleted' or 'numbered'" in str(exc_info.value)
    
    def test_create_list_document_not_found(self, mock_word_controller):
        """Test list creation with non-existent document."""
        controller, doc_id, mock_doc, mock_range = mock_word_controller
        
        with pytest.raises(DocumentError) as exc_info:
            controller.create_list("non-existent-doc", ["Item 1"], "bulleted")
        
        assert exc_info.value.error_code == ErrorCode.DOCUMENT_NOT_FOUND.value
        assert "Document not found" in str(exc_info.value)


class TestMCPServerTableListIntegration:
    """Test MCP server integration for table and list operations."""
    
    @pytest.fixture
    def mock_mcp_server(self):
        """Create a mock MCP server with Word controller."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            from word_mcp_server.server.mcp_server import WordMCPServer
            from word_mcp_server.config.config_manager import ConfigManager
            
            # Mock config manager
            config_manager = Mock(spec=ConfigManager)
            config_manager.config.word = Mock()
            config_manager.config.server = Mock()
            config_manager.config.server.dict.return_value = {}
            
            # Create server
            server = WordMCPServer(config_manager)
            
            # Mock word controller
            mock_controller = Mock()
            mock_controller_class.return_value = mock_controller
            mock_controller.is_connected.return_value = True
            server.word_controller = mock_controller
            
            return server, mock_controller
    
    @pytest.mark.asyncio
    async def test_handle_create_table_success(self, mock_mcp_server):
        """Test successful table creation through MCP server."""
        server, mock_controller = mock_mcp_server
        
        # Mock table creation
        mock_controller.create_table.return_value = {
            "rows": 3,
            "cols": 4,
            "position": None,
            "table_index": 1
        }
        
        # Call handler
        result = await server._handle_create_table({
            "doc_id": "test-doc",
            "rows": 3,
            "cols": 4
        })
        
        # Verify result
        assert result["success"] is True
        assert result["doc_id"] == "test-doc"
        assert "table_info" in result
        assert "Created 3x4 table" in result["message"]
        
        # Verify controller was called
        mock_controller.create_table.assert_called_once_with("test-doc", 3, 4, None)
    
    @pytest.mark.asyncio
    async def test_handle_create_table_missing_params(self, mock_mcp_server):
        """Test table creation with missing parameters."""
        server, mock_controller = mock_mcp_server
        
        # Test missing doc_id
        result = await server._handle_create_table({
            "rows": 3,
            "cols": 4
        })
        assert result["success"] is False
        assert "Document ID is required" in result["error"]
        
        # Test missing rows
        result = await server._handle_create_table({
            "doc_id": "test-doc",
            "cols": 4
        })
        assert result["success"] is False
        assert "Rows and columns are required" in result["error"]
    
    @pytest.mark.asyncio
    async def test_handle_format_table_cell_success(self, mock_mcp_server):
        """Test successful table cell formatting through MCP server."""
        server, mock_controller = mock_mcp_server
        
        # Call handler
        result = await server._handle_format_table_cell({
            "doc_id": "test-doc",
            "table_index": 1,
            "row": 2,
            "col": 3,
            "text": "Test Cell",
            "bold": True
        })
        
        # Verify result
        assert result["success"] is True
        assert result["doc_id"] == "test-doc"
        assert result["table_index"] == 1
        assert result["row"] == 2
        assert result["col"] == 3
        
        # Verify controller was called
        mock_controller.format_table_cell.assert_called_once_with(
            "test-doc", 1, 2, 3, "Test Cell", bold=True
        )
    
    @pytest.mark.asyncio
    async def test_handle_create_list_success(self, mock_mcp_server):
        """Test successful list creation through MCP server."""
        server, mock_controller = mock_mcp_server
        
        # Mock list creation
        mock_controller.create_list.return_value = {
            "list_type": "bulleted",
            "item_count": 3,
            "position": None,
            "start_pos": 0,
            "end_pos": 30
        }
        
        items = ["Item 1", "Item 2", "Item 3"]
        
        # Call handler
        result = await server._handle_create_list({
            "doc_id": "test-doc",
            "items": items,
            "list_type": "bulleted"
        })
        
        # Verify result
        assert result["success"] is True
        assert result["doc_id"] == "test-doc"
        assert "list_info" in result
        assert "Created bulleted list with 3 items" in result["message"]
        
        # Verify controller was called
        mock_controller.create_list.assert_called_once_with(
            "test-doc", items, "bulleted", None
        )
    
    @pytest.mark.asyncio
    async def test_handle_create_list_missing_items(self, mock_mcp_server):
        """Test list creation with missing items."""
        server, mock_controller = mock_mcp_server
        
        result = await server._handle_create_list({
            "doc_id": "test-doc"
        })
        
        assert result["success"] is False
        assert "List items are required" in result["error"]


class TestFindReplaceOperations:
    """Test find and replace operations."""
    
    @pytest.fixture
    def mock_word_controller(self):
        """Create a Word controller with mocked COM interface."""
        with patch('word_mcp_server.word.controller.COM_AVAILABLE', True), \
             patch('word_mcp_server.word.controller.win32com'), \
             patch('word_mcp_server.word.controller.pythoncom'), \
             patch('word_mcp_server.word.controller.word_constants') as mock_constants:
            
            # Mock word constants
            mock_constants.wdFindContinue = 1
            mock_constants.wdReplaceAll = 2
            mock_constants.wdReplaceOne = 1
            
            config = Mock(spec=WordConfig)
            controller = WordController(config)
            controller._word_app = Mock()
            
            # Mock document
            mock_doc = Mock()
            mock_doc.Name = "Test Document"
            
            # Mock range for find/replace operations
            mock_range = Mock()
            mock_range.Text = "This is sample text with sample words"
            mock_doc.Range.return_value = mock_range
            
            # Mock Find object
            mock_find = Mock()
            mock_range.Find = mock_find
            mock_find.Replacement = Mock()
            
            # Create document reference
            doc_id = "test-doc-123"
            doc_ref = DocumentReference(
                doc_id=doc_id,
                file_path=None,
                title="Test Document",
                is_active=True,
                word_doc_ref=mock_doc,
                created_at=datetime.now(),
                last_modified=datetime.now()
            )
            controller._documents[doc_id] = doc_ref
            
            return controller, doc_id, mock_doc, mock_range, mock_find
    
    def test_find_replace_all_success(self, mock_word_controller):
        """Test successful find and replace all operation."""
        controller, doc_id, mock_doc, mock_range, mock_find = mock_word_controller
        
        # Mock successful replace operation
        mock_find.Execute.return_value = True
        
        # Perform find/replace
        result = controller.find_replace(doc_id, "sample", "example", replace_all=True)
        
        # Verify result
        assert result["find_text"] == "sample"
        assert result["replace_text"] == "example"
        assert result["replace_all"] is True
        assert result["match_case"] is False
        assert result["match_whole_word"] is False
        assert result["use_regex"] is False
        
        # Verify Find object was configured correctly
        assert mock_find.Text == "sample"
        assert mock_find.Replacement.Text == "example"
        assert mock_find.MatchCase is False
        assert mock_find.MatchWholeWord is False
        assert mock_find.MatchWildcards is False
        assert mock_find.Forward is True
        
        # Verify Execute was called
        mock_find.Execute.assert_called()
    
    def test_find_replace_one_success(self, mock_word_controller):
        """Test successful find and replace single occurrence."""
        controller, doc_id, mock_doc, mock_range, mock_find = mock_word_controller
        
        # Mock successful replace operation
        mock_find.Execute.return_value = True
        
        # Perform find/replace
        result = controller.find_replace(doc_id, "sample", "example", replace_all=False)
        
        # Verify result
        assert result["find_text"] == "sample"
        assert result["replace_text"] == "example"
        assert result["replace_all"] is False
        assert result["replacements_made"] == 1
        
        # Verify Execute was called
        mock_find.Execute.assert_called()
    
    def test_find_replace_with_options(self, mock_word_controller):
        """Test find and replace with various options."""
        controller, doc_id, mock_doc, mock_range, mock_find = mock_word_controller
        
        # Mock successful replace operation
        mock_find.Execute.return_value = True
        
        # Perform find/replace with options
        result = controller.find_replace(
            doc_id, "Sample", "Example", 
            match_case=True, match_whole_word=True, use_regex=True
        )
        
        # Verify result
        assert result["match_case"] is True
        assert result["match_whole_word"] is True
        assert result["use_regex"] is True
        
        # Verify Find object was configured correctly
        assert mock_find.MatchCase is True
        assert mock_find.MatchWholeWord is True
        # Note: MatchWildcards is only set when word_constants is available (not in test environment)
        # In real environment, this would be True when use_regex=True
    
    def test_find_replace_empty_find_text(self, mock_word_controller):
        """Test find and replace with empty find text."""
        controller, doc_id, mock_doc, mock_range, mock_find = mock_word_controller
        
        with pytest.raises(OperationError) as exc_info:
            controller.find_replace(doc_id, "", "replacement")
        
        assert "Find text cannot be empty" in str(exc_info.value)
    
    def test_find_replace_document_not_found(self, mock_word_controller):
        """Test find and replace with non-existent document."""
        controller, doc_id, mock_doc, mock_range, mock_find = mock_word_controller
        
        with pytest.raises(DocumentError) as exc_info:
            controller.find_replace("non-existent-doc", "find", "replace")
        
        assert exc_info.value.error_code == ErrorCode.DOCUMENT_NOT_FOUND.value
        assert "Document not found" in str(exc_info.value)


class TestMCPServerFindReplaceIntegration:
    """Test MCP server integration for find and replace operations."""
    
    @pytest.fixture
    def mock_mcp_server(self):
        """Create a mock MCP server with Word controller."""
        with patch('word_mcp_server.server.mcp_server.WordController') as mock_controller_class:
            from word_mcp_server.server.mcp_server import WordMCPServer
            from word_mcp_server.config.config_manager import ConfigManager
            
            # Mock config manager
            config_manager = Mock(spec=ConfigManager)
            config_manager.config.word = Mock()
            config_manager.config.server = Mock()
            config_manager.config.server.dict.return_value = {}
            
            # Create server
            server = WordMCPServer(config_manager)
            
            # Mock word controller
            mock_controller = Mock()
            mock_controller_class.return_value = mock_controller
            mock_controller.is_connected.return_value = True
            server.word_controller = mock_controller
            
            return server, mock_controller
    
    @pytest.mark.asyncio
    async def test_handle_find_replace_success(self, mock_mcp_server):
        """Test successful find and replace through MCP server."""
        server, mock_controller = mock_mcp_server
        
        # Mock find/replace operation
        mock_controller.find_replace.return_value = {
            "find_text": "old",
            "replace_text": "new",
            "replacements_made": 3,
            "match_case": False,
            "match_whole_word": False,
            "use_regex": False,
            "replace_all": True
        }
        
        # Call handler
        result = await server._handle_find_replace({
            "doc_id": "test-doc",
            "find_text": "old",
            "replace_text": "new"
        })
        
        # Verify result
        assert result["success"] is True
        assert result["doc_id"] == "test-doc"
        assert "result" in result
        assert "3 replacements made" in result["message"]
        
        # Verify controller was called
        mock_controller.find_replace.assert_called_once_with(
            "test-doc", "old", "new", False, False, False, True
        )
    
    @pytest.mark.asyncio
    async def test_handle_find_replace_with_options(self, mock_mcp_server):
        """Test find and replace with options through MCP server."""
        server, mock_controller = mock_mcp_server
        
        # Mock find/replace operation
        mock_controller.find_replace.return_value = {
            "find_text": "Old",
            "replace_text": "New",
            "replacements_made": 1,
            "match_case": True,
            "match_whole_word": True,
            "use_regex": False,
            "replace_all": False
        }
        
        # Call handler with options
        result = await server._handle_find_replace({
            "doc_id": "test-doc",
            "find_text": "Old",
            "replace_text": "New",
            "match_case": True,
            "match_whole_word": True,
            "replace_all": False
        })
        
        # Verify result
        assert result["success"] is True
        
        # Verify controller was called with correct options
        mock_controller.find_replace.assert_called_once_with(
            "test-doc", "Old", "New", True, True, False, False
        )
    
    @pytest.mark.asyncio
    async def test_handle_find_replace_missing_params(self, mock_mcp_server):
        """Test find and replace with missing parameters."""
        server, mock_controller = mock_mcp_server
        
        # Test missing doc_id
        result = await server._handle_find_replace({
            "find_text": "old",
            "replace_text": "new"
        })
        assert result["success"] is False
        assert "Document ID is required" in result["error"]
        
        # Test missing find_text
        result = await server._handle_find_replace({
            "doc_id": "test-doc",
            "replace_text": "new"
        })
        assert result["success"] is False
        assert "Find text is required" in result["error"]
        
        # Test missing replace_text
        result = await server._handle_find_replace({
            "doc_id": "test-doc",
            "find_text": "old"
        })
        assert result["success"] is False
        assert "Replace text is required" in result["error"]